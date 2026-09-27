'use strict';

/**
 * Assembles the finished .xlsx package from the per-sheet temp files.
 *
 * Each temp file is a complete one-sheet workbook written by xlsx-write-stream.
 * Merging means keeping one sheet's package parts (styles, and whatever else the
 * writer emits) and replacing everything sheet-specific with parts generated
 * here, since a workbook has exactly one workbook.xml and one relationship set.
 */

const JSZip = require('jszip');
const fs = require('fs');
const fsp = require('fs/promises');

const CONTENT_TYPES = {
    'xl/workbook.xml':
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    'xl/styles.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
    'xl/sharedStrings.xml':
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
};

const WORKSHEET_TYPE =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const WORKSHEET_REL =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

const SHARED_STRINGS =
    'The underlying writer produced a shared strings table, whose per-sheet indices cannot be ' +
    'merged. Refusing to write a workbook that Excel would have to repair.';

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const escapeXml = (str) =>
    String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');

/** jszip@3.10 folds `level: 0` into its own default (`level || -1`), so store mode is the only way to honour it. */
const archiveOptions = (compressionLevel) =>
    compressionLevel === 0
        ? { compression: 'STORE' }
        : { compression: 'DEFLATE', compressionOptions: { level: compressionLevel } };

/** Content types are derived from the parts actually shipped: a wrong one is what makes Excel offer to repair. */
const contentTypesXml = (worksheetParts, carriedParts) => {
    const overrides = [
        ['xl/workbook.xml', CONTENT_TYPES['xl/workbook.xml']],
        ...worksheetParts.map((part) => [part, WORKSHEET_TYPE])
    ];

    for (const part of carriedParts) {
        const declared =
            CONTENT_TYPES[part] ||
            (/^xl\/theme\//.test(part)
                ? 'application/vnd.openxmlformats-officedocument.theme+xml'
                : /^docProps\//.test(part)
                  ? 'application/vnd.openxmlformats-package.core-properties+xml'
                  : null);
        if (!declared) {
            throw new Error(
                `No content type is known for the part "${part}". The underlying writer started ` +
                    'emitting a part this library does not recognise. Refusing to write a workbook ' +
                    'with a declared-but-wrong content type.'
            );
        }
        overrides.push([part, declared]);
    }

    return `${XML_HEADER}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
${overrides.map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join('\n')}
</Types>`;
};

const copyPackageParts = async (baseZip, outputZip) => {
    const carried = new Set();

    for (const fileName of Object.keys(baseZip.files)) {
        const entry = baseZip.files[fileName];
        if (entry.dir) continue;
        if (fileName === 'xl/sharedStrings.xml') throw new Error(SHARED_STRINGS);
        if (fileName.startsWith('xl/worksheets/') || fileName.includes('workbook.xml')) continue;
        if (fileName === '[Content_Types].xml' || fileName === '_rels/.rels') continue;

        carried.add(fileName);
        outputZip.file(fileName, await entry.async('nodebuffer'));
    }

    return carried;
};

const copyWorksheetParts = async (sheets, outputZip) => {
    const entries = [];
    const relationships = [];
    const parts = [];

    for (let index = 0; index < sheets.length; index++) {
        const { name, filePath } = sheets[index];
        const sheetId = index + 1;
        const fileName = `sheet${sheetId}.xml`;

        const sheetZip = await JSZip.loadAsync(await fsp.readFile(filePath));
        const worksheet = Object.keys(sheetZip.files).find(
            (entry) => entry.startsWith('xl/worksheets/') && entry.endsWith('.xml')
        );
        if (!worksheet) {
            throw new Error(
                `Sheet "${name}" produced no worksheet XML, so the workbook would declare a sheet ` +
                    'that resolves to nothing. Refusing to write it.'
            );
        }
        if (sheetZip.files['xl/sharedStrings.xml']) throw new Error(SHARED_STRINGS);

        outputZip.file(`xl/worksheets/${fileName}`, await sheetZip.files[worksheet].async('nodebuffer'));
        entries.push(`<sheet name="${escapeXml(name)}" sheetId="${sheetId}" r:id="rId${sheetId}"/>`);
        relationships.push(`<Relationship Id="rId${sheetId}" Type="${WORKSHEET_REL}" Target="worksheets/${fileName}"/>`);
        parts.push(`xl/worksheets/${fileName}`);
    }

    return { entries, relationships, parts };
};

const addWorkbookParts = (outputZip, sheetCount, entries, relationships) => {
    outputZip.file(
        'xl/workbook.xml',
        `${XML_HEADER}
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>${entries.join('')}</sheets>
</workbook>`
    );

    // The styles relationship must not reuse a sheet's id, so it sits past them all.
    const stylesRelId = `rId${sheetCount + 1}`;
    outputZip.file(
        'xl/_rels/workbook.xml.rels',
        `${XML_HEADER}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${relationships.join('\n')}
<Relationship Id="${stylesRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
    );

    outputZip.file(
        '_rels/.rels',
        `${XML_HEADER}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
    );
};

const buildArchive = async (sheets) => {
    const outputZip = new JSZip();
    const baseZip = await JSZip.loadAsync(await fsp.readFile(sheets[0].filePath));

    const carried = await copyPackageParts(baseZip, outputZip);
    const { entries, relationships, parts } = await copyWorksheetParts(sheets, outputZip);

    addWorkbookParts(outputZip, sheets.length, entries, relationships);
    outputZip.file('[Content_Types].xml', contentTypesXml(parts, carried));

    return outputZip;
};

/** Streams the archive to disk so the compressed bytes are never held twice. */
const writeArchive = (archive, outputPath, compressionLevel) =>
    new Promise((resolve, reject) => {
        const out = fs.createWriteStream(outputPath);
        out.on('finish', resolve);
        out.on('error', reject);
        archive
            .generateNodeStream({ type: 'nodebuffer', ...archiveOptions(compressionLevel) })
            .on('error', reject)
            .pipe(out);
    });

module.exports = { buildArchive, writeArchive, archiveOptions };
