'use strict';

const JSZip = require('jszip');

/**
 * Test-only helpers that read a produced archive back through the same
 * surface a spreadsheet consumer would use. Nothing here is shipped: these
 * assert on a subset of the format specification, not the whole of it.
 */

const unescapeXml = (s) =>
    s
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, '&');

const COL = (n) => {
    let s = '';
    n += 1;
    while (n > 0) {
        const r = (n - 1) % 26;
        s = String.fromCharCode(65 + r) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
};

/** Parse a produced .xlsx buffer into a plain, assertable shape. */
async function readArchive(buffer) {
    const zip = await JSZip.loadAsync(buffer);
    const parts = Object.keys(zip.files).filter((f) => !zip.files[f].dir);

    const readPart = async (name) => {
        const entry = zip.file(name);
        if (!entry) throw new Error(`part missing from archive: ${name}`);
        return entry.async('string');
    };

    const workbookXml = await readPart('xl/workbook.xml');
    const declaredSheets = [...workbookXml.matchAll(/<sheet\s+name="([^"]*)"[^>]*r:id="([^"]*)"/g)].map((m) => ({
        name: unescapeXml(m[1]),
        relId: m[2],
    }));

    const relsXml = await readPart('xl/_rels/workbook.xml.rels');
    const relationships = new Map(
        [...relsXml.matchAll(/Id="([^"]*)"[^>]*Target="([^"]*)"/g)].map((m) => [m[1], m[2]])
    );

    const sheetData = [];
    for (const sheet of declaredSheets) {
        const target = relationships.get(sheet.relId);
        if (!target) throw new Error(`sheet "${sheet.name}" references unknown ${sheet.relId}`);
        const partName = `xl/${target.replace(/^\/?/, '')}`;
        const xml = await readPart(partName);

        const rows = [];
        for (const rowMatch of xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)) {
            const cells = {};
            for (const cellMatch of rowMatch[1].matchAll(/<c\s+r="([A-Z]+)(\d+)"([^>]*)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
                const [, , , attrs, inner] = cellMatch;
                const isText = /t="inlineStr"/.test(attrs);
                const v = inner === undefined ? undefined : (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
                const t = inner === undefined ? undefined : (inner.match(/<t>([\s\S]*?)<\/t>/) || [])[1];
                cells[COL(cellMatch[1].charCodeAt(0) - 65)] = {
                    raw: v !== undefined ? unescapeXml(v) : undefined,
                    text: t !== undefined ? unescapeXml(t) : undefined,
                    numeric: !isText,
                };
            }
            rows.push(cells);
        }
        sheetData.push({ name: sheet.name, partName, xml, rows });
    }

    return { parts, declaredSheets, relationships, sheetData, zip };
}

/**
 * Assert the invariants a strict consumer checks. Returns a list of
 * violations rather than throwing, so a test can report all of them at once.
 */
async function validateArchive(buffer) {
    const violations = [];
    let archive;
    try {
        archive = await readArchive(buffer);
    } catch (err) {
        return [`archive is unreadable: ${err.message}`];
    }

    const { parts, declaredSheets, relationships, sheetData, zip } = archive;

    for (const sheet of declaredSheets) {
        const target = relationships.get(sheet.relId);
        if (!target) {
            violations.push(`declared sheet "${sheet.name}" has no resolvable relationship ${sheet.relId}`);
            continue;
        }
        const partName = `xl/${target.replace(/^\/?/, '')}`;
        if (!zip.file(partName)) {
            violations.push(`declared sheet "${sheet.name}" resolves to missing part ${partName}`);
        }
    }

    for (const entry of sheetData) {
        for (const m of entry.xml.matchAll(/<v>([^<]*)<\/v>/g)) {
            if (m[1] === 'NaN' || m[1] === 'Infinity' || m[1] === '-Infinity') {
                violations.push(`sheet "${entry.name}" contains non-finite numeric value <v>${m[1]}</v>`);
            }
        }
        for (const m of entry.xml.matchAll(/<t>([\s\S]*?)<\/t>/g)) {
            if (/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(m[1])) {
                violations.push(`sheet "${entry.name}" cell text contains a raw control character`);
            }
        }
        for (const m of entry.xml.matchAll(/<row\b[^>]*\bspans="(\d+):(\d+)"/g)) {
            if (Number(m[1]) > Number(m[2])) {
                violations.push(
                    `sheet "${entry.name}" has a row with an inverted spans range ${m[1]}:${m[2]}`
                );
            }
        }    }

    const ctXml = zip.file('[Content_Types].xml')
        ? await zip.file('[Content_Types].xml').async('string')
        : null;
    if (!ctXml) {
        violations.push('missing [Content_Types].xml');
    } else {
        const defaults = new Set(
            [...ctXml.matchAll(/<Default\s+Extension="([^"]*)"/g)].map((m) => m[1].toLowerCase())
        );
        const overrides = new Set([...ctXml.matchAll(/<Override\s+PartName="([^"]*)"/g)].map((m) => m[1]));
        for (const part of parts) {
            if (part === '[Content_Types].xml') continue;
            const ext = (part.split('.').pop() || '').toLowerCase();
            if (!overrides.has(`/${part}`) && !defaults.has(ext)) {
                violations.push(`part ${part} has no declared content type`);
            }
        }
    }

    return violations;
}

/**
 * Compressible content, so a compression level difference is measurable.
 * High-entropy and deterministic: a long run of identical characters deflates
 * to almost nothing at any level, which hides the difference entirely.
 */
const compressible = (cells, label) => {
    let seed = 0x2f6e2b1;
    const alphabet = 'abcdefghijklmnopqrstuvwxyz0123456789';
    const token = () => {
        let out = '';
        for (let i = 0; i < 64; i++) {
            seed = (seed * 1103515245 + 12345) & 0x7fffffff;
            out += alphabet[seed % alphabet.length];
        }
        return out;
    };
    return Array.from({ length: cells }, (_, i) => [`${label}-${i}-${token()}`, i]);
};

module.exports = { readArchive, validateArchive, compressible, unescapeXml };
