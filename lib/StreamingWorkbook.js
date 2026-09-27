/**
 * xlsx-stream-workbook
 * Streaming Excel workbook writer with multiple worksheets support
 * Memory-efficient for large datasets (100K+ rows)
 *
 * @author Khamidullo Khudoyberdiev
 * @license MIT
 */

const XLSXWriteStream = require('xlsx-write-stream');
const JSZip = require('jszip');
const fs = require('fs');
const fsp = require('fs/promises');
const os = require('os');
const path = require('path');
const { once } = require('events');

const PROGRESS_INTERVAL = 1000;
const RELEASE_MESSAGE =
    'The intermediate files for this workbook were released, so it can no longer be saved or ' +
    'extended. Build a new StreamingWorkbook, or pass cleanupOnSave: false and call cleanup() ' +
    'when you are finished.';

const CONTENT_TYPES = {
    'xl/workbook.xml':
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    'xl/styles.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
    'xl/sharedStrings.xml':
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
};

let instanceCounter = 0;

const columnName = (index) => {
    let name = '';
    let n = index + 1;
    while (n > 0) {
        const remainder = (n - 1) % 26;
        name = String.fromCharCode(65 + remainder) + name;
        n = Math.floor((n - 1) / 26);
    }
    return name;
};

const describe = (value) =>
    value === null ? 'null' : Array.isArray(value) ? 'an array' : `a ${typeof value}`;

/**
 * StreamingWorkbook - Create Excel files with multiple sheets using streaming
 *
 * @example
 * const workbook = new StreamingWorkbook();
 *
 * await workbook.addSheet('Users', ['ID', 'Name', 'Email'], [
 *   [1, 'John', 'john@example.com'],
 *   [2, 'Jane', 'jane@example.com']
 * ]);
 *
 * await workbook.save('output.xlsx');
 */
class StreamingWorkbook {
    /**
     * Create a new StreamingWorkbook instance
     * @param {Object} options - Configuration options
     * @param {string} [options.tempDir] - Temporary directory for intermediate files.
     *   A directory supplied here is never removed by the library.
     * @param {boolean} [options.cleanupOnSave=true] - Delete temp files after save
     * @param {number} [options.compressionLevel=6] - ZIP compression level (0-9)
     */
    constructor(options = {}) {
        const ownsTempDir = options.tempDir === undefined;

        this.options = {
            tempDir: options.tempDir ?? path.join(os.tmpdir(), `xlsx-stream-workbook-${process.pid}`),
            cleanupOnSave: options.cleanupOnSave !== false,
            compressionLevel: options.compressionLevel ?? 6
        };

        if (!Number.isInteger(this.options.compressionLevel) ||
            this.options.compressionLevel < 0 ||
            this.options.compressionLevel > 9) {
            throw new RangeError(
                `compressionLevel must be an integer between 0 and 9, received ${options.compressionLevel}`
            );
        }

        this._ownsTempDir = ownsTempDir;
        this._instanceId = ++instanceCounter;
        this._sheetCounter = 0;
        this._initialised = false;
        this._released = false;
        this._usedNames = new Set();
        this.sheets = [];
        this.tempFiles = [];
    }

    // ==================== PUBLIC API ====================

    /**
     * Add a worksheet with data
     *
     * @param {string} sheetName - Name of the worksheet
     * @param {Array<string>} headers - Column headers
     * @param {Array<Array>|Iterable<Array>} rows - Data rows, array or lazy iterable
     * @param {Object} [options] - Sheet options
     * @param {Function} [options.onProgress] - Progress callback (completed, total).
     *   `total` is undefined when the row count is not knowable up front. Must be synchronous.
     * @param {number} [options.progressInterval=1000] - Rows between progress callbacks
     * @returns {Promise<{rowCount: number, duration: number, name: string, nameChanged: boolean}>}
     */
    async addSheet(sheetName, headers, rows, options = {}) {
        return this._addSheet('array', sheetName, headers, rows, options);
    }

    /**
     * Add a worksheet with sections (vertical multi-section layout)
     *
     * @param {string} sheetName - Name of the worksheet
     * @param {Array<{title?: string, headers: string[], rows: Array<Array>}>} sections
     * @param {Object} [options] - Sheet options
     * @returns {Promise<{rowCount: number, duration: number, name: string, nameChanged: boolean}>}
     */
    async addSheetWithSections(sheetName, sections, options = {}) {
        return this._addSheet('sections', sheetName, null, sections, options);
    }

    /**
     * Add a worksheet from an iterable/generator (for very large datasets)
     *
     * @param {string} sheetName - Name of the worksheet
     * @param {Array<string>} headers - Column headers
     * @param {Iterable<Array>|AsyncIterable<Array>} rowIterator
     * @param {Object} [options] - Sheet options
     * @returns {Promise<{rowCount: number, duration: number, name: string, nameChanged: boolean}>}
     */
    async addSheetFromIterator(sheetName, headers, rowIterator, options = {}) {
        return this._addSheet('iterator', sheetName, headers, rowIterator, options);
    }

    /**
     * Save workbook to file
     *
     * @param {string} outputPath - Output file path. Missing parent directories are created.
     * @returns {Promise<{filePath: string, fileSize: number, sheetCount: number}>}
     */
    async save(outputPath) {
        this._assertUsable();
        this._assertHasSheets();
        const absolutePath = path.resolve(outputPath);
        await fsp.mkdir(path.dirname(absolutePath), { recursive: true });

        const archive = await this._buildArchive();
        await this._writeArchive(archive, absolutePath);
        await this._afterSave();

        return {
            filePath: absolutePath,
            fileSize: fs.statSync(absolutePath).size,
            sheetCount: this._orderedSheets().length
        };
    }

    /**
     * Save workbook to buffer
     *
     * @returns {Promise<Buffer>}
     */
    async saveAsBuffer() {
        this._assertUsable();
        this._assertHasSheets();
        const archive = await this._buildArchive();
        const buffer = await archive.generateAsync({
            type: 'nodebuffer',
            ...this._archiveOptions()
        });
        await this._afterSave();
        return buffer;
    }

    /**
     * Get sheet info
     * @returns {Array<{name: string, rowCount: number}>}
     */
    getSheets() {
        return this._orderedSheets().map((sheet) => ({ name: sheet.name, rowCount: sheet.rowCount }));
    }

    /**
     * Paths of the intermediate files this workbook created. Useful when
     * cleanupOnSave is false; the caller owns their lifetime once returned.
     *
     * @returns {string[]}
     */
    getTempFiles() {
        return [...this.tempFiles];
    }

    /**
     * Release the intermediate files. The workbook cannot be saved or extended afterwards.
     * @returns {Promise<void>}
     */
    async cleanup() {
        await this._cleanup();
    }

    // ==================== SHEET PLUMBING ====================

    async _addSheet(kind, sheetName, headers, payload, options) {
        this._assertUsable();

        // Everything up to the first await must stay synchronous, otherwise two
        // concurrent adds can reserve their slot in whichever order they happen
        // to resume. Sheets must come out in call order, not completion order.
        const { name, changed } = this._resolveSheetName(sheetName);
        const tempFilePath = this._nextTempPath();
        const order = this.sheets.length;
        this.sheets.push(null);
        this._usedNames.add(name.toLowerCase());

        const startTime = Date.now();
        let rowCount;
        try {
            await this._init();
            if (kind === 'sections') {
                rowCount = await this._writeSectioned(tempFilePath, name, payload, options);
            } else {
                rowCount = await this._writeTabular(tempFilePath, name, headers, payload, options, kind);
            }
        } catch (err) {
            this._usedNames.delete(name.toLowerCase());
            await this._discardTempFile(tempFilePath);
            throw err;
        }

        this.sheets[order] = { name, filePath: tempFilePath, rowCount };
        this.tempFiles.push(tempFilePath);

        return { rowCount, duration: Date.now() - startTime, name, nameChanged: changed };
    }

    /** Sheets in the order they were added, skipping any that failed. */
    _orderedSheets() {
        return this.sheets.filter(Boolean);
    }

    async _init() {
        if (!this._initialised) {
            await fsp.mkdir(this.options.tempDir, { recursive: true });
            this._initialised = true;
        }
    }

    _assertUsable() {
        if (this._released) throw new Error(RELEASE_MESSAGE);
    }

    _assertHasSheets() {
        if (this._orderedSheets().length === 0) throw new Error('No sheets added to workbook');
    }

    async _afterSave() {
        if (this.options.cleanupOnSave) await this._cleanup();
    }

    /**
     * Temp file identity never depends on timing or on shared mutable state:
     * a process-wide instance id plus a per-instance counter is unique by
     * construction, so parallel adds cannot land on the same path.
     */
    _nextTempPath() {
        return path.join(this.options.tempDir, `sheet_${this._instanceId}_${this._sheetCounter++}.xlsx`);
    }

    // ==================== WRITERS ====================

    async _writeTabular(filePath, sheetName, headers, rows, options, kind) {
        const { xlsxStream, writeStream, done } = this._openStream(filePath);
        const state = { rowNumber: 0, rowCount: 0, onProgress: options.onProgress,
            interval: this._progressInterval(options), total: undefined, reported: 0 };

        try {
            if (headers !== null) await this._emitRow(xlsxStream, headers, sheetName, state, 'header');

            const knownTotal = Array.isArray(rows) ? rows.length : undefined;
            state.total = knownTotal;

            for await (const row of rows) {
                await this._emitRow(xlsxStream, row, sheetName, state, 'data');
                state.rowCount++;
                this._reportProgress(state);
            }

            this._reportProgress(state, true);
            xlsxStream.end();
            await done;
        } catch (err) {
            xlsxStream.destroy();
            writeStream.destroy();
            throw err;
        }
        return state.rowCount;
    }

    async _writeSectioned(filePath, sheetName, sections, options) {
        if (!Array.isArray(sections)) {
            throw new TypeError(
                `Invalid sections for sheet "${sheetName}": expected an array, received ${describe(sections)}`
            );
        }

        const totals = sections.map((section, index) => {
            if (!section || typeof section !== 'object' || Array.isArray(section)) {
                throw new TypeError(
                    `Invalid section ${index + 1} in sheet "${sheetName}": expected an object, received ${describe(section)}`
                );
            }
            if (section.headers !== undefined && !Array.isArray(section.headers)) {
                throw new TypeError(
                    `Invalid headers in section ${index + 1} of sheet "${sheetName}": expected an array, received ${describe(section.headers)}`
                );
            }
            if (section.rows !== undefined && !Array.isArray(section.rows)) {
                throw new TypeError(
                    `Invalid rows in section ${index + 1} of sheet "${sheetName}": expected an array, received ${describe(section.rows)}`
                );
            }
            return (section.rows || []).length;
        });

        const { xlsxStream, writeStream, done } = this._openStream(filePath);
        const state = { rowNumber: 0, rowCount: 0, onProgress: options.onProgress,
            interval: this._progressInterval(options), total: totals.reduce((a, b) => a + b, 0), reported: 0 };

        try {
            for (let index = 0; index < sections.length; index++) {
                const section = sections[index];
                await this._emitRow(
                    xlsxStream,
                    [section.title || `Section ${index + 1}`],
                    sheetName,
                    state,
                    'title'
                );
                if (section.headers !== undefined) {
                    await this._emitRow(xlsxStream, section.headers, sheetName, state, 'header');
                }
                for (const row of section.rows || []) {
                    await this._emitRow(xlsxStream, row, sheetName, state, 'data');
                    state.rowCount++;
                    this._reportProgress(state);
                }
                if (index < sections.length - 1) {
                    // A single empty string yields a well-formed cell and a valid
                    // spans attribute; an empty array yields spans="1:0", which is not.
                    await this._emitRow(xlsxStream, [''], sheetName, state, 'spacer');
                    await this._emitRow(xlsxStream, [''], sheetName, state, 'spacer');
                }
            }

            this._reportProgress(state, true);
            xlsxStream.end();
            await done;
        } catch (err) {
            xlsxStream.destroy();
            writeStream.destroy();
            throw err;
        }
        return state.rowCount;
    }

    _openStream(filePath) {
        const xlsxStream = new XLSXWriteStream();
        const writeStream = fs.createWriteStream(filePath);
        const done = new Promise((resolve, reject) => {
            writeStream.on('finish', resolve);
            writeStream.on('error', reject);
            xlsxStream.on('error', reject);
        });
        // Nothing else awaits the write stream; keep an unhandled rejection from
        // a late stream error from crashing the process after the fact.
        done.catch(() => {});
        xlsxStream.pipe(writeStream);
        return { xlsxStream, writeStream, done };
    }

    /**
     * Validate, write, and advance the row counter. Honours backpressure so a
     * large export does not accumulate the whole sheet in memory.
     */
    async _emitRow(xlsxStream, row, sheetName, state, kind) {
        const rowNumber = state.rowNumber + 1;
        if (kind === 'data' || kind === 'header') this._validateRow(row, sheetName, rowNumber);
        if (!xlsxStream.write(row)) await once(xlsxStream, 'drain');
        state.rowNumber = rowNumber;
    }

    _validateRow(row, sheetName, rowNumber) {
        if (!Array.isArray(row)) {
            throw new TypeError(
                `Invalid row ${rowNumber} in sheet "${sheetName}": expected an array, received ${describe(row)}`
            );
        }
        for (let index = 0; index < row.length; index++) {
            const value = row[index];
            if (value === null || value === undefined) continue;
            const type = typeof value;
            if (type === 'number' && !Number.isFinite(value)) {
                throw new TypeError(
                    `Invalid value at sheet "${sheetName}", row ${rowNumber}, column ${columnName(index)}: ` +
                        `${value} cannot be written to an Excel file`
                );
            }
            if (type !== 'string' && type !== 'number' && type !== 'boolean' && !(value instanceof Date)) {
                throw new TypeError(
                    `Invalid value at sheet "${sheetName}", row ${rowNumber}, column ${columnName(index)}: ` +
                        `a ${type} cannot be written to an Excel file`
                );
            }
        }
    }

    _progressInterval(options) {
        const interval = options.progressInterval ?? PROGRESS_INTERVAL;
        if (!Number.isInteger(interval) || interval < 1) {
            throw new RangeError(
                `progressInterval must be a positive integer, received ${options.progressInterval}`
            );
        }
        return interval;
    }

    _reportProgress(state, final = false) {
        const { onProgress, interval, total } = state;
        if (!onProgress || state.rowCount === 0) return;
        if (state.rowCount === state.reported) return;

        const onBoundary = state.rowCount % interval === 0;
        if (total !== undefined) {
            if (onBoundary || final) {
                state.reported = state.rowCount;
                onProgress(state.rowCount, total);
            }
        } else if (onBoundary || state.rowCount === 1 || final) {
            state.reported = state.rowCount;
            onProgress(state.rowCount);
        }
    }

    // ==================== ARCHIVE ====================

    async _buildArchive() {
        const sheets = this._orderedSheets();
        const baseZip = await JSZip.loadAsync(await fsp.readFile(sheets[0].filePath));
        const outputZip = new JSZip();
        const carried = new Set();

        for (const fileName of Object.keys(baseZip.files)) {
            const entry = baseZip.files[fileName];
            if (entry.dir) continue;
            if (fileName.startsWith('xl/worksheets/') || fileName.includes('workbook.xml')) continue;
            if (fileName === '[Content_Types].xml' || fileName === '_rels/.rels') continue;
            if (fileName === 'xl/sharedStrings.xml') {
                throw new Error(
                    'The underlying writer produced a shared strings table, whose per-sheet indices ' +
                        'cannot be merged. Refusing to write a workbook that Excel would have to repair.'
                );
            }
            carried.add(fileName);
            outputZip.file(fileName, await entry.async('nodebuffer'));
        }

        const sheetEntries = [];
        const sheetRels = [];
        const overrides = [];

        for (let index = 0; index < sheets.length; index++) {
            const { name, filePath } = sheets[index];
            const sheetId = index + 1;
            const sheetFileName = `sheet${sheetId}.xml`;

            const sheetZip = await JSZip.loadAsync(await fsp.readFile(filePath));
            const sheetXmlFile = Object.keys(sheetZip.files).find(
                (entry) => entry.startsWith('xl/worksheets/') && entry.endsWith('.xml')
            );
            if (!sheetXmlFile) {
                throw new Error(
                    `Sheet "${name}" produced no worksheet XML, so the workbook would declare a ` +
                        'sheet that resolves to nothing. Refusing to write it.'
                );
            }
            if (sheetZip.files['xl/sharedStrings.xml']) {
                throw new Error(
                    'The underlying writer produced a shared strings table, whose per-sheet indices ' +
                        'cannot be merged. Refusing to write a workbook that Excel would have to repair.'
                );
            }

            outputZip.file(`xl/worksheets/${sheetFileName}`, await sheetZip.files[sheetXmlFile].async('nodebuffer'));
            sheetEntries.push(
                `<sheet name="${this._escapeXml(name)}" sheetId="${sheetId}" r:id="rId${sheetId}"/>`
            );
            sheetRels.push(
                `<Relationship Id="rId${sheetId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${sheetFileName}"/>`
            );
            overrides.push(`xl/worksheets/${sheetFileName}`);
        }

        outputZip.file(
            'xl/workbook.xml',
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>${sheetEntries.join('')}</sheets>
</workbook>`
        );

        const stylesRelId = `rId${sheets.length + 1}`;
        outputZip.file(
            'xl/_rels/workbook.xml.rels',
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheetRels.join('\n')}
<Relationship Id="${stylesRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
        );

        outputZip.file('[Content_Types].xml', this._contentTypes(overrides, carried));
        outputZip.file(
            '_rels/.rels',
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
        );

        return outputZip;
    }

    /**
     * Derived from the parts actually shipped. A part with no known content
     * type is a hard error rather than a generic guess: a wrong content type is
     * exactly what makes Excel offer to repair the file.
     */
    _contentTypes(worksheetParts, carriedParts) {
        const overrides = [
            ['xl/workbook.xml', CONTENT_TYPES['xl/workbook.xml']],
            ...worksheetParts.map(
                (part) =>
                    [
                        part,
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
                    ]
            )
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

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
${overrides.map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join('\n')}
</Types>`;
    }

    /**
     * jszip@3.10 folds `level: 0` into its own default (`level || -1` in
     * lib/flate.js), so a DEFLATE request can never express "no compression".
     * Store mode can, and is what a caller asking for level 0 actually wants.
     */
    _archiveOptions() {
        const level = this.options.compressionLevel;
        if (level === 0) return { compression: 'STORE' };
        return { compression: 'DEFLATE', compressionOptions: { level } };
    }

    /** Stream the archive to disk rather than materialising a second copy. */
    _writeArchive(archive, outputPath) {
        return new Promise((resolve, reject) => {
            const out = fs.createWriteStream(outputPath);
            out.on('finish', resolve);
            out.on('error', reject);
            archive
                .generateNodeStream({ type: 'nodebuffer', ...this._archiveOptions() })
                .on('error', reject)
                .pipe(out);
        });
    }

    // ==================== NAMES AND ESCAPING ====================

    /**
     * Normalise against the rules Excel enforces, then guarantee the name is
     * not already taken. Excel compares sheet names case-insensitively, and two
     * names that differ only by an illegal character are the same name.
     *
     * @returns {{name: string, changed: boolean}}
     * @private
     */
    _resolveSheetName(name) {
        const normalised = this._sanitizeSheetName(name);
        const changed = normalised !== String(name);

        if (!this._usedNames.has(normalised.toLowerCase())) {
            return { name: normalised, changed };
        }

        for (let suffix = 2; ; suffix++) {
            const tail = `_${suffix}`;
            const candidate = `${normalised.slice(0, 31 - tail.length)}${tail}`;
            if (!this._usedNames.has(candidate.toLowerCase())) {
                return { name: candidate, changed: true };
            }
        }
    }

    _sanitizeSheetName(name) {
        let safe = String(name)
            .replace(/[\u0000-\u001f]/g, '')
            .replace(/[\\\/?*:[\]]/g, '_')
            .trim();

        if (safe.startsWith("'")) safe = `_${safe.slice(1)}`;
        if (safe.endsWith("'")) safe = `${safe.slice(0, -1)}_`;

        safe = safe.substring(0, 31);
        if (safe.startsWith("'")) safe = `_${safe.slice(1)}`;
        if (safe.endsWith("'")) safe = `${safe.slice(0, -1)}_`;

        return safe || 'Sheet';
    }

    _escapeXml(str) {
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    // ==================== CLEANUP ====================

    async _discardTempFile(filePath) {
        try {
            await fsp.unlink(filePath);
        } catch {
            // The file may never have been created; nothing to remove.
        }
    }

    async _cleanup() {
        const files = [...this.tempFiles];
        for (const filePath of files) {
            try {
                await fsp.unlink(filePath);
            } catch {
                // A locked or already-removed file must not mask the caller's intent.
            }
        }
        this.tempFiles = [];
        this._released = true;

        if (this._ownsTempDir) {
            try {
                const remaining = await fsp.readdir(this.options.tempDir);
                if (remaining.length === 0) await fsp.rmdir(this.options.tempDir);
            } catch {
                // Not ours to remove, or already gone.
            }
        }
    }
}

module.exports = StreamingWorkbook;
