'use strict';

/**
 * xlsx-stream-workbook
 * Streaming Excel workbook writer with multiple worksheets support
 * Memory-efficient for large datasets (100K+ rows)
 *
 * @author Khamidullo Khudoyberdiev
 * @license MIT
 */

const fs = require('fs');
const fsp = require('fs/promises');
const os = require('os');
const path = require('path');

const { buildArchive, writeArchive, archiveOptions } = require('./archive');
const { writeTabular, writeSectioned } = require('./sheet-writer');

const RELEASE_MESSAGE =
    'The intermediate files for this workbook were released, so it can no longer be saved or ' +
    'extended. Build a new StreamingWorkbook, or pass cleanupOnSave: false and call cleanup() ' +
    'when you are finished.';

let instanceCounter = 0;

/**
 * Create Excel files with multiple sheets. Sheets are written to intermediate
 * files as they are added, then merged into one package at save time.
 *
 * @example
 * const workbook = new StreamingWorkbook();
 *
 * await workbook.addSheet('Users', ['ID', 'Name'], [[1, 'John'], [2, 'Jane']]);
 *
 * await workbook.save('output.xlsx');
 */
class StreamingWorkbook {
    /**
     * @param {Object} [options]
     * @param {string} [options.tempDir] - Where intermediate files go. Defaults to a
     *   directory under the OS temp location. A directory supplied here is never
     *   removed by this library.
     * @param {boolean} [options.cleanupOnSave=true] - Release the intermediate
     *   files after a successful save. With this on, a workbook can only be saved once.
     * @param {number} [options.compressionLevel=6] - ZIP level 0-9, where 0 stores
     *   without compressing.
     * @throws {RangeError} if compressionLevel is not an integer between 0 and 9
     */
    constructor(options = {}) {
        this.options = {
            tempDir: options.tempDir ?? path.join(os.tmpdir(), `xlsx-stream-workbook-${process.pid}`),
            cleanupOnSave: options.cleanupOnSave !== false,
            compressionLevel: options.compressionLevel ?? 6
        };

        const level = this.options.compressionLevel;
        if (!Number.isInteger(level) || level < 0 || level > 9) {
            throw new RangeError(`compressionLevel must be an integer between 0 and 9, received ${level}`);
        }

        this._ownsTempDir = options.tempDir === undefined;
        this._instanceId = ++instanceCounter;
        this._sheetCounter = 0;
        this._initialised = false;
        this._released = false;
        this._usedNames = new Set();
        this.sheets = [];
        this.tempFiles = [];
    }

    /**
     * @param {string} sheetName
     * @param {Array<string>} headers
     * @param {Array<Array>|Iterable<Array>} rows - A plain array, or any iterable.
     *   Iterables are consumed lazily and never materialised, so the row count is
     *   only knowable up front for an array.
     * @param {Object} [options] - `onProgress(completed, total)`, called at most
     *   every `progressInterval` rows and always at least once. `total` is omitted
     *   when the row count is not knowable. The callback must be synchronous.
     * @throws {TypeError} if a row is not an array, or a value cannot be written
     *   to an Excel file. The message names the sheet, row and column.
     */
    addSheet(sheetName, headers, rows, options = {}) {
        return this._addSheet('tabular', sheetName, headers, rows, options);
    }

    /**
     * Write several titled sections into one sheet, stacked vertically. `rowCount`
     * reports data rows only, not titles, headers or the spacers between sections.
     */
    addSheetWithSections(sheetName, sections, options = {}) {
        return this._addSheet('sections', sheetName, null, sections, options);
    }

    /**
     * Write rows from an iterable or async iterable. Prefer this over `addSheet`
     * for data that arrives as a stream.
     */
    addSheetFromIterator(sheetName, headers, rowIterator, options = {}) {
        return this._addSheet('tabular', sheetName, headers, rowIterator, options);
    }

    /** Missing parent directories are created. */
    async save(outputPath) {
        this._assertUsable();
        this._assertHasSheets();
        const absolutePath = path.resolve(outputPath);
        await fsp.mkdir(path.dirname(absolutePath), { recursive: true });

        const archive = await buildArchive(this._orderedSheets());
        await writeArchive(archive, absolutePath, this.options.compressionLevel);
        await this._afterSave();

        return {
            filePath: absolutePath,
            fileSize: fs.statSync(absolutePath).size,
            sheetCount: this.sheets.length
        };
    }

    async saveAsBuffer() {
        this._assertUsable();
        this._assertHasSheets();
        const buffer = await (await buildArchive(this._orderedSheets())).generateAsync({
            type: 'nodebuffer',
            ...archiveOptions(this.options.compressionLevel)
        });
        await this._afterSave();
        return buffer;
    }

    /** `rowCount` is data rows, so counts are comparable across the add methods. */
    getSheets() {
        return this._orderedSheets().map((sheet) => ({ name: sheet.name, rowCount: sheet.rowCount }));
    }

    /**
     * The intermediate files this workbook created. Reach for this when
     * `cleanupOnSave` is false; the caller owns their lifetime once returned.
     */
    getTempFiles() {
        return [...this.tempFiles];
    }

    /** Releases the intermediate files. The workbook cannot be saved or extended afterwards. */
    async cleanup() {
        await this._cleanup();
    }

    async _addSheet(kind, sheetName, headers, payload, options) {
        this._assertUsable();

        // Everything up to the first await stays synchronous, or concurrent adds can
        // reserve their slot in resume order and the sheets come out shuffled.
        const { name, changed } = this._resolveSheetName(sheetName);
        const tempFilePath = this._nextTempPath();
        const order = this.sheets.length;
        this.sheets.push(null);
        this._usedNames.add(name.toLowerCase());

        const startTime = Date.now();
        let rowCount;
        try {
            await this._init();
            rowCount = kind === 'sections'
                ? await writeSectioned(tempFilePath, name, payload, options)
                : await writeTabular(tempFilePath, name, headers, payload, options);
        } catch (err) {
            this._usedNames.delete(name.toLowerCase());
            await this._discardTempFile(tempFilePath);
            throw err;
        }

        this.sheets[order] = { name, filePath: tempFilePath, rowCount };
        this.tempFiles.push(tempFilePath);

        return { rowCount, duration: Date.now() - startTime, name, nameChanged: changed };
    }

    /** In the order they were added, skipping any that failed. */
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
        if (this.sheets.length === 0) throw new Error('No sheets added to workbook');
    }

    async _afterSave() {
        if (this.options.cleanupOnSave) await this._cleanup();
    }

    // Identity is a process-wide instance id plus a per-instance counter, never a
    // timestamp: parallel adds collide structurally if the name depends on timing.
    _nextTempPath() {
        return path.join(this.options.tempDir, `sheet_${this._instanceId}_${this._sheetCounter++}.xlsx`);
    }

    /**
     * Excel compares sheet names case-insensitively, and two names differing only
     * by an illegal character are the same name. So normalising is not enough: the
     * result still has to be unique across the workbook.
     */
    _resolveSheetName(name) {
        const normalised = this._sanitizeSheetName(name);
        const changed = normalised !== String(name);

        if (!this._usedNames.has(normalised.toLowerCase())) return { name: normalised, changed };

        for (let suffix = 2; ; suffix++) {
            const tail = `_${suffix}`;
            const candidate = `${normalised.slice(0, 31 - tail.length)}${tail}`;
            if (!this._usedNames.has(candidate.toLowerCase())) return { name: candidate, changed: true };
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

    async _discardTempFile(filePath) {
        try {
            await fsp.unlink(filePath);
        } catch {
            // The file may never have been created; nothing to remove.
        }
    }

    async _cleanup() {
        for (const filePath of [...this.tempFiles]) {
            try {
                await fsp.unlink(filePath);
            } catch {
                // A locked or already-removed file must not mask the caller's intent.
            }
        }
        this.tempFiles = [];
        this._released = true;

        // Only ever remove a directory this library created, and only when empty.
        if (!this._ownsTempDir) return;
        try {
            if ((await fsp.readdir(this.options.tempDir)).length === 0) {
                await fsp.rmdir(this.options.tempDir);
            }
        } catch {
            // Already gone, or not ours after all.
        }
    }
}

module.exports = StreamingWorkbook;
