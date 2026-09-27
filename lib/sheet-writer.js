'use strict';

/**
 * Writes one sheet's rows into a temp .xlsx file.
 *
 * Rows are validated as they are written rather than in a pass beforehand, so a
 * value that cannot be represented fails on the row that carries it instead of
 * naming a position the caller has to map back to their own data.
 */

const XLSXWriteStream = require('xlsx-write-stream');
const fs = require('fs');
const { once } = require('events');

const PROGRESS_INTERVAL = 1000;
const WRITABLE_TYPES = new Set(['string', 'number', 'boolean']);

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

const openStream = (filePath) => {
    const xlsxStream = new XLSXWriteStream();
    const writeStream = fs.createWriteStream(filePath);
    const done = new Promise((resolve, reject) => {
        writeStream.on('finish', resolve);
        writeStream.on('error', reject);
        xlsxStream.on('error', reject);
    });
    // The caller awaits `done` only on the success path, so a late stream error
    // would otherwise surface as an unhandled rejection after the fact.
    done.catch(() => {});
    xlsxStream.pipe(writeStream);
    return { xlsxStream, writeStream, done };
};

const abort = (xlsxStream, writeStream) => {
    xlsxStream.destroy();
    writeStream.destroy();
};

const validateRow = (row, sheetName, rowNumber) => {
    if (!Array.isArray(row)) {
        throw new TypeError(
            `Invalid row ${rowNumber} in sheet "${sheetName}": expected an array, received ${describe(row)}`
        );
    }
    for (let index = 0; index < row.length; index++) {
        const value = row[index];
        if (value === null || value === undefined) continue;

        const type = typeof value;
        const at = `sheet "${sheetName}", row ${rowNumber}, column ${columnName(index)}`;
        if (type === 'number' && !Number.isFinite(value)) {
            throw new TypeError(`Invalid value at ${at}: ${value} cannot be written to an Excel file`);
        }
        if (!WRITABLE_TYPES.has(type) && !(value instanceof Date)) {
            throw new TypeError(`Invalid value at ${at}: a ${type} cannot be written to an Excel file`);
        }
    }
};

const emitRow = async (xlsxStream, row, sheetName, state, kind) => {
    const rowNumber = state.rowNumber + 1;
    if (kind === 'data' || kind === 'header') validateRow(row, sheetName, rowNumber);
    if (!xlsxStream.write(row)) await once(xlsxStream, 'drain');
    state.rowNumber = rowNumber;
};

const progressInterval = (options) => {
    const interval = options.progressInterval ?? PROGRESS_INTERVAL;
    if (!Number.isInteger(interval) || interval < 1) {
        throw new RangeError(
            `progressInterval must be a positive integer, received ${options.progressInterval}`
        );
    }
    return interval;
};

const reportProgress = (state, final = false) => {
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
};

const newState = (options, total) => ({
    rowNumber: 0,
    rowCount: 0,
    reported: 0,
    total,
    interval: progressInterval(options),
    onProgress: options.onProgress
});

const writeTabular = async (filePath, sheetName, headers, rows, options) => {
    const { xlsxStream, writeStream, done } = openStream(filePath);
    const state = newState(options, Array.isArray(rows) ? rows.length : undefined);

    try {
        if (headers !== null) await emitRow(xlsxStream, headers, sheetName, state, 'header');

        for await (const row of rows) {
            await emitRow(xlsxStream, row, sheetName, state, 'data');
            state.rowCount++;
            reportProgress(state);
        }

        reportProgress(state, true);
        xlsxStream.end();
        await done;
    } catch (err) {
        abort(xlsxStream, writeStream);
        throw err;
    }
    return state.rowCount;
};

const assertSections = (sections, sheetName) => {
    if (!Array.isArray(sections)) {
        throw new TypeError(
            `Invalid sections for sheet "${sheetName}": expected an array, received ${describe(sections)}`
        );
    }
    sections.forEach((section, index) => {
        const at = `section ${index + 1} of sheet "${sheetName}"`;
        if (!section || typeof section !== 'object' || Array.isArray(section)) {
            throw new TypeError(`Invalid ${at}: expected an object, received ${describe(section)}`);
        }
        for (const field of ['headers', 'rows']) {
            if (section[field] !== undefined && !Array.isArray(section[field])) {
                throw new TypeError(
                    `Invalid ${field} in ${at}: expected an array, received ${describe(section[field])}`
                );
            }
        }
    });
};

const writeSection = async (xlsxStream, section, index, sheetName, state, isLast) => {
    await emitRow(xlsxStream, [section.title || `Section ${index + 1}`], sheetName, state, 'title');
    if (section.headers !== undefined) {
        await emitRow(xlsxStream, section.headers, sheetName, state, 'header');
    }
    for (const row of section.rows || []) {
        await emitRow(xlsxStream, row, sheetName, state, 'data');
        state.rowCount++;
        reportProgress(state);
    }
    if (!isLast) {
        // An empty array makes xlsx-write-stream emit spans="1:0". ECMA-376 puts no
        // constraint on spans, so this is an observed Excel repair trigger, not a spec rule.
        await emitRow(xlsxStream, [''], sheetName, state, 'spacer');
        await emitRow(xlsxStream, [''], sheetName, state, 'spacer');
    }
};

const writeSectioned = async (filePath, sheetName, sections, options) => {
    assertSections(sections, sheetName);
    const total = sections.reduce((sum, section) => sum + (section.rows || []).length, 0);
    const { xlsxStream, writeStream, done } = openStream(filePath);
    const state = newState(options, total);

    try {
        for (let index = 0; index < sections.length; index++) {
            await writeSection(
                xlsxStream,
                sections[index],
                index,
                sheetName,
                state,
                index === sections.length - 1
            );
        }

        reportProgress(state, true);
        xlsxStream.end();
        await done;
    } catch (err) {
        abort(xlsxStream, writeStream);
        throw err;
    }
    return state.rowCount;
};

module.exports = { writeTabular, writeSectioned };
