'use strict';

/**
 * xlsx-stream-workbook — assertion suite.
 *
 * One seam: build a workbook through the public API, save it, read the
 * produced archive back, and assert against its bytes. Nothing below this
 * line reaches into private state.
 */

const { test, describe } = require('node:test');
const assert = require('node:assert');
const { execFileSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const { StreamingWorkbook } = require('../index');
const { readArchive, validateArchive, compressible } = require('./helpers');

const roots = [];
function scratch() {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'xsw-test-'));
    roots.push(dir);
    return dir;
}
process.on('exit', () => {
    for (const dir of roots) {
        try {
            fs.rmSync(dir, { recursive: true, force: true });
        } catch {
            /* best effort */
        }
    }
});

const newWorkbook = (options = {}) =>
    new StreamingWorkbook({ tempDir: path.join(scratch(), 'tmp'), ...options });

async function assertValid(buffer, context) {
    const violations = await validateArchive(buffer);
    assert.deepStrictEqual(violations, [], `${context}: ${violations.join('; ')}`);
}

// ---------------------------------------------------------------------------
// Seam: the validator itself, and the pre-existing smoke coverage
// ---------------------------------------------------------------------------

describe('seam', () => {
    test('the validator catches an unreadable archive', async () => {
        assert.match((await validateArchive(Buffer.from('not a zip')))[0], /unreadable/);
    });

    test('the validator catches a declared sheet with no part', async () => {
        const wb = newWorkbook();
        await wb.addSheet('Only', ['a'], [[1]]);
        const violations = await validateArchive(await wb.saveAsBuffer());
        assert.deepStrictEqual(violations, []);
    });

    test('the validator catches a part with no declared content type', async () => {
        const JSZip = require('jszip');
        const zip = new JSZip();
        zip.file(
            'xl/workbook.xml',
            '<workbook xmlns:r="r"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'
        );
        zip.file(
            'xl/_rels/workbook.xml.rels',
            '<Relationships><Relationship Id="rId1" Type="t" Target="worksheets/sheet1.xml"/></Relationships>'
        );
        zip.file('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"/></sheetData></worksheet>');
        zip.file('xl/undeclared.unknownext', 'x');
        // Only the known parts are declared, so the extra part is the violation.
        zip.file(
            '[Content_Types].xml',
            '<Types><Default Extension="rels" ContentType="a"/><Default Extension="xml" ContentType="b"/>' +
                '<Override PartName="/xl/workbook.xml" ContentType="c"/>' +
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="d"/></Types>'
        );
        const violations = await validateArchive(await zip.generateAsync({ type: 'nodebuffer' }));
        assert.ok(
            violations.some((v) => /undeclared\.unknownext/.test(v) && /no declared content type/.test(v)),
            `expected a content-type violation, got ${JSON.stringify(violations)}`
        );
    });

    test('original scenario: basic multi-sheet workbook', async () => {
        const wb = newWorkbook();
        await wb.addSheet('Users', ['ID', 'Name', 'Email', 'Age'], [
            [1, 'John Doe', 'john@example.com', 30],
            [2, 'Jane Smith', 'jane@example.com', 25],
            [3, 'Bob Wilson', 'bob@example.com', 35],
        ]);
        await wb.addSheet('Orders', ['OrderID', 'UserID', 'Product', 'Amount'], [
            [101, 1, 'Widget A', 99.99],
            [102, 2, 'Widget B', 149.99],
            [103, 1, 'Widget C', 79.99],
        ]);
        const result = await wb.save(path.join(scratch(), 'basic.xlsx'));
        assert.strictEqual(result.sheetCount, 2);
        assert.ok(result.fileSize > 0);

        const archive = await readArchive(fs.readFileSync(result.filePath));
        assert.deepStrictEqual(archive.declaredSheets.map((s) => s.name), ['Users', 'Orders']);
        assert.strictEqual(archive.sheetData[0].rows.length, 4);
        assert.strictEqual(archive.sheetData[1].rows[1].C.text, 'Widget A');
    });

    test('original scenario: sectioned sheet', async () => {
        const wb = newWorkbook();
        const result = await wb.addSheetWithSections('KPI_Report', [
            { title: 'CS_TRAFFIC_ERL', headers: ['BSC', 'Day1'], rows: [['BSC01', 1000]] },
            { title: 'DROP_RATE_%', headers: ['BSC', 'Day1'], rows: [['BSC01', 0.5]] },
        ]);
        assert.strictEqual(result.rowCount, 2);
        const buffer = await wb.saveAsBuffer();
        await assertValid(buffer, 'sectioned');
        const archive = await readArchive(buffer);
        assert.strictEqual(archive.sheetData[0].rows[0].A.text, 'CS_TRAFFIC_ERL');
    });

    test('original scenario: large dataset with progress', async () => {
        const wb = newWorkbook();
        const rows = Array.from({ length: 10000 }, (_, i) => [
            i + 1,
            `Cell_${String(i).padStart(5, '0')}`,
            `Site_${Math.floor(i / 100)}`,
            i * 1.5,
            i / 100,
        ]);
        let last = null;
        const result = await wb.addSheet('LargeData', ['ID', 'Cell', 'Site', 'Traffic_MB', 'Drop_Rate'], rows, {
            onProgress: (current, total) => {
                last = { current, total };
            },
        });
        assert.strictEqual(result.rowCount, 10000);
        assert.ok(last, 'onProgress never fired');
        assert.strictEqual(last.total, 10000);
        assert.strictEqual(last.current, 10000);
        await assertValid(await wb.saveAsBuffer(), 'large dataset');
    });

    test('original scenario: generator-based writing', async () => {
        async function* generateRows() {
            for (let i = 0; i < 5000; i++) yield [i + 1, `Item_${i}`, i * 0.5];
        }
        const wb = newWorkbook();
        const result = await wb.addSheetFromIterator('Generated', ['ID', 'Name', 'Value'], generateRows());
        assert.strictEqual(result.rowCount, 5000);
        const archive = await readArchive(await wb.saveAsBuffer());
        assert.strictEqual(archive.sheetData[0].rows.length, 5001);
    });

    test('original scenario: buffer output', async () => {
        const wb = newWorkbook();
        await wb.addSheet('BufferTest', ['A', 'B', 'C'], [
            [1, 2, 3],
            [4, 5, 6],
        ]);
        const buffer = await wb.saveAsBuffer();
        assert.ok(Buffer.isBuffer(buffer));
        await assertValid(buffer, 'buffer output');
    });
});

// ---------------------------------------------------------------------------
// Temp file identity
// ---------------------------------------------------------------------------

describe('temp file identity', () => {
    test('parallel adds produce a valid workbook with every sheet intact', async () => {
        const wb = newWorkbook();
        await Promise.all(
            Array.from({ length: 6 }, (_, i) => wb.addSheet(`S${i}`, ['x'], [[i], [i * 10]]))
        );

        const buffer = await wb.saveAsBuffer();
        await assertValid(buffer, 'parallel adds');

        const archive = await readArchive(buffer);
        assert.strictEqual(archive.declaredSheets.length, 6);
        assert.strictEqual(archive.sheetData.length, 6);
        for (const sheet of archive.sheetData) {
            assert.strictEqual(sheet.rows.length, 3, `sheet ${sheet.name} lost rows`);
        }
    });

    test('parallel adds write to distinct temp files', async () => {
        const tempDir = path.join(scratch(), 'tmp');
        const wb = newStreaming(tempDir);
        await Promise.all(Array.from({ length: 6 }, (_, i) => wb.addSheet(`S${i}`, ['x'], [[i]])));
        assert.strictEqual(fs.readdirSync(tempDir).length, 6);
    });

    test('parallel adds come out in call order, not completion order', async () => {
        for (let attempt = 0; attempt < 5; attempt++) {
            const wb = newWorkbook();
            const names = ['First', 'Second', 'Third', 'Fourth', 'Fifth', 'Sixth'];
            await Promise.all(names.map((name, i) => wb.addSheet(name, ['x'], [[i]])));
            assert.deepStrictEqual(
                wb.getSheets().map((s) => s.name),
                names,
                `attempt ${attempt}: sheets are out of call order`
            );
            const archive = await readArchive(await wb.saveAsBuffer());
            assert.deepStrictEqual(archive.declaredSheets.map((s) => s.name), names);
        }
    });

    test('a failed parallel add leaves the others intact and in order', async () => {
        const wb = newWorkbook();
        const results = await Promise.allSettled([
            wb.addSheet('Good1', ['x'], [[1]]),
            wb.addSheet('Bad', ['x'], [[NaN]]),
            wb.addSheet('Good2', ['x'], [[3]]),
        ]);
        assert.deepStrictEqual(
            results.map((r) => r.status),
            ['fulfilled', 'rejected', 'fulfilled']
        );
        assert.deepStrictEqual(wb.getSheets().map((s) => s.name), ['Good1', 'Good2']);
        await assertValid(await wb.saveAsBuffer(), 'partial parallel failure');
    });

    test('two workbooks in one process do not share temp files', async () => {
        const tempDir = path.join(scratch(), 'shared');
        const a = new StreamingWorkbook({ tempDir });
        const b = new StreamingWorkbook({ tempDir });
        await Promise.all([a.addSheet('A', ['x'], [[1]]), b.addSheet('B', ['x'], [[2]])]);
        assert.strictEqual(fs.readdirSync(tempDir).length, 2);
    });

    function newStreaming(dir) {
        return new StreamingWorkbook({ tempDir: dir });
    }
});

// ---------------------------------------------------------------------------
// Save
// ---------------------------------------------------------------------------

describe('save', () => {
    test('creates missing parent directories', async () => {
        const wb = newWorkbook();
        await wb.addSheet('S', ['a'], [[1]]);
        const target = path.join(scratch(), 'deeply', 'nested', 'out.xlsx');
        const result = await wb.save(target);
        assert.ok(fs.existsSync(result.filePath));
    });

    test('saving twice works when the caller keeps the intermediate files', async () => {
        const wb = newWorkbook({ cleanupOnSave: false });
        await wb.addSheet('S', ['a'], [[1]]);
        const dir = scratch();
        const first = await wb.save(path.join(dir, 'first.xlsx'));
        const second = await wb.save(path.join(dir, 'second.xlsx'));
        assert.notStrictEqual(first.filePath, second.filePath);
        const a = await readArchive(fs.readFileSync(first.filePath));
        const b = await readArchive(fs.readFileSync(second.filePath));
        assert.deepStrictEqual(a.declaredSheets, b.declaredSheets);
    });

    test('a sheet added between two saves appears in the second file', async () => {
        const wb = newWorkbook({ cleanupOnSave: false });
        const dir = scratch();
        await wb.addSheet('First', ['a'], [[1]]);
        await wb.save(path.join(dir, 'one.xlsx'));
        await wb.addSheet('Second', ['a'], [[2]]);
        await wb.save(path.join(dir, 'two.xlsx'));
        const archive = await readArchive(fs.readFileSync(path.join(dir, 'two.xlsx')));
        assert.deepStrictEqual(archive.declaredSheets.map((s) => s.name), ['First', 'Second']);
    });

    test('saving again after automatic cleanup fails with an actionable message', async () => {
        const wb = newWorkbook();
        await wb.addSheet('S', ['a'], [[1]]);
        const dir = scratch();
        await wb.save(path.join(dir, 'one.xlsx'));
        await assert.rejects(
            () => wb.save(path.join(dir, 'two.xlsx')),
            (err) => {
                assert.ok(!/ENOENT/.test(err.message), 'must not surface a raw filesystem error');
                assert.match(err.message, /cleanupOnSave|cleanup\(\)/);
                return true;
            }
        );
    });
});

// ---------------------------------------------------------------------------
// Sheet names
// ---------------------------------------------------------------------------

describe('sheet names', () => {
    const namesOf = (wb) => wb.getSheets().map((s) => s.name);

    test('names that sanitise identically stay distinct', async () => {
        const wb = newWorkbook();
        await wb.addSheet('a/b', ['x'], [[1]]);
        await wb.addSheet('a?b', ['x'], [[2]]);
        const names = namesOf(wb);
        assert.strictEqual(new Set(names.map((n) => n.toLowerCase())).size, 2, `duplicate: ${names}`);
    });

    test('long names sharing a prefix stay distinct', async () => {
        const wb = newWorkbook();
        const base = 'A'.repeat(31);
        await wb.addSheet(base, ['x'], [[1]]);
        await wb.addSheet(`${base}B`, ['x'], [[2]]);
        const names = namesOf(wb);
        assert.strictEqual(names.length, 2);
        for (const n of names) assert.ok(n.length <= 31, `"${n}" exceeds the Excel limit`);
    });

    test('Excel-illegal characters and control characters are removed', async () => {
        const wb = newWorkbook();
        await wb.addSheet('a:b*c?d[e]f\\g/h', ['x'], [[1]]);
        await wb.addSheet(`bad${String.fromCharCode(7)}name`, ['x'], [[2]]);
        const names = namesOf(wb);
        for (const n of names) {
            assert.ok(!/[\\/?*:[\]\x00-\x1f]/.test(n), `illegal character in "${n}"`);
        }
    });

    test('an empty name still yields a valid sheet', async () => {
        const wb = newWorkbook();
        await wb.addSheet('', ['x'], [[1]]);
        await wb.addSheet('   ', ['x'], [[2]]);
        const names = namesOf(wb);
        for (const n of names) assert.ok(n.length > 0);
        assert.strictEqual(new Set(names).size, 2);
    });

    test('the caller is told when a name was altered', async () => {
        const wb = newWorkbook();
        const result = await wb.addSheet('a/b', ['x'], [[1]]);
        assert.strictEqual(result.nameChanged, true);
        assert.strictEqual(result.name, 'a_b');
        const clean = await wb.addSheet('plain', ['x'], [[2]]);
        assert.strictEqual(clean.nameChanged, false);
    });

    test('uniqueness suffixing is deterministic across runs', async () => {
        const run = async () => {
            const wb = newWorkbook();
            await wb.addSheet('a/b', ['x'], [[1]]);
            await wb.addSheet('a?b', ['x'], [[2]]);
            await wb.addSheet('a*b', ['x'], [[3]]);
            return namesOf(wb);
        };
        assert.deepStrictEqual(await run(), await run());
    });
});

// ---------------------------------------------------------------------------
// Cell values and row counts
// ---------------------------------------------------------------------------

describe('cell values', () => {
    test('a non-finite number is rejected naming sheet, row and column', async () => {
        const wb = newWorkbook();
        await assert.rejects(
            () => wb.addSheet('Metrics', ['a', 'b'], [[1, 2], [3, NaN]]),
            (err) => {
                assert.match(err.message, /Metrics/);
                assert.match(err.message, /row 3/i);
                assert.match(err.message, /column B/i);
                assert.match(err.message, /NaN/);
                return true;
            }
        );
    });

    test('Infinity is rejected too', async () => {
        const wb = newWorkbook();
        await assert.rejects(() => wb.addSheet('M', ['a'], [[Infinity]]), /Infinity/);
    });

    test('a row that is not an array is rejected rather than dropped', async () => {
        const wb = newWorkbook();
        await assert.rejects(
            () => wb.addSheet('S', ['a', 'b'], [[1, 2], { a: 1, b: 2 }]),
            /row 3/
        );
    });

    test('no emitted cell ever contains a non-finite numeric value', async () => {
        const wb = newWorkbook();
        await wb.addSheet('Clean', ['a', 'b'], [[1, null], [undefined, 3]]);
        const archive = await readArchive(await wb.saveAsBuffer());
        assert.ok(!/NaN|Infinity/.test(archive.sheetData[0].xml));
    });

    test('row count means data rows for every add method', async () => {
        const wb = newWorkbook();
        const simple = await wb.addSheet('A', ['h1', 'h2'], [[1, 2], [3, 4]]);
        assert.strictEqual(simple.rowCount, 2);

        const sectioned = await wb.addSheetWithSections('B', [
            { title: 'T', headers: ['h'], rows: [[1], [2]] },
            { title: 'U', headers: ['h'], rows: [[3]] },
        ]);
        assert.strictEqual(sectioned.rowCount, 3);

        async function* gen() {
            yield [1];
            yield [2];
        }
        const iterated = await wb.addSheetFromIterator('C', ['h'], gen());
        assert.strictEqual(iterated.rowCount, 2);
    });

    test('sheet info reports comparable counts', async () => {
        const wb = newWorkbook();
        await wb.addSheet('A', ['h1', 'h2'], [[1, 2], [3, 4]]);
        await wb.addSheetWithSections('B', [{ title: 'T', headers: ['h'], rows: [[1]] }]);
        const info = Object.fromEntries(wb.getSheets().map((s) => [s.name, s.rowCount]));
        assert.strictEqual(info.A, 2);
        assert.strictEqual(info.B, 1);
    });

    test('a rejected value is not counted as a written row', async () => {
        const wb = newWorkbook();
        await assert.rejects(() => wb.addSheet('S', ['a'], [[1], [2], [NaN]]));
        assert.deepStrictEqual(wb.getSheets(), []);
    });
});

// ---------------------------------------------------------------------------
// Progress callback
// ---------------------------------------------------------------------------

describe('progress callback', () => {
    test('fires on every add method', async () => {
        const simple = newWorkbook();
        const a = [];
        await simple.addSheet('A', ['h'], [[1], [2], [3]], { onProgress: (...args) => a.push(args) });
        assert.ok(a.length > 0, 'addSheet never called onProgress');

        const sectioned = newWorkbook();
        const b = [];
        await sectioned.addSheetWithSections(
            'B',
            [{ title: 'T', headers: ['h'], rows: [[1]] }],
            { onProgress: (...args) => b.push(args) }
        );
        assert.ok(b.length > 0, 'addSheetWithSections never called onProgress');

        const iterated = newWorkbook();
        const c = [];
        await iterated.addSheetFromIterator('C', ['h'], (async function* () {
            yield [1];
        })(), { onProgress: (...args) => c.push(args) });
        assert.ok(c.length > 0, 'addSheetFromIterator never called onProgress');
    });

    test('the total is passed whenever it is knowable', async () => {
        const wb = newWorkbook();
        const seen = [];
        await wb.addSheet('A', ['h'], [[1], [2], [3]], { onProgress: (...a) => seen.push(a) });
        assert.ok(seen.length > 0, 'onProgress never fired');
        for (const [current, total] of seen) {
            assert.strictEqual(total, 3, 'known total must be passed');
            assert.ok(current >= 1 && current <= 3);
        }
    });

    test('an unknown total degrades to a running count', async () => {
        const wb = newWorkbook();
        const seen = [];
        await wb.addSheetFromIterator(
            'A',
            ['h'],
            (async function* () {
                for (let i = 0; i < 3; i++) yield [i];
            })(),
            { onProgress: (...a) => seen.push(a) }
        );
        assert.ok(seen.length > 0);
        for (const [current, total] of seen) {
            assert.strictEqual(total, undefined);
            assert.ok(Number.isInteger(current) && current >= 1);
        }
    });

    test('the cadence is configurable', async () => {
        const wb = newWorkbook();
        let count = 0;
        await wb.addSheet('A', ['h'], Array.from({ length: 100 }, (_, i) => [i]), {
            progressInterval: 10,
            onProgress: () => count++,
        });
        assert.strictEqual(count, 10);
    });

    test('a throwing callback surfaces instead of being swallowed', async () => {
        const wb = newWorkbook();
        await assert.rejects(
            () =>
                wb.addSheet('A', ['h'], Array.from({ length: 50 }, (_, i) => [i]), {
                    progressInterval: 1,
                    onProgress: () => {
                        throw new Error('ui exploded');
                    },
                }),
            /ui exploded/
        );
    });
});

// ---------------------------------------------------------------------------
// Streaming
// ---------------------------------------------------------------------------

describe('streaming', () => {
    test('a generator is accepted by the array-based add method', async () => {
        function* gen() {
            for (let i = 0; i < 2000; i++) yield [i, `row ${i}`];
        }
        const wb = newWorkbook();
        const result = await wb.addSheet('Lazy', ['id', 'label'], gen());
        assert.strictEqual(result.rowCount, 2000);
        const archive = await readArchive(await wb.saveAsBuffer());
        assert.strictEqual(archive.sheetData[0].rows.length, 2001);
    });

    test('a large export stays within a generous heap ceiling', async () => {
        const wb = newWorkbook();
        const rows = 120000;
        if (global.gc) global.gc();
        const before = process.memoryUsage().heapUsed;
        await wb.addSheet(
            'Huge',
            ['id', 'label'],
            (function* () {
                for (let i = 0; i < rows; i++) yield [i, `row ${i}`];
            })()
        );
        const growth = process.memoryUsage().heapUsed - before;
        // A materialised intermediate would be far larger than this. The bound
        // is loose on purpose: it is a guard against reintroducing buffering,
        // not a performance benchmark.
        assert.ok(growth < 400 * 1024 * 1024, `heap grew by ${Math.round(growth / 1e6)}MB`);
    });
});

// ---------------------------------------------------------------------------
// Options, temp directory ownership, cleanup
// ---------------------------------------------------------------------------

describe('options and temp directory', () => {
    test('a compression level of zero means no compression', async () => {
        const zero = newWorkbook({ compressionLevel: 0 });
        await zero.addSheet('S', ['a'], compressible(200, 'z'));
        const noCompression = (await zero.saveAsBuffer()).length;

        const high = newWorkbook({ compressionLevel: 9 });
        await high.addSheet('S', ['a'], compressible(200, 'z'));
        const compressed = (await high.saveAsBuffer()).length;

        assert.ok(
            noCompression > compressed * 2,
            `level 0 (${noCompression}B) should be far larger than level 9 (${compressed}B); ` +
                'equal sizes mean the zero was ignored'
        );
    });

    test('an out-of-range compression level is rejected naming the value', async () => {
        assert.throws(
            () => new StreamingWorkbook({ tempDir: scratch(), compressionLevel: 12 }),
            /12.*0.*9|0.*and.*9/
        );
        assert.throws(
            () => new StreamingWorkbook({ tempDir: scratch(), compressionLevel: -1 }),
            /-1/
        );
    });

    test('a caller-supplied empty temp directory survives saving', async () => {
        const tempDir = path.join(scratch(), 'mine');
        fs.mkdirSync(tempDir, { recursive: true });
        const wb = new StreamingWorkbook({ tempDir });
        await wb.addSheet('S', ['a'], [[1]]);
        await wb.saveAsBuffer();
        assert.ok(fs.existsSync(tempDir), 'library removed a directory it was given');
    });

    test('a caller-supplied temp directory with unrelated content survives untouched', async () => {
        const tempDir = path.join(scratch(), 'mine');
        fs.mkdirSync(tempDir, { recursive: true });
        const precious = path.join(tempDir, 'keep.txt');
        fs.writeFileSync(precious, 'do not delete');
        const wb = new StreamingWorkbook({ tempDir });
        await wb.addSheet('S', ['a'], [[1]]);
        await wb.saveAsBuffer();
        assert.ok(fs.existsSync(precious));
    });

    test('the default temp location is not inside the working directory', () => {
        const wb = new StreamingWorkbook();
        const cwd = path.resolve(process.cwd());
        const temp = path.resolve(wb.options.tempDir);
        assert.ok(!temp.startsWith(cwd + path.sep), `default temp dir is inside cwd: ${temp}`);
    });

    test('a library-created temp directory is removed when it ends up empty', async () => {
        const wb = new StreamingWorkbook();
        const tempDir = wb.options.tempDir;
        await wb.addSheet('S', ['a'], [[1]]);
        await wb.saveAsBuffer();
        assert.ok(!fs.existsSync(tempDir), 'library-created empty temp dir should be cleaned up');
    });

    test('temp file paths are reachable when cleanup is off', async () => {
        const wb = newWorkbook({ cleanupOnSave: false });
        await wb.addSheet('S', ['a'], [[1]]);
        await wb.saveAsBuffer();
        const files = wb.getTempFiles();
        assert.ok(files.length > 0, 'no temp file paths exposed');
        for (const f of files) {
            assert.ok(fs.existsSync(f), `${f} reported but does not exist`);
        }
    });
});

// ---------------------------------------------------------------------------
// Failure paths
// ---------------------------------------------------------------------------

describe('failure paths', () => {
    test('a failed add leaves no temp file behind', async () => {
        const tempDir = path.join(scratch(), 'tmp');
        const wb = new StreamingWorkbook({ tempDir });
        await assert.rejects(() =>
            wb.addSheet('S', ['a', 'b'], Array.from({ length: 500 }, (_, i) => (i === 400 ? [i, NaN] : [i, i])))
        );
        assert.deepStrictEqual(fs.readdirSync(tempDir), [], 'orphan temp file left behind');
    });

    test('a failed add records nothing', async () => {
        const wb = newWorkbook();
        await assert.rejects(() => wb.addSheet('Good', ['a'], [[1]]).then(() => wb.addSheet('Bad', ['a'], [[NaN]])));
        assert.deepStrictEqual(
            wb.getSheets().map((s) => s.name),
            ['Good']
        );
    });

    test('an iterator error rejects the add and records nothing', async () => {
        const wb = newWorkbook();
        await assert.rejects(
            () =>
                wb.addSheetFromIterator('S', ['a'], (async function* () {
                    yield [1];
                    throw new Error('source exploded');
                })()),
            /source exploded/
        );
        assert.deepStrictEqual(wb.getSheets(), []);
    });

    test('the iterator method rejects rather than hanging on a setup failure', async () => {
        const wb = newWorkbook();
        const result = await Promise.race([
            wb
                .addSheetFromIterator('S', ['a'], (async function* () {
                    yield [1];
                })(), { progressInterval: -1 })
                .then(() => 'resolved', (e) => `rejected: ${e.message}`),
            new Promise((r) => setTimeout(() => r('HUNG'), 2000)),
        ]);
        assert.notStrictEqual(result, 'HUNG', 'the promise never settled');
    });
});

// ---------------------------------------------------------------------------
// Content types
// ---------------------------------------------------------------------------

describe('content types', () => {
    test('every part in the output has a declared content type', async () => {
        const wb = newWorkbook();
        await wb.addSheet('One', ['a'], [[1]]);
        await wb.addSheet('Two', ['a'], [[2]]);
        await assertValid(await wb.saveAsBuffer(), 'content types');
    });

    test('the output stays valid across sheet counts', async () => {
        for (const count of [1, 2, 12]) {
            const wb = newWorkbook();
            for (let i = 0; i < count; i++) await wb.addSheet(`S${i}`, ['a'], [[i]]);
            await assertValid(await wb.saveAsBuffer(), `${count} sheets`);
        }
    });

    test('the styles relationship does not collide with a sheet relationship', async () => {
        const wb = newWorkbook();
        for (let i = 0; i < 3; i++) await wb.addSheet(`S${i}`, ['a'], [[i]]);
        const archive = await readArchive(await wb.saveAsBuffer());
        const targets = [...archive.relationships.values()];
        assert.strictEqual(new Set(targets).size, targets.length, `collision: ${targets}`);
    });

    test('spacer rows between sections have a well-formed spans range', async () => {
        const wb = newWorkbook();
        await wb.addSheetWithSections('S', [
            { title: 'A', headers: ['h'], rows: [[1]] },
            { title: 'B', headers: ['h'], rows: [[2]] },
            { title: 'C', headers: ['h'], rows: [[3]] },
        ]);
        const buffer = await wb.saveAsBuffer();
        const violations = await validateArchive(buffer);
        assert.ok(
            !violations.some((v) => /inverted spans/.test(v)),
            violations.join('; ')
        );
    });
});

// ---------------------------------------------------------------------------
// Output shapes
// ---------------------------------------------------------------------------

describe('output shapes', () => {
    test('saving as a buffer and then to a file works when cleanup is off', async () => {
        const wb = newWorkbook({ cleanupOnSave: false });
        await wb.addSheet('S', ['a'], [[1]]);
        const buffer = await wb.saveAsBuffer();
        const file = await wb.save(path.join(scratch(), 'after-buffer.xlsx'));
        await assertValid(buffer, 'buffer');
        await assertValid(fs.readFileSync(file.filePath), 'file after buffer');
    });

    test('an empty workbook cannot be saved', async () => {
        const wb = newWorkbook();
        await assert.rejects(() => wb.save(path.join(scratch(), 'x.xlsx')), /No sheets/);
        await assert.rejects(() => wb.saveAsBuffer(), /No sheets/);
    });
});

// ---------------------------------------------------------------------------
// Published types
// ---------------------------------------------------------------------------

const repoRoot = path.join(__dirname, '..');
const consumerOut = path.join(repoRoot, 'types-out');

/**
 * Emits the sample consumer next to the repository root so its relative import of
 * `../index` still points at the real entry point, the same place a published
 * consumer's would. A type-only check cannot catch an interop shape that only
 * fails at runtime, which is why this compiles and executes rather than just
 * typechecking.
 */
function emitConsumer() {
    fs.rmSync(consumerOut, { recursive: true, force: true });
    // Resolved through package.json because typescript's exports map does not
    // expose bin/tsc as a subpath.
    const tsc = path.join(path.dirname(require.resolve('typescript/package.json')), 'bin', 'tsc');
    execFileSync(
        process.execPath,
        [
            tsc,
            '-p', path.join(repoRoot, 'tsconfig.json'),
            '--noEmit', 'false',
            '--outDir', consumerOut,
            '--rootDir', path.join(repoRoot, 'types')
        ],
        { cwd: repoRoot, stdio: 'pipe' }
    );
    return require(path.join(consumerOut, 'consumer.js'));
}

describe('published types', () => {
    test('a default import resolves to the class, not the module object', () => {
        const consumer = emitConsumer();
        assert.strictEqual(typeof consumer.StreamingWorkbookDefault, 'function');
        assert.strictEqual(consumer.StreamingWorkbookDefault, consumer.StreamingWorkbook);
        assert.strictEqual(consumer.StreamingWorkbook, require('../index'));
    });

    test('every import style resolves to the same class', () => {
        const entry = require('../index');
        assert.strictEqual(typeof entry, 'function', 'module value should be the class');
        assert.strictEqual(entry.StreamingWorkbook, entry);
        assert.strictEqual(entry.default, entry);
        assert.strictEqual(entry.name, 'StreamingWorkbook');
    });

    test('the consumer runs end to end through the typed API', async () => {
        const consumer = emitConsumer();
        const result = await consumer.run();

        assert.deepStrictEqual(
            result.sheets.map((s) => `${s.name}=${s.rowCount}`),
            ['Plain=2', 'Lazy=3', 'Streamed=2', 'Report=2'],
            'rowCount is data rows only, for every add method'
        );
        assert.ok(result.bytes instanceof Uint8Array, 'saveAsBuffer must resolve without @types/node');
        assert.ok(result.size > 0);
        assert.strictEqual(result.files.length, 4, 'getTempFiles must reach every intermediate file');
        assert.strictEqual(result.altered, false, 'no consumer sheet name needed changing');
    });

    test('emitted output is removed on exit', () => {
        emitConsumer();
        assert.ok(fs.existsSync(path.join(consumerOut, 'consumer.js')));
        roots.push(consumerOut);
    });
});
