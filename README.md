# xlsx-stream-workbook

📊 **Streaming Excel workbook writer with multiple worksheets support**

Memory-efficient solution for creating Excel files with multiple sheets, handling 100K+ rows without memory issues.

Rows are written through a backpressured stream and any iterable you pass is consumed lazily, so writing does not grow memory with row count. Merging the finished sheets into one archive is bounded by the archive itself — see [Performance Tips](#performance-tips) for the remaining caveat.

## Features

- ✅ **Streaming writes** - Memory efficient for large datasets
- ✅ **Safe parallel writes** - Sheets can be added concurrently
- ✅ **Multiple worksheets** - Create workbooks with many sheets
- ✅ **Sectioned sheets** - Vertical multi-section layouts
- ✅ **Iterator support** - Use generators for massive datasets
- ✅ **Progress callbacks** - Track write progress
- ✅ **Buffer output** - Save to file or get buffer directly
- ✅ **Loud on bad input** - Values that would corrupt the file are rejected, not dropped
- ✅ **TypeScript support** - Hand-written types, verified on every test run

## Installation

```bash
npm install xlsx-stream-workbook
```

## How it works

Writing a workbook is a two-phase job, and knowing which phase you are in explains most of this library's behaviour — including why a second `save()` needs an opt-in.

**Phase one, write.** Every `addSheet` call streams that sheet's rows straight into its own temporary `.xlsx` file and returns. Your data is written and forgotten; nothing accumulates in memory, and you never hold a complete workbook.

**Phase two, merge.** On `save()`, those temporary files are merged into one package: one `workbook.xml`, one relationship set, content types derived from the parts actually shipped, and a single compressed archive written to disk.

That split has three consequences worth knowing up front:

- **The temporary files are the only copy of your data.** Once a save releases them, the workbook cannot be saved or extended again. See [Repeat saves](#repeat-saves).
- **Adding sheets is cheap and parallel-safe.** Each sheet is independent, so `Promise.all` over several `addSheet` calls is fine — temp file identity does not depend on timing.
- **The merge is where memory goes.** Writing is streamed and backpressured, but `jszip` assembles the finished archive in memory. See [Performance Tips](#performance-tips).

[**See the two-phase write and merge pipeline →**](docs/diagrams/architecture.html)
&nbsp;&nbsp;·&nbsp;&nbsp;
[**Full walkthrough, annotated →**](docs/diagrams/architecture-full.html)

## Quick Start

```javascript
const { StreamingWorkbook } = require('xlsx-stream-workbook');

async function createReport() {
    const workbook = new StreamingWorkbook();
    
    // Add a simple sheet
    await workbook.addSheet('Users', 
        ['ID', 'Name', 'Email'],
        [
            [1, 'John Doe', 'john@example.com'],
            [2, 'Jane Smith', 'jane@example.com'],
            [3, 'Bob Wilson', 'bob@example.com']
        ]
    );
    
    // Add another sheet
    await workbook.addSheet('Orders',
        ['OrderID', 'UserID', 'Amount', 'Date'],
        [
            [101, 1, 250.00, '2024-01-15'],
            [102, 2, 175.50, '2024-01-16'],
            [103, 1, 89.99, '2024-01-17']
        ]
    );
    
    // Save to file
    const result = await workbook.save('report.xlsx');
    console.log(`Created: ${result.filePath} (${result.fileSize} bytes)`);
}

createReport();
```

## API Reference

### `new StreamingWorkbook(options?)`

Create a new workbook instance.

**Options:**
| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `tempDir` | string | a directory under the OS temp dir | Temporary directory for intermediate files. A directory you supply here is never removed by this library. |
| `cleanupOnSave` | boolean | `true` | Delete temp files after save. See [Repeat saves](#repeat-saves). |
| `compressionLevel` | number | `6` | ZIP compression level, 0-9. Level 0 stores without compression. Any other value throws a `RangeError`. |

`save()` creates missing parent directories, so you do not have to `mkdir` first.

### `workbook.addSheet(name, headers, rows, options?)`

Add a worksheet with data.

```javascript
await workbook.addSheet('Sheet1', 
    ['Col1', 'Col2', 'Col3'],
    [
        ['A1', 'B1', 'C1'],
        ['A2', 'B2', 'C2']
    ]
);
```

**With progress callback:**
```javascript
await workbook.addSheet('BigData', headers, rows, {
    onProgress: (completed, total) => {
        console.log(`Progress: ${completed}/${total}`);
    }
});
```

The callback contract is the same for all three add methods: it receives rows completed, plus the total whenever the row count is knowable up front. Streaming from an iterator or a lazy iterable means the total is unknown, so it arrives as `undefined` and the callback reports a running count. The callback fires at least once on a non-empty sheet. Control the cadence with `progressInterval` (default 1000 rows):

```javascript
await workbook.addSheet('BigData', headers, rows, {
    progressInterval: 10_000,
    onProgress: (completed, total) => { /* ... */ }
});
```

The callback must be synchronous — a returned promise is not awaited — and if it throws, the write fails rather than being silently swallowed.

### `workbook.addSheetWithSections(name, sections, options?)`

Add a worksheet with multiple sections (vertical layout).

```javascript
await workbook.addSheetWithSections('KPI Report', [
    {
        title: 'Sales Q1',
        headers: ['Product', 'Revenue', 'Units'],
        rows: [
            ['Widget A', 15000, 150],
            ['Widget B', 22000, 180]
        ]
    },
    {
        title: 'Sales Q2',
        headers: ['Product', 'Revenue', 'Units'],
        rows: [
            ['Widget A', 18000, 175],
            ['Widget B', 25000, 200]
        ]
    }
]);
```

**Output:**
```
Sales Q1
Product    | Revenue | Units
Widget A   | 15000   | 150
Widget B   | 22000   | 180

Sales Q2
Product    | Revenue | Units
Widget A   | 18000   | 175
Widget B   | 25000   | 200
```

### `workbook.addSheetFromIterator(name, headers, iterator, options?)`

Add a worksheet from an iterator or generator (for very large datasets).

```javascript
// Generator for 1 million rows
async function* generateLargeData() {
    for (let i = 0; i < 1000000; i++) {
        yield [i + 1, `User ${i}`, Math.random() * 1000];
    }
}

await workbook.addSheetFromIterator(
    'MassiveData',
    ['ID', 'Name', 'Value'],
    generateLargeData(),
    {
        onProgress: (count) => {
            if (count % 100000 === 0) {
                console.log(`Written ${count} rows`);
            }
        }
    }
);
```

### `workbook.save(outputPath)`

Save workbook to file.

```javascript
const result = await workbook.save('output.xlsx');
// result: { filePath: string, fileSize: number, sheetCount: number }
```

### `workbook.saveAsBuffer()`

Save workbook and return it as a `Uint8Array`. At runtime the value is a Node `Buffer`, so `Buffer` consumers are unaffected, but the declared type resolves without `@types/node`.

```javascript
const buffer = await workbook.saveAsBuffer();
// Use buffer for HTTP response, email attachment, etc.
```

### `workbook.getSheets()`

Get information about added sheets. `rowCount` is the number of **data** rows, so counts are comparable across `addSheet`, `addSheetWithSections` and `addSheetFromIterator`.

```javascript
const sheets = workbook.getSheets();
// [{ name: 'Sheet1', rowCount: 100 }, { name: 'Sheet2', rowCount: 50 }]
```

### `workbook.getTempFiles()`

Paths of the intermediate files this workbook created. Useful when `cleanupOnSave: false` — you own their lifetime once you have them.

```javascript
const files = workbook.getTempFiles(); // ['/tmp/xlsx-stream-workbook-1234/sheet_1_0.xlsx']
```

### `workbook.cleanup()`

Manually release the intermediate files. The workbook cannot be saved or extended afterwards.

```javascript
await workbook.cleanup();
```

## TypeScript

Types ship in `index.d.ts` and are hand-written rather than generated, so they can say why a contract is what it is. They are checked on every `npm test` against a sample consumer (`types/consumer.ts`) that exercises every public method, so a declaration that narrows what the implementation accepts fails the build instead of reaching you.

```typescript
import { StreamingWorkbook } from 'xlsx-stream-workbook';
// or: import StreamingWorkbook from 'xlsx-stream-workbook';
```

Both import styles resolve to the class. Named and default imports are the same value, as are `const { StreamingWorkbook } = require(...)` and `const StreamingWorkbook = require(...)`.

The check runs under TypeScript 7.0, a devDependency, with `strict` on. Nothing in the public types references Node's ambient globals, so you do not need `@types/node` to use this package.

## Behaviour

### Cell values

A cell value that cannot be represented in the file is rejected at the call site, naming the sheet, row and column — rather than written out as a file Excel refuses to open:

- `NaN` and `Infinity` are rejected.
- Types other than `string`, finite `number`, `boolean`, `Date`, `null` and `undefined` are rejected. A row that is not an array is rejected too, rather than silently dropped.
- `null` and `undefined` leave the cell empty, which is what you want for sparse rows.
- Cell text is written as an explicitly-typed string, so a leading `=` is text and not a formula.

### Sheet names

Names are normalised against Excel's rules (31 characters, no `\ / ? * [ ] :`, no control characters, no leading or trailing apostrophe) and then made unique across the workbook, comparing case-insensitively as Excel does. If your name was altered, `addSheet` returns `nameChanged: true` along with the `name` actually used.

### Repeat saves

The intermediate files are the only copy of your data. With the default `cleanupOnSave: true`, saving releases them, after which the workbook cannot be saved or extended again. To save more than once — or to add a sheet between saves — pass `cleanupOnSave: false` and call `cleanup()` when you are done:

```javascript
const workbook = new StreamingWorkbook({ cleanupOnSave: false });
await workbook.addSheet('Q1', headers, q1Rows);
await workbook.save('q1.xlsx');
await workbook.addSheet('Q2', headers, q2Rows);
await workbook.save('q1-q2.xlsx');
await workbook.cleanup();
```

The full state machine, including why `sealed` throws rather than returning, is drawn in [**the workbook lifecycle →**](docs/diagrams/lifecycle.html).

### Package integrity

The merge is what makes the file valid. A workbook has exactly one `workbook.xml` and one relationship set no matter how many sheets, so everything sheet-specific is replaced and everything package-level is carried forward from the first intermediate file.

**Required by the spec, and enforced here**

- Every declared sheet resolves to a real part. ECMA-376 Part 1 §12.3.24 requires the `r:id` on a `<sheet>` element to reference a worksheet part, so a dangling entry is a genuine spec violation. The merge refuses to write one.
- The Styles relationship is placed past every sheet id, so it can never collide with a worksheet's `r:id`. §12.3.20 permits at most one Styles part, reached by an *implicit* relationship.

**Stricter than the spec, on purpose**

- Every non-`.rels` part gets an explicit `Override` carrying its specific media type. OPC (Part 2 §7.2.3.2.1) would accept a matching `Default Extension="xml"` in its place. This library never takes that shortcut, because a part left on the generic fallback is technically conformant while being untyped to a SpreadsheetML consumer — an unknown part is a hard error rather than a guess.

**Verified behaviour, unverified provenance** — the confidence differs, and it is worth knowing which is which

- `NaN` and `Infinity` are rejected. The XML schema permits `<v>NaN</v>` on its own, since `ST_Xstring` is a string — but the cell is declared `t="n"`, which means "cell containing a number" (§18.18.11), and a consumer is entitled to fail the parse. The reasoning is sound; the Excel repair prompt itself is not documented anywhere.
- Spacer rows between sections are one empty cell, not an empty row. ECMA-376 places no constraint on `spans` at all, so `1:0` is schema-valid. It has been *observed* to trigger Excel's repair prompt, but no primary source documents it.

**No shared strings, deliberately.** `xlsx-write-stream` writes inline strings, so the package never contains `xl/sharedStrings.xml`. That is precisely what makes merging safe: a shared string table holds one entry per unique string across *all* worksheets, and each cell holds a positional index into it. Concatenating two per-sheet tables would leave every index in the second sheet pointing at the wrong string unless each were shifted. Inline strings avoid the problem rather than solving it.

Part-by-part detail, with the clause citations and a spec-required-versus-verified table: [`docs/research/ooxml-package-structure.md`](docs/research/ooxml-package-structure.md).

## Examples

### Large Dataset with Progress

```javascript
const { StreamingWorkbook } = require('xlsx-stream-workbook');

async function exportLargeData(data) {
    const workbook = new StreamingWorkbook();
    
    const result = await workbook.addSheet(
        'Export',
        ['ID', 'Timestamp', 'Value', 'Category'],
        data,
        {
            onProgress: (current, total) => {
                const percent = ((current / total) * 100).toFixed(1);
                process.stdout.write(`\rExporting: ${percent}%`);
            }
        }
    );
    
    console.log(`\nWrote ${result.rowCount} rows in ${result.duration}ms`);
    
    await workbook.save('large-export.xlsx');
}
```

### Multi-Sheet Report with Sections

```javascript
const { StreamingWorkbook } = require('xlsx-stream-workbook');

async function createNetworkReport() {
    const workbook = new StreamingWorkbook();
    
    // Sheet 1: Simple data
    await workbook.addSheet('Traffic', 
        ['Site', 'Hour', 'Traffic_MB'],
        trafficData
    );
    
    // Sheet 2: Sectioned KPIs
    await workbook.addSheetWithSections('KPIs by Day', [
        { title: 'CS_TRAFFIC', headers: ['BSC', 'D1', 'D2', 'D3'], rows: csData },
        { title: 'PS_TRAFFIC', headers: ['BSC', 'D1', 'D2', 'D3'], rows: psData },
        { title: 'DROP_RATE', headers: ['BSC', 'D1', 'D2', 'D3'], rows: dropData }
    ]);
    
    // Sheet 3: Large cell data
    await workbook.addSheet('Cells', cellHeaders, cellData);
    
    await workbook.save('network-report.xlsx');
}
```

### Express.js Download Endpoint

```javascript
const express = require('express');
const { StreamingWorkbook } = require('xlsx-stream-workbook');

const app = express();

app.get('/download/report', async (req, res) => {
    try {
        const workbook = new StreamingWorkbook({ cleanupOnSave: true });
        
        await workbook.addSheet('Data', headers, data);
        
        const buffer = await workbook.saveAsBuffer();
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="report.xlsx"');
        res.send(buffer);
        
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});
```

## Performance Tips

1. **Pass an iterable, not an array** - `addSheet` accepts any iterable and consumes it lazily, so a generator never has to exist in full. Same for `addSheetFromIterator` with an async generator.
2. **Adjust compression level** - Lower levels (1-3) are faster, higher (7-9) produce smaller files, and 0 disables compression entirely.
3. **Process in batches** - For database queries, fetch and write in chunks.
4. **Adding sheets in parallel is safe** - `Promise.all` over `addSheet` calls produces a valid workbook; temp file identity does not depend on timing.

**Known limitation.** Writing is streamed and backpressured, but the merge step is not. `jszip` assembles the whole archive in memory, so peak memory during `save()` is roughly the size of the finished workbook. If you need true end-to-end streaming for a very large workbook, `exceljs`'s streaming writer is the better tool.

## Requirements

- Node.js >= 18.0.0 (the test suite uses the built-in `node:test` runner)

## Dependencies

- [xlsx-write-stream](https://www.npmjs.com/package/xlsx-write-stream) - Streaming XLSX writer
- [jszip](https://www.npmjs.com/package/jszip) - ZIP file manipulation

## License

MIT © Khamidullo Khudoyberdiev

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.
