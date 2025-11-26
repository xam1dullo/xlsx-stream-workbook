# xlsx-stream-workbook

📊 **Streaming Excel workbook writer with multiple worksheets support**

Memory-efficient solution for creating Excel files with multiple sheets, handling 100K+ rows without memory issues.

## Features

- ✅ **Streaming writes** - Memory efficient for large datasets
- ✅ **Multiple worksheets** - Create workbooks with many sheets
- ✅ **Sectioned sheets** - Vertical multi-section layouts
- ✅ **Iterator support** - Use generators for massive datasets
- ✅ **Progress callbacks** - Track write progress
- ✅ **Buffer output** - Save to file or get buffer directly
- ✅ **TypeScript support** - Full type definitions included

## Installation

```bash
npm install xlsx-stream-workbook
```

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
| `tempDir` | string | `.xlsx-temp` | Temporary directory for intermediate files |
| `cleanupOnSave` | boolean | `true` | Delete temp files after save |
| `compressionLevel` | number | `6` | ZIP compression level (0-9) |

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
    onProgress: (current, total) => {
        console.log(`Progress: ${current}/${total}`);
    }
});
```

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

Save workbook and return as Buffer.

```javascript
const buffer = await workbook.saveAsBuffer();
// Use buffer for HTTP response, email attachment, etc.
```

### `workbook.getSheets()`

Get information about added sheets.

```javascript
const sheets = workbook.getSheets();
// [{ name: 'Sheet1', rowCount: 100 }, { name: 'Sheet2', rowCount: 50 }]
```

### `workbook.cleanup()`

Manually clean up temporary files.

```javascript
await workbook.cleanup();
```

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

1. **Use generators for massive datasets** - Avoids loading all data into memory
2. **Adjust compression level** - Lower levels (1-3) are faster, higher (7-9) produce smaller files
3. **Process in batches** - For database queries, fetch and write in chunks

## Requirements

- Node.js >= 14.0.0

## Dependencies

- [xlsx-write-stream](https://www.npmjs.com/package/xlsx-write-stream) - Streaming XLSX writer
- [jszip](https://www.npmjs.com/package/jszip) - ZIP file manipulation

## License

MIT © Khamidullo Khudoyberdiev

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.
# xlsx-stream-workbook
