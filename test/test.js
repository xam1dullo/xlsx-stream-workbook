/**
 * Test file for xlsx-stream-workbook
 */

const { StreamingWorkbook } = require('../index');
const path = require('path');

async function runTests() {
    console.log('🧪 Testing xlsx-stream-workbook\n');
    
    // Test 1: Basic usage
    console.log('Test 1: Basic multi-sheet workbook');
    const workbook1 = new StreamingWorkbook({
        tempDir: path.join(__dirname, 'temp')
    });
    
    await workbook1.addSheet('Users', 
        ['ID', 'Name', 'Email', 'Age'],
        [
            [1, 'John Doe', 'john@example.com', 30],
            [2, 'Jane Smith', 'jane@example.com', 25],
            [3, 'Bob Wilson', 'bob@example.com', 35]
        ]
    );
    
    await workbook1.addSheet('Orders',
        ['OrderID', 'UserID', 'Product', 'Amount'],
        [
            [101, 1, 'Widget A', 99.99],
            [102, 2, 'Widget B', 149.99],
            [103, 1, 'Widget C', 79.99]
        ]
    );
    
    const result1 = await workbook1.save(path.join(__dirname, 'output', 'test1_basic.xlsx'));
    console.log(`   ✅ Created: ${result1.filePath}`);
    console.log(`   📊 Sheets: ${result1.sheetCount}, Size: ${(result1.fileSize / 1024).toFixed(2)} KB\n`);
    
    // Test 2: Sectioned sheet
    console.log('Test 2: Sectioned sheet (vertical layout)');
    const workbook2 = new StreamingWorkbook();
    
    await workbook2.addSheetWithSections('KPI_Report', [
        {
            title: 'CS_TRAFFIC_ERL',
            headers: ['BSC', 'Day1', 'Day2', 'Day3', 'Day4', 'Day5', 'Day6', 'Day7'],
            rows: [
                ['BSC01', 1000, 1200, 1150, 1100, 1050, 900, 850],
                ['BSC02', 2000, 2100, 2200, 2150, 2100, 1900, 1850],
                ['BSC03', 1500, 1600, 1550, 1500, 1450, 1300, 1250]
            ]
        },
        {
            title: 'DROP_RATE_%',
            headers: ['BSC', 'Day1', 'Day2', 'Day3', 'Day4', 'Day5', 'Day6', 'Day7'],
            rows: [
                ['BSC01', 0.5, 0.6, 0.4, 0.5, 0.3, 0.4, 0.5],
                ['BSC02', 0.3, 0.4, 0.3, 0.4, 0.2, 0.3, 0.4],
                ['BSC03', 0.4, 0.5, 0.4, 0.3, 0.4, 0.3, 0.4]
            ]
        },
        {
            title: 'CSSR_%',
            headers: ['BSC', 'Day1', 'Day2', 'Day3', 'Day4', 'Day5', 'Day6', 'Day7'],
            rows: [
                ['BSC01', 99.5, 99.4, 99.6, 99.5, 99.7, 99.6, 99.5],
                ['BSC02', 99.7, 99.6, 99.7, 99.6, 99.8, 99.7, 99.6],
                ['BSC03', 99.6, 99.5, 99.6, 99.7, 99.6, 99.7, 99.6]
            ]
        }
    ]);
    
    const result2 = await workbook2.save(path.join(__dirname, 'output', 'test2_sections.xlsx'));
    console.log(`   ✅ Created: ${result2.filePath}`);
    console.log(`   📊 Sheets: ${result2.sheetCount}, Size: ${(result2.fileSize / 1024).toFixed(2)} KB\n`);
    
    // Test 3: Large dataset with progress
    console.log('Test 3: Large dataset (10,000 rows) with progress');
    const workbook3 = new StreamingWorkbook();
    
    const largeData = [];
    for (let i = 0; i < 10000; i++) {
        largeData.push([
            i + 1,
            `Cell_${String(i).padStart(5, '0')}`,
            `Site_${Math.floor(i / 100)}`,
            Math.random() * 1000,
            Math.random() * 100,
            new Date(2024, 0, 1 + (i % 365)).toISOString().split('T')[0]
        ]);
    }
    
    const result3Sheet = await workbook3.addSheet(
        'LargeData',
        ['ID', 'Cell', 'Site', 'Traffic_MB', 'Drop_Rate', 'Date'],
        largeData,
        {
            onProgress: (current, total) => {
                if (current % 2000 === 0 || current === total) {
                    process.stdout.write(`\r   Progress: ${current}/${total} (${((current/total)*100).toFixed(0)}%)`);
                }
            }
        }
    );
    console.log(`\n   ✅ Written ${result3Sheet.rowCount} rows in ${result3Sheet.duration}ms`);
    
    const result3 = await workbook3.save(path.join(__dirname, 'output', 'test3_large.xlsx'));
    console.log(`   📊 Size: ${(result3.fileSize / 1024).toFixed(2)} KB\n`);
    
    // Test 4: Generator/Iterator
    console.log('Test 4: Generator-based writing');
    const workbook4 = new StreamingWorkbook();
    
    async function* generateRows() {
        for (let i = 0; i < 5000; i++) {
            yield [i + 1, `Item_${i}`, Math.random() * 500];
        }
    }
    
    const result4Sheet = await workbook4.addSheetFromIterator(
        'Generated',
        ['ID', 'Name', 'Value'],
        generateRows(),
        {
            onProgress: (count) => {
                if (count % 1000 === 0) {
                    process.stdout.write(`\r   Generated: ${count} rows`);
                }
            }
        }
    );
    console.log(`\n   ✅ Generated ${result4Sheet.rowCount} rows in ${result4Sheet.duration}ms`);
    
    const result4 = await workbook4.save(path.join(__dirname, 'output', 'test4_generator.xlsx'));
    console.log(`   📊 Size: ${(result4.fileSize / 1024).toFixed(2)} KB\n`);
    
    // Test 5: Buffer output
    console.log('Test 5: Save as buffer');
    const workbook5 = new StreamingWorkbook();
    
    await workbook5.addSheet('BufferTest', ['A', 'B', 'C'], [[1, 2, 3], [4, 5, 6]]);
    
    const buffer = await workbook5.saveAsBuffer();
    console.log(`   ✅ Buffer size: ${(buffer.length / 1024).toFixed(2)} KB\n`);
    
    // Summary
    console.log('✅ All tests passed!\n');
    console.log('📁 Output files:');
    console.log('   - test/output/test1_basic.xlsx');
    console.log('   - test/output/test2_sections.xlsx');
    console.log('   - test/output/test3_large.xlsx');
    console.log('   - test/output/test4_generator.xlsx');
}

// Create output directory
const fs = require('fs');
const outputDir = path.join(__dirname, 'output');
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

// Run tests
runTests()
    .then(() => process.exit(0))
    .catch(err => {
        console.error('❌ Test failed:', err);
        process.exit(1);
    });
