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
const path = require('path');
const { promisify } = require('util');

const readFileAsync = promisify(fs.readFile);
const writeFileAsync = promisify(fs.writeFile);
const unlinkAsync = promisify(fs.unlink);
const mkdirAsync = promisify(fs.mkdir);

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
 * await workbook.addSheet('Orders', ['OrderID', 'Amount'], [
 *   [101, 250.00],
 *   [102, 175.50]
 * ]);
 * 
 * await workbook.save('output.xlsx');
 */
class StreamingWorkbook {
    /**
     * Create a new StreamingWorkbook instance
     * @param {Object} options - Configuration options
     * @param {string} [options.tempDir] - Temporary directory for intermediate files
     * @param {boolean} [options.cleanupOnSave=true] - Delete temp files after save
     * @param {number} [options.compressionLevel=6] - ZIP compression level (0-9)
     */
    constructor(options = {}) {
        this.options = {
            tempDir: options.tempDir || path.join(process.cwd(), '.xlsx-temp'),
            cleanupOnSave: options.cleanupOnSave !== false,
            compressionLevel: options.compressionLevel || 6
        };
        
        this.sheets = [];
        this.tempFiles = [];
        this._initialized = false;
    }

    /**
     * Initialize temp directory
     * @private
     */
    async _init() {
        if (!this._initialized) {
            await mkdirAsync(this.options.tempDir, { recursive: true });
            this._initialized = true;
        }
    }

    /**
     * Add a worksheet with data
     * 
     * @param {string} sheetName - Name of the worksheet
     * @param {Array<string>} headers - Column headers
     * @param {Array<Array>|Iterable} rows - Data rows (array of arrays or iterable)
     * @param {Object} [options] - Sheet options
     * @param {Function} [options.onProgress] - Progress callback (rowIndex, totalRows)
     * @returns {Promise<{rowCount: number, duration: number}>}
     * 
     * @example
     * // Simple usage
     * await workbook.addSheet('Sheet1', ['A', 'B'], [[1, 2], [3, 4]]);
     * 
     * // With progress callback
     * await workbook.addSheet('BigData', headers, rows, {
     *   onProgress: (current, total) => console.log(`${current}/${total}`)
     * });
     */
    async addSheet(sheetName, headers, rows, options = {}) {
        await this._init();
        
        const startTime = Date.now();
        const safeName = this._sanitizeSheetName(sheetName);
        const tempFilePath = path.join(this.options.tempDir, `sheet_${this.sheets.length}_${Date.now()}.xlsx`);
        
        const rowCount = await this._writeSheetToFile(tempFilePath, headers, rows, options);
        
        this.sheets.push({
            name: safeName,
            filePath: tempFilePath,
            rowCount
        });
        
        this.tempFiles.push(tempFilePath);
        
        const duration = Date.now() - startTime;
        
        return { rowCount, duration };
    }

    /**
     * Add a worksheet with sections (vertical multi-section layout)
     * 
     * @param {string} sheetName - Name of the worksheet
     * @param {Array<Object>} sections - Array of section configurations
     * @param {string} sections[].title - Section title
     * @param {Array<string>} sections[].headers - Section column headers
     * @param {Array<Array>} sections[].rows - Section data rows
     * @returns {Promise<{rowCount: number, duration: number}>}
     * 
     * @example
     * await workbook.addSheetWithSections('Report', [
     *   { title: 'Sales Q1', headers: ['Product', 'Amount'], rows: [...] },
     *   { title: 'Sales Q2', headers: ['Product', 'Amount'], rows: [...] }
     * ]);
     */
    async addSheetWithSections(sheetName, sections, options = {}) {
        await this._init();
        
        const startTime = Date.now();
        const safeName = this._sanitizeSheetName(sheetName);
        const tempFilePath = path.join(this.options.tempDir, `sheet_${this.sheets.length}_${Date.now()}.xlsx`);
        
        const rowCount = await this._writeSectionedSheetToFile(tempFilePath, sections, options);
        
        this.sheets.push({
            name: safeName,
            filePath: tempFilePath,
            rowCount
        });
        
        this.tempFiles.push(tempFilePath);
        
        const duration = Date.now() - startTime;
        
        return { rowCount, duration };
    }

    /**
     * Add a worksheet from iterable/generator (for very large datasets)
     * 
     * @param {string} sheetName - Name of the worksheet
     * @param {Array<string>} headers - Column headers
     * @param {Iterable|AsyncIterable} rowIterator - Iterator that yields rows
     * @param {Object} [options] - Options
     * @returns {Promise<{rowCount: number, duration: number}>}
     * 
     * @example
     * // Using generator for memory efficiency
     * async function* generateRows() {
     *   for (let i = 0; i < 1000000; i++) {
     *     yield [i, `Row ${i}`, Math.random()];
     *   }
     * }
     * await workbook.addSheetFromIterator('BigData', ['ID', 'Name', 'Value'], generateRows());
     */
    async addSheetFromIterator(sheetName, headers, rowIterator, options = {}) {
        await this._init();
        
        const startTime = Date.now();
        const safeName = this._sanitizeSheetName(sheetName);
        const tempFilePath = path.join(this.options.tempDir, `sheet_${this.sheets.length}_${Date.now()}.xlsx`);
        
        const rowCount = await this._writeIteratorToFile(tempFilePath, headers, rowIterator, options);
        
        this.sheets.push({
            name: safeName,
            filePath: tempFilePath,
            rowCount
        });
        
        this.tempFiles.push(tempFilePath);
        
        const duration = Date.now() - startTime;
        
        return { rowCount, duration };
    }

    /**
     * Save workbook to file
     * 
     * @param {string} outputPath - Output file path
     * @returns {Promise<{filePath: string, fileSize: number, sheetCount: number}>}
     */
    async save(outputPath) {
        if (this.sheets.length === 0) {
            throw new Error('No sheets added to workbook');
        }
        
        const absolutePath = path.resolve(outputPath);
        await this._mergeSheets(absolutePath);
        
        if (this.options.cleanupOnSave) {
            await this._cleanup();
        }
        
        const stats = fs.statSync(absolutePath);
        
        return {
            filePath: absolutePath,
            fileSize: stats.size,
            sheetCount: this.sheets.length
        };
    }

    /**
     * Save workbook to buffer
     * 
     * @returns {Promise<Buffer>}
     */
    async saveAsBuffer() {
        if (this.sheets.length === 0) {
            throw new Error('No sheets added to workbook');
        }
        
        const buffer = await this._mergeSheetsToBuffer();
        
        if (this.options.cleanupOnSave) {
            await this._cleanup();
        }
        
        return buffer;
    }

    /**
     * Get sheet info
     * @returns {Array<{name: string, rowCount: number}>}
     */
    getSheets() {
        return this.sheets.map(s => ({ name: s.name, rowCount: s.rowCount }));
    }

    /**
     * Clean up temporary files manually
     */
    async cleanup() {
        await this._cleanup();
    }

    // ==================== PRIVATE METHODS ====================

    /**
     * Write simple sheet to temp file
     * @private
     */
    _writeSheetToFile(filePath, headers, rows, options) {
        return new Promise((resolve, reject) => {
            const xlsxStream = new XLSXWriteStream();
            const writeStream = fs.createWriteStream(filePath);
            
            let rowCount = 0;
            
            writeStream.on('finish', () => resolve(rowCount));
            writeStream.on('error', reject);
            xlsxStream.on('error', reject);
            
            xlsxStream.pipe(writeStream);
            
            // Write headers
            xlsxStream.write(headers);
            
            // Write data rows
            const rowArray = Array.isArray(rows) ? rows : [...rows];
            const total = rowArray.length;
            const { onProgress } = options;
            
            for (let i = 0; i < total; i++) {
                xlsxStream.write(rowArray[i]);
                rowCount++;
                
                if (onProgress && (i % 1000 === 0 || i === total - 1)) {
                    onProgress(i + 1, total);
                }
            }
            
            xlsxStream.end();
        });
    }

    /**
     * Write sectioned sheet to temp file
     * @private
     */
    _writeSectionedSheetToFile(filePath, sections, options) {
        return new Promise((resolve, reject) => {
            const xlsxStream = new XLSXWriteStream();
            const writeStream = fs.createWriteStream(filePath);
            
            let rowCount = 0;
            
            writeStream.on('finish', () => resolve(rowCount));
            writeStream.on('error', reject);
            xlsxStream.on('error', reject);
            
            xlsxStream.pipe(writeStream);
            
            sections.forEach((section, index) => {
                // Write section title
                xlsxStream.write([section.title || `Section ${index + 1}`]);
                rowCount++;
                
                // Write headers
                xlsxStream.write(section.headers);
                rowCount++;
                
                // Write data rows
                const rows = section.rows || [];
                rows.forEach(row => {
                    xlsxStream.write(row);
                    rowCount++;
                });
                
                // Add spacing between sections
                if (index < sections.length - 1) {
                    xlsxStream.write([]);
                    xlsxStream.write([]);
                    rowCount += 2;
                }
            });
            
            xlsxStream.end();
        });
    }

    /**
     * Write from iterator to temp file
     * @private
     */
    async _writeIteratorToFile(filePath, headers, rowIterator, options) {
        return new Promise(async (resolve, reject) => {
            const xlsxStream = new XLSXWriteStream();
            const writeStream = fs.createWriteStream(filePath);
            
            let rowCount = 0;
            
            writeStream.on('finish', () => resolve(rowCount));
            writeStream.on('error', reject);
            xlsxStream.on('error', reject);
            
            xlsxStream.pipe(writeStream);
            
            // Write headers
            xlsxStream.write(headers);
            
            const { onProgress } = options;
            
            // Handle both sync and async iterators
            try {
                for await (const row of rowIterator) {
                    xlsxStream.write(row);
                    rowCount++;
                    
                    if (onProgress && rowCount % 10000 === 0) {
                        onProgress(rowCount);
                    }
                }
            } catch (err) {
                reject(err);
                return;
            }
            
            xlsxStream.end();
        });
    }

    /**
     * Merge all temp sheets into single Excel file
     * @private
     */
    async _mergeSheets(outputPath) {
        const buffer = await this._mergeSheetsToBuffer();
        await writeFileAsync(outputPath, buffer);
    }

    /**
     * Merge sheets to buffer
     * @private
     */
    async _mergeSheetsToBuffer() {
        const outputZip = new JSZip();
        
        // Read first file for base structure (styles, etc.)
        const firstFileBuffer = await readFileAsync(this.sheets[0].filePath);
        const baseZip = await JSZip.loadAsync(firstFileBuffer);
        
        // Copy base structure (styles, themes, etc.)
        for (const fileName of Object.keys(baseZip.files)) {
            if (!fileName.startsWith('xl/worksheets/') && 
                !fileName.includes('workbook.xml') &&
                !baseZip.files[fileName].dir) {
                const content = await baseZip.files[fileName].async('nodebuffer');
                outputZip.file(fileName, content);
            }
        }
        
        // Process each sheet
        const sheetEntries = [];
        const sheetRels = [];
        const contentTypes = [];
        
        for (let i = 0; i < this.sheets.length; i++) {
            const { name, filePath } = this.sheets[i];
            const sheetId = i + 1;
            const sheetFileName = `sheet${sheetId}.xml`;
            
            // Read sheet file
            const sheetBuffer = await readFileAsync(filePath);
            const sheetZip = await JSZip.loadAsync(sheetBuffer);
            
            // Find and copy sheet XML
            const sheetXmlFile = Object.keys(sheetZip.files).find(f => 
                f.includes('worksheets/sheet') && f.endsWith('.xml')
            );
            
            if (sheetXmlFile) {
                const sheetContent = await sheetZip.files[sheetXmlFile].async('nodebuffer');
                outputZip.file(`xl/worksheets/${sheetFileName}`, sheetContent);
            }
            
            // Escape XML special characters in sheet name
            const escapedName = this._escapeXml(name);
            
            sheetEntries.push(`<sheet name="${escapedName}" sheetId="${sheetId}" r:id="rId${sheetId}"/>`);
            sheetRels.push(`<Relationship Id="rId${sheetId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${sheetFileName}"/>`);
            contentTypes.push(`<Override PartName="/xl/worksheets/${sheetFileName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
        }
        
        // Create workbook.xml
        const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>${sheetEntries.join('')}</sheets>
</workbook>`;
        outputZip.file('xl/workbook.xml', workbookXml);
        
        // Create workbook.xml.rels
        const stylesRelId = this.sheets.length + 1;
        const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheetRels.join('\n')}
<Relationship Id="rId${stylesRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
        outputZip.file('xl/_rels/workbook.xml.rels', workbookRelsXml);
        
        // Create [Content_Types].xml
        const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
${contentTypes.join('\n')}
</Types>`;
        outputZip.file('[Content_Types].xml', contentTypesXml);
        
        // Create _rels/.rels
        const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
        outputZip.file('_rels/.rels', relsXml);
        
        // Generate output buffer
        return outputZip.generateAsync({
            type: 'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: this.options.compressionLevel }
        });
    }

    /**
     * Sanitize sheet name for Excel compatibility
     * @private
     */
    _sanitizeSheetName(name) {
        // Excel sheet name rules:
        // - Max 31 characters
        // - Cannot contain: \ / ? * [ ]
        // - Cannot be empty
        // - Cannot start or end with apostrophe
        
        let safe = String(name)
            .replace(/[\\\/?*\[\]]/g, '_')
            .substring(0, 31)
            .trim();
        
        if (safe.startsWith("'")) safe = '_' + safe.substring(1);
        if (safe.endsWith("'")) safe = safe.substring(0, safe.length - 1) + '_';
        
        return safe || 'Sheet';
    }

    /**
     * Escape XML special characters
     * @private
     */
    _escapeXml(str) {
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    /**
     * Clean up temp files
     * @private
     */
    async _cleanup() {
        for (const filePath of this.tempFiles) {
            try {
                if (fs.existsSync(filePath)) {
                    await unlinkAsync(filePath);
                }
            } catch (e) {
                // Ignore cleanup errors
            }
        }
        
        // Try to remove temp directory if empty
        try {
            const files = fs.readdirSync(this.options.tempDir);
            if (files.length === 0) {
                fs.rmdirSync(this.options.tempDir);
            }
        } catch (e) {
            // Ignore
        }
        
        this.tempFiles = [];
    }
}

module.exports = StreamingWorkbook;
