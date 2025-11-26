/**
 * xlsx-stream-workbook TypeScript Definitions
 */

export interface WorkbookOptions {
    /** Temporary directory for intermediate files */
    tempDir?: string;
    /** Delete temp files after save (default: true) */
    cleanupOnSave?: boolean;
    /** ZIP compression level 0-9 (default: 6) */
    compressionLevel?: number;
}

export interface SheetOptions {
    /** Progress callback */
    onProgress?: (current: number, total?: number) => void;
}

export interface Section {
    /** Section title */
    title: string;
    /** Column headers */
    headers: string[];
    /** Data rows */
    rows: any[][];
}

export interface AddSheetResult {
    /** Number of rows written */
    rowCount: number;
    /** Duration in milliseconds */
    duration: number;
}

export interface SaveResult {
    /** Absolute file path */
    filePath: string;
    /** File size in bytes */
    fileSize: number;
    /** Number of sheets */
    sheetCount: number;
}

export interface SheetInfo {
    /** Sheet name */
    name: string;
    /** Row count */
    rowCount: number;
}

export class StreamingWorkbook {
    /**
     * Create a new StreamingWorkbook instance
     */
    constructor(options?: WorkbookOptions);

    /**
     * Add a worksheet with data
     */
    addSheet(
        sheetName: string,
        headers: string[],
        rows: any[][],
        options?: SheetOptions
    ): Promise<AddSheetResult>;

    /**
     * Add a worksheet with sections (vertical multi-section layout)
     */
    addSheetWithSections(
        sheetName: string,
        sections: Section[],
        options?: SheetOptions
    ): Promise<AddSheetResult>;

    /**
     * Add a worksheet from iterable/generator
     */
    addSheetFromIterator(
        sheetName: string,
        headers: string[],
        rowIterator: Iterable<any[]> | AsyncIterable<any[]>,
        options?: SheetOptions
    ): Promise<AddSheetResult>;

    /**
     * Save workbook to file
     */
    save(outputPath: string): Promise<SaveResult>;

    /**
     * Save workbook to buffer
     */
    saveAsBuffer(): Promise<Buffer>;

    /**
     * Get sheet info
     */
    getSheets(): SheetInfo[];

    /**
     * Clean up temporary files manually
     */
    cleanup(): Promise<void>;
}

export default StreamingWorkbook;
