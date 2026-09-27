/**
 * xlsx-stream-workbook TypeScript Definitions
 */

export interface WorkbookOptions {
    /**
     * Temporary directory for intermediate files.
     * Defaults to a directory under the platform temp location. A directory
     * supplied here is never removed by the library.
     */
    tempDir?: string;
    /** Delete temp files after save (default: true). See SaveResult notes on repeat saves. */
    cleanupOnSave?: boolean;
    /**
     * ZIP compression level 0-9 (default: 6). Level 0 stores without compression.
     * Throws a RangeError for a non-integer or out-of-range value.
     */
    compressionLevel?: number;
}

export interface SheetOptions {
    /**
     * Called as rows are written. Receives the number of data rows completed
     * and, when the total is knowable up front, that total. The total is
     * undefined when streaming from an iterator or a lazy iterable.
     *
     * Must be synchronous; a returned promise is not awaited.
     */
    onProgress?: (completed: number, total?: number) => void;
    /** Rows between progress callbacks (default: 1000). */
    progressInterval?: number;
}

export interface Section {
    /** Section title. Defaults to its 1-based position when omitted. */
    title?: string;
    /** Column headers for this section */
    headers?: string[];
    /** Data rows for this section */
    rows?: any[][];
}

export interface AddSheetResult {
    /** Number of data rows written. Excludes headers, titles and spacer rows. */
    rowCount: number;
    /** Duration in milliseconds */
    duration: number;
    /** The sheet name as written, after sanitisation and uniqueness resolution */
    name: string;
    /** True when the requested name was altered */
    nameChanged: boolean;
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
    /** Number of data rows in the sheet */
    rowCount: number;
}

export class StreamingWorkbook {
    /**
     * Create a new StreamingWorkbook instance
     *
     * @throws {RangeError} if compressionLevel is not an integer between 0 and 9
     */
    constructor(options?: WorkbookOptions);

    /**
     * Add a worksheet with data.
     *
     * A cell value that cannot be written to an Excel file — a non-finite
     * number, or a type other than string, finite number, boolean, Date, null
     * or undefined — rejects with an error naming the sheet, row and column.
     * A row that is not an array rejects too, rather than being dropped.
     *
     * @param rows A plain array, or any iterable of rows. Iterables are
     *   consumed lazily and are never materialised.
     */
    addSheet(
        sheetName: string,
        headers: string[],
        rows: Iterable<any[]> | any[][],
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
     * Add a worksheet from an iterable or async iterable (for very large datasets)
     */
    addSheetFromIterator(
        sheetName: string,
        headers: string[],
        rowIterator: Iterable<any[]> | AsyncIterable<any[]>,
        options?: SheetOptions
    ): Promise<AddSheetResult>;

    /**
     * Save workbook to file, creating missing parent directories.
     *
     * With the default cleanupOnSave, saving releases the intermediate files,
     * after which this workbook can no longer be saved or extended. To save
     * more than once, pass cleanupOnSave: false and call cleanup() when done.
     */
    save(outputPath: string): Promise<SaveResult>;

    /**
     * Save workbook to buffer. See save() for the repeat-save constraint.
     *
     * Returns a Uint8Array rather than a Node Buffer so that this declaration
     * resolves without `@types/node`. At runtime the value is a Buffer, which is
     * a Uint8Array, so existing Buffer consumers are unaffected.
     */
    saveAsBuffer(): Promise<Uint8Array>;

    /**
     * Get sheet info
     */
    getSheets(): SheetInfo[];

    /**
     * Paths of the intermediate files this workbook created. Useful when
     * cleanupOnSave is false; the caller owns their lifetime once returned.
     */
    getTempFiles(): string[];

    /**
     * Release the intermediate files. The workbook cannot be saved or extended
     * afterwards.
     */
    cleanup(): Promise<void>;
}

export default StreamingWorkbook;
