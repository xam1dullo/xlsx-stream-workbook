/**
 * Sample consumer for the published type definitions.
 *
 * This file is compiled by `npm run typecheck` but never shipped. It exists so
 * that a declaration narrowing what the implementation accepts fails the build
 * instead of reaching a consumer, and so the emitted JavaScript can prove that
 * both import styles resolve to the class at runtime rather than only in types.
 *
 * It deliberately imports nothing from Node, so the project needs no
 * `@types/node` to check its own public surface.
 */

import StreamingWorkbookDefault, { StreamingWorkbook } from '../index';
import type { AddSheetResult, SaveResult, SheetInfo, WorkbookOptions } from '../index';

const options: WorkbookOptions = {
    tempDir: 'types-out',
    cleanupOnSave: false,
    compressionLevel: 6
};

function* rows(): Generator<Array<number | string>> {
    for (let i = 0; i < 3; i++) yield [i, `row ${i}`];
}

async function* streamed(): AsyncGenerator<number[]> {
    for (let i = 0; i < 2; i++) yield [i];
}

export async function run(): Promise<{
    sheets: SheetInfo[];
    files: string[];
    size: number;
    bytes: Uint8Array;
    altered: boolean;
}> {
    const workbook = new StreamingWorkbook(options);

    const plain: AddSheetResult = await workbook.addSheet('Plain', ['id'], [[1], [2]]);
    const lazy: AddSheetResult = await workbook.addSheet('Lazy', ['id', 'label'], rows());
    const iterated: AddSheetResult = await workbook.addSheetFromIterator('Streamed', ['id'], streamed());

    const sectioned: AddSheetResult = await workbook.addSheetWithSections('Report', [
        { title: 'First', headers: ['h'], rows: [[1]] },
        { headers: ['h'], rows: [[2]] }
    ]);

    const names: string[] = [plain.name, lazy.name, iterated.name, sectioned.name];
    const altered: boolean = names.some((_, index) =>
        [plain, lazy, iterated, sectioned].some((result) => result.nameChanged && result.name === names[index])
    );

    const sheets: SheetInfo[] = workbook.getSheets();
    const files: string[] = workbook.getTempFiles();

    const bytes: Uint8Array = await workbook.saveAsBuffer();
    const saved: SaveResult = await workbook.save('types-out/consumer-out.xlsx');

    await workbook.cleanup();

    return { sheets, files, size: saved.fileSize, bytes, altered };
}

export { StreamingWorkbook, StreamingWorkbookDefault };
