/**
 * xlsx-stream-workbook
 * Streaming Excel workbook writer with multiple worksheets support
 * 
 * @example
 * const { StreamingWorkbook } = require('xlsx-stream-workbook');
 * 
 * const workbook = new StreamingWorkbook();
 * await workbook.addSheet('Sheet1', ['Name', 'Age'], [['John', 30], ['Jane', 25]]);
 * await workbook.save('output.xlsx');
 */

const StreamingWorkbook = require('./lib/StreamingWorkbook');

module.exports = {
    StreamingWorkbook,
    default: StreamingWorkbook
};
