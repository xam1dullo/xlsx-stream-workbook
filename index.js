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

// The module value is the class itself, carrying itself under both interop
// names. That makes `import X from`, `import { X } from`, `const { X } = require`
// and `const X = require` all resolve to the class, which a bare object
// export cannot do: without an `__esModule` marker, a default import would
// receive the module object instead and `new` on it would throw.
module.exports = StreamingWorkbook;
module.exports.StreamingWorkbook = StreamingWorkbook;
module.exports.default = StreamingWorkbook;
