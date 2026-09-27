# xlsx-stream-workbook

Producing Excel (`.xlsx`) workbooks from JavaScript and TypeScript, with declared types that
describe what the writer will actually accept rather than what it will accept quietly.

## Language

**Workbook**:
The caller's top-level handle. Owns a set of sheets and the intermediate files that hold them.
_Avoid_: Document, file, export

**Sheet**:
The result of one add call — a name, headers, and data rows. A workbook holds many; each is
independent of the others.
_Avoid_: Tab, page, worksheet (Excel's own term, but the library's unit is a call, not a tab)

**Data row**:
A row the caller submitted. The unit every count in this library is measured in.
_Avoid_: Row (ambiguous — see structural row), record, entry

**Structural row**:
A row the library writes for layout: a header, a section title, or a spacer between sections.
Present in the file, absent from every count.
_Avoid_: Filler, padding row, chrome

**Intermediate file**:
The single-sheet `.xlsx` that one sheet's rows are written into. Between adding a sheet and
saving, it is the only place that sheet's data exists.
_Avoid_: Temp file, scratch file, staging file

**Merge**:
The one pass that combines every intermediate file into a single package.
_Avoid_: Combine, assemble, zip (zip is the format, not the act)

**Sealed**:
The state a workbook enters when its intermediate files are released. A sealed workbook can
neither be saved nor extended.
_Avoid_: Closed, finished, finalised, done

**Requested name**:
The sheet name string the caller passed in.
_Avoid_: Name (unqualified)

**Written name**:
The requested name after normalisation and de-duplication — the name that actually appears in
the file. The two differ whenever the caller is told the name changed.
_Avoid_: Sanitised name, final name

**Rejected value**:
A cell value the library refuses to write, identified by sheet, row, and column. Rejection
happens at the call, not in the produced file.
_Avoid_: Invalid value, bad input, error (an error is what the rejection throws)
