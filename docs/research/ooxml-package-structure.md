# OOXML (SpreadsheetML) package structure

Primary-source research into the internal structure of an `.xlsx` package, and the gap between
what the spec requires and what `xlsx-stream-workbook` actually emits.

The reader is deciding whether to trust the merge diagram and the README section describing it.
So this document separates three things and never blurs them:

| Label | Meaning |
| --- | --- |
| **Verified** | Read directly out of an archive this library produced, or out of the installed dependency's source. Reproducible. |
| **Required** | Normative `shall` in ECMA-376, or a documented Microsoft deviation in [MS-OI29500]. Cited. |
| **Unverified** | Believed to cause a specific behaviour, but no primary source found. Treated as folklore. |

`test/helpers.js` states the project's own position in its header comment and it is the right one:
its assertions "are on a subset of the format specification, not the whole of it." Nothing below
should be read as a claim that the test suite covers these rules.

---

## 1. Part inventory

### 1.1 The spec's minimum

ECMA-376 Part 1 §12.2 *Package Structure* gives a worked example titled "the minimal conformant
SpreadsheetML package." It contains exactly five items and no more:

```
/[Content_Types].xml
/_rels/.rels
/workbook.xml
/_rels/workbook.xml.rels
/sheet1.xml
```

The normative sentence in §12.2 is: "A SpreadsheetML package shall contain a package-relationship
item and a content-type item. The package-relationship item shall have implicit relationships with
targets of the following type: One Workbook part (§12.3.23)." Optional extras permitted at package
level: Digital Signature Origin (§15.2.7), File Property parts (§15.2.12), Thumbnail (§15.2.16).

**Required.** No Styles part, no Shared String Table, no Theme part in the minimum. Note also that
part names are the format designer's choice — the §12.2 example puts the parts at the package root,
while the three-sheet example later in the same clause uses `/xl/…`. Both are legal; see Part 2
§6.2.2 for part-name syntax.

### 1.2 The real part list in this codebase

**Verified** from `test/output/test1_basic.xlsx` (two sheets: `Users`, `Orders`).

| Part | Produced by | Declared content type |
| --- | --- | --- |
| `[Content_Types].xml` | `lib/archive.js` → `contentTypesXml()` | *not a part* (see below) |
| `_rels/.rels` | `lib/archive.js` → `addWorkbookParts()` | *(no Override)* falls back to `Default Extension="rels"` |
| `xl/_rels/workbook.xml.rels` | `lib/archive.js` → `addWorkbookParts()` | *(no Override)* falls back to `Default Extension="rels"` |
| `xl/workbook.xml` | `lib/archive.js` → `addWorkbookParts()` | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml` |
| `xl/worksheets/sheet1.xml` | copied verbatim from intermediate file 1 | `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml` |
| `xl/worksheets/sheet2.xml` | copied verbatim from intermediate file 2 | `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml` |
| `xl/styles.xml` | carried over from intermediate file 1 | `application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml` |

**Verified** — parts that are *not* emitted, and why that is fine:

| Absent part | Spec status | Note |
| --- | --- | --- |
| `xl/sharedStrings.xml` | Permitted, not required | Part 1 §12.3.15 says "A package shall contain **exactly one** Shared String Table part" — that is a conditional requirement, only bite if the part exists. `xlsx-write-stream` never emits it (uses inline strings, §5). |
| `xl/theme/theme1.xml` | **Zero or one**, per Part 1 §14.2.7 | Absent. `xl/styles.xml` still carries `<color theme="1"/>` and `<scheme val="minor"/>`, which have nothing to resolve against. Schema-valid (`CT_Color/@theme` is `xsd:unsignedInt`, optional — Part 1 Annex A.2 `sml.xsd`), semantically dangling. Excel tolerates it. |
| `docProps/core.xml`, `docProps/app.xml` | Optional (Part 1 §12.2, Part 2 §8) | Absent. Consequence: no document title/author metadata. |
| `xl/calcChain.xml` | Optional | Absent; no formulas are emitted. |

`lib/archive.js` still carries a mapping for `xl/theme/` and `docProps/` content types and a
`SHARED_STRINGS` guard, so the merge fails loudly rather than silently mis-typing a part if
`xlsx-write-stream` ever starts emitting one.

### 1.3 ZIP directory entries — a real, harmless difference

**Verified.** The merged archives contain four zero-length directory entries (`_rels/`, `xl/`,
`xl/worksheets/`, `xl/_rels/`). A single-sheet file written by `xlsx-write-stream` directly contains
none — compare `unzip -l /tmp/raw-single.xlsx` with `unzip -l test/output/test1_basic.xlsx`.

Cause: `jszip@3.10.1`'s `file()` defaults `createFolders` to `true`, so `z.file('xl/worksheets/sheet1.xml', …)`
also creates `xl/` and `xl/worksheets/`. Confirmed by direct call. The raw writer uses `archiver`,
which does not.

**Required.** Part 2 §7.3.5: "The names of all ZIP items shall be mapped to logical item names,
**except for items that do not represent files**." Directory entries are therefore not parts, get no
content type, and appear in no `Override`. The merge is correct. A diagram should show them, or
state that it shows parts only.

---

## 2. The relationship graph

Part 2 §6.5.1: relationships are the indirection that makes part references "directly discoverable
without looking at the part contents." §6.5.2.2: "The name of a package Relationships part shall be
`/_rels/.rels`." §6.5.2.3: a part Relationships part's name is the source part's name with `_rels/`
inserted before the last segment and `.rels` appended.

**Verified** chain, in resolution order, for `test1_basic.xlsx`:

1. `/_rels/.rels` → `xl/workbook.xml`
   - Id `rId1`
   - Type `http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument`
2. `/xl/_rels/workbook.xml.rels` → `xl/worksheets/sheet1.xml`
   - Id `rId1`
   - Type `http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`
3. `/xl/_rels/workbook.xml.rels` → `xl/worksheets/sheet2.xml`
   - Id `rId2`
   - Type `http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`
4. `/xl/_rels/workbook.xml.rels` → `xl/styles.xml`
   - Id `rId3`
   - Type `http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`

**Required** — step 1 is mandatory: Part 1 §12.3.23, "A package shall contain exactly one Workbook
part, and that part shall be the target of a relationship in the package-relationship item."

**Required** — steps 2–3: Part 1 §12.3.24, "A package shall contain exactly one Worksheet part per
worksheet, and those parts shall be the target of an **explicit** relationship from the Workbook
part. Specifically, the `id` attribute on the `sheet` element shall reference the desired worksheet
part." Worksheet relationships are explicit — the `r:id` is written into the XML, and Excel needs it
to find the part.

**Required** — step 4: Part 1 §12.3.20, "A package shall contain no more than one Styles part, and
that part shall be the target of an **implicit** relationship from the Workbook (§12.3.23) part."
Implicit means the consumer infers the target from the relationship *type*; no `r:id` appears in
`workbook.xml`.

**Required** — the Styles relationship is placed at `rId{sheetCount + 1}`, past all sheet ids, so it
can never collide with a worksheet's `r:id`. In `test1_basic.xlsx`: sheets take `rId1`/`rId2`, styles
take `rId3`.

Relationship type URIs, for reference:

| Type | Emitted by | Source |
| --- | --- | --- |
| `…/officeDocument/2006/relationships/officeDocument` | `lib/archive.js` `addWorkbookParts()` | Part 1 §12.3.23 |
| `…/officeDocument/2006/relationships/worksheet` | `lib/archive.js` `WORKSHEET_REL` | Part 1 §12.3.24 |
| `…/officeDocument/2006/relationships/styles` | `lib/archive.js` `addWorkbookParts()` | Part 1 §12.3.20 |
| `…/officeDocument/2006/relationships/sharedStrings` | *not emitted* | Part 1 §12.3.15 |
| `…/officeDocument/2006/relationships/theme` | *not emitted* | Part 1 §14.2.7 |

**Noteworthy for the diagram.** Part 1 §12.3.23 ends with a catch-all prohibition: "A Workbook part
shall not have implicit or explicit relationships to any other part defined by ECMA-376," beyond the
permitted lists. The merge emits exactly worksheet (explicit) + styles (implicit). No relationship
out of the worksheet part itself — which is consistent with Part 1 §12.3.24, which permits comments,
pivots, printer settings, query tables and tables from a worksheet, none of which this library writes.

**Namespace caveat.** ECMA-376 5th edition restates the relationship types under the Strict
namespace `http://purl.oclc.org/ooxml/officeDocument/relationships/…`. What is on disk is the
**Transitional** form, `http://schemas.openxmlformats.org/officeDocument/2006/relationships/…`, and
so is the markup namespace `http://schemas.openxmlformats.org/spreadsheetml/2006/main`. This is the
form Excel and `xlsx-write-stream` produce. Both are conformant; they are different conformance
classes. Do not let a diagram show the `purl.oclc.org` URIs as if they were the ones in the file.

---

## 3. Content types

### 3.1 Exact media type strings

**Verified** from `test/output/test1_basic.xlsx`:

```xml
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml"  ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml"           ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml"             ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml"  ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml"  ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
```

Required media types, Part 1 §12.3.23 / §12.3.24 / §12.3.20 / §12.3.15 and §14.2.7:

| Part type | Media type |
| --- | --- |
| Workbook | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml` |
| Worksheet | `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml` |
| Styles | `application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml` |
| Shared String Table | `application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml` |
| Theme | `application/vnd.openxmlformats-officedocument.theme+xml` |
| Relationships part | `application/vnd.openxmlformats-package.relationships+xml` (Part 2 §6.5.2.1) |
| Core Properties | `application/vnd.openxmlformats-package.core-properties+xml` (Part 2 Annex E) |
| Media Types stream | none — see below |

`[Content_Types].xml` is **not a part**. Part 2 §7.2.3.1: "The Media Types stream shall not represent a
part. This stream shall not be URI-addressable." Part 2 §7.3.7: in ZIP files it is stored in the item
named `[Content_Types].xml`. A diagram that lists it in the part inventory is wrong, even though every
tool will happily print it next to the parts.

### 3.2 Does anything rely on the generic `xml` Default?

**Verified: no.** Every non-`.rels` part in `test1_basic.xlsx` has its own `Override`. The only parts
resolving through a `Default` are the two `.rels` files, which is correct.

**Required** — and this is the interesting part, because the two rules point opposite ways:

- Part 2 §7.2.3.2.1: "For all parts of the package other than Relationships parts, the Media Types
  stream shall specify either: One matching `Default` element, or One matching `Override` element, or
  Both a matching `Default` element and a matching `Override` element, in which case, the `Override`
  element takes precedence." So *at the OPC layer*, a worksheet covered **only** by
  `Default Extension="xml" ContentType="application/xml"` is **conformant**. The generic `xml`
  Default is not itself a violation.
- Part 1 §12.2's minimal example declares no `Default Extension="xml"` at all, and every part's media
  type is normative in §12.3. So a SpreadsheetML package that leaves a part on the generic fallback is
  *technically* OPC-conformant while being *semantically* untyped for a SpreadsheetML consumer.

**This is the gap that matters.** "Spec-legal at the OPC layer" and "Excel will not offer to repair"
are different questions, and the second is a Microsoft implementation behaviour that neither ECMA-376
nor [MS-OI29500] documents. `lib/archive.js` resolves it the safe way: it refuses to write rather than
emit a part it has no media type for, and it emits an explicit `Override` for every shipped part. That
is stricter than §7.2.3.2.1 requires, and deliberately so.

### 3.3 Other Part 2 constraints on the stream

- §7.2.3.2.1: "There shall not be more than one `Default` element for any given extension, and there
  shall not be more than one `Override` element for any given part name." — satisfied.
- §7.2.3.2.1: "The order of `Default` and `Override` elements in the Media Types stream shall not be
  significant." — so the fact that `contentTypesXml()` emits workbook → styles → worksheets, while
  `xlsx-write-stream`'s template emits workbook → worksheet → styles, is a non-issue.
- §7.2.5(a): XML in OPC-defined streams "shall be encoded using either UTF-8 or UTF-16." — all parts
  are written as UTF-8 strings with `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`.
- §7.2.5(b): "DTD declarations shall not be used." — none are.
- §7.3.6: "ZIP-based packages shall not use compression algorithms except DEFLATE." — both writers
  use DEFLATE, or STORE when `compressionLevel: 0` (DEFLATE method, no compression).
- `lib/archive.js` carries a comment noting `jszip@3.10.1` folds `level: 0` into its own default via
  `level || -1`, which is why `archiveOptions()` uses `compression: 'STORE'` instead. Worth keeping
  in the diagram's footnotes.

---

## 4. Why Excel offers to repair

The four faults this repo has already hit, and what primary sources say about each. **This is the
section where verified and required diverge most sharply, and the diagram must not blur them.**

### 4.1 A declared sheet whose relationship resolves to no part

**Required.** Part 1 §12.3.24: worksheet parts "shall be the target of an explicit relationship from
the Workbook part. Specifically, the `id` attribute on the `sheet` element shall reference the desired
worksheet part." A `sheet` element naming an `r:id` that is absent from `xl/_rels/workbook.xml.rels`,
or whose `Target` names a ZIP item that does not exist, violates this directly. It is a **spec
violation**, not folklore.

This is guarded twice in `lib/archive.js`: `copyWorksheetParts()` throws if a sheet's intermediate
file yielded no `xl/worksheets/*.xml`, and `test/helpers.js` `validateArchive()` reports
`declared sheet "…" resolves to missing part …`.

### 4.2 A part with a generic instead of specific content type

**Split verdict, and this is the one to be careful about in the README.**

- **OPC layer: not a violation.** Part 2 §7.2.3.2.1 explicitly permits a matching `Default` to stand
  in for an `Override` (§3.2 above).
- **SpreadsheetML layer: the media type is normative.** Part 1 §12.3.24 gives the Worksheet part's
  content type as `…spreadsheetml.worksheet+xml`, full stop. A part called `xl/worksheets/sheet3.xml`
  typed as `application/xml` is not a Worksheet part.
- **Excel's behaviour: [unverified].** No clause in ECMA-376, and no section in [MS-OI29500] (whose
  index contains no entry for the Media Types stream, `Default`, or `Override`), states that Excel
  rejects such a file. This library's test suite treats it as a violation
  (`part … has no declared content type`), and the choice to keep it that way is defensible — but the
  claim should be written as "Excel repairs these", not "the spec forbids them".

If the diagram needs a one-line rule, the honest one is: *every part a SpreadsheetML consumer must
recognise needs its specific `Override`; the generic `xml` Default is an OPC-legal shortcut that
this library never takes.*

### 4.3 Inverted `spans` ranges on a row

**Verified mechanism.** `xlsx-write-stream@1.0.3` `dist/templates/row.js` emits
`spans="1:${values.length}"` unconditionally. An empty row array produces `spans="1:0"`. The repo
avoids it with two `['']` spacer rows — `lib/sheet-writer.js` `writeSection()`, whose comment reads:
"An empty array yields `spans=\"1:0\"`, which is not a valid range. One empty string does not."

**Required: the spec places no constraint here at all.** Part 1 §18.3.1.73, attribute `spans`:

> "Optimization only, and not required. Specifies the range of non-empty columns (in the format X:Y)
> for the block of rows to which the current row belongs. To achieve the optimization, span attribute
> values in a single block should be the same. […] **[Note: this is an optimization, and is purely
> optional. Different span values within the same row block is allowed. Not writing the span value at
> all is also allowed. end note]**"

And in the schema (Part 1 Annex A.2, `sml.xsd`):

```xsd
<xsd:simpleType name="ST_CellSpan">
  <xsd:restriction base="xsd:string"/>
</xsd:simpleType>
<xsd:simpleType name="ST_CellSpans">
  <xsd:list itemType="ST_CellSpan"/>
</xsd:simpleType>
```

`spans="1:0"` is **schema-valid**. `[MS-OI29500]`'s index has **no entry for `spans`**, so Microsoft
documents no deviation either.

**Conclusion: this is a Microsoft implementation behaviour, not a spec requirement. [Unverified] as to
provenance.** The `spans="1:0"` → repair link rests on observed behaviour, not on any primary source
found. The workaround is still right — it costs one row and removes a known trigger — but the README
should say "emits a value that triggers Excel's repair prompt", not "emits an invalid range". Note
also that the note in §18.3.1.73 says values *may* differ within a 16-row block, so the underlying
`row.js` behaviour is non-conformant with the *stated intent* (uniform `spans` per block) even though
it is conformant with the letter.

> Section numbering note: 5th edition (December 2016) numbers `row` as §18.3.1.73. The 1st/2nd
> edition number is §18.3.1.79, which is the numbering used by the Open XML SDK reference and by
> [MS-OI29500]. The text is identical.

### 4.4 A non-finite number in a numeric cell value

**Verified mechanism.** `xlsx-write-stream` `dist/templates/cell.js` interpolates the number
directly: `` return `<c r="${cell}" t="n"${maybeFormat…}><v>${value}</v></c>` `` — no finiteness
check. `NaN` becomes `<v>NaN</v>`. `lib/sheet-writer.js` `validateRow()` rejects it at the call with
`Number.isFinite`, and `test/helpers.js` scans for `<v>NaN</v>` / `<v>Infinity</v>` / `<v>-Infinity</v>`.

**Required: schema-valid, semantically undefined.** The `v` element's type is `ST_Xstring` (Part 1
Annex A.6.9, shared with `c` §18.3.1.4), i.e. a string. `<v>NaN</v>` is a perfectly good string.
The *cell* is declared `t="n"`, and `ST_CellType` (§18.18.11) defines `n` as "Cell containing a
number." A consumer that honours `t="n"` is entitled to attempt a numeric parse and fail.

**Excel's behaviour: [unverified].** No primary source found. The reasoning above is sound and
explains the failure without appealing to folklore, but it is a derivation from the schema, not a
citation. Say so if the diagram says so.

### 4.5 One more that is not folklore

**Required, and worth adding to the diagram** because a merge is exactly where it goes wrong: Part 2
§6.5.2.1, "There shall be no relationships from or to a Relationships part." A worksheet that
legitimately had, say, a printer settings part would need its own `_rels` entry — and the merge
currently carries no per-sheet `_rels` directory. It is correct today only because nothing writes
worksheet relationships. Worth a "this breaks if we ever add X" note.

---

## 5. Shared strings

### 5.1 What the part is

**Required.** Part 1 §12.3.15: "An instance of this part type contains one occurrence of each unique
string that occurs on **all worksheets in a workbook**." And: "A package shall contain **exactly
one** Shared String Table part, and that part shall be the target of an implicit relationship from
the Workbook part." Root element `sst` (§18.4.9, 5th ed; §18.2.30 in the 1st/2nd ed), a sequence of
`si` items. Cells then hold a *zero-based index* into that one table: `ST_CellType` value `s` means
"Cell containing a shared string."

### 5.2 Why per-sheet indices cannot be concatenated

The indices are positional into a single ordered list. If sheet A's table is `[X, Y]` and sheet B's
is `[Y, Z]`, then B's cell that holds local index `0` (`Y`) would read the wrong string if B's `<v>0</v>`
were copied into a table where `X` occupies slot 0.

Concretely: concatenating `A ++ B` gives `[X, Y, Y, Z]`. B's references are unchanged (`0` → now
`X`, which is wrong) unless every B reference is shifted by `|A| = 2`. Shifting is the only
mechanical fix, and it is the option this library rejects.

Correct approaches, in order of preference:

1. **Do not produce one.** Write inline strings. See §5.3.
2. **Deduplicate into one table.** Build a global table, rewrite every `t="s"` cell's `<v>` to the
   global index. Requires a full parse of every worksheet part and unbounded memory for the table —
   the opposite of this library's streaming design.
3. **Concatenate with an index shift.** Same memory cost as (2) minus the dedup, still unbounded, and
   `uniqueCount`/`count` must be recomputed.

`lib/archive.js` chooses (1) and enforces it: `copyPackageParts()` and `copyWorksheetParts()` both
throw `SHARED_STRINGS` if `xl/sharedStrings.xml` appears in any intermediate file. The message names
the reason — "whose per-sheet indices cannot be merged" — which is exactly right.

### 5.3 `xlsx-write-stream` uses inline strings, so this never arises

**Verified** in `node_modules/xlsx-write-stream@1.0.3/dist/templates/cell.js`: every string and every
boolean is written as

```xml
<c r="A1" t="inlineStr"><is><t>John Doe</t></is></c>
```

No `<v>` index, no `sharedStrings.xml` part, and — worth noting — **booleans too**, as
`t="inlineStr"` with the text `true`/`false` rather than `t="b"`. Numbers and `Date` values do use
`<v>` with `t="n"` (dates as a serial-day count via `unixTimestamp / 86400000 + 25569`).

**Required** — Part 1 §18.18.11 `ST_CellType`, `inlineStr`: "Cell containing an (inline) rich string,
i.e., one not in the shared string table. If this cell type is used, then the cell value is in the
`is` element rather than the `v` element in the cell." So the encoding is correct. `t="b"` is available
in the enumeration and would be more idiomatic, but `inlineStr` carrying `true`/`false` is conformant —
Excel reads it as text.

Implications for the merge, all favourable:

- Worksheet parts are **self-contained**. No cross-part index rewriting is needed at all, which is why
  `copyWorksheetParts()` can copy bytes without reading them.
- The bytes are larger than a shared-string table would be for highly repetitive data. That is the
  entire trade, paid knowingly.
- The guard in `lib/archive.js` is the correct tripwire: it converts a future silent-corruption bug
  into a loud failure.

---

## 6. Markup compatibility in the worksheet header

Relevant because `sheet-header.js` emits `x14ac:dyDescent` on every row, and someone auditing the
diagram will ask whether that is legal.

**Verified** — `xlsx-write-stream@1.0.3/dist/templates/sheet-header.js` declares
`xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`,
`mc:Ignorable="x14ac"`, and `xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"`.

**Required.** ECMA-376 Part 3 §7.2 (`Ignorable` Attribute) makes this the sanctioned mechanism: a
whitespace-delimited list of namespace prefixes whose in-scope namespaces are *declared ignorable*.
§9.2 Step 1 then marks an attribute as ignored when its namespace is declared ignorable and is not in
the consumer's application configuration (Part 3 §3.1) — that is, the consumer does not understand
it (Part 3 §3.7). So a strict consumer is *permitted to discard* `x14ac:dyDescent` without complaint.

Part 2 §6.2.5(d) is the constraint that makes the declaration load-bearing: XML content "shall be
schema-valid … the XML content shall not contain elements or attributes drawn from namespaces that
are not explicitly defined in the corresponding XSD schema unless the XSD schema allows elements or
attributes drawn from any namespace to be present in particular locations." Without
`mc:Ignorable="x14ac"`, the `x14ac` attributes would violate that. With it, MCE processing removes them
first and the result is valid.

**Consequence for the merge.** `mc:Ignorable` and the `x14ac` namespace declaration live in
`sheet-header.js`, i.e. inside the copied worksheet bytes, so they survive the merge untouched. No
action needed; worth one line in the diagram so a reader does not flag it.

Part 3 §9.1: "If an MCE processor detects that a document is non-conformant, the MCE processor
should indicate this non-conformance to the consuming application." That is the normative hook behind
Excel's repair prompt in the generic sense — the mechanism is documented even though the specific
triggers in §4 are not.

---

## 7. `isolatedDeclarations` / declaration emit

Not applicable. `tsconfig.json` sets `"noEmit": true` and includes only `index.d.ts` and
`types/**/*.ts`; `lib/` and `test/` are not type-checked, let alone emitted. No TypeScript compiler
setting participates in producing the `.xlsx` byte stream, so nothing in the OOXML package structure
above turns on it. Recorded here only so the omission is visibly deliberate.

---

## 8. Spec-required vs. verified: the gap

| Concern | Spec says | This codebase | Gap |
| --- | --- | --- | --- |
| Package relationship item → Workbook | **Required** (Part 1 §12.3.23) | Present, `rId1` | none |
| Sheet `r:id` → Worksheet part | **Required**, explicit (Part 1 §12.3.24) | Present, `rId1`…`rIdN` | none |
| ≤1 Styles part, implicit from Workbook | **Required if present** (Part 1 §12.3.20) | One, `rId{N+1}` | none |
| `sheets` element present in `workbook` | **Required** — `minOccurs="1"` (Part 1 Annex A.2, `CT_Workbook`) | Present | none |
| Other `workbook` children | All optional (`fileVersion`, `bookViews`, `calcPr`, `workbookPr`) | All omitted | none; the merged `workbook.xml` is minimal but conformant |
| Media type for every non-`.rels` part | OPC: `Default` **or** `Override` (§7.2.3.2.1). SpreadsheetML: specific type normative (§12.3) | Explicit `Override` for all | Stricter than OPC requires. Deliberate. |
| `Default Extension="xml"` present | Not prohibited; absent from the §12.2 minimal example | Present, unused by any part | Harmless dead declaration |
| ZIP directory entries | Not parts (Part 2 §7.3.5) | Present via jszip `createFolders` | none |
| No theme part | **Zero or one** (Part 1 §14.2.7) | Zero, but `styles.xml` has `theme="1"` / `scheme val="minor"` | Schema-valid, semantically dangling. Tolerated. |
| No shared strings | Permitted; §12.3.15 bites only if present | Absent, inline strings instead | none — the design choice that makes merging safe |
| Inverted `spans` | **No constraint** (§18.3.1.73; `ST_CellSpan` = `xsd:string`) | Avoided via 1-cell spacers | Workaround for an **[unverified]** Excel behaviour |
| Non-finite `<v>` with `t="n"` | Schema-valid (`ST_Xstring`); `t="n"` promises a number (§18.18.11) | Rejected at the call | **[unverified]** Excel behaviour, sound derivation |
| `x14ac:*` attributes | Legal via Part 3 §7.2 + §9.2 | `mc:Ignorable="x14ac"` present | none |
| No relationships from/to a `.rels` part | **Forbidden** (Part 2 §6.5.2.1) | None today | Would break silently if worksheet rels were ever added |

---

## 9. Sources

**Normative specifications** (primary; text extracted from the official ECMA PDFs, clause numbers
from the 5th editions unless noted):

- ECMA-376 Part 1, *Fundamentals and Markup Language Reference*, 5th ed., December 2016 —
  <https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip>
  — §12.2, §12.3.15, §12.3.20, §12.3.23, §12.3.24, §14.2.7, §18.2.20, §18.3.1.4, §18.3.1.73,
  §18.4.9, §18.18.11; Annex A.2 (`sml.xsd`; also distributed inside the same archive as
  `OfficeOpenXML-XMLSchema-Strict.zip`)
- ECMA-376 Part 2, *Open Packaging Conventions*, 5th ed., December 2021 —
  <https://ecma-international.org/wp-content/uploads/ECMA-376-2_5th_edition_december_2021.zip>
  — §6.2.2, §6.2.3, §6.2.5, §6.5.1, §6.5.2, §7.2.3, §7.3.1–§7.3.7; Annex E
- ECMA-376 Part 3, *Markup Compatibility and Extensibility*, 5th ed., December 2015 —
  <https://ecma-international.org/wp-content/uploads/ECMA-376-3_5th_edition_december_2015.zip>
  — §3.1, §3.7, §7.2, §9.1, §9.2

**Microsoft implementation notes** (primary; documents Office's deviations from ECMA-376):

- [MS-OI29500], *Office Implementation Information for ECMA-376 Standards Support* — index
  <https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/a9247126-8b23-4297-b15b-fef61d56a43d>
  — consulted specifically to test whether `spans` is documented; **it is not**, and there is no
  entry for the Media Types stream, `Default`, `Override`, or the shared string table
- [MS-OI29500] Part 1 Section 18.3.1.4, `c` (Cell) —
  <https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/2fd4e47f-0965-4c60-95bd-cff980b6c325>
  — documents Office's stricter rules for `@cm`, `@vm`, `@s`
- [MS-OI29500] Part 1 Section 18.2.20, `sheets` (Sheets) —
  <https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/c9766f83-abdd-4351-aeff-90859c90c9f7>
  — "Excel limits the occurrences of this element to 32767"
- Open Packaging Conventions Fundamentals (archived Windows SDK docs) —
  <https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview>
  — corroborating conceptual overview; cited for nothing normative

**Dependency source** (ground truth for what is emitted):

- `node_modules/xlsx-write-stream@1.0.3/dist/templates/` — `workbook.js`, `workbook-rels.js`,
  `rels.js`, `content-types.js`, `styles.js`, `sheet-header.js`, `sheet-footer.js`, `row.js`, `cell.js`
- `node_modules/xlsx-write-stream@1.0.3/dist/` — `XLSXRowTransform.js`, `XLSXTransformStream.js`,
  `utils.js`

**This repository** (the merge, and the generated archives):

- `lib/archive.js`, `lib/sheet-writer.js`, `lib/StreamingWorkbook.js`
- `test/output/test1_basic.xlsx`, `test2_sections.xlsx`, `test3_large.xlsx`, `test4_generator.xlsx`
- `test/helpers.js`
- `jszip@3.10.1` — `file()` `createFolders` default, verified by direct call

**Not used:** blog posts, tutorials, Stack Overflow, or any other secondary source. Where a claim
could not be traced to one of the above it is labelled **[unverified]** and appears nowhere else in
this document as though it were fact.
