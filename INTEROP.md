# GridMD Interoperability, Storage & Security

**Version 0.1 (draft).** Companion to [SPEC.md](SPEC.md). Covers the XLSX ⇄
GridMD mapping, fidelity classes, the database storage model, diff/merge
behaviour, and security.

---

## 1. Fidelity classes

Every Excel feature belongs to exactly one class. The invariant across all
classes: **nothing is silently dropped** — a converter represents natively,
carries via escape hatch, or fails loudly.

| Class | Meaning | Contract |
|---|---|---|
| **F0 — native** | First-class GridMD grammar | Lossless round-trip through the core syntax |
| **F1 — structured** | Modeled, minor normalization | Round-trips semantically; byte-level XLSX internals (e.g. style-table indices, pivot caches) are regenerated |
| **F2 — carried** | No native grammar; `fallback:` / `{raw}` | Byte-preserved and re-emitted; not readable/editable as GridMD |
| **F3 — out of scope** | Transient app state | Deliberately not represented |

## 2. XLSX ⇄ GridMD feature map

| Excel feature (OOXML home) | GridMD | Class |
|---|---|---|
| Cell values & types (`sheetData`, shared strings) | scalars in `@`/`{grid}`/`{table}` (shared strings dissolve) | F0 |
| Formulas + cached values (`<f>`/`<v>`) | `=F :: V` | F0 |
| Dynamic arrays / spill (`t="array"`, `ref`) | `{ spill: }` / `{spill-cache}` / `{ array: }` | F0 |
| Defined names, LAMBDA (`definedNames`) | frontmatter / `{sheet}` `names:` | F0 |
| Number formats (`numFmts`) | `numfmt:` codes | F0 |
| Fonts/fills/borders/alignment (`styles.xml` xf records) | cell/range props; style tables dissolve into named styles + props | F1 |
| Built-in cell styles (`cellStyles`) | built-in catalog names (FORMATTING.md §5) | F1 |
| Themes (`theme1.xml`) | frontmatter `theme:` (12 colors + 2 fonts; effects ignored) | F1 |
| Merged cells (`mergeCells`) | `@ range { merge: true }` | F0 |
| Row/col size, hidden, outline (`cols`, `row` attrs) | `{sheet}` `rows:`/`cols:`, `{outline}` | F0 |
| Freeze/split/zoom/view (`sheetViews`) | `{sheet}` `freeze/split/view` | F0 |
| Conditional formatting (`conditionalFormatting`, dxf) | `{cf}` (dxf styles inlined into `format:`) | F1 |
| Tables (`tables/table1.xml`) | `{table}` | F0 |
| AutoFilter + sort state (`autoFilter`, `sortState`) | `{filter}` / table `filter:`/`sort:` | F1 |
| Data validation (`dataValidations`) | `{validation}` | F0 |
| Charts (`charts/chart1.xml`, `charts/chartEx1.xml`) | `{chart}` — classic chartML emitted natively (column/bar/line/area/pie/doughnut/scatter/radar/bubble/stock/combo + stacked/`-3d` variants, secondary axis, trendlines, error bars, data labels, axes) AND ChartEx `cx:` parts for treemap/sunburst/waterfall/funnel/histogram/pareto/box-whisker/map with literal data points; both reverse-parsed on import | F1 |
| PivotCharts (`pivotSource`) | `{chart} pivot:` — emitted with a `c:pivotSource` binding + the pivot output range as the initial series (Excel repopulates on refresh) | F1 |
| Chart sheets (`chartsheets/`) | `{sheet} kind: chart` + `{chart} at sheet` — emitted natively | F1 |
| Sparklines (x14 extLst) | `{sparklines}` — emitted natively | F1 |
| Pivot tables (`pivotTable/CacheDefinition/CacheRecords`) | `{pivot}` — cache definition + empty records emitted with `refreshOnLoad`; Excel rebuilds from in-document source; OLAP → `fallback:` | F1/F2 |
| Slicers (x14/x15 parts) | `{slicer}` — table slicers emitted natively | F1 |
| Timelines (`timelineCaches/`, `timelines/`) | `{slicer} kind: timeline` — emitted natively for pivot targets (Excel has no table timelines; a table-targeted timeline is carried with a loud note); reverse-parsed on import | F1 |
| What-if scenarios (`scenarios`) | `{scenario}` | F1 |
| Chart sheets (`chartsheets/sheet1.xml`) | `{sheet} kind: chart` + `{chart}` `at sheet` | F1 |
| Hyperlinks (`hyperlinks` + rels) | `link:`/`tip:` props | F0 |
| Notes (legacy comments, `comments1.xml` + VML) | `note:` prop (VML shape geometry dropped) | F1 |
| Threaded comments (`threadedComments`) | `{comments}` | F0 |
| Images (`drawing1.xml`, media) | `{image}` (media by path/URL/data URI) | F1 |
| Shapes/text boxes (DrawingML) | `{shape}`/`{textbox}` common subset; preset-geometry exotica → `fallback:` | F1/F2 |
| Form controls / in-cell checkbox | `{checkbox}` / `control: checkbox` | F1 |
| Rich data types (`richData` parts) | `entity:`/`fields:` props | F1 |
| Protection (workbook/sheet/cell, password hashes) | `protection:` / `protect:` / `locked:` | F0 |
| Page setup, headers/footers, breaks (`pageSetup` etc.) | `{page}` | F0 |
| Calculation settings (`calcPr`) | frontmatter `calc:` (calcChain regenerated) | F0 |
| 1900/1904 date system (`workbookPr date1904`) | `date-system:` | F0 |
| Power Query (`customXml`/DataMashup M) | `{query}` bounded steps; original M → `fallback:` | F1/F2 |
| Office Scripts | `{script}` | F1 |
| VBA (`vbaProject.bin`) | `{raw}` base64 | F2 |
| SmartArt, WordArt effects, ink, OLE, ActiveX | `{raw}` / `fallback:` | F2 |
| External workbook links (`externalLink1.xml`) | frontmatter `links:` + `[n]Sheet!Ref` formulas + cached values | F1 |
| Custom XML parts, doc custom properties | `{raw}` / frontmatter `properties.x-*` | F2 |
| Ignored-error markers (`ignoredErrors`), custom views (`customWorkbookViews`) | `x-*` props / `{raw}` | F2 |
| Window positions, selection, scroll, undo, clipboard, Watch Window lists (no OOXML persistence) | — | F3 |

### Unit conversions (F1 normalizations)

- **Column width:** OOXML character-width ↔ px via the standard formula at the
  minor-font size; GridMD stores px (integers preferred).
- **Row height:** points ↔ px at 96 dpi (`px = pt × 96⁄72`).
- **Dates:** serial ↔ ISO 8601 per `date-system`. Excel's 1900-leap-year bug
  cell is the exception: serial 60 has no real ISO date, so converters MUST NOT
  coerce it to `1900-02-28`. They preserve the cell as numeric `60` with the
  original date `numfmt` (and MAY add an `x-excel-date-serial: 60` extension
  prop for UI display).
- **Anchors:** EMU offsets ↔ px (`px = EMU ⁄ 9525`).

## 3. Database storage model

GridMD is designed to decompose. The recommended relational shape:

```sql
workbooks (id, title, frontmatter_yaml, version, updated_at)
blocks    (id, workbook_id, sheet_seq, sheet_name, seq,
           kind,          -- 'sheet'|'grid'|'table'|'at'|'cf'|'chart'|…
           anchor,        -- 'A1', 'A1:D9', or NULL
           name,          -- table/chart/query name, or NULL
           body TEXT)     -- the block's verbatim GridMD source
-- optional hot projection for cell-level queries:
cells     (workbook_id, sheet_name, ref, row, col,
           type, value, formula, numfmt, style_json)
```

- **Block = row.** A `.gmd` file is the ordered concatenation of frontmatter +
  block bodies; assembly is `ORDER BY sheet_seq, seq`, preserving workbook sheet
  order even when sheet names sort differently. A single-cell read hits the
  `cells` projection; a single-cell write rewrites one `at` block row (or one
  grid row within a block).
- **Addressability:** `(workbook, sheet, kind, anchor|name)` is a natural key
  for every feature — the same key a UI panel edits.
- Contiguous `@` runs MAY be stored one-per-row for maximal write granularity;
  the canonical formatter reassembles deterministic output either way.
- Document stores: one document per sheet with a `blocks` array is the
  equivalent shape.

## 4. Diff & merge behaviour

- **One fact per line.** A cell edit is a one-line diff; a rule change is a
  few lines inside one block. Canonical form (SPEC §12) makes diffs stable —
  format before commit, like `gofmt`.
- **Concurrent edits** to different blocks never textually conflict. Within a
  block, line-level merge works for `@` runs and grid rows; YAML bodies
  conflict at the key line — small, reviewable conflicts.
- **Semantic merge** (recommended for products): parse both sides to blocks,
  merge block-wise by natural key, re-emit canonically. Only a same-cell /
  same-key double edit is a true conflict.
- Content hashes of canonical form give cheap change detection per block.

## 5. Security

- **Scripts are inert.** `{script}` and anything inside `{raw}` MUST NOT
  execute on load without explicit host policy + user consent. Treat like
  macro-enabled workbooks.
- **Raw package parts.** `{raw}`/`fallback` OOXML payloads are data, not trusted
  package instructions. Before re-emitting them into XLSX/XLSM, writers MUST
  canonicalize and validate `part=` paths as specified in DIRECTIVES.md §18,
  prevent package-part overwrite/smuggling, and require explicit consent for
  macro-bearing, ActiveX, OLE, external-relationship, or executable-adjacent
  parts.
- **Formula injection.** When importing untrusted CSV/text *into* GridMD,
  apply the classic guard: values beginning `= + - @ \t` are emitted as text
  (rule-4 `'` prefix), never as formulas, unless the importer is explicitly
  trusted.
- **Links.** Readers SHOULD enforce a scheme allowlist (`https`, `mailto`,
  internal `#`) on `link:` and `{image} src:` before dereferencing; `data:`
  URIs are size-capped by host policy.
- **External fetches** (`{query}` `url:`, external `links:`) are network
  egress: SSRF-guard (block private/loopback ranges) and require consent in
  server-side contexts.
- **Protection ≠ security.** Sheet/workbook protection hashes are courtesy
  locks (as in Excel). Confidentiality requires transport/storage encryption
  outside the format.
- **Resource limits.** Parsers MUST bound: total size, block count, grid
  dimensions (≤ 1,048,576 × 16,384), YAML depth (no anchors/aliases — GridMD
  YAML is the safe subset: no tags, no aliases, no multi-doc; the only scalar
  types that materialize are strings, numbers, and booleans — timestamps and
  other typed scalars are read as plain strings), and `{raw}` payload size.

## 6. Media type & versioning

- Proposed media type `text/gridmd`; extension `.gmd`; embedded fence language
  `gridmd`.
- `gridmd: "MAJOR.MINOR"` in frontmatter. Minor versions add directives/keys
  only (readers skip unknowns in lenient mode); major versions may change
  grammar. Writers emit the lowest version that expresses the document.
- Extension surface: `x-` keys (anywhere) and `{x-*}` directives; both
  round-trip untouched. Standardization path: `x-foo` → `foo` in a minor rev.
