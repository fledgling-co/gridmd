// Branch/line coverage for the XLSX writer + importer. A feature-dense
// document exercises the emission paths; round-tripping it back exercises the
// reverse-parse paths; synthetic zips exercise the error branches.

import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { deflateRawSync } from 'node:zlib';
import { lint } from '../src/index';
import { buildWorkbookModel } from '../src/xlsx/model';
import { writeXlsx } from '../src/xlsx/write';
import { xlsxToGridmd } from '../src/xlsx/read';
import { zipWrite, zipRead, crc32 } from '../src/xlsx/zip';
import type { ReportEntry, WorkbookModel } from '../src/types';

const modelOf = (src: string, baseDir = '.'): WorkbookModel => {
  const res = lint(src, { mode: 'strict' });
  assert.deepEqual(res.errors, [], `must lint clean: ${res.errors.map((e) => e.msg).join('; ')}`);
  return buildWorkbookModel(res.doc, { baseDir });
};
const convert = (src: string, baseDir = '.'): { buffer: Buffer; report: ReportEntry[] } =>
  writeXlsx(modelOf(src, baseDir));

// 1x1 PNG
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC';

test('kitchen-sink: every CF rule, validation, page + sparkline shape emits and round-trips', () => {
  const src = `---
gridmd: "0.1"
title: Everything
properties: { author: Jo, company: Acme, created: 2026-03-04 }
calc: { mode: manual, iterative: { enabled: true, max-iterations: 50, max-change: 0.01 } }
theme: { colors: { accent1: "#101010" } }
protection: { structure: true }
names:
  - { name: Tax, value: "0.1", hidden: true }
---

# Sheet1

\`\`\`{sheet}
tab-color: "#334455"
freeze: B2
view: { gridlines: false, headings: false, formulas: true, rtl: true, zoom: 80 }
cols: { A: 120, "B:C": { width: 90, hidden: true, group: 1 } }
rows: { "1": { height: 40, hidden: true, group: 1 } }
protect: { enabled: true, allow: [sort, autofilter, select-locked] }
names:
  - { name: Local, ref: "Sheet1!$A$1" }
\`\`\`

\`\`\`{grid} A1
| a | b | c |
| 1 | 2 | 3 |
| 4 | 5 | 6 |
\`\`\`

\`\`\`{cf} A1:A9
- when: "> 5"
  format: { fill: "#ffcc00", bold: true }
- when: "between 1 and 10"
- when: "not-between 3 and 4"
- contains: "x"
- not-contains: "y"
- begins: "z"
- ends: "q"
- date: today
- dupes: true
  stop: true
- unique: true
- top: "10%"
- bottom: 3
- avg: below-equal
  stddev: 1
- bars: { color: "#00ff00", min: { type: num, value: 0 }, max: { type: auto } }
- scale: ["#ff0000", "#00ff00", "#0000ff"]
- icons: 4-arrows
  reverse: true
  icons-only: true
- formula: =A1>0
\`\`\`

\`\`\`{validation} A1
type: list
values: [x, y, z]
dropdown: false
blank: false
input: { title: Pick, message: choose one }
error: { style: warning, title: Bad, message: no }
\`\`\`

\`\`\`{validation} B1
type: whole
op: ">"
value: 5
\`\`\`

\`\`\`{validation} C1
type: decimal
min: 0
max: 1
\`\`\`

\`\`\`{validation} D1
type: custom
formula: =D1>0
\`\`\`

\`\`\`{outline}
rows:
  - { range: "2:3", level: 1, collapsed: true }
cols:
  - { range: "D:E", level: 1, collapsed: true }
\`\`\`

\`\`\`{page}
orientation: landscape
paper: a4
scale: 90
margins: { top: 2, bottom: 2, left: 2, right: 2, header: 1, footer: 1 }
print-area: A1:C3
print-titles: { rows: "1:1", cols: "A:A" }
header: { left: L, center: C, right: R }
footer: Page
gridlines: true
headings: true
center: { horizontal: true, vertical: true }
breaks: { rows: [5], cols: [3] }
\`\`\`

@ E1 "note-cell" { note: "a note", link: "https://x.dev" }
@ E2 "int" { link: "#Sheet1!A1", tip: "go" }

\`\`\`{sparklines} F1:F3
type: win-loss
source: A1:C3
markers: { high: true, low: true, first: true, last: true, negative: true }
axis: { show: true }
color: accent1
\`\`\`

\`\`\`{comments} A1
- by: Ann
  at: 2026-01-01T00:00:00
  text: root
  resolved: true
  replies:
    - { by: Bob, at: 2026-01-02T00:00:00, text: reply }
\`\`\`

\`\`\`{scenario} Down
cells: { A2: 0 }
comment: dip
\`\`\`

\`\`\`{shape} star at H1:J4
text: |
  hello
fill: "#abcdef"
outline: { color: "#000000", width: 2 }
font: { size: 14, bold: true, color: "#111111" }
\`\`\`

\`\`\`{textbox} at H6
text: box
\`\`\`

\`\`\`{image} at L1
src: "data:image/png;base64,${PNG}"
alt: pixel
\`\`\`

\`\`\`{checkbox} at M1
linked: N1
\`\`\`

\`\`\`{query} Q1
source: X
\`\`\`

\`\`\`{script} Sc1 lang=js
---
console.log(1)
\`\`\`

\`\`\`{raw} ooxml part="customXml/extra.xml"
<a/>
\`\`\`
`;
  const { buffer, report } = convert(src);
  assert.deepEqual(report.filter((r) => r.action === 'not-emitted'), []);
  const parts = zipRead(buffer);
  const sheet = parts.get('xl/worksheets/sheet1.xml')!.toString();
  assert.match(sheet, /dataBar/);
  assert.match(sheet, /colorScale/);
  assert.match(sheet, /iconSet/);
  assert.match(sheet, /timePeriod/);
  assert.match(sheet, /aboveAverage/);
  // round-trip back
  const { gmd } = xlsxToGridmd(buffer);
  const res = lint(gmd, { mode: 'strict' });
  for (const e of res.errors) console.error(`rt:${e.line}: ${e.msg}`);
  assert.deepEqual(res.errors, []);
});

test('chart families: pie/doughnut/scatter/bubble/stock + data shorthand + combo axes', () => {
  const src = `---
gridmd: "0.1"
---

# S

\`\`\`{grid} A1
| cat | v | v2 | lo | hi |
| a | 1 | 5 | 1 | 9 |
| b | 2 | 6 | 2 | 8 |
\`\`\`

\`\`\`{chart} pie "Pie" at G1:K10
data: A2:A3, B2:B3
legend: { position: none }
\`\`\`

\`\`\`{chart} doughnut at G12:K20
series:
  - { name: D, cat: A2:A3, val: B2:B3 }
\`\`\`

\`\`\`{chart} scatter at M1:Q10
series:
  - { name: S, cat: B2:B3, val: C2:C3, marker: circle, labels: { show: true, position: center, numfmt: "0.0" } }
\`\`\`

\`\`\`{chart} bubble at M12:Q20
series:
  - { name: Bub, cat: B2:B3, val: C2:C3, size: B2:B3 }
\`\`\`

\`\`\`{chart} stock at A6:E14
series:
  - { name: St, cat: A2:A3, val: D2:D3, smooth: true }
\`\`\`

\`\`\`{chart} combo at S1:W12
series:
  - { name: One, cat: A2:A3, val: B2:B3, kind: column, color: "#123456", gap: 80, overlap: 20 }
  - { name: Two, cat: A2:A3, val: C2:C3, kind: line, axis: y2, trendline: { type: poly, order: 2, forecast: { forward: 1, backward: 1 }, intercept: 0, equation: true, r2: true }, error-bars: { dir: both, type: percentage, value: 5, cap: false } }
axes:
  x: { title: X, gridlines: true, reverse: true, numfmt: "0" }
  y: { title: Y, min: 0, max: 10, unit: 2, minor-unit: 1, log: true, gridlines: true }
  y2: { title: Y2 }
data-table: { show: true, legend-keys: false }
\`\`\`
`;
  const { buffer, report } = convert(src);
  assert.deepEqual(report.filter((r) => r.action === 'not-emitted'), []);
  const parts = zipRead(buffer);
  const charts = [...parts.keys()].filter((k) => /chart\d+\.xml$/.test(k)).map((k) => parts.get(k)!.toString()).join('');
  assert.match(charts, /doughnutChart/);
  assert.match(charts, /scatterChart/);
  assert.match(charts, /bubbleChart/);
  assert.match(charts, /stockChart/);
  const { gmd } = xlsxToGridmd(buffer);
  assert.deepEqual(lint(gmd, { mode: 'strict' }).errors, []);
});

test('pivot from a range source (headed range, not a table)', () => {
  const src = `---
gridmd: "0.1"
---

# Data

\`\`\`{grid} A1
| region | amount |
| AU | 10 |
| NZ | 20 |
\`\`\`

\`\`\`{pivot} P1 at E1
source: Data!A1:B3
rows:
  - { field: region }
cols:
  - { field: amount }
filters:
  - { field: region }
values:
  - { field: amount, agg: average, name: Avg, show-as: percent-of-total, numfmt: "0.00" }
layout: tabular
grand-totals: { rows: false, cols: false }
\`\`\`
`;
  const { buffer, report } = convert(src);
  assert.deepEqual(report.filter((r) => r.action === 'not-emitted'), []);
  const parts = zipRead(buffer);
  assert.ok(parts.has('xl/pivotTables/pivotTable1.xml'));
  const cache = parts.get('xl/pivotCache/pivotCacheDefinition1.xml')!.toString();
  assert.match(cache, /worksheetSource ref="A1:B3" sheet="Data"/);
});

test('write error/partial branches: remote+missing+bad image, slicer/filter/pivot edge cases', () => {
  const dir = mkdtempSync(join(tmpdir(), 'gridmd-img-'));
  writeFileSync(join(dir, 'bad.png'), Buffer.from([1, 2, 3, 4])); // not a real png/jpeg
  const src = `---
gridmd: "0.1"
---

# S

\`\`\`{image} at A1
src: "https://example.com/x.png"
\`\`\`

\`\`\`{image} at A3
src: missing.png
\`\`\`

\`\`\`{image} at A5
src: bad.png
\`\`\`

\`\`\`{sparklines} B1:B3
source: C1:D2
\`\`\`

\`\`\`{filter} A10:B12
cols:
  A: { values: [x] }
\`\`\`

\`\`\`{filter} A20:B22
cols:
  A: { top: 3 }
  B: { op: ">", value: 1 }
\`\`\`

\`\`\`{slicer} at D1
for: NoSuchTable
field: x
\`\`\`

\`\`\`{slicer} at E1
kind: timeline
for: NoSuchPivot
field: d
\`\`\`

\`\`\`{pivot} P1 at G1
source: NoSuchSource
values:
  - { field: x }
\`\`\`
`;
  const { report } = convert(src, dir);
  const notEmitted = report.filter((r) => r.action === 'not-emitted').map((r) => r.feature);
  assert.ok(notEmitted.some((f) => /remote/i.test('remote') && f.includes('example.com')));
  assert.ok(report.some((r) => r.note?.includes('file not found')));
  assert.ok(report.some((r) => r.note?.includes('unsupported image format')));
  assert.ok(report.some((r) => r.note?.includes('source rows/cols')));
  assert.ok(report.some((r) => r.feature === '{slicer}' && r.action === 'not-emitted'));
  assert.ok(report.some((r) => r.feature === '{slicer} timeline' && r.action === 'carried'));
  assert.ok(report.some((r) => r.feature === '{pivot} P1' && r.action === 'not-emitted'));
});

test('unknown chart type is carried; ChartEx with no resolvable val is not emitted', () => {
  const carried = convert(`---\ngridmd: "0.1"\n---\n\n# S\n\n\`\`\`{chart} surface at A1\nseries:\n  - { name: X, val: A1:A2 }\n\`\`\`\n`);
  assert.ok(carried.report.some((r) => r.action === 'carried' && r.feature.includes('surface')));
  const noVal = convert(`---\ngridmd: "0.1"\n---\n\n# S\n\n\`\`\`{chart} treemap at A1\nseries:\n  - { name: X, val: "NoTable[c]" }\n\`\`\`\n`);
  assert.ok(noVal.report.some((r) => r.action === 'not-emitted' && r.feature.includes('treemap')));
});

test('chart bound to a missing pivot is carried', () => {
  const r = convert(`---\ngridmd: "0.1"\n---\n\n# S\n\n\`\`\`{chart} column at A1\npivot: Ghost\n\`\`\`\n`);
  assert.ok(r.report.some((x) => x.feature.includes('PivotChart') && x.action === 'carried'));
});

// ---- zip error branches ----

test('zip reader: STORE + DEFLATE members + CRC verification', () => {
  const payload = 'x'.repeat(300);
  const deflated = deflateRawSync(Buffer.from(payload));
  const crc = crc32(Buffer.from(payload));
  // hand-build a one-entry DEFLATE zip
  const name = Buffer.from('d.txt');
  const local = Buffer.alloc(30);
  local.writeUInt32LE(0x04034b50, 0);
  local.writeUInt16LE(20, 4); local.writeUInt16LE(0, 6); local.writeUInt16LE(8, 8);
  local.writeUInt16LE(0, 10); local.writeUInt16LE(0x21, 12);
  local.writeUInt32LE(crc, 14); local.writeUInt32LE(deflated.length, 18); local.writeUInt32LE(payload.length, 22);
  local.writeUInt16LE(name.length, 26); local.writeUInt16LE(0, 28);
  const central = Buffer.alloc(46);
  central.writeUInt32LE(0x02014b50, 0);
  central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6); central.writeUInt16LE(0, 8);
  central.writeUInt16LE(8, 10); central.writeUInt16LE(0, 12); central.writeUInt16LE(0x21, 14);
  central.writeUInt32LE(crc, 16); central.writeUInt32LE(deflated.length, 20); central.writeUInt32LE(payload.length, 24);
  central.writeUInt16LE(name.length, 28); central.writeUInt32LE(0, 42);
  const localBlock = Buffer.concat([local, name, deflated]);
  const centralBlock = Buffer.concat([central, name]);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(1, 8); eocd.writeUInt16LE(1, 10);
  eocd.writeUInt32LE(centralBlock.length, 12); eocd.writeUInt32LE(localBlock.length, 16);
  const zip = Buffer.concat([localBlock, centralBlock, eocd]);
  const back = zipRead(zip);
  assert.equal(back.get('d.txt')!.toString(), payload);
});

test('zip reader: rejects malformed archives', () => {
  assert.throws(() => zipRead(Buffer.alloc(10)), /EOCD missing/);
  // valid STORE zip but corrupt one byte of the payload → CRC mismatch
  const good = zipWrite([{ name: 'a', data: 'hello' }]);
  good[31] = good[31]! ^ 0xff; // flip a payload byte (after 30-byte local header + 1-byte name)
  assert.throws(() => zipRead(good), /crc mismatch/);
});

test('xlsxToGridmd: rejects a zip without a workbook part', () => {
  const zip = zipWrite([{ name: 'random.txt', data: 'hi' }]);
  assert.throws(() => xlsxToGridmd(zip), /workbook\.xml missing/);
});
