// Final edge coverage: calc/validate/parser/model/writer error paths, a
// style-rich round-trip (drives the importer's style + serial readers), and
// patched-xlsx inputs for the external-only reader paths (shared strings,
// theme colours, carried/unrecognized parts).

import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { lint } from '../src/index';
import { buildWorkbookModel } from '../src/xlsx/model';
import { writeXlsx } from '../src/xlsx/write';
import { xlsxToGridmd } from '../src/xlsx/read';
import { zipRead, zipWrite } from '../src/xlsx/zip';
import type { ZipEntry } from '../src/xlsx/zip';
import { createEvaluator, Unsupported } from '../src/calc';
import type { WorkbookModel } from '../src/types';

const doc = (body: string): string => `---\ngridmd: "0.1"\n---\n\n# S1\n\n${body}`;
const modelOf = (src: string, baseDir = '.'): WorkbookModel => {
  const res = lint(src, { mode: 'strict' });
  assert.deepEqual(res.errors, [], `must lint clean: ${res.errors.map((e) => e.msg).join('; ')}`);
  return buildWorkbookModel(res.doc, { baseDir });
};
const evalThrows = (body: string, re?: RegExp): void => {
  assert.throws(() => createEvaluator(modelOf(doc(body))).evaluateCell('S1', 'A1'), re ?? Unsupported);
};

// ---- calc tokenizer / evaluator throws ----

test('calc: tokenizer + evaluator error branches', () => {
  evalThrows('@ A1 =#NOTANERROR', /unknown error literal/);
  evalThrows('@ A1 =~', /unexpected character/);
  const errRef = doc('@ A1 #DIV/0!\n@ B1 =A1');
  assert.deepEqual(createEvaluator(modelOf(errRef)).evaluateCell('S1', 'B1'), { err: '#DIV/0!' });
  const bad = `---\ngridmd: "0.1"\nnames:\n  - { name: Bad, formula: "LAMBDA(1,1)" }\n---\n\n# S1\n\n@ A1 =Bad(2)\n`;
  assert.throws(() => createEvaluator(modelOf(bad)).evaluateCell('S1', 'A1'), /LAMBDA parameter form/);
  const sum2d = doc('@ A1 1\n@ B1 2\n@ A2 3\n@ B2 4\n@ C1 =SUM(A1:B2)');
  assert.equal(createEvaluator(modelOf(sum2d)).evaluateCell('S1', 'C1'), 10); // flatNumbers recursion
});

// ---- parser tryProps scalar/null ----

test('parser: tryProps rejects scalars', async () => {
  const { tryProps } = await import('../src/parser');
  assert.equal(tryProps('5'), null);
  assert.equal(tryProps('null'), null);
});

// ---- validate error/warn branches ----

test('validate: frontmatter + directive + sheet-order branches', () => {
  const errsOf = (src: string): string[] => lint(src).errors.map((e) => e.msg);
  const warnsOf = (src: string): string[] => lint(src).warnings.map((e) => e.msg);

  assert.ok(errsOf('---\ngridmd: 0.1\n---\n\n# S1\n').some((m) => /MAJOR\.MINOR/.test(m)));
  assert.ok(errsOf('---\ngridmd: "0.1"\ndate-system: 1999\n---\n\n# S1\n').some((m) => /date-system must be/.test(m)));
  assert.ok(errsOf('---\ngridmd: "0.1"\nstyles: { s: 5 }\n---\n\n# S1\n').some((m) => /must be a mapping/.test(m)));
  assert.ok(errsOf(doc('```{cf} A1:A3\n- when: "> 1"\n  format: { fill: notacolour }\n```')).some((m) => /cf format\.fill: not a color/.test(m)));
  assert.ok(warnsOf(doc('```{chart} xyz at A1\nseries:\n  - { val: A1:A2 }\n```')).some((m) => /unknown chart type/.test(m)));
  assert.ok(errsOf(doc('```{cf} Other!A1:A3\n- dupes: true\n```')).some((m) => /must name the containing sheet/.test(m)));
  assert.ok(warnsOf('---\ngridmd: "0.1"\n---\n\n# S1\n\n@ A1 1\n\n```{sheet}\nhidden: true\n```\n').some((m) => /should be the first block/.test(m)));
  assert.ok(errsOf(doc('@ A1 5\n  value: 6\n')).some((m) => /inline content and body content/.test(m)));
  assert.ok(warnsOf(doc('@ A1:Z100000 =B1\n')).some((m) => /relative fill over/.test(m)));
});

// ---- model: {raw} without a part= ----

test('model: workbook {raw} without part= is reported not-emitted', () => {
  const res = lint('---\ngridmd: "0.1"\n---\n\n```{raw} ooxml\n<a/>\n```\n\n# S1\n');
  assert.deepEqual(res.errors, []);
  const model = buildWorkbookModel(res.doc, { baseDir: '.' });
  assert.ok(model.report.some((r) => r.action === 'not-emitted' && r.note === 'no part= path'));
});

// ---- writer: filter sort + partial, malformed CF rules ----

test('writer: standalone filter sort + unrecognized criterion + dead CF branches', () => {
  const src = doc([
    '```{grid} A1', '| a | b | c |', '| 1 | 2 | 3 |', '```', '',
    '```{filter} A1:C2', 'cols:', '  A: { values: [x] }', '  B: { note: 1 }',
    'sort:', '  - { col: C, order: asc, by: value }', '```', '',
    '```{cf} A1:A3', '- when: "?? bad"', '```', '',
    '```{cf} B1:B3', '- date: never-period', '```',
  ].join('\n'));
  const { buffer, report } = writeXlsx(modelOf(src));
  assert.ok(report.some((r) => r.action === 'partial' && r.feature.includes('column B')));
  const sheet = zipRead(buffer).get('xl/worksheets/sheet1.xml')!.toString();
  assert.match(sheet, /<sortState/);
});

// ---- style-rich round-trip drives the importer's style + serial readers ----

test('round-trip: rich styles, date/time serials, arrays, custom totals, validation ops', () => {
  const src = `---
gridmd: "0.1"
theme: { colors: { accent1: "#204080" } }
---

# S1

@ A1 "styled" { bold: true, italic: true, underline: double, strike: true, super: true, size: 13, color: "#ff0000", font: Georgia, fill: "#00ff00", pattern: gray-500, align: center, valign: middle, wrap: true, shrink: true, indent: 2, rotation: 45, numfmt: "0.00", locked: false, hidden: true, border: "thin #112233" }
@ A2 2026-03-04
@ A3 09:30:15
@ A4 2026-03-04T09:30
@ B1 =SORT(A2:A3) { spill: B1:B2 }
\`\`\`{spill-cache} B1
| 1 |
| 2 |
\`\`\`

@ D1
  rich:
    - { text: "a", bold: true, size: 12, color: "#123456" }
    - { text: "b" }

@ E1 "linked" { note: "hi", link: "https://x.dev", tip: "t" }

\`\`\`{table} T at A6
total:
  q: =CUSTOM()
cols:
  q: { numfmt: "0.0" }
---
| p | q |
| x | 1.5 |
| y | 2.5 |
\`\`\`

\`\`\`{validation} G1
type: list
source: =A1:A2
\`\`\`

\`\`\`{validation} G2
type: whole
op: ">="
value: 3
\`\`\`

\`\`\`{sparklines} H1:H2
type: line
source: A2:A3
color: accent1
\`\`\`
`;
  const { buffer } = writeXlsx(modelOf(src));
  const { gmd } = xlsxToGridmd(buffer);
  const res = lint(gmd, { mode: 'strict' });
  for (const e of res.errors) console.error(`rt:${e.line}: ${e.msg}`);
  assert.deepEqual(res.errors, []);
  assert.match(gmd, /```\{spill-cache\}/);
  assert.match(gmd, /rich:/);
  assert.match(gmd, /09:30:15/);
});

// ---- patched-xlsx: external-only reader paths ----

function patch(buffer: Buffer, edits: (parts: Map<string, Buffer>) => void): Buffer {
  const parts = zipRead(buffer);
  edits(parts);
  const entries: ZipEntry[] = [...parts.entries()].map(([name, data]) => ({ name, data }));
  return zipWrite(entries);
}

test('reader: shared strings, theme-colour cells, orphan binary, carried chart/pivot', () => {
  const base = writeXlsx(modelOf(doc('@ A1 1\n@ B1 =A1 :: 1\n@ C1 "hi"\n'))).buffer;
  const patched = patch(base, (parts) => {
    // shared strings + a t="s" cell + a theme-coloured, dated styled cell
    parts.set('xl/sharedStrings.xml', Buffer.from(
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">'
      + '<si><t>plain</t></si>'
      + '<si><r><rPr><b/></rPr><t>bo</t></r><r><rPr><i/></rPr><t>ld</t></r></si>'
      + '</sst>'));
    parts.set('xl/styles.xml', Buffer.from(
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>'
      + '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
      + '<font><color theme="4" tint="0.4"/><sz val="12"/><name val="Arial"/></font></fonts>'
      + '<fills count="2"><fill><patternFill patternType="none"/></fill>'
      + '<fill><patternFill patternType="solid"><fgColor rgb="3300FF00"/></patternFill></fill></fills>'
      + '<borders count="2"><border/><border><top style="thin"><color theme="4"/></top></border></borders>'
      + '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
      + '<xf numFmtId="14" fontId="0" fillId="0" borderId="0"/>'
      + '<xf numFmtId="164" fontId="1" fillId="1" borderId="1"><alignment vertical="center" textRotation="90"/><protection locked="0"/></xf></cellXfs>'
      + '<dxfs count="0"/></styleSheet>'));
    parts.set('xl/theme/theme1.xml', Buffer.from(
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x">'
      + '<a:dk1><a:sysClr val="windowText" lastClr="101010"/></a:dk1>'
      + '<a:lt1><a:srgbClr val="FEFEFE"/></a:lt1>'
      + '<a:accent1><a:srgbClr val="204080"/></a:accent1>'
      + '</a:clrScheme></a:themeElements></a:theme>'));
    parts.set('xl/worksheets/sheet1.xml', Buffer.from(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      + '<sheetData>'
      + '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c>'
      + '<c r="C1" s="1"><v>46082</v></c><c r="D1" s="2"><v>0.5</v></c></row>'
      + '</sheetData></worksheet>'));
    // an orphan binary part → carried as base64
    parts.set('xl/customData/orphan.dat', Buffer.from([1, 2, 3, 4, 5]));
    // corrupt the chart part so reverseChart returns null → carried
    for (const name of [...parts.keys()]) {
      if (/xl\/charts\/chart\d+\.xml$/.test(name)) parts.set(name, Buffer.from('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea/></c:chart></c:chartSpace>'));
    }
  });
  const { gmd, report } = xlsxToGridmd(patched);
  assert.deepEqual(lint(gmd, { mode: 'strict' }).errors, []);
  assert.match(gmd, /bold: true/);           // shared rich reversed
  assert.match(gmd, /accent1/);              // theme colour reversed
  assert.ok(report.some((r) => r.note?.includes('not reverse-parsed'))); // orphan carried (base64)
});

test('reader: unrecognized chart sheet + pivot → carried', () => {
  // A workbook with a chart sheet whose drawing has no recognizable chart.
  const wb = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Chart" sheetId="1" r:id="rId1"/></sheets></workbook>';
  const wbRels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>';
  const cs = '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><drawing r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></chartsheet>';
  const csRels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>';
  const drawing = '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:absoluteAnchor><xdr:graphicFrame><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></xdr:graphicFrame></xdr:absoluteAnchor></xdr:wsDr>';
  const zip = zipWrite([
    { name: 'xl/workbook.xml', data: wb },
    { name: 'xl/_rels/workbook.xml.rels', data: wbRels },
    { name: 'xl/chartsheets/sheet1.xml', data: cs },
    { name: 'xl/chartsheets/_rels/sheet1.xml.rels', data: csRels },
    { name: 'xl/drawings/drawing1.xml', data: drawing },
  ]);
  const { gmd, report } = xlsxToGridmd(zip);
  assert.ok(report.some((r) => r.feature.includes('chart sheet') && r.action === 'carried'));
  assert.match(gmd, /# Chart/);
});
