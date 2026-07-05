// The last uncovered branches: YAML toJS failure paths, the evaluateFormula
// seam, non-LAMBDA formula names, the dump table sort, degenerate CF rules,
// top-10 filters, name-ref chart series, and carried malformed pivot/slicer.

import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { lint } from '../src/index';
import { parseDocument, tryProps } from '../src/parser';
import { buildWorkbookModel } from '../src/xlsx/model';
import { dumpModel } from '../src/dump';
import { writeXlsx } from '../src/xlsx/write';
import { xlsxToGridmd } from '../src/xlsx/read';
import { zipRead, zipWrite } from '../src/xlsx/zip';
import type { ZipEntry } from '../src/xlsx/zip';
import { createEvaluator } from '../src/calc';
import type { WorkbookModel } from '../src/types';

const doc = (body: string): string => `---\ngridmd: "0.1"\n---\n\n# S1\n\n${body}`;
const modelOf = (src: string): WorkbookModel => {
  const res = lint(src, { mode: 'strict' });
  assert.deepEqual(res.errors, [], `must lint clean: ${res.errors.map((e) => e.msg).join('; ')}`);
  return buildWorkbookModel(res.doc, { baseDir: '.' });
};
function patch(buffer: Buffer, edits: (parts: Map<string, Buffer>) => void): Buffer {
  const parts = zipRead(buffer);
  edits(parts);
  const entries: ZipEntry[] = [...parts.entries()].map(([name, data]) => ({ name, data }));
  return zipWrite(entries);
}

// ---- YAML toJS failure branches (an unresolved alias parses cleanly but
//      throws in toJS) ----

test('parser: unresolved alias makes toJS throw (both parse paths)', () => {
  const d = parseDocument('---\ngridmd: "0.1"\n---\n\n# S1\n\n```{page}\na: *undef\n```\n');
  assert.ok(d.errors.some((e) => /YAML: /.test(e.msg))); // toJS catch pushed a YAML error
  assert.equal(tryProps('a: *undef'), null);             // tryProps catch → null
});

// ---- calc: evaluateFormula seam + non-LAMBDA formula names + more funcs ----

test('calc: evaluateFormula seam + CONCATENATE/LEN', () => {
  const ev = createEvaluator(modelOf(doc('@ A1 3\n@ A2 4')));
  assert.equal(ev.evaluateFormula('A1+A2', { sheet: 'S1' }), 7);
  assert.equal(ev.evaluateFormula('CONCATENATE("x",A1)', { sheet: 'S1' }), 'x3');
  assert.equal(ev.evaluateFormula('LEN("abcd")', { sheet: 'S1' }), 4);
});

test('calc: non-LAMBDA formula name (referenced and mis-called)', () => {
  const src = `---\ngridmd: "0.1"\nnames:\n  - { name: Plain, formula: "S1!A1*2" }\n---\n\n# S1\n\n@ A1 5\n@ B1 =Plain\n@ B2 =Plain(1)\n`;
  const ev = createEvaluator(modelOf(src));
  assert.equal(ev.evaluateCell('S1', 'B1'), 10);                 // formula name, no args
  assert.throws(() => ev.evaluateCell('S1', 'B2'), /call of non-LAMBDA/); // formula name, called
});

// ---- dump: two tables in one sheet exercises the name sort comparator ----

test('dumpModel: multiple tables sort by name', () => {
  const src = doc([
    '```{table} Zebra at A1', '---', '| a |', '| 1 |', '```', '',
    '```{table} Alpha at C1', '---', '| b |', '| 2 |', '```',
  ].join('\n'));
  const parsed = JSON.parse(dumpModel(modelOf(src)));
  assert.deepEqual(parsed.sheets[0].tables.map((t: { name: string }) => t.name), ['Alpha', 'Zebra']);
});

// ---- writer: degenerate CF rule + top-10 as the first filter column ----

test('writer: empty CF rule emits nothing; top-10 first filter column', () => {
  const m = modelOf(doc('@ A1 1'));
  m.sheets[0]!.cf.push({ sqref: 'A1:A2', rules: [{}], line: 1 });
  const { buffer } = writeXlsx(m);
  const sheet = zipRead(buffer).get('xl/worksheets/sheet1.xml')!.toString();
  assert.doesNotMatch(sheet, /conditionalFormatting/); // empty rule → no CF element

  const filt = writeXlsx(modelOf(doc(['```{grid} A1', '| a | b |', '| 1 | 2 |', '```', '',
    '```{filter} A1:B2', 'cols:', '  A: { top: "50%" }', '```'].join('\n'))));
  const fsheet = zipRead(filt.buffer).get('xl/worksheets/sheet1.xml')!.toString();
  assert.match(fsheet, /<top10 val="50" percent="1"\/>/);
});

// ---- reader: name-ref chart series ----

test('round-trip: chart series with a name-ref', () => {
  const src = doc([
    '```{grid} A1', '| n | x | y |', '| Series X | 1 | 5 |', '| b | 2 | 6 |', '```', '',
    '```{chart} column at G1', 'series:', '  - { name-ref: "=A2", cat: B2:B3, val: C2:C3 }', '```',
  ].join('\n'));
  const { buffer } = writeXlsx(modelOf(src));
  const { gmd } = xlsxToGridmd(buffer);
  assert.deepEqual(lint(gmd, { mode: 'strict' }).errors, []);
  assert.match(gmd, /name-ref:/);
});

// ---- reader: carried malformed pivot + slicer (external forms) ----

test('reader: malformed pivot and slicer parts are carried', () => {
  const src = `---\ngridmd: "0.1"\n---\n\n# Data\n\n\`\`\`{table} Sales at A1\n---\n| region | amount |\n| AU | 10 |\n| NZ | 20 |\n\`\`\`\n\n\`\`\`{pivot} P at E1\nsource: Sales\nrows:\n  - { field: region }\nvalues:\n  - { field: amount, agg: sum }\n\`\`\`\n\n\`\`\`{slicer} at H1\nfor: Sales\nfield: region\n\`\`\`\n`;
  const base = writeXlsx(modelOf(src)).buffer;
  const patched = patch(base, (parts) => {
    for (const name of [...parts.keys()]) {
      if (/pivotTables\/pivotTable\d+\.xml$/.test(name)) parts.set(name, Buffer.from('<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'));
      if (/slicerCaches\/slicerCache\d+\.xml$/.test(name)) parts.set(name, Buffer.from('<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer_region" sourceName="region"/>'));
    }
  });
  const { report } = xlsxToGridmd(patched);
  assert.ok(report.some((r) => r.feature.includes('pivotTable') && r.action === 'carried'));
  assert.ok(report.some((r) => r.feature.includes('slicer') && r.action === 'carried'));
});
