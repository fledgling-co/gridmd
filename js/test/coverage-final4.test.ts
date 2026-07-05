// The last coverable branches: spill-not-at-anchor, area chart, table
// color-sort, pivot grand-totals reversal, column-tiled sparklines, tableAt
// exhaustion, propsText-with-bad-props fallthrough, out-of-range col range,
// and a range-source pivot with a bad header row.

import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { lint } from '../src/index';
import { parseDocument } from '../src/parser';
import { parseTarget } from '../src/refs';
import { buildWorkbookModel } from '../src/xlsx/model';
import { writeXlsx } from '../src/xlsx/write';
import { xlsxToGridmd } from '../src/xlsx/read';
import { zipRead } from '../src/xlsx/zip';
import { createEvaluator, Unsupported } from '../src/calc';
import type { AtBlock, WorkbookModel } from '../src/types';

const doc = (body: string): string => `---\ngridmd: "0.1"\n---\n\n# S1\n\n${body}`;
const modelOf = (src: string): WorkbookModel => {
  const res = lint(src, { mode: 'strict' });
  assert.deepEqual(res.errors, [], `must lint clean: ${res.errors.map((e) => e.msg).join('; ')}`);
  return buildWorkbookModel(res.doc, { baseDir: '.' });
};

test('validate: spill/array range must start at the anchor cell', () => {
  const errs = lint(doc('@ A1 =SORT(C1:C3) { spill: B2:B4 }\n')).errors.map((e) => e.msg);
  assert.ok(errs.some((m) => /range must start at the anchor cell/.test(m)));
});

test('refs: out-of-range column range returns null (COLRANGE bound check)', () => {
  assert.equal(parseTarget('ZZZ:ZZZ'), null); // ZZZ > MAX_COL
});

test('parser: propsText present but rejected by tryProps falls through to scalar', () => {
  const d = parseDocument(doc('@ A1 x {Bad Key: 1}\n'));
  const block = d.sheets[0]!.blocks[0] as AtBlock;
  assert.equal(block.props, null);
  assert.equal(block.scalarText, 'x {Bad Key: 1}');
});

test('calc: tableAt exhausts all tables then reports outside-a-table', () => {
  const src = doc('```{table} T at A1\n---\n| q |\n| 1 |\n```\n\n@ Z9 =[@q]\n');
  assert.throws(() => createEvaluator(modelOf(src)).evaluateCell('S1', 'Z9'), Unsupported);
});

test('chart: area chart block emits', () => {
  const { buffer } = writeXlsx(modelOf(doc('```{grid} A1\n| c | v |\n| a | 1 |\n| b | 2 |\n```\n\n```{chart} area at D1\nseries:\n  - { name: A, cat: A2:A3, val: B2:B3 }\n```\n')));
  const zr = zipRead(buffer);
  const charts = [...zr.keys()].filter((k) => /chart\d+\.xml$/.test(k)).map((k) => zr.get(k)!.toString()).join('');
  assert.match(charts, /areaChart/);
});

test('model: table sort-by-color level reports partial', () => {
  const m = modelOf(doc('```{table} T at A1\nsort:\n  - { col: q, order: asc, by: color }\n---\n| q |\n| 1 |\n| 2 |\n```\n'));
  assert.ok(m.report.some((r) => r.feature.includes('color-sort') && r.action === 'partial'));
});

test('reader: pivot grand-totals disabled round-trips', () => {
  const src = `---\ngridmd: "0.1"\n---\n\n# D\n\n\`\`\`{table} Sales at A1\n---\n| region | amount |\n| AU | 10 |\n| NZ | 20 |\n\`\`\`\n\n\`\`\`{pivot} P at E1\nsource: Sales\nrows:\n  - { field: region }\nvalues:\n  - { field: amount, agg: sum }\ngrand-totals: { rows: false, cols: false }\n\`\`\`\n`;
  const { buffer } = writeXlsx(modelOf(src));
  const { gmd } = xlsxToGridmd(buffer);
  assert.deepEqual(lint(gmd, { mode: 'strict' }).errors, []);
  assert.match(gmd, /grand-totals:/);
});

test('writer: column-tiled sparkline source', () => {
  // 3 target cells across a row; a source with 3 columns × 2 rows tiles by column
  const src = doc('```{grid} A1\n| 1 | 2 | 3 |\n| 4 | 5 | 6 |\n```\n\n```{sparklines} A5:C5\nsource: A1:C2\n```\n');
  const { buffer } = writeXlsx(modelOf(src));
  const sheet = zipRead(buffer).get('xl/worksheets/sheet1.xml')!.toString();
  assert.match(sheet, /x14:sparklineGroup/);
});

test('pivot: range source with a non-string header is not emitted', () => {
  const src = `---\ngridmd: "0.1"\n---\n\n# Data\n\n\`\`\`{grid} A1\n| 10 | 20 |\n| x | y |\n\`\`\`\n\n\`\`\`{pivot} P at E1\nsource: Data!A1:B2\nvalues:\n  - { field: x }\n\`\`\`\n`;
  const { report } = writeXlsx(modelOf(src));
  assert.ok(report.some((r) => r.feature.includes('{pivot} P') && r.action === 'not-emitted'));
});
