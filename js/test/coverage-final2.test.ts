// The last importer branches: formula-with-note, multiline values, style-only
// cells, standalone AutoFilter reversal, and carried worksheet-chart /
// timeline / unknown-CF-rule forms.

import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { lint } from '../src/index';
import { buildWorkbookModel } from '../src/xlsx/model';
import { writeXlsx } from '../src/xlsx/write';
import { xlsxToGridmd } from '../src/xlsx/read';
import { zipRead, zipWrite } from '../src/xlsx/zip';
import type { ZipEntry } from '../src/xlsx/zip';
import type { WorkbookModel } from '../src/types';

const doc = (body: string): string => `---\ngridmd: "0.1"\n---\n\n# S1\n\n${body}`;
const modelOf = (src: string): WorkbookModel => {
  const res = lint(src, { mode: 'strict' });
  assert.deepEqual(res.errors, [], `must lint clean: ${res.errors.map((e) => e.msg).join('; ')}`);
  return buildWorkbookModel(res.doc, { baseDir: '.' });
};
const roundtrip = (src: string): string => {
  const { buffer } = writeXlsx(modelOf(src));
  const { gmd } = xlsxToGridmd(buffer);
  const res = lint(gmd, { mode: 'strict' });
  for (const e of res.errors) console.error(`rt:${e.line}: ${e.msg}`);
  assert.deepEqual(res.errors, []);
  return gmd;
};
function patch(buffer: Buffer, edits: (parts: Map<string, Buffer>) => void): Buffer {
  const parts = zipRead(buffer);
  edits(parts);
  const entries: ZipEntry[] = [...parts.entries()].map(([name, data]) => ({ name, data }));
  return zipWrite(entries);
}

test('round-trip: formula+note, multiline value, style-only cell', () => {
  const gmd = roundtrip(doc([
    '@ A1 5',
    '@ A2 =SUM(A1:A1) :: 5 { note: "check this" }',
    '@ A3',
    '  value: |',
    '    line one',
    '    line two',
    '@ Z1 { bold: true }',
  ].join('\n')));
  assert.match(gmd, /note: \|/);
  assert.match(gmd, /value: \|/);
  assert.match(gmd, /line one/);
  assert.match(gmd, /@ Z1 \{ bold: true \}/);
});

test('round-trip: standalone AutoFilter with values, top, custom op + sort', () => {
  const gmd = roundtrip(doc([
    '```{grid} A1', '| a | b | c |', '| 1 | 2 | 3 |', '| 4 | 5 | 6 |', '```', '',
    '```{filter} A1:C3',
    'cols:',
    '  A: { values: [1, 4] }',
    '  B: { top: 2 }',
    '  C: { op: ">", value: 2 }',
    'sort:',
    '  - { col: A, order: desc, by: value }',
    '```',
  ].join('\n')));
  assert.match(gmd, /```\{filter\} A1:C3/);
  assert.match(gmd, /sort:/);
});

test('reader: carried worksheet chart, malformed timeline, unknown CF rule type', () => {
  // (a) a worksheet-embedded chart, corrupted so reverseChart returns null
  const chartDoc = doc('```{grid} A1\n| x |\n| 1 |\n```\n\n```{chart} column at C1\nseries:\n  - { name: S, val: A2:A2 }\n```\n');
  const chartPatched = patch(writeXlsx(modelOf(chartDoc)).buffer, (parts) => {
    for (const name of [...parts.keys()]) {
      if (/xl\/charts\/chart\d+\.xml$/.test(name)) parts.set(name, Buffer.from('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea/></c:chart></c:chartSpace>'));
    }
  });
  assert.ok(xlsxToGridmd(chartPatched).report.some((r) => r.note?.includes('unrecognized chart form')));

  // (b) a pivot timeline whose cache lost its pivotTable reference
  const tlDoc = `---\ngridmd: "0.1"\n---\n\n# D\n\n\`\`\`{table} Sales at A1\n---\n| region | amount | when |\n| AU | 10 | 2026-01-05 |\n| NZ | 20 | 2026-02-11 |\n\`\`\`\n\n\`\`\`{pivot} P at E1\nsource: Sales\nrows:\n  - { field: region }\nvalues:\n  - { field: amount, agg: sum }\n\`\`\`\n\n\`\`\`{slicer} at H1\nkind: timeline\nfor: P\nfield: when\nlevel: months\n\`\`\`\n`;
  const tlPatched = patch(writeXlsx(modelOf(tlDoc)).buffer, (parts) => {
    for (const name of [...parts.keys()]) {
      if (/timelineCaches\/timelineCache\d+\.xml$/.test(name)) parts.set(name, Buffer.from('<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="x" sourceName="when"/>'));
    }
  });
  assert.ok(xlsxToGridmd(tlPatched).report.some((r) => r.note?.includes('unrecognized timeline form')));

  // (c) a worksheet CF rule of an unmapped type
  const cfPatched = patch(writeXlsx(modelOf(doc('@ A1 1'))).buffer, (parts) => {
    const s = parts.get('xl/worksheets/sheet1.xml')!.toString()
      .replace('</sheetData>', '</sheetData><conditionalFormatting sqref="A1:A2"><cfRule type="someUnknownType" priority="1"/></conditionalFormatting>');
    parts.set('xl/worksheets/sheet1.xml', Buffer.from(s));
  });
  assert.ok(xlsxToGridmd(cfPatched).report.some((r) => r.feature.includes('cf rule type') && r.action === 'carried'));
});
