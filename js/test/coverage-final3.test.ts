// Final real gaps: tableAt loop continuation, invalid-props fallthrough, a
// bare (content-less, style-less) imported cell, a non-FF ARGB colour, a line
// series colour (lineColor), and a reversed picture with alt text.

import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { lint } from '../src/index';
import { parseDocument } from '../src/parser';
import { buildWorkbookModel } from '../src/xlsx/model';
import { writeXlsx } from '../src/xlsx/write';
import { xlsxToGridmd } from '../src/xlsx/read';
import { zipRead, zipWrite } from '../src/xlsx/zip';
import type { ZipEntry } from '../src/xlsx/zip';
import { createEvaluator } from '../src/calc';
import type { AtBlock, WorkbookModel } from '../src/types';

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
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC';

test('calc: tableAt skips a non-matching table before the match', () => {
  const src = doc([
    '```{table} TabOne at A1', '---', '| a |', '| 1 |', '```', '',
    '```{table} TabTwo at C1', '---', '| q | r |', '| 5 | =[@q]*2 |', '```',
  ].join('\n'));
  assert.equal(createEvaluator(modelOf(src)).evaluateCell('S1', 'D2'), 10);
});

test('parser: inline braces that are not valid props fall through to a scalar', () => {
  const d = parseDocument(doc('@ A1 {Bad Key: 1}\n'));
  assert.deepEqual(d.errors, []);
  const block = d.sheets[0]!.blocks[0] as AtBlock;
  assert.equal(block.props, null); // key is not an identifier → tryProps null → fall through
  assert.ok(block.scalarText?.startsWith('{'));
});

test('reader: bare cell, non-FF ARGB colour, and orphaned styles', () => {
  const base = writeXlsx(modelOf(doc('@ A1 1\n'))).buffer;
  const patched = patch(base, (parts) => {
    parts.set('xl/styles.xml', Buffer.from(
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font>'
      + '<font><color rgb="3311AA22"/></font>'      // non-FF 8-char ARGB → #11AA2233
      + '<font><color rgb="00FF00"/></font>'        // 6-char RGB → #00FF00
      + '<font><color theme="99"/></font></fonts>'  // out-of-range theme index → no colour
      + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
      + '<borders count="1"><border/></borders>'
      + '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
      + '<xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>'
      + '<xf numFmtId="0" fontId="2" fillId="0" borderId="0"/>'
      + '<xf numFmtId="0" fontId="3" fillId="0" borderId="0"/></cellXfs><dxfs count="0"/></styleSheet>'));
    parts.set('xl/worksheets/sheet1.xml', Buffer.from(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>2</v></c><c r="C1" s="3"><v>3</v></c><c r="Z9"/></row></sheetData></worksheet>'));
  });
  const { gmd } = xlsxToGridmd(patched);
  assert.deepEqual(lint(gmd, { mode: 'strict' }).errors, []);
  assert.match(gmd, /#11AA2233/); // 8-char ARGB→#RRGGBBAA
  assert.match(gmd, /#00FF00/);   // 6-char RGB
});

test('round-trip: line series colour (lineColor) + picture with alt', () => {
  const src = doc([
    '```{grid} A1', '| cat | v |', '| a | 1 |', '| b | 2 |', '```', '',
    '```{chart} line at D1', 'series:', '  - { name: L, cat: A2:A3, val: B2:B3, color: "#ff8800" }', '```', '',
    `\`\`\`{image} at D12`, `src: "data:image/png;base64,${PNG}"`, 'alt: mypic', '```',
  ].join('\n'));
  const { buffer } = writeXlsx(modelOf(src));
  const { gmd } = xlsxToGridmd(buffer);
  assert.deepEqual(lint(gmd, { mode: 'strict' }).errors, []);
  assert.match(gmd, /color: '?#ff8800'?/i);
  assert.match(gmd, /alt: mypic/);
});
