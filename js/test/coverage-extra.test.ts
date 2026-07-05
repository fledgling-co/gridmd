// Remaining branch coverage: calc value coercions, parser/validate error
// paths, refs bounds, model body-content shapes, and the XML entity decoder.

import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { lint } from '../src/index';
import { parseDocument } from '../src/parser';
import { parseTarget } from '../src/refs';
import { buildWorkbookModel } from '../src/xlsx/model';
import { createEvaluator, Unsupported } from '../src/calc';
import { parseXml, decodeEntities, attr, textOf, findDeep } from '../src/xml';
import type { WorkbookModel } from '../src/types';

const doc = (body: string): string => `---\ngridmd: "0.1"\n---\n\n# S1\n\n${body}`;
const modelOf = (src: string): WorkbookModel => {
  const res = lint(src, { mode: 'strict' });
  assert.deepEqual(res.errors, [], `must lint clean: ${res.errors.map((e) => e.msg).join('; ')}`);
  return buildWorkbookModel(res.doc, { baseDir: '.' });
};
const evalIn = (src: string, ref: string): unknown => createEvaluator(modelOf(src)).evaluateCell('S1', ref);

// ---- calc value coercions & function edges ----

test('calc: quoted sheet name in a formula', () => {
  const src = `---\ngridmd: "0.1"\n---\n\n# S 1\n\n@ A1 7\n\n# S2\n\n@ A1 ='S 1'!A1*2\n`;
  assert.equal(createEvaluator(modelOf(src)).evaluateCell('S2', 'A1'), 14);
});

test('calc: self-column [@col] structured item ref', () => {
  const src = doc([
    '```{table} T at A1', '---', '| q | r |', '| 2 | =[@q]*10 |', '| 3 | =[@q]*10 |', '```',
  ].join('\n'));
  assert.equal(evalIn(src, 'B2'), 20);
});

test('calc: numeric coercion of text and boolean operands', () => {
  assert.equal(evalIn(doc('@ A1 "5"\n@ B1 =A1*2'), 'B1'), 10);            // text→number
  // non-numeric text poisons the arithmetic (legacy en-US semantics)
  assert.equal(typeof evalIn(doc('@ A1 abc\n@ B1 =A1+1'), 'B1'), 'string');
  assert.equal(evalIn(doc('@ A1 =1=1\n@ B1 =A1+0'), 'B1'), 1);           // boolean→number
  assert.equal(evalIn(doc('@ A1 TRUE\n@ B1 =A1'), 'B1'), true);          // scalarValue boolean
});

test('calc: aggregate over ranges (flatAll recursion) + error-in-range throws', () => {
  assert.equal(evalIn(doc('@ A1 1\n@ A2 2\n@ B1 3\n@ B2 4\n@ C1 =COUNTA(A1:B2)'), 'C1'), 4);
  const ev = createEvaluator(modelOf(doc('@ A1 =1/0\n@ A2 5\n@ B1 =SUM(A1:A2)')));
  assert.throws(() => ev.evaluateCell('S1', 'B1'), Unsupported); // error in a summed range
});

test('calc: num() array passthrough + unknown table column', () => {
  // ABS of a range → num() receives an array (poisons to NaN, no throw)
  assert.ok(Number.isNaN(evalIn(doc('@ A1 1\n@ A2 2\n@ B1 =ABS(A1:A2)') as string, 'B1') as number));
  const badCol = doc(['```{table} T at A1', '---', '| q |', '| 1 |', '```', '', '@ C1 =SUM(T[nope])'].join('\n'));
  assert.throws(() => createEvaluator(modelOf(badCol)).evaluateCell('S1', 'C1'), /unknown table column/);
});

test('calc: array constants with text, boolean and invalid elements', () => {
  assert.equal(evalIn(doc('@ A1 =CONCAT({"a","b"})'), 'A1'), 'ab');    // string elements
  assert.equal(evalIn(doc('@ A1 ={TRUE,FALSE}'), 'A1'), true);         // boolean element (spill top-left)
  assert.throws(() => createEvaluator(modelOf(doc('@ A1 ={notvalid}'))).evaluateCell('S1', 'A1'), Unsupported);
});

test('calc: calling a non-LAMBDA defined name throws', () => {
  const src = `---\ngridmd: "0.1"\nnames:\n  - { name: K, value: "5" }\n---\n\n# S1\n\n@ A1 =K(1)\n`;
  assert.throws(() => createEvaluator(modelOf(src)).evaluateCell('S1', 'A1'), /call of non-LAMBDA/);
});

// ---- parser error branches ----

test('parser: bad rows, table without separator, YAML error in a fence', () => {
  const badGrid = parseDocument(doc('```{grid} A1\n| 1 |\nnotarow\n```\n'));
  assert.ok(badGrid.errors.some((e) => /expected a pipe row/.test(e.msg)));
  const noSep = parseDocument(doc('```{table} T at A1\n| a |\n| 1 |\n```\n'));
  assert.ok(noSep.errors.some((e) => /requires a `---`-separated/.test(e.msg)));
  const badYaml = parseDocument(doc('```{page}\nx: {a: 1\n```\n'));
  assert.ok(badYaml.errors.some((e) => /^YAML:/.test(e.msg)));
});

test('parser: pure-props inline directive', () => {
  const d = parseDocument(doc('@ A1 { bold: true }\n'));
  assert.deepEqual(d.errors, []);
  const block = d.sheets[0]!.blocks[0]!;
  assert.equal(block.type, 'at');
});

// ---- refs bounds ----

test('parseTarget: out-of-range row range + unmatched range → null', () => {
  assert.equal(parseTarget('1:2000000'), null); // row beyond MAX_ROW
  assert.equal(parseTarget('A1:2'), null);       // mixed cell/row
});

// ---- validate error branches ----

test('validate: frontmatter + directive error branches', () => {
  const errsOf = (src: string): string[] => lint(src).errors.map((e) => e.msg);

  const fm = errsOf(`---\ngridmd: "0.1"\ncalc: { mode: nope }\nnames:\n  - { ref: "S1!$A$1" }\n  - { name: Dup, value: "1" }\n  - { name: Dup, value: "2" }\ntheme: { colors: { accent1: "notacolour" } }\n---\n\n# S1\n`);
  assert.ok(fm.some((m) => /calc.mode must be/.test(m)));
  assert.ok(fm.some((m) => /names entries require a name/.test(m)));
  assert.ok(fm.some((m) => /duplicate defined name/.test(m)));
  assert.ok(fm.some((m) => /theme color .* must be #RRGGBB/.test(m)));

  const badVal = errsOf(doc('```{validation} A1\ntype: whole\nerror: { style: nope }\n```\n'));
  assert.ok(badVal.some((m) => /error.style must be/.test(m)));

  const badSeries = errsOf(doc('```{chart} column at A1\nseries:\n  - { val: A1:A2, color: notacolour }\n```\n'));
  assert.ok(badSeries.some((m) => /color: not a color/.test(m)));

  const badSheetMeta = errsOf(doc('```{sheet}\nkind: nope\ntab-color: notacolour\nhidden: maybe\nfreeze: nonsense\ncols: { "1": 5 }\nrows: { "x": 5 }\n```\n'));
  assert.ok(badSheetMeta.some((m) => /kind must be/.test(m)));
  assert.ok(badSheetMeta.some((m) => /tab-color: not a color/.test(m)));
  assert.ok(badSheetMeta.some((m) => /hidden must be/.test(m)));
  assert.ok(badSheetMeta.some((m) => /freeze: must be a cell/.test(m)));
  assert.ok(badSheetMeta.some((m) => /cols key must be/.test(m)));
  assert.ok(badSheetMeta.some((m) => /rows key must be/.test(m)));

  const badProps = errsOf(doc('@ A1:B2 =C1 { merge: yes }\n@ C1 "x" { unknownprop: 1, rich: notalist, control: nope }\n'));
  assert.ok(badProps.some((m) => /merge: only `true`/.test(m)));
  assert.ok(badProps.some((m) => /rich: must be a list/.test(m)));
  assert.ok(badProps.some((m) => /control: unknown control/.test(m)));

  // warnings: unknown frontmatter/sheet/property keys + formula-without-cache
  const warned = lint(`---\ngridmd: "0.1"\nunknownfm: 1\n---\n\n# S1\n\n@ A1\n  formula: =B1\n`);
  assert.ok(warned.warnings.some((w) => /unknown frontmatter key/.test(w.msg)));
  assert.ok(warned.warnings.some((w) => /formula without a cached value/.test(w.msg)));
});

// ---- model body-content shapes ----

test('model: @ body formula/value/rich/entity + spill on cell', () => {
  const m = modelOf(doc([
    '@ A1',
    '  formula: =SUM(B1:B2)',
    '  value: 3',
    '  spill: A1:A2',
    '@ C1',
    '  value: 2026-01-01',
    '@ C2',
    '  value: hello',
    '@ D1',
    '  rich:',
    '    - { text: hi }',
  ].join('\n')));
  const s = m.sheets[0]!;
  assert.ok(s.cells.get('1,1')?.content?.arrayRef);
  assert.equal(s.cells.get('3,1')?.content?.scalar?.kind, 'date');
  assert.equal(s.cells.get('3,2')?.content?.scalar?.kind, 'text');
  assert.ok(s.cells.get('4,1')?.content?.rich);
});

// ---- xml entity decoder ----

test('xml: decodeEntities (decimal, hex, named, unknown) + helpers', () => {
  assert.equal(decodeEntities('&#65;&#x42;&amp;&nope;'), 'AB&&nope;');
  const node = parseXml('<root a="x&#38;y"><child>t&#x26;u</child></root>');
  assert.equal(attr(node, 'a'), 'x&y');
  assert.equal(textOf(node), 't&u');
  assert.equal(findDeep(node, 'child')!.name, 'child');
  assert.equal(findDeep(node, 'absent'), null);
  assert.equal(attr(node, 'missing'), undefined);
});
