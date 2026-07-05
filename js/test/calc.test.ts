import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { lint } from '../src/index';
import { buildWorkbookModel } from '../src/xlsx/model';
import { createEvaluator, verifyCachedValues } from '../src/calc';

const examplePath = fileURLToPath(new URL('../../examples/quarterly-report.gmd', import.meta.url));

const modelOf = (src: string, baseDir = '.') => {
  const res = lint(src, { mode: 'strict' });
  assert.deepEqual(res.errors, [], 'fixture must lint clean');
  return buildWorkbookModel(res.doc, { baseDir });
};

const evalIn = (src: string, sheet: string, ref: string) => createEvaluator(modelOf(src)).evaluateCell(sheet, ref);
const doc = (body: string): string => `---\ngridmd: "0.1"\n---\n\n# S1\n\n${body}`;

test('arithmetic, precedence, unary, percent, concat, comparison', () => {
  assert.equal(evalIn(doc('@ A1 =1+2*3'), 'S1', 'A1'), 7);
  assert.equal(evalIn(doc('@ A1 =(1+2)*3^2'), 'S1', 'A1'), 27);
  assert.equal(evalIn(doc('@ A1 =-3^2'), 'S1', 'A1'), 9); // (-3)^2, Excel semantics via unary binding
  assert.equal(evalIn(doc('@ A1 =50%'), 'S1', 'A1'), 0.5);
  assert.equal(evalIn(doc('@ A1 ="a"&"b"&1'), 'S1', 'A1'), 'ab1');
  assert.equal(evalIn(doc('@ A1 =2>1'), 'S1', 'A1'), true);
  assert.equal(evalIn(doc('@ A1 ="Apple"="APPLE"'), 'S1', 'A1'), true); // case-insensitive
});

test('cell refs, ranges, cross-sheet, blank coercion', () => {
  const src = `---\ngridmd: "0.1"\n---\n\n# S1\n\n@ A1 10\n@ A2 20\n@ B1 =SUM(A1:A3)\n@ B2 =A1+A9\n\n# S2\n\n@ A1 =S1!A2*2\n`;
  assert.equal(evalIn(src, 'S1', 'B1'), 30);
  assert.equal(evalIn(src, 'S1', 'B2'), 10); // blank A9 → 0
  assert.equal(evalIn(src, 'S2', 'A1'), 40);
});

test('IF / string functions / division errors', () => {
  assert.equal(evalIn(doc('@ A1 =IF(2>1,"yes","no")'), 'S1', 'A1'), 'yes');
  assert.equal(evalIn(doc('@ A1 =LEFT("Widget",3)'), 'S1', 'A1'), 'Wid');
  assert.deepEqual(evalIn(doc('@ A1 =1/0'), 'S1', 'A1'), { err: '#DIV/0!' });
  assert.equal(evalIn(doc('@ A1 =IF(A2="x :: y",1,2)'), 'S1', 'A1'), 2); // :: inside string survives
});

test('structured references + SUBTOTAL inside a table', () => {
  const src = doc([
    '```{table} T at A1',
    'total:',
    '  q: =SUBTOTAL(109,[q])',
    '---',
    '| p | q |',
    '| a | 2 |',
    '| b | 3 |',
    '```',
    '',
    '@ D1 =SUM(T[q])',
    '@ D2 =SUBTOTAL(109,T[q])',
  ].join('\n'));
  assert.equal(evalIn(src, 'S1', 'D1'), 5);
  assert.equal(evalIn(src, 'S1', 'D2'), 5);
  assert.equal(evalIn(src, 'S1', 'B4'), 5); // the total row itself
});

test('defined names: ref, constant, LAMBDA', () => {
  const src = `---\ngridmd: "0.1"\nnames:\n  - { name: Rate, ref: "S1!$A$1" }\n  - { name: Half, value: "0.5" }\n  - { name: FtoC, formula: "LAMBDA(F,(F-32)*5/9)" }\n---\n\n# S1\n\n@ A1 0.3\n@ B1 =Rate*10\n@ B2 =Half*4\n@ B3 =FtoC(212)\n`;
  assert.equal(evalIn(src, 'S1', 'B1'), 3);
  assert.equal(evalIn(src, 'S1', 'B2'), 2);
  assert.equal(evalIn(src, 'S1', 'B3'), 100);
});

test('SORT + UNIQUE spill anchors resolve to the top-left', () => {
  const src = doc('@ A1 b\n@ A2 a\n@ A3 b\n@ C1 =SORT(UNIQUE(A1:A3)) { spill: C1:C2 }\n');
  assert.equal(evalIn(src, 'S1', 'C1'), 'a');
});

test('entity field access', () => {
  const src = doc('@ B2\n  entity: { type: stock, id: "X:MSFT", text: MSFT }\n  fields: { Price: 442.1 }\n@ C2 =B2.Price*2\n');
  assert.equal(evalIn(src, 'S1', 'C2'), 884.2);
});

test('circular references are unsupported, not hangs', () => {
  const model = modelOf(doc('@ A1 =B1\n@ B1 =A1\n'));
  const { unsupported } = verifyCachedValues(model);
  assert.equal(unsupported.length, 0); // no caches to check — but evaluation must not hang
  assert.throws(() => createEvaluator(model).evaluateCell('S1', 'A1'), /circular/i);
});

test('THE GATE: every cached value in the worked example verifies', () => {
  const model = modelOf(readFileSync(examplePath, 'utf8'), dirname(examplePath));
  const { checked, mismatches, unsupported } = verifyCachedValues(model);
  for (const m of mismatches) console.error(`  MISMATCH ${m.where} =${m.formula}: cached ${JSON.stringify(m.cached)} vs computed ${JSON.stringify(m.computed)}`);
  for (const u of unsupported) console.error(`  unsupported ${u.where} =${u.formula}: ${u.reason}`);
  assert.deepEqual(mismatches, [], 'cached values must be computationally correct');
  assert.deepEqual(unsupported.map((u) => u.where), [], 'the example must stay inside the evaluator subset');
  assert.ok(checked >= 12, `expected a real number of verified caches, got ${checked}`);
});
