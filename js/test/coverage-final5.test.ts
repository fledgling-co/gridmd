// Function coverage for the pivot pageFields (filters) reversal arrows and
// every evaluator function.

import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { lint } from '../src/index';
import { buildWorkbookModel } from '../src/xlsx/model';
import { writeXlsx } from '../src/xlsx/write';
import { xlsxToGridmd } from '../src/xlsx/read';
import { createEvaluator } from '../src/calc';
import type { WorkbookModel } from '../src/types';

const doc = (body: string): string => `---\ngridmd: "0.1"\n---\n\n# S1\n\n${body}`;
const modelOf = (src: string): WorkbookModel => {
  const res = lint(src, { mode: 'strict' });
  assert.deepEqual(res.errors, [], `must lint clean: ${res.errors.map((e) => e.msg).join('; ')}`);
  return buildWorkbookModel(res.doc, { baseDir: '.' });
};

test('round-trip: pivot with rows, cols, filters (pageFields) and values', () => {
  const src = `---\ngridmd: "0.1"\n---\n\n# Data\n\n\`\`\`{table} Sales at A1\n---\n| region | quarter | rep | amount |\n| AU | Q1 | Jo | 10 |\n| NZ | Q2 | Al | 20 |\n\`\`\`\n\n\`\`\`{pivot} P at F1\nsource: Sales\nrows:\n  - { field: region }\ncols:\n  - { field: quarter }\nfilters:\n  - { field: rep }\nvalues:\n  - { field: amount, agg: sum }\n\`\`\`\n`;
  const { buffer } = writeXlsx(modelOf(src));
  const { gmd } = xlsxToGridmd(buffer);
  assert.deepEqual(lint(gmd, { mode: 'strict' }).errors, []);
  assert.match(gmd, /filters:/); // pageFields reversed
});

test('calc: every function in the library is invoked', () => {
  const src = doc([
    '@ A1 1', '@ A2 2', '@ A3 3', '@ B1 x', '@ B2 y',
    '@ R1 =SUM(A1:A3)', '@ R2 =AVERAGE(A1:A3)', '@ R3 =MIN(A1:A3)', '@ R4 =MAX(A1:A3)',
    '@ R5 =COUNT(A1:A3)', '@ R6 =COUNTA(B1:B2)', '@ R7 =PRODUCT(A1:A3)', '@ R8 =ABS(-2)',
    '@ R9 =ROUND(1.234,1)', '@ R10 =SQRT(9)', '@ R11 =IF(1=1,"y","n")', '@ R12 =AND(1=1,2=2)',
    '@ R13 =OR(1=2,2=2)', '@ R14 =NOT(1=2)', '@ R15 =CONCAT("a","b")', '@ R16 =CONCATENATE("c","d")',
    '@ R17 =LEN("abc")', '@ R18 =LEFT("hello",2)', '@ R19 =RIGHT("hello",2)', '@ R20 =UPPER("hi")',
    '@ R21 =LOWER("HI")', '@ R22 =SUBTOTAL(9,A1:A3)', '@ R23 =SORT(A1:A3)', '@ R24 =UNIQUE(A1:A3)',
  ].join('\n'));
  const ev = createEvaluator(modelOf(src));
  const checks: [string, unknown][] = [
    ['R1', 6], ['R2', 2], ['R3', 1], ['R4', 3], ['R5', 3], ['R6', 2], ['R7', 6], ['R8', 2],
    ['R9', 1.2], ['R10', 3], ['R11', 'y'], ['R12', true], ['R13', true], ['R14', true],
    ['R15', 'ab'], ['R16', 'cd'], ['R17', 3], ['R18', 'he'], ['R19', 'lo'], ['R20', 'HI'],
    ['R21', 'hi'], ['R22', 6],
  ];
  for (const [ref, want] of checks) assert.equal(ev.evaluateCell('S1', ref), want, ref);
  // SORT / UNIQUE spill to the top-left value
  assert.equal(ev.evaluateCell('S1', 'R23'), 1);
  assert.equal(ev.evaluateCell('S1', 'R24'), 1);
});

test('calc: a call to a wholly unknown function throws', () => {
  assert.throws(
    () => createEvaluator(modelOf(doc('@ A1 =TOTALLYUNKNOWNFN(1,2)'))).evaluateCell('S1', 'A1'),
    /function TOTALLYUNKNOWNFN/,
  );
});

test('calc: evaluateFormula shotgun over operators and edge cases', () => {
  const ev = createEvaluator(modelOf(doc('@ A1 2\n@ A2 3\n@ B1 x\n@ B2 y')));
  const f = (s: string): unknown => ev.evaluateFormula(s, { sheet: 'S1' });
  assert.equal(f('1+2-3*4/2^2'), 0);
  assert.equal(f('-2^2'), 4);
  assert.equal(f('+5'), 5);
  assert.equal(f('50%'), 0.5);
  assert.equal(f('"a"&"b"&2'), 'ab2');
  assert.equal(f('2=2'), true);
  assert.equal(f('2<>3'), true);
  assert.equal(f('"a"="A"'), true);
  assert.equal(f('IF(1=2,10)'), false);   // IF with a missing else
  assert.equal(f('LEFT("word")'), 'w');   // LEFT default length
  assert.equal(f('RIGHT("word",0)'), ''); // RIGHT 0
  assert.equal(f('MIN(A1:A2)'), 2);
  assert.deepEqual(f('SORT(A1:A2)'), [[2], [3]]); // evaluateFormula returns the spill array
  assert.deepEqual(f('AVERAGE(B1:B2)'), { err: '#DIV/0!' }); // AVERAGE of text → no numbers
});

test('calc: SORT/UNIQUE over mixed numbers, text and booleans', () => {
  const src = doc([
    '@ A1 3', '@ A2 pear', '@ A3 1', '@ A4 apple', '@ A5 =1=1', '@ A6 pear',
    '@ B1 =SORT(A1:A6) { spill: B1:B6 }', '@ C1 =UNIQUE(A1:A6) { spill: C1:C5 }',
  ].join('\n'));
  const ev = createEvaluator(modelOf(src));
  // SORT ranks numbers < text < logicals; the top-left is the smallest number
  assert.equal(ev.evaluateCell('S1', 'B1'), 1);
  // UNIQUE keeps first-seen; top-left is 3
  assert.equal(ev.evaluateCell('S1', 'C1'), 3);
});
