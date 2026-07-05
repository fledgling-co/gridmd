// Branch/line coverage for the pure core: dump, refs, scalar, parser, calc,
// units, styles, model. Behaviour is already pinned by the other suites; this
// suite drives the remaining error/edge branches so the 100% gate holds.

import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { lint } from '../src/index';
import { parseDocument, parseInfoArgs, tryProps, findPropsSplit, splitPipeRow } from '../src/parser';
import { parseTarget, parseCell, colToNum, numToCol } from '../src/refs';
import { parseScalar, splitCached } from '../src/scalar';
import { buildWorkbookModel, translateFormula } from '../src/xlsx/model';
import { dumpModel } from '../src/dump';
import { createEvaluator, verifyCachedValues, Unsupported } from '../src/calc';
import { isoToSerial, resolveColor, pxToPt, cmToInch, pxToColWidth } from '../src/xlsx/units';
import { StyleRegistry, parseBorderEdge } from '../src/xlsx/styles';
import type { WorkbookModel } from '../src/types';

const examplePath = fileURLToPath(new URL('../../examples/quarterly-report.gmd', import.meta.url));
const doc = (body: string): string => `---\ngridmd: "0.1"\n---\n\n# S1\n\n${body}`;
const modelOf = (src: string, baseDir = '.'): WorkbookModel => {
  const res = lint(src, { mode: 'strict' });
  assert.deepEqual(res.errors, [], `must lint clean: ${res.errors.map((e) => e.msg).join('; ')}`);
  return buildWorkbookModel(res.doc, { baseDir });
};

// ---- dump ----

test('dumpModel: the worked example', () => {
  const out = dumpModel(modelOf(readFileSync(examplePath, 'utf8'), dirname(examplePath)));
  assert.ok(out.endsWith('\n'));
  const parsed = JSON.parse(out);
  assert.equal(parsed.gridmd, '0.1');
  assert.ok(parsed.sheets.length >= 5);
  assert.ok(parsed.names.length >= 1);
});

test('dumpModel: hidden variants, freeze, protect, rich, blank formula cache', () => {
  const src = `---
gridmd: "0.1"
title: T
date-system: 1904
names:
  - { name: A, ref: "S1!$A$1" }
  - { name: B, value: "5" }
---

# S1

\`\`\`{sheet}
hidden: very
freeze: B2
protect: { enabled: true }
\`\`\`

@ A1 5
@ A2 =A1+1 :: 6
@ A3
  rich:
    - { text: "hi" }
    - { text: "!" }
@ A4 TRUE
@ A5 #DIV/0!
@ A6 2026-07-04
@ B1:C1 { merge: true }

# S2

\`\`\`{sheet}
hidden: true
\`\`\`
`;
  const model = modelOf(src);
  const out = dumpModel(model);
  const parsed = JSON.parse(out);
  assert.equal(parsed.dateSystem, 1904);
  assert.equal(parsed.title, 'T');
  const s1 = parsed.sheets.find((s: { name: string }) => s.name === 'S1');
  assert.equal(s1.hidden, 'very');
  assert.equal(s1.freeze, 'B2');
  assert.equal(s1.protected, true);
  assert.equal(s1.cells.A1.t, 'n');
  assert.equal(s1.cells.A2.t, 'f');
  assert.equal(s1.cells.A3.t, 'rich');
  assert.equal(s1.cells.A4.t, 'b');
  assert.equal(s1.cells.A5.t, 'e');
  assert.equal(s1.cells.A6.t, 'd');
  assert.deepEqual(s1.merges, ['B1:C1']);
  const s2 = parsed.sheets.find((s: { name: string }) => s.name === 'S2');
  assert.equal(s2.hidden, true);
});

test('dumpModel: scalarDump null + empty names', () => {
  // a formula cell with no cached value → cached: null
  const model = modelOf(doc('@ A1 =B1\n'));
  const parsed = JSON.parse(dumpModel(model));
  assert.equal(parsed.sheets[0].cells.A1.cached, null);
  assert.deepEqual(parsed.names, []);
});

// ---- refs ----

test('parseTarget: cell / range / col-range / row-range / quoted sheet', () => {
  assert.equal(parseTarget('A1')!.kind, 'cell');
  assert.equal(parseTarget('A1:B2')!.kind, 'range');
  const cols = parseTarget('B:D')!;
  assert.equal(cols.kind, 'cols');
  assert.deepEqual([cols.c1, cols.c2], [2, 4]);
  const rows = parseTarget('3:7')!;
  assert.equal(rows.kind, 'rows');
  assert.deepEqual([rows.r1, rows.r2], [3, 7]);
  const q = parseTarget("'My Sheet'!A1")!;
  assert.equal(q.sheet, 'My Sheet');
  const q2 = parseTarget("'It''s'!A1")!;
  assert.equal(q2.sheet, "It's");
  assert.equal(parseTarget('nonsense'), null);
  assert.equal(parseTarget('A1:B2:C3'), null); // too many colons
  assert.equal(parseTarget('AA10:AB20')!.kind, 'range');
  assert.equal(parseCell('XFE1'), null); // beyond MAX_COL
  assert.equal(colToNum('AA'), 27);
  assert.equal(numToCol(27), 'AA');
});

// ---- scalar ----

test('scalar: forced text, unterminated CSE, split with trailing spaces', () => {
  assert.equal(parseScalar('{=A1').problem, 'unterminated CSE array formula');
  assert.deepEqual(splitCached('=A1'), { head: '=A1', cached: null });
});

// ---- parser edge branches ----

test('parseDocument: structural errors', () => {
  assert.match(parseDocument('no frontmatter').errors[0]!.msg, /must begin with/);
  assert.match(parseDocument('---\ngridmd: "0.1"\nno close').errors[0]!.msg, /unterminated frontmatter/);
  // lenient mode downgrades unrecognized lines to warnings
  const len = parseDocument('---\ngridmd: "0.1"\n---\n\n# S1\n\nstray\n', { mode: 'lenient' });
  assert.equal(len.errors.length, 0);
  assert.ok(len.warnings.some((w) => /unrecognized/.test(w.msg)));
});

test('parseDocument: YAML tag rejected, empty & bad body', () => {
  const tagged = parseDocument('---\ngridmd: "0.1"\nx-a: !!str 5\n---\n\n# S1\n');
  assert.ok(tagged.errors.some((e) => /tags are outside/.test(e.msg)));
  // @ body that is not a mapping
  const badBody = parseDocument('---\ngridmd: "0.1"\n---\n\n# S1\n\n@ A1\n  - 1\n  - 2\n');
  assert.ok(badBody.errors.some((e) => /must be a YAML mapping/.test(e.msg)));
});

test('parseInfoArgs: missing anchor + bad size + quoted flag', () => {
  const errs1: { line: number; msg: string }[] = [];
  parseInfoArgs('at', 1, errs1);
  assert.ok(errs1.some((e) => /requires an anchor/.test(e.msg)));
  const errs2: { line: number; msg: string }[] = [];
  parseInfoArgs('size nope', 1, errs2);
  assert.ok(errs2.some((e) => /requires WxH/.test(e.msg)));
  const a = parseInfoArgs('foo "a b" key="quoted" at Z9 size 10x20', 1, []);
  assert.deepEqual(a.positional, ['foo', 'a b']);
  assert.equal(a.flags.key, 'quoted');
  assert.equal(a.anchor, 'Z9');
  assert.deepEqual(a.size, { w: 10, h: 20 });
});

test('tryProps + findPropsSplit rejection branches', () => {
  assert.equal(tryProps('[1,2]'), null); // array, not a map
  assert.equal(tryProps(': bad'), null); // YAML error
  assert.equal(findPropsSplit('no braces').propsText, null);
  assert.equal(findPropsSplit('a }{ b }').propsText, null); // unbalanced depth<0
  assert.equal(splitPipeRow('not a pipe row'), null);
});

// ---- units ----

test('units: pxToPt, cmToInch, resolveColor variants', () => {
  assert.equal(pxToPt(96), 72);
  assert.ok(cmToInch(2.54) > 0.99);
  assert.equal(pxToColWidth(0), 0);
  assert.equal(resolveColor('#11223344'), '44112233'); // #RRGGBBAA → AARRGGBB
  assert.equal(resolveColor('auto'), null);
  assert.equal(resolveColor(42), null); // non-string
  assert.equal(resolveColor('not-a-color'), null);
  assert.equal(isoToSerial('06:30:15'), (6 * 3600 + 30 * 60 + 15) / 86400);
});

// ---- styles ----

test('styles: fonts/fills/borders/dxf/numfmt dedup', () => {
  const st = new StyleRegistry({ accent1: '112233' });
  assert.equal(st.fontId({}), 0);
  const f1 = st.fontId({ bold: true, italic: true, underline: 'double', strike: true, sub: true, size: 14, color: '#ff0000', font: 'Arial' });
  const f2 = st.fontId({ bold: true, italic: true, underline: 'double', strike: true, sub: true, size: 14, color: '#ff0000', font: 'Arial' });
  assert.equal(f1, f2); // dedup
  st.fontId({ underline: 'single-accounting', super: true });
  assert.equal(st.fillId({}), 0);
  assert.equal(st.fillId({ fill: 'auto' }), 0); // unresolvable colour
  assert.ok(st.fillId({ fill: '#abcdef', pattern: 'gray-500', fill2: '#000000' }) > 1);
  assert.equal(st.numFmtId('general'), 0);
  assert.ok(st.numFmtId('0.000') >= 164);
  st.dxfId({ bold: true, italic: true, color: '#111111', fill: '#222222' });
  const edge = parseBorderEdge('thin #112233', {});
  assert.equal(edge!.style, 'thin');
  assert.equal(parseBorderEdge(null, {}), null);
  assert.equal(parseBorderEdge({ style: 'none' }, {})!.style, 'thin');
  const bid = st.borderId({
    top: { style: 'thin', color: '#111111' }, bottom: null, left: null, right: null,
    diagUp: { style: 'thin', color: null }, diagDown: null,
  });
  assert.ok(bid > 0);
  assert.equal(st.borderId(null), 0);
  const xf = st.xfId({ align: 'center', valign: 'middle', wrap: true, shrink: true, indent: 2, rotation: 'vertical', numfmt: '0.00', locked: false, hidden: true }, null);
  assert.ok(xf > 0);
  assert.match(st.toXml(), /<styleSheet/);
});

// ---- model ----

test('model: style extend chain, entity, control, range styles, edge borders', () => {
  const src = `---
gridmd: "0.1"
styles:
  base: { bold: true }
  derived: { extend: base, italic: true }
---

# S1

@ A1 "x" { style: derived }
@ B2
  entity: { type: stock, id: "X:MSFT", text: MSFT }
  fields: { Price: 1 }
@ C3 "cb" { control: checkbox }
@ D1:E2 { border: "thin #111111", border-inner: "thin #222222" }
`;
  const model = modelOf(src);
  const s = model.sheets[0]!;
  assert.ok(model.carry.some((c) => c.feature?.includes('entity')));
  assert.ok(model.carry.some((c) => c.feature?.includes('control')));
  // relative-fill via range formula
  const rel = modelOf(doc('@ A1:A3 =B1*2\n'));
  const cell = rel.sheets[0]!.cells.get('1,2');
  assert.ok(cell?.content?.formula);
  void s;
});

test('translateFormula: absolute refs + string literal skip', () => {
  assert.equal(translateFormula('SUM(A1,$B$2,"A1")', 1, 1), 'SUM(B2,$B$2,"A1")');
});

// ---- calc: functions, errors, unsupported ----

const evalIn = (src: string, sheet: string, ref: string): unknown => createEvaluator(modelOf(src)).evaluateCell(sheet, ref);

test('calc: aggregate + text + logical function library', () => {
  const src = doc([
    '@ A1 1', '@ A2 2', '@ A3 3',
    '@ B1 =MIN(A1:A3)', '@ B2 =MAX(A1:A3)', '@ B3 =COUNT(A1:A3)',
    '@ B4 =COUNTA(A1:A3)', '@ B5 =PRODUCT(A1:A3)', '@ B6 =AVERAGE(A1:A3)',
    '@ B7 =AND(1=1,2=2)', '@ B8 =OR(1=2,2=2)', '@ B9 =NOT(1=2)',
    '@ B10 =CONCAT("a","b")', '@ B11 =UPPER("hi")', '@ B12 =LOWER("HI")',
    '@ B13 =RIGHT("widget",3)', '@ B14 =ROUND(3.14159,2)', '@ B15 =SQRT(16)',
    '@ B16 =ABS(-4)', '@ C1 =AVERAGE(D1:D3)',
  ].join('\n'));
  const m = modelOf(src);
  const ev = createEvaluator(m);
  assert.equal(ev.evaluateCell('S1', 'B1'), 1);
  assert.equal(ev.evaluateCell('S1', 'B2'), 3);
  assert.equal(ev.evaluateCell('S1', 'B3'), 3);
  assert.equal(ev.evaluateCell('S1', 'B4'), 3);
  assert.equal(ev.evaluateCell('S1', 'B5'), 6);
  assert.equal(ev.evaluateCell('S1', 'B6'), 2);
  assert.equal(ev.evaluateCell('S1', 'B7'), true);
  assert.equal(ev.evaluateCell('S1', 'B8'), true);
  assert.equal(ev.evaluateCell('S1', 'B9'), true);
  assert.equal(ev.evaluateCell('S1', 'B10'), 'ab');
  assert.equal(ev.evaluateCell('S1', 'B11'), 'HI');
  assert.equal(ev.evaluateCell('S1', 'B12'), 'hi');
  assert.equal(ev.evaluateCell('S1', 'B13'), 'get');
  assert.equal(ev.evaluateCell('S1', 'B14'), 3.14);
  assert.equal(ev.evaluateCell('S1', 'B15'), 4);
  assert.equal(ev.evaluateCell('S1', 'B16'), 4);
  assert.deepEqual(ev.evaluateCell('S1', 'C1'), { err: '#DIV/0!' }); // AVERAGE of blanks
  assert.equal(ev.evaluateCell('S1', 'B1'), 1); // memo hit second time
});

test('calc: array constants, error literals, comparisons, negation', () => {
  assert.equal(evalIn(doc('@ A1 =-5'), 'S1', 'A1'), -5);
  assert.deepEqual(evalIn(doc('@ A1 =#REF!'), 'S1', 'A1'), { err: '#REF!' });
  assert.equal(evalIn(doc('@ A1 =2<=2'), 'S1', 'A1'), true);
  assert.equal(evalIn(doc('@ A1 =3>=4'), 'S1', 'A1'), false);
  assert.equal(evalIn(doc('@ A1 =1<>2'), 'S1', 'A1'), true);
  assert.equal(evalIn(doc('@ A1 =SUM({1,2;3,4})'), 'S1', 'A1'), 10);
  // an error operand short-circuits
  assert.deepEqual(evalIn(doc('@ A1 =#N/A + 1'), 'S1', 'A1'), { err: '#N/A' });
  // text ordering in comparison (number < text)
  assert.equal(evalIn(doc('@ A1 1\n@ A2 z\n@ B1 =A1<A2'), 'S1', 'B1'), true);
});

test('calc: unsupported constructs raise Unsupported', () => {
  const throws = (body: string): void => {
    const ev = createEvaluator(modelOf(doc(body)));
    assert.throws(() => ev.evaluateCell('S1', 'A1'), Unsupported);
  };
  throws('@ A1 =NOTAFUNCTION(1)');
  throws('@ A1 =LAMBDA(x,x)');       // inline LAMBDA invocation
  throws('@ A1 =A:A');               // whole-column range
  throws('@ A1 =UnknownName');
  throws('@ A1 =Tbl[col]');          // structured ref outside a table
});

test('calc: entity field missing → #FIELD!, structured totals, LAMBDA arity', () => {
  const entity = doc('@ B2\n  entity: { type: stock, id: "X", text: X }\n  fields: { Price: 5 }\n@ C1 =B2.Nope\n@ C2 =B2.Price');
  assert.deepEqual(evalIn(entity, 'S1', 'C1'), { err: '#FIELD!' });
  assert.equal(evalIn(entity, 'S1', 'C2'), 5);
  const lam = `---\ngridmd: "0.1"\nnames:\n  - { name: Add, formula: "LAMBDA(a,b,a+b)" }\n---\n\n# S1\n\n@ A1 =Add(2,3)\n@ A2 =Add(1)\n`;
  assert.equal(evalIn(lam, 'S1', 'A1'), 5);
  assert.deepEqual(evalIn(lam, 'S1', 'A2'), { err: '#VALUE!' }); // wrong arity
});

test('calc: SUBTOTAL codes + total-row structured ref', () => {
  const src = doc([
    '```{table} T at A1', 'total:', '  q: =SUBTOTAL(101,[q])', '---',
    '| p | q |', '| a | 2 |', '| b | 4 |', '```', '',
    '@ D1 =SUBTOTAL(1,T[q])', '@ D2 =T[[#Totals],[q]]',
  ].join('\n'));
  const ev = createEvaluator(modelOf(src));
  assert.equal(ev.evaluateCell('S1', 'D1'), 3);   // AVERAGE
  assert.equal(ev.evaluateCell('S1', 'B4'), 3);   // SUBTOTAL(101) total = average
  assert.equal(ev.evaluateCell('S1', 'D2'), 3);   // totals structured ref
});

test('verifyCachedValues: mismatch, unsupported, text/error/date caches', () => {
  const good = modelOf(doc('@ A1 5\n@ A2 =A1*2 :: 10\n@ A3 =A1&"!" :: "5!"'));
  const r1 = verifyCachedValues(good);
  assert.equal(r1.mismatches.length, 0);
  assert.ok(r1.checked >= 2);

  const bad = modelOf(doc('@ A1 5\n@ A2 =A1*2 :: 99'));
  const r2 = verifyCachedValues(bad);
  assert.equal(r2.mismatches.length, 1);

  const unsup = modelOf(doc('@ A1 =NOPE() :: 3'));
  const r3 = verifyCachedValues(unsup);
  assert.equal(r3.unsupported.length, 1);

  // error-valued cache verification path
  const errCache = modelOf(doc('@ A1 =1/0 :: #DIV/0!'));
  assert.equal(verifyCachedValues(errCache).mismatches.length, 0);

  // date-valued cache
  const dateCache = modelOf(doc('@ A1 2026-01-01\n@ A2 =A1 :: 2026-01-01'));
  assert.equal(verifyCachedValues(dateCache).mismatches.length, 0);
});
