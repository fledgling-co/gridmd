import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { findPropsSplit, tryProps, splitPipeRow, parseInfoArgs, parseDocument } from '../src/parser';
import type { AtBlock, FenceBlock } from '../src/types';

// ---- props split (SPEC §9.1 / Appendix A rule 2) ----

test('props split: formula + cache + props', () => {
  const { scalarText, propsText } = findPropsSplit('=FtoC(100) :: 37.8 { numfmt: "0.0" }');
  assert.equal(scalarText, '=FtoC(100) :: 37.8');
  assert.equal(propsText, '{ numfmt: "0.0" }');
  assert.deepEqual(tryProps(propsText!), { numfmt: '0.0' });
});

test('props split: quoted text + props with URL', () => {
  const { scalarText, propsText } = findPropsSplit('"Source" { link: "https://x.dev/q?a=1,b", tip: "t" }');
  assert.equal(scalarText, '"Source"');
  assert.deepEqual(tryProps(propsText!), { link: 'https://x.dev/q?a=1,b', tip: 't' });
});

test('props split: trailing array constant is NOT props', () => {
  // {1,2} parses to a YAML map with null values → rejected by tryProps
  const { propsText } = findPropsSplit('=A1+ {1,2}');
  assert.ok(propsText === null || tryProps(propsText) === null);
});

test('props split: bare array-constant formula stays whole', () => {
  const { scalarText, propsText } = findPropsSplit('={1,2;3,4}');
  assert.equal(propsText, null);
  assert.equal(scalarText, '={1,2;3,4}');
});

test('props split: braces inside formula strings are ignored', () => {
  const { scalarText, propsText } = findPropsSplit('=IF(A1="{",1,2) { bold: true }');
  assert.equal(scalarText, '=IF(A1="{",1,2)');
  assert.deepEqual(tryProps(propsText!), { bold: true });
});

test('props keys must be identifiers with non-null values', () => {
  assert.equal(tryProps('{ Bad Key: 1 }'), null);
  assert.equal(tryProps('{ bold }'), null);
  assert.deepEqual(tryProps('{ spill: B9:B11 }'), { spill: 'B9:B11' });
});

// ---- pipe rows ----

test('pipe row basics + trimming', () => {
  assert.deepEqual(splitPipeRow('| a | 12 | =A1 |'), ['a', '12', '=A1']);
});

test('pipe row escapes', () => {
  assert.deepEqual(splitPipeRow('| a\\|b | c |'), ['a|b', 'c']);
});

test('pipe row empty cells', () => {
  assert.deepEqual(splitPipeRow('|  | x |  |'), ['', 'x', '']);
});

test('pipe row must close', () => {
  assert.equal(splitPipeRow('| a | b'), null);
});

// ---- info strings ----

test('info args: chart fence line', () => {
  const a = parseInfoArgs('column "Revenue by product" at G2:N20', 1, []);
  assert.deepEqual(a.positional, ['column', 'Revenue by product']);
  assert.equal(a.anchor, 'G2:N20');
});

test('info args: at sheet, size, flags', () => {
  const a = parseInfoArgs('ooxml part="xl/charts/chart1.xml" encoding=base64', 1, []);
  assert.deepEqual(a.positional, ['ooxml']);
  assert.equal(a.flags.part, 'xl/charts/chart1.xml');
  assert.equal(a.flags.encoding, 'base64');
  const b = parseInfoArgs('at M2 size 150x200', 1, []);
  assert.equal(b.anchor, 'M2');
  assert.deepEqual(b.size, { w: 150, h: 200 });
});

// ---- document structure ----

const MINI = `---
gridmd: "0.1"
---

# S1

@ A1 "hello" { bold: true }
@ B2
  formula: =A1&"!"
  value: "hello!"

\`\`\`{grid} D1
| 1 | 2 |
\`\`\`
`;

test('mini document parses', () => {
  const doc = parseDocument(MINI);
  assert.deepEqual(doc.errors, []);
  assert.equal(doc.sheets.length, 1);
  const [at1, at2, grid] = doc.sheets[0]!.blocks as [AtBlock, AtBlock, FenceBlock];
  assert.equal(at1.scalarText, '"hello"');
  assert.deepEqual(at1.props, { bold: true });
  assert.equal(at2.body.formula, '=A1&"!"');
  assert.equal(at2.body.value, 'hello!');
  assert.equal(grid.kind, 'grid');
  assert.deepEqual(grid.rows!.map((r) => r.cells), [['1', '2']]);
});

test('multiline body dedent rule: blank lines inside, terminator outside', () => {
  const doc = parseDocument(`---
gridmd: "0.1"
---

# S1

@ A1
  note: |
    line one

    line three
@ B1 2
`);
  assert.deepEqual(doc.errors, []);
  const [a, b] = doc.sheets[0]!.blocks as [AtBlock, AtBlock];
  assert.equal(a.body.note, 'line one\n\nline three\n');
  assert.equal(b.scalarText, '2');
});

test('level-2 headings are comments; unrecognized lines error in strict', () => {
  const doc = parseDocument(`---
gridmd: "0.1"
---

# S1

## organizer heading

stray prose
`);
  assert.equal(doc.errors.length, 1);
  assert.match(doc.errors[0]!.msg, /unrecognized line/);
});

test('unclosed fence is an error', () => {
  const doc = parseDocument(`---
gridmd: "0.1"
---

# S1

\`\`\`{grid} A1
| 1 |
`);
  assert.ok(doc.errors.some((e) => /unclosed/.test(e.msg)));
});

test('YAML aliases rejected (safe subset)', () => {
  const doc = parseDocument(`---
gridmd: "0.1"
x-a: &a [1]
x-b: *a
---

# S1
`);
  assert.ok(doc.errors.some((e) => /aliases/.test(e.msg)));
});
