import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { lint, isValidPartPath } from '../src/index';

const doc = (body: string): string => `---\ngridmd: "0.1"\n---\n\n${body}`;
const errsOf = (src: string): string[] => lint(src).errors.map((e) => e.msg);

test('duplicate cell definition is an error', () => {
  const errs = errsOf(doc(`# S1\n\n@ A1 1\n@ A1 2\n`));
  assert.ok(errs.some((m) => /defined more than once/.test(m)));
});

test('annotation of a defined cell is NOT a duplicate (SPEC §9.4)', () => {
  const res = lint(doc(`# S1\n\n@ A1 1 { bold: true }\n@ A1\n  note: |\n    fine\n`));
  assert.deepEqual(res.errors, []);
});

test('grid + @ overlap detected; blank grid cells define nothing', () => {
  const src = doc(`# S1\n\n\`\`\`{grid} A1\n| 1 |  |\n\`\`\`\n\n@ B1 2\n@ A1 9\n`);
  const errs = errsOf(src);
  assert.equal(errs.filter((m) => /defined more than once/.test(m)).length, 1); // A1 only
});

test('entity body defines the cell', () => {
  const errs = errsOf(doc(`# S1\n\n@ B2\n  entity: { type: stock, id: "X" }\n@ B2 1\n`));
  assert.ok(errs.some((m) => /defined more than once/.test(m)));
});

test('range content requires a formula', () => {
  const errs = errsOf(doc(`# S1\n\n@ A1:A3 5\n`));
  assert.ok(errs.some((m) => /formula content only/.test(m)));
});

test('relative fill defines every cell (overlap caught)', () => {
  const errs = errsOf(doc(`# S1\n\n@ A1:A3 =B1*2\n@ A2 7\n`));
  assert.ok(errs.some((m) => /defined more than once/.test(m)));
});

test('spill-cache requires an owning spill formula', () => {
  const orphan = errsOf(doc(`# S1\n\n\`\`\`{spill-cache} D2\n| 1 |\n\`\`\`\n`));
  assert.ok(orphan.some((m) => /no owning spill/.test(m)));

  const ok = lint(doc(`# S1\n\n@ D2 =SORT(A1:A3) { spill: D2:D4 }\n\`\`\`{spill-cache} D2\n| 1 |\n| 2 |\n| 3 |\n\`\`\`\n`));
  assert.deepEqual(ok.errors, []);

  const oob = errsOf(doc(`# S1\n\n@ D2 =SORT(A1:A3) { spill: D2:D3 }\n\`\`\`{spill-cache} D2\n| 1 |\n| 2 |\n| 3 |\n\`\`\`\n`));
  assert.ok(oob.some((m) => /exceeds the declared spill/.test(m)));
});

test('spill-cache cells are not definitions', () => {
  const res = lint(doc(`# S1\n\n@ D2 =SORT(A1:A2) { spill: D2:D3 }\n\`\`\`{spill-cache} D2\n| 1 |\n| 2 |\n\`\`\`\n@ D3\n  note: annotation ok\n`));
  assert.deepEqual(res.errors, []);
});

test('table: duplicate and non-text headers are errors', () => {
  const dup = errsOf(doc(`# S1\n\n\`\`\`{table} T1 at A1\n---\n| a | A |\n| 1 | 2 |\n\`\`\`\n`));
  assert.ok(dup.some((m) => /duplicate table column name/.test(m)));

  const numeric = errsOf(doc(`# S1\n\n\`\`\`{table} T2 at A1\n---\n| a | 2026-01-01 |\n| 1 | 2 |\n\`\`\`\n`));
  assert.ok(numeric.some((m) => /header cells must be non-empty text/.test(m)));
});

test('table: meta referencing unknown columns', () => {
  const errs = errsOf(doc(`# S1\n\n\`\`\`{table} T3 at A1\ncols: { nope: { numfmt: "0" } }\n---\n| a | b |\n| 1 | 2 |\n\`\`\`\n`));
  assert.ok(errs.some((m) => /unknown column: nope/.test(m)));
});

test('chart sheet rules', () => {
  const good = lint(doc(`# C1\n\n\`\`\`{sheet}\nkind: chart\n\`\`\`\n\n\`\`\`{chart} column "T" at sheet\ndata: A1:A5\n\`\`\`\n`));
  assert.deepEqual(good.errors, []);

  const withGrid = errsOf(doc(`# C1\n\n\`\`\`{sheet}\nkind: chart\n\`\`\`\n\n\`\`\`{chart} column "T" at sheet\ndata: A1:A5\n\`\`\`\n\n\`\`\`{grid} A1\n| 1 |\n\`\`\`\n`));
  assert.ok(withGrid.some((m) => /cannot carry worksheet grid content/.test(m)));

  const noKind = errsOf(doc(`# S1\n\n\`\`\`{chart} column "T" at sheet\ndata: A1:A5\n\`\`\`\n`));
  assert.ok(noKind.some((m) => /require \{sheet\} kind: chart/.test(m)));
});

test('cf: rule shape + explicit priority', () => {
  const errs = errsOf(doc(`# S1\n\n\`\`\`{cf} A1:A9\n- when: "> 5"\n  top: 3\n- formula: =A1>0\n  priority: 0\n\`\`\`\n`));
  assert.ok(errs.some((m) => /exactly one distinguishing key/.test(m)));
  assert.ok(errs.some((m) => /priority must be a positive integer/.test(m)));
});

test('unsafe link and image schemes rejected', () => {
  const errs = errsOf(doc(`# S1\n\n@ A1 "x" { link: "javascript:alert(1)" }\n\n\`\`\`{image} at B2 size 10x10\nsrc: "data:text/html;base64,AAAA"\n\`\`\`\n`));
  assert.ok(errs.some((m) => /link: scheme/.test(m)));
  assert.ok(errs.some((m) => /image src fails/.test(m)));
});

test('anchor sheet qualifier must match the containing sheet', () => {
  const errs = errsOf(doc(`# S1\n\n@ Other!A1 1\n`));
  assert.ok(errs.some((m) => /must name the containing sheet/.test(m)));
});

test('sheet-scoped directives rejected at workbook level', () => {
  const errs = errsOf(doc(`\`\`\`{grid} A1\n| 1 |\n\`\`\`\n\n# S1\n`));
  assert.ok(errs.some((m) => /sheet-scoped/.test(m)));
});

test('raw part path canonicalization', () => {
  assert.equal(isValidPartPath('xl/charts/chart1.xml'), true);
  assert.equal(isValidPartPath('xl/vbaProject.bin'), true);
  assert.equal(isValidPartPath('/absolute.xml'), false);
  assert.equal(isValidPartPath('a/../b.xml'), false);
  assert.equal(isValidPartPath('a//b.xml'), false);
  assert.equal(isValidPartPath('a\\b.xml'), false);
  assert.equal(isValidPartPath('a/%2e%2e/b.xml'), false);
  assert.equal(isValidPartPath('has space.xml'), false);
});

test('duplicate sheet names + bad sheet chars', () => {
  const errs = errsOf(`---\ngridmd: "0.1"\n---\n\n# Data\n\n# data\n\n# Bad[1]\n`);
  assert.ok(errs.some((m) => /duplicate sheet name/.test(m)));
  assert.ok(errs.some((m) => /forbidden character/.test(m)));
});

test('table/pivot names share the defined-name namespace', () => {
  const errs = errsOf(`---\ngridmd: "0.1"\nnames:\n  - { name: Sales, ref: "S1!$A$1" }\n---\n\n# S1\n\n\`\`\`{table} Sales at A1\n---\n| a |\n| 1 |\n\`\`\`\n`);
  assert.ok(errs.some((m) => /collides with an existing name/.test(m)));
});
