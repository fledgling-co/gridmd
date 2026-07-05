import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { parseScalar, splitCached } from '../src/scalar';

test('blank cell', () => {
  assert.deepEqual(parseScalar(''), { kind: 'blank' });
});

test('numbers (JSON grammar)', () => {
  assert.equal(parseScalar('45020.5').kind, 'number');
  assert.equal(parseScalar('45020.5').value, 45020.5);
  assert.equal(parseScalar('-0.3').value, -0.3);
  assert.equal(parseScalar('1e3').value, 1000);
  assert.equal(parseScalar('0.30').kind, 'number'); // non-canonical but valid
  assert.equal(parseScalar('01').kind, 'text');     // leading zero → not a JSON number
  assert.equal(parseScalar('1,000').kind, 'text');  // separators → text
});

test('booleans, case-insensitive', () => {
  assert.deepEqual(parseScalar('TRUE'), { kind: 'boolean', value: true });
  assert.deepEqual(parseScalar('false'), { kind: 'boolean', value: false });
});

test('dates and times', () => {
  assert.equal(parseScalar('2026-07-04').kind, 'date');
  assert.equal(parseScalar('2026-07-04T09:12').kind, 'date');
  assert.equal(parseScalar('2026-07-04T09:12:33').kind, 'date');
  assert.equal(parseScalar('12:30').kind, 'time');
  assert.equal(parseScalar('3:2').kind, 'text'); // hh must be two digits
});

test('errors', () => {
  assert.deepEqual(parseScalar('#DIV/0!'), { kind: 'error', value: '#DIV/0!' });
  assert.equal(parseScalar('#SPILL!').kind, 'error');
  assert.equal(parseScalar('#NOPE!').kind, 'text');
});

test('tick text forces verbatim', () => {
  assert.deepEqual(parseScalar("'0042"), { kind: 'text', value: '0042', forced: true });
  assert.deepEqual(parseScalar("'=not a formula"), { kind: 'text', value: '=not a formula', forced: true });
});

test('quoted text with "" doubling', () => {
  assert.equal(parseScalar('"TRUE"').value, 'TRUE');
  assert.equal(parseScalar('"say ""hi"""').value, 'say "hi"');
  assert.equal(parseScalar('"unterminated').problem, 'unterminated quoted text');
});

test('bare text', () => {
  assert.deepEqual(parseScalar('Widget A'), { kind: 'text', value: 'Widget A' });
});

test('formula without cache', () => {
  const s = parseScalar('=SUM(B4:B10)');
  assert.equal(s.kind, 'formula');
  assert.equal(s.cse, false);
  assert.equal(s.formula, 'SUM(B4:B10)');
  assert.equal(s.cached, null);
});

test('formula with cached value', () => {
  const s = parseScalar('=SUM(B4:B10) :: 45020.5');
  assert.equal(s.formula, 'SUM(B4:B10)');
  assert.deepEqual(s.cached, { kind: 'number', value: 45020.5 });
});

test(':: inside a formula string literal does not split', () => {
  const s = parseScalar('=IF(A1="x :: y",1,2) :: 2');
  assert.equal(s.formula, 'IF(A1="x :: y",1,2)');
  assert.deepEqual(s.cached, { kind: 'number', value: 2 });
});

test('split takes the LAST :: outside quotes', () => {
  const { head, cached } = splitCached('=A1&" :: " :: "a :: b"');
  assert.equal(head, '=A1&" :: "');
  assert.equal(cached, '"a :: b"');
});

test('CSE array formula', () => {
  const s = parseScalar('{=TRANSPOSE(A1:B5)}');
  assert.equal(s.kind, 'formula');
  assert.equal(s.cse, true);
  assert.equal(s.formula, 'TRANSPOSE(A1:B5)');
});

test('CSE with cache', () => {
  const s = parseScalar('{=SUM(A1:A3*B1:B3)} :: 42');
  assert.equal(s.cse, true);
  assert.deepEqual(s.cached, { kind: 'number', value: 42 });
});

test('cached side must not be a formula', () => {
  const s = parseScalar('=A1 :: =B1');
  assert.equal(s.cached!.kind, 'invalid');
});

test('cached text', () => {
  const s = parseScalar('=A1&"!" :: "hi!"');
  assert.deepEqual(s.cached, { kind: 'text', value: 'hi!', quoted: true });
});
