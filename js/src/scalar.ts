// Cell scalar micro-grammar (SPEC.md §6).

import type { Scalar } from './types';

export const ERROR_VALUES = new Set([
  '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A',
  '#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!',
]);

const NUMBER_RE = /^-?(0|[1-9]\d*)(\.\d+)?([eE][+-]?\d+)?$/;
const DATE_RE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?)?$/;
const TIME_RE = /^\d{2}:\d{2}(:\d{2})?$/;
const DQ_RE = /^"((?:[^"]|"")*)"$/;

// Splits "formula :: cached" at the LAST " :: " outside double-quoted
// string literals (SPEC §6). Returns { head, cached } (cached null if none).
export function splitCached(text: string): { head: string; cached: string | null } {
  let inQ = false;
  let idx = -1;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') { inQ = !inQ; continue; }
    if (!inQ && ch === ' ' && text.startsWith(' :: ', i)) idx = i;
  }
  if (idx === -1) return { head: text, cached: null };
  return { head: text.slice(0, idx), cached: text.slice(idx + 4).trim() };
}

// Parses one cell scalar. `raw` must already be trimmed (pipe-row cells are
// trimmed by the row splitter; @-directive scalars by the line parser).
export function parseScalar(raw: string): Scalar {
  if (raw === '') return { kind: 'blank' };

  if (raw.startsWith('{=')) {
    const { head, cached } = splitCached(raw);
    if (!head.endsWith('}')) {
      return { kind: 'text', value: raw, problem: 'unterminated CSE array formula' };
    }
    return { kind: 'formula', cse: true, formula: head.slice(2, -1), cached: parseCached(cached) };
  }
  if (raw.startsWith('=')) {
    const { head, cached } = splitCached(raw);
    return { kind: 'formula', cse: false, formula: head.slice(1), cached: parseCached(cached) };
  }
  if (raw.startsWith("'")) return { kind: 'text', value: raw.slice(1), forced: true };
  if (raw.startsWith('"')) {
    const m = DQ_RE.exec(raw);
    if (m) return { kind: 'text', value: m[1]!.replace(/""/g, '"'), quoted: true };
    return { kind: 'text', value: raw, problem: 'unterminated quoted text' };
  }
  if (NUMBER_RE.test(raw)) return { kind: 'number', value: Number(raw) };
  const up = raw.toUpperCase();
  if (up === 'TRUE' || up === 'FALSE') return { kind: 'boolean', value: up === 'TRUE' };
  if (DATE_RE.test(raw)) return { kind: 'date', value: raw };
  if (TIME_RE.test(raw)) return { kind: 'time', value: raw };
  if (ERROR_VALUES.has(up)) return { kind: 'error', value: up };
  return { kind: 'text', value: raw };
}

function parseCached(cachedText: string | null): Scalar | null {
  if (cachedText === null) return null;
  const v = parseScalar(cachedText);
  if (v.kind === 'formula') return { kind: 'invalid', problem: 'cached value must not be a formula' };
  return v;
}
