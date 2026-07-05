// A1-reference parsing (SPEC.md §8.2, Appendix A).

import type { CellPos, Target } from './types';

export const MAX_COL = 16384; // XFD
export const MAX_ROW = 1048576;

export function colToNum(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n: number): string {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = (n - 1 - r) / 26;
  }
  return s;
}

const CELL_RE = /^(\$?)([A-Z]{1,3})(\$?)([1-9]\d{0,6})$/;
const COLRANGE_RE = /^\$?([A-Z]{1,3}):\$?([A-Z]{1,3})$/;
const ROWRANGE_RE = /^\$?([1-9]\d{0,6}):\$?([1-9]\d{0,6})$/;

export function parseCell(text: string): CellPos | null {
  const m = CELL_RE.exec(text);
  if (!m) return null;
  const col = colToNum(m[2]!);
  const row = Number(m[4]!);
  if (col > MAX_COL || row > MAX_ROW) return null;
  return { col, row };
}

// Parses a target: cell | cell:cell | col:col | row:row, with an optional
// leading Sheet! qualifier ('quoted' names supported). Column and row ranges
// carry sentinel coordinates for the missing axis (a full column/row); those
// fields are never read for `cols`/`rows` kinds.
export function parseTarget(input: string): Target | null {
  let text = input;
  let sheet: string | null = null;
  const bang = text.lastIndexOf('!');
  if (bang !== -1) {
    sheet = text.slice(0, bang);
    if (sheet.startsWith("'") && sheet.endsWith("'") && sheet.length >= 2) {
      sheet = sheet.slice(1, -1).replace(/''/g, "'");
    }
    text = text.slice(bang + 1);
  }
  const cell = parseCell(text);
  if (cell) return { kind: 'cell', sheet, c1: cell.col, r1: cell.row, c2: cell.col, r2: cell.row };
  if (text.includes(':')) {
    const parts = text.split(':');
    if (parts.length === 2) {
      const a = parseCell(parts[0]!);
      const b = parseCell(parts[1]!);
      if (a && b) {
        return {
          kind: 'range', sheet,
          c1: Math.min(a.col, b.col), r1: Math.min(a.row, b.row),
          c2: Math.max(a.col, b.col), r2: Math.max(a.row, b.row),
        };
      }
      let m = COLRANGE_RE.exec(text);
      if (m) {
        const c1 = colToNum(m[1]!), c2 = colToNum(m[2]!);
        if (c1 <= MAX_COL && c2 <= MAX_COL) {
          return { kind: 'cols', sheet, c1: Math.min(c1, c2), r1: 1, c2: Math.max(c1, c2), r2: MAX_ROW };
        }
      }
      m = ROWRANGE_RE.exec(text);
      if (m) {
        const r1 = Number(m[1]!), r2 = Number(m[2]!);
        if (r1 <= MAX_ROW && r2 <= MAX_ROW) {
          return { kind: 'rows', sheet, c1: 1, r1: Math.min(r1, r2), c2: MAX_COL, r2: Math.max(r1, r2) };
        }
      }
    }
  }
  return null;
}

export const refKey = (col: number, row: number): string => `${col},${row}`;
