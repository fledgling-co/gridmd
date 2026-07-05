// GridMD public API surface.

import { parseDocument } from './parser';
import { validateDocument } from './validate';
import type { Diagnostic, LintResult } from './types';

export { parseDocument } from './parser';
export { validateDocument, isValidPartPath } from './validate';
export { parseScalar, splitCached, ERROR_VALUES } from './scalar';
export { parseTarget, parseCell, colToNum, numToCol, MAX_COL, MAX_ROW } from './refs';
export { findPropsSplit, tryProps, splitPipeRow, parseInfoArgs, parseYaml, RESERVED_KINDS } from './parser';
export { buildWorkbookModel, translateFormula } from './xlsx/model';
export { writeXlsx } from './xlsx/write';
export { xlsxToGridmd } from './xlsx/read';
export { dumpModel } from './dump';
export { verifyCachedValues, createEvaluator, Unsupported } from './calc';
export { zipRead, zipWrite, crc32 } from './xlsx/zip';

export type * from './types';

export function lint(source: string, opts: { mode?: string } = {}): LintResult {
  const doc = parseDocument(source, opts);
  validateDocument(doc);
  const byLine = (a: Diagnostic, b: Diagnostic): number => a.line - b.line;
  return {
    doc,
    errors: [...doc.errors].sort(byLine),
    warnings: [...doc.warnings].sort(byLine),
    sheets: doc.sheets.length,
    cells: doc.stats?.defs ?? 0,
    blocks: doc.stats?.blocks ?? 0,
  };
}
