// Canonical model dump — the cross-language conformance contract.
// Every GridMD implementation must produce byte-identical JSON for a given
// document (conformance/README.md). Keys are emitted in a fixed order and
// numbers use shortest round-trip form, so dumps diff cleanly.

import { numToCol } from './refs';
import type { Scalar, WorkbookModel } from './types';

interface ScalarDump {
  t: 'n' | 'b' | 'e' | 'd' | 's';
  v: string | number | boolean;
}

const scalarDump = (s: Scalar | null | undefined): ScalarDump | null => {
  if (!s) return null;
  switch (s.kind) {
    case 'number': return { t: 'n', v: s.value as number };
    case 'boolean': return { t: 'b', v: s.value as boolean };
    case 'error': return { t: 'e', v: s.value as string };
    case 'date':
    case 'time': return { t: 'd', v: s.value as string };
    default: return { t: 's', v: String(s.value ?? '') };
  }
};

export function dumpModel(model: WorkbookModel): string {
  const out = {
    gridmd: model.fm.gridmd ?? null,
    title: model.fm.title ?? null,
    dateSystem: model.fm['date-system'] === 1904 ? 1904 : 1900,
    names: (model.fm.names ?? [])
      .map((n: Record<string, unknown>) => ({ name: n.name, ref: n.ref ?? null, formula: n.formula ?? null, value: n.value !== undefined ? String(n.value) : null }))
      .sort((a: { name: string }, b: { name: string }) => a.name.localeCompare(b.name)),
    sheets: model.sheets.map((s) => ({
      name: s.name,
      kind: s.kind,
      hidden: s.meta.hidden === true ? true : s.meta.hidden === 'very' ? 'very' : false,
      freeze: s.meta.freeze ?? null,
      protected: Boolean(s.meta.protect?.enabled),
      cells: Object.fromEntries(
        [...s.cells.values()]
          .filter((c) => c.content)
          .sort((a, b) => a.row - b.row || a.col - b.col)
          .map((c) => {
            const ref = `${numToCol(c.col)}${c.row}`;
            const ct = c.content!;
            if (ct.rich) return [ref, { t: 'rich', v: ct.rich.map((r: { text: string }) => r.text).join('') }];
            if (ct.formula !== undefined && ct.formula !== null) {
              return [ref, {
                t: 'f', f: ct.formula,
                cached: scalarDump(ct.cached),
                array: ct.arrayRef ?? null,
              }];
            }
            return [ref, scalarDump(ct.scalar)];
          }),
      ),
      merges: s.merges
        .map((m) => `${numToCol(m.c1)}${m.r1}:${numToCol(m.c2)}${m.r2}`)
        .sort(),
      tables: s.tables
        .map((t) => ({
          name: t.name,
          anchor: `${numToCol(t.anchor.col)}${t.anchor.row}`,
          columns: t.columns,
          bodyRows: t.bodyRows,
          hasTotals: Boolean(t.total),
        }))
        .sort((a, b) => a.name.localeCompare(b.name)),
      counts: {
        cf: s.cf.reduce((n, b) => n + (Array.isArray(b.rules) ? b.rules.length : 0), 0),
        validations: s.validations.length,
        notes: s.notes.length,
        threads: s.threads.length,
        scenarios: s.scenarios.length,
        sparklines: s.sparklines.length,
        charts: s.charts.length,
        pivots: s.pivots.length,
        slicers: s.slicers.length,
        images: s.images.length,
        shapes: s.shapes.length,
        hyperlinks: s.hyperlinks.length,
      },
    })),
  };
  return `${JSON.stringify(out, null, 1)}\n`;
}
