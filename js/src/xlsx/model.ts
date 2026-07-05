// Materializes a parsed GridMD document into a per-sheet workbook model the
// XLSX writer can emit: effective cells, merges, tables, CF, validation,
// notes, hyperlinks, outline, page setup, charts, pivots, sparklines,
// slicers, images, shapes/textboxes, threaded comments, scenarios, filters.
// Features with no documented OOXML form (queries, scripts, in-cell cell
// controls, rich-value entities) are CARRIED inside the package as a
// round-trippable custom part — never silently dropped (SPEC §11).

import { parseTarget, parseCell, refKey, numToCol } from '../refs';
import { parseScalar } from '../scalar';
import { isValidPartPath } from '../validate';
import type {
  AtBlock, Cell, CellContent, FenceBlock, InfoArgs, Meta, ParsedDocument,
  RawPart, ReportEntry, Scalar, Sheet, Target, TableIndexEntry, WorkbookModel,
} from '../types';

const DATE_RE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?)?$/;
const TIME_RE = /^\d{2}:\d{2}(:\d{2})?$/;

// No documented native OOXML form → carried in customXml/gridmdCarry1.xml.
const CARRIED_KINDS = new Set(['query', 'script']);

interface CarryInput {
  kind?: string;
  line: number;
  args?: InfoArgs | null;
  meta?: Meta;
  code?: string | null;
  body?: Meta;
}

export function buildWorkbookModel(doc: ParsedDocument, { baseDir = '.' }: { baseDir?: string } = {}): WorkbookModel {
  const fm = doc.frontmatter ?? {};
  const report: ReportEntry[] = [];
  const rawParts: RawPart[] = [];
  const carry: WorkbookModel['carry'] = [];
  const themeColors: Record<string, string> = {};
  for (const [k, v] of Object.entries(fm.theme?.colors ?? {})) themeColors[k] = String(v).replace('#', '').toUpperCase();

  const styleDefs = fm.styles ?? {};
  const resolveStyle = (name: string, seen = new Set<string>()): Meta => {
    const def = styleDefs[name];
    if (!def || seen.has(name)) return {};
    seen.add(name);
    const base = def.extend ? resolveStyle(def.extend, seen) : {};
    const { extend, ...rest } = def;
    return { ...base, ...rest };
  };
  const expandPatch = (props: Meta): Meta => {
    if (!props) return {};
    const { style, ...rest } = props;
    return style ? { ...resolveStyle(style), ...rest } : rest;
  };

  const carryBlock = (b: CarryInput, feature: string, note?: string): void => {
    carry.push({ kind: b.kind ?? 'prop', line: b.line, feature, args: b.args ?? null, meta: b.meta ?? null, code: b.code ?? null, body: b.body ?? null });
    report.push({ line: b.line, feature, action: 'carried', note });
  };

  const handleRaw = (b: FenceBlock): void => {
    if (b.args.flags.part && isValidPartPath(b.args.flags.part)) {
      rawParts.push({ part: b.args.flags.part, payload: b.payload, encoding: b.args.flags.encoding, line: b.line });
      report.push({ line: b.line, feature: `{raw} ${b.args.flags.part}`, action: 'carried' });
    } else {
      report.push({ line: b.line, feature: '{raw}', action: 'not-emitted', note: 'no part= path' });
    }
  };

  for (const b of doc.workbookBlocks) {
    if (b.type !== 'fence') continue;
    if (b.kind === 'raw') handleRaw(b);
    else if (CARRIED_KINDS.has(b.kind)) {
      carryBlock(b, `{${b.kind}} ${b.args.positional[0] ?? ''}`.trim(),
        b.kind === 'query' ? 'query pipelines have no portable OOXML form (DataMashup is opaque M binary)' : 'scripts are stored outside .xlsx packages');
    }
  }

  const sheets: Sheet[] = [];
  for (const sheet of doc.sheets) {
    const meta = sheet.blocks.find((b): b is FenceBlock => b.type === 'fence' && b.kind === 'sheet')?.meta ?? {};
    const s: Sheet = {
      name: sheet.name, meta, kind: meta.kind === 'chart' ? 'chart' : 'worksheet',
      cells: new Map(), merges: [], tables: [], cf: [], validations: [], notes: [],
      hyperlinks: [], outline: { rows: [], cols: [] }, page: null,
      charts: [], sparklines: [], pivots: [], slicers: [], images: [],
      shapes: [], threads: [], scenarios: [], filters: [],
    };
    sheets.push(s);

    const cellAt = (col: number, row: number): Cell => {
      const k = refKey(col, row);
      let c = s.cells.get(k);
      if (!c) { c = { col, row, content: null, patches: [] }; s.cells.set(k, c); }
      return c;
    };
    const setContent = (col: number, row: number, content: CellContent): void => {
      const c = cellAt(col, row);
      if (!c.content) c.content = content;
      else if (content.cached !== undefined && c.content.formula !== undefined && c.content.cached == null) {
        c.content.cached = content.cached;
      }
    };
    const scalarContent = (sc: Scalar): CellContent => {
      if (sc.kind === 'formula') return { formula: sc.formula, cse: sc.cse, cached: sc.cached ?? null };
      return { scalar: sc };
    };

    for (const b of sheet.blocks) {
      if (b.type === 'at') { applyAt(b); continue; }
      switch (b.kind) {
        case 'sheet': break;
        case 'grid': {
          const a = parseCell(b.args.positional[0]!)!;
          (b.rows ?? []).forEach((row, ri) => row.cells.forEach((text, ci) => {
            const sc = parseScalar(text);
            if (sc.kind !== 'blank') setContent(a.col + ci, a.row + ri, scalarContent(sc));
          }));
          break;
        }
        case 'spill-cache': {
          const a = parseCell(b.args.positional[0]!)!;
          (b.rows ?? []).forEach((row, ri) => row.cells.forEach((text, ci) => {
            const sc = parseScalar(text);
            if (sc.kind === 'blank') return;
            if (ri === 0 && ci === 0) setContent(a.col, a.row, { cached: sc });
            else setContent(a.col + ci, a.row + ri, { scalar: sc, spillCache: true });
          }));
          break;
        }
        case 'table': applyTable(b); break;
        case 'cf': s.cf.push({ sqref: b.args.positional[0]!, rules: b.meta, line: b.line }); break;
        case 'validation': s.validations.push({ sqref: b.args.positional[0]!, meta: b.meta }); break;
        case 'filter': s.filters.push({ sqref: b.args.positional[0]!, meta: b.meta, line: b.line }); break;
        case 'outline':
          s.outline.rows.push(...(b.meta.rows ?? []));
          s.outline.cols.push(...(b.meta.cols ?? []));
          break;
        case 'page': s.page = b.meta; break;
        case 'chart': s.charts.push({ type: b.args.positional[0], title: b.args.positional[1] ?? null, anchor: b.args.anchor, size: b.args.size, meta: b.meta, line: b.line }); break;
        case 'sparklines': s.sparklines.push({ sqref: b.args.positional[0]!, meta: b.meta, line: b.line }); break;
        case 'pivot': s.pivots.push({ name: b.args.positional[0]!, anchor: b.args.anchor, meta: b.meta, line: b.line }); break;
        case 'slicer': s.slicers.push({ anchor: b.args.anchor, size: b.args.size, meta: b.meta, kind: b.meta.kind ?? 'slicer', line: b.line }); break;
        case 'image': {
          const src = String(b.meta.src ?? '');
          s.images.push({ anchor: b.args.anchor, size: b.args.size, src, alt: b.meta.alt ?? '', line: b.line });
          break;
        }
        case 'shape':
          s.shapes.push({ preset: b.args.positional[0] ?? 'rect', anchor: b.args.anchor, size: b.args.size, meta: b.meta, line: b.line });
          break;
        case 'textbox':
          s.shapes.push({ preset: 'textbox', anchor: b.args.anchor, size: b.args.size, meta: b.meta, line: b.line });
          break;
        case 'checkbox':
          carryBlock(b, `{checkbox} at ${b.args.anchor}`, 'form controls pending; state carried');
          break;
        case 'comments': s.threads.push({ ref: b.args.positional[0]!, comments: b.meta, line: b.line }); break;
        case 'scenario': s.scenarios.push({ name: b.args.positional[0]!, meta: b.meta, line: b.line }); break;
        case 'raw': handleRaw(b); break;
        default:
          if (CARRIED_KINDS.has(b.kind)) carryBlock(b, `{${b.kind}} ${b.args.positional[0] ?? ''}`.trim());
      }
    }

    function applyAt(b: AtBlock): void {
      const t = parseTarget(b.targetText);
      if (!t) return;
      const body = b.body ?? {};
      const flow = b.props ?? {};

      if (b.scalarText !== null) {
        const sc = parseScalar(b.scalarText);
        if (t.kind === 'cell' && sc.kind !== 'blank') {
          const content = scalarContent(sc);
          if (content.formula !== undefined) {
            const spill = flow.spill ?? body.spill;
            const arr = flow.array ?? body.array;
            if (spill || arr) content.arrayRef = String(spill ?? arr);
            if (arr) content.cse = true;
          }
          setContent(t.c1, t.r1, content);
        } else if (t.kind === 'range' && sc.kind === 'formula') {
          for (let r = t.r1; r <= t.r2; r++) for (let c = t.c1; c <= t.c2; c++) {
            setContent(c, r, { formula: translateFormula(sc.formula!, r - t.r1, c - t.c1), cse: false, cached: null });
          }
        }
      } else {
        const content = bodyContent(body, flow);
        if (content && t.kind === 'cell') setContent(t.c1, t.r1, content);
      }

      const patch = expandPatch({ ...flow, ...bodyProps(body) });
      delete patch.spill; delete patch.array; delete patch.value; delete patch.formula; delete patch.rich; delete patch.entity; delete patch.fields;
      const { merge, link, tip, note, control, ...styleProps } = patch;
      if (merge === true && t.kind === 'range') s.merges.push(t);
      if (link) s.hyperlinks.push({ col: t.c1, row: t.r1, target: link, tip });
      if (note ?? body.note) s.notes.push({ col: t.c1, row: t.r1, text: String(note ?? body.note) });
      if (control) carryBlock({ line: b.line, args: null, meta: { target: b.targetText, control } }, `control: ${control} at ${b.targetText}`, 'in-cell cell controls have no documented OOXML form; boolean value emitted');
      if (body.entity) {
        carryBlock({ line: b.line, args: null, meta: { target: b.targetText, entity: body.entity, fields: body.fields ?? null } }, `entity cell at ${b.targetText}`, 'rich-value parts are undocumented; display text emitted');
      }
      if (Object.keys(styleProps).length) {
        if (t.kind === 'cell') cellAt(t.c1, t.r1).patches.push(styleProps);
        else if (t.kind === 'range') {
          for (let r = t.r1; r <= t.r2; r++) for (let c = t.c1; c <= t.c2; c++) {
            cellAt(c, r).patches.push(rangeEdgeProps(styleProps, t, c, r));
          }
        }
      }
    }

    function bodyContent(body: Meta, flow: Meta): CellContent | null {
      if (body.formula !== undefined) {
        const f = String(body.formula).replace(/^=/, '');
        const content: CellContent = { formula: f, cse: false, cached: body.value !== undefined ? yamlScalar(body.value) : null };
        const spill = body.spill ?? flow.spill;
        const arr = body.array ?? flow.array;
        if (spill || arr) content.arrayRef = String(spill ?? arr);
        if (arr) content.cse = true;
        return content;
      }
      if (body.rich !== undefined) return { rich: body.rich };
      if (body.entity !== undefined) {
        return {
          scalar: { kind: 'text', value: String(body.entity.text ?? body.entity.id ?? '') },
          entityFields: body.fields ?? {},
        };
      }
      if (body.value !== undefined) return { scalar: yamlScalar(body.value) };
      return null;
    }

    function applyTable(b: FenceBlock): void {
      const a = parseCell(b.args.anchor!)!;
      const tm = b.meta ?? {};
      const header = tm.header !== false;
      const columns: string[] = [];
      const rows = b.rows ?? [];
      rows.forEach((row, ri) => row.cells.forEach((text, ci) => {
        const sc = parseScalar(text);
        if (header && ri === 0 && sc.kind === 'text') columns.push(String(sc.value));
        if (sc.kind !== 'blank') setContent(a.col + ci, a.row + ri, scalarContent(sc));
      }));
      const totalRow = tm.total ? a.row + rows.length : null;
      if (tm.total) {
        for (const [colName, v] of Object.entries(tm.total)) {
          const ci = columns.findIndex((c) => c.toLowerCase() === String(colName).toLowerCase());
          if (ci === -1) continue;
          const sc = parseScalar(String(v));
          setContent(a.col + ci, totalRow!, scalarContent(sc));
        }
      }
      for (const [colName, cprops] of Object.entries(tm.cols ?? {})) {
        const ci = columns.findIndex((c) => c.toLowerCase() === String(colName).toLowerCase());
        if (ci === -1 || !cprops || typeof cprops !== 'object') continue;
        const patch = expandPatch(cprops);
        const lastRow = totalRow ?? (a.row + rows.length - 1);
        for (let r = a.row + (header ? 1 : 0); r <= lastRow; r++) cellAt(a.col + ci, r).patches.push(patch);
      }
      const sortLevels = (tm.sort ?? []).filter((lvl: Meta) => (lvl?.by ?? 'value') === 'value');
      if ((tm.sort ?? []).length > sortLevels.length) {
        report.push({ line: b.line, feature: `table ${b.args.positional[0]} color-sort level`, action: 'partial', note: 'sort-by-color levels are not serialized; value levels emitted' });
      }
      s.tables.push({
        name: b.args.positional[0]!, anchor: a, columns,
        headerRow: header, bodyRows: rows.length - (header ? 1 : 0),
        total: tm.total ?? null, style: tm.style, banded: tm.banded ?? 'rows',
        filter: tm.filter ?? null, sort: sortLevels, line: b.line,
      });
    }
  }

  // Global table index for chart/pivot/slicer data-reference resolution.
  const tableIndex = new Map<string, TableIndexEntry>();
  for (const s of sheets) {
    for (const t of s.tables) tableIndex.set(t.name.toLowerCase(), { ...t, sheetName: s.name });
  }

  return { fm, themeColors, sheets, report, rawParts, carry, tableIndex, baseDir };
}

function bodyProps(body: Meta): Meta {
  const { value, formula, rich, entity, fields, spill, array, ...rest } = body ?? {};
  return rest;
}

function yamlScalar(v: unknown): Scalar {
  if (typeof v === 'number') return { kind: 'number', value: v };
  if (typeof v === 'boolean') return { kind: 'boolean', value: v };
  const str = String(v);
  if (DATE_RE.test(str)) return { kind: 'date', value: str };
  if (TIME_RE.test(str)) return { kind: 'time', value: str };
  return { kind: 'text', value: str };
}

function rangeEdgeProps(props: Meta, t: Target, col: number, row: number): Meta {
  const out = { ...props };
  if (out.border !== undefined) {
    if (row === t.r1) out['border-top'] = out.border;
    if (row === t.r2) out['border-bottom'] = out.border;
    if (col === t.c1) out['border-left'] = out.border;
    if (col === t.c2) out['border-right'] = out.border;
    delete out.border;
  }
  if (out['border-inner'] !== undefined) {
    const inner = out['border-inner'];
    if (row > t.r1) out['border-top'] ??= inner;
    if (col > t.c1) out['border-left'] ??= inner;
    delete out['border-inner'];
  }
  return out;
}

// Relative fill (SPEC §8.5): shift unanchored A1 refs by (dr, dc), skipping
// string literals and quoted sheet names.
export function translateFormula(formula: string, dr: number, dc: number): string {
  const colToNumLocal = (str: string): number => [...str].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  let out = '';
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i]!;
    if (ch === '"' || ch === "'") {
      const q = ch;
      let j = i + 1;
      while (j < formula.length) {
        if (formula[j] === q) {
          if (formula[j + 1] === q) { j += 2; continue; }
          break;
        }
        j++;
      }
      out += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    const rest = formula.slice(i);
    const m = /^(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(])/.exec(rest);
    const prev = out.slice(-1);
    if (m && !/[A-Za-z0-9_.]/.test(prev)) {
      const [whole, cd, colL, rd, rowS] = m;
      const col = cd === '$' ? colL : numToCol(Math.max(1, colToNumLocal(colL!) + dc));
      const row = rd === '$' ? rowS : String(Math.max(1, Number(rowS) + dr));
      out += `${cd}${col}${rd}${row}`;
      i += whole!.length;
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}
