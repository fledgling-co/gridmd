// GridMD semantic validation (SPEC.md §9.4, §12–§13; DIRECTIVES.md).

import { parseTarget, parseCell, refKey, MAX_COL, MAX_ROW } from './refs';
import { parseScalar } from './scalar';
import { RESERVED_KINDS } from './parser';
import type {
  AtBlock, FenceBlock, ParsedDocument, ParseStats, SheetBlock, Target, TargetKind,
} from './types';

const SHEET_NAME_BAD = /[:\\/?*[\]]/;
const TABLE_NAME_RE = /^[A-Za-z_\\][A-Za-z0-9_.\\]{0,254}$/;
const CELLISH_NAME_RE = /^[A-Za-z]{1,3}\d+$/;
const COLOR_RE = /^#[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/;
const THEME_COLOR_RE = /^(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)(@-?\d{1,3})?$/;

const WORKBOOK_KINDS = new Set(['query', 'script', 'raw']);
const CONTENT_KEYS = ['value', 'formula', 'rich', 'entity'];
const FILL_ENUMERATION_CAP = 10000;

const KNOWN_PROPS = new Set([
  'style', 'font', 'size', 'bold', 'italic', 'underline', 'strike', 'sub',
  'super', 'color', 'fill', 'pattern', 'fill2', 'border', 'border-top',
  'border-right', 'border-bottom', 'border-left', 'border-diag-up',
  'border-diag-down', 'border-inner', 'border-inner-h', 'border-inner-v',
  'align', 'valign', 'rotation', 'indent', 'wrap', 'shrink', 'numfmt',
  'merge', 'locked', 'hidden', 'link', 'tip', 'note', 'rich', 'spill',
  'array', 'control', 'entity', 'fields', 'value', 'formula',
]);

const SHEET_META_KEYS = new Set([
  'kind', 'tab-color', 'hidden', 'freeze', 'split', 'view',
  'default-row-height', 'default-col-width', 'cols', 'rows', 'protect', 'names',
]);

const FRONTMATTER_KEYS = new Set([
  'gridmd', 'title', 'properties', 'locale', 'date-system', 'calc', 'theme',
  'names', 'styles', 'table-styles', 'links', 'protection',
]);

const CHART_TYPES = new Set([
  'column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'bubble',
  'radar', 'stock', 'surface', 'histogram', 'pareto', 'box-whisker',
  'treemap', 'sunburst', 'waterfall', 'funnel', 'map', 'combo',
]);

const SHAPE_KINDS = new Set([
  'rect', 'rounded-rect', 'ellipse', 'triangle', 'right-triangle', 'diamond',
  'pentagon', 'hexagon', 'star', 'arrow-right', 'arrow-left', 'arrow-up',
  'arrow-down', 'chevron', 'callout', 'line', 'connector',
]);

const VALIDATION_TYPES = new Set(['list', 'whole', 'decimal', 'date', 'time', 'text-length', 'custom']);
const CF_RULE_KEYS = ['when', 'contains', 'not-contains', 'begins', 'ends', 'date', 'dupes', 'unique', 'top', 'bottom', 'avg', 'bars', 'scale', 'icons', 'formula'];

function chartBaseType(t: string): string {
  let base = t;
  for (const suf of ['-stacked100', '-stacked', '-3d']) {
    if (base.endsWith(suf)) base = base.slice(0, -suf.length);
  }
  return base;
}

const isColor = (v: unknown): boolean =>
  typeof v === 'string' && (v === 'auto' || COLOR_RE.test(v) || THEME_COLOR_RE.test(v));

const isSafeLink = (v: unknown): boolean => typeof v === 'string' && /^(https:\/\/|mailto:|#)/.test(v);

function isSafeImageSrc(v: unknown): boolean {
  if (typeof v !== 'string') return false;
  if (/^(javascript|vbscript|file):/i.test(v)) return false;
  if (/^data:/i.test(v)) return /^data:image\//i.test(v);
  if (/^[a-z][a-z0-9+.-]*:/i.test(v)) return /^https:/i.test(v);
  return true; // relative path
}

// {raw} part= path rules (DIRECTIVES §18).
export function isValidPartPath(p: unknown): boolean {
  if (typeof p !== 'string' || p === '') return false;
  if (p.startsWith('/') || p.includes('\\')) return false;
  if (/[\x00-\x1f ]/.test(p)) return false;
  if (/%2e|%2f|%5c/i.test(p)) return false;
  const segs = p.split('/');
  return segs.every((s) => s !== '' && s !== '.' && s !== '..');
}

type TargetFn = (text: string | null | undefined, line: number, kinds: TargetKind[], what: string) => Target | null;
type AddDefFn = (col: number, row: number, line: number, what: string) => void;
interface ValidateCtx { target: TargetFn; addDef: AddDefFn }

export function validateDocument(doc: ParsedDocument): ParsedDocument {
  const errors = doc.errors;
  const warnings = doc.warnings;
  const err = (line: number, msg: string): void => { errors.push({ line, msg }); };
  const warn = (line: number, msg: string): void => { warnings.push({ line, msg }); };
  const stats: ParseStats = { defs: 0, blocks: 0 };
  doc.stats = stats;

  // Excel's single namespace for defined names + table/pivot names.
  const globalNames = new Map<string, string>();

  // ---- frontmatter ----
  const fm = doc.frontmatter ?? {};
  if (typeof fm.gridmd !== 'string' || !/^\d+\.\d+$/.test(fm.gridmd)) {
    err(2, 'frontmatter requires gridmd: "MAJOR.MINOR" (quoted string)');
  }
  for (const k of Object.keys(fm)) {
    if (!FRONTMATTER_KEYS.has(k) && !k.startsWith('x-')) warn(2, `unknown frontmatter key: ${k}`);
  }
  if (fm['date-system'] !== undefined && ![1900, 1904].includes(fm['date-system'])) {
    err(2, 'date-system must be 1900 or 1904');
  }
  if (fm.calc?.mode !== undefined && !['auto', 'auto-no-tables', 'manual'].includes(fm.calc.mode)) {
    err(2, `calc.mode must be auto | auto-no-tables | manual, got ${fm.calc.mode}`);
  }
  for (const n of fm.names ?? []) {
    if (!n || typeof n.name !== 'string') { err(2, 'names entries require a name'); continue; }
    const forms = ['ref', 'formula', 'value'].filter((k) => n[k] !== undefined);
    if (forms.length !== 1) err(2, `name ${n.name}: exactly one of ref | formula | value required`);
    if (globalNames.has(n.name.toLowerCase())) err(2, `duplicate defined name: ${n.name}`);
    globalNames.set(n.name.toLowerCase(), 'name');
  }
  for (const [name, style] of Object.entries(fm.styles ?? {})) {
    if (style === null || typeof style !== 'object' || Array.isArray(style)) {
      err(2, `style ${name} must be a mapping`);
    }
  }
  for (const [slot, v] of Object.entries(fm.theme?.colors ?? {})) {
    if (!/^(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)$/.test(slot)) warn(2, `unknown theme color slot: ${slot}`);
    else if (!COLOR_RE.test(String(v))) err(2, `theme color ${slot} must be #RRGGBB`);
  }

  // ---- fence validation (shared by workbook + sheet scope) ----
  // `ctx` is present only in sheet scope. At workbook level the sole reachable
  // kinds are query/script/raw, none of which touch ctx.target/ctx.addDef, so
  // it is omitted there and asserted (`ctx!`) in the sheet-only cases.
  function validateFence(b: FenceBlock, ctx?: ValidateCtx): void {
    const meta = b.meta ?? {};
    const pos = b.args.positional;
    const need = (cond: unknown, msg: string): void => { if (!cond) err(b.line, `{${b.kind}} ${msg}`); };

    switch (b.kind) {
      case 'grid': {
        const anchor = parseCell(pos[0] ?? '');
        need(anchor, 'requires a cell anchor');
        if (!anchor) break;
        for (const [ri, row] of (b.rows ?? []).entries()) {
          for (const [ci, cellText] of row.cells.entries()) {
            const s = parseScalar(cellText);
            if (s.problem) err(row.line, `grid cell: ${s.problem}`);
            if (s.kind !== 'blank') ctx!.addDef(anchor.col + ci, anchor.row + ri, row.line, '{grid}');
          }
        }
        break;
      }
      case 'table': {
        const name = pos[0];
        need(typeof name === 'string' && TABLE_NAME_RE.test(name) && !CELLISH_NAME_RE.test(name), 'requires a valid table name');
        const anchor = parseCell(b.args.anchor ?? '');
        need(anchor, 'requires `at <cell>`');
        if (typeof name === 'string') {
          const key = name.toLowerCase();
          if (globalNames.has(key)) err(b.line, `table name collides with an existing name: ${name}`);
          globalNames.set(key, 'table');
        }
        if (!anchor || !(b.rows ?? []).length) { need((b.rows ?? []).length, 'requires payload rows'); break; }
        const header = meta.header !== false;
        const columns: string[] = [];
        for (const [ri, row] of (b.rows ?? []).entries()) {
          for (const [ci, cellText] of row.cells.entries()) {
            const s = parseScalar(cellText);
            if (s.problem) err(row.line, `table cell: ${s.problem}`);
            if (header && ri === 0) {
              if (s.kind !== 'text' || s.value === '') err(row.line, `table header cells must be non-empty text (column ${ci + 1})`);
              else columns.push(String(s.value));
              ctx!.addDef(anchor.col + ci, anchor.row + ri, row.line, '{table} header');
              continue;
            }
            if (s.kind !== 'blank') ctx!.addDef(anchor.col + ci, anchor.row + ri, row.line, '{table}');
          }
        }
        const lower = columns.map((c) => c.toLowerCase());
        lower.forEach((c, i) => {
          if (lower.indexOf(c) !== i) err(b.line, `duplicate table column name: ${columns[i]}`);
        });
        const colSet = new Set(lower);
        const checkCols = (obj: Record<string, unknown> | null | undefined, what: string): void => {
          for (const k of Object.keys(obj ?? {})) {
            if (!colSet.has(String(k).toLowerCase())) err(b.line, `${what} references unknown column: ${k}`);
          }
        };
        checkCols(meta.cols, 'cols');
        checkCols(meta.total, 'total');
        checkCols(meta.filter, 'filter');
        for (const s of meta.sort ?? []) {
          if (!colSet.has(String(s?.col ?? '').toLowerCase())) err(b.line, `sort references unknown column: ${s?.col}`);
        }
        if (meta.total && typeof meta.total === 'object') {
          const totalRow = anchor.row + (b.rows ?? []).length;
          for (const colName of Object.keys(meta.total)) {
            const ci = lower.indexOf(String(colName).toLowerCase());
            if (ci !== -1) ctx!.addDef(anchor.col + ci, totalRow, b.line, '{table} total');
          }
        }
        break;
      }
      case 'cf': {
        need(ctx!.target(pos[0], b.line, ['cell', 'range', 'cols', 'rows'], '{cf}'), 'requires a target range');
        const rules = Array.isArray(meta) ? meta : null;
        need(rules, 'body must be a YAML list of rules');
        for (const rule of rules ?? []) {
          const kinds = CF_RULE_KEYS.filter((k) => rule?.[k] !== undefined);
          if (kinds.length !== 1) err(b.line, 'each cf rule needs exactly one distinguishing key');
          if (rule?.priority !== undefined && (!Number.isInteger(rule.priority) || rule.priority < 1)) {
            err(b.line, 'cf priority must be a positive integer');
          }
          for (const key of ['fill', 'color']) {
            if (rule?.format?.[key] !== undefined && !isColor(rule.format[key])) {
              err(b.line, `cf format.${key}: not a color: ${rule.format[key]}`);
            }
          }
        }
        break;
      }
      case 'validation':
        need(ctx!.target(pos[0], b.line, ['cell', 'range', 'cols', 'rows'], '{validation}'), 'requires a target');
        need(VALIDATION_TYPES.has(meta.type), `type must be one of ${[...VALIDATION_TYPES].join(' | ')}`);
        if (meta.type === 'list') need(meta.values !== undefined || meta.source !== undefined, 'list validation requires values: or source:');
        if (meta.error?.style !== undefined) need(['stop', 'warning', 'information'].includes(meta.error.style), 'error.style must be stop | warning | information');
        break;
      case 'filter':
        need(ctx!.target(pos[0], b.line, ['range'], '{filter}'), 'requires a range');
        for (const k of Object.keys(meta.cols ?? {})) {
          if (!/^[A-Z]{1,3}$/.test(k)) err(b.line, `filter cols keys are column letters on plain ranges: ${k}`);
        }
        break;
      case 'chart': {
        const type = pos[0];
        if (type !== undefined && !CHART_TYPES.has(chartBaseType(type))) {
          warn(b.line, `unknown chart type ${type} — a converter must carry it via fallback:`);
        }
        need(b.args.anchor, 'requires `at <anchor>` (or `at sheet` on a chart sheet)');
        if (b.args.anchor && b.args.anchor !== 'sheet') {
          ctx!.target(b.args.anchor, b.line, ['cell', 'range'], '{chart} at');
        }
        need(meta.series !== undefined || meta.data !== undefined || meta.pivot !== undefined, 'requires series:, data:, or pivot:');
        for (const [i, s] of (Array.isArray(meta.series) ? meta.series : []).entries()) {
          if (!s || (s.val === undefined && meta.pivot === undefined)) err(b.line, `series[${i}] requires val:`);
          if (s?.color !== undefined && !isColor(s.color)) err(b.line, `series[${i}].color: not a color`);
        }
        break;
      }
      case 'sparklines':
        need(ctx!.target(pos[0], b.line, ['cell', 'range'], '{sparklines}'), 'requires a target range');
        need(meta.source !== undefined, 'requires source:');
        if (meta.type !== undefined) need(['line', 'column', 'win-loss'].includes(meta.type), 'type must be line | column | win-loss');
        break;
      case 'pivot':
        need(typeof pos[0] === 'string', 'requires a name');
        need(parseCell((b.args.anchor ?? '').replace(/^.*!/, '')), 'requires `at <cell>`');
        need(meta.source !== undefined, 'requires source:');
        if (typeof pos[0] === 'string') {
          const key = pos[0].toLowerCase();
          if (globalNames.has(key)) err(b.line, `pivot name collides with an existing name: ${pos[0]}`);
          globalNames.set(key, 'pivot');
        }
        break;
      case 'slicer':
        need(b.args.anchor, 'requires an anchor');
        need(meta.for !== undefined && meta.field !== undefined, 'requires for: and field:');
        break;
      case 'image':
        need(b.args.anchor, 'requires an anchor');
        need(typeof meta.src === 'string', 'requires src:');
        if (typeof meta.src === 'string' && !isSafeImageSrc(meta.src)) err(b.line, `image src fails the scheme allowlist: ${meta.src}`);
        break;
      case 'shape':
        if (pos[0] !== undefined && !SHAPE_KINDS.has(pos[0])) warn(b.line, `unknown shape kind ${pos[0]} — carry exotic geometry via fallback:`);
        need(b.args.anchor, 'requires an anchor');
        break;
      case 'textbox':
        need(b.args.anchor, 'requires an anchor');
        break;
      case 'checkbox':
        need(b.args.anchor, 'requires an anchor');
        need(meta.linked === undefined || parseCell(String(meta.linked).replace(/\$/g, '')), 'linked: must be a cell');
        break;
      case 'comments': {
        need(ctx!.target(pos[0], b.line, ['cell'], '{comments}'), 'requires a cell target');
        const list = Array.isArray(meta) ? meta : null;
        need(list, 'body must be a YAML list of comments');
        for (const c of list ?? []) {
          if (!c?.by || !c?.at || !c?.text) err(b.line, 'each comment requires by:, at:, text:');
        }
        break;
      }
      case 'outline':
        for (const r of meta.rows ?? []) {
          if (!/^\d+:\d+$/.test(String(r?.range ?? ''))) err(b.line, `outline rows range must be "n:m": ${r?.range}`);
        }
        for (const c of meta.cols ?? []) {
          if (!/^[A-Z]{1,3}:[A-Z]{1,3}$/.test(String(c?.range ?? ''))) err(b.line, `outline cols range must be "A:B": ${c?.range}`);
        }
        break;
      case 'page':
        if (meta.orientation !== undefined) need(['portrait', 'landscape'].includes(meta.orientation), 'orientation must be portrait | landscape');
        need(!(meta.scale !== undefined && meta.fit !== undefined), 'scale: and fit: are mutually exclusive');
        break;
      case 'query':
        need(typeof pos[0] === 'string', 'requires a name');
        need(meta.source !== undefined, 'requires source:');
        need(meta.steps === undefined || Array.isArray(meta.steps), 'steps: must be a list');
        break;
      case 'script':
        need(typeof pos[0] === 'string', 'requires a name');
        need(typeof b.args.flags.lang === 'string', 'requires lang=');
        need((b.code ?? '').trim() !== '', 'requires a code payload after ---');
        break;
      case 'scenario':
        need(typeof pos[0] === 'string', 'requires a name');
        need(meta.cells && typeof meta.cells === 'object', 'requires cells:');
        for (const k of Object.keys(meta.cells ?? {})) {
          if (!parseCell(k.replace(/\$/g, ''))) err(b.line, `scenario cells key must be a cell: ${k}`);
        }
        break;
      case 'raw':
        need(['ooxml', 'json', 'text'].includes(pos[0] ?? ''), 'format must be ooxml | json | text');
        if (b.args.flags.part !== undefined) {
          need(isValidPartPath(b.args.flags.part), `part= fails package-path canonicalization: ${b.args.flags.part}`);
        }
        if (b.args.flags.encoding !== undefined) need(b.args.flags.encoding === 'base64', 'encoding must be base64');
        break;
      default:
        break;
    }
  }

  // ---- workbook-level blocks ----
  for (const b of doc.workbookBlocks) {
    stats.blocks++;
    if (b.type === 'at') { err(b.line, '@ directives are not allowed before the first sheet'); continue; }
    if (b.kind.startsWith('x-')) continue;
    if (!RESERVED_KINDS.has(b.kind)) { err(b.line, `unknown directive {${b.kind}}`); continue; }
    if (!WORKBOOK_KINDS.has(b.kind)) {
      err(b.line, `{${b.kind}} is sheet-scoped and cannot appear before the first sheet`);
      continue;
    }
    validateFence(b);
  }

  // ---- sheets ----
  if (doc.sheets.length === 0) err(1, 'a workbook requires at least one sheet (a level-1 heading)');
  const sheetNames = new Map<string, SheetBlock>();
  for (const sheet of doc.sheets) {
    const nameKey = sheet.name.toLowerCase();
    if (sheet.name.length > 31) err(sheet.line, `sheet name exceeds 31 chars: ${sheet.name}`);
    if (SHEET_NAME_BAD.test(sheet.name)) err(sheet.line, `sheet name contains a forbidden character (: \\ / ? * [ ]): ${sheet.name}`);
    if (sheetNames.has(nameKey)) err(sheet.line, `duplicate sheet name: ${sheet.name}`);
    sheetNames.set(nameKey, sheet);
    validateSheet(sheet);
  }

  function validateSheet(sheet: SheetBlock): void {
    const defs = new Map<string, number>(); // refKey -> line
    const spills: Target[] = [];      // ranges from spill:/array: props
    const spillCaches: FenceBlock[] = [];
    const sheetMetas: FenceBlock[] = [];
    let chartsAtSheet = 0;
    let gridContent = 0;

    const addDef: AddDefFn = (col, row, line, what) => {
      if (col > MAX_COL || row > MAX_ROW) { err(line, `${what}: cell out of bounds`); return; }
      const k = refKey(col, row);
      const prev = defs.get(k);
      if (prev !== undefined) {
        err(line, `${what}: cell defined more than once (previous definition at line ${prev})`);
        return;
      }
      defs.set(k, line);
      stats.defs++;
    };

    const target: TargetFn = (text, line, kinds, what) => {
      const t = parseTarget(text ?? '');
      if (!t || !kinds.includes(t.kind)) { err(line, `${what}: invalid target ${text}`); return null; }
      if (t.sheet && t.sheet.toLowerCase() !== sheet.name.toLowerCase()) {
        err(line, `${what}: anchor qualifier ${t.sheet}! must name the containing sheet (${sheet.name})`);
      }
      return t;
    };

    const ctx: ValidateCtx = { target, addDef };

    for (const b of sheet.blocks) {
      stats.blocks++;
      if (b.type === 'at') { validateAt(b); continue; }
      if (b.kind.startsWith('x-')) continue;
      if (!RESERVED_KINDS.has(b.kind)) { err(b.line, `unknown directive {${b.kind}}`); continue; }
      if (b.kind === 'sheet') { sheetMetas.push(b); validateSheetMeta(b); continue; }
      if (b.kind === 'grid' || b.kind === 'table') gridContent++;
      if (b.kind === 'spill-cache') { spillCaches.push(b); continue; }
      if (b.kind === 'chart' && b.args.anchor === 'sheet') chartsAtSheet++;
      validateFence(b, ctx);
    }

    if (sheetMetas.length > 1) err(sheetMetas[1]!.line, 'multiple {sheet} blocks in one sheet');
    if (sheetMetas.length && sheet.blocks[0] !== sheetMetas[0]) {
      warn(sheetMetas[0]!.line, '{sheet} should be the first block of its sheet');
    }
    const meta = sheetMetas[0]?.meta ?? {};

    // Chart sheets (SPEC §5, DIRECTIVES §5).
    if (meta.kind === 'chart') {
      if (chartsAtSheet !== 1) err(sheet.line, `a chart sheet requires exactly one {chart} anchored \`at sheet\` (found ${chartsAtSheet})`);
      if (gridContent > 0 || defs.size > 0) err(sheet.line, 'a chart sheet cannot carry worksheet grid content');
    } else if (chartsAtSheet > 0) {
      err(sheet.line, '`at sheet` chart anchors require {sheet} kind: chart');
    }

    // {spill-cache} (SPEC §8.3, §13).
    for (const sc of spillCaches) {
      const anchor = parseCell(sc.args.positional[0] ?? '');
      if (!anchor) { err(sc.line, '{spill-cache} requires a cell anchor'); continue; }
      const rows = sc.rows ?? [];
      const h = rows.length;
      const w = Math.max(0, ...rows.map((r) => r.cells.length));
      const owner = spills.find((s) => s.c1 === anchor.col && s.r1 === anchor.row);
      if (!owner) {
        err(sc.line, `{spill-cache} at ${sc.args.positional[0]} has no owning spill/array formula at that anchor`);
        continue;
      }
      if (anchor.row + h - 1 > owner.r2 || anchor.col + w - 1 > owner.c2) {
        err(sc.line, '{spill-cache} rectangle exceeds the declared spill/array range');
      }
    }

    function validateAt(b: AtBlock): void {
      const t = parseTarget(b.targetText);
      if (!t) { err(b.line, `invalid @ target: ${b.targetText}`); return; }
      if (t.sheet && t.sheet.toLowerCase() !== sheet.name.toLowerCase()) {
        err(b.line, `@ target qualifier ${t.sheet}! must name the containing sheet`);
      }
      const body = b.body ?? {};
      const props = { ...(b.props ?? {}), ...body };

      const bodyContentKeys = CONTENT_KEYS.filter((k) => body[k] !== undefined);
      let scalar = null;
      if (b.scalarText !== null) {
        scalar = parseScalar(b.scalarText);
        if (scalar.problem) err(b.line, `scalar: ${scalar.problem}`);
        if (scalar.cached?.kind === 'invalid') err(b.line, `scalar: ${scalar.cached.problem}`);
        const cachedOnly = bodyContentKeys.length === 1 && bodyContentKeys[0] === 'value' && scalar.kind === 'formula';
        if (bodyContentKeys.length && !cachedOnly) {
          err(b.line, 'inline content and body content keys on the same @ directive');
        }
      }
      const hasFormula = scalar?.kind === 'formula' || body.formula !== undefined;
      const hasContent = (scalar !== null && scalar.kind !== 'blank') || bodyContentKeys.length > 0;

      if (hasContent) {
        if (t.kind === 'cell') {
          addDef(t.c1, t.r1, b.line, '@');
        } else if (t.kind === 'range' && hasFormula) {
          const count = (t.r2 - t.r1 + 1) * (t.c2 - t.c1 + 1);
          if (count > FILL_ENUMERATION_CAP) {
            warn(b.line, `relative fill over ${count} cells — overlap checking skipped`);
          } else {
            for (let r = t.r1; r <= t.r2; r++) for (let c = t.c1; c <= t.c2; c++) addDef(c, r, b.line, '@ fill');
          }
        } else {
          err(b.line, 'range targets accept formula content only (relative fill, SPEC §8.5/§9.4)');
        }
      }

      for (const [k, v] of Object.entries(props)) {
        if (!KNOWN_PROPS.has(k) && !k.startsWith('x-')) warn(b.line, `unknown property: ${k}`);
        if ((k === 'fill' || k === 'color') && !isColor(v)) err(b.line, `${k}: not a color: ${v}`);
        if (k === 'link' && !isSafeLink(v)) err(b.line, `link: scheme must be https:, mailto:, or internal #: ${v}`);
        if (k === 'merge') {
          if (t.kind !== 'range') err(b.line, 'merge: requires a range target');
          if (v !== true) err(b.line, 'merge: only `true` is valid');
        }
        if (k === 'spill' || k === 'array') {
          const st = parseTarget(String(v));
          if (!st || st.kind !== 'range') { err(b.line, `${k}: must be a range`); continue; }
          if (t.kind !== 'cell' || st.c1 !== t.c1 || st.r1 !== t.r1) {
            err(b.line, `${k}: range must start at the anchor cell`);
          }
          spills.push({ ...st, line: b.line });
        }
        if (k === 'rich' && !Array.isArray(v)) err(b.line, 'rich: must be a list of runs');
        if (k === 'control' && v !== 'checkbox') err(b.line, `control: unknown control ${v}`);
      }
      if (body.formula !== undefined && body.value === undefined) {
        warn(b.line, 'formula without a cached value: readers will need a calc engine to display');
      }
    }

    function validateSheetMeta(b: FenceBlock): void {
      const m = b.meta ?? {};
      for (const k of Object.keys(m)) {
        if (!SHEET_META_KEYS.has(k) && !k.startsWith('x-')) warn(b.line, `unknown {sheet} key: ${k}`);
      }
      if (m.kind !== undefined && !['worksheet', 'chart'].includes(m.kind)) err(b.line, '{sheet} kind must be worksheet | chart');
      if (m['tab-color'] !== undefined && !isColor(m['tab-color'])) err(b.line, `tab-color: not a color: ${m['tab-color']}`);
      if (m.hidden !== undefined && ![true, false, 'very'].includes(m.hidden)) err(b.line, 'hidden must be false | true | very');
      for (const key of ['freeze', 'split']) {
        if (m[key] !== undefined && !parseCell(String(m[key]))) err(b.line, `${key}: must be a cell reference`);
      }
      for (const [k, v] of Object.entries(m.cols ?? {})) {
        if (!/^[A-Z]{1,3}(:[A-Z]{1,3})?$/.test(String(k))) err(b.line, `cols key must be a column or column range: ${k}`);
        if (typeof v !== 'number' && (typeof v !== 'object' || v === null)) err(b.line, `cols.${k}: must be a width or a mapping`);
      }
      for (const k of Object.keys(m.rows ?? {})) {
        if (!/^\d+(:\d+)?$/.test(String(k))) err(b.line, `rows key must be a row or row range: ${k}`);
      }
    }
  }

  return doc;
}
