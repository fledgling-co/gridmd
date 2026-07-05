// A bounded formula evaluator used to VERIFY cached values (SPEC §6: writers
// must never fabricate caches — this checks them). Covers the en-US formula
// canon subset: arithmetic/comparison/concat, A1 + cross-sheet refs, ranges,
// structured table references, defined names (refs, constants, LAMBDA),
// entity field access, array constants, and a bounded function library.
// Anything outside the subset is reported as unsupported — never guessed.

import { parseCell, colToNum, numToCol, refKey } from './refs';
import { isoToSerial } from './xlsx/units';
import { ERROR_VALUES } from './scalar';
import type { Scalar, Sheet, TableIndexEntry, WorkbookModel } from './types';

export class Unsupported extends Error {}
const BLANK: unique symbol = Symbol('blank');
type Blank = typeof BLANK;

interface ErrVal { err: string }
type Value = number | string | boolean | ErrVal | Blank | Value[];

// Always throws; typed `Value` so call sites read as an ordinary return (this
// keeps the enclosing switch case's control flow simple and fully covered).
const unsupportedFunction = (name: string): Value => { throw new Unsupported(`function ${name}`); };

interface Ctx { sheet: string; col?: number; row?: number; env?: Map<string, Value> }

interface Node {
  t: string;
  v?: string | number;
  op?: string;
  l?: Node; r?: Node; e?: Node;
  sheet?: string | null;
  ref?: string; a?: string; b?: string;
  table?: string | null; spec?: string; field?: string;
  name?: string; rawName?: string; args?: Node[];
  env?: Map<string, Value>;
}

// ---------- tokenizer ----------
const OPS = ['<>', '<=', '>=', '=', '<', '>', '+', '-', '*', '/', '^', '&', '%', '(', ')', ',', ':', ';', '!'];

type Token =
  | { t: 'num'; v: number }
  | { t: 'str' | 'sheet' | 'err' | 'arrconst' | 'bracket' | 'ident' | 'op'; v: string };
type StringToken = { t: 'str' | 'sheet' | 'err' | 'arrconst' | 'bracket' | 'ident' | 'op'; v: string };

function tokenize(src: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i]!;
    if (ch === ' ' || ch === '\n' || ch === '\t') { i++; continue; }
    if (ch === '"') {
      let j = i + 1, val = '';
      while (j < src.length) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') { val += '"'; j += 2; continue; }
          break;
        }
        val += src[j++];
      }
      out.push({ t: 'str', v: val });
      i = j + 1;
      continue;
    }
    if (ch === "'") {
      let j = i + 1, val = '';
      while (j < src.length) {
        if (src[j] === "'") {
          if (src[j + 1] === "'") { val += "'"; j += 2; continue; }
          break;
        }
        val += src[j++];
      }
      out.push({ t: 'sheet', v: val });
      i = j + 1;
      continue;
    }
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(src[i + 1] ?? ''))) {
      const m = /^\d*\.?\d+([eE][+-]?\d+)?/.exec(src.slice(i))!;
      out.push({ t: 'num', v: Number(m[0]) });
      i += m[0].length;
      continue;
    }
    if (ch === '#') {
      const err = [...ERROR_VALUES].find((e) => src.startsWith(e, i));
      if (err) { out.push({ t: 'err', v: err }); i += err.length; continue; }
      throw new Unsupported(`unknown error literal at ${src.slice(i, i + 10)}`);
    }
    if (ch === '{') {
      const end = src.indexOf('}', i);
      if (end === -1) throw new Unsupported('unterminated array constant');
      out.push({ t: 'arrconst', v: src.slice(i + 1, end) });
      i = end + 1;
      continue;
    }
    if (ch === '[') {
      let depth = 0, j = i;
      while (j < src.length) {
        if (src[j] === '[') depth++;
        else if (src[j] === ']') { depth--; if (depth === 0) break; }
        j++;
      }
      out.push({ t: 'bracket', v: src.slice(i + 1, j) });
      i = j + 1;
      continue;
    }
    if (/[A-Za-z_$\\]/.test(ch)) {
      const m = /^[A-Za-z_$\\][A-Za-z0-9_.$\\]*/.exec(src.slice(i))!;
      out.push({ t: 'ident', v: m[0] });
      i += m[0].length;
      continue;
    }
    const op = OPS.find((o) => src.startsWith(o, i));
    if (op) { out.push({ t: 'op', v: op }); i += op.length; continue; }
    throw new Unsupported(`unexpected character ${ch}`);
  }
  return out;
}

// ---------- parser (precedence climbing) ----------
function parse(src: string): Node {
  const toks = tokenize(src);
  let p = 0;
  const peek = (): Token | undefined => toks[p];
  const next = (): Token | undefined => toks[p++];
  const eat = (t: Token['t'], v?: string): Token => {
    const tok = next();
    if (!tok || tok.t !== t || (v !== undefined && tok.v !== v)) throw new Unsupported(`expected ${v ?? t}`);
    return tok;
  };

  function expr(): Node { return cmp(); }
  function cmp(): Node {
    let left = concat();
    while (peek()?.t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(String(peek()!.v))) {
      const op = String(next()!.v);
      left = { t: 'bin', op, l: left, r: concat() };
    }
    return left;
  }
  function concat(): Node {
    let left = add();
    while (peek()?.t === 'op' && peek()!.v === '&') { next(); left = { t: 'bin', op: '&', l: left, r: add() }; }
    return left;
  }
  function add(): Node {
    let left = mul();
    while (peek()?.t === 'op' && (peek()!.v === '+' || peek()!.v === '-')) {
      const op = String(next()!.v);
      left = { t: 'bin', op, l: left, r: mul() };
    }
    return left;
  }
  function mul(): Node {
    let left = pow();
    while (peek()?.t === 'op' && (peek()!.v === '*' || peek()!.v === '/')) {
      const op = String(next()!.v);
      left = { t: 'bin', op, l: left, r: pow() };
    }
    return left;
  }
  function pow(): Node {
    let left = unary();
    while (peek()?.t === 'op' && peek()!.v === '^') { next(); left = { t: 'bin', op: '^', l: left, r: unary() }; }
    return left;
  }
  function unary(): Node {
    if (peek()?.t === 'op' && (peek()!.v === '-' || peek()!.v === '+')) {
      const op = String(next()!.v);
      return op === '-' ? { t: 'neg', e: unary() } : unary();
    }
    return postfix();
  }
  function postfix(): Node {
    let e = atom();
    while (peek()?.t === 'op' && peek()!.v === '%') { next(); e = { t: 'pct', e }; }
    return e;
  }

  function atom(): Node {
    const tok = next();
    if (!tok) throw new Unsupported('unexpected end of formula');
    if (tok.t === 'num') return { t: 'num', v: tok.v };
    if (tok.t === 'str') return { t: 'str', v: tok.v };
    if (tok.t === 'err') return { t: 'err', v: tok.v };
    if (tok.t === 'arrconst') return { t: 'arrconst', v: tok.v };
    if (tok.t === 'op' && tok.v === '(') {
      const e = expr();
      eat('op', ')');
      return e;
    }
    if (tok.t === 'bracket') return { t: 'sref', table: null, spec: tok.v };
    if (tok.t === 'sheet' || tok.t === 'ident') return identLike(tok);
    throw new Unsupported(`unexpected token ${tok.v}`);
  }

  function identLike(tok: StringToken): Node {
    // Sheet-qualified reference.
    if (peek()?.t === 'op' && peek()!.v === '!') {
      next();
      const sheet = tok.v;
      const refTok = eat('ident');
      return refOrRange(String(refTok.v), sheet);
    }
    if (tok.t === 'sheet') throw new Unsupported('quoted sheet name without !');
    // Structured reference Table[...]
    if (peek()?.t === 'bracket') {
      const spec = String(next()!.v);
      return { t: 'sref', table: tok.v, spec };
    }
    // Function or LAMBDA-name call.
    if (peek()?.t === 'op' && peek()!.v === '(') {
      next();
      const args: Node[] = [];
      if (!(peek()?.t === 'op' && peek()!.v === ')')) {
        args.push(expr());
        while (peek()?.t === 'op' && peek()!.v === ',') { next(); args.push(expr()); }
      }
      eat('op', ')');
      return { t: 'call', name: tok.v.toUpperCase(), rawName: tok.v, args };
    }
    return refOrRange(tok.v, null);
  }

  function refOrRange(text: string, sheet: string | null): Node {
    const cellish = /^\$?[A-Z]{1,3}\$?\d+$/;
    if (cellish.test(text)) {
      if (peek()?.t === 'op' && peek()!.v === ':') {
        next();
        const b = String(eat('ident').v);
        if (!cellish.test(b)) throw new Unsupported(`bad range end ${b}`);
        return { t: 'range', sheet, a: text.replace(/\$/g, ''), b: b.replace(/\$/g, '') };
      }
      // Entity field access: B6.Price lexes as one ident.
      return { t: 'ref', sheet, ref: text.replace(/\$/g, '') };
    }
    const fieldMatch = /^(\$?[A-Z]{1,3}\$?\d+)\.(\w+)$/.exec(text);
    if (fieldMatch) return { t: 'field', sheet, ref: fieldMatch[1]!.replace(/\$/g, ''), field: fieldMatch[2]! };
    if (/^\$?[A-Z]{1,3}$/.test(text) && peek()?.t === 'op' && peek()!.v === ':') {
      throw new Unsupported('whole-column ranges');
    }
    return { t: 'name', name: text };
  }

  const ast = expr();
  if (p !== toks.length) throw new Unsupported(`trailing tokens in formula`);
  return ast;
}

export interface Evaluator {
  evaluateCell: (sheetName: string, ref: string) => Value;
  evaluateFormula: (formula: string, ctx: Ctx) => Value;
  BLANK: Blank;
}

// ---------- evaluator ----------
export function createEvaluator(model: WorkbookModel): Evaluator {
  const sheets = new Map<string, Sheet>(model.sheets.map((s) => [s.name.toLowerCase(), s]));
  const names = new Map<string, Record<string, unknown>>((model.fm.names ?? []).map((n: Record<string, unknown>) => [String(n.name).toLowerCase(), n]));
  const memo = new Map<string, Value>();     // 'sheet ref' -> value
  const evaluating = new Set<string>();

  const tableAt = (sheetName: string, col: number, row: number): TableIndexEntry | null => {
    for (const t of model.tableIndex.values()) {
      if (t.sheetName.toLowerCase() !== sheetName.toLowerCase()) continue;
      const r1 = t.anchor.row, r2 = r1 + (t.headerRow ? 1 : 0) + t.bodyRows - 1 + (t.total ? 1 : 0);
      const c1 = t.anchor.col, c2 = c1 + t.columns.length - 1;
      if (col >= c1 && col <= c2 && row >= r1 && row <= r2) return t;
    }
    return null;
  };

  function cellValue(sheetName: string, ref: string): Value {
    const key = `${sheetName.toLowerCase()} ${ref}`;
    if (memo.has(key)) return memo.get(key)!;
    const sheet = sheets.get(sheetName.toLowerCase());
    if (!sheet) throw new Unsupported(`unknown sheet ${sheetName}`);
    const pos = parseCell(ref)!;
    const cell = sheet.cells.get(refKey(pos.col, pos.row));
    const content = cell?.content;
    let value: Value;
    if (!content) value = BLANK;
    else if (content.formula !== undefined && content.formula !== null && !content.spillCache) {
      if (evaluating.has(key)) throw new Unsupported(`circular reference at ${sheetName}!${ref}`);
      evaluating.add(key);
      try {
        value = evalAst(parse(content.formula), { sheet: sheet.name, col: pos.col, row: pos.row });
        if (Array.isArray(value)) { const grid = value as Value[][]; value = grid[0]?.[0] ?? BLANK; } // spill anchor shows top-left
      } finally {
        evaluating.delete(key);
      }
    } else value = scalarValue(content.scalar);
    memo.set(key, value);
    return value;
  }

  function scalarValue(sc: Scalar | undefined | null): Value {
    if (!sc) return BLANK;
    switch (sc.kind) {
      case 'number': return sc.value!;
      case 'boolean': return sc.value!;
      case 'text': return sc.value!;
      case 'date':
      case 'time': return isoToSerial(String(sc.value), model.fm['date-system'] === 1904 ? 1904 : 1900);
      case 'error': return { err: String(sc.value) };
      default: return BLANK;
    }
  }

  function rangeValues(sheetName: string, a: string, b: string): Value[][] {
    const pa = parseCell(a)!, pb = parseCell(b)!;
    const out: Value[][] = [];
    for (let r = Math.min(pa.row, pb.row); r <= Math.max(pa.row, pb.row); r++) {
      const row: Value[] = [];
      for (let c = Math.min(pa.col, pb.col); c <= Math.max(pa.col, pb.col); c++) {
        row.push(cellValue(sheetName, `${numToCol(c)}${r}`));
      }
      out.push(row);
    }
    return out;
  }

  function structuredValues(node: Node, ctx: Ctx): Value {
    const table = node.table
      ? model.tableIndex.get(node.table.toLowerCase())
      : tableAt(ctx.sheet, ctx.col!, ctx.row!);
    if (!table) throw new Unsupported(`structured reference outside a table: [${node.spec}]`);
    const spec = node.spec!.trim();
    const colOf = (name: string): number => {
      const i = table.columns.findIndex((c) => c.toLowerCase() === name.trim().toLowerCase());
      if (i === -1) throw new Unsupported(`unknown table column ${name}`);
      return table.anchor.col + i;
    };
    const bodyR1 = table.anchor.row + (table.headerRow ? 1 : 0);
    const bodyR2 = bodyR1 + table.bodyRows - 1;
    if (spec.startsWith('@')) {
      const col = colOf(spec.slice(1));
      if (ctx.row! < bodyR1 || ctx.row! > bodyR2 + (table.total ? 1 : 0)) throw new Unsupported('[@col] outside table rows');
      return cellValue(table.sheetName, `${numToCol(col)}${ctx.row}`);
    }
    const totals = /^\[#Totals\],\[(.+)\]$/.exec(spec);
    if (totals) {
      const col = colOf(totals[1]!);
      return cellValue(table.sheetName, `${numToCol(col)}${bodyR2 + 1}`);
    }
    if (spec.startsWith('[')) throw new Unsupported(`structured item specifier [${spec}]`);
    const col = colOf(spec);
    const out: Value[][] = [];
    for (let r = bodyR1; r <= bodyR2; r++) out.push([cellValue(table.sheetName, `${numToCol(col)}${r}`)]);
    return out;
  }

  function nameValue(name: string, ctx: Ctx, args: Value[] | null = null): Value {
    const n = names.get(name.toLowerCase());
    if (!n) throw new Unsupported(`unknown name ${name}`);
    if (n.formula !== undefined) {
      const body = String(n.formula);
      if (/^LAMBDA\s*\(/i.test(body)) {
        const lam = parse(body);
        if (args === null) throw new Unsupported(`LAMBDA name ${name} referenced without arguments`);
        const lamArgs = lam.args!;
        const params = lamArgs.slice(0, -1).map((pn) => {
          if (pn.t === 'name') return pn.name!.toLowerCase();
          if (pn.t === 'ref') return pn.ref!.toLowerCase(); // short params like F lex as refs
          throw new Unsupported('LAMBDA parameter form');
        });
        if (args.length !== params.length) return { err: '#VALUE!' };
        const env = new Map<string, Value>(params.map((pn, i) => [pn, args[i]!]));
        return evalAst(lamArgs[lamArgs.length - 1]!, { ...ctx, env });
      }
      if (args !== null) throw new Unsupported(`call of non-LAMBDA name ${name}`);
      return evalAst(parse(body), ctx);
    }
    if (args !== null) throw new Unsupported(`call of non-LAMBDA name ${name}`);
    if (n.ref !== undefined) return evalAst(parse(String(n.ref)), ctx);
    return evalAst(parse(String(n.value)), ctx);
  }

  const num = (v: Value): number => {
    if (v === BLANK) return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string') {
      const n = Number(v);
      if (Number.isFinite(n) && v.trim() !== '') return n;
      return { err: '#VALUE!' } as unknown as number;
    }
    return v as unknown as number; // error object / array passthrough (poisons arithmetic, legacy semantics)
  };
  const isErr = (v: Value): v is ErrVal => typeof v === 'object' && v !== null && !Array.isArray(v) && 'err' in v;
  const flatNumbers = (vals: Value[]): number[] => {
    const out: number[] = [];
    for (const v of vals) {
      if (Array.isArray(v)) { out.push(...flatNumbers(v)); continue; }
      if (isErr(v)) throw v.err ? Object.assign(new Unsupported('error in range'), { excelError: v }) : v;
      if (typeof v === 'number') out.push(v);
    }
    return out;
  };
  const flatAll = (vals: Value[]): Value[] => {
    const out: Value[] = [];
    for (const v of vals) {
      if (Array.isArray(v)) out.push(...flatAll(v));
      else out.push(v);
    }
    return out;
  };

  const FNS: Record<string, (a: Value[]) => Value> = {
    SUM: (a) => flatNumbers(a).reduce((x, y) => x + y, 0),
    AVERAGE: (a) => { const n = flatNumbers(a); return n.length ? n.reduce((x, y) => x + y, 0) / n.length : { err: '#DIV/0!' }; },
    MIN: (a) => { const n = flatNumbers(a); return n.length ? Math.min(...n) : 0; },
    MAX: (a) => { const n = flatNumbers(a); return n.length ? Math.max(...n) : 0; },
    COUNT: (a) => flatNumbers(a).length,
    COUNTA: (a) => flatAll(a).filter((v) => v !== BLANK).length,
    PRODUCT: (a) => flatNumbers(a).reduce((x, y) => x * y, 1),
    ABS: (a) => Math.abs(num(a[0]!)),
    ROUND: (a) => { const f = 10 ** num(a[1]!); return Math.round(num(a[0]!) * f) / f; },
    SQRT: (a) => Math.sqrt(num(a[0]!)),
    IF: (a) => (truthy(a[0]!) ? a[1] ?? true : a[2] ?? false),
    AND: (a) => flatAll(a).every(truthy),
    OR: (a) => flatAll(a).some(truthy),
    NOT: (a) => !truthy(a[0]!),
    CONCAT: (a) => flatAll(a).map(text).join(''),
    CONCATENATE: (a) => flatAll(a).map(text).join(''),
    LEN: (a) => text(a[0]!).length,
    LEFT: (a) => text(a[0]!).slice(0, a[1] === undefined ? 1 : num(a[1])),
    RIGHT: (a) => { const n2 = a[1] === undefined ? 1 : num(a[1]); return n2 === 0 ? '' : text(a[0]!).slice(-n2); },
    UPPER: (a) => text(a[0]!).toUpperCase(),
    LOWER: (a) => text(a[0]!).toLowerCase(),
    SUBTOTAL: (a) => {
      const code = num(a[0]!) % 100;
      const rest = a.slice(1);
      const map: Record<number, string> = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 9: 'SUM' };
      if (!map[code]) throw new Unsupported(`SUBTOTAL code ${code}`);
      return FNS[map[code]!]!(rest);
    },
    SORT: (a) => {
      const flat = flatAll(a).filter((v) => v !== BLANK);
      flat.sort((x, y) => {
        const rank = (v: Value): number => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
        if (rank(x) !== rank(y)) return rank(x) - rank(y);
        if (typeof x === 'number') return x - (y as number);
        return String(x).localeCompare(String(y));
      });
      return flat.map((v) => [v]);
    },
    UNIQUE: (a) => {
      const seen = new Set<string>();
      const out: Value[][] = [];
      for (const v of flatAll(a)) {
        const k = typeof v === 'string' ? `s:${v.toLowerCase()}` : `v:${(v as { toString(): string })?.toString?.()}`;
        if (v === BLANK || seen.has(k)) continue;
        seen.add(k);
        out.push([v]);
      }
      return out;
    },
  };
  const truthy = (v: Value): boolean => {
    if (Array.isArray(v)) throw new Unsupported('array in boolean context');
    if (v === BLANK) return false;
    if (typeof v === 'string') return v !== '';
    if (isErr(v)) throw new Unsupported(`error value ${v.err}`);
    return Boolean(v);
  };
  const text = (v: Value): string => {
    if (v === BLANK) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (isErr(v)) throw new Unsupported(`error value ${v.err}`);
    return String(v);
  };

  function evalAst(node: Node, ctx: Ctx): Value {
    switch (node.t) {
      case 'num': return node.v!;
      case 'str': return node.v!;
      case 'err': return { err: String(node.v) };
      case 'arrconst': {
        return String(node.v).split(';').map((row) => row.split(',').map((el): Value => {
          const s = el.trim();
          if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
          if (/^".*"$/.test(s)) return s.slice(1, -1).replace(/""/g, '"');
          if (/^(TRUE|FALSE)$/i.test(s)) return s.toUpperCase() === 'TRUE';
          throw new Unsupported(`array constant element ${s}`);
        }));
      }
      case 'neg': { const v = evalAst(node.e!, ctx); return isErr(v) ? v : -num(v); }
      case 'pct': { const v = evalAst(node.e!, ctx); return isErr(v) ? v : num(v) / 100; }
      case 'bin': return binop(node, ctx);
      case 'ref': {
        if (ctx.env?.has(node.ref!.toLowerCase())) return ctx.env.get(node.ref!.toLowerCase())!;
        return cellValue(node.sheet ?? ctx.sheet, node.ref!);
      }
      case 'range': return rangeValues(node.sheet ?? ctx.sheet, node.a!, node.b!);
      case 'sref': return structuredValues(node, ctx);
      case 'field': {
        const sheet = sheets.get((node.sheet ?? ctx.sheet).toLowerCase());
        const pos = parseCell(node.ref!)!;
        const content = sheet?.cells.get(refKey(pos.col, pos.row))?.content;
        const fields = content?.entityFields;
        if (!fields || fields[node.field!] === undefined) return { err: '#FIELD!' };
        const v = fields[node.field!];
        return typeof v === 'number' || typeof v === 'boolean' ? v : String(v);
      }
      case 'name': {
        if (ctx.env?.has(node.name!.toLowerCase())) return ctx.env.get(node.name!.toLowerCase())!;
        return nameValue(node.name!, ctx);
      }
      case 'call': return callValue(node, ctx);
      default:
        throw new Unsupported(`node ${node.t}`);
    }
  }

  function callValue(node: Node, ctx: Ctx): Value {
    if (node.name === 'LAMBDA') throw new Unsupported('inline LAMBDA invocation');
    const args = node.args!.map((a) => evalAst(a, ctx));
    const isName = names.has(node.rawName!.toLowerCase());
    return FNS[node.name!] && !isName ? FNS[node.name!]!(args)
      : isName ? nameValue(node.rawName!, ctx, args)
      : unsupportedFunction(String(node.name));
  }

  function binop(node: Node, ctx: Ctx): Value {
    const l = evalAst(node.l!, ctx);
    const r = evalAst(node.r!, ctx);
    if (isErr(l)) return l;
    if (isErr(r)) return r;
    if (Array.isArray(l) || Array.isArray(r)) throw new Unsupported('array operand in binary op');
    switch (node.op) {
      case '+': return num(l) + num(r);
      case '-': return num(l) - num(r);
      case '*': return num(l) * num(r);
      case '/': return num(r) === 0 ? { err: '#DIV/0!' } : num(l) / num(r);
      case '^': return num(l) ** num(r);
      case '&': return text(l) + text(r);
      default: {
        // The parser only ever emits the six comparison operators here.
        const cmp = compare(l, r);
        return node.op === '=' ? cmp === 0
          : node.op === '<>' ? cmp !== 0
          : node.op === '<' ? cmp < 0
          : node.op === '>' ? cmp > 0
          : node.op === '<=' ? cmp <= 0
          : cmp >= 0; // '>='
      }
    }
  }

  function compare(l: Value, r: Value): number {
    const lv = l === BLANK ? 0 : l;
    const rv = r === BLANK ? 0 : r;
    if (typeof lv === 'string' && typeof rv === 'string') return lv.toLowerCase().localeCompare(rv.toLowerCase());
    if (typeof lv === 'number' && typeof rv === 'number') return lv - rv;
    if (typeof lv === 'boolean' || typeof rv === 'boolean') return (lv === rv) ? 0 : lv ? 1 : -1;
    // Excel type ordering: numbers < text < logicals
    const rank = (v: Value): number => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
    return rank(lv) - rank(rv);
  }

  return {
    evaluateCell: (sheetName: string, ref: string): Value => cellValue(sheetName, ref),
    evaluateFormula: (formula: string, ctx: Ctx): Value => evalAst(parse(formula), ctx),
    BLANK,
  };
}

export interface Mismatch { where: string; formula: string; cached: Value | null; computed: Value }
export interface UnsupportedItem { where: string; formula: string; reason: string }
export interface VerifyResult { checked: number; mismatches: Mismatch[]; unsupported: UnsupportedItem[] }

// ---------- cached-value verification ----------
export function verifyCachedValues(model: WorkbookModel, { tolerance = 1e-9 }: { tolerance?: number } = {}): VerifyResult {
  const ev = createEvaluator(model);
  const mismatches: Mismatch[] = [];
  const unsupported: UnsupportedItem[] = [];
  let checked = 0;

  const close = (a: number, b: number): boolean => Math.abs(a - b) <= tolerance * Math.max(1, Math.abs(a), Math.abs(b));
  const cachedActual = (cached: Scalar): Value | null => {
    switch (cached.kind) {
      case 'number': return cached.value!;
      case 'boolean': return cached.value!;
      case 'text': return cached.value!;
      case 'error': return { err: String(cached.value) };
      case 'date':
      case 'time': return isoToSerial(String(cached.value), model.fm['date-system'] === 1904 ? 1904 : 1900);
      default: return null;
    }
  };

  for (const sheet of model.sheets) {
    if (sheet.kind === 'chart') continue;
    for (const cell of sheet.cells.values()) {
      const content = cell.content;
      if (!content || content.formula === undefined || content.formula === null || content.spillCache) continue;
      if (content.cached == null) continue;
      const where = `${sheet.name}!${numToCol(cell.col)}${cell.row}`;
      let computed: Value;
      try {
        computed = ev.evaluateCell(sheet.name, `${numToCol(cell.col)}${cell.row}`);
      } catch (e: unknown) {
        if (e instanceof Unsupported) { unsupported.push({ where, formula: content.formula, reason: e.message }); continue; }
        throw e;
      }
      checked++;
      const want = cachedActual(content.cached);
      const wantErr = (want as { err?: string } | null)?.err;
      const compErr = (computed as { err?: string })?.err;
      const ok = typeof want === 'number' && typeof computed === 'number' ? close(computed, want)
        : typeof want === 'object' && want !== null ? (typeof computed === 'object' && compErr === wantErr)
        : computed === want;
      if (!ok) mismatches.push({ where, formula: content.formula, cached: want, computed });
    }
  }
  return { checked, mismatches, unsupported };
}
