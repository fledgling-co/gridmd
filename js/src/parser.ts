// GridMD document parser (SPEC.md §2–§10, Appendix A).
// Produces a block tree; semantic checks live in validate.js.

import { parseDocument as yamlDocument, visit } from 'yaml';
import type {
  AtBlock, Block, Diagnostic, FenceBlock, InfoArgs, Meta, ParsedDocument, Row, SheetBlock,
} from './types';

export const RESERVED_KINDS = new Set([
  'sheet', 'grid', 'spill-cache', 'table', 'cf', 'validation', 'filter',
  'chart', 'sparklines', 'pivot', 'slicer', 'image', 'shape', 'textbox',
  'checkbox', 'comments', 'outline', 'page', 'query', 'script', 'scenario',
  'raw',
]);

const IDENT_KEY_RE = /^(x-)?[a-z][a-z0-9-]*$/;

export function parseYaml(text: string, line: number, errors: Diagnostic[]): Meta {
  if (text.trim() === '') return {};
  const doc = yamlDocument(text, { version: '1.2', uniqueKeys: true });
  for (const e of doc.errors) {
    errors.push({ line, msg: `YAML: ${String(e.message).split('\n')[0]!}` });
  }
  let alias = false;
  let tagged = false;
  visit(doc, {
    Alias() { alias = true; },
    Scalar(_, node) { if (node.tag) tagged = true; },
    Map(_, node) { if (node.tag) tagged = true; },
    Seq(_, node) { if (node.tag) tagged = true; },
  });
  if (alias) errors.push({ line, msg: 'YAML anchors/aliases are outside the GridMD safe subset' });
  if (tagged) errors.push({ line, msg: 'YAML tags are outside the GridMD safe subset' });
  try {
    return doc.toJS() ?? {};
  } catch (e: unknown) {
    errors.push({ line, msg: `YAML: ${e instanceof Error ? e.message : String(e)}` });
    return {};
  }
}

// YAML flow-map candidate for @-directive props: must parse to a mapping in
// which every top-level key is an identifier and every value is non-null
// (SPEC Appendix A, props rule).
export function tryProps(text: string): Meta | null {
  let v: unknown;
  try {
    const doc = yamlDocument(text, { version: '1.2', uniqueKeys: true });
    if (doc.errors.length) return null;
    v = doc.toJS();
  } catch {
    return null;
  }
  if (v === null || typeof v !== 'object' || Array.isArray(v)) return null;
  for (const [k, val] of Object.entries(v)) {
    if (!IDENT_KEY_RE.test(k)) return null;
    if (val === null) return null;
  }
  return v;
}

// Right-edge props split (SPEC §9.1 / Appendix A): the last brace-balanced
// {…} group (double quotes respected) that runs to end-of-line and is
// preceded by whitespace.
export function findPropsSplit(text: string): { scalarText: string; propsText: string | null } {
  if (!text.endsWith('}')) return { scalarText: text, propsText: null };
  let inQ = false;
  let depth = 0;
  let start = -1;
  let lastGroup: [number, number] | null = null;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') { inQ = !inQ; continue; }
    if (inQ) continue;
    if (ch === '{') {
      if (depth === 0) start = i;
      depth++;
    } else if (ch === '}') {
      depth--;
      if (depth === 0 && start !== -1) lastGroup = [start, i];
      if (depth < 0) return { scalarText: text, propsText: null };
    }
  }
  if (!lastGroup) return { scalarText: text, propsText: null };
  const [s, e] = lastGroup;
  if (e !== text.length - 1 || s === 0 || text[s - 1] !== ' ') {
    return { scalarText: text, propsText: null };
  }
  return { scalarText: text.slice(0, s).trimEnd(), propsText: text.slice(s) };
}

// Pipe row → trimmed cell strings; backslash escapes the next character.
// Returns null if the line is not a well-formed pipe row.
export function splitPipeRow(rawLine: string): string[] | null {
  const line = rawLine.replace(/\s+$/, '');
  if (!line.startsWith('|') || line.length < 2) return null;
  const cells: string[] = [];
  let cell = '';
  let opened = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '\\' && i + 1 < line.length) { cell += line[i + 1]!; i++; continue; }
    if (ch === '|') {
      if (!opened) { opened = true; continue; }
      cells.push(cell.trim());
      cell = '';
      continue;
    }
    cell += ch;
  }
  if (cell.trim() !== '') return null; // no unescaped closing pipe
  return cells;
}

interface Token { v: string; q: boolean }

// Fence info string: positional args (quoted "" doubling), at-anchors,
// size WxH, key=val flags.
export function parseInfoArgs(rest: string, line: number, errors: Diagnostic[]): InfoArgs {
  const out: InfoArgs = { positional: [], flags: {}, anchor: null, size: null };
  const re = /"((?:[^"]|"")*)"|\S+/g;
  const tokens: Token[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(rest))) {
    tokens.push(m[1] !== undefined ? { v: m[1].replace(/""/g, '"'), q: true } : { v: m[0]!, q: false });
  }
  for (let k = 0; k < tokens.length; k++) {
    const tok = tokens[k]!;
    if (!tok.q && tok.v === 'at') {
      const nxt = tokens[++k];
      if (!nxt) { errors.push({ line, msg: '`at` requires an anchor' }); break; }
      out.anchor = nxt.v;
      continue;
    }
    if (!tok.q && tok.v === 'size') {
      const nxt = tokens[++k];
      const sm = nxt && /^(\d+)x(\d+)$/.exec(nxt.v);
      if (!sm) { errors.push({ line, msg: '`size` requires WxH (e.g. 480x320)' }); continue; }
      out.size = { w: Number(sm[1]), h: Number(sm[2]) };
      continue;
    }
    const fm = !tok.q && /^([A-Za-z][A-Za-z0-9-]*)=(.*)$/.exec(tok.v);
    if (fm) {
      out.flags[fm[1]!] = fm[2]!.replace(/^"(.*)"$/s, '$1').replace(/""/g, '"');
      continue;
    }
    out.positional.push(tok.v);
  }
  return out;
}

const FENCE_OPEN_RE = /^(`{3,})\{([A-Za-z][A-Za-z0-9-]*)\}(.*)$/;

function parseFence(lines: string[], i: number, m: RegExpExecArray, errors: Diagnostic[]): [FenceBlock, number] {
  const open = m[1]!.length;
  const kind = m[2]!;
  const args = parseInfoArgs(m[3] ?? '', i + 1, errors);
  const body: string[] = [];
  let j = i + 1;
  let closed = false;
  const closeRe = new RegExp(`^\`{${open},}\\s*$`);
  while (j < lines.length) {
    if (closeRe.test(lines[j]!)) { closed = true; j++; break; }
    body.push(lines[j]!);
    j++;
  }
  if (!closed) errors.push({ line: i + 1, msg: `unclosed {${kind}} fence` });
  const block: FenceBlock = { type: 'fence', kind, args, body, line: i + 1 };
  refineFence(block, errors);
  return [block, j];
}

function parseRows(bodyLines: string[], baseLine: number, errors: Diagnostic[]): Row[] {
  const rows: Row[] = [];
  for (let k = 0; k < bodyLines.length; k++) {
    const l = bodyLines[k]!;
    if (l.trim() === '') continue;
    const cells = splitPipeRow(l);
    if (cells === null) {
      errors.push({ line: baseLine + k + 1, msg: `expected a pipe row, got: ${l.slice(0, 50)}` });
      continue;
    }
    rows.push({ cells, line: baseLine + k + 1 });
  }
  return rows;
}

function refineFence(block: FenceBlock, errors: Diagnostic[]): void {
  const { kind, body, line } = block;
  const meta = (arr: string[], off: number): Meta => parseYaml(arr.join('\n'), line + off, errors);
  if (kind === 'grid' || kind === 'spill-cache') {
    block.rows = parseRows(body, line, errors);
  } else if (kind === 'table') {
    const d = body.indexOf('---');
    if (d === -1) {
      errors.push({ line, msg: '{table} requires a `---`-separated payload of pipe rows' });
      block.meta = meta(body, 1);
      block.rows = [];
    } else {
      block.meta = meta(body.slice(0, d), 1);
      block.rows = parseRows(body.slice(d + 1), line + d + 1, errors);
    }
  } else if (kind === 'script') {
    const d = body.indexOf('---');
    if (d === -1) { block.meta = {}; block.code = body.join('\n'); }
    else { block.meta = meta(body.slice(0, d), 1); block.code = body.slice(d + 1).join('\n'); }
  } else if (kind === 'raw' || kind.startsWith('x-')) {
    block.payload = body.join('\n');
  } else {
    block.meta = meta(body, 1);
  }
}

function parseAt(lines: string[], i: number, errors: Diagnostic[]): [AtBlock, number] {
  const line = lines[i]!;
  const rest = line.slice(2);
  const sp = rest.indexOf(' ');
  const targetText = sp === -1 ? rest : rest.slice(0, sp);
  const inline = sp === -1 ? '' : rest.slice(sp + 1).trim();

  // Multiline body: maximal run of blank-or-2-space-indented lines, trailing
  // blanks excluded (Appendix A, dedent rule).
  let j = i + 1;
  let taken = 0;
  let lastTake = 0;
  const acc: string[] = [];
  while (j < lines.length) {
    const l = lines[j]!;
    if (l.trim() === '') { acc.push(''); j++; taken++; continue; }
    if (/^ {2}/.test(l)) { acc.push(l.slice(2)); j++; taken++; lastTake = taken; continue; }
    break;
  }
  const bodyLines = lastTake > 0 ? acc.slice(0, lastTake) : null;
  const next = i + 1 + lastTake;

  const block: AtBlock = {
    type: 'at', targetText, line: i + 1,
    scalarText: null, props: null, body: null,
  };
  if (bodyLines) {
    const parsed = parseYaml(bodyLines.join('\n'), i + 2, errors);
    if (parsed === null || typeof parsed !== 'object' || Array.isArray(parsed)) {
      errors.push({ line: i + 2, msg: '@ directive body must be a YAML mapping' });
    } else {
      block.body = parsed;
    }
  }
  if (inline !== '') {
    if (inline.startsWith('{') && !inline.startsWith('{=')) {
      const props = tryProps(inline);
      if (props) {
        block.props = props;
        return [block, next];
      }
    }
    const { scalarText, propsText } = findPropsSplit(inline);
    if (propsText) {
      const props = tryProps(propsText);
      if (props) {
        block.props = props;
        block.scalarText = scalarText === '' ? null : scalarText;
        return [block, next];
      }
    }
    block.scalarText = inline;
  }
  return [block, next];
}

export function parseDocument(source: string, { mode = 'strict' }: { mode?: string } = {}): ParsedDocument {
  const errors: Diagnostic[] = [];
  const warnings: Diagnostic[] = [];
  const lines = source.split(/\r?\n/);
  const doc: ParsedDocument = { frontmatter: {}, workbookBlocks: [], sheets: [], errors, warnings, mode };

  if (lines[0] !== '---') {
    errors.push({ line: 1, msg: 'document must begin with `---` YAML frontmatter' });
    return doc;
  }
  let fmEnd = -1;
  for (let k = 1; k < lines.length; k++) {
    if (lines[k] === '---') { fmEnd = k; break; }
  }
  if (fmEnd === -1) {
    errors.push({ line: 1, msg: 'unterminated frontmatter (missing closing `---`)' });
    return doc;
  }
  doc.frontmatter = parseYaml(lines.slice(1, fmEnd).join('\n'), 2, errors);

  let i = fmEnd + 1;
  let cur: SheetBlock | null = null;
  const push = (b: Block): number => (cur ? cur.blocks : doc.workbookBlocks).push(b);

  while (i < lines.length) {
    const line = lines[i]!;
    if (line.trim() === '' || /^>/.test(line) || /^#{2,}(\s|$)/.test(line)) { i++; continue; }
    let m: RegExpExecArray | null;
    if ((m = /^# (.+)$/.exec(line))) {
      cur = { name: m[1]!.trim(), line: i + 1, blocks: [] };
      doc.sheets.push(cur);
      i++;
      continue;
    }
    if ((m = FENCE_OPEN_RE.exec(line))) {
      const [block, next] = parseFence(lines, i, m, errors);
      push(block);
      i = next;
      continue;
    }
    if (/^@ /.test(line)) {
      const [block, next] = parseAt(lines, i, errors);
      push(block);
      i = next;
      continue;
    }
    (mode === 'strict' ? errors : warnings).push({
      line: i + 1,
      msg: `unrecognized line: ${line.slice(0, 60)}`,
    });
    i++;
  }
  return doc;
}
