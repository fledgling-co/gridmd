// styles.xml registry: dedups numFmts / fonts / fills / borders into cellXfs,
// plus dxfs for conditional formatting.

import { resolveColor, NUMFMT_ALIASES, BUILTIN_NUMFMT_IDS } from './units';
import type { Meta } from '../types';

const esc = (s: unknown): string => String(s)
  .replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;');

const BORDER_STYLES = new Map<string, string | null>(Object.entries({
  hair: 'hair', thin: 'thin', medium: 'medium', thick: 'thick', double: 'double',
  dotted: 'dotted', dashed: 'dashed', 'dash-dot': 'dashDot',
  'dash-dot-dot': 'dashDotDot', 'medium-dashed': 'mediumDashed',
  'medium-dash-dot': 'mediumDashDot', 'medium-dash-dot-dot': 'mediumDashDotDot',
  'slant-dash-dot': 'slantDashDot', none: null,
}));

const PATTERNS = new Map<string, string>(Object.entries({
  solid: 'solid', 'gray-750': 'darkGray', 'gray-500': 'mediumGray',
  'gray-250': 'lightGray', 'gray-125': 'gray125', 'gray-0625': 'gray0625',
}));

export interface BorderEdge { style: string; color: string | null }
export interface BorderEdges {
  top: BorderEdge | null; right: BorderEdge | null; bottom: BorderEdge | null;
  left: BorderEdge | null; diagUp: BorderEdge | null; diagDown: BorderEdge | null;
}

// "thin #D6D9E0" | { style, color } → { style, color } (normalized)
export function parseBorderEdge(v: Meta, theme: Record<string, string>): BorderEdge | null {
  if (v == null) return null;
  if (typeof v === 'string') {
    const [style, color] = v.split(/\s+/);
    return { style: BORDER_STYLES.get(style!) ?? 'thin', color: resolveColor(color ?? '', theme) };
  }
  return { style: BORDER_STYLES.get(v.style) ?? 'thin', color: resolveColor(v.color ?? '', theme) };
}

export class StyleRegistry {
  theme: Record<string, string>;
  numFmts: Map<string, number>;
  fonts: string[];
  fontKeys: Map<string, number>;
  fills: string[];
  fillKeys: Map<string, number>;
  borders: string[];
  borderKeys: Map<string, number>;
  xfs: string[];
  xfKeys: Map<string, number>;
  dxfs: string[];

  constructor(themeColors: Record<string, string> = {}) {
    this.theme = themeColors;
    this.numFmts = new Map();          // code -> id (custom from 164)
    this.fonts = ['<font><sz val="11"/><name val="Calibri"/></font>'];
    this.fontKeys = new Map([[this.fonts[0]!, 0]]);
    this.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    this.fillKeys = new Map(this.fills.map((f, i) => [f, i]));
    this.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    this.borderKeys = new Map([[this.borders[0]!, 0]]);
    this.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'];
    this.xfKeys = new Map([[this.xfs[0]!, 0]]);
    this.dxfs = [];
  }

  numFmtId(code: string): number {
    const resolved = NUMFMT_ALIASES[code] ?? code;
    const builtin = BUILTIN_NUMFMT_IDS.get(resolved);
    if (builtin !== undefined) return builtin;
    if (!this.numFmts.has(resolved)) this.numFmts.set(resolved, 164 + this.numFmts.size);
    return this.numFmts.get(resolved)!;
  }

  fontXml(p: Meta): string {
    const bits: string[] = [];
    if (p.bold) bits.push('<b/>');
    if (p.italic) bits.push('<i/>');
    if (p.underline === true) bits.push('<u/>');
    else if (p.underline === 'double') bits.push('<u val="double"/>');
    else if (p.underline === 'single-accounting') bits.push('<u val="singleAccounting"/>');
    else if (p.underline === 'double-accounting') bits.push('<u val="doubleAccounting"/>');
    if (p.strike) bits.push('<strike/>');
    if (p.sub) bits.push('<vertAlign val="subscript"/>');
    if (p.super) bits.push('<vertAlign val="superscript"/>');
    bits.push(`<sz val="${p.size ?? 11}"/>`);
    const color = resolveColor(p.color ?? '', this.theme);
    if (color) bits.push(`<color rgb="${color}"/>`);
    bits.push(`<name val="${esc(p.font ?? 'Calibri')}"/>`);
    return `<font>${bits.join('')}</font>`;
  }

  fontId(p: Meta): number {
    if (!p.bold && !p.italic && !p.underline && !p.strike && !p.sub && !p.super
      && p.size === undefined && p.color === undefined && p.font === undefined) return 0;
    const xml = this.fontXml(p);
    if (!this.fontKeys.has(xml)) { this.fonts.push(xml); this.fontKeys.set(xml, this.fonts.length - 1); }
    return this.fontKeys.get(xml)!;
  }

  fillId(p: Meta): number {
    if (p.fill === undefined) return 0;
    const fg = resolveColor(p.fill, this.theme);
    if (!fg) return 0;
    const pattern = PATTERNS.get(p.pattern ?? 'solid') ?? 'solid';
    const bg = p.fill2 ? `<bgColor rgb="${resolveColor(p.fill2, this.theme)}"/>` : '';
    const xml = `<fill><patternFill patternType="${pattern}"><fgColor rgb="${fg}"/>${bg}</patternFill></fill>`;
    if (!this.fillKeys.has(xml)) { this.fills.push(xml); this.fillKeys.set(xml, this.fills.length - 1); }
    return this.fillKeys.get(xml)!;
  }

  borderId(edges: BorderEdges | null): number {
    // edges: { top, right, bottom, left, diagUp, diagDown } of parsed edges
    if (!edges || Object.values(edges).every((e) => !e)) return 0;
    const side = (name: string, e: BorderEdge | null): string => {
      if (!e || !e.style) return `<${name}/>`;
      const c = e.color ? `<color rgb="${e.color}"/>` : '';
      return `<${name} style="${e.style}">${c}</${name}>`;
    };
    const diag = edges.diagUp || edges.diagDown;
    const attrs = `${edges.diagUp ? ' diagonalUp="1"' : ''}${edges.diagDown ? ' diagonalDown="1"' : ''}`;
    const xml = `<border${attrs}>${side('left', edges.left)}${side('right', edges.right)}${side('top', edges.top)}${side('bottom', edges.bottom)}${side('diagonal', diag)}</border>`;
    if (!this.borderKeys.has(xml)) { this.borders.push(xml); this.borderKeys.set(xml, this.borders.length - 1); }
    return this.borderKeys.get(xml)!;
  }

  alignmentXml(p: Meta): string {
    const a: string[] = [];
    const H: Record<string, string> = { left: 'left', center: 'center', right: 'right', justify: 'justify', fill: 'fill', 'center-across': 'centerContinuous', distributed: 'distributed' };
    const V: Record<string, string> = { top: 'top', middle: 'center', bottom: 'bottom', justify: 'justify', distributed: 'distributed' };
    if (p.align) a.push(`horizontal="${H[p.align] ?? 'general'}"`);
    if (p.valign) a.push(`vertical="${V[p.valign] ?? 'bottom'}"`);
    if (p.wrap) a.push('wrapText="1"');
    if (p.shrink) a.push('shrinkToFit="1"');
    if (p.indent) a.push(`indent="${p.indent}"`);
    if (p.rotation !== undefined) a.push(`textRotation="${p.rotation === 'vertical' ? 255 : p.rotation}"`);
    return a.length ? `<alignment ${a.join(' ')}/>` : '';
  }

  // props (already style-resolved) → cellXf index
  xfId(p: Meta, borders: BorderEdges | null): number {
    const numFmtId = p.numfmt !== undefined ? this.numFmtId(p.numfmt) : 0;
    const fontId = this.fontId(p);
    const fillId = this.fillId(p);
    const borderId = this.borderId(borders);
    const align = this.alignmentXml(p);
    const prot = (p.locked === false || p.hidden === true)
      ? `<protection${p.locked === false ? ' locked="0"' : ''}${p.hidden ? ' hidden="1"' : ''}/>` : '';
    const applies = [
      numFmtId ? 'applyNumberFormat="1"' : '',
      fontId ? 'applyFont="1"' : '',
      fillId ? 'applyFill="1"' : '',
      borderId ? 'applyBorder="1"' : '',
      align ? 'applyAlignment="1"' : '',
      prot ? 'applyProtection="1"' : '',
    ].filter(Boolean).join(' ');
    const inner = align + prot;
    const xml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}"${applies ? ' ' + applies : ''}${inner ? `>${inner}</xf>` : '/>'}`;
    if (!this.xfKeys.has(xml)) { this.xfs.push(xml); this.xfKeys.set(xml, this.xfs.length - 1); }
    return this.xfKeys.get(xml)!;
  }

  // CF format: {…} → dxf index
  dxfId(fmt: Meta): number {
    const font: string[] = [];
    if (fmt.bold) font.push('<b/>');
    if (fmt.italic) font.push('<i/>');
    const color = resolveColor(fmt.color ?? '', this.theme);
    if (color) font.push(`<color rgb="${color}"/>`);
    const fill = resolveColor(fmt.fill ?? '', this.theme);
    const xml = '<dxf>'
      + (font.length ? `<font>${font.join('')}</font>` : '')
      + (fill ? `<fill><patternFill><bgColor rgb="${fill}"/></patternFill></fill>` : '')
      + '</dxf>';
    this.dxfs.push(xml);
    return this.dxfs.length - 1;
  }

  toXml(): string {
    const numFmtXml = this.numFmts.size
      ? `<numFmts count="${this.numFmts.size}">${[...this.numFmts.entries()].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`).join('')}</numFmts>`
      : '';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + numFmtXml
      + `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>`
      + `<fills count="${this.fills.length}">${this.fills.join('')}</fills>`
      + `<borders count="${this.borders.length}">${this.borders.join('')}</borders>`
      + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      + `<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>`
      + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
      + `<dxfs count="${this.dxfs.length}">${this.dxfs.join('')}</dxfs>`
      + '</styleSheet>';
  }
}
