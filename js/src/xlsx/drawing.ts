// Drawing parts (xl/drawings/drawingN.xml): chart frames, pictures,
// shapes/textboxes, and slicer graphic frames, with GridMD anchors mapped to
// twoCellAnchor / oneCellAnchor / absoluteAnchor.

import { parseTarget } from '../refs';
import { resolveColor } from './units';
import type { ShapeModel } from '../types';

const esc = (s: unknown): string => String(s)
  .replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;');

const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const EMU = 9525; // per px

const SHAPE_PRESETS = new Map<string, string>(Object.entries({
  rect: 'rect', 'rounded-rect': 'roundRect', ellipse: 'ellipse', triangle: 'triangle',
  'right-triangle': 'rtTriangle', diamond: 'diamond', pentagon: 'pentagon',
  hexagon: 'hexagon', star: 'star5', 'arrow-right': 'rightArrow',
  'arrow-left': 'leftArrow', 'arrow-up': 'upArrow', 'arrow-down': 'downArrow',
  chevron: 'chevron', callout: 'wedgeRectCallout', line: 'line',
  connector: 'bentConnector3', textbox: 'rect',
}));

export interface AnchorMarker { col: number; row: number }
export interface Anchor {
  kind: 'absolute' | 'two' | 'one';
  x?: number; y?: number; cx?: number; cy?: number;
  from?: AnchorMarker; to?: AnchorMarker;
}

// anchorStr: 'B2:K18' | 'B2' (+size) | '120,80' (+size) → anchor descriptor.
export function parseAnchor(anchorStr: string | null, size?: { w: number; h: number } | null): Anchor {
  const px = /^(\d+),(\d+)$/.exec(anchorStr ?? '');
  if (px) return { kind: 'absolute', x: +px[1]! * EMU, y: +px[2]! * EMU, cx: (size?.w ?? 480) * EMU, cy: (size?.h ?? 320) * EMU };
  const t = parseTarget(anchorStr ?? '');
  if (!t) return { kind: 'absolute', x: 0, y: 0, cx: (size?.w ?? 480) * EMU, cy: (size?.h ?? 320) * EMU };
  if (t.kind === 'range' && !size) {
    return { kind: 'two', from: { col: t.c1 - 1, row: t.r1 - 1 }, to: { col: t.c2, row: t.r2 } };
  }
  return { kind: 'one', from: { col: t.c1 - 1, row: t.r1 - 1 }, cx: (size?.w ?? 480) * EMU, cy: (size?.h ?? 320) * EMU };
}

const marker = (tag: string, m: AnchorMarker): string =>
  `<xdr:${tag}><xdr:col>${m.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${m.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`;

function wrapAnchor(anchor: Anchor, inner: string): string {
  if (anchor.kind === 'two') {
    return `<xdr:twoCellAnchor editAs="oneCell">${marker('from', anchor.from!)}${marker('to', anchor.to!)}${inner}<xdr:clientData/></xdr:twoCellAnchor>`;
  }
  if (anchor.kind === 'one') {
    return `<xdr:oneCellAnchor>${marker('from', anchor.from!)}<xdr:ext cx="${anchor.cx}" cy="${anchor.cy}"/>${inner}<xdr:clientData/></xdr:oneCellAnchor>`;
  }
  return `<xdr:absoluteAnchor><xdr:pos x="${anchor.x}" y="${anchor.y}"/><xdr:ext cx="${anchor.cx}" cy="${anchor.cy}"/>${inner}<xdr:clientData/></xdr:absoluteAnchor>`;
}

export function chartFrame(id: number, name: string, relId: string, anchor: Anchor): string {
  const inner = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
    + '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
    + `<a:graphic><a:graphicData uri="${C_NS}"><c:chart xmlns:c="${C_NS}" xmlns:r="${R_NS}" r:id="${relId}"/></a:graphicData></a:graphic>`
    + '</xdr:graphicFrame>';
  return wrapAnchor(anchor, inner);
}

export function pictureFrame(id: number, name: string, relId: string, anchor: Anchor, alt: string | undefined): string {
  const inner = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${esc(name)}" descr="${esc(alt ?? '')}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>`
    + `<xdr:blipFill><a:blip xmlns:r="${R_NS}" r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>`
    + `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${anchor.cx ?? 0}" cy="${anchor.cy ?? 0}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
  return wrapAnchor(anchor, inner);
}

export function shapeFrame(id: number, shape: ShapeModel, anchor: Anchor, themeColors: Record<string, string>): string {
  const preset = SHAPE_PRESETS.get(shape.preset) ?? 'rect';
  const meta = shape.meta ?? {};
  const fill = resolveColor(meta.fill ?? '', themeColors);
  const outline = meta.outline ? resolveColor(meta.outline.color ?? '', themeColors) : null;
  const font = meta.font ?? {};
  const fontColor = resolveColor(font.color ?? '', themeColors);
  const lines: string[] = String(meta.text ?? '').replace(/\n$/, '').split('\n');
  const rPr = `<a:rPr lang="en-US" sz="${Math.round((font.size ?? 11) * 100)}"${font.bold ? ' b="1"' : ''}${font.italic ? ' i="1"' : ''}>`
    + (fontColor ? `<a:solidFill><a:srgbClr val="${fontColor.slice(2)}"/></a:solidFill>` : '')
    + '</a:rPr>';
  const paras = lines.map((l) => `<a:p>${l === '' ? '' : `<a:r>${rPr}<a:t>${esc(l)}</a:t></a:r>`}</a:p>`).join('');
  const inner = `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${esc(shape.preset)} ${id}"/><xdr:cNvSpPr/></xdr:nvSpPr>`
    + `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${anchor.cx ?? 0}" cy="${anchor.cy ?? 0}"/></a:xfrm><a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>`
    + (fill ? `<a:solidFill><a:srgbClr val="${fill.slice(2)}"/></a:solidFill>` : shape.preset === 'textbox' ? '<a:noFill/>' : '')
    + (outline ? `<a:ln${meta.outline.width ? ` w="${Math.round(meta.outline.width * 12700)}"` : ''}><a:solidFill><a:srgbClr val="${outline.slice(2)}"/></a:solidFill></a:ln>` : shape.preset === 'textbox' ? '<a:ln><a:noFill/></a:ln>' : '')
    + '</xdr:spPr>'
    + `<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="square" rtlCol="0"/><a:lstStyle/>${paras}</xdr:txBody></xdr:sp>`;
  return wrapAnchor(anchor, inner);
}

export function chartExFrame(id: number, name: string, relId: string, anchor: Anchor): string {
  const CX_DRAWING = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
  const inner = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
    + '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
    + `<a:graphic><a:graphicData uri="${CX_DRAWING}"><cx:chart xmlns:cx="${CX_DRAWING}" xmlns:r="${R_NS}" r:id="${relId}"/></a:graphicData></a:graphic>`
    + '</xdr:graphicFrame>';
  return wrapAnchor(anchor, inner);
}

export function timesliceFrame(id: number, name: string, anchor: Anchor): string {
  const TSLE = 'http://schemas.microsoft.com/office/drawing/2012/timeslicer';
  const inner = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
    + '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
    + `<a:graphic><a:graphicData uri="${TSLE}"><tsle:timeslicer xmlns:tsle="${TSLE}" name="${esc(name)}"/></a:graphicData></a:graphic>`
    + '</xdr:graphicFrame>';
  return wrapAnchor(anchor, inner);
}

export function slicerFrame(id: number, name: string, anchor: Anchor): string {
  const inner = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
    + '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
    + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer">'
    + `<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="${esc(name)}"/>`
    + '</a:graphicData></a:graphic></xdr:graphicFrame>';
  return wrapAnchor(anchor, inner);
}

export function drawingXml(anchors: string[]): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${A_NS}">${anchors.join('')}</xdr:wsDr>`;
}
