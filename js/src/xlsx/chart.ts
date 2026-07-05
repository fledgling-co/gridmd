// DrawingML chart generator (c: chartML). Emits the classic chart types —
// column/bar/line/area/pie/doughnut/scatter (+ -stacked / -stacked100) and
// combo with a secondary axis — with series styling, data labels, trendlines
// (forecast/equation/R²), error bars, axis bounds/units/gridlines/titles/
// log/reverse, legend, and the data table. ChartEx-only types (treemap,
// sunburst, waterfall, funnel, histogram, pareto, box-whisker, map) are not
// classic chartML and are reported by the caller.

import { numToCol } from '../refs';
import { resolveColor } from './units';
import type { ChartModel, Meta, ReportEntry, TableIndexEntry } from '../types';

const esc = (s: unknown): string => String(s)
  .replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;');

const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const CLASSIC_TYPES = new Set(['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'combo', 'radar', 'bubble', 'stock']);
export const CHARTEX_TYPES = new Set(['treemap', 'sunburst', 'waterfall', 'funnel', 'histogram', 'pareto', 'box-whisker', 'map']);

export function isClassicChart(type: unknown): boolean {
  return CLASSIC_TYPES.has(baseType(type));
}

function baseType(t: unknown): string {
  let b = String(t ?? 'column');
  for (const suf of ['-stacked100', '-stacked', '-3d']) if (b.endsWith(suf)) b = b.slice(0, -suf.length);
  return b;
}
function grouping(t: unknown): string | null {
  if (String(t).endsWith('-stacked100')) return 'percentStacked';
  if (String(t).endsWith('-stacked')) return 'stacked';
  return null;
}

const quoteSheet = (n: string): string => /^[A-Za-z_][A-Za-z0-9_]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`;

// GridMD data ref → absolute sheet-qualified A1 (charts cannot hold
// structured references in <c:f>; Excel stores them resolved).
export function resolveDataRef(ref: unknown, ownSheet: string, tableIndex: Map<string, TableIndexEntry>): string | null {
  const text = String(ref).trim();
  const tm = /^([A-Za-z_\\][A-Za-z0-9_.\\]*)\[([^\]#@]+)\]$/.exec(text);
  if (tm) {
    const t = tableIndex.get(tm[1]!.toLowerCase());
    if (!t) return null;
    const ci = t.columns.findIndex((c) => c.toLowerCase() === tm[2]!.trim().toLowerCase());
    if (ci === -1) return null;
    const col = numToCol(t.anchor.col + ci);
    const r1 = t.anchor.row + (t.headerRow ? 1 : 0);
    const r2 = r1 + t.bodyRows - 1;
    return `${quoteSheet(t.sheetName)}!$${col}$${r1}:$${col}$${r2}`;
  }
  const sm = /^(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))!(.+)$/.exec(text);
  const sheet = sm ? (sm[1] ?? sm[2])! : ownSheet;
  const range = (sm ? sm[3]! : text).replace(/\$/g, '');
  const abs = range.split(':').map((cell) => cell.replace(/^([A-Z]{1,3})(\d+)$/, '$$$1$$$2')).join(':');
  return `${quoteSheet(sheet)}!${abs}`;
}

interface PlotBlock { kind: string; axis: string; grouping: string | null; threeD: boolean; series: Meta[] }
interface ChartSpaceOpts { ownSheet: string; tableIndex: Map<string, TableIndexEntry>; themeColors: Record<string, string>; report: ReportEntry[] }

export function chartSpaceXml(chart: ChartModel, { ownSheet, tableIndex, themeColors }: ChartSpaceOpts): string {
  const meta = chart.meta ?? {};
  const type = chart.type ?? 'column';
  const ref = (r: unknown): string | null => resolveDataRef(r, ownSheet, tableIndex);

  // Normalize series list (`data:` shorthand → cat + val series).
  let series = meta.series;
  if (!series && meta.data !== undefined) {
    const parts = splitTopLevel(String(meta.data));
    const [cat, ...vals] = parts;
    series = vals.length ? vals.map((v, i) => ({ name: `Series ${i + 1}`, cat, val: v })) : [{ name: 'Series 1', val: cat }];
  }
  series = (series ?? []).map((s: Meta, i: number) => ({
    ...s, _idx: i, _kind: baseType(s.kind ?? type), _grouping: grouping(s.kind ?? type),
    _3d: String(s.kind ?? type).endsWith('-3d'), _axis: s.axis === 'y2' ? 'secondary' : 'primary',
  }));

  // Group series into plot blocks by (kind, axis, 3d).
  const blocks: PlotBlock[] = [];
  for (const s of series as Meta[]) {
    let blk = blocks.find((b) => b.kind === s._kind && b.axis === s._axis && b.threeD === s._3d);
    if (!blk) { blk = { kind: s._kind, axis: s._axis, grouping: s._grouping, threeD: s._3d, series: [] }; blocks.push(blk); }
    blk.series.push(s);
  }
  const any3d = blocks.some((b) => b.threeD);

  const AX: Record<string, [number, number]> = { primary: [1001, 1002], secondary: [1003, 1004] };
  const AX_SER = 1005;
  const usesSecondary = blocks.some((b) => b.axis === 'secondary');
  const color = (v: unknown, fallback?: unknown): string | null => resolveColor(v ?? fallback ?? '', themeColors);

  const solidFill = (rgb: string | null): string => rgb ? `<a:solidFill><a:srgbClr val="${rgb.slice(2)}"/></a:solidFill>` : '';
  const spPrFor = (s: Meta, isLine: boolean): string => {
    const rgb = color(s.color ?? s.fill?.color);
    const ln = s.outline
      ? `<a:ln${s.outline.width ? ` w="${Math.round(s.outline.width * 12700)}"` : ''}>${solidFill(color(s.outline.color))}</a:ln>`
      : isLine && rgb ? `<a:ln w="19050">${solidFill(rgb)}</a:ln>` : '';
    if (!rgb && !ln) return '';
    return `<c:spPr>${isLine ? '' : solidFill(rgb)}${ln}</c:spPr>`;
  };

  const dLblsXml = (l: Meta): string => {
    if (!l || l.show === false) return '';
    const POS: Record<string, string> = { center: 'ctr', 'inside-end': 'inEnd', 'inside-base': 'inBase', 'outside-end': 'outEnd', above: 't' };
    const contains = l.contains ?? ['value'];
    return '<c:dLbls>'
      + (l.numfmt ? `<c:numFmt formatCode="${esc(l.numfmt)}" sourceLinked="0"/>` : '')
      + (l.position && POS[l.position] ? `<c:dLblPos val="${POS[l.position]}"/>` : '')
      + '<c:showLegendKey val="0"/>'
      + `<c:showVal val="${contains.includes('value') ? 1 : 0}"/>`
      + `<c:showCatName val="${contains.includes('category') ? 1 : 0}"/>`
      + `<c:showSerName val="${contains.includes('series') ? 1 : 0}"/>`
      + `<c:showPercent val="${contains.includes('percent') ? 1 : 0}"/>`
      + '<c:showBubbleSize val="0"/></c:dLbls>';
  };

  const trendlineXml = (t: Meta): string => {
    if (!t) return '';
    const TYPES: Record<string, string> = { linear: 'linear', poly: 'poly', exp: 'exp', log: 'log', power: 'power', 'moving-average': 'movingAvg' };
    return '<c:trendline>'
      + `<c:trendlineType val="${TYPES[t.type] ?? 'linear'}"/>`
      + (t.order ? `<c:order val="${t.order}"/>` : '')
      + (t.window ? `<c:period val="${t.window}"/>` : '')
      + (t.forecast?.forward ? `<c:forward val="${t.forecast.forward}"/>` : '')
      + (t.forecast?.backward ? `<c:backward val="${t.forecast.backward}"/>` : '')
      + (t.intercept !== undefined ? `<c:intercept val="${t.intercept}"/>` : '')
      + `<c:dispRSqr val="${t.r2 ? 1 : 0}"/><c:dispEq val="${t.equation ? 1 : 0}"/>`
      + '</c:trendline>';
  };

  const errBarsXml = (e: Meta): string => {
    if (!e) return '';
    const TYPES: Record<string, string> = { 'std-error': 'stdErr', percentage: 'percentage', 'std-dev': 'stdDev', fixed: 'fixedVal' };
    const type = TYPES[e.type] ?? 'stdErr';
    const needsVal = type === 'percentage' || type === 'stdDev' || type === 'fixedVal';
    return '<c:errBars>'
      + `<c:errDir val="y"/><c:errBarType val="${e.dir ?? 'both'}"/>`
      + `<c:errValType val="${type}"/><c:noEndCap val="${e.cap === false ? 1 : 0}"/>`
      + (needsVal && e.value !== undefined ? `<c:val val="${e.value}"/>` : '')
      + '</c:errBars>';
  };

  const catXml = (s: Meta): string => {
    const c = s.cat ?? (series as Meta[]).find((x) => x.cat)?.cat;
    if (!c) return '';
    const f = ref(c);
    return f ? `<c:cat><c:strRef><c:f>${esc(f)}</c:f></c:strRef></c:cat>` : '';
  };
  const valXml = (s: Meta, tag = 'c:val'): string => {
    const f = ref(s.val);
    return f ? `<${tag}><c:numRef><c:f>${esc(f)}</c:f></c:numRef></${tag}>` : '';
  };

  const serXml = (s: Meta, blockKind: string): string => {
    const isLine = blockKind === 'line' || blockKind === 'stock' || blockKind === 'radar';
    const isScatter = blockKind === 'scatter' || blockKind === 'bubble';
    const name = s['name-ref']
      ? `<c:tx><c:strRef><c:f>${esc(String(s['name-ref']).replace(/^=/, ''))}</c:f></c:strRef></c:tx>`
      : s.name !== undefined ? `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>` : '';
    const marker = isLine || isScatter
      ? `<c:marker><c:symbol val="${s.marker ?? 'none'}"/></c:marker>` : '';
    const data = isScatter
      ? `${s.cat ? valXml({ val: s.cat }, 'c:xVal') : ''}${valXml(s, 'c:yVal')}${blockKind === 'bubble' && s.size ? valXml({ val: s.size }, 'c:bubbleSize') : ''}`
      : `${catXml(s)}${valXml(s)}`;
    return `<c:ser><c:idx val="${s._idx}"/><c:order val="${s._idx}"/>${name}`
      + spPrFor(s, isLine || isScatter)
      + marker
      + dLblsXml(s.labels)
      + trendlineXml(s.trendline)
      + errBarsXml(s['error-bars'])
      + data
      + (blockKind === 'line' || blockKind === 'stock' ? `<c:smooth val="${s.smooth ? 1 : 0}"/>` : '')
      + '</c:ser>';
  };

  const blockXml = (blk: PlotBlock): string => {
    const [catAx, valAx] = AX[blk.axis]!;
    const sers = blk.series.map((s) => serXml(s, blk.kind)).join('');
    const d3 = blk.threeD;
    const axes = `<c:axId val="${catAx}"/><c:axId val="${valAx}"/>${d3 ? `<c:axId val="${AX_SER}"/>` : ''}`;
    switch (blk.kind) {
      case 'column':
      case 'bar': {
        const first = blk.series[0]!;
        return `<c:bar${d3 ? '3D' : ''}Chart><c:barDir val="${blk.kind === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${blk.grouping ?? 'clustered'}"/><c:varyColors val="0"/>${sers}`
          + `<c:gapWidth val="${first.gap ?? 150}"/>`
          + (!d3 && first.overlap !== undefined ? `<c:overlap val="${first.overlap}"/>` : !d3 && blk.grouping ? '<c:overlap val="100"/>' : '')
          + (d3 ? '<c:shape val="box"/>' : '')
          + `${axes}</c:bar${d3 ? '3D' : ''}Chart>`;
      }
      case 'line':
        return `<c:line${d3 ? '3D' : ''}Chart><c:grouping val="${blk.grouping ?? 'standard'}"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>${axes}</c:line${d3 ? '3D' : ''}Chart>`;
      case 'area':
        return `<c:area${d3 ? '3D' : ''}Chart><c:grouping val="${blk.grouping ?? 'standard'}"/><c:varyColors val="0"/>${sers}${axes}</c:area${d3 ? '3D' : ''}Chart>`;
      case 'pie':
        return `<c:pie${d3 ? '3D' : ''}Chart><c:varyColors val="1"/>${sers}${d3 ? '' : '<c:firstSliceAng val="0"/>'}</c:pie${d3 ? '3D' : ''}Chart>`;
      case 'doughnut':
        return `<c:doughnutChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`;
      case 'scatter':
        return `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}${axes}</c:scatterChart>`;
      case 'radar':
        return `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${sers}${axes}</c:radarChart>`;
      case 'bubble':
        return `<c:bubbleChart><c:varyColors val="0"/>${sers}<c:bubbleScale val="100"/>${axes}</c:bubbleChart>`;
      case 'stock':
        return `<c:stockChart>${sers}<c:hiLowLines/>${axes}</c:stockChart>`;
      default:
        return '';
    }
  };

  const titleXml = (text: string | null): string => text
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>';

  const axTitle = (text: unknown): string => text
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : '';

  const axesMeta = meta.axes ?? {};
  const isPieish = blocks.every((b) => b.kind === 'pie' || b.kind === 'doughnut');

  const catAxXml = (axId: number, crossId: number, del: boolean, m: Meta = {}): string =>
    `<c:catAx><c:axId val="${axId}"/><c:scaling><c:orientation val="${m.reverse ? 'maxMin' : 'minMax'}"/></c:scaling>`
    + `<c:delete val="${del ? 1 : 0}"/><c:axPos val="b"/>`
    + (m.gridlines ? '<c:majorGridlines/>' : '')
    + axTitle(m.title)
    + (m.numfmt ? `<c:numFmt formatCode="${esc(m.numfmt)}" sourceLinked="0"/>` : '')
    + `<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="${m.labels === false ? 'none' : 'nextTo'}"/>`
    + `<c:crossAx val="${crossId}"/></c:catAx>`;

  const valAxXml = (axId: number, crossId: number, pos: string, m: Meta = {}, crossesMax = false): string =>
    `<c:valAx><c:axId val="${axId}"/><c:scaling>`
    + (m.log ? `<c:logBase val="10"/>` : '')
    + `<c:orientation val="${m.reverse ? 'maxMin' : 'minMax'}"/>`
    + (m.max !== undefined ? `<c:max val="${m.max}"/>` : '')
    + (m.min !== undefined ? `<c:min val="${m.min}"/>` : '')
    + '</c:scaling>'
    + `<c:delete val="0"/><c:axPos val="${pos}"/>`
    + (m.gridlines ? '<c:majorGridlines/>' : '')
    + axTitle(m.title)
    + (m.numfmt ? `<c:numFmt formatCode="${esc(m.numfmt)}" sourceLinked="0"/>` : '')
    + `<c:majorTickMark val="${m.ticks ?? 'out'}"/><c:minorTickMark val="${m['minor-ticks'] ?? 'none'}"/><c:tickLblPos val="nextTo"/>`
    + `<c:crossAx val="${crossId}"/>`
    + (crossesMax ? '<c:crosses val="max"/>' : '')
    + (m.unit !== undefined ? `<c:majorUnit val="${m.unit}"/>` : '')
    + (m['minor-unit'] !== undefined ? `<c:minorUnit val="${m['minor-unit']}"/>` : '')
    + '</c:valAx>';

  let axesXml = '';
  if (!isPieish) {
    axesXml = catAxXml(AX.primary![0], AX.primary![1], false, axesMeta.x ?? {})
      + valAxXml(AX.primary![1], AX.primary![0], 'l', axesMeta.y ?? {});
    if (usesSecondary) {
      axesXml += valAxXml(AX.secondary![1], AX.secondary![0], 'r', axesMeta.y2 ?? {}, true)
        + catAxXml(AX.secondary![0], AX.secondary![1], true, {});
    }
    if (any3d) {
      axesXml += `<c:serAx><c:axId val="${AX_SER}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="${AX.primary![1]}"/></c:serAx>`;
    }
  }

  const dt = meta['data-table'];
  const dTableXml = dt?.show
    ? `<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="${dt['legend-keys'] === false ? 0 : 1}"/></c:dTable>` : '';

  const legend = meta.legend ?? {};
  const LEGEND_POS: Record<string, string> = { left: 'l', right: 'r', top: 't', bottom: 'b' };
  const legendXml = legend.position === 'none'
    ? '' : `<c:legend><c:legendPos val="${LEGEND_POS[legend.position] ?? 'r'}"/><c:overlay val="${legend.overlay ? 1 : 0}"/></c:legend>`;

  return XMLD()
    + `<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">`
    + '<c:chart>'
    + titleXml(chart.title)
    + (any3d ? '<c:view3D><c:rotX val="15"/><c:rotY val="20"/><c:rAngAx val="1"/></c:view3D>' : '')
    + `<c:plotArea><c:layout/>${blocks.map(blockXml).join('')}${axesXml}${dTableXml}</c:plotArea>`
    + legendXml
    + '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>'
    + '</c:chart>'
    + '</c:chartSpace>';
}

function XMLD(): string { return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'; }

// Split on top-level commas (bracket-aware — structured refs carry commas).
export function splitTopLevel(text: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let cur = '';
  for (const ch of text) {
    if (ch === '[') depth++;
    else if (ch === ']') depth--;
    if (ch === ',' && depth === 0) { out.push(cur.trim()); cur = ''; continue; }
    cur += ch;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
}
