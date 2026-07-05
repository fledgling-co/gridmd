// ChartEx (cx:) generator — the 2016+ chart family: treemap, sunburst,
// waterfall, funnel, histogram, pareto, box-whisker, map (regionMap).
// ChartEx parts carry literal data points alongside the formula refs, so the
// caller supplies a range-literal resolver backed by the workbook model.

import { createHash } from 'node:crypto';
import type { ChartModel } from '../types';

const esc = (s: unknown): string => String(s)
  .replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;');

export const CX_NS = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const LAYOUT_IDS = new Map<string, string>(Object.entries({
  treemap: 'treemap', sunburst: 'sunburst', waterfall: 'waterfall',
  funnel: 'funnel', histogram: 'clusteredColumn', pareto: 'clusteredColumn',
  'box-whisker': 'boxWhisker', map: 'regionMap',
}));
const LEGEND_POS: Record<string, string> = { left: 'l', right: 'r', top: 't', bottom: 'b' };

const guidFor = (seed: string): string => {
  const h = createHash('sha1').update(seed).digest('hex').toUpperCase();
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20, 32)}}`;
};

export const isChartExType = (t: unknown): boolean => LAYOUT_IDS.has(String(t ?? ''));

export interface LiteralResult { f: string; values: (string | number | null)[] }
export type LiteralResolver = (ref: string) => LiteralResult | null;

// chart: { type, title, meta } — resolveLiterals(ref) → { f, values } | null.
export function chartExXml(chart: ChartModel, resolveLiterals: LiteralResolver): string | null {
  const meta = chart.meta ?? {};
  const type = String(chart.type);
  const layout = LAYOUT_IDS.get(type);
  const s0 = (meta.series ?? [])[0] ?? {};
  const cat = s0.cat ?? null;
  const val = s0.val ?? null;
  const catLit = cat ? resolveLiterals(cat) : null;
  const valLit = val ? resolveLiterals(val) : null;
  if (!valLit) return null;

  const strDim = catLit
    ? `<cx:strDim type="cat"><cx:f>${esc(catLit.f)}</cx:f><cx:lvl ptCount="${catLit.values.length}">`
      + catLit.values.map((v, i) => `<cx:pt idx="${i}">${esc(v ?? '')}</cx:pt>`).join('')
      + '</cx:lvl></cx:strDim>'
    : '';
  const numType = type === 'map' ? 'colorVal' : 'val';
  const numDim = `<cx:numDim type="${numType}"><cx:f>${esc(valLit.f)}</cx:f><cx:lvl ptCount="${valLit.values.length}" formatCode="General">`
    + valLit.values.map((v, i) => (typeof v === 'number' ? `<cx:pt idx="${i}">${v}</cx:pt>` : '')).join('')
    + '</cx:lvl></cx:numDim>';

  const layoutPr = type === 'histogram' || type === 'pareto'
    ? '<cx:layoutPr><cx:binning intervalClosed="r"><cx:binCount val="0"/></cx:binning></cx:layoutPr>'
    : '';
  const seriesName = s0.name
    ? `<cx:tx><cx:txData><cx:v>${esc(s0.name)}</cx:v></cx:txData></cx:tx>` : '';
  const mainSeries = `<cx:series layoutId="${layout}" uniqueId="${guidFor(`cx:${type}:0`)}">${seriesName}${layoutPr}<cx:dataId val="0"/></cx:series>`;
  const paretoSeries = type === 'pareto'
    ? `<cx:series layoutId="paretoLine" ownerIdx="0" uniqueId="${guidFor(`cx:${type}:1`)}"><cx:axisId val="2"/></cx:series>`
    : '';

  const hasAxes = ['waterfall', 'funnel', 'histogram', 'pareto', 'box-whisker'].includes(type);
  const axes = hasAxes
    ? '<cx:axis id="0"><cx:catScaling gapWidth="1"/><cx:tickLabels/></cx:axis>'
      + '<cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis>'
      + (type === 'pareto' ? '<cx:axis id="2" hidden="0"><cx:valScaling max="1" min="0"/><cx:units unitsLabel="percentage"/><cx:tickLabels/></cx:axis>' : '')
    : '';

  const titleXml = chart.title
    ? `<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich><a:bodyPr/><a:p><a:r><a:t>${esc(chart.title)}</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>`
    : '';
  const legend = meta.legend ?? {};
  const legendXml = legend.position && legend.position !== 'none'
    ? `<cx:legend pos="${LEGEND_POS[legend.position] ?? 'r'}" align="ctr" overlay="0"/>` : '';

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<cx:chartSpace xmlns:cx="${CX_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">`
    + `<cx:chartData><cx:data id="0">${strDim}${numDim}</cx:data></cx:chartData>`
    + '<cx:chart>'
    + titleXml
    + `<cx:plotArea><cx:plotAreaRegion>${mainSeries}${paretoSeries}</cx:plotAreaRegion>${axes}</cx:plotArea>`
    + legendXml
    + '</cx:chart>'
    + '</cx:chartSpace>';
}
