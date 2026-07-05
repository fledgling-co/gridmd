// Pivot table parts: pivotCacheDefinition (refreshOnLoad, no records),
// pivotCacheRecords (empty), pivotTableDefinition. Excel rebuilds the layout
// from the in-document source data on open.

import { numToCol, parseTarget } from '../refs';
import type { Meta, PivotModel, Target, TableIndexEntry } from '../types';

const esc = (s: unknown): string => String(s)
  .replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;');

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const AGG = new Map<string, string>(Object.entries({
  sum: 'sum', count: 'count', average: 'average', max: 'max', min: 'min',
  product: 'product', 'count-numbers': 'countNums', 'std-dev': 'stdDev',
  'std-devp': 'stdDevP', var: 'var', varp: 'varP',
}));
const SHOW_AS = new Map<string, string>(Object.entries({
  'percent-of-total': 'percentOfTotal', 'percent-of-row': 'percentOfRow',
  'percent-of-column': 'percentOfCol', 'percent-of': 'percentOf',
  'running-total': 'runTotal', 'difference-from': 'difference',
  'percent-difference-from': 'percentDiff', index: 'index',
}));

type HeaderFn = (tgt: Target) => string[] | null;

// Resolve the pivot source to { fields:[names], sourceXml }.
function resolveSource(source: unknown, tableIndex: Map<string, TableIndexEntry>, getHeaderCells: HeaderFn | undefined): { fields: string[]; sourceXml: string } | null {
  const text = String(source);
  const t = tableIndex.get(text.toLowerCase());
  if (t) {
    return {
      fields: t.columns,
      sourceXml: `<cacheSource type="worksheet"><worksheetSource name="${esc(t.name)}"/></cacheSource>`,
    };
  }
  const tgt = parseTarget(text);
  if (tgt && tgt.sheet && getHeaderCells) {
    const fields = getHeaderCells(tgt);
    if (fields) {
      const ref = `${numToCol(tgt.c1)}${tgt.r1}:${numToCol(tgt.c2)}${tgt.r2}`;
      return {
        fields,
        sourceXml: `<cacheSource type="worksheet"><worksheetSource ref="${ref}" sheet="${esc(tgt.sheet)}"/></cacheSource>`,
      };
    }
  }
  return null;
}

export interface PivotParts { cacheDef: string; cacheRecords: string; pivotTable: string }
export interface PivotBuildOpts {
  pivotId: number;
  cacheId: number;
  tableIndex: Map<string, TableIndexEntry>;
  getHeaderCells: HeaderFn;
  numFmtId: (code: string) => number;
}

export function buildPivotParts(pivot: PivotModel, { cacheId, tableIndex, getHeaderCells, numFmtId }: PivotBuildOpts): PivotParts | null {
  const meta = pivot.meta ?? {};
  const src = resolveSource(meta.source, tableIndex, getHeaderCells);
  if (!src) return null;
  const { fields } = src;
  const fieldIdx = (name: unknown): number => fields.findIndex((f) => f.toLowerCase() === String(name).toLowerCase());

  const rows: number[] = (meta.rows ?? []).map((r: Meta) => fieldIdx(r.field)).filter((i: number) => i !== -1);
  const cols: number[] = (meta.cols ?? []).map((c: Meta) => fieldIdx(c.field)).filter((i: number) => i !== -1);
  const pages: number[] = (meta.filters ?? []).map((f: Meta) => fieldIdx(f.field)).filter((i: number) => i !== -1);
  const values = (meta.values ?? []).map((v: Meta) => ({ ...v, _idx: fieldIdx(v.field) })).filter((v: Meta) => v._idx !== -1);

  const cacheDef = XMLDECL
    + `<pivotCacheDefinition xmlns="${MAIN_NS}" xmlns:r="${R_NS}" r:id="rId1" refreshOnLoad="1" refreshedBy="GridMD" refreshedVersion="6" createdVersion="6" minRefreshableVersion="3" recordCount="0">`
    + src.sourceXml
    + `<cacheFields count="${fields.length}">`
    + fields.map((f) => `<cacheField name="${esc(f)}" numFmtId="0"><sharedItems/></cacheField>`).join('')
    + '</cacheFields></pivotCacheDefinition>';

  const cacheRecords = `${XMLDECL}<pivotCacheRecords xmlns="${MAIN_NS}" count="0"/>`;

  const layout = meta.layout ?? 'compact';
  const compact = layout === 'compact' ? 1 : 0;
  const outline = layout === 'tabular' ? 0 : 1;
  const gtRows = meta['grand-totals']?.rows !== false;
  const gtCols = meta['grand-totals']?.cols !== false;

  const pivotFields = fields.map((_, i) => {
    let axis: string | null = null;
    if (rows.includes(i)) axis = 'axisRow';
    else if (cols.includes(i)) axis = 'axisCol';
    else if (pages.includes(i)) axis = 'axisPage';
    const isData = values.some((v: Meta) => v._idx === i);
    const attrs = [axis ? `axis="${axis}"` : '', isData ? 'dataField="1"' : '', `compact="${compact}"`, `outline="${outline}"`, 'showAll="0"'].filter(Boolean).join(' ');
    const items = axis ? '<items count="1"><item t="default"/></items>' : '';
    return `<pivotField ${attrs}>${items}</pivotField>`;
  }).join('');

  const anchor = parseTarget(pivot.anchor ?? 'A3');
  const c1 = anchor?.c1 ?? 1, r1 = anchor?.r1 ?? 3;
  const width = Math.max(2, cols.length + values.length + 1);
  const height = Math.max(4, rows.length + 6);
  const locRef = `${numToCol(c1)}${r1}:${numToCol(c1 + width)}${r1 + height}`;

  const dataFields = values.map((v: Meta) => {
    const agg = AGG.get(v.agg ?? 'sum') ?? 'sum';
    const attrs = [`name="${esc(v.name ?? `${(v.agg ?? 'sum')} of ${v.field}`)}"`, `fld="${v._idx}"`, 'baseField="0"', 'baseItem="0"'];
    if (agg !== 'sum') attrs.push(`subtotal="${agg}"`);
    if (v['show-as'] && SHOW_AS.has(v['show-as'])) attrs.push(`showDataAs="${SHOW_AS.get(v['show-as'])}"`);
    if (v.numfmt !== undefined && numFmtId) attrs.push(`numFmtId="${numFmtId(v.numfmt)}"`);
    return `<dataField ${attrs.join(' ')}/>`;
  }).join('');

  const pivotTable = XMLDECL
    + `<pivotTableDefinition xmlns="${MAIN_NS}" name="${esc(pivot.name)}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values"`
    + ` updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" compact="${compact}" compactData="${compact}" outline="${outline}" outlineData="${outline}" multipleFieldFilters="0"`
    + `${gtRows ? '' : ' rowGrandTotals="0"'}${gtCols ? '' : ' colGrandTotals="0"'}>`
    + `<location ref="${locRef}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"${pages.length ? ' rowPageCount="1" colPageCount="1"' : ''}/>`
    + `<pivotFields count="${fields.length}">${pivotFields}</pivotFields>`
    + (rows.length ? `<rowFields count="${rows.length}">${rows.map((i) => `<field x="${i}"/>`).join('')}</rowFields>` : '')
    + '<rowItems count="1"><i t="grand"><x/></i></rowItems>'
    + (cols.length ? `<colFields count="${cols.length}">${cols.map((i) => `<field x="${i}"/>`).join('')}</colFields>` : '')
    + '<colItems count="1"><i t="grand"><x/></i></colItems>'
    + (pages.length ? `<pageFields count="${pages.length}">${pages.map((i) => `<pageField fld="${i}" hier="-1"/>`).join('')}</pageFields>` : '')
    + (values.length ? `<dataFields count="${values.length}">${dataFields}</dataFields>` : '')
    + '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
    + '</pivotTableDefinition>';

  return { cacheDef, cacheRecords, pivotTable };
}
