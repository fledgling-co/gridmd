// XLSX package writer: GridMD model → .xlsx buffer.
// Full-feature emission: worksheet core + charts (chartML), chart sheets,
// pivots (refresh-on-load caches), sparklines (x14), slicers (x14/x15),
// images, shapes/textboxes, threaded comments, scenarios, filters, sort
// state. Features with no documented OOXML form are carried in-package
// (customXml/gridmdCarry1.xml) — loud, never silent (SPEC §11 / INTEROP §1).

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { createHash } from 'node:crypto';
import { zipWrite } from './zip';
import type { ZipEntry } from './zip';
import { StyleRegistry, parseBorderEdge } from './styles';
import type { BorderEdges } from './styles';
import { isoToSerial, pxToColWidth, pxToPt, cmToInch, resolveColor } from './units';
import { numToCol, colToNum, parseTarget, parseCell, refKey } from '../refs';
import { chartSpaceXml, isClassicChart, resolveDataRef } from './chart';
import { chartExXml, isChartExType } from './chartex';
import { parseAnchor, chartFrame, chartExFrame, pictureFrame, shapeFrame, slicerFrame, timesliceFrame, drawingXml } from './drawing';
import { buildPivotParts } from './pivot';
import type {
  CarryEntry, Cell, CellContent, ChartModel, ImageModel, Meta, ReportEntry,
  Sheet, SparklineModel, Target, ValidationBlock, WorkbookModel,
} from '../types';

const esc = (s: unknown): string => String(s)
  .replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;');
const A1 = (col: number, row: number): string => `${numToCol(col)}${row}`;
const XMLDECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const X15_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main';
const TC_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';

const CT = {
  workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  worksheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  chartsheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
  comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  pivotTable: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
  pivotCacheDef: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  pivotCacheRec: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
  threaded: 'application/vnd.ms-excel.threadedcomments+xml',
  person: 'application/vnd.ms-excel.person+xml',
  slicer: 'application/vnd.ms-excel.slicer+xml',
  slicerCache: 'application/vnd.ms-excel.slicerCache+xml',
  chartEx: 'application/vnd.ms-office.chartex+xml',
  timeline: 'application/vnd.ms-excel.timeline+xml',
  timelineCache: 'application/vnd.ms-excel.timelineCacheDefinition+xml',
};
const REL = {
  worksheet: `${R_NS}/worksheet`, chartsheet: `${R_NS}/chartsheet`,
  styles: `${R_NS}/styles`, table: `${R_NS}/table`, comments: `${R_NS}/comments`,
  vml: `${R_NS}/vmlDrawing`, hyperlink: `${R_NS}/hyperlink`, drawing: `${R_NS}/drawing`,
  chart: `${R_NS}/chart`, image: `${R_NS}/image`, pivotTable: `${R_NS}/pivotTable`,
  pivotCacheDef: `${R_NS}/pivotCacheDefinition`, pivotCacheRec: `${R_NS}/pivotCacheRecords`,
  threaded: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
  person: 'http://schemas.microsoft.com/office/2017/10/relationships/person',
  slicer: 'http://schemas.microsoft.com/office/2007/relationships/slicer',
  slicerCache: 'http://schemas.microsoft.com/office/2007/relationships/slicerCache',
  chartEx: 'http://schemas.microsoft.com/office/2014/relationships/chartEx',
  timeline: 'http://schemas.microsoft.com/office/2011/relationships/timeline',
  timelineCache: 'http://schemas.microsoft.com/office/2011/relationships/timelineCacheDefinition',
};

const PAPER: Record<string, number> = { letter: 1, tabloid: 3, legal: 5, a3: 8, a4: 9, a5: 11 };
const ICON_SETS = new Map<string, string>(Object.entries({
  '3-arrows': '3Arrows', '3-arrows-gray': '3ArrowsGray', '3-flags': '3Flags',
  '3-traffic-lights': '3TrafficLights1', '3-traffic-lights-rimmed': '3TrafficLights2',
  '3-signs': '3Signs', '3-symbols': '3Symbols', '3-symbols-circled': '3Symbols2',
  '3-stars': '3Stars', '3-triangles': '3Triangles',
  '4-arrows': '4Arrows', '4-arrows-gray': '4ArrowsGray', '4-red-to-black': '4RedToBlack',
  '4-ratings': '4Rating', '4-traffic-lights': '4TrafficLights',
  '5-arrows': '5Arrows', '5-arrows-gray': '5ArrowsGray', '5-ratings': '5Rating',
  '5-quarters': '5Quarters', '5-boxes': '5Boxes',
}));
const SUBTOTAL_FN = new Map<number, string>([[101, 'average'], [102, 'countNums'], [103, 'count'],
  [104, 'max'], [105, 'min'], [106, 'product'], [107, 'stdDev'], [109, 'sum'], [110, 'var']]);
const CELLIS_OPS = new Map<string, string>(Object.entries({
  '=': 'equal', '<>': 'notEqual', '>': 'greaterThan', '>=': 'greaterThanOrEqual',
  '<': 'lessThan', '<=': 'lessThanOrEqual',
}));
const FILTER_OPS = new Map<string, string>(Object.entries({
  '=': 'equal', '<>': 'notEqual', '>': 'greaterThan', '>=': 'greaterThanOrEqual',
  '<': 'lessThan', '<=': 'lessThanOrEqual',
}));
const TIME_PERIODS = new Map<string, string>(Object.entries({
  yesterday: 'yesterday', today: 'today', tomorrow: 'tomorrow',
  'last-7-days': 'last7Days', 'this-week': 'thisWeek', 'last-week': 'lastWeek',
  'next-week': 'nextWeek', 'this-month': 'thisMonth', 'last-month': 'lastMonth',
  'next-month': 'nextMonth',
}));
const PROTECT_ATTRS = new Map<string, string>(Object.entries({
  'format-cells': 'formatCells', 'format-columns': 'formatColumns', 'format-rows': 'formatRows',
  'insert-columns': 'insertColumns', 'insert-rows': 'insertRows', 'insert-hyperlinks': 'insertHyperlinks',
  'delete-columns': 'deleteColumns', 'delete-rows': 'deleteRows', sort: 'sort',
  autofilter: 'autoFilter', 'pivot-tables': 'pivotTables', objects: 'objects', scenarios: 'scenarios',
}));

const quoteSheet = (name: string): string => /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
const guidFor = (seed: string): string => {
  const h = createHash('sha1').update(seed).digest('hex').toUpperCase();
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20, 32)}}`;
};

interface RowCell { col: number; xml: string }
interface RowProp { ht?: number; hidden?: boolean; level?: number }
interface PivotReg { cacheId: number; sheetIdx: number; sheetName: string; anchor: string | null }
interface ChartEmit { part: string; ex: boolean }
interface ImageMedia { ext: string; ct: string; data: Buffer }
interface WriteResult { buffer: Buffer; report: ReportEntry[] }

export function writeXlsx(model: WorkbookModel): WriteResult {
  const { fm, themeColors, sheets, report, rawParts, carry, tableIndex, baseDir = '.' } = model;
  const dateSystem = fm['date-system'] === 1904 ? 1904 : 1900;
  const styles = new StyleRegistry(themeColors);
  const entries: ZipEntry[] = [];
  const overrides: [string, string][] = [];
  const mediaDefaults = new Set<string>();
  const carryList: CarryEntry[] = [...(carry ?? [])];
  let wsCount = 0, csCount = 0, chartCount = 0, chartExCount = 0, drawingCount = 0, tableCount = 0;
  let commentCount = 0, tcCount = 0, pivotCount = 0, slicerCount = 0, mediaCount = 0, timelineCount = 0;
  let anyVml = false;
  const wbPivotCaches: { cacheId: number; target: string }[] = [];  // { cacheId, target }
  const wbSlicerCaches: { target: string }[] = []; // { target }
  const wbTimelineCaches: { target: string }[] = []; // { target }
  const persons = new Map<string, string>(); // name -> guid

  // Pre-assign table part ids (slicers reference tables across sheets).
  const tableIds = new Map<string, number>();
  {
    let id = 0;
    for (const s of sheets) for (const t of s.tables) tableIds.set(t.name.toLowerCase(), ++id);
  }

  // Pivot range sources read header names straight from the model cells.
  const getHeaderCells = (tgt: Target): string[] | null => {
    const sheet = sheets.find((x) => x.name.toLowerCase() === String(tgt.sheet ?? '').toLowerCase());
    if (!sheet) return null;
    const out: string[] = [];
    for (let c = tgt.c1; c <= tgt.c2; c++) {
      const cell = sheet.cells.get(refKey(c, tgt.r1));
      const v = cell?.content?.scalar?.value;
      if (typeof v !== 'string' || v === '') return null;
      out.push(v);
    }
    return out;
  };

  const addPart = (name: string, data: Buffer | string, contentType: string | null): void => {
    entries.push({ name, data });
    if (contentType) overrides.push([`/${name}`, contentType]);
  };
  const relsXml = (rels: string[]): string =>
    `${XMLDECL}<Relationships xmlns="${PKG_REL_NS}">${rels.join('')}</Relationships>`;

  // Pivot registry (name → cacheId/sheet) — timelines + PivotCharts need it
  // before/across sheet rendering. cacheIds assigned in emission order.
  const pivotRegistry = new Map<string, PivotReg>();
  {
    let cid = 0;
    sheets.forEach((sh, i) => {
      if (sh.kind === 'chart') return;
      for (const p of sh.pivots) {
        cid++;
        pivotRegistry.set(String(p.name).toLowerCase(), { cacheId: cid, sheetIdx: i, sheetName: sh.name, anchor: p.anchor });
      }
    });
  }

  // Literal range values for ChartEx data points, read from the model cells.
  const rangeLiterals = (refText: string, ownSheet: string): { f: string; values: (string | number | null)[] } | null => {
    const f = resolveDataRef(refText, ownSheet, tableIndex);
    if (!f) return null;
    const m = /^(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))!\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)$/.exec(f);
    if (!m) return null;
    const sheet = sheets.find((x) => x.name.toLowerCase() === (m[1] ?? m[2])!.toLowerCase());
    if (!sheet) return null;
    const c1 = colToNum(m[3]!), r1 = Number(m[4]), c2 = colToNum(m[5]!), r2 = Number(m[6]);
    const values: (string | number | null)[] = [];
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
      const content = sheet.cells.get(refKey(c, r))?.content;
      const sc = content?.scalar ?? content?.cached;
      values.push(sc?.kind === 'number' ? (sc.value as number) : sc?.value !== undefined ? String(sc.value) : null);
    }
    return { f, values };
  };

  const rendered = sheets.map((s) => s.kind === 'chart' ? renderChartSheet(s) : renderSheet(s));

  // ---- workbook ----
  const definedNames: string[] = [];
  for (const n of fm.names ?? []) {
    definedNames.push(`<definedName name="${esc(n.name)}"${n.hidden ? ' hidden="1"' : ''}>${esc(n.ref ?? n.formula ?? n.value ?? '')}</definedName>`);
  }
  sheets.forEach((s, i) => {
    for (const n of s.meta.names ?? []) {
      definedNames.push(`<definedName name="${esc(n.name)}" localSheetId="${i}">${esc(n.ref ?? n.formula ?? n.value ?? '')}</definedName>`);
    }
    if (s.page?.['print-area']) {
      const t = parseTarget(s.page['print-area']);
      if (t) definedNames.push(`<definedName name="_xlnm.Print_Area" localSheetId="${i}">${esc(`${quoteSheet(s.name)}!$${numToCol(t.c1)}$${t.r1}:$${numToCol(t.c2)}$${t.r2}`)}</definedName>`);
    }
    const pt = s.page?.['print-titles'];
    if (pt) {
      const parts: string[] = [];
      if (pt.rows) parts.push(`${quoteSheet(s.name)}!$${pt.rows.split(':')[0]}:$${pt.rows.split(':')[1]}`);
      if (pt.cols) parts.push(`${quoteSheet(s.name)}!$${pt.cols.split(':')[0]}:$${pt.cols.split(':')[1]}`);
      if (parts.length) definedNames.push(`<definedName name="_xlnm.Print_Titles" localSheetId="${i}">${esc(parts.join(','))}</definedName>`);
    }
  });

  const wbRels: string[] = [];
  let wbRelId = 0;
  const nextWbRel = (): string => `rId${++wbRelId}`;
  const sheetEls = sheets.map((s, i) => {
    const rid = nextWbRel();
    wbRels.push(`<Relationship Id="${rid}" Type="${s.kind === 'chart' ? REL.chartsheet : REL.worksheet}" Target="${rendered[i]!.target}"/>`);
    const state = s.meta.hidden === true ? ' state="hidden"' : s.meta.hidden === 'very' ? ' state="veryHidden"' : '';
    return `<sheet name="${esc(s.name)}"${state} sheetId="${i + 1}" r:id="${rid}"/>`;
  });
  const stylesRid = nextWbRel();
  wbRels.push(`<Relationship Id="${stylesRid}" Type="${REL.styles}" Target="styles.xml"/>`);
  const pivotCachesXml = wbPivotCaches.length
    ? `<pivotCaches>${wbPivotCaches.map((p) => {
      const rid = nextWbRel();
      wbRels.push(`<Relationship Id="${rid}" Type="${REL.pivotCacheDef}" Target="${p.target}"/>`);
      return `<pivotCache cacheId="${p.cacheId}" r:id="${rid}"/>`;
    }).join('')}</pivotCaches>` : '';
  let slicerCachesExt = '';
  const wbExts: string[] = [];
  if (wbSlicerCaches.length) {
    const refs = wbSlicerCaches.map((sc) => {
      const rid = nextWbRel();
      wbRels.push(`<Relationship Id="${rid}" Type="${REL.slicerCache}" Target="${sc.target}"/>`);
      return `<x14:slicerCache r:id="${rid}"/>`;
    }).join('');
    wbExts.push(`<ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x15="${X15_NS}"><x15:slicerCaches xmlns:x14="${X14_NS}">${refs}</x15:slicerCaches></ext>`);
  }
  if (wbTimelineCaches.length) {
    const refs = wbTimelineCaches.map((tc) => {
      const rid = nextWbRel();
      wbRels.push(`<Relationship Id="${rid}" Type="${REL.timelineCache}" Target="${tc.target}"/>`);
      return `<x15:timelineCacheRef r:id="${rid}"/>`;
    }).join('');
    wbExts.push(`<ext uri="{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}" xmlns:x15="${X15_NS}"><x15:timelineCacheRefs xmlns:r="${R_NS}">${refs}</x15:timelineCacheRefs></ext>`);
  }
  if (wbExts.length) slicerCachesExt = `<extLst>${wbExts.join('')}</extLst>`;
  if (persons.size) {
    const rid = nextWbRel();
    wbRels.push(`<Relationship Id="${rid}" Type="${REL.person}" Target="persons/person.xml"/>`);
    const list = [...persons.entries()].map(([name, id]) =>
      `<person displayName="${esc(name)}" id="${id}" userId="${esc(name)}" providerId="None"/>`).join('');
    addPart('xl/persons/person.xml', `${XMLDECL}<personList xmlns="${TC_NS}">${list}</personList>`, CT.person);
  }

  const calcMode = fm.calc?.mode === 'manual' ? 'manual' : fm.calc?.mode === 'auto-no-tables' ? 'autoNoTable' : 'auto';
  const workbook = XMLDECL
    + `<workbook xmlns="${MAIN_NS}" xmlns:r="${R_NS}">`
    + (dateSystem === 1904 ? '<workbookPr date1904="1"/>' : '<workbookPr/>')
    + (fm.protection?.structure ? '<workbookProtection lockStructure="1"/>' : '')
    + `<sheets>${sheetEls.join('')}</sheets>`
    + (definedNames.length ? `<definedNames>${definedNames.join('')}</definedNames>` : '')
    + `<calcPr calcId="0" calcMode="${calcMode}" fullCalcOnLoad="1"${fm.calc?.iterative?.enabled ? ` iterate="1" iterateCount="${fm.calc.iterative['max-iterations'] ?? 100}" iterateDelta="${fm.calc.iterative['max-change'] ?? 0.001}"` : ''}/>`
    + pivotCachesXml
    + slicerCachesExt
    + '</workbook>';

  addPart('xl/workbook.xml', workbook, CT.workbook);
  addPart('xl/_rels/workbook.xml.rels', relsXml(wbRels), null);
  addPart('xl/styles.xml', styles.toXml(), CT.styles);

  const created = fm.properties?.created ? `${fm.properties.created}`.slice(0, 10) + 'T00:00:00Z' : '2026-01-01T00:00:00Z';
  addPart('docProps/core.xml',
    `${XMLDECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(fm.title ?? '')}</dc:title><dc:creator>${esc(fm.properties?.author ?? 'GridMD')}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${esc(created)}</dcterms:created></cp:coreProperties>`,
    CT.core);
  addPart('docProps/app.xml',
    `${XMLDECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>GridMD</Application><Company>${esc(fm.properties?.company ?? '')}</Company></Properties>`,
    CT.app);

  for (const rp of rawParts) {
    const data = rp.encoding === 'base64' ? Buffer.from((rp.payload ?? '').replace(/\s+/g, ''), 'base64') : (rp.payload ?? '');
    addPart(rp.part, data, null);
  }
  if (carryList.length) {
    const json = JSON.stringify({ gridmd: fm.gridmd ?? '0.1', carried: carryList }, null, 1);
    addPart('customXml/gridmdCarry1.xml', `${XMLDECL}<gridmdCarry xmlns="urn:gridmd:carry">${esc(json)}</gridmdCarry>`, null);
  }

  const contentTypes = XMLDECL
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + (anyVml ? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' : '')
    + [...mediaDefaults].map((s) => JSON.parse(s) as [string, string]).map(([ext, ct]) => `<Default Extension="${ext}" ContentType="${ct}"/>`).join('')
    + overrides.map(([p, t]) => `<Override PartName="${p}" ContentType="${t}"/>`).join('')
    + '</Types>';

  const rootRels = relsXml([
    `<Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="xl/workbook.xml"/>`,
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`,
    `<Relationship Id="rId3" Type="${R_NS}/extended-properties" Target="docProps/app.xml"/>`,
  ]);

  const buffer = zipWrite([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: rootRels },
    ...entries,
  ]);
  return { buffer, report };

  // ================= chart emission (shared) =================
  function emitChart(chart: ChartModel, ownSheetName: string): ChartEmit | null {
    if (isChartExType(chart.type)) {
      const xml = chartExXml(chart, (ref) => rangeLiterals(ref, ownSheetName));
      if (!xml) {
        report.push({ line: chart.line, feature: `{chart} ${chart.type}`, action: 'not-emitted', note: 'ChartEx requires a resolvable val: range' });
        return null;
      }
      chartExCount++;
      const partName = `xl/charts/chartEx${chartExCount}.xml`;
      addPart(partName, xml, CT.chartEx);
      return { part: partName, ex: true };
    }
    if (!isClassicChart(chart.type)) {
      carryList.push({ kind: 'chart', line: chart.line, feature: `{chart} ${chart.type}`, meta: chart.meta });
      report.push({ line: chart.line, feature: `{chart} ${chart.type}`, action: 'carried', note: 'unknown chart type; definition carried in-package' });
      return null;
    }
    let effective = chart;
    let pivotSource = '';
    if (chart.meta?.pivot !== undefined) {
      const reg = pivotRegistry.get(String(chart.meta.pivot).toLowerCase());
      if (!reg) {
        carryList.push({ kind: 'chart', line: chart.line, feature: '{chart} pivot-bound', meta: chart.meta });
        report.push({ line: chart.line, feature: '{chart} (PivotChart)', action: 'carried', note: 'pivot target not found' });
        return null;
      }
      // Native PivotChart: declare the pivotSource; series plot the pivot
      // output area and Excel repopulates them on refresh.
      const t = parseTarget(reg.anchor ?? 'A3') ?? { c1: 1, r1: 3 };
      const dataCol = numToCol(t.c1 + 1);
      const valRange = `${quoteSheet(reg.sheetName)}!$${dataCol}$${t.r1 + 1}:$${dataCol}$${t.r1 + 4}`;
      pivotSource = `<c:pivotSource><c:name>[]${esc(reg.sheetName)}!${esc(String(chart.meta.pivot))}</c:name><c:fmtId val="0"/></c:pivotSource>`;
      effective = { ...chart, meta: { ...chart.meta, pivot: undefined, series: [{ name: String(chart.meta.pivot), val: valRange }] } };
    }
    chartCount++;
    const partName = `xl/charts/chart${chartCount}.xml`;
    let xml = chartSpaceXml(effective, { ownSheet: ownSheetName, tableIndex, themeColors, report });
    if (pivotSource) xml = xml.replace('<c:chart>', `${pivotSource}<c:chart>`);
    addPart(partName, xml, CT.chart);
    return { part: partName, ex: false };
  }

  // ================= worksheet =================
  function renderSheet(s: Sheet): { target: string } {
    wsCount++;
    const partName = `xl/worksheets/sheet${wsCount}.xml`;
    const rels: string[] = [];
    let relId = 0;
    const nextRel = (): string => `rId${++relId}`;

    // ---- cells ----
    const rowsMap = new Map<number, RowCell[]>();
    let minC = Infinity, minR = Infinity, maxC = 0, maxR = 0;
    for (const cell of s.cells.values()) {
      const props: Meta = Object.assign({}, ...cell.patches);
      const content = cell.content;
      const kindOf = content?.scalar?.kind ?? content?.cached?.kind;
      if (props.numfmt === undefined && (kindOf === 'date' || kindOf === 'time')) {
        const v = content?.scalar?.value ?? content?.cached?.value ?? '';
        props.numfmt = kindOf === 'time' ? 'hh:mm:ss' : String(v).includes('T') ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd';
      }
      const xf = styles.xfId(props, collectEdges(props));
      const xml = cellXml(cell, content, xf);
      if (!xml) continue;
      if (!rowsMap.has(cell.row)) rowsMap.set(cell.row, []);
      rowsMap.get(cell.row)!.push({ col: cell.col, xml });
      minC = Math.min(minC, cell.col); maxC = Math.max(maxC, cell.col);
      minR = Math.min(minR, cell.row); maxR = Math.max(maxR, cell.row);
    }
    const dimension = maxC ? `${A1(minC, minR)}:${A1(maxC, maxR)}` : 'A1';

    const rowProps = new Map<number, RowProp>();
    const rp = (r: number): RowProp => { if (!rowProps.has(r)) rowProps.set(r, {}); return rowProps.get(r)!; };
    for (const [k, v] of Object.entries(s.meta.rows ?? {})) {
      const [a, b] = String(k).split(':').map(Number);
      for (let r = a!; r <= (b ?? a!); r++) {
        const o = rp(r);
        if (typeof v === 'number') o.ht = pxToPt(v);
        else if (v && typeof v === 'object') {
          const vv = v as Meta;
          if (vv.height !== undefined) o.ht = pxToPt(vv.height);
          if (vv.hidden) o.hidden = true;
          if (vv.group) o.level = vv.group;
        }
      }
    }
    for (const g of s.outline.rows) {
      const [a, b] = String(g.range).split(':').map(Number);
      for (let r = a!; r <= b!; r++) {
        const o = rp(r);
        o.level = Math.max(o.level ?? 0, g.level ?? 1);
        if (g.collapsed) o.hidden = true;
      }
    }

    const allRows = [...new Set([...rowsMap.keys(), ...rowProps.keys()])].sort((a, b) => a - b);
    const sheetData = allRows.map((r) => {
      const cells = (rowsMap.get(r) ?? []).sort((a, b) => a.col - b.col).map((c) => c.xml).join('');
      const p = rowProps.get(r) ?? {};
      const attrs = [`r="${r}"`,
        p.ht !== undefined ? `ht="${p.ht}" customHeight="1"` : '',
        p.hidden ? 'hidden="1"' : '',
        p.level ? `outlineLevel="${p.level}"` : ''].filter(Boolean).join(' ');
      return `<row ${attrs}>${cells}</row>`;
    }).join('');

    const colXml: string[] = [];
    for (const [k, v] of Object.entries(s.meta.cols ?? {})) {
      const [a, b] = String(k).split(':');
      const cfg: Meta = typeof v === 'number' ? { width: v } : (v ?? {});
      const attrs = [`min="${colToNum(a!)}"`, `max="${colToNum(b ?? a!)}"`];
      if (cfg.width !== undefined) attrs.push(`width="${pxToColWidth(cfg.width)}"`, 'customWidth="1"');
      if (cfg.hidden) attrs.push('hidden="1"');
      if (cfg.group) attrs.push(`outlineLevel="${cfg.group}"`);
      colXml.push(`<col ${attrs.join(' ')}/>`);
    }
    for (const g of s.outline.cols) {
      const [a, b] = String(g.range).split(':');
      colXml.push(`<col min="${colToNum(a!)}" max="${colToNum(b!)}" outlineLevel="${g.level ?? 1}"${g.collapsed ? ' hidden="1"' : ''}/>`);
    }

    let pane = '';
    if (s.meta.freeze) {
      const c = parseTarget(String(s.meta.freeze))!;
      const x = c.c1 - 1, y = c.r1 - 1;
      if (x || y) {
        pane = `<pane ${[x ? `xSplit="${x}"` : '', y ? `ySplit="${y}"` : '',
          `topLeftCell="${A1(c.c1, c.r1)}"`,
          `activePane="${x && y ? 'bottomRight' : y ? 'bottomLeft' : 'topRight'}"`,
          'state="frozen"'].filter(Boolean).join(' ')}/>`;
      }
    }
    const view = s.meta.view ?? {};
    const sheetView = `<sheetViews><sheetView workbookViewId="0"`
      + (view.gridlines === false ? ' showGridLines="0"' : '')
      + (view.headings === false ? ' showRowColHeaders="0"' : '')
      + (view.formulas ? ' showFormulas="1"' : '')
      + (view.rtl ? ' rightToLeft="1"' : '')
      + (view.zoom ? ` zoomScale="${view.zoom}"` : '')
      + `>${pane}</sheetView></sheetViews>`;

    // ---- scenarios ----
    const scenXml = s.scenarios.length
      ? `<scenarios>${s.scenarios.map((sc) => {
        const cells = Object.entries(sc.meta.cells ?? {});
        return `<scenario name="${esc(sc.name)}" locked="1" count="${cells.length}" user="GridMD"${sc.meta.comment ? ` comment="${esc(sc.meta.comment)}"` : ''}>`
          + cells.map(([r, v]) => `<inputCells r="${esc(r.replace(/\$/g, ''))}" val="${esc(v)}"/>`).join('')
          + '</scenario>';
      }).join('')}</scenarios>` : '';

    // ---- standalone AutoFilter + sortState ----
    let afXml = '', sortXml = '';
    if (s.filters.length) {
      const f = s.filters[0]!;
      if (s.filters.length > 1) report.push({ line: s.filters[1]!.line, feature: '{filter}', action: 'partial', note: 'one AutoFilter per sheet; extras skipped' });
      const t = parseTarget(f.sqref)!;
      afXml = `<autoFilter ref="${esc(f.sqref)}">${Object.entries(f.meta.cols ?? {}).map(([letter, crit]) => {
        const c = crit as Meta;
        const colId = colToNum(letter) - t.c1;
        if (c.values) return `<filterColumn colId="${colId}"><filters>${c.values.map((v: unknown) => `<filter val="${esc(v)}"/>`).join('')}</filters></filterColumn>`;
        if (c.top !== undefined) {
          const raw = String(c.top ?? c.bottom);
          return `<filterColumn colId="${colId}"><top10 val="${parseInt(raw, 10)}"${raw.endsWith('%') ? ' percent="1"' : ''}/></filterColumn>`;
        }
        if (c.op !== undefined && FILTER_OPS.has(c.op)) {
          return `<filterColumn colId="${colId}"><customFilters><customFilter operator="${FILTER_OPS.get(c.op)}" val="${esc(c.value)}"/></customFilters></filterColumn>`;
        }
        report.push({ line: f.line, feature: `{filter} column ${letter}`, action: 'partial', note: 'color/icon/text-operator criteria not serialized' });
        return '';
      }).join('')}</autoFilter>`;
      if (Array.isArray(f.meta.sort) && f.meta.sort.length) {
        const headers = f.meta.headers !== false;
        const bodyR1 = t.r1 + (headers ? 1 : 0);
        const conds = f.meta.sort.filter((lvl: Meta) => (lvl.by ?? 'value') === 'value').map((lvl: Meta) => {
          const col = numToCol(colToNum(lvl.col));
          return `<sortCondition ref="${col}${bodyR1}:${col}${t.r2}"${lvl.order === 'desc' ? ' descending="1"' : ''}/>`;
        }).join('');
        if (conds) sortXml = `<sortState ref="${A1(t.c1, bodyR1)}:${A1(t.c2, t.r2)}">${conds}</sortState>`;
      }
    }

    const mergeXml = s.merges.length
      ? `<mergeCells count="${s.merges.length}">${s.merges.map((m) => `<mergeCell ref="${A1(m.c1, m.r1)}:${A1(m.c2, m.r2)}"/>`).join('')}</mergeCells>` : '';

    let cfPriority = 0;
    const cfXml = s.cf.map((block) => {
      const rules = (Array.isArray(block.rules) ? block.rules : []).map((rule: Meta) =>
        cfRuleXml(rule, rule.priority ?? ++cfPriority, block.sqref)).filter(Boolean).join('');
      return rules ? `<conditionalFormatting sqref="${esc(block.sqref)}">${rules}</conditionalFormatting>` : '';
    }).join('');

    const dvXml = s.validations.length
      ? `<dataValidations count="${s.validations.length}">${s.validations.map(validationXml).join('')}</dataValidations>` : '';

    let linkXml = '';
    if (s.hyperlinks.length) {
      linkXml = `<hyperlinks>${s.hyperlinks.map((h) => {
        const tip = h.tip ? ` tooltip="${esc(h.tip)}"` : '';
        if (h.target.startsWith('#')) return `<hyperlink ref="${A1(h.col, h.row)}" location="${esc(h.target.slice(1))}"${tip}/>`;
        const id = nextRel();
        rels.push(`<Relationship Id="${id}" Type="${REL.hyperlink}" Target="${esc(h.target)}" TargetMode="External"/>`);
        return `<hyperlink ref="${A1(h.col, h.row)}" r:id="${id}"${tip}/>`;
      }).join('')}</hyperlinks>`;
    }

    let protXml = '';
    if (s.meta.protect?.enabled) {
      const allow = new Set<string>(s.meta.protect.allow ?? []);
      const attrs = ['sheet="1"'];
      for (const [key, attr] of PROTECT_ATTRS) attrs.push(`${attr}="${allow.has(key) ? 0 : 1}"`);
      attrs.push(`selectLockedCells="${allow.has('select-locked') ? 0 : 1}"`);
      attrs.push(`selectUnlockedCells="${allow.has('select-unlocked') ? 0 : 1}"`);
      protXml = `<sheetProtection ${attrs.join(' ')}/>`;
    }

    const pg = s.page ?? {};
    const m = pg.margins ?? {};
    const marginXml = `<pageMargins left="${cmToInch(m.left ?? 1.78)}" right="${cmToInch(m.right ?? 1.78)}" top="${cmToInch(m.top ?? 1.91)}" bottom="${cmToInch(m.bottom ?? 1.91)}" header="${cmToInch(m.header ?? 0.76)}" footer="${cmToInch(m.footer ?? 0.76)}"/>`;
    const setupAttrs: string[] = [];
    if (pg.orientation) setupAttrs.push(`orientation="${pg.orientation}"`);
    if (pg.paper) setupAttrs.push(`paperSize="${PAPER[String(pg.paper).toLowerCase()] ?? 9}"`);
    if (pg.scale !== undefined) setupAttrs.push(`scale="${pg.scale}"`);
    if (pg.fit) setupAttrs.push(`fitToWidth="${pg.fit.width ?? 1}"`, `fitToHeight="${pg.fit.height ?? 1}"`);
    const setupXml = setupAttrs.length ? `<pageSetup ${setupAttrs.join(' ')}/>` : '';
    const sheetPrXml = (s.meta['tab-color'] || pg.fit)
      ? `<sheetPr>${s.meta['tab-color'] ? `<tabColor rgb="${resolveColor(s.meta['tab-color'], themeColors)}"/>` : ''}${pg.fit ? '<pageSetUpPr fitToPage="1"/>' : ''}</sheetPr>` : '';
    const printOptsXml = (pg.gridlines || pg.headings || pg.center)
      ? `<printOptions${pg.gridlines ? ' gridLines="1"' : ''}${pg.headings ? ' headings="1"' : ''}${pg.center?.horizontal ? ' horizontalCentered="1"' : ''}${pg.center?.vertical ? ' verticalCentered="1"' : ''}/>` : '';
    const headerFooterXml = (pg.header || pg.footer)
      ? `<headerFooter>${pg.header ? `<oddHeader>${esc(hf(pg.header))}</oddHeader>` : ''}${pg.footer ? `<oddFooter>${esc(hf(pg.footer))}</oddFooter>` : ''}</headerFooter>` : '';
    const breaksXml = (pg.breaks?.rows?.length ? `<rowBreaks count="${pg.breaks.rows.length}" manualBreakCount="${pg.breaks.rows.length}">${pg.breaks.rows.map((r: unknown) => `<brk id="${r}" max="16383" man="1"/>`).join('')}</rowBreaks>` : '')
      + (pg.breaks?.cols?.length ? `<colBreaks count="${pg.breaks.cols.length}" manualBreakCount="${pg.breaks.cols.length}">${pg.breaks.cols.map((c: unknown) => `<brk id="${c}" max="1048575" man="1"/>`).join('')}</colBreaks>` : '');

    // ---- notes (legacy comments + VML) ----
    let legacyXml = '';
    if (s.notes.length) {
      anyVml = true;
      commentCount++;
      const comments = `${XMLDECL}<comments xmlns="${MAIN_NS}"><authors><author>GridMD</author></authors><commentList>`
        + s.notes.map((n) => `<comment ref="${A1(n.col, n.row)}" authorId="0"><text><r><t xml:space="preserve">${esc(n.text)}</t></r></text></comment>`).join('')
        + '</commentList></comments>';
      const vml = '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">'
        + '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'
        + '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
        + s.notes.map((n, i) => `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:6pt;width:108pt;height:44pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:AutoFill Type="False"/><x:Row>${n.row - 1}</x:Row><x:Column>${n.col - 1}</x:Column></x:ClientData></v:shape>`).join('')
        + '</xml>';
      const cid = nextRel();
      rels.push(`<Relationship Id="${cid}" Type="${REL.comments}" Target="../comments${commentCount}.xml"/>`);
      const vid = nextRel();
      rels.push(`<Relationship Id="${vid}" Type="${REL.vml}" Target="../drawings/vmlDrawing${commentCount}.vml"/>`);
      addPart(`xl/comments${commentCount}.xml`, comments, CT.comments);
      addPart(`xl/drawings/vmlDrawing${commentCount}.vml`, vml, null);
      legacyXml = `<legacyDrawing r:id="${vid}"/>`;
    }

    // ---- threaded comments ----
    if (s.threads.length) {
      tcCount++;
      const items: string[] = [];
      for (const th of s.threads) {
        (th.comments ?? []).forEach((c: Meta, i: number) => {
          if (!persons.has(c.by)) persons.set(c.by, guidFor(`person:${c.by}`));
          const rootId = guidFor(`tc:${s.name}:${th.ref}:${i}`);
          items.push(`<threadedComment ref="${esc(th.ref)}" dT="${esc(String(c.at).replace(/Z$/, ''))}" personId="${persons.get(c.by)}" id="${rootId}"${c.resolved ? ' done="1"' : ''}><text>${esc(c.text)}</text></threadedComment>`);
          (c.replies ?? []).forEach((rep: Meta, j: number) => {
            if (!persons.has(rep.by)) persons.set(rep.by, guidFor(`person:${rep.by}`));
            items.push(`<threadedComment ref="${esc(th.ref)}" dT="${esc(String(rep.at).replace(/Z$/, ''))}" personId="${persons.get(rep.by)}" id="${guidFor(`tc:${s.name}:${th.ref}:${i}:${j}`)}" parentId="${rootId}"><text>${esc(rep.text)}</text></threadedComment>`);
          });
        });
      }
      const tid = nextRel();
      rels.push(`<Relationship Id="${tid}" Type="${REL.threaded}" Target="../threadedComments/threadedComment${tcCount}.xml"/>`);
      addPart(`xl/threadedComments/threadedComment${tcCount}.xml`,
        `${XMLDECL}<ThreadedComments xmlns="${TC_NS}">${items.join('')}</ThreadedComments>`, CT.threaded);
    }

    // ---- pivots ----
    for (const p of s.pivots) {
      const built = buildPivotParts(p, {
        pivotId: pivotCount + 1, cacheId: pivotCount + 1, tableIndex, getHeaderCells,
        numFmtId: (code) => styles.numFmtId(code),
      });
      if (!built) { report.push({ line: p.line, feature: `{pivot} ${p.name}`, action: 'not-emitted', note: 'source could not be resolved to a table or headed range' }); continue; }
      pivotCount++;
      addPart(`xl/pivotCache/pivotCacheDefinition${pivotCount}.xml`, built.cacheDef, CT.pivotCacheDef);
      addPart(`xl/pivotCache/_rels/pivotCacheDefinition${pivotCount}.xml.rels`,
        relsXml([`<Relationship Id="rId1" Type="${REL.pivotCacheRec}" Target="pivotCacheRecords${pivotCount}.xml"/>`]), null);
      addPart(`xl/pivotCache/pivotCacheRecords${pivotCount}.xml`, built.cacheRecords, CT.pivotCacheRec);
      addPart(`xl/pivotTables/pivotTable${pivotCount}.xml`, built.pivotTable, CT.pivotTable);
      addPart(`xl/pivotTables/_rels/pivotTable${pivotCount}.xml.rels`,
        relsXml([`<Relationship Id="rId1" Type="${REL.pivotCacheDef}" Target="../pivotCache/pivotCacheDefinition${pivotCount}.xml"/>`]), null);
      const rid = nextRel();
      rels.push(`<Relationship Id="${rid}" Type="${REL.pivotTable}" Target="../pivotTables/pivotTable${pivotCount}.xml"/>`);
      wbPivotCaches.push({ cacheId: pivotCount, target: `pivotCache/pivotCacheDefinition${pivotCount}.xml` });
    }

    // ---- slicers ----
    const sheetSlicerRelIds: string[] = [];
    const sheetTimelineRelIds: string[] = [];
    for (const sl of s.slicers) {
      if (sl.kind === 'timeline') {
        const reg = pivotRegistry.get(String(sl.meta.for ?? '').toLowerCase());
        if (!reg) {
          carryList.push({ kind: 'slicer', line: sl.line, meta: sl.meta });
          report.push({ line: sl.line, feature: '{slicer} timeline', action: 'carried', note: 'Excel timelines require a pivot target; definition carried in-package' });
          continue;
        }
        timelineCount++;
        const field = String(sl.meta.field);
        const cacheName = `NativeTimeline_${field.replace(/[^A-Za-z0-9_]/g, '_')}`;
        const LEVELS: Record<string, number> = { years: 0, quarters: 1, months: 2, days: 3 };
        const level = LEVELS[sl.meta.level] ?? 2;
        const [d1, d2] = Array.isArray(sl.meta.range) ? sl.meta.range : [null, null];
        const dt = (d: unknown): string => `${d ?? '1900-01-01'}T00:00:00`;
        addPart(`xl/timelineCaches/timelineCache${timelineCount}.xml`,
          `${XMLDECL}<timelineCacheDefinition xmlns="${X15_NS}" xmlns:r="${R_NS}" name="${esc(cacheName)}" sourceName="${esc(field)}">`
          + `<pivotTables><pivotTable tabId="${reg.sheetIdx + 1}" name="${esc(sl.meta.for)}"/></pivotTables>`
          + `<state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="${reg.cacheId}" filterType="unknown">`
          + `<selection startDate="${dt(d1)}" endDate="${dt(d2 ?? d1)}"/><bounds startDate="${dt(d1)}" endDate="${dt(d2 ?? d1)}"/>`
          + '</state></timelineCacheDefinition>', CT.timelineCache);
        addPart(`xl/timelines/timeline${timelineCount}.xml`,
          `${XMLDECL}<timelines xmlns="${X15_NS}"><timeline name="${esc(field)}" cache="${esc(cacheName)}" caption="${esc(field)}" level="${level}" selectionLevel="${level}" scrollPosition="${dt(d1)}"/></timelines>`, CT.timeline);
        const rid = nextRel();
        rels.push(`<Relationship Id="${rid}" Type="${REL.timeline}" Target="../timelines/timeline${timelineCount}.xml"/>`);
        sheetTimelineRelIds.push(rid);
        wbTimelineCaches.push({ target: `timelineCaches/timelineCache${timelineCount}.xml` });
        sl._name = field;
        sl._timeline = true;
        continue;
      }
      const table = tableIndex.get(String(sl.meta.for ?? '').toLowerCase());
      const colIdx = table ? table.columns.findIndex((c) => c.toLowerCase() === String(sl.meta.field ?? '').toLowerCase()) : -1;
      if (!table || colIdx === -1) {
        report.push({ line: sl.line, feature: '{slicer}', action: 'not-emitted', note: 'slicer target must be a table column (pivot slicers pending)' });
        continue;
      }
      slicerCount++;
      const cacheName = `Slicer_${String(sl.meta.field).replace(/[^A-Za-z0-9_]/g, '_')}`;
      const caption = String(sl.meta.field);
      addPart(`xl/slicerCaches/slicerCache${slicerCount}.xml`,
        `${XMLDECL}<slicerCacheDefinition xmlns="${X14_NS}" xmlns:x="${MAIN_NS}" name="${esc(cacheName)}" sourceName="${esc(caption)}">`
        + `<extLst><x:ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:x15="${X15_NS}"><x15:tableSlicerCache tableId="${tableIds.get(table.name.toLowerCase())}" column="${colIdx + 1}"/></x:ext></extLst>`
        + '</slicerCacheDefinition>', CT.slicerCache);
      addPart(`xl/slicers/slicer${slicerCount}.xml`,
        `${XMLDECL}<slicers xmlns="${X14_NS}" xmlns:x="${MAIN_NS}"><slicer name="${esc(caption)}" cache="${esc(cacheName)}" caption="${esc(caption)}" rowHeight="241300"/></slicers>`, CT.slicer);
      const rid = nextRel();
      rels.push(`<Relationship Id="${rid}" Type="${REL.slicer}" Target="../slicers/slicer${slicerCount}.xml"/>`);
      sheetSlicerRelIds.push(rid);
      wbSlicerCaches.push({ target: `slicerCaches/slicerCache${slicerCount}.xml` });
      sl._name = caption;
      if (sl.meta.selected) report.push({ line: sl.line, feature: '{slicer} selected values', action: 'partial', note: 'selection state applies via the table filter; set table filter: to persist it' });
    }

    // ---- tables ----
    let tablePartsXml = '';
    if (s.tables.length) {
      const ids: string[] = [];
      for (const t of s.tables) {
        tableCount++;
        const tid = tableIds.get(t.name.toLowerCase()) ?? tableCount;
        const lastBodyRow = t.anchor.row + (t.headerRow ? 1 : 0) + t.bodyRows - 1;
        const lastRow = t.total ? lastBodyRow + 1 : lastBodyRow;
        const bodyR1 = t.anchor.row + (t.headerRow ? 1 : 0);
        const refAll = `${A1(t.anchor.col, t.anchor.row)}:${A1(t.anchor.col + t.columns.length - 1, lastRow)}`;
        const refFilter = `${A1(t.anchor.col, t.anchor.row)}:${A1(t.anchor.col + t.columns.length - 1, lastBodyRow)}`;
        const filterCols = Object.entries(t.filter ?? {}).map(([colName, crit]) => {
          const c = crit as Meta;
          const ci = t.columns.findIndex((col) => col.toLowerCase() === colName.toLowerCase());
          if (ci === -1 || !c?.values) return '';
          return `<filterColumn colId="${ci}"><filters>${c.values.map((v: unknown) => `<filter val="${esc(v)}"/>`).join('')}</filters></filterColumn>`;
        }).join('');
        const sortState = (t.sort ?? []).length
          ? `<sortState ref="${A1(t.anchor.col, bodyR1)}:${A1(t.anchor.col + t.columns.length - 1, lastBodyRow)}">`
          + t.sort.map((lvl: Meta) => {
            const ci = t.columns.findIndex((c) => c.toLowerCase() === String(lvl.col).toLowerCase());
            if (ci === -1) return '';
            const col = numToCol(t.anchor.col + ci);
            return `<sortCondition ref="${col}${bodyR1}:${col}${lastBodyRow}"${lvl.order === 'desc' ? ' descending="1"' : ''}/>`;
          }).join('') + '</sortState>' : '';
        const colsXml = t.columns.map((c, i) => {
          let totals = '';
          const tv = Object.entries(t.total ?? {}).find(([k]) => k.toLowerCase() === c.toLowerCase())?.[1];
          if (tv !== undefined) {
            const fn = /^=SUBTOTAL\((\d{3}),/.exec(String(tv));
            if (fn && SUBTOTAL_FN.has(Number(fn[1]))) totals = ` totalsRowFunction="${SUBTOTAL_FN.get(Number(fn[1]))}"`;
            else if (!String(tv).startsWith('=')) totals = ` totalsRowLabel="${esc(tv)}"`;
            else totals = ' totalsRowFunction="custom"';
          }
          return `<tableColumn id="${i + 1}" name="${esc(c)}"${totals}/>`;
        }).join('');
        const styleName = t.style ? tableStyleName(t.style) : null;
        const tableXml = `${XMLDECL}<table xmlns="${MAIN_NS}" id="${tid}" name="${esc(t.name)}" displayName="${esc(t.name)}" ref="${refAll}"${t.total ? ' totalsRowCount="1"' : ''}>`
          + `<autoFilter ref="${refFilter}">${filterCols}</autoFilter>`
          + sortState
          + `<tableColumns count="${t.columns.length}">${colsXml}</tableColumns>`
          + (styleName ? `<tableStyleInfo name="${styleName}" showFirstColumn="0" showLastColumn="0" showRowStripes="${t.banded === 'rows' || t.banded === 'both' ? 1 : 0}" showColumnStripes="${t.banded === 'cols' || t.banded === 'both' ? 1 : 0}"/>` : '')
          + '</table>';
        addPart(`xl/tables/table${tableCount}.xml`, tableXml, CT.table);
        const rid = nextRel();
        rels.push(`<Relationship Id="${rid}" Type="${REL.table}" Target="../tables/table${tableCount}.xml"/>`);
        ids.push(rid);
      }
      tablePartsXml = `<tableParts count="${ids.length}">${ids.map((id) => `<tablePart r:id="${id}"/>`).join('')}</tableParts>`;
    }

    // ---- drawing (charts + images + shapes + slicer frames) ----
    let drawingRef = '';
    const anchors: string[] = [];
    const drawingRels: string[] = [];
    let dRelId = 0;
    const nextDRel = (): string => `rId${++dRelId}`;
    let objId = 1;
    for (const chart of s.charts) {
      const em = emitChart(chart, s.name);
      if (!em) continue;
      const rid = nextDRel();
      drawingRels.push(`<Relationship Id="${rid}" Type="${em.ex ? REL.chartEx : REL.chart}" Target="../charts/${em.part.split('/').pop()}"/>`);
      const frame = em.ex ? chartExFrame : chartFrame;
      anchors.push(frame(++objId, chart.title ?? `Chart ${objId}`, rid, parseAnchor(chart.anchor, chart.size)));
    }
    for (const img of s.images) {
      const media = loadImage(img);
      if (!media) continue;
      mediaCount++;
      const name = `xl/media/image${mediaCount}.${media.ext}`;
      addPart(name, media.data, null);
      mediaDefaults.add(JSON.stringify([media.ext, media.ct]));
      const rid = nextDRel();
      drawingRels.push(`<Relationship Id="${rid}" Type="${REL.image}" Target="../media/image${mediaCount}.${media.ext}"/>`);
      anchors.push(pictureFrame(++objId, `Picture ${objId}`, rid, parseAnchor(img.anchor, img.size), img.alt));
    }
    for (const shape of s.shapes) {
      anchors.push(shapeFrame(++objId, shape, parseAnchor(shape.anchor, shape.size), themeColors));
    }
    for (const sl of s.slicers) {
      if (!sl._name) continue;
      const frame = sl._timeline ? timesliceFrame : slicerFrame;
      anchors.push(frame(++objId, sl._name, parseAnchor(sl.anchor, sl.size ?? (sl._timeline ? { w: 320, h: 110 } : { w: 160, h: 200 }))));
    }
    if (anchors.length) {
      drawingCount++;
      addPart(`xl/drawings/drawing${drawingCount}.xml`, drawingXml(anchors), CT.drawing);
      if (drawingRels.length) addPart(`xl/drawings/_rels/drawing${drawingCount}.xml.rels`, relsXml(drawingRels), null);
      const rid = nextRel();
      rels.push(`<Relationship Id="${rid}" Type="${REL.drawing}" Target="../drawings/drawing${drawingCount}.xml"/>`);
      drawingRef = `<drawing r:id="${rid}"/>`;
    }

    // ---- extLst: sparklines + slicer list ----
    const exts: string[] = [];
    if (s.sparklines.length) {
      const groups = s.sparklines.map((g) => sparklineGroupXml(g, s.name)).filter(Boolean).join('');
      if (groups) {
        exts.push(`<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14_NS}">`
          + `<x14:sparklineGroups xmlns:xm="${XM_NS}">${groups}</x14:sparklineGroups></ext>`);
      }
    }
    if (sheetSlicerRelIds.length) {
      exts.push(`<ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="${X14_NS}">`
        + `<x14:slicerList>${sheetSlicerRelIds.map((rid) => `<x14:slicer xmlns:r="${R_NS}" r:id="${rid}"/>`).join('')}</x14:slicerList></ext>`);
    }
    if (sheetTimelineRelIds.length) {
      exts.push(`<ext uri="{7E03D99C-DC04-49d9-9315-930204A7B6E9}" xmlns:x15="${X15_NS}">`
        + `<x15:timelineRefs>${sheetTimelineRelIds.map((rid) => `<x15:timelineRef xmlns:r="${R_NS}" r:id="${rid}"/>`).join('')}</x15:timelineRefs></ext>`);
    }
    const extXml = exts.length ? `<extLst>${exts.join('')}</extLst>` : '';

    const xml = XMLDECL
      + `<worksheet xmlns="${MAIN_NS}" xmlns:r="${R_NS}">`
      + sheetPrXml
      + `<dimension ref="${dimension}"/>`
      + sheetView
      + '<sheetFormatPr defaultRowHeight="15"/>'
      + (colXml.length ? `<cols>${colXml.join('')}</cols>` : '')
      + `<sheetData>${sheetData}</sheetData>`
      + protXml
      + scenXml
      + afXml
      + sortXml
      + mergeXml
      + cfXml
      + dvXml
      + linkXml
      + printOptsXml
      + marginXml
      + setupXml
      + headerFooterXml
      + breaksXml
      + drawingRef
      + legacyXml
      + tablePartsXml
      + extXml
      + '</worksheet>';

    addPart(partName, xml, CT.worksheet);
    if (rels.length) addPart(`xl/worksheets/_rels/sheet${wsCount}.xml.rels`, relsXml(rels), null);
    return { target: `worksheets/sheet${wsCount}.xml` };
  }

  // ================= chart sheet =================
  function renderChartSheet(s: Sheet): { target: string } {
    csCount++;
    const partName = `xl/chartsheets/sheet${csCount}.xml`;
    const rels: string[] = [];
    const chart = s.charts[0];
    let drawingEl = '';
    if (chart) {
      const em = emitChart(chart, s.name);
      if (em) {
        drawingCount++;
        const rid = 'rId1';
        const frameFn = em.ex ? chartExFrame : chartFrame;
        const frame = frameFn(2, chart.title ?? s.name, rid,
          { kind: 'absolute', x: 0, y: 0, cx: 9144000, cy: 6858000 });
        addPart(`xl/drawings/drawing${drawingCount}.xml`, drawingXml([frame]), CT.drawing);
        addPart(`xl/drawings/_rels/drawing${drawingCount}.xml.rels`,
          relsXml([`<Relationship Id="${rid}" Type="${em.ex ? REL.chartEx : REL.chart}" Target="../charts/${em.part.split('/').pop()}"/>`]), null);
        rels.push(`<Relationship Id="rId1" Type="${REL.drawing}" Target="../drawings/drawing${drawingCount}.xml"/>`);
        drawingEl = '<drawing r:id="rId1"/>';
      }
    }
    const tab = s.meta['tab-color'] ? `<sheetPr><tabColor rgb="${resolveColor(s.meta['tab-color'], themeColors)}"/></sheetPr>` : '<sheetPr/>';
    const xml = `${XMLDECL}<chartsheet xmlns="${MAIN_NS}" xmlns:r="${R_NS}">${tab}`
      + '<sheetViews><sheetView workbookViewId="0" zoomToFit="1"/></sheetViews>'
      + `${drawingEl}</chartsheet>`;
    addPart(partName, xml, CT.chartsheet);
    if (rels.length) addPart(`xl/chartsheets/_rels/sheet${csCount}.xml.rels`, relsXml(rels), null);
    return { target: `chartsheets/sheet${csCount}.xml` };
  }

  // ================= helpers =================
  function loadImage(img: ImageModel): ImageMedia | null {
    const dm = /^data:image\/(png|jpeg|jpg|gif);base64,(.+)$/i.exec(img.src);
    if (dm) {
      const ext = dm[1]!.toLowerCase() === 'jpg' ? 'jpeg' : dm[1]!.toLowerCase();
      return { ext, ct: `image/${ext}`, data: Buffer.from(dm[2]!, 'base64') };
    }
    if (/^[a-z][a-z0-9+.-]*:/i.test(img.src)) {
      report.push({ line: img.line, feature: `{image} ${img.src.slice(0, 40)}`, action: 'not-emitted', note: 'remote images are not fetched at convert time' });
      return null;
    }
    try {
      const data = readFileSync(join(baseDir, img.src));
      const ext = data.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) ? 'png'
        : data[0] === 0xff && data[1] === 0xd8 ? 'jpeg' : null;
      if (!ext) { report.push({ line: img.line, feature: `{image} ${img.src}`, action: 'not-emitted', note: 'unsupported image format (png/jpeg only)' }); return null; }
      return { ext, ct: `image/${ext}`, data };
    } catch {
      report.push({ line: img.line, feature: `{image} ${img.src}`, action: 'not-emitted', note: 'file not found relative to the document' });
      return null;
    }
  }

  function sparklineGroupXml(g: SparklineModel, ownSheet: string): string {
    const meta = g.meta ?? {};
    const targets = parseTarget(g.sqref);
    const source = parseTarget(String(meta.source));
    if (!targets || !source) return '';
    const tCells: { c: number; r: number }[] = [];
    for (let r = targets.r1; r <= targets.r2; r++) for (let c = targets.c1; c <= targets.c2; c++) tCells.push({ c, r });
    const srcRows = source.r2 - source.r1 + 1;
    const srcCols = source.c2 - source.c1 + 1;
    const bySheet = source.sheet ?? ownSheet;
    const slices: string[] = [];
    if (srcRows === tCells.length) {
      for (let i = 0; i < tCells.length; i++) slices.push(`${quoteSheet(bySheet)}!${A1(source.c1, source.r1 + i)}:${A1(source.c2, source.r1 + i)}`);
    } else if (srcCols === tCells.length) {
      for (let i = 0; i < tCells.length; i++) slices.push(`${quoteSheet(bySheet)}!${A1(source.c1 + i, source.r1)}:${A1(source.c1 + i, source.r2)}`);
    } else {
      report.push({ line: g.line, feature: '{sparklines}', action: 'not-emitted', note: 'source rows/cols must match the target cell count' });
      return '';
    }
    const mk = meta.markers ?? {};
    const color = resolveColor(meta.color ?? 'accent1', themeColors) ?? 'FF376092';
    const typeAttr = meta.type === 'column' ? ' type="column"' : meta.type === 'win-loss' ? ' type="stacked"' : '';
    const markAttrs = `${mk === true || mk.high ? ' high="1"' : ''}${mk === true || mk.low ? ' low="1"' : ''}${mk.first ? ' first="1"' : ''}${mk.last ? ' last="1"' : ''}${mk.negative ? ' negative="1"' : ''}${(mk === true || mk.high || mk.low || mk.first || mk.last) ? ' markers="1"' : ''}`;
    const axis = meta.axis?.show ? ' displayXAxis="1"' : '';
    return `<x14:sparklineGroup displayEmptyCellsAs="gap"${typeAttr}${markAttrs}${axis}>`
      + `<x14:colorSeries rgb="${color}"/><x14:colorNegative rgb="FFD00000"/><x14:colorAxis rgb="FF000000"/>`
      + `<x14:colorMarkers rgb="${color}"/><x14:colorFirst rgb="${color}"/><x14:colorLast rgb="${color}"/>`
      + `<x14:colorHigh rgb="FF00B050"/><x14:colorLow rgb="FFD00000"/>`
      + `<x14:sparklines>${tCells.map((t, i) => `<x14:sparkline><xm:f>${esc(slices[i])}</xm:f><xm:sqref>${A1(t.c, t.r)}</xm:sqref></x14:sparkline>`).join('')}</x14:sparklines>`
      + '</x14:sparklineGroup>';
  }

  function cellXml(cell: Cell, content: CellContent | null, xf: number): string | null {
    const ref = A1(cell.col, cell.row);
    const sAttr = xf ? ` s="${xf}"` : '';
    if (!content) return xf ? `<c r="${ref}"${sAttr}/>` : null;
    if (content.rich) {
      const runs = content.rich.map((r: Meta) => {
        const pr: string[] = [];
        if (r.bold) pr.push('<b/>');
        if (r.italic) pr.push('<i/>');
        if (r.size) pr.push(`<sz val="${r.size}"/>`);
        const c = resolveColor(r.color ?? '', themeColors);
        if (c) pr.push(`<color rgb="${c}"/>`);
        return `<r>${pr.length ? `<rPr>${pr.join('')}</rPr>` : ''}<t xml:space="preserve">${esc(r.text ?? '')}</t></r>`;
      }).join('');
      return `<c r="${ref}"${sAttr} t="inlineStr"><is>${runs}</is></c>`;
    }
    if (content.formula !== undefined) {
      const fAttr = (content.cse || content.arrayRef) ? ` t="array" ref="${content.arrayRef ?? `${ref}:${ref}`}"` : '';
      const f = `<f${fAttr}>${esc(content.formula)}</f>`;
      const cached = content.cached;
      if (!cached || cached.kind === 'blank') return `<c r="${ref}"${sAttr}>${f}</c>`;
      if (cached.kind === 'number') return `<c r="${ref}"${sAttr}>${f}<v>${cached.value}</v></c>`;
      if (cached.kind === 'boolean') return `<c r="${ref}"${sAttr} t="b">${f}<v>${cached.value ? 1 : 0}</v></c>`;
      if (cached.kind === 'error') return `<c r="${ref}"${sAttr} t="e">${f}<v>${cached.value}</v></c>`;
      if (cached.kind === 'date' || cached.kind === 'time') return `<c r="${ref}"${sAttr}>${f}<v>${isoToSerial(String(cached.value), dateSystem)}</v></c>`;
      return `<c r="${ref}"${sAttr} t="str">${f}<v>${esc(cached.value)}</v></c>`;
    }
    const sc = content.scalar;
    if (!sc) return xf ? `<c r="${ref}"${sAttr}/>` : null;
    switch (sc.kind) {
      case 'number': return `<c r="${ref}"${sAttr}><v>${sc.value}</v></c>`;
      case 'boolean': return `<c r="${ref}"${sAttr} t="b"><v>${sc.value ? 1 : 0}</v></c>`;
      case 'error': return `<c r="${ref}"${sAttr} t="e"><v>${sc.value}</v></c>`;
      case 'date':
      case 'time': return `<c r="${ref}"${sAttr}><v>${isoToSerial(String(sc.value), dateSystem)}</v></c>`;
      default: return `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${esc(sc.value)}</t></is></c>`;
    }
  }

  function collectEdges(props: Meta): BorderEdges | null {
    const single = props.border !== undefined && props['border-top'] === undefined
      && props['border-bottom'] === undefined && props['border-left'] === undefined
      && props['border-right'] === undefined ? parseBorderEdge(props.border, themeColors) : null;
    const edge = (k: string) => props[k] !== undefined ? parseBorderEdge(props[k], themeColors) : single;
    const edges: BorderEdges = {
      top: edge('border-top'), bottom: edge('border-bottom'),
      left: edge('border-left'), right: edge('border-right'),
      diagUp: props['border-diag-up'] ? parseBorderEdge(props['border-diag-up'], themeColors) : null,
      diagDown: props['border-diag-down'] ? parseBorderEdge(props['border-diag-down'], themeColors) : null,
    };
    return Object.values(edges).some(Boolean) ? edges : null;
  }

  function cfRuleXml(rule: Meta, priority: number, sqref: string): string {
    const first = sqref.split(':')[0]!;
    const dxf = (rule.format || rule.style) ? ` dxfId="${styles.dxfId(rule.format ?? {})}"` : '';
    const stop = rule.stop ? ' stopIfTrue="1"' : '';
    const base = `priority="${priority}"${dxf}${stop}`;
    const lit = (v: unknown): string => {
      const sv = String(v);
      if (sv.startsWith('=')) return esc(sv.slice(1));
      if (/^-?\d+(\.\d+)?$/.test(sv)) return sv;
      return esc(`"${sv.replace(/"/g, '""')}"`);
    };
    if (rule.when !== undefined) {
      const bm = /^(not-)?between\s+(.+?)\s+and\s+(.+)$/i.exec(String(rule.when).trim());
      if (bm) return `<cfRule type="cellIs" ${base} operator="${bm[1] ? 'notBetween' : 'between'}"><formula>${lit(bm[2])}</formula><formula>${lit(bm[3])}</formula></cfRule>`;
      const om = /^(<>|>=|<=|=|>|<)\s*(.+)$/.exec(String(rule.when).trim());
      if (!om) return '';
      return `<cfRule type="cellIs" ${base} operator="${CELLIS_OPS.get(om[1]!)}"><formula>${lit(om[2])}</formula></cfRule>`;
    }
    if (rule.contains !== undefined) {
      const t = esc(rule.contains);
      return `<cfRule type="containsText" ${base} operator="containsText" text="${t}"><formula>NOT(ISERROR(SEARCH("${t}",${first})))</formula></cfRule>`;
    }
    if (rule['not-contains'] !== undefined) {
      const t = esc(rule['not-contains']);
      return `<cfRule type="notContainsText" ${base} operator="notContains" text="${t}"><formula>ISERROR(SEARCH("${t}",${first}))</formula></cfRule>`;
    }
    if (rule.begins !== undefined) {
      const t = esc(rule.begins);
      return `<cfRule type="beginsWith" ${base} operator="beginsWith" text="${t}"><formula>LEFT(${first},LEN("${t}"))="${t}"</formula></cfRule>`;
    }
    if (rule.ends !== undefined) {
      const t = esc(rule.ends);
      return `<cfRule type="endsWith" ${base} operator="endsWith" text="${t}"><formula>RIGHT(${first},LEN("${t}"))="${t}"</formula></cfRule>`;
    }
    if (rule.date !== undefined) {
      const tp = TIME_PERIODS.get(rule.date);
      return tp ? `<cfRule type="timePeriod" ${base} timePeriod="${tp}"/>` : '';
    }
    if (rule.dupes !== undefined) return `<cfRule type="duplicateValues" ${base}/>`;
    if (rule.unique !== undefined) return `<cfRule type="uniqueValues" ${base}/>`;
    if (rule.top !== undefined || rule.bottom !== undefined) {
      const raw = String(rule.top ?? rule.bottom);
      return `<cfRule type="top10" ${base} rank="${parseInt(raw, 10)}"${raw.endsWith('%') ? ' percent="1"' : ''}${rule.bottom !== undefined ? ' bottom="1"' : ''}/>`;
    }
    if (rule.avg !== undefined) {
      const below = String(rule.avg).startsWith('below');
      const equal = String(rule.avg).endsWith('-equal');
      return `<cfRule type="aboveAverage" ${base}${below ? ' aboveAverage="0"' : ''}${equal ? ' equalAverage="1"' : ''}${rule.stddev ? ` stdDev="${rule.stddev}"` : ''}/>`;
    }
    if (rule.bars !== undefined) {
      const b = rule.bars === true ? {} : rule.bars;
      const cfvo = (side: string, def: string): string => {
        const cv = b[side];
        if (!cv || cv.type === 'auto') return `<cfvo type="${def}"/>`;
        return `<cfvo type="${cv.type}"${cv.value !== undefined ? ` val="${cv.value}"` : ''}/>`;
      };
      const color = resolveColor(b.color ?? '#638EC6', themeColors);
      return `<cfRule type="dataBar" ${base.replace(/ dxfId="\d+"/, '')}><dataBar>${cfvo('min', 'min')}${cfvo('max', 'max')}<color rgb="${color}"/></dataBar></cfRule>`;
    }
    if (rule.scale !== undefined) {
      const colors = rule.scale.map((c: unknown) => `<color rgb="${resolveColor(c, themeColors)}"/>`).join('');
      const cfvos = rule.scale.length === 3
        ? '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>'
        : '<cfvo type="min"/><cfvo type="max"/>';
      return `<cfRule type="colorScale" ${base.replace(/ dxfId="\d+"/, '')}><colorScale>${cfvos}${colors}</colorScale></cfRule>`;
    }
    if (rule.icons !== undefined) {
      const set = ICON_SETS.get(rule.icons) ?? '3Arrows';
      const n = parseInt(set, 10) || 3;
      const steps = Array.from({ length: n }, (_, i) => `<cfvo type="percent" val="${Math.round((100 / n) * i)}"/>`).join('');
      return `<cfRule type="iconSet" ${base.replace(/ dxfId="\d+"/, '')}><iconSet iconSet="${set}"${rule.reverse ? ' reverse="1"' : ''}${rule['icons-only'] ? ' showValue="0"' : ''}>${steps}</iconSet></cfRule>`;
    }
    if (rule.formula !== undefined) {
      return `<cfRule type="expression" ${base}><formula>${esc(String(rule.formula).replace(/^=/, ''))}</formula></cfRule>`;
    }
    return '';
  }

  function validationXml(v: ValidationBlock): string {
    const meta = v.meta;
    const TYPE: Record<string, string> = { list: 'list', whole: 'whole', decimal: 'decimal', date: 'date', time: 'time', 'text-length': 'textLength', custom: 'custom' };
    const OPS: Record<string, string> = { between: 'between', 'not-between': 'notBetween', '=': 'equal', '<>': 'notEqual', '>': 'greaterThan', '>=': 'greaterThanOrEqual', '<': 'lessThan', '<=': 'lessThanOrEqual' };
    const attrs = [`type="${TYPE[meta.type]}"`, `sqref="${esc(v.sqref)}"`];
    if (meta.blank !== false) attrs.push('allowBlank="1"');
    if (meta.type === 'list' && meta.dropdown === false) attrs.push('showDropDown="1"');
    if (meta.error?.style && meta.error.style !== 'stop') attrs.push(`errorStyle="${meta.error.style}"`);
    if (meta.input) attrs.push('showInputMessage="1"', meta.input.title ? `promptTitle="${esc(meta.input.title)}"` : '', meta.input.message ? `prompt="${esc(meta.input.message)}"` : '');
    if (meta.error) attrs.push('showErrorMessage="1"', meta.error.title ? `errorTitle="${esc(meta.error.title)}"` : '', meta.error.message ? `error="${esc(meta.error.message)}"` : '');
    let op = meta.op;
    if (!op && meta.min !== undefined && meta.max !== undefined) op = 'between';
    if (op && OPS[op]) attrs.push(`operator="${OPS[op]}"`);
    let f1: string | null = null, f2: string | null = null;
    if (meta.type === 'list') f1 = meta.values ? `"${meta.values.join(',')}"` : String(meta.source).replace(/^=/, '');
    else if (meta.type === 'custom') f1 = String(meta.formula ?? '').replace(/^=/, '');
    else if (op === 'between' || op === 'not-between') { f1 = String(meta.min); f2 = String(meta.max); }
    else if (meta.value !== undefined) f1 = String(meta.value).replace(/^=/, '');
    return `<dataValidation ${attrs.filter(Boolean).join(' ')}>${f1 !== null ? `<formula1>${esc(f1)}</formula1>` : ''}${f2 !== null ? `<formula2>${esc(f2)}</formula2>` : ''}</dataValidation>`;
  }
}

function tableStyleName(style: unknown): string | null {
  const m = /^(light|medium|dark)-(\d+)$/.exec(String(style));
  if (!m) return null;
  return `TableStyle${m[1]![0]!.toUpperCase()}${m[1]!.slice(1)}${m[2]}`;
}

function hf(spec: Meta): string {
  if (typeof spec === 'string') return spec;
  return `${spec.left ? `&L${spec.left}` : ''}${spec.center ? `&C${spec.center}` : ''}${spec.right ? `&R${spec.right}` : ''}`;
}
