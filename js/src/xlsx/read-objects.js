// Reverse-parsing of drawing-hosted objects and analysis parts for the
// importer: charts (chartML → {chart}), pictures (→ {image} data: URIs),
// shapes/textboxes (→ {shape}/{textbox}), pivot tables (→ {pivot}) and
// table slicers (→ {slicer}). Anything unrecognized returns null so the
// caller falls back to carrying the raw part.

import { parseCell, numToCol } from '../refs.js';
import { one, all, textOf, attr } from '../xml.js';

const A1 = (col, row) => `${numToCol(col)}${row}`;

const LEGEND_REV = new Map(Object.entries({ l: 'left', r: 'right', t: 'top', b: 'bottom' }));
const DLBL_POS_REV = new Map(Object.entries({ ctr: 'center', inEnd: 'inside-end', inBase: 'inside-base', outEnd: 'outside-end', t: 'above' }));
const TREND_REV = new Map(Object.entries({ linear: 'linear', poly: 'poly', exp: 'exp', log: 'log', power: 'power', movingAvg: 'moving-average' }));
const ERR_TYPE_REV = new Map(Object.entries({ stdErr: 'std-error', percentage: 'percentage', stdDev: 'std-dev', fixedVal: 'fixed' }));
const SHOW_AS_REV = new Map(Object.entries({
  percentOfTotal: 'percent-of-total', percentOfRow: 'percent-of-row', percentOfCol: 'percent-of-column',
  percentOf: 'percent-of', runTotal: 'running-total', difference: 'difference-from',
  percentDiff: 'percent-difference-from', index: 'index',
}));
const AGG_REV = new Map(Object.entries({
  average: 'average', count: 'count', countNums: 'count-numbers', max: 'max', min: 'min',
  product: 'product', stdDev: 'std-dev', stdDevP: 'std-devp', var: 'var', varP: 'varp',
}));
const PRESET_REV = new Map(Object.entries({
  rect: 'rect', roundRect: 'rounded-rect', ellipse: 'ellipse', triangle: 'triangle',
  rtTriangle: 'right-triangle', diamond: 'diamond', pentagon: 'pentagon', hexagon: 'hexagon',
  star5: 'star', rightArrow: 'arrow-right', leftArrow: 'arrow-left', upArrow: 'arrow-up',
  downArrow: 'arrow-down', chevron: 'chevron', wedgeRectCallout: 'callout', line: 'line',
  bentConnector3: 'connector',
}));

const rgbOf = (el) => {
  const clr = el && (one(el, 'srgbClr') ?? null);
  return clr?.attrs.val ? `#${clr.attrs.val.toUpperCase()}` : null;
};
const solidFillColor = (spPr) => {
  const sf = spPr && one(spPr, 'solidFill');
  return sf ? rgbOf(sf) : null;
};
const yaml = (s) => {
  const str = String(s);
  if (/^[A-Za-z0-9][A-Za-z0-9 _./()+-]*$/.test(str) && !/^(true|false|null)$/i.test(str) && !/^-?\d/.test(str)) return str;
  return `'${str.replace(/'/g, "''")}'`;
};

// ---- drawing anchors ----
export function anchorsOf(drawingDoc) {
  const out = [];
  for (const kind of ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor']) {
    for (const a of all(drawingDoc, kind)) out.push({ kind, el: a });
  }
  return out;
}

export function anchorText(anchor) {
  if (anchor.kind === 'absoluteAnchor') return { at: 'sheet' };
  const from = one(anchor.el, 'from');
  const cell = { col: Number(textOf(one(from, 'col'))) + 1, row: Number(textOf(one(from, 'row'))) + 1 };
  if (anchor.kind === 'twoCellAnchor') {
    const to = one(anchor.el, 'to');
    return { at: `${A1(cell.col, cell.row)}:${A1(Number(textOf(one(to, 'col'))), Number(textOf(one(to, 'row'))))}` };
  }
  const ext = one(anchor.el, 'ext');
  return {
    at: A1(cell.col, cell.row),
    size: { w: Math.round(Number(ext?.attrs.cx ?? 0) / 9525), h: Math.round(Number(ext?.attrs.cy ?? 0) / 9525) },
  };
}

const fenceAnchor = (a) => a.at === 'sheet' ? 'at sheet' : `at ${a.at}${a.size ? ` size ${a.size.w}x${a.size.h}` : ''}`;

// ---- chart reversal ----
const BLOCK_KINDS = new Map(Object.entries({
  barChart: 'bar', lineChart: 'line', areaChart: 'area', pieChart: 'pie',
  doughnutChart: 'doughnut', scatterChart: 'scatter', radarChart: 'radar',
  bubbleChart: 'bubble', stockChart: 'stock',
  bar3DChart: 'bar', line3DChart: 'line', area3DChart: 'area', pie3DChart: 'pie',
}));
const BLOCK_3D = new Set(['bar3DChart', 'line3DChart', 'area3DChart', 'pie3DChart']);
const CX_TYPE_REV = new Map(Object.entries({
  treemap: 'treemap', sunburst: 'sunburst', waterfall: 'waterfall',
  funnel: 'funnel', boxWhisker: 'box-whisker', regionMap: 'map',
}));

export function reverseChart(chartDoc, anchorSpec) {
  if (one(chartDoc, 'chartData')) return reverseChartEx(chartDoc, anchorSpec);
  const chart = one(chartDoc, 'chart');
  const plot = chart && one(chart, 'plotArea');
  if (!plot) return null;
  const blocks = [];
  for (const child of plot.children) {
    if (!BLOCK_KINDS.has(child.name)) continue;
    blocks.push({ el: child, kind: BLOCK_KINDS.get(child.name), threeD: BLOCK_3D.has(child.name) });
  }
  if (!blocks.length) return null; // unknown → carry raw

  const valAxes = all(plot, 'valAx');
  const secondAxIds = valAxes.length > 1 ? new Set([attrVal(valAxes[1], 'axId')]) : new Set();

  const series = [];
  for (const blk of blocks) {
    const grouping = one(blk.el, 'grouping')?.attrs.val;
    const suffix = grouping === 'stacked' ? '-stacked' : grouping === 'percentStacked' ? '-stacked100' : '';
    const barDir = one(blk.el, 'barDir')?.attrs.val;
    const kind = (blk.kind === 'bar' ? (barDir === 'bar' ? 'bar' : 'column') + suffix : blk.kind + suffix)
      + (blk.threeD ? '-3d' : '');
    const axIds = all(blk.el, 'axId').map((x) => x.attrs.val);
    const secondary = axIds.some((id) => secondAxIds.has(id));
    const gap = one(blk.el, 'gapWidth')?.attrs.val;
    const overlap = one(blk.el, 'overlap')?.attrs.val;
    for (const ser of all(blk.el, 'ser')) {
      series.push({ ser, kind, secondary, gap, overlap });
    }
  }
  if (!series.length) return null;

  const typeCounts = new Set(series.map((s) => s.kind));
  const chartType = typeCounts.size > 1 || series.some((s) => s.secondary) ? 'combo' : series[0].kind;
  const title = chart && one(chart, 'title') ? textOf(one(chart, 'title')).trim() : null;

  const lines = [`\`\`\`{chart} ${chartType}${title ? ` "${title.replace(/"/g, '""')}"` : ''} ${fenceAnchor(anchorSpec)}`, 'series:'];
  for (const { ser, kind, secondary, gap, overlap } of series) {
    const tx = one(ser, 'tx');
    const name = tx ? (one(tx, 'strRef') ? null : textOf(tx).trim()) : null;
    const nameRef = tx && one(tx, 'strRef') ? textOf(one(one(tx, 'strRef'), 'f')).trim() : null;
    if (name) lines.push(`  - name: ${yaml(name)}`);
    else if (nameRef) lines.push(`  - name-ref: =${nameRef}`);
    else lines.push('  - name: Series');
    const cat = one(ser, 'cat') ?? one(ser, 'xVal');
    const val = one(ser, 'val') ?? one(ser, 'yVal');
    const refOf = (el) => el && textOf(one(one(el, 'strRef') ?? one(el, 'numRef') ?? el, 'f')).trim();
    const catRef = refOf(cat);
    const valRef = refOf(val);
    if (catRef) lines.push(`    cat: ${yaml(catRef)}`);
    if (valRef) lines.push(`    val: ${yaml(valRef)}`);
    if (chartType === 'combo') lines.push(`    kind: ${kind}`);
    if (secondary) lines.push('    axis: y2');
    const color = solidFillColor(one(ser, 'spPr')) ?? lineColor(one(ser, 'spPr'));
    if (color) lines.push(`    color: ${yaml(color)}`);
    const marker = one(ser, 'marker') && one(one(ser, 'marker'), 'symbol')?.attrs.val;
    if (marker && marker !== 'none') lines.push(`    marker: ${marker}`);
    if (one(ser, 'smooth')?.attrs.val === '1') lines.push('    smooth: true');
    if (gap !== undefined && gap !== '150') lines.push(`    gap: ${gap}`);
    if (overlap !== undefined && overlap !== '100' && overlap !== '0') lines.push(`    overlap: ${overlap}`);
    const dLbls = one(ser, 'dLbls');
    if (dLbls && one(dLbls, 'showVal')?.attrs.val === '1') {
      const bits = ['show: true'];
      const pos = one(dLbls, 'dLblPos')?.attrs.val;
      if (pos && DLBL_POS_REV.has(pos)) bits.push(`position: ${DLBL_POS_REV.get(pos)}`);
      const nf = one(dLbls, 'numFmt')?.attrs.formatCode;
      if (nf) bits.push(`numfmt: ${yaml(nf)}`);
      lines.push(`    labels: { ${bits.join(', ')} }`);
    }
    const tr = one(ser, 'trendline');
    if (tr) {
      const bits = [`type: ${TREND_REV.get(one(tr, 'trendlineType')?.attrs.val) ?? 'linear'}`];
      const order = one(tr, 'order')?.attrs.val;
      if (order) bits.push(`order: ${order}`);
      const period = one(tr, 'period')?.attrs.val;
      if (period) bits.push(`window: ${period}`);
      const fwd = one(tr, 'forward')?.attrs.val;
      const back = one(tr, 'backward')?.attrs.val;
      if (Number(fwd) || Number(back)) bits.push(`forecast: { forward: ${fwd ?? 0}, backward: ${back ?? 0} }`);
      const icpt = one(tr, 'intercept')?.attrs.val;
      if (icpt !== undefined) bits.push(`intercept: ${icpt}`);
      if (one(tr, 'dispEq')?.attrs.val === '1') bits.push('equation: true');
      if (one(tr, 'dispRSqr')?.attrs.val === '1') bits.push('r2: true');
      lines.push(`    trendline: { ${bits.join(', ')} }`);
    }
    const eb = one(ser, 'errBars');
    if (eb) {
      const bits = [`dir: ${one(eb, 'errBarType')?.attrs.val ?? 'both'}`,
        `type: ${ERR_TYPE_REV.get(one(eb, 'errValType')?.attrs.val) ?? 'std-error'}`];
      const val = one(eb, 'val')?.attrs.val;
      if (val !== undefined) bits.push(`value: ${val}`);
      if (one(eb, 'noEndCap')?.attrs.val !== '1') bits.push('cap: true');
      lines.push(`    error-bars: { ${bits.join(', ')} }`);
    }
  }

  // axes
  const catAx = one(plot, 'catAx');
  const axLines = [];
  const axisYaml = (ax, isCat) => {
    if (!ax) return null;
    const bits = [];
    const titleEl = one(ax, 'title');
    if (titleEl) bits.push(`title: ${yaml(textOf(titleEl).trim())}`);
    const scaling = one(ax, 'scaling');
    if (!isCat && scaling) {
      const min = one(scaling, 'min')?.attrs.val;
      const max = one(scaling, 'max')?.attrs.val;
      if (min !== undefined) bits.push(`min: ${min}`);
      if (max !== undefined) bits.push(`max: ${max}`);
      if (one(scaling, 'logBase')) bits.push('log: true');
    }
    if (one(scaling ?? { children: [] }, 'orientation')?.attrs.val === 'maxMin') bits.push('reverse: true');
    const mu = one(ax, 'majorUnit')?.attrs.val;
    if (mu !== undefined) bits.push(`unit: ${mu}`);
    const mnu = one(ax, 'minorUnit')?.attrs.val;
    if (mnu !== undefined) bits.push(`minor-unit: ${mnu}`);
    if (one(ax, 'majorGridlines')) bits.push('gridlines: true');
    const nf = one(ax, 'numFmt')?.attrs.formatCode;
    if (nf) bits.push(`numfmt: ${yaml(nf)}`);
    return bits.length ? `{ ${bits.join(', ')} }` : null;
  };
  const x = axisYaml(catAx, true);
  const y = axisYaml(valAxes[0], false);
  const y2 = axisYaml(valAxes[1], false);
  if (x) axLines.push(`  x: ${x}`);
  if (y) axLines.push(`  y: ${y}`);
  if (y2) axLines.push(`  y2: ${y2}`);
  if (axLines.length) lines.push('axes:', ...axLines);

  const legend = one(chart, 'legend');
  if (legend) {
    const pos = LEGEND_REV.get(one(legend, 'legendPos')?.attrs.val) ?? 'right';
    lines.push(`legend: { position: ${pos} }`);
  } else lines.push('legend: { position: none }');
  const dt = one(plot, 'dTable');
  if (dt) lines.push(`data-table: { show: true, legend-keys: ${one(dt, 'showKeys')?.attrs.val !== '0'} }`);
  lines.push('```', '');
  return lines;
}

function attrVal(el, childName) {
  const c = one(el, childName);
  return c?.attrs.val;
}

function lineColor(spPr) {
  const ln = spPr && one(spPr, 'ln');
  return ln ? solidFillColor(ln) : null;
}

// ---- picture reversal ----
export function reversePicture(pic, anchorSpec, mediaLookup) {
  const blip = one(one(pic, 'blipFill') ?? { children: [] }, 'blip');
  const relId = blip && attr(blip, 'embed');
  const media = relId && mediaLookup(relId);
  if (!media) return null;
  const alt = one(one(pic, 'nvPicPr') ?? { children: [] }, 'cNvPr')?.attrs.descr;
  const lines = [`\`\`\`{image} ${fenceAnchor(anchorSpec)}`, `src: data:${media.mime};base64,${media.data.toString('base64')}`];
  if (alt) lines.push(`alt: ${yaml(alt)}`);
  lines.push('```', '');
  return lines;
}

// ---- shape / textbox reversal ----
export function reverseShape(sp, anchorSpec) {
  const spPr = one(sp, 'spPr');
  const preset = one(spPr ?? { children: [] }, 'prstGeom')?.attrs.prst;
  const kind = PRESET_REV.get(preset) ?? 'rect';
  const noFill = spPr && one(spPr, 'noFill');
  const fill = solidFillColor(spPr);
  const txBody = one(sp, 'txBody');
  const paras = txBody ? all(txBody, 'p').map((p) => all(p, 'r').map((r) => textOf(one(r, 't'))).join('')) : [];
  const text = paras.join('\n');
  const isTextbox = kind === 'rect' && (noFill || !fill);
  const lines = [`\`\`\`{${isTextbox ? 'textbox' : 'shape'}}${isTextbox ? '' : ` ${kind}`} ${fenceAnchor(anchorSpec)}`];
  if (text) {
    lines.push('text: |');
    for (const l of text.split('\n')) lines.push(`  ${l}`);
  }
  if (!isTextbox && fill) lines.push(`fill: ${yaml(fill)}`);
  const firstRPr = txBody && one(all(one(txBody, 'p') ?? { children: [] }, 'r')[0] ?? { children: [] }, 'rPr');
  if (firstRPr) {
    const bits = [];
    if (firstRPr.attrs.sz) bits.push(`size: ${Number(firstRPr.attrs.sz) / 100}`);
    if (firstRPr.attrs.b === '1') bits.push('bold: true');
    const c = solidFillColor(firstRPr);
    if (c) bits.push(`color: ${yaml(c)}`);
    if (bits.length) lines.push(`font: { ${bits.join(', ')} }`);
  }
  lines.push('```', '');
  return lines;
}

// ---- pivot reversal ----
export function reversePivot(ptDoc, cacheDoc) {
  const loc = one(ptDoc, 'location');
  if (!loc || !cacheDoc) return null;
  const fields = all(one(cacheDoc, 'cacheFields') ?? { children: [] }, 'cacheField').map((f) => f.attrs.name);
  if (!fields.length) return null;
  const anchor = loc.attrs.ref.split(':')[0];
  const src = one(one(cacheDoc, 'cacheSource') ?? { children: [] }, 'worksheetSource');
  const source = src?.attrs.name ?? (src?.attrs.sheet ? `${src.attrs.sheet}!${src.attrs.ref}` : null);
  if (!source) return null;
  const lines = [`\`\`\`{pivot} ${ptDoc.attrs.name ?? 'Pivot'} at ${anchor}`, `source: ${yaml(source)}`];
  const fieldList = (elName) => all(one(ptDoc, elName) ?? { children: [] }, 'field')
    .map((f) => fields[Number(f.attrs.x)]).filter(Boolean);
  const rows = fieldList('rowFields');
  const cols = fieldList('colFields');
  if (rows.length) lines.push('rows:', ...rows.map((f) => `  - { field: ${yaml(f)} }`));
  if (cols.length) lines.push('cols:', ...cols.map((f) => `  - { field: ${yaml(f)} }`));
  const pages = all(one(ptDoc, 'pageFields') ?? { children: [] }, 'pageField')
    .map((f) => fields[Number(f.attrs.fld)]).filter(Boolean);
  if (pages.length) lines.push('filters:', ...pages.map((f) => `  - { field: ${yaml(f)} }`));
  const dataFields = all(one(ptDoc, 'dataFields') ?? { children: [] }, 'dataField');
  if (dataFields.length) {
    lines.push('values:');
    for (const df of dataFields) {
      const bits = [`field: ${yaml(fields[Number(df.attrs.fld)] ?? '')}`];
      const agg = AGG_REV.get(df.attrs.subtotal) ?? 'sum';
      bits.push(`agg: ${agg}`);
      if (df.attrs.name) bits.push(`name: ${yaml(df.attrs.name)}`);
      if (df.attrs.showDataAs && SHOW_AS_REV.has(df.attrs.showDataAs)) bits.push(`show-as: ${SHOW_AS_REV.get(df.attrs.showDataAs)}`);
      lines.push(`  - { ${bits.join(', ')} }`);
    }
  }
  const layout = ptDoc.attrs.compact === '0' ? (ptDoc.attrs.outline === '0' ? 'tabular' : 'outline') : 'compact';
  if (layout !== 'compact') lines.push(`layout: ${layout}`);
  if (ptDoc.attrs.rowGrandTotals === '0' || ptDoc.attrs.colGrandTotals === '0') {
    lines.push(`grand-totals: { rows: ${ptDoc.attrs.rowGrandTotals !== '0'}, cols: ${ptDoc.attrs.colGrandTotals !== '0'} }`);
  }
  lines.push('```', '');
  return lines;
}

// ---- slicer reversal ----
export function reverseSlicer(slicerEl, cacheDoc, tableNameById, anchorSpec) {
  const cacheExt = cacheDoc && findDeep(cacheDoc, 'tableSlicerCache');
  if (!cacheExt) return null;
  const tableName = tableNameById.get(Number(cacheExt.attrs.tableId));
  if (!tableName) return null;
  const field = cacheDoc.attrs.sourceName;
  const lines = [`\`\`\`{slicer} ${fenceAnchor(anchorSpec ?? { at: 'A1', size: { w: 160, h: 200 } })}`,
    `for: ${yaml(tableName)}`, `field: ${yaml(field)}`, '```', ''];
  return lines;
}

function findDeep(el, name) {
  if (el.name === name) return el;
  for (const c of el.children) {
    const found = findDeep(c, name);
    if (found) return found;
  }
  return null;
}

// ---- ChartEx (cx:) reversal — restores type, series refs, title, legend ----
const FENCE = '```';

export function reverseChartEx(doc, anchorSpec) {
  const chart = one(doc, 'chart');
  const region = chart && one(one(chart, 'plotArea') ?? { children: [] }, 'plotAreaRegion');
  if (!region) return null;
  const seriesEls = all(region, 'series');
  if (!seriesEls.length) return null;
  const layouts = seriesEls.map((s) => s.attrs.layoutId);
  let type;
  if (layouts.includes('paretoLine')) type = 'pareto';
  else if (layouts[0] === 'clusteredColumn') type = 'histogram';
  else type = CX_TYPE_REV.get(layouts[0]) ?? null;
  if (!type) return null;
  const data = one(one(doc, 'chartData') ?? { children: [] }, 'data');
  const strDim = data && one(data, 'strDim');
  const numDim = data && one(data, 'numDim');
  const catRef = strDim ? textOf(one(strDim, 'f')).trim() : null;
  const valRef = numDim ? textOf(one(numDim, 'f')).trim() : null;
  if (!valRef) return null;
  const main = seriesEls.find((s) => s.attrs.layoutId !== 'paretoLine');
  const name = main && one(main, 'tx') ? textOf(one(main, 'tx')).trim() : null;
  const title = chart && one(chart, 'title') ? textOf(one(chart, 'title')).trim() : null;
  const lines = [
    `${FENCE}{chart} ${type}${title ? ` "${title.replace(/"/g, '""')}"` : ''} ${fenceAnchor(anchorSpec)}`,
    'series:',
    `  - name: ${yaml(name || 'Series 1')}`,
  ];
  if (catRef) lines.push(`    cat: ${yaml(catRef)}`);
  lines.push(`    val: ${yaml(valRef)}`);
  const legend = chart && one(chart, 'legend');
  if (legend) {
    const POS = { l: 'left', r: 'right', t: 'top', b: 'bottom' };
    lines.push(`legend: { position: ${POS[legend.attrs.pos] ?? 'right'} }`);
  } else lines.push('legend: { position: none }');
  lines.push(FENCE, '');
  return lines;
}
