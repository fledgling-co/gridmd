// XLSX → GridMD importer. Reverses the worksheet core natively (cells,
// formulas + cached values, styles, merges, tables + sort/filter state, CF,
// validation, notes, threaded comments, scenarios, sparklines, page setup,
// names, protection, freeze/view). Parts it cannot reverse-parse yet
// (charts, drawings, pivots, slicers, media, custom parts) are carried as
// {raw} blocks — loud, never silent.

import { zipRead } from './zip.js';
import { parseXml, attr, one, all, textOf, findDeep } from '../xml.js';
import { numToCol, colToNum, parseCell } from '../refs.js';
import { DEFAULT_THEME } from './units.js';
import { anchorsOf, anchorText, reverseChart, reversePicture, reverseShape, reversePivot, reverseSlicer } from './read-objects.js';

const DAY_MS = 86400000;

const BUILTIN_NUMFMTS = new Map([
  [1, '0'], [2, '0.00'], [3, '#,##0'], [4, '#,##0.00'], [9, '0%'], [10, '0.00%'],
  [11, '0.00E+00'], [12, '# ?/?'], [13, '# ??/??'], [14, 'm/d/yyyy'], [15, 'd-mmm-yy'],
  [16, 'd-mmm'], [17, 'mmm-yy'], [18, 'h:mm AM/PM'], [19, 'h:mm:ss AM/PM'],
  [20, 'h:mm'], [21, 'h:mm:ss'], [22, 'm/d/yyyy h:mm'], [37, '#,##0_);(#,##0)'],
  [38, '#,##0_);[Red](#,##0)'], [45, 'mm:ss'], [46, '[h]:mm:ss'], [47, 'mm:ss.0'], [49, '@'],
]);
const THEME_SLOTS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
const PAPER_REV = new Map([[1, 'letter'], [3, 'tabloid'], [5, 'legal'], [8, 'A3'], [9, 'A4'], [11, 'A5']]);
const CELLIS_REV = new Map(Object.entries({
  equal: '=', notEqual: '<>', greaterThan: '>', greaterThanOrEqual: '>=',
  lessThan: '<', lessThanOrEqual: '<=',
}));
const TIME_PERIOD_REV = new Map(Object.entries({
  yesterday: 'yesterday', today: 'today', tomorrow: 'tomorrow',
  last7Days: 'last-7-days', thisWeek: 'this-week', lastWeek: 'last-week',
  nextWeek: 'next-week', thisMonth: 'this-month', lastMonth: 'last-month', nextMonth: 'next-month',
}));
const ICON_REV = new Map([
  ['3Arrows', '3-arrows'], ['3ArrowsGray', '3-arrows-gray'], ['3Flags', '3-flags'],
  ['3TrafficLights1', '3-traffic-lights'], ['3TrafficLights2', '3-traffic-lights-rimmed'],
  ['3Signs', '3-signs'], ['3Symbols', '3-symbols'], ['3Symbols2', '3-symbols-circled'],
  ['3Stars', '3-stars'], ['3Triangles', '3-triangles'], ['4Arrows', '4-arrows'],
  ['4ArrowsGray', '4-arrows-gray'], ['4RedToBlack', '4-red-to-black'], ['4Rating', '4-ratings'],
  ['4TrafficLights', '4-traffic-lights'], ['5Arrows', '5-arrows'], ['5ArrowsGray', '5-arrows-gray'],
  ['5Rating', '5-ratings'], ['5Quarters', '5-quarters'], ['5Boxes', '5-boxes'],
]);
const SUBTOTAL_REV = new Map([['average', 101], ['countNums', 102], ['count', 103],
  ['max', 104], ['min', 105], ['product', 106], ['stdDev', 107], ['sum', 109], ['var', 110]]);
const PROTECT_REV = new Map(Object.entries({
  formatCells: 'format-cells', formatColumns: 'format-columns', formatRows: 'format-rows',
  insertColumns: 'insert-columns', insertRows: 'insert-rows', insertHyperlinks: 'insert-hyperlinks',
  deleteColumns: 'delete-columns', deleteRows: 'delete-rows', sort: 'sort',
  autoFilter: 'autofilter', pivotTables: 'pivot-tables', objects: 'objects', scenarios: 'scenarios',
}));
const BORDER_STYLE_REV = new Map(Object.entries({
  hair: 'hair', thin: 'thin', medium: 'medium', thick: 'thick', double: 'double',
  dotted: 'dotted', dashed: 'dashed', dashDot: 'dash-dot', dashDotDot: 'dash-dot-dot',
  mediumDashed: 'medium-dashed', mediumDashDot: 'medium-dash-dot',
  mediumDashDotDot: 'medium-dash-dot-dot', slantDashDot: 'slant-dash-dot',
}));

const CONSUMED_PREFIXES = [/^\[Content_Types\]\.xml$/, /^_rels\//, /^docProps\//,
  /^xl\/workbook\.xml$/, /^xl\/_rels\//, /^xl\/styles\.xml$/, /^xl\/sharedStrings\.xml$/,
  /^xl\/theme\//, /^xl\/worksheets\//, /^xl\/chartsheets\//, /^xl\/tables\//,
  /^xl\/comments\d+\.xml$/, /^xl\/threadedComments\//, /^xl\/persons\//];

export function xlsxToGridmd(buffer) {
  const parts = zipRead(buffer);
  const report = [];
  const consumed = new Set();
  const xml = (name) => {
    if (!parts.has(name)) return null;
    consumed.add(name);
    return parseXml(parts.get(name).toString('utf8'));
  };
  const relsFor = (partName) => {
    const dir = partName.slice(0, partName.lastIndexOf('/'));
    const relName = `${dir}/_rels/${partName.slice(partName.lastIndexOf('/') + 1)}.rels`;
    const map = new Map();
    const doc = xml(relName);
    if (!doc) return map;
    for (const r of all(doc, 'Relationship')) {
      let target = r.attrs.Target;
      if (r.attrs.TargetMode === 'External') {
        map.set(r.attrs.Id, { target, type: r.attrs.Type, mode: 'External' });
        continue;
      }
      if (!target.startsWith('/')) {
        const segs = `${dir}/${target}`.split('/');
        const outSegs = [];
        for (const s of segs) {
          if (s === '..') outSegs.pop();
          else if (s !== '.') outSegs.push(s);
        }
        target = outSegs.join('/');
      } else target = target.slice(1);
      map.set(r.attrs.Id, { target, type: r.attrs.Type, mode: r.attrs.TargetMode });
    }
    return map;
  };

  // ---- workbook ----
  const wb = xml('xl/workbook.xml');
  if (!wb) throw new Error('not an xlsx: xl/workbook.xml missing');
  const wbRels = relsFor('xl/workbook.xml');
  const date1904 = one(wb, 'workbookPr') && attr(one(wb, 'workbookPr'), 'date1904') === '1';
  const dateSystem = date1904 ? 1904 : 1900;
  const calcPr = one(wb, 'calcPr');
  const sheetsEl = all(one(wb, 'sheets'), 'sheet');
  const core = xml('docProps/core.xml');
  consumed.add('docProps/app.xml');

  // ---- theme + shared strings + styles ----
  const themeColors = readTheme(xml('xl/theme/theme1.xml'));
  const shared = readSharedStrings(xml('xl/sharedStrings.xml'));
  const st = readStyles(xml('xl/styles.xml'), themeColors);

  // Pre-scan table + slicer-cache parts (slicers reference tables by numeric
  // id, possibly across sheets; peeking does not consume).
  const tableNameById = new Map();
  for (const name of parts.keys()) {
    if (/^xl\/tables\/table\d+\.xml$/.test(name)) {
      const doc = parseXml(parts.get(name).toString('utf8'));
      tableNameById.set(Number(doc.attrs.id), doc.attrs.name);
    }
  }
  const slicerCaches = new Map(); // cache name -> { part, doc }
  for (const name of parts.keys()) {
    if (/^xl\/slicerCaches\//.test(name)) {
      const doc = parseXml(parts.get(name).toString('utf8'));
      slicerCaches.set(doc.attrs.name, { part: name, doc });
    }
  }
  const timelineCaches = new Map(); // cache name -> { part, doc }
  for (const name of parts.keys()) {
    if (/^xl\/timelineCaches\//.test(name)) {
      const doc = parseXml(parts.get(name).toString('utf8'));
      timelineCaches.set(doc.attrs.name, { part: name, doc });
    }
  }

  // ---- defined names ----
  const names = [];
  const pageNames = new Map(); // sheetIdx -> { printArea, printTitles }
  for (const dn of all(one(wb, 'definedNames') ?? { children: [] }, 'definedName')) {
    const name = dn.attrs.name;
    const value = textOf(dn).trim();
    const localId = dn.attrs.localSheetId !== undefined ? Number(dn.attrs.localSheetId) : null;
    if (name === '_xlnm.Print_Area' || name === '_xlnm.Print_Titles') {
      if (localId !== null) {
        const pn = pageNames.get(localId) ?? {};
        pn[name === '_xlnm.Print_Area' ? 'area' : 'titles'] = value;
        pageNames.set(localId, pn);
      }
      continue;
    }
    names.push({ name, value, localId, hidden: dn.attrs.hidden === '1' });
  }

  // ---- frontmatter ----
  const fmLines = ['---', 'gridmd: "0.1"'];
  const title = core ? textOf(one(core, 'title') ?? { children: [], text: '' }).trim() : '';
  const creator = core ? textOf(one(core, 'creator') ?? { children: [], text: '' }).trim() : '';
  if (title) fmLines.push(`title: ${yamlStr(title)}`);
  if (creator && creator !== 'GridMD') fmLines.push(`properties:`, `  author: ${yamlStr(creator)}`);
  if (dateSystem === 1904) fmLines.push('date-system: 1904');
  if (calcPr) {
    const mode = calcPr.attrs.calcMode;
    const iterate = calcPr.attrs.iterate === '1';
    if (mode === 'manual' || mode === 'autoNoTable' || iterate) {
      fmLines.push('calc:');
      fmLines.push(`  mode: ${mode === 'manual' ? 'manual' : mode === 'autoNoTable' ? 'auto-no-tables' : 'auto'}`);
      if (iterate) fmLines.push(`  iterative: { enabled: true, max-iterations: ${calcPr.attrs.iterateCount ?? 100}, max-change: ${calcPr.attrs.iterateDelta ?? 0.001} }`);
    }
  }
  const wbNames = names.filter((n) => n.localId === null);
  if (wbNames.length) {
    fmLines.push('names:');
    for (const n of wbNames) {
      fmLines.push(`  - { name: ${yamlStr(n.name)}, ${classifyName(n.value)}: ${yamlStr(n.value)}${n.hidden ? ', hidden: true' : ''} }`);
    }
  }
  if (one(wb, 'workbookProtection')?.attrs.lockStructure === '1') {
    fmLines.push('protection:', '  structure: true');
  }
  fmLines.push('---', '');

  // ---- sheets ----
  const body = [];
  sheetsEl.forEach((sh, sheetIdx) => {
    const rel = wbRels.get(attr(sh, 'id'));
    const target = rel?.target;
    if (!target || !parts.has(target)) return;
    if (target.startsWith('xl/chartsheets/')) {
      xml(target);
      const csRels = relsFor(target);
      let chartLines = null;
      for (const [, rel] of csRels) {
        if (!rel.target.startsWith('xl/drawings/')) continue;
        const dDoc = xml(rel.target);
        const dRels = relsFor(rel.target);
        for (const anchor of anchorsOf(dDoc ?? { children: [] })) {
          const chartEl = findDeep(anchor.el, 'chart');
          const cTarget = chartEl && dRels.get(attr(chartEl, 'id'))?.target;
          if (!cTarget || !parts.has(cTarget)) continue;
          const lines = reverseChart(parseXml(parts.get(cTarget).toString('utf8')), { at: 'sheet' });
          if (lines) { consumed.add(cTarget); chartLines = lines; }
        }
      }
      if (chartLines) {
        report.push({ feature: `chart sheet "${sh.attrs.name}"`, action: 'native' });
        body.push(`# ${sh.attrs.name}`, '', '```{sheet}', 'kind: chart', '```', '', ...chartLines);
      } else {
        report.push({ feature: `chart sheet "${sh.attrs.name}"`, action: 'carried', note: 'unrecognized chart; parts carried as {raw}' });
        body.push(`# ${sh.attrs.name}`, '',
          '> Imported chart sheet — its chart is carried as a {raw} part before the sheets.', '');
      }
      return;
    }
    const ws = xml(target);
    const wsRels = relsFor(target);
    body.push(...renderSheet(sh, ws, wsRels, sheetIdx));
  });

  // ---- carried leftovers ----
  const rawBlocks = [];
  for (const name of parts.keys()) {
    if (consumed.has(name)) continue;
    if (CONSUMED_PREFIXES.some((re) => re.test(name))) continue;
    const data = parts.get(name);
    const isXml = name.endsWith('.xml') || name.endsWith('.rels') || name.endsWith('.vml');
    report.push({ feature: name, action: 'carried', note: 'carried as {raw} (not reverse-parsed)' });
    if (isXml) {
      rawBlocks.push(`\`\`\`{raw} ooxml part="${name}"`, data.toString('utf8').replace(/^﻿/, ''), '```', '');
    } else {
      const b64 = data.toString('base64').replace(/(.{76})/g, '$1\n').trimEnd();
      rawBlocks.push(`\`\`\`{raw} ooxml part="${name}" encoding=base64`, b64, '```', '');
    }
  }

  const gmd = [...fmLines, ...(rawBlocks.length ? ['> Parts carried from the source package (not yet reverse-parsed).', '', ...rawBlocks] : []), ...body].join('\n');
  return { gmd, report };

  // ================= per-sheet =================
  function renderSheet(sh, ws, wsRels, sheetIdx) {
    const out = [`# ${sh.attrs.name}`, ''];
    const cellsByRef = new Map();
    const arrayAnchors = [];

    // ---- sheet meta ----
    const metaLines = [];
    if (one(ws, 'sheetPr') && one(one(ws, 'sheetPr'), 'tabColor')) {
      const c = colorFrom(one(one(ws, 'sheetPr'), 'tabColor'), themeColors);
      if (c) metaLines.push(`tab-color: ${yamlStr(c)}`);
    }
    if (sh.attrs.state === 'hidden') metaLines.push('hidden: true');
    if (sh.attrs.state === 'veryHidden') metaLines.push('hidden: very');
    const view = one(ws, 'sheetViews') && one(one(ws, 'sheetViews'), 'sheetView');
    if (view) {
      const pane = one(view, 'pane');
      if (pane?.attrs.state === 'frozen' && pane.attrs.topLeftCell) metaLines.push(`freeze: ${pane.attrs.topLeftCell}`);
      const vbits = [];
      if (view.attrs.showGridLines === '0') vbits.push('gridlines: false');
      if (view.attrs.showRowColHeaders === '0') vbits.push('headings: false');
      if (view.attrs.showFormulas === '1') vbits.push('formulas: true');
      if (view.attrs.rightToLeft === '1') vbits.push('rtl: true');
      if (view.attrs.zoomScale && view.attrs.zoomScale !== '100') vbits.push(`zoom: ${view.attrs.zoomScale}`);
      if (vbits.length) metaLines.push(`view: { ${vbits.join(', ')} }`);
    }
    const colEls = one(ws, 'cols') ? all(one(ws, 'cols'), 'col') : [];
    if (colEls.length) {
      metaLines.push('cols:');
      for (const c of colEls) {
        const min = Number(c.attrs.min), max = Number(c.attrs.max);
        const key = min === max ? numToCol(min) : `"${numToCol(min)}:${numToCol(max)}"`;
        const bits = [];
        if (c.attrs.width !== undefined && c.attrs.customWidth === '1') bits.push(`width: ${Math.round(Number(c.attrs.width) * 7 + 5)}`);
        if (c.attrs.hidden === '1') bits.push('hidden: true');
        if (c.attrs.outlineLevel) bits.push(`group: ${c.attrs.outlineLevel}`);
        if (bits.length === 1 && bits[0].startsWith('width:')) metaLines.push(`  ${key}: ${bits[0].slice(7)}`);
        else if (bits.length) metaLines.push(`  ${key}: { ${bits.join(', ')} }`);
      }
    }
    const rowMeta = [];
    for (const r of all(one(ws, 'sheetData') ?? { children: [] }, 'row')) {
      const bits = [];
      if (r.attrs.ht !== undefined && r.attrs.customHeight === '1') bits.push(`height: ${Math.round(Number(r.attrs.ht) * 96 / 72)}`);
      if (r.attrs.hidden === '1') bits.push('hidden: true');
      if (r.attrs.outlineLevel) bits.push(`group: ${r.attrs.outlineLevel}`);
      if (bits.length) rowMeta.push(`  ${r.attrs.r}: { ${bits.join(', ')} }`);
    }
    if (rowMeta.length) metaLines.push('rows:', ...rowMeta);
    const prot = one(ws, 'sheetProtection');
    if (prot?.attrs.sheet === '1') {
      const allow = [];
      for (const [xmlAttr, gm] of PROTECT_REV) if (prot.attrs[xmlAttr] === '0') allow.push(gm);
      if (prot.attrs.selectLockedCells !== '1') allow.push('select-locked');
      if (prot.attrs.selectUnlockedCells !== '1') allow.push('select-unlocked');
      metaLines.push('protect:', '  enabled: true', `  allow: [${allow.join(', ')}]`);
    }
    const sheetNames = names.filter((x) => x.localId === sheetIdx);
    if (sheetNames.length) {
      metaLines.push('names:');
      for (const n of sheetNames) metaLines.push(`  - { name: ${yamlStr(n.name)}, ${classifyName(n.value)}: ${yamlStr(n.value)} }`);
    }
    if (metaLines.length) out.push('```{sheet}', ...metaLines, '```', '');

    // ---- cells ----
    for (const r of all(one(ws, 'sheetData') ?? { children: [] }, 'row')) {
      for (const c of all(r, 'c')) {
        const ref = c.attrs.r;
        const pos = parseCell(ref);
        if (!pos) continue;
        const cell = readCell(c, pos);
        cellsByRef.set(ref, cell);
        if (cell.arrayRef && cell.arrayRef !== `${ref}:${ref}`) arrayAnchors.push(cell);
      }
    }

    // ---- tables (consume their cells) ----
    const tableRanges = [];
    const tableBlocks = [];
    for (const [, rel] of wsRels) {
      if (!rel.target.startsWith('xl/tables/')) continue;
      const t = xml(rel.target);
      if (!t) continue;
      tableBlocks.push(renderTable(t, cellsByRef, tableRanges));
    }

    // ---- drawings: charts / pictures / shapes / slicer frames ----
    const drawingBlocks = [];
    const slicerAnchors = new Map();
    for (const [, rel] of wsRels) {
      if (!/^xl\/drawings\/drawing\d+\.xml$/.test(rel.target)) continue;
      const dDoc = xml(rel.target);
      if (!dDoc) continue;
      const dRels = relsFor(rel.target);
      const mediaLookup = (rid) => {
        const t = dRels.get(rid)?.target;
        if (!t || !parts.has(t)) return null;
        consumed.add(t);
        const ext = t.split('.').pop().toLowerCase();
        const mime = ext === 'png' ? 'image/png' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : 'application/octet-stream';
        return { mime, data: parts.get(t) };
      };
      for (const anchor of anchorsOf(dDoc)) {
        const spec = anchorText(anchor);
        const chartEl = findDeep(anchor.el, 'chart');
        if (chartEl) {
          const cTarget = dRels.get(attr(chartEl, 'id'))?.target;
          if (cTarget && parts.has(cTarget)) {
            const lines = reverseChart(parseXml(parts.get(cTarget).toString('utf8')), spec);
            if (lines) { consumed.add(cTarget); drawingBlocks.push(...lines); continue; }
            report.push({ feature: cTarget, action: 'carried', note: 'unrecognized chart form (ChartEx?) — carried as {raw}' });
          }
          continue;
        }
        const slicerEl = findDeep(anchor.el, 'slicer');
        if (slicerEl?.attrs.name) { slicerAnchors.set(slicerEl.attrs.name, spec); continue; }
        const tsEl = findDeep(anchor.el, 'timeslicer');
        if (tsEl?.attrs.name) { slicerAnchors.set(tsEl.attrs.name, spec); continue; }
        const pic = one(anchor.el, 'pic');
        if (pic) {
          const lines = reversePicture(pic, spec, mediaLookup);
          if (lines) drawingBlocks.push(...lines);
          continue;
        }
        const sp = one(anchor.el, 'sp');
        if (sp) drawingBlocks.push(...reverseShape(sp, spec));
      }
    }

    // ---- pivots ----
    for (const [, rel] of wsRels) {
      if (!rel.target.startsWith('xl/pivotTables/')) continue;
      const ptDoc = xml(rel.target);
      if (!ptDoc) continue;
      const ptRels = relsFor(rel.target);
      const cacheTarget = [...ptRels.values()].find((r) => r.target.includes('pivotCacheDefinition'))?.target;
      const cacheDoc = cacheTarget ? xml(cacheTarget) : null;
      if (cacheTarget) {
        for (const [, r2] of relsFor(cacheTarget)) if (r2.target.includes('Records')) consumed.add(r2.target);
      }
      const lines = cacheDoc && reversePivot(ptDoc, cacheDoc);
      if (lines) drawingBlocks.push(...lines);
      else report.push({ feature: rel.target, action: 'carried', note: 'unrecognized pivot form — carried as {raw}' });
    }

    // ---- slicers ----
    for (const [, rel] of wsRels) {
      if (!rel.target.startsWith('xl/slicers/')) continue;
      const sDoc = xml(rel.target);
      if (!sDoc) continue;
      for (const sl of all(sDoc, 'slicer')) {
        const cache = slicerCaches.get(sl.attrs.cache);
        const lines = cache && reverseSlicer(sl, cache.doc, tableNameById, slicerAnchors.get(sl.attrs.name));
        if (lines) { consumed.add(cache.part); drawingBlocks.push(...lines); }
        else report.push({ feature: rel.target, action: 'carried', note: 'unrecognized slicer form — carried as {raw}' });
      }
    }

    // ---- timelines ----
    for (const [, rel] of wsRels) {
      if (!rel.target.startsWith('xl/timelines/')) continue;
      const tDoc = xml(rel.target);
      if (!tDoc) continue;
      for (const tl of all(tDoc, 'timeline')) {
        const cache = timelineCaches.get(tl.attrs.cache);
        const pivotName = cache && one(one(cache.doc, 'pivotTables') ?? { children: [] }, 'pivotTable')?.attrs.name;
        if (!pivotName) {
          report.push({ feature: rel.target, action: 'carried', note: 'unrecognized timeline form — carried as {raw}' });
          continue;
        }
        consumed.add(cache.part);
        const spec = slicerAnchors.get(tl.attrs.name) ?? { at: 'A1', size: { w: 320, h: 110 } };
        const LEVEL_REV = ['years', 'quarters', 'months', 'days'];
        drawingBlocks.push(
          `\`\`\`{slicer} at ${spec.at}${spec.size ? ` size ${spec.size.w}x${spec.size.h}` : ''}`,
          'kind: timeline',
          `for: ${pivotName}`,
          `field: ${cache.doc.attrs.sourceName}`,
          `level: ${LEVEL_REV[Number(tl.attrs.level ?? 2)] ?? 'months'}`,
          '```', '');
      }
    }

    // spill-cache coverage
    const spillCovered = new Set();
    for (const anc of arrayAnchors) {
      const [a, b] = anc.arrayRef.split(':').map(parseCell);
      for (let rr = a.row; rr <= b.row; rr++) for (let cc = a.col; cc <= b.col; cc++) {
        if (rr === a.row && cc === a.col) continue;
        spillCovered.add(`${numToCol(cc)}${rr}`);
      }
    }
    const inTable = (pos) => tableRanges.some((tr) => pos.col >= tr.c1 && pos.col <= tr.c2 && pos.row >= tr.r1 && pos.row <= tr.r2);

    // hyperlinks / notes / threads attach to cells
    const linkByRef = new Map();
    for (const h of all(one(ws, 'hyperlinks') ?? { children: [] }, 'hyperlink')) {
      const target = h.attrs.location ? `#${h.attrs.location}` : wsRels.get(attr(h, 'id'))?.target;
      if (target) linkByRef.set(h.attrs.ref, { link: target, tip: h.attrs.tooltip });
    }
    const noteByRef = new Map();
    const threads = [];
    const threadRefs = new Set();
    for (const [, rel] of wsRels) {
      if (rel.target.startsWith('xl/threadedComments/')) {
        const tc = xml(rel.target);
        const personsMap = readPersons();
        const roots = new Map();
        for (const el of all(tc, 'threadedComment')) {
          const item = {
            ref: el.attrs.ref, by: personsMap.get(el.attrs.personId) ?? 'Unknown',
            at: el.attrs.dT, text: textOf(one(el, 'text')), done: el.attrs.done === '1',
            id: el.attrs.id, parentId: el.attrs.parentId,
          };
          threadRefs.add(item.ref);
          if (!item.parentId) { roots.set(item.id, { ...item, replies: [] }); }
          else roots.get(item.parentId)?.replies.push(item);
        }
        for (const root of roots.values()) threads.push(root);
      }
    }
    for (const [, rel] of wsRels) {
      if (/^xl\/comments\d+\.xml$/.test(rel.target)) {
        const cm = xml(rel.target);
        for (const el of all(one(cm, 'commentList'), 'comment')) {
          if (threadRefs.has(el.attrs.ref)) continue; // threaded-comment legacy shim
          noteByRef.set(el.attrs.ref, textOf(one(el, 'text')));
        }
      }
      if (/vmlDrawing\d+\.vml$/.test(rel.target)) consumed.add(rel.target);
    }

    // ---- emit cells ----
    const ordered = [...cellsByRef.entries()]
      .map(([ref, cell]) => ({ ref, cell, pos: parseCell(ref) }))
      .sort((a, b) => a.pos.row - b.pos.row || a.pos.col - b.pos.col);
    for (const { ref, cell, pos } of ordered) {
      if (inTable(pos)) continue;
      const extras = { ...(linkByRef.get(ref) ?? {}) };
      const note = noteByRef.get(ref);
      if (spillCovered.has(ref)) continue; // emitted via {spill-cache}
      const line = cellLine(ref, cell, extras, note);
      if (line) out.push(...line);
      if (cell.arrayRef && cell.arrayRef !== `${ref}:${ref}`) {
        out.push(`\`\`\`{spill-cache} ${ref}`);
        const [a, b] = cell.arrayRef.split(':').map(parseCell);
        for (let rr = a.row; rr <= b.row; rr++) {
          const rowCells = [];
          for (let cc = a.col; cc <= b.col; cc++) {
            const covered = cellsByRef.get(`${numToCol(cc)}${rr}`);
            rowCells.push(rr === a.row && cc === a.col
              ? (cell.cachedText ?? '')
              : covered ? scalarTextOf(covered) : '');
          }
          out.push(`| ${rowCells.join(' | ')} |`);
        }
        out.push('```');
      }
    }
    if (ordered.length) out.push('');

    // merges
    for (const m of all(one(ws, 'mergeCells') ?? { children: [] }, 'mergeCell')) {
      out.push(`@ ${m.attrs.ref} { merge: true }`);
    }

    out.push(...tableBlocks.flat());
    out.push(...drawingBlocks);

    // conditional formatting
    for (const cf of all(ws, 'conditionalFormatting')) {
      const rules = all(cf, 'cfRule').map(cfRuleToYaml).filter(Boolean);
      if (rules.length) out.push(`\`\`\`{cf} ${cf.attrs.sqref}`, ...rules.flat(), '```', '');
    }

    // validations
    const dv = one(ws, 'dataValidations');
    for (const v of all(dv ?? { children: [] }, 'dataValidation')) {
      out.push(...validationToBlock(v));
    }

    // standalone filter + sort
    const af = one(ws, 'autoFilter');
    if (af) out.push(...filterToBlock(af, one(ws, 'sortState')));

    // sparklines
    for (const ext of all(one(ws, 'extLst') ?? { children: [] }, 'ext')) {
      const groups = one(ext, 'sparklineGroups');
      if (groups) for (const g of all(groups, 'sparklineGroup')) out.push(...sparklineToBlock(g));
    }

    // threads
    for (const th of threads) {
      out.push(`\`\`\`{comments} ${th.ref}`);
      out.push(`- by: ${yamlStr(th.by)}`, `  at: ${th.at}Z`.replace(/ZZ$/, 'Z'), `  text: ${yamlStr(th.text)}`);
      if (th.done) out.push('  resolved: true');
      if (th.replies.length) {
        out.push('  replies:');
        for (const rep of th.replies) out.push(`    - { by: ${yamlStr(rep.by)}, at: ${rep.at}Z, text: ${yamlStr(rep.text)} }`.replace(/ZZ,/, 'Z,'));
      }
      out.push('```', '');
    }

    // scenarios
    for (const sc of all(one(ws, 'scenarios') ?? { children: [] }, 'scenario')) {
      out.push(`\`\`\`{scenario} ${sc.attrs.name}`);
      const cells = all(sc, 'inputCells').map((ic) => `${ic.attrs.r}: ${yamlMaybeNum(ic.attrs.val)}`);
      out.push(`cells: { ${cells.join(', ')} }`);
      if (sc.attrs.comment) out.push(`comment: ${yamlStr(sc.attrs.comment)}`);
      out.push('```', '');
    }

    // page setup
    out.push(...pageToBlock(ws, pageNames.get(sheetIdx)));
    return out;
  }

  // ================= cell machinery =================
  function readCell(c, pos) {
    const t = c.attrs.t ?? 'n';
    const xf = c.attrs.s ? Number(c.attrs.s) : 0;
    const props = st.propsForXf(xf);
    const fEl = one(c, 'f');
    const vEl = one(c, 'v');
    const isEl = one(c, 'is');
    const cell = { props, formula: null, arrayRef: null, cachedText: null, scalarText: null, rich: null };

    const numText = (raw) => {
      const n = Number(raw);
      if (props.numfmt && isDateFormat(props.numfmt)) {
        const iso = serialToIso(n, dateSystem, props.numfmt);
        if (iso) { delete cell.props.numfmt; return iso; }
      }
      return String(n);
    };

    if (fEl) {
      cell.formula = textOf(fEl);
      if (fEl.attrs.t === 'array' && fEl.attrs.ref) cell.arrayRef = fEl.attrs.ref;
      if (vEl) {
        const raw = textOf(vEl);
        cell.cachedText = t === 'b' ? (raw === '1' ? 'TRUE' : 'FALSE')
          : t === 'e' ? raw
          : t === 'str' ? quoteText(raw, true)
          : numText(raw);
      }
      return cell;
    }
    if (isEl) {
      const runs = all(isEl, 'r');
      if (runs.length > 1 || (runs.length === 1 && one(runs[0], 'rPr'))) {
        cell.rich = runs.map((run) => {
          const pr = one(run, 'rPr');
          const item = { text: textOf(one(run, 't')) };
          if (pr) {
            if (one(pr, 'b')) item.bold = true;
            if (one(pr, 'i')) item.italic = true;
            const sz = one(pr, 'sz');
            if (sz) item.size = Number(sz.attrs.val);
            const col = one(pr, 'color');
            const cc = col && colorFrom(col, themeColors);
            if (cc) item.color = cc;
          }
          return item;
        });
      } else cell.scalarText = quoteText(textOf(isEl), false);
      return cell;
    }
    if (!vEl) return cell;
    const raw = textOf(vEl);
    cell.scalarText = t === 's' ? sharedText(Number(raw), cell)
      : t === 'b' ? (raw === '1' ? 'TRUE' : 'FALSE')
      : t === 'e' ? raw
      : t === 'str' ? quoteText(raw, false)
      : numText(raw);
    return cell;
  }

  function sharedText(idx, cell) {
    const si = shared[idx];
    if (!si) return '""';
    if (si.rich) { cell.rich = si.rich; return null; }
    return quoteText(si.text, false);
  }

  function scalarTextOf(cell) {
    return cell.scalarText ?? cell.cachedText ?? '';
  }

  function cellLine(ref, cell, extras, note) {
    const props = { ...cell.props };
    if (extras.link) { props.link = extras.link; if (extras.tip) props.tip = extras.tip; }
    const propKeys = Object.keys(props);
    const flow = propKeys.length ? ` { ${propKeys.map((k) => `${k}: ${yamlPropVal(props[k])}`).join(', ')} }` : '';

    if (cell.rich || note !== undefined || (cell.scalarText ?? '').includes('\n')) {
      const lines = [`@ ${ref}`];
      if (cell.rich) {
        lines.push('  rich:');
        for (const r of cell.rich) {
          const bits = [`text: ${yamlStr(r.text)}`];
          if (r.bold) bits.push('bold: true');
          if (r.italic) bits.push('italic: true');
          if (r.size) bits.push(`size: ${r.size}`);
          if (r.color) bits.push(`color: ${yamlStr(r.color)}`);
          lines.push(`    - { ${bits.join(', ')} }`);
        }
      } else if (cell.formula) {
        lines.push(`  formula: =${cell.formula}`);
        if (cell.cachedText !== null) lines.push(`  value: ${bodyValue(cell.cachedText)}`);
      } else if (cell.scalarText !== null) {
        if (cell.scalarText.includes('\n')) {
          lines.push('  value: |');
          for (const l of unquote(cell.scalarText).split('\n')) lines.push(`    ${l}`);
        } else lines.push(`  value: ${bodyValue(cell.scalarText)}`);
      }
      for (const k of propKeys) lines.push(`  ${k}: ${yamlPropVal(props[k])}`);
      if (note !== undefined) {
        lines.push('  note: |');
        for (const l of String(note).split('\n')) lines.push(`    ${l}`);
      }
      return lines;
    }

    if (cell.formula !== null) {
      const arr = cell.arrayRef && cell.arrayRef !== `${ref}:${ref}` ? { array: cell.arrayRef } : cell.arrayRef ? { array: cell.arrayRef } : null;
      const arrFlow = arr ? (propKeys.length ? flow.replace(' }', `, array: ${arr.array} }`) : ` { array: ${arr.array} }`) : flow;
      const cachedPart = cell.cachedText !== null ? ` :: ${cell.cachedText}` : '';
      return [`@ ${ref} =${cell.formula}${cachedPart}${arrFlow}`];
    }
    if (cell.scalarText !== null && cell.scalarText !== '') return [`@ ${ref} ${cell.scalarText}${flow}`];
    if (propKeys.length) return [`@ ${ref}${flow}`];
    return null;
  }

  // ================= directive reversal =================
  function renderTable(t, cellsByRef, tableRanges) {
    const [a, b] = t.attrs.ref.split(':').map(parseCell);
    tableRanges.push({ c1: a.col, r1: a.row, c2: b.col, r2: b.row });
    const totals = t.attrs.totalsRowCount === '1';
    const headerRows = t.attrs.headerRowCount === '0' ? 0 : 1;
    const colsEl = all(one(t, 'tableColumns'), 'tableColumn');
    const columns = colsEl.map((c) => c.attrs.name);
    const meta = [];
    const info = one(t, 'tableStyleInfo');
    if (info?.attrs.name) {
      const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(info.attrs.name);
      if (m) meta.push(`style: ${m[1].toLowerCase()}-${m[2]}`);
      const rows = info.attrs.showRowStripes === '1', colsB = info.attrs.showColumnStripes === '1';
      meta.push(`banded: ${rows && colsB ? 'both' : colsB ? 'cols' : rows ? 'rows' : 'none'}`);
    }
    if (totals) {
      meta.push('total:');
      colsEl.forEach((c, i) => {
        if (c.attrs.totalsRowLabel) meta.push(`  ${yamlStr(c.attrs.name)}: ${yamlStr(c.attrs.totalsRowLabel)}`);
        else if (c.attrs.totalsRowFunction && SUBTOTAL_REV.has(c.attrs.totalsRowFunction)) {
          meta.push(`  ${yamlStr(c.attrs.name)}: =SUBTOTAL(${SUBTOTAL_REV.get(c.attrs.totalsRowFunction)},[${c.attrs.name}])`);
        } else if (c.attrs.totalsRowFunction === 'custom') {
          const cellRef = `${numToCol(a.col + i)}${b.row}`;
          const cell = cellsByRef.get(cellRef);
          if (cell?.formula) meta.push(`  ${yamlStr(c.attrs.name)}: =${cell.formula}`);
        }
      });
    }
    const af = one(t, 'autoFilter');
    const filterEntries = [];
    for (const fc of all(af ?? { children: [] }, 'filterColumn')) {
      const colName = columns[Number(fc.attrs.colId)];
      const filters = one(fc, 'filters');
      if (colName && filters) {
        filterEntries.push(`  ${yamlStr(colName)}: { values: [${all(filters, 'filter').map((f) => yamlStr(f.attrs.val)).join(', ')}] }`);
      }
    }
    if (filterEntries.length) meta.push('filter:', ...filterEntries);
    const ss = one(t, 'sortState');
    if (ss) {
      const levels = all(ss, 'sortCondition').map((scond) => {
        const colLetter = /^([A-Z]{1,3})/.exec(scond.attrs.ref)?.[1];
        const idx = colLetter ? colToNum(colLetter) - a.col : -1;
        const colName = columns[idx];
        return colName ? `  - { col: ${yamlStr(colName)}, order: ${scond.attrs.descending === '1' ? 'desc' : 'asc'}, by: value }` : null;
      }).filter(Boolean);
      if (levels.length) meta.push('sort:', ...levels);
    }
    // uniform column numfmt → cols: meta
    const colMeta = [];
    const bodyR1 = a.row + headerRows;
    const bodyR2 = totals ? b.row - 1 : b.row;
    columns.forEach((name, i) => {
      const fmts = new Set();
      for (let rr = bodyR1; rr <= bodyR2; rr++) {
        const cell = cellsByRef.get(`${numToCol(a.col + i)}${rr}`);
        fmts.add(cell?.props?.numfmt);
      }
      if (fmts.size === 1) {
        const fmt = [...fmts][0];
        if (fmt !== undefined) {
          colMeta.push(`  ${yamlStr(name)}: { numfmt: ${yamlStr(fmt)} }`);
          for (let rr = bodyR1; rr <= bodyR2; rr++) delete cellsByRef.get(`${numToCol(a.col + i)}${rr}`)?.props?.numfmt;
        }
      }
    });
    if (colMeta.length) meta.push('cols:', ...colMeta);

    const rows = [];
    const lastPayloadRow = totals ? b.row - 1 : b.row;
    for (let rr = a.row; rr <= lastPayloadRow; rr++) {
      const rowCells = [];
      for (let cc = a.col; cc <= b.col; cc++) {
        const cell = cellsByRef.get(`${numToCol(cc)}${rr}`);
        if (!cell) { rowCells.push(''); continue; }
        let text;
        if (cell.formula !== null) text = `=${cell.formula}${cell.cachedText !== null ? ` :: ${cell.cachedText}` : ''}`;
        else text = scalarTextOf(cell);
        rowCells.push(String(text).replace(/\|/g, '\\|'));
      }
      rows.push(`| ${rowCells.join(' | ')} |`);
    }
    const out = [`\`\`\`{table} ${t.attrs.name} at ${numToCol(a.col)}${a.row}`, ...meta, '---', ...rows, '```', ''];
    // leftover per-cell props inside the table → @ annotations
    for (let rr = a.row; rr <= b.row; rr++) for (let cc = a.col; cc <= b.col; cc++) {
      const cell = cellsByRef.get(`${numToCol(cc)}${rr}`);
      const keys = Object.keys(cell?.props ?? {});
      if (keys.length) out.push(`@ ${numToCol(cc)}${rr} { ${keys.map((k) => `${k}: ${yamlPropVal(cell.props[k])}`).join(', ')} }`);
    }
    return out;
  }

  function cfRuleToYaml(rule) {
    const fmt = rule.attrs.dxfId !== undefined ? st.dxfProps(Number(rule.attrs.dxfId)) : null;
    const tail = [];
    if (fmt && Object.keys(fmt).length) tail.push(`  format: { ${Object.entries(fmt).map(([k, v]) => `${k}: ${yamlPropVal(v)}`).join(', ')} }`);
    if (rule.attrs.stopIfTrue === '1') tail.push('  stop: true');
    const formulas = all(rule, 'formula').map(textOf);
    const operand = (f) => /^-?\d+(\.\d+)?$/.test(f) ? f : /^".*"$/.test(f) ? f.slice(1, -1).replace(/""/g, '"') : `=${f}`;
    switch (rule.attrs.type) {
      case 'cellIs': {
        const op = rule.attrs.operator;
        if (op === 'between' || op === 'notBetween') {
          return [`- when: ${yamlStr(`${op === 'notBetween' ? 'not-between' : 'between'} ${operand(formulas[0])} and ${operand(formulas[1])}`)}`, ...tail];
        }
        return [`- when: ${yamlStr(`${CELLIS_REV.get(op)} ${operand(formulas[0])}`)}`, ...tail];
      }
      case 'containsText': return [`- contains: ${yamlStr(rule.attrs.text ?? '')}`, ...tail];
      case 'notContainsText': return [`- not-contains: ${yamlStr(rule.attrs.text ?? '')}`, ...tail];
      case 'beginsWith': return [`- begins: ${yamlStr(rule.attrs.text ?? '')}`, ...tail];
      case 'endsWith': return [`- ends: ${yamlStr(rule.attrs.text ?? '')}`, ...tail];
      case 'timePeriod': return [`- date: ${TIME_PERIOD_REV.get(rule.attrs.timePeriod) ?? 'today'}`, ...tail];
      case 'duplicateValues': return ['- dupes: true', ...tail];
      case 'uniqueValues': return ['- unique: true', ...tail];
      case 'top10': {
        const key = rule.attrs.bottom === '1' ? 'bottom' : 'top';
        const val = `${rule.attrs.rank ?? 10}${rule.attrs.percent === '1' ? '%' : ''}`;
        return [`- ${key}: ${rule.attrs.percent === '1' ? yamlStr(val) : val}`, ...tail];
      }
      case 'aboveAverage': {
        const below = rule.attrs.aboveAverage === '0';
        const equal = rule.attrs.equalAverage === '1';
        const lines = [`- avg: ${below ? 'below' : 'above'}${equal ? '-equal' : ''}`];
        if (rule.attrs.stdDev) lines.push(`  stddev: ${rule.attrs.stdDev}`);
        return [...lines, ...tail];
      }
      case 'dataBar': {
        const bar = one(rule, 'dataBar');
        const color = one(bar, 'color') && colorFrom(one(bar, 'color'), themeColors);
        return [`- bars: { color: ${yamlStr(color ?? '#638EC6')} }`, ...tail];
      }
      case 'colorScale': {
        const cs = one(rule, 'colorScale');
        const colors = all(cs, 'color').map((c) => yamlStr(colorFrom(c, themeColors) ?? '#FFFFFF'));
        return [`- scale: [${colors.join(', ')}]`, ...tail];
      }
      case 'iconSet': {
        const is = one(rule, 'iconSet');
        const name = ICON_REV.get(is?.attrs.iconSet ?? '3Arrows') ?? '3-arrows';
        const lines = [`- icons: ${name}`];
        if (is?.attrs.reverse === '1') lines.push('  reverse: true');
        if (is?.attrs.showValue === '0') lines.push('  icons-only: true');
        return [...lines, ...tail];
      }
      case 'expression': return [`- formula: =${formulas[0] ?? ''}`, ...tail];
      default:
        report.push({ feature: `cf rule type ${rule.attrs.type}`, action: 'carried', note: 'unmapped rule type skipped' });
        return null;
    }
  }

  function validationToBlock(v) {
    const TYPE_REV = { list: 'list', whole: 'whole', decimal: 'decimal', date: 'date', time: 'time', textLength: 'text-length', custom: 'custom' };
    const out = [`\`\`\`{validation} ${v.attrs.sqref}`, `type: ${TYPE_REV[v.attrs.type] ?? 'custom'}`];
    const f1 = textOf(one(v, 'formula1') ?? { children: [], text: '' });
    const f2 = textOf(one(v, 'formula2') ?? { children: [], text: '' });
    if (v.attrs.type === 'list') {
      if (/^".*"$/.test(f1)) out.push(`values: [${f1.slice(1, -1).split(',').map((x) => yamlStr(x)).join(', ')}]`);
      else out.push(`source: =${f1}`);
      if (v.attrs.showDropDown === '1') out.push('dropdown: false');
    } else if (v.attrs.type === 'custom') {
      out.push(`formula: =${f1}`);
    } else if (v.attrs.operator === 'between' || v.attrs.operator === 'notBetween' || (!v.attrs.operator && f2)) {
      if (v.attrs.operator === 'notBetween') out.push('op: not-between');
      out.push(`min: ${yamlMaybeNum(f1)}`, `max: ${yamlMaybeNum(f2)}`);
    } else if (f1) {
      const OP_REV = { equal: '=', notEqual: '<>', greaterThan: '>', greaterThanOrEqual: '>=', lessThan: '<', lessThanOrEqual: '<=' };
      if (v.attrs.operator && OP_REV[v.attrs.operator]) out.push(`op: "${OP_REV[v.attrs.operator]}"`);
      out.push(`value: ${yamlMaybeNum(f1)}`);
    }
    if (v.attrs.allowBlank !== '1') out.push('blank: false');
    if (v.attrs.promptTitle || v.attrs.prompt) out.push(`input: { title: ${yamlStr(v.attrs.promptTitle ?? '')}, message: ${yamlStr(v.attrs.prompt ?? '')} }`);
    if (v.attrs.errorTitle || v.attrs.error) out.push(`error: { style: ${v.attrs.errorStyle ?? 'stop'}, title: ${yamlStr(v.attrs.errorTitle ?? '')}, message: ${yamlStr(v.attrs.error ?? '')} }`);
    out.push('```', '');
    return out;
  }

  function filterToBlock(af, sortState) {
    const t = af.attrs.ref.split(':').map(parseCell);
    const out = [`\`\`\`{filter} ${af.attrs.ref}`];
    const cols = [];
    for (const fc of all(af, 'filterColumn')) {
      const letter = numToCol(t[0].col + Number(fc.attrs.colId));
      const filters = one(fc, 'filters');
      const custom = one(fc, 'customFilters');
      const top10 = one(fc, 'top10');
      if (filters) cols.push(`  ${letter}: { values: [${all(filters, 'filter').map((f) => yamlStr(f.attrs.val)).join(', ')}] }`);
      else if (top10) cols.push(`  ${letter}: { ${top10.attrs.bottom === '1' ? 'bottom' : 'top'}: ${top10.attrs.percent === '1' ? yamlStr(`${top10.attrs.val}%`) : top10.attrs.val} }`);
      else if (custom) {
        const cf = one(custom, 'customFilter');
        const OP_REV = { equal: '=', notEqual: '<>', greaterThan: '>', greaterThanOrEqual: '>=', lessThan: '<', lessThanOrEqual: '<=' };
        if (cf && OP_REV[cf.attrs.operator]) cols.push(`  ${letter}: { op: "${OP_REV[cf.attrs.operator]}", value: ${yamlMaybeNum(cf.attrs.val)} }`);
      }
    }
    if (cols.length) out.push('cols:', ...cols);
    if (sortState) {
      const levels = all(sortState, 'sortCondition').map((scond) => {
        const colLetter = /^([A-Z]{1,3})/.exec(scond.attrs.ref)?.[1];
        return colLetter ? `  - { col: ${colLetter}, order: ${scond.attrs.descending === '1' ? 'desc' : 'asc'}, by: value }` : null;
      }).filter(Boolean);
      if (levels.length) out.push('sort:', ...levels);
    }
    out.push('```', '');
    return out;
  }

  function sparklineToBlock(g) {
    const lines = all(one(g, 'sparklines'), 'sparkline').map((sp) => ({
      f: textOf(one(sp, 'f')), sqref: textOf(one(sp, 'sqref')),
    }));
    if (!lines.length) return [];
    const targets = lines.map((l) => l.sqref).join(' ');
    const first = parseCell(lines[0].sqref);
    const last = parseCell(lines[lines.length - 1].sqref);
    const sqref = lines.length > 1 ? `${lines[0].sqref}:${lines[lines.length - 1].sqref}` : lines[0].sqref;
    // reconstruct a contiguous source range when the slices tile a block
    const srcParts = lines.map((l) => l.f.split('!'));
    const srcSheet = srcParts[0][0];
    const firstRange = srcParts[0][1].split(':').map(parseCell);
    const lastRange = srcParts[srcParts.length - 1][1].split(':').map(parseCell);
    const source = `${srcSheet}!${numToCol(firstRange[0].col)}${firstRange[0].row}:${numToCol(lastRange[1].col)}${lastRange[1].row}`;
    const out = [`\`\`\`{sparklines} ${sqref}`];
    if (g.attrs.type === 'column') out.push('type: column');
    else if (g.attrs.type === 'stacked') out.push('type: win-loss');
    else out.push('type: line');
    out.push(`source: ${source}`);
    const marks = [];
    for (const [xmlKey, key] of [['high', 'high'], ['low', 'low'], ['first', 'first'], ['last', 'last'], ['negative', 'negative']]) {
      if (g.attrs[xmlKey] === '1') marks.push(`${key}: true`);
    }
    if (marks.length) out.push(`markers: { ${marks.join(', ')} }`);
    const cs = one(g, 'colorSeries');
    const color = cs && colorFrom(cs, themeColors);
    if (color) out.push(`color: ${yamlStr(color)}`);
    out.push('```', '');
    return out;
  }

  function pageToBlock(ws, pn) {
    const setup = one(ws, 'pageSetup');
    const hfEl = one(ws, 'headerFooter');
    const po = one(ws, 'printOptions');
    if (!setup && !hfEl && !po && !pn) return [];
    const out = ['```{page}'];
    if (setup?.attrs.orientation) out.push(`orientation: ${setup.attrs.orientation}`);
    if (setup?.attrs.paperSize && PAPER_REV.has(Number(setup.attrs.paperSize))) out.push(`paper: ${PAPER_REV.get(Number(setup.attrs.paperSize))}`);
    const margins = one(ws, 'pageMargins');
    if (margins) {
      const cm = (v) => Math.round(Number(v) * 2.54 * 100) / 100;
      out.push(`margins: { top: ${cm(margins.attrs.top)}, bottom: ${cm(margins.attrs.bottom)}, left: ${cm(margins.attrs.left)}, right: ${cm(margins.attrs.right)}, header: ${cm(margins.attrs.header)}, footer: ${cm(margins.attrs.footer)} }`);
    }
    if (setup?.attrs.scale && setup.attrs.scale !== '100') out.push(`scale: ${setup.attrs.scale}`);
    else if (setup?.attrs.fitToWidth !== undefined || setup?.attrs.fitToHeight !== undefined) {
      out.push(`fit: { width: ${setup.attrs.fitToWidth ?? 1}, height: ${setup.attrs.fitToHeight ?? 1} }`);
    }
    if (pn?.area) {
      const range = pn.area.split('!').pop().replace(/\$/g, '');
      out.push(`print-area: ${range}`);
    }
    if (pn?.titles) {
      const bits = [];
      for (const part of pn.titles.split(',')) {
        const spec = part.split('!').pop().replace(/\$/g, '');
        if (/^\d+:\d+$/.test(spec)) bits.push(`rows: "${spec}"`);
        else bits.push(`cols: "${spec}"`);
      }
      if (bits.length) out.push(`print-titles: { ${bits.join(', ')} }`);
    }
    const hfParse = (s) => {
      const seg = { left: '', center: '', right: '' };
      let cur = 'center';
      for (const piece of s.split(/&(?=[LCR])/)) {
        if (piece.startsWith('L')) { cur = 'left'; seg.left += piece.slice(1); }
        else if (piece.startsWith('C')) { cur = 'center'; seg.center += piece.slice(1); }
        else if (piece.startsWith('R')) { cur = 'right'; seg.right += piece.slice(1); }
        else seg[cur] += piece;
      }
      return seg;
    };
    for (const [el, key] of [[one(hfEl ?? { children: [] }, 'oddHeader'), 'header'], [one(hfEl ?? { children: [] }, 'oddFooter'), 'footer']]) {
      if (!el) continue;
      const seg = hfParse(textOf(el));
      const bits = Object.entries(seg).filter(([, v]) => v !== '').map(([k, v]) => `${k}: ${yamlStr(v)}`);
      if (bits.length) out.push(`${key}: { ${bits.join(', ')} }`);
    }
    if (po?.attrs.gridLines === '1') out.push('gridlines: true');
    if (po?.attrs.headings === '1') out.push('headings: true');
    if (po?.attrs.horizontalCentered === '1' || po?.attrs.verticalCentered === '1') {
      out.push(`center: { horizontal: ${po.attrs.horizontalCentered === '1'}, vertical: ${po.attrs.verticalCentered === '1'} }`);
    }
    out.push('```', '');
    return out;
  }

  function readPersons() {
    const map = new Map();
    const doc = xml('xl/persons/person.xml');
    if (doc) for (const p of all(doc, 'person')) map.set(p.attrs.id, p.attrs.displayName);
    return map;
  }
}

// ================= styles =================
function readStyles(doc, themeColors) {
  const numFmts = new Map(BUILTIN_NUMFMTS);
  const fonts = [];
  const fills = [];
  const borders = [];
  const xfs = [];
  const dxfs = [];
  if (doc) {
    for (const nf of all(one(doc, 'numFmts') ?? { children: [] }, 'numFmt')) {
      numFmts.set(Number(nf.attrs.numFmtId), nf.attrs.formatCode);
    }
    for (const f of all(one(doc, 'fonts') ?? { children: [] }, 'font')) fonts.push(fontProps(f, themeColors));
    for (const f of all(one(doc, 'fills') ?? { children: [] }, 'fill')) fills.push(fillProps(f, themeColors));
    for (const b of all(one(doc, 'borders') ?? { children: [] }, 'border')) borders.push(borderProps(b, themeColors));
    for (const x of all(one(doc, 'cellXfs') ?? { children: [] }, 'xf')) xfs.push(x);
    for (const d of all(one(doc, 'dxfs') ?? { children: [] }, 'dxf')) dxfs.push(d);
  }
  return {
    propsForXf(idx) {
      const x = xfs[idx];
      if (!x || idx === 0) return {};
      const props = {};
      const numFmtId = Number(x.attrs.numFmtId ?? 0);
      if (numFmtId !== 0 && numFmts.has(numFmtId)) props.numfmt = numFmts.get(numFmtId);
      Object.assign(props, fonts[Number(x.attrs.fontId ?? 0)] ?? {});
      Object.assign(props, fills[Number(x.attrs.fillId ?? 0)] ?? {});
      Object.assign(props, borders[Number(x.attrs.borderId ?? 0)] ?? {});
      const al = one(x, 'alignment');
      if (al) {
        const H = { left: 'left', center: 'center', right: 'right', justify: 'justify', fill: 'fill', centerContinuous: 'center-across', distributed: 'distributed' };
        const V = { top: 'top', center: 'middle', bottom: 'bottom', justify: 'justify', distributed: 'distributed' };
        if (al.attrs.horizontal && H[al.attrs.horizontal]) props.align = H[al.attrs.horizontal];
        if (al.attrs.vertical && V[al.attrs.vertical]) props.valign = V[al.attrs.vertical];
        if (al.attrs.wrapText === '1') props.wrap = true;
        if (al.attrs.shrinkToFit === '1') props.shrink = true;
        if (al.attrs.indent) props.indent = Number(al.attrs.indent);
        if (al.attrs.textRotation) props.rotation = Number(al.attrs.textRotation) === 255 ? 'vertical' : Number(al.attrs.textRotation);
      }
      const pr = one(x, 'protection');
      if (pr?.attrs.locked === '0') props.locked = false;
      if (pr?.attrs.hidden === '1') props.hidden = true;
      return props;
    },
    dxfProps(idx) {
      const d = dxfs[idx];
      if (!d) return {};
      const props = {};
      const f = one(d, 'font');
      if (f) Object.assign(props, fontProps(f, themeColors, true));
      const fill = one(d, 'fill');
      const pf = fill && one(fill, 'patternFill');
      const bg = pf && (one(pf, 'bgColor') ?? one(pf, 'fgColor'));
      const c = bg && colorFrom(bg, themeColors);
      if (c) props.fill = c;
      return props;
    },
  };
}

function fontProps(f, themeColors, isDxf = false) {
  const props = {};
  if (one(f, 'b')) props.bold = true;
  if (one(f, 'i')) props.italic = true;
  if (one(f, 'strike')) props.strike = true;
  const u = one(f, 'u');
  if (u) props.underline = u.attrs.val === 'double' ? 'double' : u.attrs.val === 'singleAccounting' ? 'single-accounting' : u.attrs.val === 'doubleAccounting' ? 'double-accounting' : true;
  const va = one(f, 'vertAlign');
  if (va?.attrs.val === 'subscript') props.sub = true;
  if (va?.attrs.val === 'superscript') props.super = true;
  const sz = one(f, 'sz');
  if (sz && Number(sz.attrs.val) !== 11) props.size = Number(sz.attrs.val);
  const nm = one(f, 'name');
  if (nm && nm.attrs.val !== 'Calibri') props.font = nm.attrs.val;
  const col = one(f, 'color');
  const c = col && colorFrom(col, themeColors);
  if (c) props.color = c;
  return props;
}

function fillProps(f, themeColors) {
  const pf = one(f, 'patternFill');
  if (!pf || !pf.attrs.patternType || pf.attrs.patternType === 'none' || pf.attrs.patternType === 'gray125') return {};
  const fg = one(pf, 'fgColor');
  const c = fg && colorFrom(fg, themeColors);
  return c ? { fill: c } : {};
}

function borderProps(b, themeColors) {
  const props = {};
  for (const [el, key] of [['top', 'border-top'], ['bottom', 'border-bottom'], ['left', 'border-left'], ['right', 'border-right']]) {
    const e = one(b, el);
    if (e?.attrs.style) {
      const style = BORDER_STYLE_REV.get(e.attrs.style) ?? 'thin';
      const col = one(e, 'color');
      const c = col && colorFrom(col, themeColors);
      props[key] = c ? `${style} ${c}` : style;
    }
  }
  return props;
}

function readSharedStrings(doc) {
  if (!doc) return [];
  return all(doc, 'si').map((si) => {
    const runs = all(si, 'r');
    if (runs.length > 1 || (runs.length === 1 && one(runs[0], 'rPr'))) {
      return {
        rich: runs.map((run) => {
          const pr = one(run, 'rPr');
          const item = { text: textOf(one(run, 't')) };
          if (pr) {
            if (one(pr, 'b')) item.bold = true;
            if (one(pr, 'i')) item.italic = true;
          }
          return item;
        }),
      };
    }
    return { text: textOf(si) };
  });
}

function readTheme(doc) {
  const colors = { ...DEFAULT_THEME };
  if (!doc) return colors;
  const scheme = findDeep(doc, 'clrScheme');
  if (!scheme) return colors;
  const ORDER = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  for (const slot of ORDER) {
    const el = one(scheme, slot);
    if (!el) continue;
    const srgb = one(el, 'srgbClr');
    const sys = one(el, 'sysClr');
    if (srgb?.attrs.val) colors[slot] = srgb.attrs.val.toUpperCase();
    else if (sys?.attrs.lastClr) colors[slot] = sys.attrs.lastClr.toUpperCase();
  }
  return colors;
}

// styles <color> element → GridMD color string.
function colorFrom(el, themeColors) {
  if (!el) return null;
  if (el.attrs.auto === '1') return null;
  if (el.attrs.rgb) {
    const argb = el.attrs.rgb.toUpperCase();
    if (argb.length === 8) {
      return argb.startsWith('FF') ? `#${argb.slice(2)}` : `#${argb.slice(2)}${argb.slice(0, 2)}`;
    }
    return `#${argb}`;
  }
  if (el.attrs.theme !== undefined) {
    const slot = THEME_SLOTS[Number(el.attrs.theme)];
    if (!slot) return null;
    const tint = el.attrs.tint !== undefined ? Math.round(Number(el.attrs.tint) * 100) : 0;
    return tint ? `${slot}@${tint}` : slot;
  }
  return null;
}

// ================= scalars & YAML emission =================
const NUMBER_RE = /^-?(0|[1-9]\d*)(\.\d+)?([eE][+-]?\d+)?$/;
const DATEISH_RE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?)?$|^\d{2}:\d{2}(:\d{2})?$/;
const ERRORS = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!']);

function quoteText(s, forceQuote) {
  if (s === '') return '""';
  const needs = forceQuote || /^['="{]/.test(s) || /[|]/.test(s) || s.includes(' :: ')
    || s.endsWith('}') || /^\s|\s$/.test(s) || NUMBER_RE.test(s) || DATEISH_RE.test(s)
    || /^(true|false)$/i.test(s) || ERRORS.has(s.toUpperCase());
  return needs && !s.includes('\n') ? `"${s.replace(/"/g, '""')}"` : s;
}

function unquote(s) {
  return /^".*"$/s.test(s) ? s.slice(1, -1).replace(/""/g, '"') : s;
}

function yamlStr(s) {
  const str = String(s);
  if (/^[A-Za-z0-9][A-Za-z0-9 _./()+-]*$/.test(str) && !/^(true|false|null|yes|no|on|off)$/i.test(str) && !NUMBER_RE.test(str) && !/^\s|\s$/.test(str)) return str;
  return `'${str.replace(/'/g, "''")}'`;
}

const yamlMaybeNum = (s) => NUMBER_RE.test(String(s)) ? String(s) : yamlStr(s);

function yamlPropVal(v) {
  if (typeof v === 'boolean' || typeof v === 'number') return String(v);
  return yamlStr(v);
}

function bodyValue(scalarText) {
  const s = String(scalarText);
  if (NUMBER_RE.test(s) || /^(TRUE|FALSE)$/.test(s)) return s.toLowerCase() === s ? s : (s === 'TRUE' ? 'true' : s === 'FALSE' ? 'false' : s);
  return yamlStr(unquote(s));
}

function classifyName(value) {
  if (/^(?:'[^']+'|[A-Za-z_][A-Za-z0-9_. ]*)!/.test(value)) return 'ref';
  if (/^[-\d"'{[]/.test(value)) return 'value';
  return 'formula';
}

function isDateFormat(code) {
  const stripped = String(code).replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
  return /[ymdh]/i.test(stripped) && !/[#0?]/.test(stripped);
}

function serialToIso(serial, dateSystem, code) {
  if (!Number.isFinite(serial) || serial < 0) return null;
  if (dateSystem === 1900 && Math.floor(serial) === 60) return null; // phantom day: keep numeric
  const frac = serial - Math.floor(serial);
  let days = Math.floor(serial);
  let dateStr = null;
  if (days > 0) {
    let diff = days;
    let base;
    if (dateSystem === 1904) base = Date.UTC(1904, 0, 1);
    else { base = Date.UTC(1899, 11, 30); if (days < 60) diff = days + 1; }
    const d = new Date(base + diff * DAY_MS);
    dateStr = `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
  }
  const hasTime = frac > 1e-9;
  if (!hasTime) return dateStr;
  const totalSec = Math.round(frac * 86400);
  const hh = String(Math.floor(totalSec / 3600)).padStart(2, '0');
  const mm = String(Math.floor((totalSec % 3600) / 60)).padStart(2, '0');
  const ss = totalSec % 60;
  const time = ss ? `${hh}:${mm}:${String(ss).padStart(2, '0')}` : `${hh}:${mm}`;
  return dateStr ? `${dateStr}T${time}` : time;
}
