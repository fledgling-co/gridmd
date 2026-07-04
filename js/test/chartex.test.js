// ChartEx (cx:), 3D/radar/bubble/stock chartML, PivotCharts and pivot
// timelines — emission + full round-trip survival.

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { lint } from '../src/index.js';
import { buildWorkbookModel } from '../src/xlsx/model.js';
import { writeXlsx } from '../src/xlsx/write.js';
import { zipRead } from '../src/xlsx/zip.js';
import { xlsxToGridmd } from '../src/xlsx/read.js';

const DOC = `---
gridmd: "0.1"
title: Chart families
---

# Data

\`\`\`{table} Sales at A1
---
| region | amount | when |
| AU | 40 | 2026-01-05 |
| NZ | 25 | 2026-02-11 |
| UK | 35 | 2026-03-20 |
\`\`\`

\`\`\`{pivot} ByRegion at E1
source: Sales
rows:
  - { field: region }
values:
  - { field: amount, agg: sum }
\`\`\`

\`\`\`{slicer} at H1 size 320x110
kind: timeline
for: ByRegion
field: when
level: months
range: [2026-01-01, 2026-03-31]
\`\`\`

# Charts

\`\`\`{chart} treemap "Share" at A1:F12
series:
  - name: Share
    cat: Sales[region]
    val: Sales[amount]
legend: { position: bottom }
\`\`\`

\`\`\`{chart} histogram at A14:F24
series:
  - name: Spread
    val: Sales[amount]
\`\`\`

\`\`\`{chart} pareto at H1:M12
series:
  - name: Pareto
    cat: Sales[region]
    val: Sales[amount]
\`\`\`

\`\`\`{chart} radar at H14:M24
series:
  - name: R
    cat: Sales[region]
    val: Sales[amount]
\`\`\`

\`\`\`{chart} column-3d at O1:T12
series:
  - name: C3
    cat: Sales[region]
    val: Sales[amount]
\`\`\`

\`\`\`{chart} column "Pivoted" at O14:T24
pivot: ByRegion
\`\`\`
`;

const convert = (src) => {
  const res = lint(src, { mode: 'strict' });
  assert.deepEqual(res.errors, [], 'fixture must lint clean');
  return writeXlsx(buildWorkbookModel(res.doc, { baseDir: '.' }));
};

test('ChartEx, 3D, radar, PivotChart and timeline parts all emit', () => {
  const { buffer, report } = convert(DOC);
  const parts = zipRead(buffer);

  // ChartEx parts: treemap, histogram, pareto.
  assert.ok(parts.has('xl/charts/chartEx1.xml'));
  assert.ok(parts.has('xl/charts/chartEx3.xml'));
  const cx1 = parts.get('xl/charts/chartEx1.xml').toString();
  assert.match(cx1, /<cx:series layoutId="treemap"/);
  assert.match(cx1, /<cx:pt idx="0">AU<\/cx:pt>/); // literal data points
  assert.match(cx1, /<cx:numDim type="val">/);
  const cx3 = parts.get('xl/charts/chartEx3.xml').toString();
  assert.match(cx3, /layoutId="paretoLine"/);

  // Classic additions: radar + 3D column with series axis + view3D.
  const chartFiles = [...parts.keys()].filter((k) => /^xl\/charts\/chart\d+\.xml$/.test(k));
  const all = chartFiles.map((k) => parts.get(k).toString()).join('');
  assert.match(all, /<c:radarChart>/);
  assert.match(all, /<c:bar3DChart>/);
  assert.match(all, /<c:view3D>/);
  assert.match(all, /<c:serAx>/);

  // PivotChart: pivotSource declared.
  assert.match(all, /<c:pivotSource><c:name>/);

  // Timeline parts + workbook/worksheet wiring.
  assert.ok(parts.has('xl/timelineCaches/timelineCache1.xml'));
  assert.ok(parts.has('xl/timelines/timeline1.xml'));
  const tc = parts.get('xl/timelineCaches/timelineCache1.xml').toString();
  assert.match(tc, /sourceName="when"/);
  assert.match(tc, /<pivotTable tabId="1" name="ByRegion"\/>/);
  assert.match(tc, /selection startDate="2026-01-01T00:00:00"/);
  assert.match(parts.get('xl/workbook.xml').toString(), /x15:timelineCacheRefs/);
  assert.match(parts.get('xl/worksheets/sheet1.xml').toString(), /x15:timelineRefs/);
  assert.match(parts.get('xl/drawings/drawing1.xml').toString(), /tsle:timeslicer/);

  // Content types registered.
  const ct = parts.get('[Content_Types].xml').toString();
  assert.match(ct, /chartex\+xml/);
  assert.match(ct, /timeline\+xml/);

  // NOTHING dropped.
  assert.deepEqual(report.filter((r) => r.action === 'not-emitted'), []);
});

test('the new families survive the full round trip', () => {
  const res1 = lint(DOC, { mode: 'strict' });
  const model1 = buildWorkbookModel(res1.doc, { baseDir: '.' });
  const { buffer } = writeXlsx(model1);
  const { gmd: gmd2 } = xlsxToGridmd(buffer);

  const res2 = lint(gmd2, { mode: 'strict' });
  for (const e of res2.errors) console.error(`  gmd2:${e.line}: ${e.msg}`);
  assert.deepEqual(res2.errors, [], 'round-tripped GridMD lints clean');
  const model2 = buildWorkbookModel(res2.doc, { baseDir: '.' });

  for (const s1 of model1.sheets) {
    const s2 = model2.sheets.find((x) => x.name === s1.name);
    assert.equal(s2.charts.length, s1.charts.length, `${s1.name}: chart count`);
    assert.equal(s2.pivots.length, s1.pivots.length, `${s1.name}: pivot count`);
    assert.equal(s2.slicers.length, s1.slicers.length, `${s1.name}: slicer/timeline count`);
  }
  const charts2 = model2.sheets.find((s) => s.name === 'Charts').charts;
  const types2 = charts2.map((c) => c.type).sort();
  assert.ok(types2.includes('treemap'), 'treemap survives');
  assert.ok(types2.includes('histogram'), 'histogram survives');
  assert.ok(types2.includes('pareto'), 'pareto survives');
  assert.ok(types2.includes('radar'), 'radar survives');
  assert.ok(types2.includes('column-3d'), '3d survives');
  const tl2 = model2.sheets.find((s) => s.name === 'Data').slicers.find((x) => x.kind === 'timeline');
  assert.ok(tl2, 'timeline survives');
  assert.equal(tl2.meta.for, 'ByRegion');
  assert.equal(tl2.meta.field, 'when');
});
