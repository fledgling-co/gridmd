// The round-trip harness: example.gmd → .xlsx → .gmd′ where .gmd′ lints
// clean and the worksheet-core semantic model matches the original.

import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { lint } from '../src/index';
import { buildWorkbookModel } from '../src/xlsx/model';
import { writeXlsx } from '../src/xlsx/write';
import { xlsxToGridmd } from '../src/xlsx/read';
import { dumpModel } from '../src/dump';
import { refKey } from '../src/refs';
import type { CellContent, Scalar, TableModel } from '../src/types';

const examplePath = fileURLToPath(new URL('../../examples/quarterly-report.gmd', import.meta.url));

function modelOf(source: string, baseDir = '.') {
  const res = lint(source, { mode: 'strict' });
  assert.deepEqual(res.errors, [], 'document must lint clean');
  return buildWorkbookModel(res.doc, { baseDir });
}

const scalarKey = (s: Scalar | null | undefined): string | null => {
  if (!s) return null;
  if (s.kind === 'number') return `n:${Math.round((s.value as number) * 1e9) / 1e9}`;
  if (s.kind === 'boolean') return `b:${s.value}`;
  if (s.kind === 'error') return `e:${s.value}`;
  if (s.kind === 'date' || s.kind === 'time') return `d:${s.value}`;
  return `t:${s.value}`;
};

const contentKey = (c: CellContent | null | undefined): string | null => {
  if (!c) return null;
  if (c.rich) return `rich:${c.rich.map((r: { text: string }) => r.text).join('')}`;
  if (c.formula !== undefined && c.formula !== null) {
    return `f:${c.formula}|${scalarKey(c.cached)}|${c.arrayRef ?? ''}`;
  }
  return scalarKey(c.scalar);
};

test('round-trip: gmd → xlsx → gmd′ lints clean and the core model matches', () => {
  const src = readFileSync(examplePath, 'utf8');
  const model1 = modelOf(src, dirname(examplePath));
  const { buffer } = writeXlsx(model1);

  const { gmd: gmd2, report } = xlsxToGridmd(buffer);

  // 1. The importer's output is valid strict-mode GridMD.
  const res2 = lint(gmd2, { mode: 'strict' });
  for (const e of res2.errors) console.error(`  gmd2:${e.line}: ${e.msg}`);
  assert.deepEqual(res2.errors, [], 'imported GridMD must lint clean');

  const model2 = buildWorkbookModel(res2.doc, { baseDir: '.' });

  // 2. Sheet names + order + KINDS survive (chart sheet stays a chart sheet).
  assert.deepEqual(model2.sheets.map((s) => s.name), model1.sheets.map((s) => s.name));
  assert.deepEqual(model2.sheets.map((s) => s.kind), model1.sheets.map((s) => s.kind));

  // 3. Cell-level content round-trips on every worksheet.
  for (const s1 of model1.sheets) {
    if (s1.kind === 'chart') continue;
    const s2 = model2.sheets.find((x) => x.name === s1.name);
    assert.ok(s2, `sheet ${s1.name} survives`);
    for (const [key, cell1] of s1.cells) {
      const want = contentKey(cell1.content);
      if (want === null) continue;
      const cell2 = s2.cells.get(key);
      assert.ok(cell2?.content, `${s1.name}!${key}: content survives`);
      assert.equal(contentKey(cell2.content), want, `${s1.name}!${key}: content matches`);
    }
    // No content invented either.
    for (const [key, cell2] of s2.cells) {
      if (contentKey(cell2.content) !== null) {
        assert.ok(s1.cells.get(key)?.content, `${s2.name}!${key}: no invented content`);
      }
    }
    // Structure counts.
    assert.equal(s2.merges.length, s1.merges.length, `${s1.name}: merges`);
    assert.deepEqual(s2.tables.map((t) => t.name).sort(), s1.tables.map((t) => t.name).sort(), `${s1.name}: tables`);
    for (const t1 of s1.tables) {
      const t2: TableModel = s2.tables.find((t) => t.name === t1.name)!;
      assert.deepEqual(t2.columns, t1.columns, `table ${t1.name}: columns`);
      assert.deepEqual({ col: t2.anchor.col, row: t2.anchor.row }, { col: t1.anchor.col, row: t1.anchor.row }, `table ${t1.name}: anchor`);
      assert.equal(t2.bodyRows, t1.bodyRows, `table ${t1.name}: body rows`);
    }
    assert.equal(s2.cf.length, s1.cf.length, `${s1.name}: cf blocks`);
    assert.equal(
      s2.cf.reduce((n, b) => n + b.rules.length, 0),
      s1.cf.reduce((n, b) => n + b.rules.length, 0),
      `${s1.name}: cf rules`,
    );
    assert.equal(s2.validations.length, s1.validations.length, `${s1.name}: validations`);
    assert.equal(s2.notes.length, s1.notes.length, `${s1.name}: notes`);
    assert.equal(s2.threads.length, s1.threads.length, `${s1.name}: comment threads`);
    assert.equal(s2.scenarios.length, s1.scenarios.length, `${s1.name}: scenarios`);
    assert.equal(s2.sparklines.length, s1.sparklines.length, `${s1.name}: sparkline groups`);
    assert.equal(Boolean(s2.page), Boolean(s1.page), `${s1.name}: page setup presence`);
    assert.equal(s2.meta.freeze, s1.meta.freeze, `${s1.name}: freeze`);
    assert.equal(Boolean(s2.meta.protect?.enabled), Boolean(s1.meta.protect?.enabled), `${s1.name}: protection`);
    // Drawing-hosted objects now import natively.
    assert.equal(s2.charts.length, s1.charts.length, `${s1.name}: charts`);
    assert.equal(s2.pivots.length, s1.pivots.length, `${s1.name}: pivots`);
    assert.equal(s2.slicers.length, s1.slicers.length, `${s1.name}: slicers`);
    assert.equal(s2.images.length, s1.images.length, `${s1.name}: images`);
    assert.equal(s2.shapes.length, s1.shapes.length, `${s1.name}: shapes/textboxes`);
  }
  // Chart meta survives: the Summary combo chart keeps its series + trendline.
  const combo1 = model1.sheets.find((s) => s.name === 'Summary')!.charts[0]!;
  const combo2 = model2.sheets.find((s) => s.name === 'Summary')!.charts[0]!;
  assert.equal(combo2.type, 'combo');
  assert.equal(combo2.meta.series.length, combo1.meta.series.length);
  assert.ok(combo2.meta.series[0].trendline, 'trendline survives');
  assert.ok(combo2.meta.series[0]['error-bars'], 'error bars survive');
  assert.equal(combo2.meta.series[1].axis, 'y2', 'secondary axis survives');
  // Pivot meta survives.
  const piv2 = model2.sheets.find((s) => s.name === 'Summary')!.pivots[0]!;
  assert.equal(piv2.name, 'RevenueByRegion');
  assert.equal(String(piv2.meta.source), 'Sales');
  assert.equal(piv2.meta.rows[0].field, 'region');
  assert.equal(piv2.meta.values[0].agg, 'sum');

  // 4. Workbook names survive.
  const names1 = (model1.fm.names ?? []).map((n: { name: string }) => n.name).sort();
  const names2 = (model2.fm.names ?? []).map((n: { name: string }) => n.name).sort();
  assert.deepEqual(names2, names1, 'defined names survive');

  // 5. Charts/pivots/slicers/media now import NATIVELY — no longer carried.
  const carried = report.filter((r) => r.action === 'carried').map((r) => r.feature);
  for (const gone of ['xl/charts/chart1.xml', 'xl/pivotTables/pivotTable1.xml', 'xl/slicers/slicer1.xml', 'xl/media/image1.png']) {
    assert.ok(!carried.includes(gone), `should be native, not carried: ${gone}`);
  }
  assert.doesNotMatch(gmd2, /```\{raw\} ooxml part="xl\/charts\/chart1\.xml"/);
  assert.match(gmd2, /```\{chart\} combo/);
  assert.match(gmd2, /```\{pivot\} RevenueByRegion/);
  assert.match(gmd2, /```\{slicer\}/);
  assert.match(gmd2, /```\{image\}[\s\S]{0,200}src: data:image\/png;base64,/);
  assert.match(gmd2, /```\{textbox\}/);
});

test('round-trip: threaded comments, notes and scenarios survive with content', () => {
  const src = readFileSync(examplePath, 'utf8');
  const { buffer } = writeXlsx(modelOf(src, dirname(examplePath)));
  const { gmd: gmd2 } = xlsxToGridmd(buffer);
  assert.match(gmd2, /```\{comments\} G8/);
  assert.match(gmd2, /by: Priya N/);
  assert.match(gmd2, /resolved: true/);
  assert.match(gmd2, /note: \|/);
  assert.match(gmd2, /Blended company rate/);
  assert.match(gmd2, /```\{scenario\} Downside/);
  assert.match(gmd2, /```\{sparklines\} C14:C16/);
  assert.match(gmd2, /source: Data!B2:G4/);
  assert.match(gmd2, /```\{spill-cache\} B9/);
});

// Law 3 for every conformance fixture — the cross-language gate runs this
// same loop; keeping it in-suite catches regressions like note-only cells
// being dropped on import (the 02-structure bug).
test('law 3: every conformance fixture round-trips dump-stable', () => {
  const fixtures = ['01-cells', '02-structure', '03-features'];
  for (const name of fixtures) {
    const dir = dirname(fileURLToPath(new URL(`../../conformance/fixtures/${name}.gmd`, import.meta.url)));
    const src = readFileSync(`${dir}/${name}.gmd`, 'utf8');
    const expected = readFileSync(`${dir}/../expected/${name}.json`, 'utf8');
    const model1 = modelOf(src, dir);
    assert.equal(dumpModel(model1), expected, `${name}: baseline dump matches expectation`);
    const { buffer } = writeXlsx(model1);
    const { gmd: gmd2 } = xlsxToGridmd(buffer);
    const res2 = lint(gmd2, { mode: 'strict' });
    assert.deepEqual(res2.errors, [], `${name}: round-tripped gmd lints clean`);
    const model2 = buildWorkbookModel(res2.doc, { baseDir: dir });
    assert.equal(dumpModel(model2), expected, `${name}: round-tripped dump byte-identical`);
  }
});
