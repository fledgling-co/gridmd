import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { lint } from '../src/index';
import { buildWorkbookModel, translateFormula } from '../src/xlsx/model';
import { writeXlsx } from '../src/xlsx/write';
import { zipWrite, zipRead, crc32 } from '../src/xlsx/zip';
import { isoToSerial, pxToColWidth, resolveColor } from '../src/xlsx/units';
import { resolveDataRef, splitTopLevel } from '../src/xlsx/chart';
import type { TableIndexEntry } from '../src/types';

const examplePath = fileURLToPath(new URL('../../examples/quarterly-report.gmd', import.meta.url));

const convert = (src: string, baseDir = '.') => {
  const res = lint(src, { mode: 'strict' });
  assert.deepEqual(res.errors, [], 'document must lint clean before conversion');
  return writeXlsx(buildWorkbookModel(res.doc, { baseDir }));
};

// ---- units ----

test('isoToSerial: 1900 system incl. the phantom leap day', () => {
  assert.equal(isoToSerial('1900-01-01'), 1);
  assert.equal(isoToSerial('1900-02-28'), 59);
  assert.equal(isoToSerial('1900-03-01'), 61); // serial 60 = phantom 1900-02-29
  assert.equal(isoToSerial('2026-07-04'), 46207);
  assert.equal(isoToSerial('12:00'), 0.5);
  assert.equal(isoToSerial('2026-07-04T06:00'), 46207.25);
  assert.equal(isoToSerial('1904-01-02', 1904), 1);
});

test('px→width and colors', () => {
  assert.equal(pxToColWidth(96), 13);
  assert.equal(resolveColor('#1F3FA6'), 'FF1F3FA6');
  assert.equal(resolveColor('accent1', { accent1: '1F3FA6' }), 'FF1F3FA6');
  assert.equal(resolveColor('accent1@-100', { accent1: '1F3FA6' }), 'FF000000');
  assert.equal(resolveColor('accent1@100', { accent1: '1F3FA6' }), 'FFFFFFFF');
});

// ---- zip ----

test('crc32 reference value', () => {
  assert.equal(crc32(Buffer.from('123456789')), 0xcbf43926);
});

test('zip write → read round-trip', () => {
  const buf = zipWrite([
    { name: 'a.txt', data: 'hello' },
    { name: 'dir/b.bin', data: Buffer.from([1, 2, 3]) },
  ]);
  const back = zipRead(buf);
  assert.equal(back.get('a.txt')!.toString(), 'hello');
  assert.deepEqual([...back.get('dir/b.bin')!], [1, 2, 3]);
});

// ---- chart data references ----

test('resolveDataRef: structured refs resolve to absolute A1 ranges', () => {
  const tableIndex = new Map<string, TableIndexEntry>([['sales', {
    name: 'Sales', sheetName: 'Sales', anchor: { col: 1, row: 4 },
    columns: ['id', 'product', 'qty'], headerRow: true, bodyRows: 7,
    total: null, style: undefined, banded: 'rows', filter: null, sort: [], line: 1,
  }]]);
  assert.equal(resolveDataRef('Sales[product]', 'Summary', tableIndex), 'Sales!$B$5:$B$11');
  assert.equal(resolveDataRef('Data!B2:G4', 'Summary', tableIndex), 'Data!$B$2:$G$4');
  assert.equal(resolveDataRef('A1:A5', 'Summary', tableIndex), 'Summary!$A$1:$A$5');
});

test('splitTopLevel respects brackets', () => {
  assert.deepEqual(splitTopLevel('Sales[product], Sales[total]'), ['Sales[product]', 'Sales[total]']);
  assert.deepEqual(splitTopLevel('T[[#Totals],[x]], A1:B2'), ['T[[#Totals],[x]]', 'A1:B2']);
});

// ---- relative fill translation ----

test('translateFormula shifts relative refs only', () => {
  assert.equal(translateFormula('A2*1.1', 3, 0), 'A5*1.1');
  assert.equal(translateFormula('$A$2+B2', 1, 1), '$A$2+C3');
  assert.equal(translateFormula('LOG10(A1)', 1, 0), 'LOG10(A2)');
  assert.equal(translateFormula('"A1 "&A1', 2, 0), '"A1 "&A3');
  assert.equal(translateFormula('Sheet2!A1+[@qty]', 1, 0), 'Sheet2!A2+[@qty]');
});

// ---- synthetic conversions ----

test('dates serialize as serials with a date numfmt', () => {
  const { buffer } = convert(`---\ngridmd: "0.1"\n---\n\n# S1\n\n@ A1 2026-07-04\n`);
  const parts = zipRead(buffer);
  const sheet = parts.get('xl/worksheets/sheet1.xml')!.toString();
  assert.match(sheet, /<c r="A1" s="\d+"><v>46207<\/v><\/c>/);
  assert.match(parts.get('xl/styles.xml')!.toString(), /yyyy-mm-dd/);
});

test('relative fill emits translated formulas per cell', () => {
  const { buffer } = convert(`---\ngridmd: "0.1"\n---\n\n# S1\n\n@ B1:B3 =A1*2\n`);
  const sheet = zipRead(buffer).get('xl/worksheets/sheet1.xml')!.toString();
  assert.match(sheet, /<c r="B2"[^>]*><f>A2\*2<\/f><\/c>/);
  assert.match(sheet, /<c r="B3"[^>]*><f>A3\*2<\/f><\/c>/);
});

test('scenarios + standalone filter + sortState emit in order', () => {
  const src = `---\ngridmd: "0.1"\n---\n\n# S1\n\n`
    + '```{grid} A1\n| a | b |\n| 1 | 9 |\n| 2 | 8 |\n```\n\n'
    + '```{filter} A1:B3\ncols:\n  B: { op: ">", value: 5 }\nsort:\n  - { col: B, order: desc, by: value }\n```\n\n'
    + '```{scenario} Up\ncells: { A2: 5 }\n```\n';
  const { buffer } = convert(src);
  const sheet = zipRead(buffer).get('xl/worksheets/sheet1.xml')!.toString();
  assert.match(sheet, /<scenarios><scenario name="Up"[^>]*><inputCells r="A2" val="5"\/><\/scenario><\/scenarios>/);
  assert.match(sheet, /<autoFilter ref="A1:B3"><filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="5"\/>/);
  assert.match(sheet, /<sortState ref="A2:B3"><sortCondition ref="B2:B3" descending="1"\/><\/sortState>/);
  const idx = (re: RegExp): number => sheet.search(re);
  assert.ok(idx(/<scenarios>/) < idx(/<autoFilter/), 'scenarios before autoFilter');
  assert.ok(idx(/<autoFilter/) < idx(/<sortState/), 'autoFilter before sortState');
});

// ---- the worked example: full-feature emission ----

test('example converts with the FULL feature set emitted', () => {
  const src = readFileSync(examplePath, 'utf8');
  const { buffer, report } = convert(src, dirname(examplePath));
  const parts = zipRead(buffer);

  // Package skeleton — worksheet core AND the full feature parts.
  for (const name of ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml',
    'xl/styles.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet4.xml',
    'xl/tables/table1.xml', 'xl/tables/table2.xml',
    'xl/charts/chart1.xml', 'xl/charts/chart2.xml',
    'xl/drawings/drawing1.xml', 'xl/chartsheets/sheet1.xml',
    'xl/pivotTables/pivotTable1.xml', 'xl/pivotCache/pivotCacheDefinition1.xml',
    'xl/pivotCache/pivotCacheRecords1.xml',
    'xl/slicers/slicer1.xml', 'xl/slicerCaches/slicerCache1.xml',
    'xl/media/image1.png', 'xl/threadedComments/threadedComment1.xml',
    'xl/persons/person.xml', 'xl/comments1.xml',
    'customXml/item1.xml', 'customXml/gridmdCarry1.xml']) {
    assert.ok(parts.has(name), `missing part ${name}`);
  }

  const wb = parts.get('xl/workbook.xml')!.toString();
  assert.match(wb, /<sheet name="Summary"/);
  assert.match(wb, /<sheet name="Data" state="hidden"/);
  assert.match(wb, /<sheet name="Revenue Chart"/); // chart sheet now emitted
  assert.match(wb, /<pivotCaches><pivotCache cacheId="1"/);
  assert.match(wb, /x15:slicerCaches/);
  assert.match(wb, /definedName name="FtoC"[^>]*>LAMBDA\(F,\(F-32\)\*5\/9\)/);

  // Summary: combo chart with trendline/error-bars/secondary axis; sparklines.
  const chart1 = parts.get('xl/charts/chart1.xml')!.toString();
  assert.match(chart1, /<c:barChart>/);
  assert.match(chart1, /<c:lineChart>/);
  assert.match(chart1, /<c:f>Sales!\$B\$5:\$B\$11<\/c:f>/); // Sales[product] resolved
  assert.match(chart1, /<c:trendlineType val="linear"\/>/);
  assert.match(chart1, /<c:forward val="1"\/>/);
  assert.match(chart1, /<c:dispRSqr val="1"\/><c:dispEq val="1"\/>/);
  assert.match(chart1, /<c:errBars>/);
  assert.match(chart1, /<c:max val="700"\/>/);
  assert.match(chart1, /<c:majorUnit val="100"\/>/);
  assert.match(chart1, /<c:legendPos val="b"\/>/);
  assert.equal((chart1.match(/<c:valAx>/g) ?? []).length, 2, 'primary + secondary value axes');

  const summary = parts.get('xl/worksheets/sheet1.xml')!.toString();
  assert.match(summary, /x14:sparklineGroup/);
  assert.match(summary, /<xm:f>Data!B2:G2<\/xm:f><xm:sqref>C14<\/xm:sqref>/);
  assert.match(summary, /<drawing r:id="rId\d+"\/>/);
  assert.match(summary, /<f t="array" ref="B9:B11">SORT\(UNIQUE\(Sales\[region\]\)\)<\/f>/);

  // Chart sheet: chart2 plots the table by region.
  const chart2 = parts.get('xl/charts/chart2.xml')!.toString();
  assert.match(chart2, /<c:f>Sales!\$C\$5:\$C\$11<\/c:f>/); // Sales[region]
  const chartsheet = parts.get('xl/chartsheets/sheet1.xml')!.toString();
  assert.match(chartsheet, /<drawing r:id="rId1"\/>/);

  // Pivot: cache + definition wired to the Sales table.
  const cacheDef = parts.get('xl/pivotCache/pivotCacheDefinition1.xml')!.toString();
  assert.match(cacheDef, /refreshOnLoad="1"/);
  assert.match(cacheDef, /<worksheetSource name="Sales"\/>/);
  assert.match(cacheDef, /<cacheField name="region"/);
  const pt = parts.get('xl/pivotTables/pivotTable1.xml')!.toString();
  assert.match(pt, /name="RevenueByRegion" cacheId="1"/);
  assert.match(pt, /<rowFields count="1"><field x="2"\/><\/rowFields>/);   // region
  assert.match(pt, /<colFields count="1"><field x="3"\/><\/colFields>/);   // quarter
  assert.match(pt, /<dataField name="Sum of total" fld="6"/);

  // Slicer: table slicer cache bound to Sales.product.
  const slicerCache = parts.get('xl/slicerCaches/slicerCache1.xml')!.toString();
  assert.match(slicerCache, /x15:tableSlicerCache tableId="1" column="2"/);
  const sales = parts.get('xl/worksheets/sheet2.xml')!.toString();
  assert.match(sales, /<tablePart r:id=/);
  assert.match(sales, /cfRule type="dataBar"/);
  assert.match(summary, /x14:slicerList/); // the slicer is anchored on Summary
  assert.match(parts.get('xl/drawings/drawing1.xml')!.toString(), /sle:slicer/); // frame in the drawing

  // Table sort state now emitted.
  const table1 = parts.get('xl/tables/table1.xml')!.toString();
  assert.match(table1, /<sortState ref="A5:I11"><sortCondition ref="G5:G11" descending="1"\/><\/sortState>/);

  // Threaded comments + persons.
  const tc = parts.get('xl/threadedComments/threadedComment1.xml')!.toString();
  assert.match(tc, /<threadedComment ref="G8"[^>]*done="1">/);
  assert.match(tc, /parentId=/);
  assert.match(parts.get('xl/persons/person.xml')!.toString(), /displayName="Priya N"/);

  // Scenario on Assumptions.
  const assumptions = parts.get('xl/worksheets/sheet3.xml')!.toString();
  assert.match(assumptions, /<scenario name="Downside"[^>]*comment="FX -12%, volume shrinks 10\.?%?\.?"/);
  assert.match(assumptions, /<inputCells r="B3" val="0.58"\/>/);

  // Fidelity: NOTHING is silently dropped — no not-emitted entries at all,
  // and the only carried items are the four no-OOXML-form features + raw.
  const notEmitted = report.filter((r) => r.action === 'not-emitted');
  assert.deepEqual(notEmitted, [], `expected zero not-emitted, got: ${JSON.stringify(notEmitted)}`);
  const carried = report.filter((r) => r.action === 'carried').map((r) => r.feature);
  for (const expected of ['{query} FxRates', '{script} tidy-headers', 'control: checkbox at B5', 'entity cell at B6', '{raw} customXml/item1.xml']) {
    assert.ok(carried.includes(expected), `carry list should include ${expected}`);
  }
  assert.match(parts.get('customXml/gridmdCarry1.xml')!.toString(), /gridmdCarry/);
});
