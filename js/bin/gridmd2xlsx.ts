#!/usr/bin/env node
// GridMD → XLSX transformer (v0: F0/F1 worksheet core; everything else is
// reported loudly, never silently dropped). --strict exits 1 if anything
// could not be emitted natively.

import { readFileSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { lint } from '../src/index';
import { buildWorkbookModel } from '../src/xlsx/model';
import { writeXlsx } from '../src/xlsx/write';

const argv = process.argv.slice(2);
const strict = argv.includes('--strict');
const files = argv.filter((a) => !a.startsWith('-'));
const oIdx = argv.indexOf('-o');
const outArg = oIdx !== -1 ? argv[oIdx + 1] : null;
const input = files.filter((f) => f !== outArg)[0];

if (!input) {
  console.error('usage: gridmd2xlsx <file.gmd> [-o out.xlsx] [--strict]');
  process.exit(2);
}

const source = readFileSync(input, 'utf8');
const res = lint(source, { mode: 'strict' });
if (res.errors.length) {
  for (const e of res.errors) console.error(`${input}:${e.line}: error: ${e.msg}`);
  console.error(`${input}: ${res.errors.length} error(s) — fix the document before converting`);
  process.exit(1);
}

const model = buildWorkbookModel(res.doc, { baseDir: dirname(input) });
const { buffer, report } = writeXlsx(model);
const out = outArg ?? input.replace(/\.gmd$/, '') + '.xlsx';
writeFileSync(out, buffer);

const notEmitted = report.filter((r) => r.action === 'not-emitted');
const carried = report.filter((r) => r.action === 'carried');
const partial = report.filter((r) => r.action === 'partial');
for (const r of report) {
  console.log(`${input}:${r.line}: ${r.action}: ${r.feature}${r.note ? ` (${r.note})` : ''}`);
}
console.log(`${out}: written (${buffer.length} bytes) — ${carried.length} carried, ${partial.length} partial, ${notEmitted.length} not emitted`);
if (notEmitted.length) {
  console.log('fidelity: the features above are NOT in the .xlsx — v0 emits the worksheet core only (see INTEROP.md §2)');
}
process.exit(strict && notEmitted.length ? 1 : 0);
