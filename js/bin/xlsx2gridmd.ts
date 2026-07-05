#!/usr/bin/env node
// XLSX → GridMD importer. Worksheet core imported natively; parts not yet
// reverse-parsed (charts, drawings, pivots, slicers, media) are carried as
// {raw} blocks — loud, never silent.

import { readFileSync, writeFileSync } from 'node:fs';
import { xlsxToGridmd } from '../src/xlsx/read';
import { lint } from '../src/index';

const argv = process.argv.slice(2);
const files = argv.filter((a) => !a.startsWith('-'));
const oIdx = argv.indexOf('-o');
const outArg = oIdx !== -1 ? argv[oIdx + 1] : null;
const input = files.filter((f) => f !== outArg)[0];

if (!input) {
  console.error('usage: xlsx2gridmd <file.xlsx> [-o out.gmd]');
  process.exit(2);
}

const { gmd, report } = xlsxToGridmd(readFileSync(input));
for (const r of report) console.log(`${input}: ${r.action}: ${r.feature}${r.note ? ` (${r.note})` : ''}`);

// The importer's output must itself be valid GridMD — self-check before writing.
const res = lint(gmd, { mode: 'strict' });
for (const e of res.errors) console.error(`self-check:${e.line}: error: ${e.msg}`);
const out = outArg ?? input.replace(/\.(xlsx|xlsm)$/, '') + '.gmd';
writeFileSync(out, gmd);
console.log(`${out}: written — ${res.sheets} sheet(s), ${res.cells} defined cell(s); self-check ${res.errors.length === 0 ? 'clean' : `FAILED (${res.errors.length} error(s))`}`);
process.exit(res.errors.length ? 1 : 0);
