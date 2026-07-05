#!/usr/bin/env node
// Verifies a document's cached values (` :: ` caches) against the bounded
// formula evaluator. Exit 1 on any mismatch; unsupported formulas are listed
// but never guessed at (SPEC §6 honesty rule).

import { readFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { lint } from '../src/index';
import { buildWorkbookModel } from '../src/xlsx/model';
import { verifyCachedValues } from '../src/calc';

const files = process.argv.slice(2).filter((a) => !a.startsWith('-'));
if (!files.length) {
  console.error('usage: gridmd-calc <file.gmd> …');
  process.exit(2);
}

let failed = false;
for (const file of files) {
  const res = lint(readFileSync(file, 'utf8'), { mode: 'strict' });
  if (res.errors.length) {
    for (const e of res.errors) console.error(`${file}:${e.line}: error: ${e.msg}`);
    failed = true;
    continue;
  }
  const model = buildWorkbookModel(res.doc, { baseDir: dirname(file) });
  const { checked, mismatches, unsupported } = verifyCachedValues(model);
  for (const m of mismatches) {
    console.log(`${file}: MISMATCH ${m.where} =${m.formula} cached ${JSON.stringify(m.cached)} computed ${JSON.stringify(m.computed)}`);
  }
  for (const u of unsupported) {
    console.log(`${file}: unsupported ${u.where} =${u.formula} (${u.reason})`);
  }
  console.log(`${file}: ${checked} cached value(s) verified, ${mismatches.length} mismatch(es), ${unsupported.length} unsupported`);
  if (mismatches.length) failed = true;
}
process.exit(failed ? 1 : 0);
