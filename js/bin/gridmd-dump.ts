#!/usr/bin/env node
// Emits the canonical conformance model dump (conformance/README.md) for a
// .gmd document. All language implementations must produce identical output.

import { readFileSync } from 'node:fs';
import { dirname } from 'node:path';
import { lint } from '../src/index';
import { buildWorkbookModel } from '../src/xlsx/model';
import { dumpModel } from '../src/dump';

const input = process.argv[2];
if (!input) {
  console.error('usage: gridmd-dump <file.gmd>');
  process.exit(2);
}
const res = lint(readFileSync(input, 'utf8'), { mode: 'strict' });
if (res.errors.length) {
  for (const e of res.errors) console.error(`${input}:${e.line}: error: ${e.msg}`);
  process.exit(1);
}
process.stdout.write(dumpModel(buildWorkbookModel(res.doc, { baseDir: dirname(input) })));
