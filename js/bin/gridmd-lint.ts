#!/usr/bin/env node
// GridMD reference linter — validates .gmd files against SPEC.md (strict mode
// by default; --lenient downgrades unrecognized lines to warnings).

import { readFileSync } from 'node:fs';
import { lint } from '../src/index';

const argv = process.argv.slice(2);
const lenient = argv.includes('--lenient');
const files = argv.filter((a) => !a.startsWith('--'));

if (files.length === 0) {
  console.error('usage: gridmd-lint [--lenient] <file.gmd> …');
  process.exit(2);
}

let failed = false;
for (const file of files) {
  let source: string;
  try {
    source = readFileSync(file, 'utf8');
  } catch (e: unknown) {
    console.error(`${file}: ${e instanceof Error ? e.message : String(e)}`);
    failed = true;
    continue;
  }
  const res = lint(source, { mode: lenient ? 'lenient' : 'strict' });
  for (const w of res.warnings) console.log(`${file}:${w.line}: warning: ${w.msg}`);
  for (const e of res.errors) console.log(`${file}:${e.line}: error: ${e.msg}`);
  console.log(
    `${file}: ${res.errors.length} error(s), ${res.warnings.length} warning(s) — ` +
    `${res.sheets} sheet(s), ${res.cells} defined cell(s), ${res.blocks} block(s)`,
  );
  if (res.errors.length) failed = true;
}
process.exit(failed ? 1 : 0);
