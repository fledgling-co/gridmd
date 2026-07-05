import { test } from 'bun:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { lint } from '../src/index';

const path = fileURLToPath(new URL('../../examples/quarterly-report.gmd', import.meta.url));

test('the worked example machine-validates with zero errors', () => {
  const res = lint(readFileSync(path, 'utf8'), { mode: 'strict' });
  for (const e of res.errors) console.error(`  error  L${e.line}: ${e.msg}`);
  for (const w of res.warnings) console.error(`  warn   L${w.line}: ${w.msg}`);
  assert.deepEqual(res.errors, []);
  assert.equal(res.sheets, 5); // Summary, Sales, Assumptions, Data, Revenue Chart
  assert.ok(res.cells > 80, `expected a real cell count, got ${res.cells}`);
});
