// Build the npm package into dist/:
//   1. Bun bundles src/index.ts + the three CLI bins (target node, ESM;
//      `yaml` kept external so consumers dedupe it).
//   2. Node shebangs are (re)applied to the bin outputs.
//   3. tsgo --emitDeclarationOnly (native TS7 compiler) writes the .d.ts surface.

import { rm, readFile, writeFile, chmod } from 'node:fs/promises';
import { join } from 'node:path';

const root = new URL('..', import.meta.url).pathname;
const dist = join(root, 'dist');
const bins = ['gridmd-lint', 'gridmd2xlsx', 'xlsx2gridmd'];

await rm(dist, { recursive: true, force: true });

const check = (label: string, r: Awaited<ReturnType<typeof Bun.build>>): void => {
  if (!r.success) { for (const log of r.logs) console.error(log); throw new Error(`bun build (${label}) failed`); }
};

// The library entry → dist/index.js (single bundled file).
check('index', await Bun.build({
  entrypoints: [join(root, 'src/index.ts')],
  outdir: dist,
  target: 'node',
  format: 'esm',
  external: ['yaml'],
}));

// The CLI bins → dist/bin/*.js (shared code split into dist/bin chunks).
check('bins', await Bun.build({
  entrypoints: bins.map((b) => join(root, `bin/${b}.ts`)),
  outdir: join(dist, 'bin'),
  target: 'node',
  format: 'esm',
  external: ['yaml'],
  splitting: true,
}));

// Node shebang + executable bit on the bin outputs.
const SHEBANG = '#!/usr/bin/env node\n';
for (const b of bins) {
  const file = join(dist, 'bin', `${b}.js`);
  const src = await readFile(file, 'utf8');
  if (!src.startsWith('#!')) await writeFile(file, SHEBANG + src);
  await chmod(file, 0o755);
}

// Type declarations (moduleResolution=bundler forbids JS emit; tsgo emits d.ts only
// emits .d.ts; bun already produced the JS).
const tsc = Bun.spawnSync(['bunx', 'tsgo', '-p', 'tsconfig.build.json'], { cwd: root, stdout: 'inherit', stderr: 'inherit' });
if (tsc.exitCode !== 0) throw new Error('tsgo --emitDeclarationOnly failed');

console.log('build: dist/ ready (bundled JS + .d.ts).');
