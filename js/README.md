# gridmd

The **GridMD** reference implementation: a parser + linter and a two-way
`.gmd` ⇄ `.xlsx` converter, plus a bounded formula evaluator that *verifies*
cached values (SPEC §6 — writers must never fabricate caches). Strict
TypeScript, runs on [Bun](https://bun.sh), published to npm as `gridmd`.

This is the semantic reference for the cross-language conformance suite
(`../conformance/`): its canonical model dump defines the contract every other
implementation (`../go`, `../rust`, `../swift`) must match byte-for-byte.

## Setup

```bash
bun install          # one lockfile (bun.lock), one runtime dep (yaml)
```

Requires Bun ≥ 1.3. Node ≥ 20 is only needed to *run the published package*
(the built `dist/` is plain ESM); development uses Bun to run the `.ts` sources
directly.

## Commands

```bash
bun test                 # run the suite (bunfig gate: 100% line coverage)
bun run test:coverage    # same, with the coverage table
bun run typecheck        # tsc --noEmit (strict) over src, bin, test
bun run build            # bundle dist/ (JS via bun build, .d.ts via tsc)

# dev tools (run the .ts sources directly; paths relative to js/)
bun run lint:example         # lint ../examples/quarterly-report.gmd
bun run xlsx:example         # → ../examples/quarterly-report.xlsx
bun run roundtrip:example    # gmd → xlsx → gmd, then re-lint
bun run dump <file.gmd>      # canonical conformance dump to stdout

# the CLIs directly (dev)
bun bin/gridmd-lint.ts [--lenient] <file.gmd> …
bun bin/gridmd-dump.ts <file.gmd>              # conformance dump
bun bin/gridmd2xlsx.ts <file.gmd> [-o out.xlsx] [--strict]
bun bin/xlsx2gridmd.ts <file.xlsx> [-o out.gmd]
bun bin/gridmd-calc.ts <file.gmd> …            # verify cached values
```

## Library API

```ts
import {
  lint, parseDocument,
  buildWorkbookModel, writeXlsx, xlsxToGridmd, dumpModel,
  verifyCachedValues, createEvaluator,
} from 'gridmd';
import type { WorkbookModel, Sheet, CellContent, Scalar, Diagnostic, LintResult } from 'gridmd';

const { doc, errors } = lint(source);            // parse + validate
const model = buildWorkbookModel(doc);           // → workbook model
const { buffer, report } = writeXlsx(model);     // → .xlsx (Buffer) + fidelity report
const dump = dumpModel(model);                   // → canonical conformance JSON
const { gmd } = xlsxToGridmd(buffer);            // .xlsx → GridMD
const { mismatches, unsupported } = verifyCachedValues(model);
```

The full model/diagnostic shapes are exported as types (`Sheet`, `Cell`,
`CellContent`, `Scalar`, `WorkbookModel`, `ChartModel`, …).

## Coverage gate

The `bunfig.toml` gate enforces **100 % line coverage** (`bun test` fails
otherwise). The suite hits every line of `src/**` — verified via lcov.

**Why line coverage and not function coverage?** Bun's *function*-coverage
metric has a confirmed false negative: it reports 98.8 % functions for
`src/calc.ts` and blames one anonymous arrow it cannot name (lcov emits **no**
`FN`/`FNDA` entry for it). To rule out a real gap, every function and callback
in `calc.ts` was instrumented with a file-based hit marker and the whole suite
run — **all of them fire**. So the code is genuinely 100 %-covered; the 98.8 %
is a Bun instrumentation artifact, not a missing test. Line coverage (the
metric Bun measures correctly, and the one the cross-language contract cares
about) is therefore the gate. This is the one documented coverage exception.

## npm publishing

```jsonc
// package.json (already configured)
"exports": { ".": { "types": "./dist/index.d.ts", "import": "./dist/index.js" } }
"bin":     { "gridmd-lint", "gridmd2xlsx", "xlsx2gridmd" }  // → dist/bin/*.js (node shebang)
"files":   ["dist"]
"prepublishOnly": "bun run build && bun test"
```

`bun run build` produces `dist/index.js` (bundled, `yaml` kept external so
consumers dedupe it), `dist/bin/*.js` (node-runnable CLIs), and the `.d.ts`
type surface. `npm publish` runs `prepublishOnly` (build + full suite) first.
Consumers that touch the `Buffer`-typed APIs (`writeXlsx`, `zipRead/zipWrite`)
need `@types/node` — standard for a Node library.

## CODING_PRACTICES adherence

Applies `~/Dev/bella-team-files/CODING_PRACTICES.md` §1 (TypeScript boundary
safety): explicit return types on exports; strict + `noUncheckedIndexedAccess`
+ `verbatimModuleSyntax`; `catch (e: unknown)` with narrowing; exhaustive
switches with `never`/throw defaults; no `as`-casts on untrusted input; the
GridMD source is validated by the `lint` pass before any model is built.

### Deliberate divergences

This is a zero-backend polyglot library, not a Next/Nest app, so the
Next/Nest-specific rules don't apply. Two intentional, contained divergences:

1. **YAML-derived data is typed as a permissive `Meta` alias (`= any`), not a
   web of speculative interfaces.** The `SPEC` allows `x-` extension keys and
   open per-directive shapes, so directive metadata (`meta`/`props`/`body`,
   frontmatter) enters as parsed YAML and is validated by the `lint` pass, not
   the type system. The *stable* surface — scalars, cells, sheets, the
   workbook model, diagnostics, the block tree — is fully, strictly typed
   (`src/types.ts`), and the byte-identical conformance suite + 100 % line
   coverage are the real validation. Typing the open YAML fully would have
   risked the port's byte-for-byte guarantee for no safety gain. See the header
   of `src/types.ts`.

2. **The theme-tokens/no-`any` rule is relaxed only for that YAML surface**;
   everywhere else (the model, XLSX emit/parse, the evaluator) is `any`-free.

The port's overriding invariant is **behaviour-identical, byte-identical
output**: the four conformance dumps, the example `.xlsx`, and the round-trip
`.gmd` are all verified byte-for-byte against the pre-migration JS reference.
```
