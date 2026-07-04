# GridMD delivery plan (multi-language)

The durable tracker for Luke's 2026-07-05 directive. Survives context
compaction — keep it current as phases land.

## The directive (verbatim intent)

1. Finish ALL remaining spec/implementation work — no deferred parts
   (ChartEx-family charts, timeline caches, remaining classic chart types,
   chart/pivot/slicer reverse parsing ✅, calc verification ✅).
2. **Two-way converters in Go, Swift, and Rust** alongside the JS/TS one.
3. **Bun** as the JS/TS runtime/bundler/test runner.
4. **100 % test coverage** on all converters, including JS/TS.
5. GitHub **release pipeline**: compile + publish binaries as GitHub Releases.
6. **npm** deployment for the TypeScript lib; **SPM** consumption for Swift
   (requires `Package.swift` at the repo ROOT pointing into `swift/`).
7. Easy clone-and-go setup for anyone.
8. Adhere to `~/Dev/bella-team-files/CODING_PRACTICES.md` +
   `NEW_PROJECT_BEST_PRACTICES.md` where sensible (they are Next/Nest-oriented;
   applicable here: TS boundary safety §1, latest pinned versions, day-one
   quality gate, scripts/tooling conventions, no invented APIs, self-review).
   Divergences (this is a zero-backend polyglot library, not a Next/Nest app)
   are deliberate and noted in per-package READMEs.
9. **Opus + xhigh-effort runners** for the heavy parallel work.
10. ⏸ **PAUSE before the repo restructure step (task #3)** — Luke wants to
    weigh in before the Bun/TS/npm packaging decisions are executed.

## Layout (agreed so far)

```
/                 spec docs (SPEC/DIRECTIVES/FORMATTING/INTEROP/HANDOVER), PLAN.md
conformance/      cross-language fixtures + expected dumps + contract README
examples/         quarterly-report.gmd (+ assets)
js/               reference implementation (currently plain JS on node:test;
                  Bun/TS/npm migration = task #3, PAUSED for Luke's input)
go/               Go module (runner-built)
rust/             Cargo crate (runner-built)
swift/            Swift sources; root Package.swift for SPM (runner-built)
.github/          CI + release workflows (task #8)
```

## Task state (mirror of the session task list)

- [x] #1 Importer reverse-parsing (charts/pivots/slicers/images/shapes native)
- [x] #2 Final spec gaps: ChartEx (cx:), pivot timelines, radar/bubble/stock/-3d, PivotCharts — all native both directions
- [⏸] #3 Bun + TS + npm packaging restructure — PAUSED pending Luke
- [x] #4 Conformance suite (3 laws: dump-match, reject-invalid, round-trip;
      canonical dump = js/src/dump.js; fixtures + expected generated)
- [~] #5 Go port — Opus/xhigh runner LAUNCHED (workflow wf_dfc76715-365)
- [~] #6 Rust port — Opus/xhigh runner LAUNCHED (same workflow)
- [~] #7 Swift port + SPM — Opus/xhigh runner LAUNCHED (same workflow)
- [ ] #8 CI: coverage gates (bun/go/llvm-cov/swift), release binaries on tag,
      npm publish flow, SPM tag — after #3 unpauses

## Port runner contract (tasks #5–#7)

Each runner owns exactly its directory (`go/`, `rust/`, `swift/` + root
`Package.swift` for Swift), snapshots `js/src` + `SPEC.md` + `conformance/` as
reference at start, and must deliver:
- Tier-1 conformance (see conformance/README.md): parse→dump byte-identical,
  invalid fixtures rejected, gmd→xlsx→gmd round-trip dump-stable, {raw} carry
  for anything not natively converted (never silent loss).
- The three CLI verbs: `dump`, `to-xlsx`, `from-xlsx`.
- Idiomatic, zero-heavy-dependency code (allowed: a YAML lib; zip via stdlib
  where available — Go archive/zip ✓, Rust may use `zip`/`quick-xml`/`serde_yaml`
  equivalents, Swift may use its own minimal zip like the JS impl).
- Tests targeting 100 % line coverage of their own code, wired to the
  language's coverage tool; a `make test`-style one-liner documented.
- A README with setup + usage.

## Facts runners/future sessions need

- Toolchains: bun + go installed via brew this session; rust installing;
  swift 6.3.3 present. No Excel on this machine — structural verification only.
- Repo is a fresh local git (branch main, no remote yet). gh CLI available.
- The JS implementation is the semantic reference; conformance/expected/*.json
  are GENERATED from it (never hand-edit).
- The one npm dependency in js/ is `yaml` (safe-subset enforced at parse).
