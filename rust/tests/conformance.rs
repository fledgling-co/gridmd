//! The three conformance laws (conformance/README.md) over every valid fixture
//! and the example workbook, plus invalid rejection.

use gridmd::model::build_model;
use gridmd::parser::Mode;
use gridmd::xlsx::{write_xlsx, xlsx_to_gridmd};
use gridmd::{dump_source, lint};
use std::path::PathBuf;

fn repo(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("..").join(rel)
}

fn read(rel: &str) -> String {
    std::fs::read_to_string(repo(rel)).unwrap_or_else(|e| panic!("read {rel}: {e}"))
}

const VALID: &[(&str, &str)] = &[
    ("conformance/fixtures/01-cells.gmd", "conformance/expected/01-cells.json"),
    ("conformance/fixtures/02-structure.gmd", "conformance/expected/02-structure.json"),
    ("conformance/fixtures/03-features.gmd", "conformance/expected/03-features.json"),
    ("examples/quarterly-report.gmd", "conformance/expected/quarterly-report.json"),
];

#[test]
fn law1_dump_is_byte_identical() {
    for (src, expected) in VALID {
        let dump = dump_source(&read(src)).unwrap_or_else(|e| panic!("{src} failed to dump: {e:?}"));
        let want = read(expected);
        assert_eq!(dump, want, "dump mismatch for {src}");
    }
}

#[test]
fn law2_invalid_fixtures_are_rejected() {
    for name in ["bad-table-headers", "duplicate-cell", "orphan-spill-cache"] {
        let src = read(&format!("conformance/invalid/{name}.gmd"));
        let res = dump_source(&src);
        assert!(res.is_err(), "{name} should be rejected");
        assert!(!res.unwrap_err().is_empty(), "{name} should carry >=1 error");
    }
}

#[test]
fn law3_round_trip_is_dump_stable() {
    for (src, _) in VALID {
        let source = read(src);
        let res = lint(&source, Mode::Strict);
        assert!(res.errors.is_empty(), "{src} should lint clean");
        let model = build_model(&res.doc);
        let xlsx = write_xlsx(&model, &source);
        let (back, _report) = xlsx_to_gridmd(&xlsx).expect("import");
        let orig = dump_source(&source).unwrap();
        let round = dump_source(&back).unwrap_or_else(|e| panic!("{src} round-trip dump: {e:?}"));
        assert_eq!(orig, round, "round-trip dump mismatch for {src}");
        // The importer's output must itself pass strict lint (CLI contract).
        assert!(lint(&back, Mode::Strict).errors.is_empty(), "{src} import must lint clean");
    }
}

#[test]
fn native_import_of_foreign_xlsx_is_lint_clean() {
    // The JS-written example xlsx has no GridMD carry part → exercises the
    // native (DEFLATE + worksheet) reverse-parser.
    let bytes = std::fs::read(repo("conformance/fixtures-xlsx/quarterly-report.xlsx")).unwrap();
    let (gmd, report) = xlsx_to_gridmd(&bytes).expect("native import");
    assert!(report.iter().any(|r| r.action == "imported"));
    let res = lint(&gmd, Mode::Strict);
    assert!(res.errors.is_empty(), "native import must lint clean: {:?}", res.errors);
    assert_eq!(res.doc.sheets.len(), 5);
}
