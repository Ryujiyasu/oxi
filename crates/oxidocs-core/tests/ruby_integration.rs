//! Integration tests: parse Ruby (furigana) fixtures end-to-end.
//!
//! Fixtures live in `tools/fixtures/ruby_samples/` and are minimal copies of
//! the COM-measurement repros from `tools/metrics/build_ruby_repro_fixtures.py`.
//! Each fixture has one or more paragraphs containing a `<w:ruby>` element;
//! these tests verify that `parser::parse_docx` populates `Run.ruby` with the
//! expected base/text/align/hps/hpsRaise.
//!
//! Companion to `omml_integration.rs` (math) and `comments_fixtures.rs`
//! (tracked-changes balloons).

use std::fs;

use oxidocs_core::ir::{Block, Document, RubyAlign, Run};
use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    // Tests run with CWD at the crate root; fixtures are two levels up.
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root.join("tools").join("fixtures").join("ruby_samples").join(name)
}

/// Walk every paragraph in document order and collect references to runs that
/// carry a `Ruby`. Returned in document order — useful for asserting that the
/// Nth ruby in the file has property X.
fn collect_ruby_runs(doc: &Document) -> Vec<&Run> {
    doc.pages.iter()
        .flat_map(|p| p.blocks.iter())
        .filter_map(|b| if let Block::Paragraph(p) = b { Some(p) } else { None })
        .flat_map(|p| p.runs.iter())
        .filter(|r| r.ruby.is_some())
        .collect()
}

fn load(name: &str) -> Option<Document> {
    let path = fixture_path(name);
    if !path.exists() {
        eprintln!("skipping: {} not found", path.display());
        return None;
    }
    let data = fs::read(&path).expect("read fixture");
    Some(parse_docx(&data).expect("parse fixture"))
}

#[test]
fn v10_minimal_ruby_has_base_text_lang_only() {
    let Some(doc) = load("v10_base_minimal.docx") else { return };
    let rubies: Vec<_> = collect_ruby_runs(&doc).into_iter()
        .map(|r| r.ruby.as_ref().unwrap())
        .collect();
    assert_eq!(rubies.len(), 1, "v10 should have exactly 1 ruby");
    let ru = rubies[0];
    assert_eq!(ru.base, "含");
    assert_eq!(ru.text, "ふく");
    assert_eq!(ru.lang.as_deref(), Some("ja-JP"));
    // Minimal fixture has no rubyPr hps/raise/align — all optional fields None.
    assert!(ru.align.is_none(), "expected no rubyAlign override");
    assert!(ru.hps_halfpt.is_none(), "expected no hps");
    assert!(ru.hps_raise_halfpt.is_none(), "expected no hpsRaise");
    assert!(ru.hps_base_text_halfpt.is_none(), "expected no hpsBaseText");
}

#[test]
fn v11_align_variants_cover_five_modes() {
    // Authored as 5 paragraphs, one per w:rubyAlign value (center,
    // distributeLetter, distributeSpace, left, right) — same base/text in each.
    let Some(doc) = load("v11_align_variants.docx") else { return };
    let rubies: Vec<_> = collect_ruby_runs(&doc).into_iter()
        .map(|r| r.ruby.as_ref().unwrap())
        .collect();
    assert_eq!(rubies.len(), 5, "v11 should have 5 rubies, one per align mode");

    let expected_align = [
        RubyAlign::Center,
        RubyAlign::DistributeLetter,
        RubyAlign::DistributeSpace,
        RubyAlign::Left,
        RubyAlign::Right,
    ];
    for (i, exp) in expected_align.iter().enumerate() {
        let ru = rubies[i];
        assert_eq!(ru.base, "漢字");
        assert_eq!(ru.text, "かん");
        assert_eq!(ru.align, Some(*exp), "ruby {i} should be {exp:?}");
        // Each carries hps=11 (half-points) → 5.5pt derived font_size.
        assert_eq!(ru.hps_halfpt, Some(11));
        assert_eq!(ru.font_size, Some(5.5));
    }
}

#[test]
fn v12_atomic_wrap_attaches_ruby_after_long_base_text() {
    // The fixture's first paragraph has a long Hiragana run followed by a
    // ruby ("特定"/"とくてい") and a trailing text run. Verifies parser tolerates
    // ruby positioned *after* a >30-char text run in the same paragraph.
    let Some(doc) = load("v12_atomic_wrap.docx") else { return };
    let rubies: Vec<_> = collect_ruby_runs(&doc).into_iter()
        .map(|r| r.ruby.as_ref().unwrap())
        .collect();
    assert_eq!(rubies.len(), 1);
    let ru = rubies[0];
    assert_eq!(ru.base, "特定");
    assert_eq!(ru.text, "とくてい");
    assert_eq!(ru.hps_halfpt, Some(11));
    assert_eq!(ru.font_size, Some(5.5));
}

#[test]
fn v13_hps_raise_variants_carry_distinct_metrics() {
    // 4 ruby paragraphs at base=12pt:
    //   P1 default        — hps=12   raise=None
    //   P3 raise=12halfpt — hps=12   raise=12
    //   P5 raise=24halfpt — hps=12   raise=24
    //   P7 hps=24halfpt   — hps=24   raise=None
    let Some(doc) = load("v13_hps_raise_variants.docx") else { return };
    let rubies: Vec<_> = collect_ruby_runs(&doc).into_iter()
        .map(|r| r.ruby.as_ref().unwrap())
        .collect();
    assert_eq!(rubies.len(), 4);
    // (hps_halfpt, raise_halfpt) tuples in document order.
    let expected: [(u32, Option<u32>); 4] = [
        (12, None), (12, Some(12)), (12, Some(24)), (24, None),
    ];
    for (i, (exp_hps, exp_raise)) in expected.iter().enumerate() {
        let ru = rubies[i];
        assert_eq!(ru.hps_halfpt, Some(*exp_hps), "ruby[{i}] hps");
        assert_eq!(ru.hps_raise_halfpt, *exp_raise, "ruby[{i}] hpsRaise");
        // font_size derives from hps (half-pt → pt).
        assert_eq!(ru.font_size, Some(*exp_hps as f32 / 2.0));
    }
}

#[test]
fn v16_raise_sweep_produces_five_monotonic_raises() {
    // 5 ruby paragraphs with raise sweeping 12, 24, 36, 48, 72 half-points
    // (label suffixes r6, r12, r18, r24, r36 in the fixture text — point values).
    let Some(doc) = load("v16_raise_sweep.docx") else { return };
    let rubies: Vec<_> = collect_ruby_runs(&doc).into_iter()
        .map(|r| r.ruby.as_ref().unwrap())
        .collect();
    assert_eq!(rubies.len(), 5, "v16 should have 5 raise-sweep rubies");
    let expected_raise = [12u32, 24, 36, 48, 72];
    for (i, exp) in expected_raise.iter().enumerate() {
        let ru = rubies[i];
        assert_eq!(ru.base, "含");
        assert_eq!(ru.text, "ふく");
        assert_eq!(ru.hps_halfpt, Some(14), "all sweep rubies share hps=14");
        assert_eq!(ru.hps_raise_halfpt, Some(*exp), "ruby[{i}] expected raise={exp}");
    }
}

#[test]
fn all_five_fixtures_parse_without_error() {
    // Smoke test: every committed ruby fixture parses and emits >=1 ruby
    // (or, for any pure-control fixture in the future, simply parses).
    for name in [
        "v10_base_minimal.docx",
        "v11_align_variants.docx",
        "v12_atomic_wrap.docx",
        "v13_hps_raise_variants.docx",
        "v16_raise_sweep.docx",
    ] {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let rubies = collect_ruby_runs(&doc);
        assert!(!rubies.is_empty(), "{} should produce at least 1 Ruby", name);
    }
}
