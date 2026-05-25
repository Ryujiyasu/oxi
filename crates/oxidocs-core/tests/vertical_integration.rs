//! Integration tests: parse vertical-writing (w:textDirection="tbRlV") fixtures
//! end-to-end.
//!
//! Fixtures live in `tools/fixtures/vertical_samples/` and are minimal copies of
//! the COM-measurement repros from `tools/metrics/build_vertical_repro_fixtures.py`.
//! Each fixture has a 2-cell table: cell[0] carries `<w:textDirection w:val="tbRlV"/>`
//! in its `<w:tcPr>`; cell[1] is a horizontal filler. These tests verify
//! `parser::parse_docx` populates `TableCell.text_direction` and preserves the
//! run text/font_size/font_family_east_asia we authored.
//!
//! Companion to `ruby_integration.rs` (furigana), `omml_integration.rs` (math),
//! and `comments_fixtures.rs` (tracked-changes balloons).

use std::fs;

use oxidocs_core::ir::{Block, Document, TableCell};
use oxidocs_core::parser::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    // Tests run with CWD at the crate root; fixtures are two levels up.
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root.join("tools").join("fixtures").join("vertical_samples").join(name)
}

/// Walk every table cell in document order. Returned in document order so the
/// Nth cell in the file can be asserted by index.
fn collect_cells(doc: &Document) -> Vec<&TableCell> {
    let mut out = Vec::new();
    for page in &doc.pages {
        for b in &page.blocks {
            if let Block::Table(t) = b {
                for row in &t.rows {
                    for c in &row.cells {
                        out.push(c);
                    }
                }
            }
        }
    }
    out
}

/// Concatenate all run text inside a cell (paragraph order, run order).
fn cell_text(cell: &TableCell) -> String {
    cell.blocks.iter()
        .filter_map(|b| if let Block::Paragraph(p) = b { Some(p) } else { None })
        .flat_map(|p| p.runs.iter())
        .map(|r| r.text.as_str())
        .collect::<String>()
}

/// First run's font_size (Option), useful for size-variant assertions.
fn cell_first_font_size(cell: &TableCell) -> Option<f32> {
    cell.blocks.iter()
        .filter_map(|b| if let Block::Paragraph(p) = b { Some(p) } else { None })
        .flat_map(|p| p.runs.iter())
        .next()
        .and_then(|r| r.style.font_size)
}

/// First run's font_family_east_asia (Option), useful for font-variant assertions.
fn cell_first_font_ea(cell: &TableCell) -> Option<String> {
    cell.blocks.iter()
        .filter_map(|b| if let Block::Paragraph(p) = b { Some(p) } else { None })
        .flat_map(|p| p.runs.iter())
        .next()
        .and_then(|r| r.style.font_family_east_asia.clone())
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
fn v1_basic_cell_carries_tb_rl_v_and_three_char_text() {
    let Some(doc) = load("v1_basic.docx") else { return };
    let cells = collect_cells(&doc);
    assert_eq!(cells.len(), 2, "v1_basic table has 1 row × 2 cells");
    // cell[0] = vertical label
    assert_eq!(cells[0].text_direction.as_deref(), Some("tbRlV"));
    assert_eq!(cell_text(cells[0]), "申請者");
    assert_eq!(cell_first_font_size(cells[0]), Some(10.5));
    assert_eq!(cell_first_font_ea(cells[0]).as_deref(), Some("ＭＳ 明朝"));
    // cell[1] = horizontal filler — no text_direction override
    assert!(cells[1].text_direction.is_none(),
        "filler cell should NOT carry tbRlV");
}

#[test]
fn v1_long_cell_keeps_six_char_sequence_intact() {
    let Some(doc) = load("v1_long.docx") else { return };
    let cells = collect_cells(&doc);
    assert_eq!(cells[0].text_direction.as_deref(), Some("tbRlV"));
    assert_eq!(cell_text(cells[0]), "連絡担当窓口");
    // 6 chars × 10.5pt = 63pt of vertical extent — verifies parser doesn't
    // split or reorder CJK runs in vertical cells.
    assert_eq!(cell_text(cells[0]).chars().count(), 6);
}

#[test]
fn v1_msmincho_14pt_carries_size_override() {
    let Some(doc) = load("v1_msmincho_14pt.docx") else { return };
    let cells = collect_cells(&doc);
    assert_eq!(cells[0].text_direction.as_deref(), Some("tbRlV"));
    assert_eq!(cell_text(cells[0]), "申請者");
    // sz=28 halfpt → 14pt. Filler cell stays at default 10.5pt.
    assert_eq!(cell_first_font_size(cells[0]), Some(14.0),
        "vertical cell should reflect run-level sz=28 halfpt");
    assert_eq!(cell_first_font_size(cells[1]), Some(10.5),
        "filler cell font size unchanged");
}

#[test]
fn v1_yu_mincho_font_family_preserved_in_vertical_cell() {
    let Some(doc) = load("v1_yu_mincho.docx") else { return };
    let cells = collect_cells(&doc);
    assert_eq!(cells[0].text_direction.as_deref(), Some("tbRlV"));
    assert_eq!(cell_first_font_ea(cells[0]).as_deref(), Some("Yu Mincho"),
        "Yu Mincho eastAsia font should round-trip through parser");
    // Filler cell keeps the default ＭＳ 明朝.
    assert_eq!(cell_first_font_ea(cells[1]).as_deref(), Some("ＭＳ 明朝"));
}

#[test]
fn v1_two_cols_preserves_full_28_char_run() {
    let Some(doc) = load("v1_two_cols.docx") else { return };
    let cells = collect_cells(&doc);
    assert_eq!(cells[0].text_direction.as_deref(), Some("tbRlV"));
    let text = cell_text(cells[0]);
    // Parser must NOT collapse / truncate / wrap-break the run text — wrap is
    // a layout concern, not a parser concern.
    assert_eq!(text, "現在登録されている連絡担当窓口情報の更新の有無について確認");
    assert_eq!(text.chars().count(), 29);
}

#[test]
fn all_five_fixtures_parse_with_one_vertical_cell_each() {
    // Smoke test: every committed vertical fixture parses and produces exactly
    // one cell carrying text_direction="tbRlV" plus one horizontal filler.
    for name in [
        "v1_basic.docx",
        "v1_long.docx",
        "v1_msmincho_14pt.docx",
        "v1_yu_mincho.docx",
        "v1_two_cols.docx",
    ] {
        let path = fixture_path(name);
        if !path.exists() {
            eprintln!("skipping {}", name);
            continue;
        }
        let data = fs::read(&path).unwrap();
        let doc = parse_docx(&data)
            .unwrap_or_else(|e| panic!("failed to parse {}: {:?}", name, e));
        let cells = collect_cells(&doc);
        let n_vert = cells.iter()
            .filter(|c| c.text_direction.as_deref() == Some("tbRlV"))
            .count();
        assert_eq!(n_vert, 1, "{} should have exactly 1 tbRlV cell", name);
    }
}
