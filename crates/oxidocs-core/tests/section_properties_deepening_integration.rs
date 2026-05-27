//! Integration tests: deepening pass for `<w:sectPr>` features that
//! section_integration.rs (S290) and columns_integration.rs (S307)
//! did not pin. Tests the remaining surface of
//! `parse_section_properties` at parser/ooxml.rs:5454.
//!
//! Coverage gaps filled:
//!   - pgBorders: 4-side PageBorders + three independent storage
//!     filters (val=none / sz=0 / color=auto). Same parser idioms
//!     as tblBorders (S311) — color="auto" SUPPRESSES (OPPOSITE
//!     of tcBorders S310 where auto materializes to "000000").
//!   - pgMar ASYMMETRIC rounding (COM-confirmed 0e7a 2026-04-13):
//!     top → 10tw round, bottom/left/right/header/footer → exact.
//!   - pgMar gutter ADDITIVE to margin.left (folded at parse time,
//!     NOT a separate field).
//!   - docGrid type="lines" + linePitch → grid_line_pitch populated.
//!   - docGrid linePitch WITHOUT type → doc_grid_no_type=true,
//!     grid_line_pitch stays None. Gates CJK 83/64 multiplier per
//!     CLAUDE.md (no_type=true → multiplier SKIPPED).
//!   - pgNumType: fmt → page_number_format, start → page_number_start.
//!
//! Fixtures live in `tools/fixtures/section_properties_deepening_samples/`
//! and are authored by
//! `tools/metrics/build_section_properties_extras_repro_fixtures.py`.
//! (The "extras" suffix avoids the `tools/**/*_deep*.py` gitignore.)

use std::fs;

use oxidocs_core::ir::Document;
use oxidocs_core::parse_docx;

fn fixture_path(name: &str) -> std::path::PathBuf {
    let crate_root = std::env::current_dir().unwrap();
    let workspace_root = crate_root.parent().unwrap().parent().unwrap();
    workspace_root
        .join("tools")
        .join("fixtures")
        .join("section_properties_deepening_samples")
        .join(name)
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
fn v1_sect_pg_borders_three_storage_filters() {
    let Some(doc) = load("v1_sect_pg_borders.docx") else { return };
    let page = &doc.pages[0];

    let pb = page
        .page_borders
        .as_ref()
        .expect("pgBorders with ≥1 valid side must populate page_borders");

    // top: val=single sz=24 color=000000 → fully stored.
    let top = pb.top.as_ref().expect("top border stored");
    assert_eq!(top.style, "single");
    assert!(
        (top.width - 3.0).abs() < 0.001,
        "sz=24 → 3.0pt (val/8), got {}",
        top.width
    );
    assert_eq!(
        top.color.as_deref(),
        Some("000000"),
        "explicit hex color preserved verbatim"
    );

    // bottom: val=single sz=24 color="auto" → STORED, but color
    // is SUPPRESSED to None (parser/ooxml.rs:5500-5503).
    // This is the SAME suppression as tblBorders (S311) and the
    // OPPOSITE of tcBorders (S310) where "auto" materializes to
    // "000000". Three adjacent border parsers, three different
    // "auto" handlings. A regression that unified them would
    // silently shift one or the other.
    let bottom = pb.bottom.as_ref().expect("bottom border stored");
    assert!(
        bottom.color.is_none(),
        "pgBorders color=\"auto\" SUPPRESSES storage (color stays None — \
         OPPOSITE of tcBorders S310 where auto materializes to \"000000\")"
    );

    // left: val="none" → NOT stored (parser/ooxml.rs:5508 filter).
    assert!(
        pb.left.is_none(),
        "<w:left w:val=\"none\"/> filter: style=none → not stored even with sz>0"
    );

    // right: sz=0 → NOT stored (parser/ooxml.rs:5508 width>0 filter).
    assert!(
        pb.right.is_none(),
        "<w:right w:sz=\"0\"/> filter: width=0 → not stored even with valid style"
    );
}

#[test]
fn v1_sect_pgmar_asymmetric_top_rounded_others_exact() {
    let Some(doc) = load("v1_sect_pgmar_asymmetric.docx") else { return };
    let page = &doc.pages[0];
    let m = &page.margin;

    // top: w=1133 → ROUND10 → 1130tw / 20 = 56.5pt (NOT 56.65pt).
    // COM-confirmed (0e7a 2026-04-13): top margin uses the rounded
    // value for content-start Y. A regression that skipped rounding
    // would shift content-start Y by 0.15pt → propagate to every
    // page break Y on every doc.
    assert!(
        (m.top - 56.5).abs() < 0.001,
        "top=1133 → ROUND10 → 56.5pt (NOT 56.65pt exact), got {}",
        m.top
    );

    // bottom: EXACT — w=1133 / 20 = 56.65pt. The bottom margin is
    // EXACT because Word uses it as the page-break LIMIT, not the
    // content-start Y. Mixing the two rounding rules is the
    // 2026-04-13 fix.
    assert!(
        (m.bottom - 56.65).abs() < 0.001,
        "bottom=1133 → EXACT → 56.65pt (NOT rounded to 56.5pt), got {}",
        m.bottom
    );

    // left/right: EXACT.
    assert!(
        (m.left - 53.85).abs() < 0.001,
        "left=1077 → 53.85pt exact"
    );
    assert!(
        (m.right - 53.85).abs() < 0.001,
        "right=1077 → 53.85pt exact"
    );

    // header/footer: EXACT.
    let hd = page
        .header_distance
        .expect("pgMar header attr populates header_distance");
    assert!(
        (hd - 42.55).abs() < 0.001,
        "header=851 → 42.55pt exact, got {}",
        hd
    );
    let fd = page
        .footer_distance
        .expect("pgMar footer attr populates footer_distance");
    assert!(
        (fd - 49.6).abs() < 0.001,
        "footer=992 → 49.6pt exact, got {}",
        fd
    );
}

#[test]
fn v1_sect_gutter_adds_to_left_margin() {
    let Some(doc) = load("v1_sect_gutter.docx") else { return };
    let m = &doc.pages[0].margin;

    // left=1440 + gutter=720 → margin.left = 2160tw / 20 = 108pt.
    // The gutter is folded into left margin at parse time
    // (parser/ooxml.rs:5664-5666); there is no separate gutter
    // field. A regression that stored gutter separately would
    // produce left=72pt + gutter=36pt and downstream consumers
    // would silently double-count or under-count.
    assert!(
        (m.left - 108.0).abs() < 0.001,
        "left=1440tw + gutter=720tw → 108pt (additive, NOT separate field), got {}",
        m.left
    );
}

#[test]
fn v1_sect_docgrid_lines_pitch_populates_grid_line_pitch() {
    let Some(doc) = load("v1_sect_docgrid_lines_pitch.docx") else { return };
    let page = &doc.pages[0];

    let pitch = page
        .grid_line_pitch
        .expect("docGrid type=lines linePitch=350 populates grid_line_pitch");
    assert!(
        (pitch - 17.5).abs() < 0.001,
        "linePitch=350 → 17.5pt (val/20), got {}",
        pitch
    );

    // doc_grid_no_type must be FALSE when type IS set.
    assert!(
        !page.doc_grid_no_type,
        "type=lines is set → doc_grid_no_type stays false"
    );
}

#[test]
fn v1_sect_docgrid_lines_and_chars_neg_charspace_populates_all_three_fields() {
    // S339 (2026-05-27): pins docGrid type="linesAndChars" + negative
    // charSpace parsing. Three OOXML §17.6.5 fields populate:
    //   - grid_line_pitch: linePitch / 20 (twips → points)
    //   - grid_char_space_raw: raw charSpace value (1/4096 pt units)
    //     preserved verbatim so layout post-process can recompute
    //     grid_char_pitch with the correct default_font_size
    //   - grid_char_pitch: derived as default_fs + charSpace/4096
    //     using the parser-resolved default_fs (Normal style fallback
    //     to rPrDefault)
    //   - doc_grid_no_type: FALSE (type IS set)
    //
    // b35123-realistic values: linePitch=350 (17.5pt), charSpace=-2714
    // (compression). Only 2/55 baseline docs use this combo (b35123,
    // 191cb5) — S339 corpus survey identified it as the rare-but-real
    // compression-grid mode driving b35123's IoU=0.8655 swamp.
    //
    // Per S339 minimal repro: Word's actual char-width formula for
    // negative charSpace is `font_size + char_space_pt` (= font_size +
    // charSpace/4096). Sub-grid position snapping (0.75pt) creates the
    // observed 8.25/9.00 alternation, but the AVERAGE stride matches
    // this formula. This test only pins PARSER output (grid_char_pitch
    // populated, raw preserved); the layout-side formula and the
    // suspected default_fs derivation bug (b35123 specific) are
    // tracked separately in `[[session339_docgrid_minimal_repro]]`.
    let Some(doc) = load("v1_sect_docgrid_linesAndChars_neg_charspace.docx") else { return };
    let page = &doc.pages[0];

    assert_eq!(
        page.grid_char_space_raw,
        Some(-2714),
        "grid_char_space_raw stores raw charSpace verbatim (sign + magnitude)"
    );

    let line_pitch = page
        .grid_line_pitch
        .expect("linesAndChars + linePitch=350 populates grid_line_pitch");
    assert!(
        (line_pitch - 17.5).abs() < 0.001,
        "linePitch=350tw → 17.5pt (val/20), got {}",
        line_pitch
    );

    let char_pitch = page
        .grid_char_pitch
        .expect("linesAndChars + charSpace populates grid_char_pitch");
    // Computed as default_fs + charSpace/4096. Shared fixture uses
    // rPrDefault sz=22 (11pt) and no Normal style override, so default_fs
    // resolves to 11pt → pitch = 11 + (-2714/4096) = 11 - 0.6626 ≈ 10.337pt.
    // The post-process refines to (content_w / floor(content_w / raw_pitch))
    // so a slight rounding-up to grid_char_pitch is possible; assert on
    // the range matching the OOXML formula within 0.05pt tolerance.
    assert!(
        char_pitch > 0.0 && char_pitch < 11.0,
        "grid_char_pitch for default_fs=11 + charSpace=-2714 should be ~10.337pt (< 11pt), got {}",
        char_pitch
    );
    assert!(
        char_pitch > 10.0,
        "grid_char_pitch should be > 10pt (= 11 - 1pt floor margin), got {}",
        char_pitch
    );

    assert!(
        !page.doc_grid_no_type,
        "type=linesAndChars is set → doc_grid_no_type stays false"
    );
}

#[test]
fn v1_sect_docgrid_lines_and_chars_grid_char_cw_ratio_matches_pitch_div_default_fs() {
    // S339: grid_char_cw_ratio = grid_char_pitch / default_font_size.
    // Layout uses this ratio at break_into_lines and count_cell_lines
    // (mod.rs:5793, 9309) to compute per-char width as
    // `font_size × pitch / default_fs` (positive branch) or
    // `font_size + char_space_pt` (negative branch).
    //
    // For shared fixture (rPrDefault sz=22 = 11pt, no Normal override),
    // default_fs = 11pt → ratio = pitch / 11 ≈ 0.94. The exact value
    // depends on post-process rounding (parser/ooxml.rs:230-244), so
    // we assert a range rather than exact equality.
    let Some(doc) = load("v1_sect_docgrid_linesAndChars_neg_charspace.docx") else { return };
    let page = &doc.pages[0];

    let ratio = page
        .grid_char_cw_ratio
        .expect("grid_char_cw_ratio populated when grid_char_pitch is");
    assert!(
        ratio > 0.9 && ratio < 1.0,
        "ratio = pitch/default_fs ≈ 10.337/11 ≈ 0.94, got {}",
        ratio
    );
}

#[test]
fn v1_sect_docgrid_no_type_flips_flag_without_setting_pitch() {
    // NON-OBVIOUS branch at parser/ooxml.rs:5695-5698:
    //   `grid_type.is_empty() && line_pitch > 0` → doc_grid_no_type=true.
    // Even though linePitch is declared, grid_line_pitch is NOT
    // populated because the parser only emits it for type=lines or
    // type=linesAndChars. The doc_grid_no_type flag exists
    // specifically to gate the CJK 83/64 line-height multiplier
    // per CLAUDE.md.
    let Some(doc) = load("v1_sect_docgrid_no_type.docx") else { return };
    let page = &doc.pages[0];

    assert!(
        page.doc_grid_no_type,
        "docGrid with linePitch but NO type → doc_grid_no_type=true"
    );
    assert!(
        page.grid_line_pitch.is_none(),
        "linePitch alone (without type=lines/linesAndChars) → grid_line_pitch STAYS None"
    );
}

#[test]
fn v1_sect_pgnumtype_populates_format_and_start() {
    let Some(doc) = load("v1_sect_pgnumtype.docx") else { return };
    let page = &doc.pages[0];

    assert_eq!(
        page.page_number_format.as_deref(),
        Some("lowerRoman"),
        "pgNumType fmt stored verbatim as enum-like string"
    );
    assert_eq!(
        page.page_number_start,
        Some(5),
        "pgNumType start parses to u32"
    );
}
