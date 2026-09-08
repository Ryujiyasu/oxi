// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use super::*;

#[test]
fn section_parity_and_number_restart_have_distinct_padding_rules() {
    let cases: &[(&[u8], usize)] = &[
        (include_bytes!("../../../../tests/fixtures/section_parity/oddPage_c1_s2_h0.docx"), 3),
        (include_bytes!("../../../../tests/fixtures/section_parity/oddPage_c1_s2_h1.docx"), 4),
        (include_bytes!("../../../../tests/fixtures/section_parity/evenPage_c1_s1_h0.docx"), 3),
        (include_bytes!("../../../../tests/fixtures/section_parity/evenPage_c2_s1_h1.docx"), 4),
        (include_bytes!("../../../../tests/fixtures/section_parity/oddPage_c1_sNone_h0.docx"), 4),
        (include_bytes!("../../../../tests/fixtures/section_parity/oddPage_c2_sNone_h0.docx"), 3),
        (include_bytes!("../../../../tests/fixtures/section_parity/nextPage_c1_s1_h1.docx"), 4),
        (include_bytes!("../../../../tests/fixtures/section_parity/nextPage_c1_s2_h1.docx"), 3),
    ];
    for (index, (bytes, expected)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let expected_start = [Some(3), Some(3), Some(2), Some(2), None, None, Some(1), Some(2)][index];
        assert_eq!(section_page_number_start(&doc.pages[1]), expected_start);
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), *expected, "case {index}");
        let content_pages: Vec<_> = layout.pages.iter().enumerate().filter_map(|(i, p)| {
            p.elements.iter().any(|e| matches!(&e.content, LayoutContent::Text { text, .. }
                if text == "SECTION2")).then_some(i + 1)
        }).collect();
        assert_eq!(content_pages, vec![expected - 1], "section 2 in case {index}");
    }
}

#[test]
fn document_break_count_distrust_preserves_locally_valid_table_hint() {
    let bytes = include_bytes!(concat!(env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/table_row_cached_break.docx"));
    let doc = crate::parser::parse_docx(bytes).unwrap();
    let engine = LayoutEngine::for_document(&doc);
    let trusted = engine.layout_pass(&doc);
    assert_eq!(trusted.pages.len(), 2);
    engine.lrpb_count_distrust.set(true);
    let distrusted = engine.layout_pass(&doc);
    assert_eq!(distrusted.pages.len(), trusted.pages.len());
    assert!(distrusted.pages[1].elements.iter().any(|e| {
        matches!(&e.content, LayoutContent::Text { text, .. } if text.contains("Second"))
    }));
}

#[test]
fn vertical_pair_compression_does_not_charge_back_structural_spacing() {
    let mut engine = LayoutEngine::new();
    engine.default_font_size = 10.5;
    engine.doc_regime_fs = 10.5;
    engine.compress_punctuation = true;
    let style = RunStyle {
        font_size: Some(10.5),
        font_family_east_asia: Some("メイリオ".into()),
        ..RunStyle::default()
    };
    let field = RunStyle {
        ruby_field: true,
        ..style.clone()
    };
    let small = RunStyle {
        font_size: Some(9.0),
        ..style.clone()
    };
    let mut head = vec!['あ'; 34];
    for (i, ch) in [(6, '。'), (7, '』'), (15, '。'), (31, '。'), (32, '」')] {
        head[i] = ch;
    }
    let head: String = head.into_iter().collect();
    let gloss = "い".repeat(60);
    let fragments = vec![
        (head.as_str(), &style, None, 0, 0),
        ("あ", &field, None, 1, 0),
        ("あああああああ、", &style, None, 2, 0),
        (gloss.as_str(), &small, None, 3, 0),
    ];
    let lines = engine.break_into_lines(
        &fragments,
        425.25,
        0.0,
        &ParagraphStyle::default(),
        None,
        None,
        false,
        true,
        true,
        false,
        false,
        true,
        true,
    );
    assert_eq!(
        lines[0]
            .fragments
            .iter()
            .map(|f| f.text.chars().count())
            .sum::<usize>(),
        43
    );
}

#[test]
fn vertical_breaking_prices_each_runs_own_font_size() {
    let engine = LayoutEngine::new();
    let large = RunStyle {
        font_size: Some(12.0),
        font_family_east_asia: Some("ＭＳ 明朝".into()),
        ..RunStyle::default()
    };
    let small = RunStyle {
        font_size: Some(9.0),
        ..large.clone()
    };
    let fragments = vec![
        ("本文", &large, None, 0, 0),
        ("あいうえおかきく", &small, None, 1, 0),
    ];
    let lines = engine.break_into_lines(
        &fragments,
        90.0,
        0.0,
        &ParagraphStyle::default(),
        None,
        None,
        false,
        true,
        true,
        false,
        false,
        true,
        true,
    );
    assert_eq!(lines.len(), 2);
    assert_eq!(
        lines[0]
            .fragments
            .iter()
            .map(|f| f.text.as_str())
            .collect::<String>(),
        "本文あいうえおかき"
    );
    assert_eq!(
        lines[1]
            .fragments
            .iter()
            .map(|f| f.text.as_str())
            .collect::<String>(),
        "く"
    );
}

fn text_row(x: f32, y: f32, height: f32, index: usize) -> LayoutElement {
    let mut element = LayoutElement::new(
        x,
        y,
        20.0,
        height,
        LayoutContent::Text {
            text: index.to_string(),
            font_size: 10.0,
            font_family: None,
            bold: false,
            italic: false,
            underline: false,
            underline_style: None,
            strikethrough: false,
            double_strikethrough: false,
            color: None,
            highlight: None,
            character_spacing: 0.0,
            field_type: None,
            text_scale: 100.0,
            is_vertical: false,
            effects: TextEffects::default(),
        },
    );
    element.paragraph_index = Some(index);
    element
}

#[test]
fn text_balance_preserves_order_across_existing_columns() {
    let mut elements: Vec<_> = (0..4)
        .map(|i| text_row(0.0, i as f32 * 10.0, 10.0, i))
        .collect();
    elements.push(text_row(100.0, 0.0, 10.0, 4));
    elements.push(text_row(100.0, 10.0, 10.0, 5));
    assert_eq!(
        LayoutEngine::rebalance_text_columns(&mut elements, 0.0, &[0.0, 100.0]),
        Some(30.0)
    );
    assert_eq!(
        elements.iter().map(|e| (e.x, e.y)).collect::<Vec<_>>(),
        vec![
            (0.0, 0.0),
            (0.0, 10.0),
            (0.0, 20.0),
            (100.0, 0.0),
            (100.0, 10.0),
            (100.0, 20.0)
        ]
    );
}

#[test]
fn text_balance_uses_row_advances_and_keeps_ruby_with_base() {
    let mut elements = vec![
        text_row(0.0, 0.0, 30.0, 0),
        text_row(0.0, 30.0, 10.0, 1),
        text_row(0.0, 40.0, 10.0, 2),
        text_row(0.0, 50.0, 10.0, 3),
    ];
    let mut ruby = text_row(2.0, 24.0, 5.0, 1);
    ruby.flow_line_offset = -6.0;
    elements.push(ruby);
    assert_eq!(
        LayoutEngine::rebalance_text_columns(&mut elements, 0.0, &[0.0, 100.0]),
        Some(30.0)
    );
    assert_eq!((elements[1].x, elements[1].y), (100.0, 0.0));
    assert_eq!((elements[4].x, elements[4].y), (102.0, -6.0));
}

#[test]
fn text_balance_does_not_split_a_hanging_line_at_the_column_boundary() {
    let mut elements = vec![text_row(90.0, 0.0, 10.0, 0), text_row(110.0, 0.0, 10.0, 0)];
    assert!(LayoutEngine::rebalance_text_columns(&mut elements, 0.0, &[0.0, 100.0]).is_none());
    assert_eq!((elements[0].y, elements[1].y), (0.0, 0.0));
}

#[test]
fn text_balance_leaves_table_content_for_table_path() {
    let mut elements = vec![text_row(0.0, 0.0, 10.0, 0), text_row(0.0, 10.0, 10.0, 1)];
    elements[1].cell_paragraph_index = Some(0);
    assert!(LayoutEngine::rebalance_text_columns(&mut elements, 0.0, &[0.0, 100.0]).is_none());
    assert_eq!(elements[1].y, 10.0);
}

#[test]
fn text_balance_keeps_page_continuation_before_later_paragraphs() {
    let mut elements = vec![
        text_row(100.0, 0.0, 10.0, 2),
        text_row(0.0, 20.0, 10.0, 3),
        text_row(0.0, 30.0, 10.0, 4),
    ];
    let before: Vec<_> = elements.iter().map(|e| (e.x, e.y)).collect();
    assert!(LayoutEngine::rebalance_text_columns(&mut elements, 0.0, &[0.0, 100.0]).is_none());
    assert_eq!(elements.iter().map(|e| (e.x, e.y)).collect::<Vec<_>>(), before);
}

#[test]
fn table_continuations_use_destination_header_geometry() {
    let cases: &[(&[u8], usize, usize)] = &[
        (include_bytes!("../../../../tests/fixtures/table_page_geometry/even_rows.docx"), 26, 48),
        (include_bytes!("../../../../tests/fixtures/table_page_geometry/even_split.docx"), 26, 48),
        (include_bytes!("../../../../tests/fixtures/table_page_geometry/even_whole.docx"), 26, 46),
        (include_bytes!("../../../../tests/fixtures/table_page_geometry/first_rows.docx"), 20, 46),
        (include_bytes!("../../../../tests/fixtures/table_page_geometry/first_split.docx"), 20, 46),
        (include_bytes!("../../../../tests/fixtures/table_page_geometry/first_whole.docx"), 16, 41),
        (include_bytes!("../../../../tests/fixtures/table_page_geometry/same_rows.docx"), 26, 52),
        (include_bytes!("../../../../tests/fixtures/table_page_geometry/same_split.docx"), 26, 52),
        (include_bytes!("../../../../tests/fixtures/table_page_geometry/same_whole.docx"), 26, 51),
    ];
    for (case, (bytes, second_start, third_start)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 3, "case {case}");
        for row in 1..=60 {
            let expected_page = if row < *second_start { 1 }
                else if row < *third_start { 2 } else { 3 };
            let marker = format!("ROW {row:03}");
            let actual_pages: Vec<_> = layout.pages.iter().enumerate().filter_map(|(i, p)| {
                p.elements.iter().any(|e| matches!(&e.content, LayoutContent::Text { text, .. }
                    if text == &marker)).then_some(i + 1)
            }).collect();
            assert_eq!(actual_pages, vec![expected_page], "case {case}, {marker}");
        }
    }
}

#[test]
fn coanchored_floats_share_origin_and_move_to_next_page_together() {
    // Word-exported page memberships for two wrap modes on one paragraph.
    let cases: &[(&[u8], usize, [usize; 5])] = &[
        (include_bytes!("../../../../tests/fixtures/shared_float_anchor/y200_tb120_sq220.docx"), 1, [1, 1, 1, 1, 1]),
        (include_bytes!("../../../../tests/fixtures/shared_float_anchor/y200_tb120_sq80.docx"), 1, [1, 1, 1, 1, 1]),
        (include_bytes!("../../../../tests/fixtures/shared_float_anchor/y200_tb200_sq220.docx"), 1, [1, 1, 1, 1, 1]),
        (include_bytes!("../../../../tests/fixtures/shared_float_anchor/y200_tb200_sq80.docx"), 1, [1, 1, 1, 1, 1]),
        (include_bytes!("../../../../tests/fixtures/shared_float_anchor/y300_tb120_sq220.docx"), 1, [1, 1, 1, 1, 1]),
        (include_bytes!("../../../../tests/fixtures/shared_float_anchor/y300_tb120_sq80.docx"), 1, [1, 1, 1, 1, 1]),
        (include_bytes!("../../../../tests/fixtures/shared_float_anchor/y300_tb200_sq220.docx"), 1, [1, 1, 1, 1, 1]),
        (include_bytes!("../../../../tests/fixtures/shared_float_anchor/y300_tb200_sq80.docx"), 1, [1, 1, 1, 1, 1]),
        (include_bytes!("../../../../tests/fixtures/shared_float_anchor/y400_tb120_sq220.docx"), 2, [1, 2, 2, 2, 2]),
        (include_bytes!("../../../../tests/fixtures/shared_float_anchor/y400_tb120_sq80.docx"), 1, [1, 1, 1, 1, 1]),
        (include_bytes!("../../../../tests/fixtures/shared_float_anchor/y400_tb200_sq220.docx"), 2, [1, 2, 2, 2, 2]),
        (include_bytes!("../../../../tests/fixtures/shared_float_anchor/y400_tb200_sq80.docx"), 2, [1, 2, 2, 2, 2]),
    ];
    for (case, (bytes, page_count, expected_pages)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), *page_count, "case {case}");
        for (marker, expected) in ["FILL", "ANCHOR", "AFTER", "TOPBOX", "SIDEBOX"]
            .iter().zip(expected_pages)
        {
            let actual: Vec<_> = layout.pages.iter().enumerate().filter_map(|(i, p)| {
                p.elements.iter().any(|e| matches!(&e.content, LayoutContent::Text { text, .. }
                    if text == marker)).then_some(i + 1)
            }).collect();
            assert_eq!(actual, vec![*expected], "case {case}, {marker}");
        }
    }
}

#[test]
fn list_markers_use_text_descent_and_untyped_grid_keeps_natural_capacity() {
    let cases: &[(&[u8], usize, &[usize])] = &[
        (include_bytes!("../../../../tests/fixtures/list_pagination/Arial.docx"), 48, &[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42, 44, 46, 48]),
        (include_bytes!("../../../../tests/fixtures/list_pagination/Arial_notype.docx"), 48, &[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42, 44, 46, 48]),
        (include_bytes!("../../../../tests/fixtures/list_pagination/Symbol.docx"), 55, &[1, 2, 3, 4, 5, 6, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35, 37, 39, 41, 43, 45, 47, 49, 51, 53, 55]),
        (include_bytes!("../../../../tests/fixtures/list_pagination/Symbol_notype.docx"), 55, &[1, 2, 3, 4, 5, 6, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35, 37, 39, 41, 43, 45, 47, 49, 51, 53, 55]),
    ];
    for (case, (bytes, count, expected)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), *count, "case {case}");
        for (i, page) in expected.iter().enumerate() {
            let marker = format!("ITEM{}", 5320 + i * 2);
            let actual: Vec<_> = layout.pages.iter().enumerate().filter_map(|(j, p)| {
                p.elements.iter().any(|e| matches!(&e.content, LayoutContent::Text { text, .. }
                    if text == &marker)).then_some(j + 1)
            }).collect();
            assert_eq!(actual, vec![*page], "case {case}, {marker}");
        }
    }
}

#[test]
fn modern_justified_hanging_preserves_decimal_paren_exclusion() {
    let cases: &[(&[u8], &[usize])] = &[
        (include_bytes!("../../../../tests/fixtures/list_pagination/c14_paren.docx"), &[2]),
        (include_bytes!("../../../../tests/fixtures/list_pagination/c15_none.docx"), &[2]),
        (include_bytes!("../../../../tests/fixtures/list_pagination/c15_paren.docx"), &[3, 3, 3, 2]),
    ];
    for (case, (bytes, expected)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), expected.len(), "case {case}");
        for (page, count) in layout.pages.iter().zip(*expected) {
            let mut ys: Vec<_> = page.elements.iter().filter_map(|e| match &e.content {
                LayoutContent::Text { text, .. } if !text.trim().is_empty()
                    && !text.trim().trim_end_matches([')', '.']).chars().all(|c| c.is_ascii_digit()) => Some(e.y),
                _ => None,
            }).collect();
            ys.sort_by(f32::total_cmp);
            ys.dedup_by(|a, b| (*a - *b).abs() < 0.1);
            assert_eq!(ys.len(), *count, "case {case}");
        }
    }
}

#[test]
fn top_margins_retain_precision_and_cjk_layout_keeps_legacy_origins() {
    let cases: &[(&[u8], &[f32])] = &[
        (include_bytes!("../../../../tests/fixtures/top_margin_precision/topmargin.docx"),
            &[56.5, 56.7, 57.0, 118.5, 118.8, 119.0]),
        (include_bytes!("../../../../tests/fixtures/top_margin_precision/topmargin_cjk.docx"),
            &[56.5, 56.5, 57.0, 118.5, 119.0, 119.0]),
        (include_bytes!("../../../../tests/fixtures/top_margin_precision/topmargin_cjk_box.docx"),
            &[56.5, 56.5, 57.0, 118.5, 119.0, 119.0]),
    ];
    let declared = [56.5, 56.7, 57.0, 118.5, 118.8, 119.0];
    for (case, (bytes, expected)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        assert_eq!(doc.pages.len(), declared.len(), "case {case}");
        for (page, top) in doc.pages.iter().zip(declared) {
            assert!((page.margin.top - top).abs() < 0.001, "case {case}");
        }
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), expected.len(), "case {case}");
        for (page, top) in layout.pages.iter().zip(*expected) {
            let marker = page.elements.iter().find(|e| matches!(&e.content,
                LayoutContent::Text { text, .. } if text.starts_with("TOPMARK_"))).unwrap();
            assert!((marker.y - top).abs() < 0.001, "case {case}: {} vs {top}", marker.y);
        }
        // Rendering compatibility must not alter the source geometry.
        for (page, top) in doc.pages.iter().zip(declared) {
            assert!((page.margin.top - top).abs() < 0.001, "case {case}");
        }
    }
}

#[test]
fn widow_lookahead_uses_the_same_at_least_height_as_line_fitting() {
    let cases: &[(&[u8], usize, &[usize])] = &[
        (include_bytes!("../../../../tests/fixtures/top_margin_precision/widow3.docx"), 3, &[1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4, 5, 5, 5, 6, 6, 6, 8, 8, 8, 10, 10, 10, 12, 12, 12, 14, 14, 14, 16, 16, 16]),
        (include_bytes!("../../../../tests/fixtures/top_margin_precision/widow5.docx"), 5, &[1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 4, 4, 5, 5, 5, 5, 5, 6, 6, 6, 6, 6, 7, 7, 7, 8, 8, 9, 9, 9, 10, 10, 11, 11, 11, 12, 12, 13, 13, 13, 14, 14, 15, 15, 15, 16, 16]),
    ];
    for (case, (bytes, line_count, expected)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut rows = std::collections::HashMap::<(usize, u32), String>::new();
        for (page, p) in layout.pages.iter().enumerate() {
            for e in &p.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    rows.entry((page + 1, e.y.to_bits())).or_default()
                        .extend(text.chars().filter(|c| !c.is_whitespace()));
                }
            }
        }
        for (i, page) in expected.iter().enumerate() {
            let marker = format!("CASE{:02}LINE{}", i / line_count, i % line_count);
            let actual: Vec<_> = rows.iter().filter_map(|((p, _), text)|
                (text == &marker).then_some(*p)).collect();
            assert_eq!(actual, vec![*page], "case {case}, {marker}");
        }
    }
}


#[test]
fn cell_nbsp_uses_measured_space_advance_without_changing_break_semantics() {
    let cases: &[(&[u8], &[usize])] = &[
        (include_bytes!("../../../../tests/fixtures/cell_pagination/nbsp_plain.docx"), &[0, 0, 0, 0]),
        (include_bytes!("../../../../tests/fixtures/cell_pagination/nbsp_space.docx"), &[0, 0, 0, 0]),
        (include_bytes!("../../../../tests/fixtures/cell_pagination/nbsp_nbsp.docx"), &[1, 0, 0, 0]),
        (include_bytes!("../../../../tests/fixtures/cell_pagination/nbsp_split.docx"), &[1, 0, 0, 0]),
        (include_bytes!("../../../../tests/fixtures/cell_pagination/nbsp_nbsp2.docx"), &[1, 1, 1, 1]),
        (include_bytes!("../../../../tests/fixtures/cell_pagination/nbsp_internal.docx"), &[1, 1, 1, 1]),
        (include_bytes!("../../../../tests/fixtures/cell_pagination/nbsp_linebreak.docx"), &[2, 1, 1, 1]),
    ];
    for (case, (bytes, extra_lines)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 4, "case {case}");
        for (page, extra) in layout.pages.iter().zip(*extra_lines) {
            let after = page.elements.iter().find(|e| matches!(&e.content,
                LayoutContent::Text { text, .. } if text == "AFTER")).unwrap();
            // Word's four cell-width probes add these whole Calibri 10pt lines.
            let expected = 44.414 + *extra as f32 * 12.207;
            assert!((after.y - expected).abs() < 0.01,
                "case {case}: {} vs {expected}", after.y);
        }
    }
}

#[test]
fn cell_orphan_control_keeps_first_two_lines_after_earlier_cell_paragraphs() {
    let cases: &[(&[u8], usize, bool)] = &[
        (include_bytes!("../../../../tests/fixtures/cell_pagination/row2_0.docx"), 2, false),
        (include_bytes!("../../../../tests/fixtures/cell_pagination/row2_1.docx"), 2, true),
        (include_bytes!("../../../../tests/fixtures/cell_pagination/row5_0.docx"), 5, false),
        (include_bytes!("../../../../tests/fixtures/cell_pagination/row5_1.docx"), 5, true),
    ];
    for (case, (bytes, a_lines, widow_on)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 12, "case {case}");
        for (column, count) in [('A', *a_lines), ('B', 9)] {
            for (arm, natural_kept) in [3usize, 2, 2, 1, 1, 1].iter().enumerate() {
                let kept = if *widow_on && *natural_kept == 1 { 0 }
                    else { (*natural_kept).min(count) };
                for line in 0..count {
                    let marker = format!("{column}{arm}LINE{line}");
                    let pages: Vec<_> = layout.pages.iter().enumerate()
                        .filter_map(|(i, p)| p.elements.iter().any(|e| matches!(&e.content,
                            LayoutContent::Text { text, .. } if text == &marker)).then_some(i + 1))
                        .collect();
                    let expected = arm * 2 + if line < kept { 1 } else { 2 };
                    assert_eq!(pages, vec![expected], "case {case}, {marker}");
                }
            }
        }
    }
}


#[test]
fn cell_list_markers_follow_their_paragraph_across_an_orphan_split() {
    let mut doc = crate::parser::parse_docx(include_bytes!(
        "../../../../tests/fixtures/cell_pagination/row2_1.docx")).unwrap();
    let mut arm = 0;
    for page in &mut doc.pages {
        for block in &mut page.blocks {
            if let crate::ir::Block::Table(table) = block {
                for (col, cell) in table.rows[0].cells.iter_mut().enumerate() {
                    if let crate::ir::Block::Paragraph(para) = &mut cell.blocks[1] {
                        para.style.list_marker = Some(format!("MARK{arm}{col}"));
                        para.style.list_marker_size = Some(10.0);
                        para.style.list_indent = Some(12.0);
                    }
                }
                arm += 1;
            }
        }
    }
    assert_eq!(arm, 6);
    let layout = LayoutEngine::for_document(&doc).layout(&doc);
    for (arm, expected) in [1usize, 3, 5, 8, 10, 12].iter().enumerate() {
        for col in 0..2 {
            let marker = format!("MARK{arm}{col}");
            let found: Vec<_> = layout.pages.iter().enumerate()
                .filter_map(|(i, page)| page.elements.iter().any(|e| matches!(&e.content,
                    LayoutContent::Text { text, .. } if text == &marker)).then_some(i + 1))
                .collect();
            assert_eq!(found, vec![*expected], "{marker}");
        }
    }
}


#[test]
fn nested_cell_paragraph_indices_do_not_create_a_false_orphan() {
    let doc = crate::parser::parse_docx(include_bytes!(
        "../../../../tests/fixtures/cell_pagination/nested.docx")).unwrap();
    let layout = LayoutEngine::for_document(&doc).layout(&doc);
    assert_eq!(layout.pages.len(), 11);
    let titles = [1usize, 2, 4, 6, 8, 10];
    let nested_first = [1usize, 2, 4, 7, 9, 11];
    let nested_second = [1usize, 3, 5, 7, 9, 11];
    for arm in 0..6 {
        for (col, label) in ['A', 'B'].iter().enumerate() {
            for (marker, expected) in [
                (format!("{label}{arm}LINE0"), titles[arm]),
                (format!("N{arm}{col}LINE0"), nested_first[arm]),
                (format!("N{arm}{col}LINE1"), nested_second[arm]),
            ] {
                let found: Vec<_> = layout.pages.iter().enumerate()
                    .filter_map(|(i, page)| page.elements.iter().any(|e| matches!(&e.content,
                        LayoutContent::Text { text, .. } if text == &marker)).then_some(i + 1))
                    .collect();
                assert_eq!(found, vec![expected], "{marker}");
            }
        }
    }
}

#[test]
fn legacy_table_paragraphs_allow_a_lone_first_line() {
    let cases: &[&[u8]] = &[
        include_bytes!("../../../../tests/fixtures/cell_pagination/orphan_compat12.docx"),
        include_bytes!("../../../../tests/fixtures/cell_pagination/orphan_compat14.docx"),
        include_bytes!("../../../../tests/fixtures/cell_pagination/orphan_compat_absent.docx"),
    ];
    for (case, bytes) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 12, "case {case}");
        for (column, count) in [('A', 2), ('B', 9)] {
            for (arm, kept) in [3usize, 2, 2, 1, 1, 1].iter().enumerate() {
                for line in 0..count {
                    let marker = format!("{column}{arm}LINE{line}");
                    let pages: Vec<_> = layout.pages.iter().enumerate()
                        .filter_map(|(i, p)| p.elements.iter().any(|e| matches!(&e.content,
                            LayoutContent::Text { text, .. } if text == &marker)).then_some(i + 1))
                        .collect();
                    let expected = arm * 2 + if line < *kept { 1 } else { 2 };
                    assert_eq!(pages, vec![expected], "case {case}, {marker}");
                }
            }
        }
    }
}

#[test]
fn cell_line_spacing_preserves_defaults_until_a_style_overrides_them() {
    let cases: &[(&[u8], f32)] = &[
        (include_bytes!("../../../../tests/fixtures/cell_line_spacing/none.docx"), 15.84),
        (include_bytes!("../../../../tests/fixtures/cell_line_spacing/empty.docx"), 15.84),
        (include_bytes!("../../../../tests/fixtures/cell_line_spacing/single.docx"), 13.80),
        (include_bytes!("../../../../tests/fixtures/cell_line_spacing/multiple.docx"), 20.64),
        (include_bytes!("../../../../tests/fixtures/cell_line_spacing/direct.docx"), 27.60),
    ];
    for (case, (bytes, word_pitch)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 1, "case {case}");
        let y = |label: &str| layout.pages[0].elements.iter()
            .find(|e| matches!(&e.content, LayoutContent::Text { text, .. } if text == label))
            .unwrap().y;
        assert!((y("SECOND") - y("FIRST") - word_pitch).abs() < 0.08,
            "case {case}: Word pitch {word_pitch}");
        // The row estimate and emitted lines must reserve the same space.
        assert!((y("AFTER") - y("FIRST") - 2.0 * (y("SECOND") - y("FIRST"))).abs() < 0.01,
            "case {case}: row height disagrees with its lines");
    }
}

#[test]
fn whitespace_footer_lines_reserve_before_spacing_only_once() {
    let cases: &[&[u8]] = &[
        include_bytes!("../../../../tests/fixtures/footer_spacing/empty.docx"),
        include_bytes!("../../../../tests/fixtures/footer_spacing/tabs.docx"),
        include_bytes!("../../../../tests/fixtures/footer_spacing/spaces.docx"),
    ];
    for (case, bytes) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 11, "case {case}");
        for (arm, expected) in [1usize, 2, 3, 5, 8, 11].iter().enumerate() {
            let marker = format!("MARK{arm}");
            let pages: Vec<_> = layout.pages.iter().enumerate()
                .filter_map(|(i, page)| page.elements.iter().any(|e| matches!(&e.content,
                    LayoutContent::Text { text, .. } if text == &marker)).then_some(i + 1))
                .collect();
            assert_eq!(pages, vec![*expected], "case {case}, {marker}");
        }
    }
}

#[test]
fn empty_runs_do_not_override_the_paragraph_mark_size() {
    let cases: &[&[u8]] = &[
        include_bytes!("../../../../tests/fixtures/empty_run_mark/none.docx"),
        include_bytes!("../../../../tests/fixtures/empty_run_mark/empty11.docx"),
        include_bytes!("../../../../tests/fixtures/empty_run_mark/empty20.docx"),
        include_bytes!("../../../../tests/fixtures/empty_run_mark/space11.docx"),
        include_bytes!("../../../../tests/fixtures/empty_run_mark/space20.docx"),
    ];
    for (case, bytes) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 1, "case {case}");
        let y = |label: &str| layout.pages[0].elements.iter()
            .find(|e| matches!(&e.content, LayoutContent::Text { text, .. } if text == label))
            .unwrap().y;
        assert!((y("AFTER") - y("FIRST") - 26.4).abs() < 0.08,
            "case {case}: empty run changed the Word paragraph-mark advance");
    }
}

#[test]
fn latin_segment_punctuation_matches_word_line_breaks() {
    let cases: &[(&[u8], &[&str])] = &[
        (include_bytes!("../../../../tests/fixtures/latin_segment_hang/14_comma.docx"), &["AAA AAAAA,", "NEXT"]),
        (include_bytes!("../../../../tests/fixtures/latin_segment_hang/14_letter.docx"), &["AAA", "AAAAAx", "NEXT"]),
        (include_bytes!("../../../../tests/fixtures/latin_segment_hang/14_long_comma.docx"), &["AAAAAAAAA", "AAAAAAAAA", "AA, NEXT"]),
        (include_bytes!("../../../../tests/fixtures/latin_segment_hang/14_period.docx"), &["AAA AAAAA.", "NEXT"]),
        (include_bytes!("../../../../tests/fixtures/latin_segment_hang/14_semicolon.docx"), &["AAA", "AAAAA;", "NEXT"]),
        (include_bytes!("../../../../tests/fixtures/latin_segment_hang/15_comma.docx"), &["AAA", "AAAAA,", "NEXT"]),
        (include_bytes!("../../../../tests/fixtures/latin_segment_hang/15_letter.docx"), &["AAA", "AAAAAx", "NEXT"]),
        (include_bytes!("../../../../tests/fixtures/latin_segment_hang/15_long_comma.docx"), &["AAAAAAAAA", "AAAAAAAAA", "AA, NEXT"]),
        (include_bytes!("../../../../tests/fixtures/latin_segment_hang/15_period.docx"), &["AAA", "AAAAA.", "NEXT"]),
        (include_bytes!("../../../../tests/fixtures/latin_segment_hang/15_semicolon.docx"), &["AAA", "AAAAA;", "NEXT"]),
    ];
    for (case, (bytes, expected)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 1, "case {case}");
        let mut lines: Vec<(f32, String)> = Vec::new();
        for element in &layout.pages[0].elements {
            if let LayoutContent::Text { text, .. } = &element.content {
                if let Some((_, line)) = lines.iter_mut().find(|(y, _)| (*y - element.y).abs() < 0.01) {
                    line.push_str(text);
                } else { lines.push((element.y, text.clone())); }
            }
        }
        lines.sort_by(|a, b| a.0.partial_cmp(&b.0).unwrap());
        let actual: Vec<_> = lines.iter().map(|(_, text)| text.trim()).filter(|text| !text.is_empty()).collect();
        assert_eq!(actual.as_slice(), *expected, "case {case}");
    }
}

#[test]
fn cant_split_row_fits_the_rendered_text_height() {
    // Word: 13 Times New Roman 11pt lines need 164.436pt, not 156pt.
    // With 160pt left, cantSplit moves the entire row to the next page.
    let cases: &[(&[u8], bool, bool)] = &[
        (include_bytes!("../../../../tests/fixtures/cant_split_row_fit/lead100_cant1.docx"), true, true),
        (include_bytes!("../../../../tests/fixtures/cant_split_row_fit/lead100_cant0.docx"), true, false),
        (include_bytes!("../../../../tests/fixtures/cant_split_row_fit/lead90_cant1.docx"), false, true),
        (include_bytes!("../../../../tests/fixtures/cant_split_row_fit/lead90_cant0.docx"), false, false),
    ];
    for (case, (bytes, overflow, cant_split)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for line in 1..=13 {
            let label = format!("ROW{line:03}");
            let pages: Vec<_> = layout.pages.iter().enumerate().filter_map(|(i, p)| {
                p.elements.iter().any(|e| matches!(&e.content, LayoutContent::Text { text, .. }
                    if text == &label)).then_some(i + 1)
            }).collect();
            let expected = if *overflow && (*cant_split || line == 13) { 2 } else { 1 };
            assert_eq!(pages, vec![expected], "case {case}, {label}");
        }
    }
}

#[test]
fn structural_row_revisions_respect_each_display_mode() {
    let cases: &[(&[u8], Option<&str>, bool, bool)] = &[
        (include_bytes!("../../../../tests/fixtures/structural_revisions/row_live.docx"), None, false, false),
        (include_bytes!("../../../../tests/fixtures/structural_revisions/row_delete.docx"), Some("delete"), false, false),
        (include_bytes!("../../../../tests/fixtures/structural_revisions/row_delete_expanded.docx"), Some("delete"), false, false),
        (include_bytes!("../../../../tests/fixtures/structural_revisions/row_delete_empty_header.docx"), Some("delete"), true, false),
        (include_bytes!("../../../../tests/fixtures/structural_revisions/row_delete_all.docx"), Some("delete"), false, true),
        (include_bytes!("../../../../tests/fixtures/structural_revisions/row_insert.docx"), Some("insert"), false, false),
    ];
    for (case, (bytes, kind, empty, only_row)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let table = doc.pages[0].blocks.iter().find_map(|b| match b {
            Block::Table(t) => Some(t), _ => None,
        }).unwrap();
        let revision = table.rows[0].tracked_change.as_ref();
        assert_eq!(revision.map(|r| r.change_type.as_str()), *kind, "case {case}");
        if let Some(revision) = revision {
            assert_eq!(revision.author.as_deref(), Some("Test"));
            assert_eq!(revision.pair_id.as_deref(), Some("1"));
        }
        for view in [ShowRevisions::Final, ShowRevisions::Original, ShowRevisions::All, ShowRevisions::Simple] {
            let removed = (view == ShowRevisions::Final && *kind == Some("delete"))
                || (view == ShowRevisions::Original && *kind == Some("insert"));
            let layout = LayoutEngine::for_document(&doc).with_show_revisions(view).layout(&doc);
            assert_eq!(layout.pages.len(), 1, "case {case}, {view:?}");
            let y = |label: &str| layout.pages[0].elements.iter().find_map(|e| match &e.content {
                LayoutContent::Text { text, .. } if text == label => Some(e.y), _ => None,
            });
            assert_eq!(y("REVISED").is_some(), !removed && !empty, "case {case}, {view:?}");
            if !only_row {
                assert!((y("KEEP").unwrap() - if removed { 72.0 } else { 112.0 }).abs() < 0.1,
                    "case {case}, {view:?}");
            }
            let after = 72.0 + if removed { 0.0 } else { 40.0 }
                + if *only_row { 0.0 } else { 11.5 };
            assert!((y("AFTER").unwrap() - after).abs() < 0.1, "case {case}, {view:?}");
        }
    }
}

#[test]
fn deleted_empty_break_before_table_keeps_its_bookmark() {
    let cases: &[(&[u8], f32)] = &[
        (include_bytes!("../../../../tests/fixtures/structural_revisions/gap_live.docx"), 56.4),
        (include_bytes!("../../../../tests/fixtures/structural_revisions/gap_deleted.docx"), 13.2),
        (include_bytes!("../../../../tests/fixtures/structural_revisions/gap_chain.docx"), 13.2),
    ];
    for (case, (bytes, advance)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).with_show_revisions(ShowRevisions::Final).layout(&doc);
        let y = |label: &str| layout.pages[0].elements.iter().find_map(|e| match &e.content {
            LayoutContent::Text { text, .. } if text == label => Some(e.y), _ => None,
        }).unwrap();
        assert!((y("SECOND") - y("FIRST") - advance).abs() < 0.1, "case {case}");
    }
    let mut doc = crate::parser::parse_docx(cases[1].0).unwrap();
    filter_runs_for_show_revisions(&mut doc, true);
    let table = doc.pages[0].blocks.iter().filter_map(|b| match b {
        Block::Table(t) => Some(t), _ => None,
    }).nth(1).unwrap();
    let Block::Paragraph(p) = &table.rows[0].cells[0].blocks[0] else { panic!("cell paragraph"); };
    assert!(p.runs.iter().any(|r| r.bookmark_name.as_deref() == Some("GapAnchor")));
    assert!(p.runs.iter().any(|r| r.text == "SECOND"));
}

#[test]
fn minimum_row_height_controls_first_and_continuation_fragments() {
    let cases: &[(&[u8], usize, &[usize])] = &[
        (include_bytes!("../../../../tests/fixtures/minimum_row_height/long_min180_fresh0.docx"), 3, &[2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3]),
        (include_bytes!("../../../../tests/fixtures/minimum_row_height/long_min60_fresh0.docx"), 3, &[1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3]),
        (include_bytes!("../../../../tests/fixtures/minimum_row_height/long_min800_fresh0.docx"), 4, &[2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 3, 4]),
        (include_bytes!("../../../../tests/fixtures/minimum_row_height/long_min800_fresh1.docx"), 3, &[1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3]),
        (include_bytes!("../../../../tests/fixtures/minimum_row_height/min0.docx"), 2, &[1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2]),
        (include_bytes!("../../../../tests/fixtures/minimum_row_height/min180.docx"), 2, &[2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2]),
        (include_bytes!("../../../../tests/fixtures/minimum_row_height/min400.docx"), 3, &[2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3]),
        (include_bytes!("../../../../tests/fixtures/minimum_row_height/min401.docx"), 3, &[2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3]),
        (include_bytes!("../../../../tests/fixtures/minimum_row_height/min60.docx"), 2, &[1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2]),
        (include_bytes!("../../../../tests/fixtures/minimum_row_height/min800.docx"), 3, &[2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3]),
        (include_bytes!("../../../../tests/fixtures/minimum_row_height/short_min180.docx"), 2, &[2, 2, 2, 2, 2, 2]),
        (include_bytes!("../../../../tests/fixtures/minimum_row_height/short_min800.docx"), 3, &[2, 2, 2, 2, 2, 3]),
    ];
    for (case, (bytes, page_count, expected)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), *page_count, "case {case}");
        for (index, expected_page) in expected.iter().enumerate() {
            let label = if index + 1 == expected.len() { "AFTER".to_string() } else { format!("ROW{:03}", index + 1) };
            let actual: Vec<_> = layout.pages.iter().enumerate().filter_map(|(i, page)| {
                page.elements.iter().any(|e| matches!(&e.content, LayoutContent::Text { text, .. }
                    if text == &label)).then_some(i + 1)
            }).collect();
            assert_eq!(actual, vec![*expected_page], "case {case}, {label}");
        }
    }
}
