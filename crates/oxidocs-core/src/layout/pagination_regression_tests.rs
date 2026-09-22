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
fn document_break_count_distrust_preserves_measured_table_pagination() {
    let bytes = include_bytes!(concat!(env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/table_row_cached_break_explicit_tail.docx"));
    let doc = crate::parser::parse_docx(bytes).unwrap();
    let engine = LayoutEngine::for_document(&doc);
    let trusted = engine.layout_pass(&doc);
    assert_eq!(trusted.pages.len(), 2);
    engine.lrpb_count_distrust.set(true);
    let distrusted = engine.layout_pass(&doc);
    assert_eq!(distrusted.pages.len(), trusted.pages.len());
    // Word keeps both rows on page one; only the trailing paragraph overflows.
    assert!(distrusted.pages[0].elements.iter().any(|e| {
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
        LayoutEngine::rebalance_text_columns(&mut elements, 0.0, &[0.0, 100.0], &[], 0.0, 14, None),
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
        LayoutEngine::rebalance_text_columns(&mut elements, 0.0, &[0.0, 100.0], &[], 0.0, 14, None),
        Some(30.0)
    );
    assert_eq!((elements[1].x, elements[1].y), (100.0, 0.0));
    assert_eq!((elements[4].x, elements[4].y), (102.0, -6.0));
}

#[test]
fn text_balance_does_not_split_a_hanging_line_at_the_column_boundary() {
    let mut elements = vec![text_row(90.0, 0.0, 10.0, 0), text_row(110.0, 0.0, 10.0, 0)];
    assert!(LayoutEngine::rebalance_text_columns(&mut elements, 0.0, &[0.0, 100.0], &[], 0.0, 14, None).is_none());
    assert_eq!((elements[0].y, elements[1].y), (0.0, 0.0));
}

#[test]
fn text_balance_leaves_table_content_for_table_path() {
    let mut elements = vec![text_row(0.0, 0.0, 10.0, 0), text_row(0.0, 10.0, 10.0, 1)];
    elements[1].cell_paragraph_index = Some(0);
    assert!(LayoutEngine::rebalance_text_columns(&mut elements, 0.0, &[0.0, 100.0], &[], 0.0, 14, None).is_none());
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
    assert!(LayoutEngine::rebalance_text_columns(&mut elements, 0.0, &[0.0, 100.0], &[], 0.0, 14, None).is_none());
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
fn top_margins_retain_precision_in_latin_and_cjk_layout() {
    let cases: &[(&[u8], &[f32])] = &[
        (include_bytes!("../../../../tests/fixtures/top_margin_precision/topmargin.docx"),
            &[56.5, 56.7, 57.0, 118.5, 118.8, 119.0]),
        (include_bytes!("../../../../tests/fixtures/top_margin_precision/topmargin_cjk.docx"),
            &[56.5, 56.7, 57.0, 118.5, 118.8, 119.0]),
        (include_bytes!("../../../../tests/fixtures/top_margin_precision/topmargin_cjk_box.docx"),
            &[56.5, 56.7, 57.0, 118.5, 118.8, 119.0]),
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
        // Layout must not alter the source geometry.
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

#[test]
fn heading_and_fitting_table_keep_their_word_page_boundaries() {
    let cases: &[(&[u8], usize, [usize; 3])] = &[
        (include_bytes!("../../../../tests/fixtures/keep_next_table/rem24_keep0.docx"), 2, [1, 2, 2]),
        (include_bytes!("../../../../tests/fixtures/keep_next_table/rem24_keep1.docx"), 2, [2, 2, 2]),
        (include_bytes!("../../../../tests/fixtures/keep_next_table/rem28.8_keep0.docx"), 2, [1, 2, 2]),
        (include_bytes!("../../../../tests/fixtures/keep_next_table/rem28.8_keep1.docx"), 2, [2, 2, 2]),
        (include_bytes!("../../../../tests/fixtures/keep_next_table/rem29_keep0.docx"), 2, [1, 2, 2]),
        (include_bytes!("../../../../tests/fixtures/keep_next_table/rem29_keep1.docx"), 2, [2, 2, 2]),
        (include_bytes!("../../../../tests/fixtures/keep_next_table/rem35_keep0.docx"), 2, [1, 2, 2]),
        (include_bytes!("../../../../tests/fixtures/keep_next_table/rem35_keep1.docx"), 2, [2, 2, 2]),
        (include_bytes!("../../../../tests/fixtures/keep_next_table/rem60_keep0.docx"), 2, [1, 1, 2]),
        (include_bytes!("../../../../tests/fixtures/keep_next_table/rem60_keep1.docx"), 2, [1, 1, 2]),
        (include_bytes!("../../../../tests/fixtures/keep_next_table/rem80_keep0.docx"), 1, [1, 1, 1]),
        (include_bytes!("../../../../tests/fixtures/keep_next_table/rem80_keep1.docx"), 1, [1, 1, 1]),
    ];
    for (case, (bytes, page_count, expected)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), *page_count, "case {case}");
        for (label, expected_page) in ["HEADING", "ROW001", "AFTER"].iter().zip(expected) {
            let actual: Vec<_> = layout.pages.iter().enumerate().filter_map(|(i, page)| {
                page.elements.iter().any(|e| matches!(&e.content, LayoutContent::Text { text, .. }
                    if text == label)).then_some(i + 1)
            }).collect();
            assert_eq!(actual, vec![*expected_page], "case {case}, {label}");
        }
    }
}

#[test]
fn inline_symbols_preserve_word_line_advances() {
    let cases: &[(&[u8], &[f32], Option<char>)] = &[
        (include_bytes!("../../../../tests/fixtures/run_symbols/plain_mixed0.docx"), &[0.0, 11.546, 23.066, 34.586, 46.106, 57.626, 69.026, 80.546], None),
        (include_bytes!("../../../../tests/fixtures/run_symbols/plain_mixed1.docx"), &[0.0, 11.546, 23.066, 34.586, 46.106, 57.626, 69.026, 80.546], None),
        (include_bytes!("../../../../tests/fixtures/run_symbols/sorts_mixed0.docx"), &[0.0, 11.786, 23.546, 35.186, 46.946, 58.706, 70.466, 82.106], Some('\u{F070}')),
        (include_bytes!("../../../../tests/fixtures/run_symbols/sorts_mixed1.docx"), &[0.0, 11.786, 23.546, 35.186, 46.946, 58.706, 70.466, 82.106], Some('\u{F070}')),
        (include_bytes!("../../../../tests/fixtures/run_symbols/symbol_mixed0.docx"), &[0.0, 12.266, 24.506, 36.746, 48.986, 61.226, 73.466, 85.706], Some('\u{F0B7}')),
        (include_bytes!("../../../../tests/fixtures/run_symbols/symbol_mixed1.docx"), &[0.0, 12.266, 24.506, 36.746, 48.986, 61.226, 73.466, 85.706], Some('\u{F0B7}')),
        (include_bytes!("../../../../tests/fixtures/run_symbols/wing_mixed0.docx"), &[0.0, 11.546, 23.066, 34.586, 46.106, 57.626, 69.026, 80.546], Some('\u{F0FC}')),
        (include_bytes!("../../../../tests/fixtures/run_symbols/wing_mixed1.docx"), &[0.0, 11.546, 23.066, 34.586, 46.106, 57.626, 69.026, 80.546], Some('\u{F0FC}')),
    ];
    for (case, (bytes, expected, symbol)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 2, "case {case}");
        let mut positions = Vec::new();
        let mut symbols = 0;
        for (page, p) in layout.pages.iter().enumerate() {
            for e in &p.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    if text.starts_with("ROW") {
                        assert_eq!(page, 0, "case {case}");
                        positions.push(e.y);
                    }
                    if text == "AFTER" { assert_eq!(page, 1, "case {case}"); }
                    if let Some(symbol) = symbol { symbols += text.chars().filter(|c| c == symbol).count(); }
                }
            }
        }
        assert_eq!(positions.len(), expected.len(), "case {case}");
        for (y, expected) in positions.iter().zip(*expected) {
            assert!((y - positions[0] - expected).abs() < 0.15, "case {case}: {} vs {expected}", y - positions[0]);
        }
        assert_eq!(symbols, if symbol.is_some() { 8 } else { 0 }, "case {case}");
    }
}

#[test]
fn declared_legacy_symbols_use_the_word_substitute() {
    let cases: &[&[u8]] = &[
        include_bytes!("../../../../tests/fixtures/run_symbols/declared_ansi.docx"),
        include_bytes!("../../../../tests/fixtures/run_symbols/declared_symbol.docx"),
        include_bytes!("../../../../tests/fixtures/run_symbols/declared_zero.docx"),
        include_bytes!("../../../../tests/fixtures/run_symbols/declared_segoe.docx"),
    ];
    for (case, bytes) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 2, "case {case}");
        for line in 1..=8 {
            let label = format!("ROW{line:03}");
            let pages: Vec<_> = layout.pages.iter().enumerate().filter_map(|(i, page)| {
                page.elements.iter().any(|e| matches!(&e.content, LayoutContent::Text { text, .. }
                    if text == &label)).then_some(i + 1)
            }).collect();
            assert_eq!(pages, vec![if line == 8 { 2 } else { 1 }], "case {case}, {label}");
        }
        let after: Vec<_> = layout.pages.iter().enumerate().filter_map(|(i, page)| {
            page.elements.iter().any(|e| matches!(&e.content, LayoutContent::Text { text, .. }
                if text == "AFTER")).then_some(i + 1)
        }).collect();
        assert_eq!(after, vec![2], "case {case}");
    }
}

#[test]
fn glued_symbols_keep_their_font_without_adding_word_breaks() {
    let cases: &[(&[u8], char, &str)] = &[
        (include_bytes!("../../../../tests/fixtures/run_symbols/glued_symbol.docx"), '\u{F0AF}', "Symbol"),
        (include_bytes!("../../../../tests/fixtures/run_symbols/glued_wingdings.docx"), '\u{F0E0}', "Wingdings"),
    ];
    for &(bytes, symbol, family) in cases {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut count = 0;
        for page in &layout.pages {
            for e in &page.elements {
                if let LayoutContent::Text { text, font_family, .. } = &e.content {
                    if text.contains(symbol) {
                        assert_eq!(font_family.as_deref(), Some(family), "{family}: {text:?}");
                        assert!(!text.chars().any(|c| c.is_ascii_alphabetic()), "symbol font leaked into Latin text: {text:?}");
                        count += text.chars().filter(|&c| c == symbol).count();
                    }
                }
            }
        }
        assert_eq!(count, 13, "{family}");
        // Word keeps the complete BBBB + symbol + CCCC token together at 90pt.
        // Flushing on each font boundary would incorrectly leave BBBB on line 1.
        let p = 9;
        let elements: Vec<_> = layout.pages.iter().flat_map(|p| &p.elements)
            .filter(|e| e.paragraph_index == Some(p)).collect();
        let a = elements.iter().find(|e| matches!(&e.content, LayoutContent::Text { text, .. } if text.contains("AAAA"))).unwrap();
        let b = elements.iter().find(|e| matches!(&e.content, LayoutContent::Text { text, .. } if text.contains("BBBB"))).unwrap();
        assert!(b.y > a.y + 5.0, "{family}: font boundary changed word wrapping");
    }
}

#[test]
fn inline_symbol_row_combines_ascent_and_descent() {
    let cases: &[&[u8]] = &[
        include_bytes!("../../../../tests/fixtures/run_symbols/calibri_row_plain.docx"),
        include_bytes!("../../../../tests/fixtures/run_symbols/calibri_row_symbol.docx"),
    ];
    let ys: Vec<_> = cases.iter().map(|bytes| {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        layout.pages.iter().flat_map(|p| &p.elements).find_map(|e| {
            matches!(&e.content, LayoutContent::Text { text, .. } if text == "AFTER").then_some(e.y)
        }).unwrap()
    }).collect();
    // Fresh Word exports: 139.00998 - 138.529999 = 0.47998pt.
    assert!((ys[1] - ys[0] - 0.48).abs() < 0.03, "{ys:?}");
}

#[test]
fn inline_symbol_height_respects_spacing_rules() {
    let cases: &[(&[u8], &[u8])] = &[
        (include_bytes!("../../../../tests/fixtures/run_symbols/spacing_body_plain.docx"), include_bytes!("../../../../tests/fixtures/run_symbols/spacing_body_symbol.docx")),
        (include_bytes!("../../../../tests/fixtures/run_symbols/spacing_cell_plain.docx"), include_bytes!("../../../../tests/fixtures/run_symbols/spacing_cell_symbol.docx")),
    ];
    let expected = [0.0, 0.0, 0.0, 0.0, 0.0, 0.48, 0.48, 0.0, 0.48, 0.0, 6.12, 9.24, 0.0, 6.12, 5.16];
    for (context, &(plain, symbol)) in cases.iter().enumerate() {
        let positions: Vec<_> = [plain, symbol].iter().map(|bytes| {
            let doc = crate::parser::parse_docx(bytes).unwrap();
            let layout = LayoutEngine::for_document(&doc).layout(&doc);
            layout.pages.iter().flat_map(|p| &p.elements).filter_map(|e| {
                if let LayoutContent::Text { text, .. } = &e.content {
                    text.strip_prefix("AFTER").and_then(|s| s.parse::<usize>().ok()).map(|i| (i, e.y))
                } else { None }
            }).collect::<std::collections::HashMap<_, _>>()
        }).collect();
        for (i, expected) in expected.iter().enumerate() {
            let delta = positions[1][&i] - positions[0][&i];
            assert!((delta - expected).abs() < 0.06, "context {context}, case {i}: {delta} vs {expected}");
        }
    }
}

#[test]
fn inline_symbol_cell_alignment_is_independent_of_column_order() {
    let cases: &[(&[u8], &[u8], &str, f32)] = &[
        (include_bytes!("../../../../tests/fixtures/run_symbols/center_first_symbol0.docx"), include_bytes!("../../../../tests/fixtures/run_symbols/center_first_symbol1.docx"), "CELL00", 0.24),
        (include_bytes!("../../../../tests/fixtures/run_symbols/center_last_symbol0.docx"), include_bytes!("../../../../tests/fixtures/run_symbols/center_last_symbol1.docx"), "CELL01", 0.24),
        (include_bytes!("../../../../tests/fixtures/run_symbols/bottom_first_symbol0.docx"), include_bytes!("../../../../tests/fixtures/run_symbols/bottom_first_symbol1.docx"), "CELL00", 0.48),
    ];
    for &(plain, symbol, label, expected) in cases {
        let positions: Vec<_> = [plain, symbol].iter().map(|bytes| {
            let doc = crate::parser::parse_docx(bytes).unwrap();
            let layout = LayoutEngine::for_document(&doc).layout(&doc);
            layout.pages.iter().flat_map(|p| &p.elements).filter_map(|e| {
                if let LayoutContent::Text { text, .. } = &e.content {
                    Some((text.clone(), e.y))
                } else { None }
            }).collect::<std::collections::HashMap<_, _>>()
        }).collect();
        // Fresh Word exports: center +0.24pt in either column; bottom +0.48pt.
        let delta = positions[1][label] - positions[0][label];
        assert!((delta - expected).abs() < 0.03, "{label}: {delta} vs {expected}");
        let after = positions[1]["AFTER"] - positions[0]["AFTER"];
        assert!((after - 0.48).abs() < 0.03, "row advance: {after}");
    }
}

#[test]
fn cell_page_break_before_preserves_row_and_header_boundaries() {
    let expected: serde_json::Value = serde_json::from_str(include_str!("../../../../tests/fixtures/cell_page_break_before/word.json")).unwrap();
    let cases: &[(&str, &[u8])] = &[
        ("control", include_bytes!("../../../../tests/fixtures/cell_page_break_before/control.docx")),
        ("first_fresh", include_bytes!("../../../../tests/fixtures/cell_page_break_before/first_fresh.docx")),
        ("first_lead", include_bytes!("../../../../tests/fixtures/cell_page_break_before/first_lead.docx")),
        ("left", include_bytes!("../../../../tests/fixtures/cell_page_break_before/left.docx")),
        ("left_cant", include_bytes!("../../../../tests/fixtures/cell_page_break_before/left_cant.docx")),
        ("mid_left", include_bytes!("../../../../tests/fixtures/cell_page_break_before/mid_left.docx")),
        ("mid_right", include_bytes!("../../../../tests/fixtures/cell_page_break_before/mid_right.docx")),
        ("right", include_bytes!("../../../../tests/fixtures/cell_page_break_before/right.docx")),
        ("left_keep", include_bytes!("../../../../tests/fixtures/cell_page_break_before/left_keep.docx")),
        ("left_header", include_bytes!("../../../../tests/fixtures/cell_page_break_before/left_header.docx")),
        ("left_header_keep", include_bytes!("../../../../tests/fixtures/cell_page_break_before/left_header_keep.docx")),
        ("tall_0", include_bytes!("../../../../tests/fixtures/cell_page_break_before/tall_0.docx")),
        ("tall_1", include_bytes!("../../../../tests/fixtures/cell_page_break_before/tall_1.docx")),
        ("inherited", include_bytes!("../../../../tests/fixtures/cell_page_break_before/inherited.docx")),
        ("direct_off", include_bytes!("../../../../tests/fixtures/cell_page_break_before/direct_off.docx")),
        ("derived_off", include_bytes!("../../../../tests/fixtures/cell_page_break_before/derived_off.docx")),
        ("manual_leading", include_bytes!("../../../../tests/fixtures/cell_page_break_before/manual_leading.docx")),
        ("manual_same_run", include_bytes!("../../../../tests/fixtures/cell_page_break_before/manual_same_run.docx")),
    ];
    for &(name, bytes) in cases {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let word = &expected[name];
        assert_eq!(layout.pages.len(), word["pages"].as_u64().unwrap() as usize, "{name}");
        let labels = word["labels"].as_object().unwrap();
        let mut actual = std::collections::HashMap::<String, Vec<usize>>::new();
        for (page_idx, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    for label in text.split_whitespace().filter(|t| labels.contains_key(*t)) {
                        actual.entry(label.to_owned()).or_default().push(page_idx + 1);
                    }
                }
            }
        }
        for (label, pages) in labels {
            let want: Vec<usize> = pages.as_array().unwrap().iter().map(|p| p.as_u64().unwrap() as usize).collect();
            assert_eq!(actual.get(label), Some(&want), "{name}: {label}");
        }
    }
}

#[test]
fn split_header_extent_is_independent_of_saved_page_breaks() {
    let cases: &[([&[u8]; 4], usize, f32)] = &[
        ([
            include_bytes!("../../../../tests/fixtures/split_header_extent/short_header0_marker0.docx"),
            include_bytes!("../../../../tests/fixtures/split_header_extent/short_header0_marker1.docx"),
            include_bytes!("../../../../tests/fixtures/split_header_extent/short_header1_marker0.docx"),
            include_bytes!("../../../../tests/fixtures/split_header_extent/short_header1_marker1.docx"),
        ], 2, 24.48),
        ([
            include_bytes!("../../../../tests/fixtures/split_header_extent/long_header0_marker0.docx"),
            include_bytes!("../../../../tests/fixtures/split_header_extent/long_header0_marker1.docx"),
            include_bytes!("../../../../tests/fixtures/split_header_extent/long_header1_marker0.docx"),
            include_bytes!("../../../../tests/fixtures/split_header_extent/long_header1_marker1.docx"),
        ], 4, 72.51),
    ];
    for (cases, expected_pages, header_delta) in cases {
        let mut positions = Vec::new();
        for bytes in cases {
            let doc = crate::parser::parse_docx(bytes).unwrap();
            let layout = LayoutEngine::for_document(&doc).layout(&doc);
            assert_eq!(layout.pages.len(), *expected_pages);
            let y = layout.pages.last().unwrap().elements.iter().find_map(|e| {
                matches!(&e.content, LayoutContent::Text { text, .. } if text == "NEXT0").then_some(e.y)
            }).unwrap();
            positions.push(y);
        }
        // Repeated headers must retain their measured extent with or without
        // saved hints, across one and three continuation pages. The unhinted
        // no-header control isolates the advance measured in fresh Word exports.
        // Hinted splits without repeated headers retain their separate cursor
        // path; their line reanchoring is outside this header regression.
        assert!((positions[2] - positions[3]).abs() < 0.03, "repeated header: {positions:?}");
        for repeated in &positions[2..] {
            assert!((repeated - positions[0] - header_delta).abs() < 0.3, "header advance: {positions:?}");
        }
    }
}

#[test]
fn keep_lines_fits_the_last_line_before_trailing_spacing() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/keep_lines_page_end");
    let expected: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in expected.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), case["pages"].as_u64().unwrap() as usize, "{name}");
        let labels = case["labels"].as_object().unwrap();
        let mut actual = std::collections::HashMap::new();
        for (page_idx, page) in layout.pages.iter().enumerate() {
            for element in &page.elements {
                if let LayoutContent::Text { text, .. } = &element.content {
                    if labels.contains_key(text.trim()) {
                        actual.insert(text.trim().to_owned(), page_idx + 1);
                    }
                }
            }
        }
        for (label, position) in labels {
            assert_eq!(actual.get(label).copied(), position["page"].as_u64().map(|n| n as usize), "{name}: {label}");
        }
    }
}

#[test]
fn cell_image_paragraph_spacing_contributes_to_footer_extent() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cell_image_spacing");
    for mode in ["direct", "style"] {
        for font in [9, 11] {
            let mut positions = Vec::new();
            for gap in [0, 2] {
                let name = format!("{mode}_font{font}_gap{gap}");
                let doc = crate::parser::parse_docx(
                    &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
                ).unwrap();
                let layout = LayoutEngine::for_document(&doc).layout(&doc);
                assert_eq!(layout.pages.len(), 1, "{name}");
                let text_y = |label: &str| layout.pages[0].elements.iter().find_map(|e| {
                    matches!(&e.content, LayoutContent::Text { text, .. } if text == label).then_some(e.y)
                }).unwrap();
                let image_y = layout.pages[0].elements.iter().find_map(|e| {
                    matches!(&e.content, LayoutContent::Image { .. }).then_some(e.y)
                }).unwrap();
                positions.push((text_y("BEFORE"), text_y("AFTER"), image_y));
            }
            // Fresh Word exports: the two 2pt margins grow the footer stack
            // by 4.104pt; the anchored footer end stays fixed and the picture
            // moves upward by 2pt, including when spacing comes from a style.
            assert!((positions[0].0 - positions[1].0 - 4.104).abs() < 0.2, "{mode}/{font}: {positions:?}");
            assert!((positions[0].1 - positions[1].1).abs() < 0.03, "{mode}/{font}: {positions:?}");
            assert!((positions[0].2 - positions[1].2 - 2.0).abs() < 0.2, "{mode}/{font}: {positions:?}");
        }
    }
}

#[test]
fn minimum_row_height_allows_fitting_continuation_fragments() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/row_split_minimum");
    let expected: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in expected.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        // This matrix checks row-fragment placement. Word also emits a trailing
        // empty page in three short-row cases; that separate end-paragraph
        // pagination difference is not covered by this regression test.
        let labels = case["labels"].as_object().unwrap();
        let mut actual = std::collections::HashMap::new();
        for (page_idx, page) in layout.pages.iter().enumerate() {
            for element in &page.elements {
                if let LayoutContent::Text { text, .. } = &element.content {
                    if labels.contains_key(text.trim()) {
                        actual.insert(text.trim().to_owned(), page_idx + 1);
                    }
                }
            }
        }
        for (label, position) in labels {
            assert_eq!(actual.get(label).copied(), position["page"].as_u64().map(|n| n as usize), "{name}: {label}");
        }
    }
}

#[test]
fn footer_cell_top_border_is_counted_once() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cell_image_spacing");
    for gap in [0, 2] {
        let mut positions = Vec::new();
        for mode in ["direct", "border"] {
            let doc = crate::parser::parse_docx(&std::fs::read(
                fixtures.join(format!("{mode}_font9_gap{gap}.docx")),
            ).unwrap()).unwrap();
            let layout = LayoutEngine::for_document(&doc).layout(&doc);
            let text_y = |label: &str| layout.pages[0].elements.iter().find_map(|e| {
                matches!(&e.content, LayoutContent::Text { text, .. } if text == label).then_some(e.y)
            }).unwrap();
            positions.push((text_y("BEFORE"), text_y("AFTER")));
        }
        assert!((positions[0].0 - positions[1].0 - 1.0).abs() < 0.2, "gap{gap}: {positions:?}");
        assert!((positions[0].1 - positions[1].1).abs() < 0.03, "gap{gap}: {positions:?}");
    }
}

#[test]
fn cell_image_spacing_is_shared_within_a_paragraph() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cell_image_spacing");
    for (mode, growth) in [("same", 4.104_f32), ("separate", 6.0_f32)] {
        let mut positions = Vec::new();
        for gap in [0, 2] {
            let doc = crate::parser::parse_docx(&std::fs::read(
                fixtures.join(format!("multi_{mode}_gap{gap}.docx")),
            ).unwrap()).unwrap();
            let layout = LayoutEngine::for_document(&doc).layout(&doc);
            assert_eq!(layout.pages.len(), 1);
            let y = |label: &str| layout.pages[0].elements.iter().find_map(|e| {
                matches!(&e.content, LayoutContent::Text { text, .. } if text == label).then_some(e.y)
            }).unwrap();
            positions.push((y("BEFORE"), y("AFTER")));
        }
        assert!((positions[0].0 - positions[1].0 - growth).abs() < 0.2, "{mode}: {positions:?}");
        assert!((positions[0].1 - positions[1].1).abs() < 0.03, "{mode}: {positions:?}");
    }
}

#[test]
fn empty_header_mark_uses_its_inherited_paragraph_font() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/header_inherited_mark");
    let mut body_y = Vec::new();
    for font in ["TimesNewRoman", "Calibri"] {
        let mut ys = Vec::new();
        for explicit in [0, 1] {
            let doc = crate::parser::parse_docx(&std::fs::read(
                fixtures.join(format!("{font}_explicit{explicit}.docx")),
            ).unwrap()).unwrap();
            let layout = LayoutEngine::for_document(&doc).layout(&doc);
            assert_eq!(layout.pages.len(), 1);
            ys.push(layout.pages[0].elements.iter().find_map(|e| {
                matches!(&e.content, LayoutContent::Text { text, .. } if text == "BODY").then_some(e.y)
            }).unwrap());
        }
        // Word's BODY baseline is identical for inherited and direct marks.
        assert!((ys[0] - ys[1]).abs() < 0.03, "{font}: {ys:?}");
        body_y.push(ys[0]);
    }
    // Word: Calibri adds 0.84pt over the Times New Roman header mark.
    assert!((body_y[1] - body_y[0] - 0.84).abs() < 0.15, "{body_y:?}");
}

#[test]
fn modern_short_cell_widow_moves_the_whole_row() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/short_cell_widow");
    let expected: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in expected.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(&std::fs::read(
            fixtures.join(format!("{name}.docx")),
        ).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), case["pages"].as_u64().unwrap() as usize, "{name}");
        let labels = case["labels"].as_object().unwrap();
        let mut actual = std::collections::HashMap::new();
        for (i, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    for token in text.split_whitespace() {
                        if labels.contains_key(token) { actual.insert(token.to_owned(), i + 1); }
                    }
                }
            }
        }
        for (label, page) in labels {
            assert_eq!(actual.get(label).copied(), page.as_u64().map(|p| p as usize), "{name}: {label}");
        }
    }
}

#[test]
fn legacy_cjk_cell_lines_do_not_gain_modern_widow_control() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cjk_cell_leading");
    let expected: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in expected.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(&std::fs::read(
            fixtures.join(format!("{name}.docx")),
        ).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), case["pages"].as_u64().unwrap() as usize, "{name}");
        let labels = case["labels"].as_object().unwrap();
        let mut actual = std::collections::HashMap::new();
        for (i, page) in layout.pages.iter().enumerate() {
            let mut lines = std::collections::BTreeMap::<i32, String>::new();
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    lines.entry((e.y * 100.0).round() as i32).or_default().push_str(text);
                }
            }
            for line in lines.values() {
                for label in labels.keys() {
                    if line.contains(label.as_str()) { actual.insert(label.clone(), i + 1); }
                }
            }
        }
        for (label, position) in labels {
            assert_eq!(actual.get(label).copied(), position["page"].as_u64().map(|p| p as usize), "{name}: {label}");
        }
    }
}

#[test]
fn cjk_exact_lines_preserve_fractional_top_origins_with_each_grid_mode() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cjk_top_origin");
    let expected: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in expected.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(&std::fs::read(
            fixtures.join(format!("{name}.docx")),
        ).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 1, "{name}");
        let mut ys: Vec<f32> = layout.pages[0].elements.iter().filter_map(|e|
            matches!(&e.content, LayoutContent::Text { text, .. } if !text.is_empty())
                .then_some(e.y)
        ).collect();
        ys.sort_by(f32::total_cmp);
        ys.dedup_by(|a, b| (*a - *b).abs() < 0.001);
        let top = case["top"].as_f64().unwrap() as f32;
        assert_eq!(ys.len(), 2, "{name}");
        assert!((ys[0] - top).abs() < 0.001, "{name}: {} vs {top}", ys[0]);
        assert!((ys[1] - ys[0] - 18.0).abs() < 0.001, "{name}");
    }
}

#[test]
fn fixed_image_top_bottom_bands_displace_intersecting_paragraphs() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/fixed_image_band");
    let expected: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    let cases = expected.as_array().unwrap();
    let reference = cases.iter().find(|c| c["name"] == "page_y2_fill0").unwrap();
    let baseline_offset = reference["labels"]["MARK"]["baseline"].as_f64().unwrap() as f32 - 72.0;
    for case in cases {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(&std::fs::read(
            fixtures.join(format!("{name}.docx")),
        ).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 1, "{name}");
        for (label, position) in case["labels"].as_object().unwrap() {
            let element = layout.pages[0].elements.iter().find(|e|
                matches!(&e.content, LayoutContent::Text { text, .. } if text == label)
            ).unwrap();
            let expected_y = position["baseline"].as_f64().unwrap() as f32 - baseline_offset;
            assert!((element.y - expected_y).abs() < 0.15,
                "{name}: {label} {} vs {expected_y}", element.y);
        }
    }
}

#[test]
fn contextual_cell_tail_suppresses_after_spacing_for_text_and_images() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cell_contextual_tail");
    let expected: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    let mut measured = std::collections::HashMap::new();
    for case in expected.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(&std::fs::read(
            fixtures.join(format!("{name}.docx")),
        ).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 1, "{name}");
        let y = layout.pages[0].elements.iter().find(|e|
            matches!(&e.content, LayoutContent::Text { text, .. } if text == "AFTER")
        ).unwrap().y;
        measured.insert(name.to_owned(), (y, case["after_baseline"].as_f64().unwrap() as f32));
    }
    for (name, (actual, word)) in &measured {
        let kind = name.split('_').next().unwrap();
        let reference = &measured[&format!("{kind}_ctx0_after0_other0")];
        assert!(((actual - reference.0) - (word - reference.1)).abs() < 0.15,
            "{name}: Oxi delta {} vs Word {}", actual - reference.0, word - reference.1);
    }
}

#[test]
fn split_row_borders_close_both_cells_near_the_page_bottom() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/row_split_minimum");
    // Centers of the horizontal rules in fresh Word PDF exports. The
    // one-line fragment also includes the completed LABEL cell's 2pt after.
    for (name, bottom) in [
        ("lines6_height0_free22", 214.32_f32),
        ("lines6_height0_free40", 217.32_f32),
        ("lines6_height30_free40", 217.32_f32),
    ] {
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 2, "{name}");
        for (page, expected_y) in [(0, bottom), (1, 20.28)] {
            for x in [100.0_f32, 260.0] {
                assert!(layout.pages[page].elements.iter().any(|e| matches!(&e.content,
                    LayoutContent::TableBorder { x1, x2, y1, y2, .. }
                    if (*y1 - *y2).abs() < 0.1 && (*y1 - expected_y).abs() < 0.8
                        && *x1 <= x && *x2 >= x)),
                    "{name}: page {} missing rule at x={x}, y={expected_y}", page + 1);
            }
        }
    }
}

#[test]
fn cjk_split_row_border_encloses_the_full_line_box() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cjk_split_border");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 2, "{name}");
        let edges = case["pages"][0]["horizontal_y"].as_array().unwrap();
        let bottom = (edges[edges.len()-1].as_f64().unwrap()
            + edges[edges.len()-2].as_f64().unwrap()) as f32 / 2.0;
        assert!(layout.pages[0].elements.iter().any(|e| matches!(&e.content,
            LayoutContent::TableBorder { x1, x2, y1, y2, .. }
                if (*y1-*y2).abs()<0.1 && (*y1-bottom).abs()<0.8
                    && *x1 < 100.0 && *x2 > 100.0)), "{name}: bottom={bottom}");
        for (i, page) in layout.pages.iter().enumerate() {
            let expected = case["pages"][i]["text"].as_str().unwrap()
                .lines().filter(|s| s.contains("\u{65e5}\u{672c}")).count();
            let actual = page.elements.iter().filter(|e| matches!(&e.content,
                LayoutContent::Text { text, .. } if text.contains("\u{65e5}\u{672c}"))).count();
            assert_eq!(actual, expected, "{name}: page {}", i+1);
        }
    }
}

#[test]
fn split_cells_repeat_only_their_declared_horizontal_edges() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/declared_fragment_edges");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["doc"].as_str().unwrap();
        let doc = crate::parser::parse_docx(&std::fs::read(fixtures.join(name)).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let pages = case["pages"].as_array().unwrap();
        assert_eq!(layout.pages.len(), pages.len(), "{name}");
        for (i, page) in layout.pages.iter().enumerate() {
            let mut actual: Vec<f32> = page.elements.iter().filter_map(|e| match &e.content {
                LayoutContent::TableBorder { x1, x2, y1, y2, .. }
                    if (*y1 - *y2).abs() < 0.1 && *x1 < 100.0 && *x2 > 100.0 => Some(*y1),
                _ => None,
            }).collect();
            actual.sort_by(f32::total_cmp);
            let mut expected: Vec<f32> = pages[i]["horizontal"].as_array().unwrap().iter()
                .map(|r| ((r[1].as_f64().unwrap() + r[3].as_f64().unwrap()) / 2.0) as f32)
                .collect();
            expected.sort_by(f32::total_cmp);
            assert_eq!(actual.len(), expected.len(), "{name}: page {}", i + 1);
            for (a, b) in actual.iter().zip(expected.iter()) {
                // Allow the existing sub-point cell-origin difference; edge presence is exact.
                assert!((a - b).abs() < 1.0, "{name}: page {}: {a} vs {b}", i + 1);
            }
        }
    }
}

#[test]
fn collapsed_row_top_is_repeated_unless_the_cell_suppresses_it() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/collapsed_fragment_top");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 2, "{name}");
        for x in [100.0_f32, 260.0] {
            let expected = case["pages"][1]["horizontal"].as_array().unwrap().iter().any(|r|
                r[0].as_f64().unwrap() < x as f64 && r[2].as_f64().unwrap() > x as f64
                    && r[1].as_f64().unwrap() < 21.0);
            let actual = layout.pages[1].elements.iter().any(|e| matches!(&e.content,
                LayoutContent::TableBorder { x1, x2, y1, y2, .. }
                    if (*y1 - *y2).abs() < 0.1 && (*y1 - 20.28).abs() < 0.8
                        && *x1 < x && *x2 > x));
            assert_eq!(actual, expected, "{name}: continuation top at x={x}");
        }
    }
}

#[test]
fn footer_inline_images_keep_word_positions_and_line_spacing() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/footer_images");
    let cases: serde_json::Value = serde_json::from_slice(&std::fs::read(fixtures.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(&std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 1, "{name}");
        let actual: Vec<_> = layout.pages[0].elements.iter().filter(|e| matches!(&e.content, LayoutContent::Image { .. })).collect();
        let expected = case["images"].as_array().unwrap();
        assert_eq!(actual.len(), expected.len(), "{name}");
        for (image, bbox) in actual.iter().zip(expected) {
            let coords = [image.x, image.y, image.x + image.width, image.y + image.height];
            for (axis, value) in coords.iter().enumerate() {
                let word = bbox[axis].as_f64().unwrap() as f32;
                assert!((value - word).abs() < 0.05, "{name} axis {axis}: {value} vs Word {word}");
            }
        }
    }
}

#[test]
fn oversized_first_line_does_not_create_an_empty_page() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/oversized_first_line");
    let cases: serde_json::Value = serde_json::from_slice(&std::fs::read(fixtures.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(&std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), case["pages"].as_u64().unwrap() as usize, "{name}");
        for (text, expected) in case["positions"].as_object().unwrap() {
            let actual: Vec<usize> = layout.pages.iter().enumerate().filter_map(|(i,p)|
                p.elements.iter().any(|e| matches!(&e.content, LayoutContent::Text { text: actual_text, .. } if actual_text == text)).then_some(i+1)).collect();
            let word: Vec<usize> = expected.as_array().unwrap().iter().map(|p| p.as_u64().unwrap() as usize).collect();
            assert_eq!(actual, word, "{name}: {text}");
        }
    }
}

#[test]
fn short_left_tab_gaps_land_on_the_word_tab_stop() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/short_tab_gap");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let text = layout.pages[0].elements.iter().find(|e|
            matches!(&e.content, LayoutContent::Text { text, .. } if text == "X")
        ).unwrap();
        let word_x = case["x"].as_f64().unwrap() as f32;
        assert!((text.x - word_x).abs() < 0.05, "{name}: {} vs Word {word_x}", text.x);
    }
}

#[test]
fn continuous_section_boundaries_keep_word_spacing() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/continuous_section_spacing");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let text_position = |wanted: &str| {
            layout.pages.iter().enumerate().find_map(|(i, p)| {
                p.elements.iter().find_map(|e| match &e.content {
                    LayoutContent::Text { text, .. } if text == wanted => Some((i + 1, e.y)),
                    _ => None,
                })
            }).unwrap()
        };
        match case["kind"].as_str().unwrap() {
            "gap" => {
                let (pre_page, pre_y) = text_position("PRE");
                let (item_page, item_y) = text_position("ITEM00");
                assert_eq!(pre_page, item_page, "{name}");
                let expected = case["gap"].as_f64().unwrap() as f32;
                assert!((item_y - pre_y - expected).abs() < 0.1, "{name}: first gap");
                if let Some(expected) = case["right_gap"].as_f64() {
                    let right_y = layout.pages[0].elements.iter().filter_map(|e| {
                        match &e.content {
                            LayoutContent::Text { text, .. }
                                if text.starts_with("ITEM") && e.x > 200.0 => Some(e.y),
                            _ => None,
                        }
                    }).reduce(f32::min).unwrap();
                    assert!((right_y - pre_y - expected as f32).abs() < 0.1,
                        "{name}: common column origin");
                }
            }
            "page_top" => {
                let (page, y) = text_position("ITEM00");
                assert_eq!(page, case["page"].as_u64().unwrap() as usize, "{name}");
                let expected = doc.pages[0].margin.top + case["before"].as_f64().unwrap() as f32;
                assert!((y - expected).abs() < 0.1, "{name}: {y} vs {expected}");
            }
            "image" => {
                let actual: Vec<_> = layout.pages.iter().enumerate().flat_map(|(i, p)| {
                    p.elements.iter().filter_map(move |e| {
                        matches!(&e.content, LayoutContent::Image { .. })
                            .then_some((i + 1, [e.x, e.y, e.x + e.width, e.y + e.height]))
                    })
                }).collect();
                let expected = case["images"].as_array().unwrap();
                assert_eq!(actual.len(), expected.len(), "{name}");
                for ((page, bbox), word) in actual.iter().zip(expected) {
                    assert_eq!(*page, word["page"].as_u64().unwrap() as usize, "{name}");
                    for (axis, value) in bbox.iter().enumerate() {
                        let expected = word["bbox"][axis].as_f64().unwrap() as f32;
                        assert!((value - expected).abs() < 0.05, "{name}: image axis {axis}");
                    }
                }
            }
            kind => panic!("unknown fixture kind: {kind}"),
        }
    }
}


#[test]
fn hidden_paragraph_keeps_follow_the_leading_paragraph() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/hidden_paragraph_keep");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (needle, key) in [("HEAD", "head_page"), ("NEXT", "next_page")] {
            let page = layout.pages.iter().position(|page| page.elements.iter().any(|e|
                matches!(&e.content, LayoutContent::Text { text, .. } if text.contains(needle))
            )).unwrap() + 1;
            assert_eq!(page, case[key].as_u64().unwrap() as usize, "{name}: {needle}");
        }
    }
}

#[test]
fn partial_table_margins_inherit_each_missing_edge() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/table_margin_inheritance");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let table = doc.pages.iter().flat_map(|p| &p.blocks).find_map(|b|
            if let crate::ir::Block::Table(t) = b { Some(t) } else { None }
        ).unwrap();
        let defaults = table.style.default_cell_margins.as_ref().unwrap();
        let cell = &table.rows[0].cells[0];
        let local = cell.margins.as_ref();
        for (edge, actual) in [
            ("left", local.and_then(|m| m.left).or(defaults.left)),
            ("right", local.and_then(|m| m.right).or(defaults.right)),
            ("top", local.and_then(|m| m.top).or(defaults.top)),
            ("bottom", local.and_then(|m| m.bottom).or(defaults.bottom)),
        ] {
            let expected = case["padding"][edge].as_f64().unwrap() as f32;
            assert!((actual.unwrap() - expected).abs() < 0.001, "{name}: {edge}");
        }
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut lines: Vec<(f32, String)> = Vec::new();
        for element in &layout.pages[0].elements {
            if let LayoutContent::Text { text, .. } = &element.content {
                if let Some((_, line)) = lines.iter_mut().find(|(y, _)| (element.y - *y).abs() < 0.01) {
                    line.push_str(text);
                } else {
                    lines.push((element.y, text.clone()));
                }
            }
        }
        lines.sort_by(|a,b| a.0.total_cmp(&b.0));
        let actual: Vec<_> = lines.iter().map(|(_, line)| line.trim_end()).collect();
        let expected: Vec<_> = case["lines"].as_array().unwrap().iter()
            .map(|l| l["text"].as_str().unwrap()).collect();
        assert_eq!(actual, expected, "{name}: Word line breaks");
    }
}

#[test]
fn floating_shape_contours_preserve_word_page_breaks() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/shape_wrap_pagination");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let actual: Vec<_> = layout.pages.iter().enumerate().filter_map(|(i, page)| {
            page.elements.iter().find(|e| e.paragraph_index == Some(1))
                .map(|e| (i + 1, e.y))
        }).collect();
        assert_eq!(actual.len(), 1, "{name}: target paragraph must occur once");
        assert_eq!(actual[0].0, case["page"].as_u64().unwrap() as usize, "{name}");
        assert!((actual[0].1 - case["y"].as_f64().unwrap() as f32).abs() < 0.8,
            "{name}: paragraph position");
    }
}

#[test]
fn nested_table_paragraphs_have_independent_widow_control() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/nested_paragraph_identity");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["pages"].as_object().unwrap() {
            let actual: Vec<_> = layout.pages.iter().enumerate().filter_map(|(i, page)|
                page.elements.iter().any(|e| matches!(&e.content,
                    LayoutContent::Text { text, .. } if text.contains(label)))
                    .then_some(i + 1)
            ).collect();
            assert_eq!(actual, vec![expected.as_u64().unwrap() as usize], "{name}: {label}");
        }
    }
}

#[test]
fn table_continuation_keeps_following_rows_below_images() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/table_image_continuation");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["text"].as_object().unwrap() {
            let (pi, element) = layout.pages.iter().enumerate().find_map(|(pi, p)|
                p.elements.iter().find(|e| matches!(&e.content,
                    LayoutContent::Text { text, .. } if text.contains(label)))
                    .map(|e| (pi, e))
            ).unwrap();
            assert_eq!(pi + 1, expected["page"].as_u64().unwrap() as usize, "{name}: {label}");
            if label == "FOLLOW" {
                let image_bottom = layout.pages[pi].elements.iter().filter_map(|e|
                    matches!(e.content, LayoutContent::Image { .. }).then_some(e.y + e.height)
                ).fold(f32::NEG_INFINITY, f32::max);
                assert!(image_bottom.is_finite() && element.y >= image_bottom - 0.01,
                    "{name}: following row overlaps image");
                let baseline = element.y + element.baseline_offset.expect("exact cell baseline");
                assert!((baseline - expected["baseline"].as_f64().unwrap() as f32).abs() < 0.8,
                    "{name}: following baseline {baseline} differs from Word {expected:?}");
            }
        }
    }
}

#[test]
fn word_flow_footnote_boundaries_preserve_body_and_notes() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/footnote_flow_boundary");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["places"].as_object().unwrap() {
            let actual: Vec<_> = layout.pages.iter().enumerate().flat_map(|(i, page)|
                page.elements.iter().filter_map(move |e| matches!(&e.content,
                    LayoutContent::Text { text, .. } if text.contains(label))
                    .then_some(i + 1))
            ).collect();
            assert_eq!(actual, vec![expected["page"].as_u64().unwrap() as usize],
                "{name}: {label} must be present exactly once on its Word page");
        }
    }
}

#[test]
fn word_flow_hyphenation_keeps_punctuation_and_following_tokens() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/hyphen_terminal_punctuation");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut lines: Vec<(usize, f32, String)> = Vec::new();
        for (pi, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    if let Some((_, _, line)) = lines.iter_mut().find(|(p, y, _)|
                        *p == pi && (e.y - *y).abs() < 0.01) {
                        line.push_str(text);
                    } else {
                        lines.push((pi, e.y, text.clone()));
                    }
                }
            }
        }
        lines.sort_by(|a,b| a.0.cmp(&b.0).then(a.1.total_cmp(&b.1)));
        let actual: Vec<_> = lines.iter().map(|(_, _, text)| text.trim_end()).collect();
        let expected: Vec<_> = case["lines"].as_array().unwrap().iter()
            .map(|line| line.as_str().unwrap()).collect();
        assert_eq!(actual, expected, "{name}: Word line breaks and complete text");
    }
}

#[test]
fn word_autofit_infeasible_preserves_cell_insets() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/autofit_infeasible");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let engine = LayoutEngine::for_document(&doc);
        let page = &doc.pages[0];
        let table = page.blocks.iter().find_map(|b| match b {
            Block::Table(table) => Some(table), _ => None,
        }).unwrap();
        let width = page.size.width - page.margin.left - page.margin.right;
        let actual = engine.resolve_table_col_widths_n(table, width, false);
        let expected = case["widths"].as_array().unwrap();
        assert_eq!(actual.len(), expected.len(), "{name}");
        for (column, (actual, expected)) in actual.iter().zip(expected).enumerate() {
            let expected = expected.as_f64().unwrap() as f32;
            assert!(*actual > 0.0 && (*actual - expected).abs() <= 0.05,
                "{name} column {column}: actual={actual}, Word={expected}");
        }
    }
}

#[test]
fn word_autofit_application_defaults_preserve_theme_inheritance() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/application_font_defaults");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let actual: Vec<_> = layout.pages.iter().flat_map(|page| page.elements.iter())
            .filter(|e| matches!(&e.content, LayoutContent::Text { text, .. } if text == "FOLLOW"))
            .map(|e| e.y).collect();
        let expected = case["places"].as_array().unwrap().iter()
            .find(|p| p["text"].as_str().unwrap().trim() == "FOLLOW")
            .unwrap()["y"].as_f64().unwrap() as f32;
        assert_eq!(actual.len(), 1, "{name}: FOLLOW must occur once");
        assert!((actual[0] - expected).abs() < 0.1,
            "{name}: actual={}, Word={expected}", actual[0]);
    }
}


#[test]
fn text_balance_grid_keeps_the_word_column_boundary_at_fractional_origins() {
    // A double-height heading followed by 24 grid rows. Word keeps twelve
    // rows on the left, including when the band follows one-column text.
    for top in [65.2, 265.2, 277.5] {
        let mut y = top;
        let mut rows: Vec<_> = (0..25).map(|i| {
            let height = if i == 0 { 40.55 } else { 20.55 };
            let row = text_row(0.0, y, height, i);
            y += height;
            row
        }).collect();
        LayoutEngine::rebalance_text_columns(
            &mut rows, top, &[0.0, 100.0], &[], 0.0, 14, Some(20.55),
        ).unwrap();
        assert_eq!(rows.iter().filter(|row| row.x == 0.0).count(), 12,
            "Word column boundary at band origin {top}");
        assert!((rows[12].y - top).abs() < 0.001);
        assert_eq!(rows.iter().map(|row| row.paragraph_index.unwrap()).collect::<Vec<_>>(),
            (0..25).collect::<Vec<_>>());
    }
}

#[test]
fn text_balance_grid_does_not_pull_a_tall_last_line_into_the_wrong_column() {
    // Word's ten-line controls use five left rows at the normal grid height,
    // but six when the final line requires two grid slots. Without an active
    // grid, even a smaller increase in the final line retains six left rows.
    for (pitch, height, last_height, expected) in [
        (Some(20.55), 20.55, 20.55, 5),
        (Some(20.55), 20.55, 40.65, 6),
        (None, 9.941, 11.93, 6),
    ] {
        let mut rows: Vec<_> = (0..10).map(|i| {
            text_row(0.0, i as f32 * height,
                if i == 9 { last_height } else { height }, i)
        }).collect();
        LayoutEngine::rebalance_text_columns(
            &mut rows, 0.0, &[0.0, 100.0], &[], 0.0, 14, pitch,
        ).unwrap();
        assert_eq!(rows.iter().filter(|row| row.x == 0.0).count(), expected,
            "active grid {pitch:?}, final height {last_height}");
    }
}


#[test]
fn case_expansion_fragments_retain_original_character_spans() {
    let map = CaseSourceMap::uppercase("aßZ");
    assert_eq!(map.fragment(0, "ASS"), Some((0, 2, "aß".into())));
    assert_eq!(map.fragment(1, "SS"), Some((1, 1, "ß".into())));
    assert_eq!(map.fragment(2, "S"), Some((1, 1, String::new())));
    assert_eq!(map.fragment(3, "Z"), Some((2, 1, "Z".into())));
    assert_eq!(map.fragment(0, "AX"), None);
    assert_eq!(map.fragment(usize::MAX, "S"), None);
}

#[test]
fn merged_cell_coordinates_do_not_use_available_height_as_page_stride() {
    for source_page in [0, 3] {
        let mut flow = MergedCellTextFlow {
            identity: std::sync::Arc::new(()), key: 0, start_row: 0,
            source_page, origin: 80.0, page_top: 10.0, page_height: 100.0,
            coordinate_stride: 400.0, header: 0.0, pad_top: 0.0, pad_bottom: 0.0,
            content_height: 50.0, element_count: 2,
            paragraphs: vec![MergedCellFlowParagraph {
                lines: vec![
                    MergedCellFlowLine { y: 0.0, fit: 20.0, elements: vec![(0, 0.0)] },
                    MergedCellFlowLine { y: 30.0, fit: 20.0, elements: vec![(1, 30.0)] },
                ], keep_lines: false, before: 0.0, after: 0.0,
            }], cuts: std::collections::BTreeMap::new(),
        };
        let (positions, end) = flow.paginate();
        assert_eq!(positions, vec![Some((source_page, 80.0)), Some((source_page + 1, 10.0))]);
        assert_eq!(end - (source_page + 1) as f32 * 400.0, 30.0);
        flow.cuts.insert(source_page, (90.0, 40.0));
        flow.cuts.insert(source_page + 1, (70.0, 10.0));
        let (positions, end) = flow.paginate();
        assert_eq!(positions, vec![Some((source_page + 1, 40.0)), Some((source_page + 2, 10.0))]);
        assert_eq!(end - (source_page + 2) as f32 * 400.0, 30.0);
    }
}


#[test]
fn exact_cell_baselines_match_word_across_fonts_and_sizes() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cell_exact_baselines");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap(),
    ).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap(),
        ).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 1, "{name}: page count");
        let e = layout.pages[0].elements.iter().find(|e| matches!(&e.content,
            LayoutContent::Text { text, .. } if text == "FOLLOW")).unwrap();
        let actual = e.y + e.baseline_offset.expect("exact cell baseline");
        let expected = case["baseline"].as_f64().unwrap() as f32;
        assert!((actual - expected).abs() < 0.2, "{name}: {actual} vs {expected}");
    }
}


#[test]
fn circled_number_fonts_follow_the_hint_in_mixed_script_documents() {
    for has_cjk in [false, true] {
        let mut engine = LayoutEngine::new();
        engine.doc_body_has_real_cjk = has_cjk;
        let para = ParagraphStyle::default();
        for hint in [None, Some(false), Some(true)] {
            let run = RunStyle {
                font_family: Some("Calibri".into()),
                font_family_east_asia: Some("Microsoft JhengHei Light".into()),
                font_size: Some(11.0),
                font_hint_east_asia: hint,
                ..RunStyle::default()
            };
            let expected = if hint == Some(true) {
                "Microsoft JhengHei Light"
            } else { "Calibri" };
            for text in ["①", "⑩", "A ① B", "A ⑩ B"] {
                assert_eq!(engine.resolve_font_family_for_text(text, &run, &para),
                    Some(expected), "{text}: hint {hint:?}, CJK {has_cjk}");
                assert_eq!(engine.metrics_for_text(text, &run, &para).family,
                    expected, "{text}: hint {hint:?}, CJK {has_cjk}");
            }
        }
    }
}


#[test]
fn no_grid_empty_lines_keep_the_paragraph_mark_font() {
    for has_cjk in [false, true] {
        let mut engine = LayoutEngine::new();
        engine.doc_body_has_real_cjk = has_cjk;
        for hint in [None, Some(false), Some(true)] {
            let para = ParagraphStyle {
                line_spacing: Some(1.0),
                ppr_rpr: Some(RunStyle {
                    font_family: Some("Calibri".into()),
                    font_family_east_asia: Some("Microsoft JhengHei Light".into()),
                    font_size: Some(11.0),
                    font_hint_east_asia: hint,
                    ..RunStyle::default()
                }),
                ..ParagraphStyle::default()
            };
            let actual = engine.line_height_for_line(
                &Line::default(), &para, 11.0, true, None, false);
            assert!((actual - 13.44).abs() < 0.2,
                "hint {hint:?}, CJK {has_cjk}: {actual}");
        }
    }
}


#[test]
fn cell_image_leading_excludes_paragraph_spacing() {
    fn image_in(blocks: &[Block]) -> Option<&crate::ir::Image> {
        for block in blocks {
            match block {
                Block::Image(image) => return Some(image),
                Block::Table(table) => {
                    for row in &table.rows {
                        for cell in &row.cells {
                            if let Some(image) = image_in(&cell.blocks) { return Some(image); }
                        }
                    }
                }
                _ => {}
            }
        }
        None
    }
    let doc = crate::parser::parse_docx(include_bytes!(
        "../../../../tests/fixtures/table_image_continuation/filler64_image40.docx"
    )).unwrap();
    let mut engine = LayoutEngine::for_document(&doc);
    let source = image_in(&doc.pages[0].blocks).expect("inline image");
    for has_cjk in [false, true] {
    engine.doc_body_has_real_cjk = has_cjk;
    for after in [0.0, 8.0, 16.0] {
        let mut image = source.clone();
        image.host_paragraph.as_mut().unwrap().style.space_after = Some(after);
        let actual = engine.s971_image_line_h(&image, 1.0e6, None, false);
        assert!((actual - 41.92).abs() < 0.2,
            "paragraph spacing {after}: image line {actual}");
    }
    }
}


#[test]
fn cell_space_width_does_not_change_line_height() {
    let base = crate::parser::parse_docx(include_bytes!(
        "../../../../tests/fixtures/table_image_continuation/filler64_image40.docx"
    )).unwrap();
    let table = base.pages[0].blocks.iter().find_map(|b| match b {
        Block::Table(t) => Some(t.clone()), _ => None,
    }).unwrap();
    let template = table.rows[0].cells[0].blocks.iter().find_map(|b| match b {
        Block::Paragraph(p) => Some(p.clone()), _ => None,
    }).unwrap();
    let mut reference = None;
    for position in 0..3 {
        for size in [10.0, 11.0, 20.0] {
            let mut p = template.clone();
            let font = RunStyle { font_family: Some("Arial".into()),
                font_size: Some(10.0), ..RunStyle::default() };
            p.style = ParagraphStyle { line_spacing: Some(1.0),
                space_before: Some(0.0), space_after: Some(0.0),
                ppr_rpr: Some(font.clone()), default_run_style: Some(font.clone()),
                ..ParagraphStyle::default() };
            let mut a = p.runs[0].clone(); a.style = font.clone(); a.text = "Before".into();
            let mut b = a.clone(); b.text = "After".into();
            let mut space = a.clone(); space.text = "  ".into(); space.style.font_size = Some(size);
            p.runs = match position { 0 => vec![space, a, b], 1 => vec![a, space, b], _ => vec![a, b, space] };
            let mut after = p.clone(); after.runs.truncate(1);
            after.runs[0].style = font; after.runs[0].text = "SENTINEL".into();
            let mut t = table.clone(); t.rows.truncate(1); t.rows[0].cells.truncate(1);
            t.rows[0].height = None; t.rows[0].height_rule = None;
            t.rows[0].cells[0].blocks = vec![Block::Paragraph(p)];
            let mut doc = base.clone();
            doc.pages[0].blocks = vec![Block::Table(t), Block::Paragraph(after)];
            let layout = LayoutEngine::for_document(&doc).layout(&doc);
            assert_eq!(layout.pages.len(), 1);
            let y = layout.pages[0].elements.iter().find(|e| matches!(&e.content,
                LayoutContent::Text { text, .. } if text == "SENTINEL")).unwrap().y;
            let expected = *reference.get_or_insert(y);
            assert!((y - expected).abs() < 0.01,
                "position {position}, space {size}: {y} vs {expected}");
        }
    }
}


#[test]
fn image_effects_survive_whole_row_page_moves() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/image_effect_wholepush");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for expected in case["images"].as_array().unwrap() {
            let images: Vec<_> = layout.pages.iter().enumerate().flat_map(|(i,p)|
                p.elements.iter().filter_map(move |e|
                    matches!(e.content, LayoutContent::Image { .. }).then_some((i+1,e))))
                .collect();
            assert_eq!(images.len(), 1, "{name}");
            assert_eq!(images[0].0, expected["page"].as_u64().unwrap() as usize, "{name}");
            assert!((images[0].1.y - expected["bbox"][1].as_f64().unwrap() as f32).abs() < 0.2,
                "{name}: image y {} vs {expected}", images[0].1.y);
        }
        let expected = &case["text"]["FOLLOW"];
        let (pi, e) = layout.pages.iter().enumerate().find_map(|(i,p)|
            p.elements.iter().find(|e| matches!(&e.content,
                LayoutContent::Text { text, .. } if text == "FOLLOW")).map(|e| (i+1,e))).unwrap();
        assert_eq!(pi, expected["page"].as_u64().unwrap() as usize, "{name}");
        let baseline = e.y + e.baseline_offset.expect("exact baseline");
        assert!((baseline - expected["baseline"].as_f64().unwrap() as f32).abs() < 0.8,
            "{name}: following baseline {baseline} vs {expected}");
    }
}

#[test]
fn outer_table_edges_preserve_word_row_origins_and_following_paragraph() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/table_outer_edges");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 1, "{name}");
        let y = |label: &str| layout.pages[0].elements.iter().find_map(|e| {
            matches!(&e.content, LayoutContent::Text { text, .. } if text == label)
                .then_some(e.y)
        }).unwrap();
        for label in ["ROW0", "ROW1", "AFTER"] {
            let word = case["word"][label]["y"].as_f64().unwrap()
                - case["word"]["BEFORE"]["y"].as_f64().unwrap();
            let actual = (y(label) - y("BEFORE")) as f64;
            assert!((actual - word).abs() < 0.2,
                "{name} {label}: Word {word}, actual {actual}");
        }
    }
}

#[test]
fn table_bottom_edge_participates_in_page_fit() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/table_outer_edge_fit");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), case["pages"].as_u64().unwrap() as usize, "{name}");
        for (label, expected) in case["paragraphs"].as_object().unwrap() {
            let actual = layout.pages.iter().position(|p| p.elements.iter().any(|e| {
                matches!(&e.content, LayoutContent::Text { text, .. } if text == label)
            })).unwrap() + 1;
            assert_eq!(actual, expected.as_u64().unwrap() as usize, "{name} {label}");
        }
    }
}

#[test]
fn leading_page_break_preserves_excess_before_spacing_in_both_scripts() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/leading_break_spacing");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), case["pages"].as_u64().unwrap() as usize, "{name}");
        let (page, element) = layout.pages.iter().enumerate().find_map(|(i, p)| {
            p.elements.iter().find(|e| matches!(&e.content,
                LayoutContent::Text { text, .. } if text == "SENTINEL"))
                .map(|e| (i + 1, e))
        }).unwrap();
        assert_eq!(page, case["word"]["page"].as_u64().unwrap() as usize, "{name}");
        // All fixtures use a 15pt exact line: Word's baseline is 12pt in.
        let baseline = element.y as f64 + 12.0;
        let expected = case["word"]["y"].as_f64().unwrap();
        assert!((baseline - expected).abs() < 0.2,
            "{name}: Word {expected}, actual {baseline}");
    }
}

#[test]
fn cjk_adjacent_spaces_follow_balance_setting_independent_of_run_boundaries() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/space_balance");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 1, "{name}");
        let x = |label: &str| layout.pages[0].elements.iter().find_map(|e| {
            matches!(&e.content, LayoutContent::Text { text, .. } if text == label)
                .then_some(e.x)
        }).unwrap();
        let actual = (x("\u{4e19}") - x("\u{7532}")) as f64;
        let word = case["word_extent"].as_f64().unwrap();
        assert!((actual - word).abs() < 0.2,
            "{name}: Word {word}, actual {actual}");
    }
}

#[test]
fn mixed_script_lines_advance_by_their_own_heights() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/body_mixed_line_heights");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 1, "{name}");
        let y = |label: &str| layout.pages[0].elements.iter().find_map(|e| {
            matches!(&e.content, LayoutContent::Text { text, .. } if text == label)
                .then_some(e.y)
        }).unwrap();
        let actual = (y("AFTER") - y("BEFORE")) as f64;
        let word = case["height"].as_f64().unwrap();
        assert!((actual - word).abs() < 0.2,
            "{name}: Word {word}, actual {actual}");
    }
}

#[test]
fn table_fragment_bottom_edge_participates_in_page_fit() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/table_fragment_border_fit");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["paragraphs"].as_object().unwrap() {
            let actual = layout.pages.iter().position(|p| p.elements.iter().any(|e| {
                matches!(&e.content, LayoutContent::Text { text, .. } if text == label)
            })).unwrap() + 1;
            assert_eq!(actual, expected.as_u64().unwrap() as usize, "{name} {label}");
        }
    }
}

#[test]
fn continuous_column_moves_preserve_required_line_capacity() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/continuous_column_capacity");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["paragraphs"].as_object().unwrap() {
            let actual = layout.pages.iter().position(|p| p.elements.iter().any(|e| {
                matches!(&e.content, LayoutContent::Text { text, .. } if text == label)
            })).unwrap() + 1;
            assert_eq!(actual, expected.as_u64().unwrap() as usize, "{name} {label}");
        }
    }
}

#[test]
fn table_fragment_border_capacity_is_not_counted_twice() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/table_fragment_border_tie");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["paragraphs"].as_object().unwrap() {
            let actual: Vec<usize> = layout.pages.iter().enumerate().filter_map(|(i, p)| {
                p.elements.iter().any(|e| {
                    matches!(&e.content, LayoutContent::Text { text, .. } if text == label)
                }).then_some(i + 1)
            }).collect();
            let expected: Vec<usize> = expected.as_array().unwrap().iter()
                .map(|p| p.as_u64().unwrap() as usize).collect();
            assert_eq!(actual, expected, "{name} {label}");
        }
    }
}

#[test]
fn table_continuation_capacity_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/table_continuation_capacity");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), case["pages"].as_u64().unwrap() as usize, "{name}");
        for (label, expected) in case["paragraphs"].as_object().unwrap() {
            let actual: Vec<usize> = layout.pages.iter().enumerate().filter_map(|(i, p)| {
                p.elements.iter().any(|e| matches!(&e.content,
                    LayoutContent::Text { text, .. } if text == label)).then_some(i + 1)
            }).collect();
            let expected: Vec<usize> = expected.as_array().unwrap().iter()
                .map(|p| p.as_u64().unwrap() as usize).collect();
            assert_eq!(actual, expected, "{name} {label}");
        }
        let after = layout.pages.iter().flat_map(|p| &p.elements)
            .find(|e| matches!(&e.content, LayoutContent::Text { text, .. } if text == "AFTER")).unwrap();
        // Every sentinel uses Arial 10pt in a 12pt exact line.
        let baseline = after.y as f64 + 9.6;
        let word = case["after"]["baseline"].as_f64().unwrap();
        assert!((baseline - word).abs() < 0.2, "{name}: AFTER Word {word}, actual {baseline}");
    }
}

#[test]
fn cell_hyphen_breaks_match_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cell_hyphen_breaks");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut lines: std::collections::BTreeMap<(usize, i32), Vec<(f32, String)>> = Default::default();
        let mut before = None;
        let mut after = None;
        for (pi, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    match text.as_str() {
                        "BEFORE" => before = Some(e.y),
                        "AFTER" => after = Some(e.y),
                        _ => lines.entry((pi, (e.y * 1000.0).round() as i32))
                            .or_default().push((e.x, text.clone())),
                    }
                }
            }
        }
        let actual: Vec<String> = lines.into_values().filter_map(|mut pieces| {
            pieces.sort_by(|a, b| a.0.total_cmp(&b.0));
            let text = pieces.into_iter().map(|(_, t)| t).collect::<String>()
                .replace('\u{a0}', " ").trim().to_string();
            (!text.is_empty()).then_some(text)
        }).collect();
        let expected: Vec<String> = case["lines"].as_array().unwrap().iter()
            .map(|t| t.as_str().unwrap().to_string()).collect();
        assert_eq!(actual, expected, "{name}");
        let extent = (after.unwrap() - before.unwrap()) as f64;
        let word = case["extent"].as_f64().unwrap();
        assert!((extent - word).abs() < 0.3, "{name}: Word {word}, actual {extent}");
    }
}

#[test]
fn cell_tab_line_height_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cell_tab_line_height");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 1, "{name}");
        let marker = |label: &str| layout.pages[0].elements.iter()
            .find_map(|e| matches!(&e.content, LayoutContent::Text { text, .. }
                if text == label).then_some(e.y)).unwrap();
        // Identical sentinel fonts cancel their baseline offset.
        let extent = (marker("AFTER") - marker("BEFORE")) as f64;
        let word = case["extent"].as_f64().unwrap();
        assert!((extent - word).abs() < 0.2, "{name}: Word {word}, actual {extent}");
    }
}

#[test]
fn section_float_band_stays_with_its_anchor() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/section_float_band");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 2, "{name}");
        for (label, word) in case["paragraphs"].as_object().unwrap() {
            let (page, y) = layout.pages.iter().enumerate().find_map(|(i, p)|
                p.elements.iter().find_map(|e| matches!(&e.content,
                    LayoutContent::Text { text, .. } if text == label).then_some((i + 1, e.y)))).unwrap();
            assert_eq!(page, word["page"].as_u64().unwrap() as usize, "{name} {label}");
            let baseline = y as f64 + 9.6;
            let expected = word["baseline"].as_f64().unwrap();
            assert!((baseline - expected).abs() < 0.2,
                "{name} {label}: Word {expected}, actual {baseline}");
        }
    }
}

#[test]
fn body_punctuation_language_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/body_punctuation_language");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut lines: std::collections::BTreeMap<(usize, i32), Vec<(f32, String)>> = Default::default();
        for (pi, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    match text.as_str() {
                        "BEFORE" => {},
                        "AFTER" => {},
                        _ => lines.entry((pi, (e.y * 1000.0).round() as i32))
                            .or_default().push((e.x, text.clone())),
                    }
                }
            }
        }
        let actual: Vec<String> = lines.into_values().filter_map(|mut pieces| {
            pieces.sort_by(|a, b| a.0.total_cmp(&b.0));
            let text = pieces.into_iter().map(|(_, t)| t).collect::<String>()
                .replace('\u{a0}', " ").trim().to_string();
            (!text.is_empty()).then_some(text)
        }).collect();
        let expected: Vec<String> = case["lines"].as_array().unwrap().iter()
            .map(|t| t.as_str().unwrap().to_string()).collect();
        if actual != expected {
            failures.push(format!("{name}: Word {expected:?}, actual {actual:?}"));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn footnote_table_carry_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/footnote_table_carry");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize { failures.push(format!("{name}: total pages Word {}, actual {}", case["pages"], layout.pages.len())); }
        for (label, expected) in case["paragraphs"].as_object().unwrap() {
            let actual: Vec<usize> = layout.pages.iter().enumerate().filter_map(|(i, p)| {
                // FIRST also occurs inside the footnote's sentence. Identify
                // the sentinel by its source body block as well as its text.
                p.elements.iter().any(|e| e.paragraph_index == Some(
                    if label == "FIRST" { 2 } else if label == "LAST" { 3 } else { 1 }
                ) && matches!(&e.content,
                    LayoutContent::Text { text, .. } if text == label)).then_some(i + 1)
            }).collect();
            let expected: Vec<usize> = expected.as_array().unwrap().iter()
                .map(|p| p.as_u64().unwrap() as usize).collect();
            if actual != expected { failures.push(format!("{name} {label}: Word {expected:?}, actual {actual:?}")); }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn cell_pair_boundaries_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cell_pair_boundaries");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut lines: std::collections::BTreeMap<(usize, i32), Vec<(f32, String)>> = Default::default();
        for (pi, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    match text.as_str() {
                        "BEFORE" => {},
                        "AFTER" => {},
                        _ => lines.entry((pi, (e.y * 1000.0).round() as i32))
                            .or_default().push((e.x, text.clone())),
                    }
                }
            }
        }
        let actual: Vec<String> = lines.into_values().filter_map(|mut pieces| {
            pieces.sort_by(|a, b| a.0.total_cmp(&b.0));
            let text = pieces.into_iter().map(|(_, t)| t).collect::<String>()
                .replace('\u{a0}', " ").trim().to_string();
            (!text.is_empty()).then_some(text)
        }).collect();
        let expected: Vec<String> = case["lines"].as_array().unwrap().iter()
            .map(|t| t.as_str().unwrap().to_string()).collect();
        if actual != expected {
            failures.push(format!("{name}: Word {expected:?}, actual {actual:?}"));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn multiple_empty_footer_paragraphs_reserve_the_word_stack() {
    let cases: &[(&[u8], usize)] = &[
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n1_ink0_h480.docx"), 1),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n1_ink0_h490.docx"), 1),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n1_ink0_h500.docx"), 2),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n1_ink1_h480.docx"), 1),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n1_ink1_h490.docx"), 1),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n1_ink1_h500.docx"), 2),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n2_ink0_h480.docx"), 1),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n2_ink0_h490.docx"), 2),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n2_ink0_h500.docx"), 2),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n2_ink1_h480.docx"), 1),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n2_ink1_h490.docx"), 2),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n2_ink1_h500.docx"), 2),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n3_ink0_h480.docx"), 2),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n3_ink0_h490.docx"), 2),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n3_ink0_h500.docx"), 2),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n3_ink1_h480.docx"), 2),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n3_ink1_h490.docx"), 2),
        (include_bytes!("../../../../tests/fixtures/empty_footer_stack/n3_ink1_h500.docx"), 2),
    ];
    for (case, (bytes, expected)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let pages: Vec<_> = layout.pages.iter().enumerate()
            .filter_map(|(i, page)| page.elements.iter().any(|e| matches!(&e.content,
                LayoutContent::Text { text, .. } if text == "TARGET")).then_some(i + 1))
            .collect();
        assert_eq!(pages, vec![*expected], "case {case}");
    }
}

#[test]
fn cell_style_spacing_is_preserved_in_mixed_script_documents() {
    let cases: &[(&[u8], f32)] = &[
        (include_bytes!("../../../../tests/fixtures/cell_inherited_spacing/ja0_custom0_s0.docx"), 12.96),
        (include_bytes!("../../../../tests/fixtures/cell_inherited_spacing/ja0_custom0_s1.8.docx"), 16.56),
        (include_bytes!("../../../../tests/fixtures/cell_inherited_spacing/ja0_custom1_s0.docx"), 12.96),
        (include_bytes!("../../../../tests/fixtures/cell_inherited_spacing/ja0_custom1_s1.8.docx"), 16.56),
        (include_bytes!("../../../../tests/fixtures/cell_inherited_spacing/ja1_custom0_s0.docx"), 12.84),
        (include_bytes!("../../../../tests/fixtures/cell_inherited_spacing/ja1_custom0_s1.8.docx"), 16.44),
        (include_bytes!("../../../../tests/fixtures/cell_inherited_spacing/ja1_custom1_s0.docx"), 12.84),
        (include_bytes!("../../../../tests/fixtures/cell_inherited_spacing/ja1_custom1_s1.8.docx"), 16.44),
    ];
    for (case, (bytes, expected)) in cases.iter().enumerate() {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), 1, "case {case}");
        let y = |label: &str| layout.pages[0].elements.iter()
            .find(|e| matches!(&e.content, LayoutContent::Text { text, .. } if text == label))
            .unwrap().y;
        let actual = y("XTWO") - y("XONE");
        assert!((actual - expected).abs() < 0.15,
            "case {case}: Word pitch {expected}, actual {actual}");
    }
}

#[test]
fn nbsp_lines_keep_the_run_height() {
    let cases: &[([&[u8]; 4], f32)] = &[
        ([include_bytes!("../../../../tests/fixtures/body_nbsp_height/nbsp_1.docx"), include_bytes!("../../../../tests/fixtures/body_nbsp_height/letter_1.docx"), include_bytes!("../../../../tests/fixtures/body_nbsp_height/space_1.docx"), include_bytes!("../../../../tests/fixtures/body_nbsp_height/empty_1.docx")], 2.28),
        ([include_bytes!("../../../../tests/fixtures/body_nbsp_height/nbsp_1.15.docx"), include_bytes!("../../../../tests/fixtures/body_nbsp_height/letter_1.15.docx"), include_bytes!("../../../../tests/fixtures/body_nbsp_height/space_1.15.docx"), include_bytes!("../../../../tests/fixtures/body_nbsp_height/empty_1.15.docx")], 2.64),
    ];
    for (case, (arms, expected_gap)) in cases.iter().enumerate() {
        let ys: Vec<f32> = arms.iter().map(|bytes| {
            let doc = crate::parser::parse_docx(bytes).unwrap();
            let layout = LayoutEngine::for_document(&doc).layout(&doc);
            assert_eq!(layout.pages.len(), 1);
            layout.pages[0].elements.iter().find(|e| matches!(&e.content,
                LayoutContent::Text { text, .. } if text == "AFTER")).unwrap().y
        }).collect();
        assert!((ys[0] - ys[1]).abs() < 0.01, "case {case}: NBSP differs from text");
        assert!((ys[2] - ys[3]).abs() < 0.01, "case {case}: ASCII space differs from empty");
        assert!((ys[2] - ys[0] - expected_gap).abs() < 0.08,
            "case {case}: Word gap {expected_gap}, actual {}", ys[2] - ys[0]);
    }
}

#[test]
fn center_tab_capacity_respects_preceding_text() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/center_tab_prefix");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut lines: std::collections::BTreeMap<(usize, i32), Vec<(f32, String)>> = Default::default();
        for (pi, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    match text.as_str() {
                        "BEFORE" => {},
                        "AFTER" => {},
                        _ => lines.entry((pi, (e.y * 1000.0).round() as i32))
                            .or_default().push((e.x, text.clone())),
                    }
                }
            }
        }
        let actual: Vec<String> = lines.into_values().filter_map(|mut pieces| {
            pieces.sort_by(|a, b| a.0.total_cmp(&b.0));
            let text = pieces.into_iter().map(|(_, t)| t).collect::<String>()
                .chars().filter(|c| !c.is_whitespace()).collect::<String>();
            (!text.is_empty()).then_some(text)
        }).collect();
        let expected: Vec<String> = case["lines"].as_array().unwrap().iter()
            .map(|t| t.as_str().unwrap().to_string()).collect();
        if actual != expected {
            failures.push(format!("{name}: Word {expected:?}, actual {actual:?}"));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}


#[test]
fn inline_images_fit_their_centered_grid_box() {
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/inline_image_grid_fit");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(fixtures.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(fixtures.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let images: Vec<_> = layout.pages.iter().enumerate().flat_map(|(i, p)|
            p.elements.iter().filter_map(move |e|
                matches!(e.content, LayoutContent::Image { .. }).then_some(i + 1))).collect();
        let expected_pages = case["pages"].as_u64().unwrap() as usize;
        let expected_image = case["image_page"].as_u64().unwrap() as usize;
        if layout.pages.len() != expected_pages || images != vec![expected_image] {
            failures.push(format!("{name}: Word pages {expected_pages}, image {expected_image}; actual pages {}, images {images:?}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn paragraph_widow_style_precedes_document_defaults() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/widow_style_inheritance");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["positions"].as_object().unwrap() {
            let page = layout.pages.iter().position(|p| p.elements.iter().any(|e|
                matches!(&e.content, LayoutContent::Text { text, .. } if text.contains(label))));
            if page.map(|p| p + 1) != Some(expected.as_u64().unwrap() as usize) {
                failures.push(format!("{name} {label}: Word {expected}, actual {page:?} (zero-based)"));
            }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: wrong total page count {}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn drawingml_textbox_page_fit_respects_compatibility_mode() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/dml_textbox_page_fit");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["positions"].as_object().unwrap() {
            let page = layout.pages.iter().position(|p| p.elements.iter().any(|e|
                matches!(&e.content, LayoutContent::Text { text, .. } if text.contains(label))));
            if page.map(|p| p + 1) != Some(expected["page"].as_u64().unwrap() as usize) {
                failures.push(format!("{name} {label}: Word {expected}, actual {page:?} (zero-based)"));
            }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: wrong total page count {}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn empty_paragraphs_use_the_declared_default_style() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/empty_default_style_id");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["positions"].as_object().unwrap() {
            let page = layout.pages.iter().position(|p| p.elements.iter().any(|e|
                matches!(&e.content, LayoutContent::Text { text, .. } if text.contains(label))));
            if page.map(|p| p + 1) != Some(expected["page"].as_u64().unwrap() as usize) {
                failures.push(format!("{name} {label}: Word {expected}, actual {page:?} (zero-based)"));
            }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: wrong total page count {}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn cell_first_paragraph_orphan_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cell_first_orphan");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["positions"].as_object().unwrap() {
            let page = layout.pages.iter().position(|p| p.elements.iter().any(|e|
                matches!(&e.content, LayoutContent::Text { text, .. } if text.contains(label))));
            if page.map(|p| p + 1) != Some(expected["page"].as_u64().unwrap() as usize) {
                failures.push(format!("{name} {label}: Word {expected}, actual {page:?} (zero-based)"));
            }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: wrong total page count {}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn header_intermediate_border_reserves_body_space() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/header_intermediate_border");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["positions"].as_object().unwrap() {
            let page = layout.pages.iter().position(|p| p.elements.iter().any(|e|
                matches!(&e.content, LayoutContent::Text { text, .. } if text.contains(label))));
            if page.map(|p| p + 1) != Some(expected["page"].as_u64().unwrap() as usize) {
                failures.push(format!("{name} {label}: Word {expected}, actual {page:?} (zero-based)"));
            }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: wrong total page count {}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn header_image_line_spacing_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/header_image_leading");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["positions"].as_object().unwrap() {
            let page = layout.pages.iter().position(|p| p.elements.iter().any(|e|
                matches!(&e.content, LayoutContent::Text { text, .. } if text.contains(label))));
            if page.map(|p| p + 1) != Some(expected["page"].as_u64().unwrap() as usize) {
                failures.push(format!("{name} {label}: Word {expected}, actual {page:?} (zero-based)"));
            }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: wrong total page count {}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn cell_proportional_grid_widths_match_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cell_proportional_grid");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut lines: std::collections::BTreeMap<(usize, i32), Vec<(f32, String)>> = Default::default();
        for (pi, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    match text.as_str() {
                        "BEFORE" => {},
                        "AFTER" => {},
                        _ => lines.entry((pi, (e.y * 1000.0).round() as i32))
                            .or_default().push((e.x, text.clone())),
                    }
                }
            }
        }
        let actual: Vec<String> = lines.into_values().filter_map(|mut pieces| {
            pieces.sort_by(|a, b| a.0.total_cmp(&b.0));
            let text = pieces.into_iter().map(|(_, t)| t).collect::<String>()
                .replace('\u{a0}', " ").trim().to_string();
            (!text.is_empty()).then_some(text)
        }).collect();
        let expected: Vec<String> = case["lines"].as_array().unwrap().iter()
            .map(|t| t.as_str().unwrap().to_string()).collect();
        if actual != expected {
            failures.push(format!("{name}: Word {expected:?}, actual {actual:?}"));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}


#[test]
fn cell_grid_space_balance_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cell_grid_space_balance");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut lines: std::collections::BTreeMap<(usize, i32), Vec<(f32, String)>> = Default::default();
        for (pi, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    match text.as_str() {
                        "BEFORE" => {},
                        "AFTER" => {},
                        _ => lines.entry((pi, (e.y * 1000.0).round() as i32))
                            .or_default().push((e.x, text.clone())),
                    }
                }
            }
        }
        let actual: Vec<String> = lines.into_values().filter_map(|mut pieces| {
            pieces.sort_by(|a, b| a.0.total_cmp(&b.0));
            let text = pieces.into_iter().map(|(_, t)| t).collect::<String>()
                .replace('\u{a0}', " ").replace('\u{3000}', " ").trim().to_string();
            (!text.is_empty()).then_some(text)
        }).collect();
        let expected: Vec<String> = case["lines"].as_array().unwrap().iter()
            .map(|t| t.as_str().unwrap().to_string()).collect();
        if actual != expected {
            failures.push(format!("{name}: Word {expected:?}, actual {actual:?}"));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}


#[test]
fn floating_table_keep_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/floating_table_keep");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["positions"].as_object().unwrap() {
            let page = layout.pages.iter().position(|p| p.elements.iter().any(|e|
                matches!(&e.content, LayoutContent::Text { text, .. } if text.contains(label))));
            if page.map(|p| p + 1) != Some(expected["page"].as_u64().unwrap() as usize) {
                failures.push(format!("{name} {label}: Word {expected}, actual {page:?} (zero-based)"));
            }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: wrong total page count {}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}


#[test]
fn latin_keep_lastline_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/latin_keep_lastline");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["positions"].as_object().unwrap() {
            let page = layout.pages.iter().position(|p| p.elements.iter().any(|e|
                matches!(&e.content, LayoutContent::Text { text, .. } if text.contains(label))));
            if page.map(|p| p + 1) != Some(expected["page"].as_u64().unwrap() as usize) {
                failures.push(format!("{name} {label}: Word {expected}, actual {page:?} (zero-based)"));
            }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: wrong total page count {}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}


#[test]
fn word_punctuation_credit_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/word_punctuation_credit");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut lines: std::collections::BTreeMap<(usize, i32), Vec<(f32, String)>> = Default::default();
        for (pi, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    match text.as_str() {
                        "BEFORE" => {},
                        "AFTER" => {},
                        _ => lines.entry((pi, (e.y * 1000.0).round() as i32))
                            .or_default().push((e.x, text.clone())),
                    }
                }
            }
        }
        let actual: Vec<String> = lines.into_values().filter_map(|mut pieces| {
            pieces.sort_by(|a, b| a.0.total_cmp(&b.0));
            let text = pieces.into_iter().map(|(_, t)| t).collect::<String>()
                .replace('\u{a0}', " ").trim().to_string();
            (!text.is_empty()).then(|| text.chars().filter(|c| !c.is_whitespace()).collect::<String>())
        }).collect();
        let expected: Vec<String> = case["lines"].as_array().unwrap().iter()
            .map(|t| t.as_str().unwrap().to_string()).collect();
        if actual != expected {
            failures.push(format!("{name}: Word {expected:?}, actual {actual:?}"));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}


#[test]
fn cell_atleast_grid_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/cell_atleast_grid");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let texts: Vec<String> = layout.pages.iter().map(|page| page.elements.iter()
            .filter_map(|e| match &e.content { LayoutContent::Text { text, .. } => Some(text.as_str()), _ => None })
            .collect()).collect();
        for (label, expected) in case["positions"].as_object().unwrap() {
            let page = texts.iter().position(|text| text.contains(label)).map(|p| p + 1);
            if page != Some(expected["page"].as_u64().unwrap() as usize) {
                failures.push(format!("{name} {label}: Word {expected}, actual {page:?}"));
            }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: wrong total page count {}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn floating_table_following_flow_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/floating_table_following_flow");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let locate = |label: &str| layout.pages.iter().enumerate().find_map(|(pi, page)|
            page.elements.iter().find_map(|e| match &e.content {
                LayoutContent::Text { text, .. } if text == label => Some((pi + 1, e.y + e.text_y_off)), _ => None,
            }));
        let (tp, ty) = locate("ROW0LINE0").unwrap();
        for label in ["BEFORE", "AFTER"] {
            let (ap, ay) = locate(label).unwrap();
            let expected = case["positions"][label]["y"].as_f64().unwrap()
                - case["positions"]["ROW0LINE0"]["y"].as_f64().unwrap();
            if ap != tp || (f64::from(ay - ty) - expected).abs() > 0.75 {
                failures.push(format!("{name} {label}: Word relative y {expected}, actual {} pages {ap}/{tp}", ay - ty));
            }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: wrong page count {}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn inherited_frames_match_word_equivalent_direct_documents() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/frame_style_inheritance");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    for variant in ["inherit", "wrap", "zero", "space"] {
        let name = |mode: &str| format!("{variant}_{mode}");
        let truth = |mode: &str| cases.as_array().unwrap().iter()
            .find(|c| c["name"] == name(mode)).unwrap();
        assert_eq!(truth("styled")["positions"], truth("inline")["positions"]);
        let parse = |mode: &str| crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{}.docx", name(mode)))).unwrap()).unwrap();
        let styled = parse("styled");
        let direct = parse("inline");
        let first_frame = |doc: &crate::ir::Document| {
            doc.pages.iter().flat_map(|p| &p.blocks).find_map(|b| match b {
                crate::ir::Block::Paragraph(p) => p.style.frame_pr.as_ref(), _ => None,
            }).map(|fp| serde_json::to_value(fp).unwrap()).unwrap()
        };
        assert_eq!(first_frame(&styled), first_frame(&direct), "{variant}: resolved frame");
        let render = |doc: &crate::ir::Document| {
            let layout = LayoutEngine::for_document(doc).layout(doc);
            let positions: Vec<_> = layout.pages.iter().enumerate().flat_map(|(pi, page)| {
                page.elements.iter().filter_map(move |e| match &e.content {
                    LayoutContent::Text { text, .. } => Some((pi, text.clone(), e.x, e.y + e.text_y_off)),
                    _ => None,
                })
            }).collect();
            (layout.pages.len(), positions)
        };
        assert_eq!(render(&styled), render(&direct), "{variant}: Word-equivalent layout");
    }
}

#[test]
fn frame_exclusion_across_columns_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/frame_two_columns");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(&std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let first = layout.pages.iter().flat_map(|page| &page.elements).find(|e|
            matches!(&e.content, LayoutContent::Text { text, .. } if text == "BODY00")).unwrap();
        let native_origin = first.y + first.text_y_off;
        let word_origin = case["positions"]["BODY00"]["y"].as_f64().unwrap();
        for (label, expected) in case["positions"].as_object().unwrap() {
            if !label.starts_with("BODY") { continue; }
            let actual = layout.pages.iter().enumerate().find_map(|(pi, page)|
                page.elements.iter().find_map(|e| match &e.content {
                    LayoutContent::Text { text, .. } if text == label => {
                        let dy = f64::from(e.y + e.text_y_off - native_origin);
                        let expected_dy = expected["y"].as_f64().unwrap() - word_origin;
                        assert!((dy - expected_dy).abs() <= 0.75,
                            "{name} {label}: Word relative y {expected_dy:.2}, actual {dy:.2}");
                        Some((pi + 1, e.x > 250.0))
                    },
                    _ => None,
                }));
            let word = (expected["page"].as_u64().unwrap() as usize, expected["x"].as_f64().unwrap() > 250.0);
            if actual != Some(word) { failures.push(format!("{name} {label}: Word {word:?}, actual {actual:?}")); }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: wrong page count {}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn negative_floating_table_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/negative_floating_table");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        for (label, expected) in case["positions"].as_object().unwrap() {
            let actual = layout.pages.iter().enumerate().find_map(|(pi, page)|
                page.elements.iter().find_map(|e| match &e.content {
                    LayoutContent::Text { text, .. } if text == label => Some(pi + 1),
                    _ => None,
                }));
            if actual != Some(expected["page"].as_u64().unwrap() as usize) {
                failures.push(format!("{name} {label}: Word {}, actual {actual:?}", expected["page"]));
            }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: wrong page count {}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn legacy_kerning_space_capacity_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/legacy_kerning_space_capacity");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut lines: std::collections::BTreeMap<(usize, i32), Vec<(f32, String)>> = Default::default();
        for (pi, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    match text.as_str() {
                        "BEFORE" => {},
                        "AFTER" => {},
                        _ => lines.entry((pi, (e.y * 1000.0).round() as i32))
                            .or_default().push((e.x, text.clone())),
                    }
                }
            }
        }
        let actual: Vec<String> = lines.into_values().filter_map(|mut pieces| {
            pieces.sort_by(|a, b| a.0.total_cmp(&b.0));
            let text = pieces.into_iter().map(|(_, t)| t).collect::<String>()
                .replace('\u{a0}', " ").trim().to_string();
            (!text.is_empty()).then(|| text.chars().filter(|c| !c.is_whitespace()).collect::<String>())
        }).collect();
        let expected: Vec<String> = case["lines"].as_array().unwrap().iter()
            .map(|t| t.as_str().unwrap().to_string()).collect();
        if actual != expected {
            failures.push(format!("{name}: Word {expected:?}, actual {actual:?}"));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}


#[test]
fn floating_table_narrow_lane_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/floating_table_narrow_lane");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let native_origin = layout.pages.iter().flat_map(|p| &p.elements).find_map(|e|
            match &e.content { LayoutContent::Text { text, .. } if text == "ROW0LINE0" => Some(e.y + e.text_y_off), _ => None }).unwrap();
        let word_origin = case["positions"]["ROW0LINE0"]["y"].as_f64().unwrap();
        for (label, expected) in case["positions"].as_object().unwrap() {
            let actual = layout.pages.iter().enumerate().find_map(|(pi, page)|
                page.elements.iter().find_map(|e| match &e.content {
                    LayoutContent::Text { text, .. } if text == label => Some((pi + 1, e.y + e.text_y_off)),
                    _ => None,
                }));
            if actual.map_or(true, |(pi, y)| pi != expected["page"].as_u64().unwrap() as usize || (f64::from(y - native_origin) - (expected["y"].as_f64().unwrap() - word_origin)).abs() > 0.75) {
                let nearby: Vec<_> = layout.pages.iter().flat_map(|p| &p.elements)
                    .filter_map(|e| match &e.content { LayoutContent::Text { text, .. }
                        if !text.starts_with("ROW") && text != "BEFORE" => Some((text.as_str(), e.x, e.y, e.width)), _ => None }).collect();
                failures.push(format!("{name} {label}: Word {}, actual {actual:?}, body {nearby:?}", expected["page"]));
            }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: wrong page count {}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn text_frame_group_anchor_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/text_frame_group_anchor");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let native_origin = layout.pages.iter().flat_map(|p| &p.elements).find_map(|e|
            match &e.content { LayoutContent::Text { text, .. } if text == "FRAME" => Some(e.y + e.text_y_off), _ => None }).unwrap();
        let word_origin = case["positions"]["FRAME"]["y"].as_f64().unwrap();
        for (label, expected) in case["positions"].as_object().unwrap() {
            let actual = layout.pages.iter().enumerate().find_map(|(pi, page)| {
                let mut rows: std::collections::BTreeMap<(Option<usize>, i64), Vec<&LayoutElement>> = std::collections::BTreeMap::new();
                for e in &page.elements { if matches!(&e.content, LayoutContent::Text { .. }) {
                    rows.entry((e.paragraph_index, (e.y * 1000.0).round() as i64)).or_default().push(e);
                }}
                rows.values_mut().find_map(|row| {
                    row.sort_by(|a,b| a.x.total_cmp(&b.x));
                    let text: String = row.iter().filter_map(|e| match &e.content { LayoutContent::Text { text, .. } => Some(text.as_str()), _ => None }).collect();
                    (text.trim() == label).then(|| (pi + 1, row[0].y + row[0].text_y_off))
                })
            });
            if actual.map_or(true, |(pi, y)| pi != expected["page"].as_u64().unwrap() as usize || (f64::from(y - native_origin) - (expected["y"].as_f64().unwrap() - word_origin)).abs() > 0.75) {
                let nearby: Vec<_> = layout.pages.iter().flat_map(|p| &p.elements)
                    .filter_map(|e| match &e.content { LayoutContent::Text { text, .. }
                        if !text.starts_with("ROW") && text != "BEFORE" => Some((text.as_str(), e.x, e.y, e.width)), _ => None }).collect();
                failures.push(format!("{name} {label}: Word {}, actual {actual:?}, body {nearby:?}", expected["page"]));
            }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: wrong page count {}", layout.pages.len()));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}


#[test]
fn latin_no_grid_bottom_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/latin_no_grid_bottom");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: page count {}", layout.pages.len()));
        }
        for (label, expected) in case["positions"].as_object().unwrap() {
            let actual = layout.pages.iter().enumerate().find_map(|(pi, page)| {
                page.elements.iter().any(|e| matches!(&e.content,
                    LayoutContent::Text { text, .. } if text == label)).then_some(pi + 1)
            });
            let want = expected["page"].as_u64().unwrap() as usize;
            if actual != Some(want) {
                failures.push(format!("{name} {label}: Word {want}, actual {actual:?}"));
            }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}


#[test]
fn latin_table_line_pitch_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/latin_table_line_pitch");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut positions = std::collections::BTreeMap::new();
        for (pi, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    positions.insert(text.clone(), (pi + 1, e.y));
                }
            }
        }
        if layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: page count {}", layout.pages.len()));
        }
        let origin = positions.get("LINE0").expect("first line missing").1;
        let expected_origin = case["positions"]["LINE0"]["y"].as_f64().unwrap() as f32;
        for i in 0..16 {
            let label = format!("LINE{i}");
            let expected = &case["positions"][&label];
            let want_page = expected["page"].as_u64().unwrap() as usize;
            let want_y = expected["y"].as_f64().unwrap() as f32 - expected_origin;
            if !positions.get(&label).is_some_and(|&(page, y)|
                page == want_page && (y - origin - want_y).abs() < 0.2)
            {
                failures.push(format!("{name} {label}: Word {want_page}/{want_y}, actual {:?}", positions.get(&label)));
            }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}


#[test]
fn image_paragraph_page_break_before_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/image_paragraph_page_break_before");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut labels = std::collections::BTreeMap::new();
        let mut images = Vec::new();
        for (pi, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                match &e.content {
                    LayoutContent::Text { text, .. } => { labels.insert(text.clone(), pi + 1); },
                    LayoutContent::Image { .. } => images.push(pi + 1),
                    _ => {},
                }
            }
        }
        let expected_images: Vec<usize> = case["images"].as_array().unwrap().iter()
            .map(|p| p.as_u64().unwrap() as usize).collect();
        if images != expected_images || layout.pages.len() != case["pages"].as_u64().unwrap() as usize {
            failures.push(format!("{name}: images {images:?}, pages {}", layout.pages.len()));
        }
        for (text, page) in case["labels"].as_object().unwrap() {
            let want = page.as_u64().unwrap() as usize;
            if labels.get(text) != Some(&want) {
                failures.push(format!("{name} {text}: Word {want}, actual {:?}", labels.get(text)));
            }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}


#[test]
fn header_text_frame_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/header_text_frame");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut positions = std::collections::BTreeMap::new();
        for (pi, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    positions.insert(text.clone(), (pi + 1, e.x, e.y + e.text_y_off));
                }
            }
        }
        let origin = positions.get("HEADER").expect("header missing").2;
        let word_origin = case["positions"]["HEADER"]["y"].as_f64().unwrap() as f32;
        for (label, expected) in case["positions"].as_object().unwrap() {
            let want_page = expected["page"].as_u64().unwrap() as usize;
            let want_x = expected["x"].as_f64().unwrap() as f32;
            let want_y = expected["y"].as_f64().unwrap() as f32 - word_origin;
            if !positions.get(label).is_some_and(|&(page, x, y)|
                page == want_page && (x - want_x).abs() < 0.3 && (y - origin - want_y).abs() < 0.2)
            {
                failures.push(format!("{name} {label}: Word ({want_page}, {want_x}, relative {want_y}), actual {:?}, origin {origin}", positions.get(label)));
            }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn inline_picture_position_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/inline_picture_position");
    let cases: serde_json::Value = serde_json::from_slice(&std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for c in cases.as_array().unwrap() {
        let name = c["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(&std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let find = |label: &str| layout.pages.iter().enumerate().find_map(|(pi, page)| {
            page.elements.iter().find_map(|e| match &e.content {
                LayoutContent::Text { text, .. } if text.trim() == label => Some((pi + 1, e.y + e.text_y_off)),
                _ => None,
            })
        });
        let before = find("BEFORE").unwrap();
        let picture = layout.pages.iter().enumerate().find_map(|(pi, page)| {
            page.elements.iter().find_map(|e| matches!(&e.content, LayoutContent::Image { .. }).then_some((pi + 1, e.y)))
        });
        let image_y = c["image"]["bbox"][1].as_f64().unwrap();
        match picture {
            Some((page, y)) if page == c["image"]["page"].as_u64().unwrap() as usize
                && (y as f64 - image_y).abs() < 0.15 => {},
            actual => failures.push(format!("{name} image: Word y {image_y}, actual {actual:?}")),
        }
        for label in ["AFTER", "X", "(1)"] {
            let expected = &c["positions"][label];
            let delta = expected["y"].as_f64().unwrap() - c["positions"]["BEFORE"]["y"].as_f64().unwrap();
            match find(label) {
                Some((page, y)) if page == expected["page"].as_u64().unwrap() as usize
                    && ((y - before.1) as f64 - delta).abs() < 0.25 => {},
                actual => failures.push(format!("{name} {label}: expected relative y {delta}, actual {actual:?}, before {before:?}")),
            }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}


#[test]
fn synthetic_ms_gothic_bold_width_matches_word() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/synthetic_ms_gothic_bold_width");
    let cases: serde_json::Value = serde_json::from_slice(
        &std::fs::read(root.join("word.json")).unwrap()).unwrap();
    let mut failures = Vec::new();
    for case in cases.as_array().unwrap() {
        let name = case["name"].as_str().unwrap();
        let doc = crate::parser::parse_docx(
            &std::fs::read(root.join(format!("{name}.docx"))).unwrap()).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        let mut lines: std::collections::BTreeMap<(usize, i32), Vec<(f32, String)>> = Default::default();
        for (pi, page) in layout.pages.iter().enumerate() {
            for e in &page.elements {
                if let LayoutContent::Text { text, .. } = &e.content {
                    match text.as_str() {
                        "BEFORE" => {},
                        "AFTER" => {},
                        _ => lines.entry((pi, (e.y * 1000.0).round() as i32))
                            .or_default().push((e.x, text.clone())),
                    }
                }
            }
        }
        let actual: Vec<String> = lines.into_values().filter_map(|mut pieces| {
            pieces.sort_by(|a, b| a.0.total_cmp(&b.0));
            let text = pieces.into_iter().map(|(_, t)| t).collect::<String>()
                .replace('\u{a0}', " ").trim().to_string();
            (!text.is_empty()).then(|| text.chars().filter(|c| !c.is_whitespace()).collect::<String>())
        }).collect();
        let expected: Vec<String> = case["lines"].as_array().unwrap().iter()
            .map(|t| t.as_str().unwrap().to_string()).collect();
        if actual != expected {
            failures.push(format!("{name}: Word {expected:?}, actual {actual:?}"));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

#[test]
fn keep_next_across_last_column_matches_word() {
    let cases: &[(&str, &[u8], usize, usize, f32, f32, usize)] = &[
        ("first-line_widow0_remaining15", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line_widow0_remaining15.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line_widow0_remaining20", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line_widow0_remaining20.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line_widow0_remaining25", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line_widow0_remaining25.docx"), 2, 1, 30.0000, 245.2500, 1),
        ("first-line_widow0_remaining30", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line_widow0_remaining30.docx"), 2, 1, 30.0000, 240.0000, 1),
        ("first-line_widow0_remaining35", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line_widow0_remaining35.docx"), 2, 1, 30.0000, 234.7500, 1),
        ("first-line_widow0_remaining40", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line_widow0_remaining40.docx"), 2, 1, 30.0000, 230.2500, 1),
        ("first-line_widow1_remaining15", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line_widow1_remaining15.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line_widow1_remaining20", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line_widow1_remaining20.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line_widow1_remaining25", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line_widow1_remaining25.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line_widow1_remaining30", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line_widow1_remaining30.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line_widow1_remaining35", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line_widow1_remaining35.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line_widow1_remaining40", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line_widow1_remaining40.docx"), 2, 1, 30.0000, 230.2500, 1),
        ("first-line-auto_widow0_remaining15", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-auto_widow0_remaining15.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-auto_widow0_remaining20", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-auto_widow0_remaining20.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-auto_widow0_remaining25", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-auto_widow0_remaining25.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-auto_widow0_remaining30", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-auto_widow0_remaining30.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-auto_widow0_remaining35", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-auto_widow0_remaining35.docx"), 2, 1, 30.0000, 240.7500, 1),
        ("first-line-auto_widow0_remaining40", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-auto_widow0_remaining40.docx"), 2, 1, 30.0000, 236.2500, 1),
        ("first-line-auto_widow1_remaining15", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-auto_widow1_remaining15.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-auto_widow1_remaining20", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-auto_widow1_remaining20.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-auto_widow1_remaining25", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-auto_widow1_remaining25.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-auto_widow1_remaining30", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-auto_widow1_remaining30.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-auto_widow1_remaining35", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-auto_widow1_remaining35.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-auto_widow1_remaining40", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-auto_widow1_remaining40.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-short_widow0_remaining15", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-short_widow0_remaining15.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-short_widow0_remaining20", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-short_widow0_remaining20.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-short_widow0_remaining25", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-short_widow0_remaining25.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-short_widow0_remaining30", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-short_widow0_remaining30.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-short_widow0_remaining35", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-short_widow0_remaining35.docx"), 2, 1, 30.0000, 240.7500, 1),
        ("first-line-short_widow0_remaining40", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-short_widow0_remaining40.docx"), 2, 1, 30.0000, 236.2500, 1),
        ("first-line-short_widow1_remaining15", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-short_widow1_remaining15.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-short_widow1_remaining20", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-short_widow1_remaining20.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-short_widow1_remaining25", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-short_widow1_remaining25.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-short_widow1_remaining30", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-short_widow1_remaining30.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-short_widow1_remaining35", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-short_widow1_remaining35.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("first-line-short_widow1_remaining40", include_bytes!("../../../../tests/fixtures/keepnext_first_line/first-line-short_widow1_remaining40.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("last-column_widow0_remaining15", include_bytes!("../../../../tests/fixtures/keepnext_first_line/last-column_widow0_remaining15.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("last-column_widow0_remaining20", include_bytes!("../../../../tests/fixtures/keepnext_first_line/last-column_widow0_remaining20.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("last-column_widow0_remaining25", include_bytes!("../../../../tests/fixtures/keepnext_first_line/last-column_widow0_remaining25.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("last-column_widow0_remaining30", include_bytes!("../../../../tests/fixtures/keepnext_first_line/last-column_widow0_remaining30.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("last-column_widow0_remaining35", include_bytes!("../../../../tests/fixtures/keepnext_first_line/last-column_widow0_remaining35.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("last-column_widow0_remaining40", include_bytes!("../../../../tests/fixtures/keepnext_first_line/last-column_widow0_remaining40.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("last-column_widow1_remaining15", include_bytes!("../../../../tests/fixtures/keepnext_first_line/last-column_widow1_remaining15.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("last-column_widow1_remaining20", include_bytes!("../../../../tests/fixtures/keepnext_first_line/last-column_widow1_remaining20.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("last-column_widow1_remaining25", include_bytes!("../../../../tests/fixtures/keepnext_first_line/last-column_widow1_remaining25.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("last-column_widow1_remaining30", include_bytes!("../../../../tests/fixtures/keepnext_first_line/last-column_widow1_remaining30.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("last-column_widow1_remaining35", include_bytes!("../../../../tests/fixtures/keepnext_first_line/last-column_widow1_remaining35.docx"), 2, 2, 30.0000, 30.0000, 2),
        ("last-column_widow1_remaining40", include_bytes!("../../../../tests/fixtures/keepnext_first_line/last-column_widow1_remaining40.docx"), 2, 2, 30.0000, 30.0000, 2),
    ];
    for &(name, bytes, page_count, heading_page, heading_x, heading_y, body_page) in cases {
        let doc = crate::parser::parse_docx(bytes).unwrap();
        let layout = LayoutEngine::for_document(&doc).layout(&doc);
        assert_eq!(layout.pages.len(), page_count, "{name}: page count");
        let heading_index = if name.starts_with("last-column_") { 2 } else { 1 };
        let headings: Vec<_> = layout.pages.iter().enumerate().flat_map(|(page, p)| {
            p.elements.iter().filter_map(move |e| match &e.content {
                LayoutContent::Text { text, .. } if e.paragraph_index == Some(heading_index) && !text.is_empty() => Some((page + 1, e)),
                _ => None,
            })
        }).collect();
        let heading_text: String = headings.iter().filter_map(|(_, e)| match &e.content {
            LayoutContent::Text { text, .. } => Some(text.as_str()),
            _ => None,
        }).collect();
        assert_eq!(heading_text.trim(), "Kept heading", "{name}: heading text");
        assert!(headings.iter().all(|(p, _)| *p == heading_page), "{name}: heading fragments on expected page");
        let (page, heading) = headings[0];
        assert_eq!(page, heading_page, "{name}: heading page");
        assert!((heading.x - heading_x).abs() <= 0.5, "{name}: heading x {} expected {heading_x}", heading.x);
        assert!((heading.y - heading_y).abs() <= 0.5, "{name}: heading y {} expected {heading_y}", heading.y);
        let first_body_page = layout.pages.iter().position(|p| p.elements.iter().any(|e| {
            matches!(&e.content, LayoutContent::Text { text, .. } if e.paragraph_index == Some(heading_index + 1) && !text.is_empty())
        })).map(|p| p + 1);
        assert_eq!(first_body_page, Some(body_page), "{name}: first body line");
    }
}
