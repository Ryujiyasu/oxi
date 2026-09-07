// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use super::*;

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
