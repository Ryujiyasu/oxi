// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Contextual paragraph spacing uses source order across table containers.
//! Resolve boundary before/after spacing on the layout copy, so measurement,
//! continuation replay and painting all consume the same point value.

use crate::ir::{Block, Document, Paragraph};

struct Story<'a> {
    default_style: Option<&'a str>,
    previous: Option<(usize, Option<String>)>,
    next_container: usize,
    reverse: bool,
    row_boundary_style: Option<&'a str>,
}

impl<'a> Story<'a> {
    fn new(default_style: Option<&'a str>) -> Self {
        Self {
            default_style,
            previous: None,
            next_container: 1,
            reverse: false,
            row_boundary_style: None,
        }
    }

    fn before(default_style: Option<&'a str>, end_style: Option<&'a str>) -> Self {
        Self {
            row_boundary_style: end_style,
            ..Self::new(default_style)
        }
    }

    fn after(default_style: Option<&'a str>) -> Self {
        Self {
            reverse: true,
            ..Self::new(default_style)
        }
    }

    fn after_for(default_style: Option<&'a str>, end_style: Option<&'a str>) -> Self {
        Self {
            row_boundary_style: end_style,
            ..Self::after(default_style)
        }
    }

    fn paragraph(&mut self, para: &mut Paragraph, container: usize, in_cell: bool) -> bool {
        let style = para.style.style_id.as_deref().or(self.default_style);
        let suppress = in_cell
            && para.style.contextual_spacing
            && self
                .previous
                .as_ref()
                .is_some_and(|(previous_container, previous_style)| {
                    *previous_container != container && previous_style.as_deref() == style
                });
        if suppress {
            // Each side uses the current paragraph's contextual flag.
            // Different cells never exchange numeric spacing values.
            if self.reverse {
                para.style.space_after = Some(0.0);
                para.style.space_after_from_doc_defaults = false;
                para.style.after_lines = None;
                para.style.after_autospacing = false;
                para.style.after_autospacing_off = true;
            } else {
                para.style.space_before = Some(0.0);
                para.style.space_before_from_doc_defaults = false;
                para.style.before_lines = None;
                para.style.before_autospacing = false;
                para.style.before_autospacing_off = true;
            }
        }
        self.previous = Some((container, style.map(str::to_owned)));
        suppress
    }

    fn blocks(&mut self, blocks: &mut [Block], container: usize, in_cell: bool) {
        for block in source_order(blocks, self.reverse) {
            match block {
                Block::Paragraph(para) => {
                    self.paragraph(para, container, in_cell);
                }
                Block::Table(table) => {
                    for (row_index, row) in source_order(&mut table.rows, self.reverse).enumerate()
                    {
                        if self.reverse {
                            // A row's final cell compares its after spacing
                            // with the end mark, not a paragraph outside the row.
                            // Resolve that after spacing on the layout copy,
                            // including the shortcut nested-row estimates.
                            self.previous = self
                                .row_boundary_style
                                .map(|id| (usize::MAX, Some(id.to_owned())));
                        } else if row_index > 0 {
                            // The next row starts after a built-in Normal end
                            // mark. Source and vertical neighbors do not change
                            // its identity; use the document's resolved style ID.
                            self.previous = self
                                .row_boundary_style
                                .map(|id| (usize::MAX, Some(id.to_owned())));
                        }
                        for cell in source_order(&mut row.cells, self.reverse) {
                            let child_container = self.next_container;
                            self.next_container += 1;
                            self.blocks(&mut cell.blocks, child_container, true);
                            // A text box is a separate story, not the next
                            // paragraph of its anchor's surrounding cell.
                            for text_box in &mut cell.cell_text_boxes {
                                let mut story = if self.reverse {
                                    Story::after_for(self.default_style, self.row_boundary_style)
                                } else {
                                    Story::before(self.default_style, self.row_boundary_style)
                                };
                                story.blocks(&mut text_box.blocks, 0, false);
                            }
                        }
                    }
                    if self.reverse {
                        // A table separates the preceding paragraph from the
                        // paragraphs inside it (also for nested tables).
                        self.previous = None;
                    } else {
                        // The paragraph after a nested table follows its end
                        // mark, not the last source paragraph inside a cell.
                        self.previous = self
                            .row_boundary_style
                            .map(|id| (usize::MAX, Some(id.to_owned())));
                    }
                }
                Block::Image(image) => {
                    if let Some(host) = image.host_paragraph.as_mut() {
                        if self.paragraph(host, container, in_cell) {
                            if self.reverse {
                                image.paragraph_space_after = 0.0;
                            } else {
                                image.paragraph_space_before = 0.0;
                            }
                        }
                    } else {
                        // Without a host paragraph, its source style is not
                        // available; do not infer adjacency through it.
                        self.previous = None;
                    }
                }
                Block::Math(_) | Block::UnsupportedElement(_) => {
                    self.previous = None;
                }
            }
        }
    }
}

fn source_order<T>(items: &mut [T], reverse: bool) -> Box<dyn Iterator<Item = &mut T> + '_> {
    if reverse {
        Box::new(items.iter_mut().rev())
    } else {
        Box::new(items.iter_mut())
    }
}

pub(super) fn apply(doc: &mut Document, end_style: Option<&str>) {
    let default_style = doc
        .styles
        .default_paragraph_style_id
        .clone()
        .or_else(|| end_style.map(str::to_owned));
    let mut body = Story::before(default_style.as_deref(), end_style);
    for page in &mut doc.pages {
        body.blocks(&mut page.blocks, 0, false);
        for blocks in [
            &mut page.header,
            &mut page.footer,
            &mut page.header_first,
            &mut page.footer_first,
            &mut page.header_even,
            &mut page.footer_even,
        ] {
            Story::before(default_style.as_deref(), end_style).blocks(blocks, 0, false);
        }
        for note in &mut page.footnotes {
            Story::before(default_style.as_deref(), end_style).blocks(&mut note.blocks, 0, false);
        }
        for note in &mut page.endnotes {
            Story::before(default_style.as_deref(), end_style).blocks(&mut note.blocks, 0, false);
        }
        for text_box in &mut page.text_boxes {
            Story::before(default_style.as_deref(), end_style).blocks(
                &mut text_box.blocks,
                0,
                false,
            );
        }
    }
    let mut body = Story::after_for(default_style.as_deref(), end_style);
    for page in doc.pages.iter_mut().rev() {
        body.blocks(&mut page.blocks, 0, false);
        for blocks in [
            &mut page.header,
            &mut page.footer,
            &mut page.header_first,
            &mut page.footer_first,
            &mut page.header_even,
            &mut page.footer_even,
        ] {
            Story::after_for(default_style.as_deref(), end_style).blocks(blocks, 0, false);
        }
        for note in &mut page.footnotes {
            Story::after_for(default_style.as_deref(), end_style).blocks(
                &mut note.blocks,
                0,
                false,
            );
        }
        for note in &mut page.endnotes {
            Story::after_for(default_style.as_deref(), end_style).blocks(
                &mut note.blocks,
                0,
                false,
            );
        }
        for text_box in &mut page.text_boxes {
            Story::after_for(default_style.as_deref(), end_style).blocks(
                &mut text_box.blocks,
                0,
                false,
            );
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::{Alignment, ParagraphStyle, Table, TableCell, TableRow, TableStyle};

    fn para(style: Option<&str>, contextual: bool) -> Paragraph {
        Paragraph {
            runs: Vec::new(),
            style: ParagraphStyle {
                style_id: style.map(str::to_owned),
                contextual_spacing: contextual,
                space_before: Some(12.0),
                space_after: Some(6.0),
                before_lines: Some(100.0),
                ..Default::default()
            },
            alignment: Alignment::Left,
            shapes: Vec::new(),
            ppr_change: None,
            paragraph_mark_revision: None,
        }
    }

    fn cell(blocks: Vec<Block>) -> TableCell {
        serde_json::from_value(serde_json::json!({"blocks": blocks, "width": null})).unwrap()
    }

    fn table(cells: Vec<TableCell>) -> Block {
        let row: TableRow = serde_json::from_value(serde_json::json!({"cells": cells})).unwrap();
        Block::Table(Table {
            rows: vec![row],
            style: TableStyle::default(),
            grid_columns: Vec::new(),
        })
    }

    #[test]
    fn different_cells_use_current_context_and_preserve_numeric_after() {
        let mut story = Story::new(Some("Normal"));
        let mut first = para(Some("A"), false);
        let mut current = para(Some("A"), true);
        story.paragraph(&mut first, 1, true);
        assert!(story.paragraph(&mut current, 2, true));
        assert_eq!(current.style.space_before, Some(0.0));
        assert_eq!(current.style.before_lines, None);
        assert_eq!(first.style.space_after, Some(6.0));
        assert_eq!(current.style.space_after, Some(6.0));
        let mut no_context = para(Some("A"), false);
        assert!(!story.paragraph(&mut no_context, 3, true));
        assert_eq!(no_context.style.space_before, Some(12.0));
    }

    #[test]
    fn previous_cell_last_paragraph_controls_adjacency() {
        let mut story = Story::new(Some("Normal"));
        story.paragraph(&mut para(Some("A"), true), 1, true);
        story.paragraph(&mut para(Some("B"), true), 1, true);
        let mut current = para(Some("A"), true);
        assert!(!story.paragraph(&mut current, 2, true));
        assert_eq!(current.style.space_before, Some(12.0));
        story.paragraph(&mut para(Some("A"), true), 2, true);
        assert!(story.paragraph(&mut current, 3, true));
    }

    #[test]
    fn nested_entry_uses_source_and_exit_uses_end_mark() {
        let mut blocks = vec![
            Block::Paragraph(para(Some("A"), true)),
            table(vec![cell(vec![
                Block::Paragraph(para(Some("A"), true)),
                table(vec![cell(vec![Block::Paragraph(para(Some("B"), true))])]),
                Block::Paragraph(para(Some("B"), true)),
            ])]),
        ];
        Story::before(Some("Normal"), Some("Normal")).blocks(&mut blocks, 0, false);
        let Block::Table(outer) = &blocks[1] else {
            panic!()
        };
        let Block::Paragraph(first) = &outer.rows[0].cells[0].blocks[0] else {
            panic!()
        };
        let Block::Table(inner) = &outer.rows[0].cells[0].blocks[1] else {
            panic!()
        };
        let Block::Paragraph(inner_first) = &inner.rows[0].cells[0].blocks[0] else {
            panic!()
        };
        let Block::Paragraph(after) = &outer.rows[0].cells[0].blocks[2] else {
            panic!()
        };
        assert_eq!(first.style.space_before, Some(0.0));
        assert_eq!(inner_first.style.space_before, Some(12.0));
        assert_eq!(after.style.space_before, Some(12.0));
    }

    #[test]
    fn implicit_default_and_named_default_match_without_story_or_body_bleed() {
        let mut first_story = Story::new(Some("Normal"));
        first_story.paragraph(&mut para(None, true), 0, false);
        let mut current = para(Some("Normal"), true);
        assert!(first_story.paragraph(&mut current, 1, true));
        let mut second_story = Story::new(Some("Normal"));
        current = para(Some("Normal"), true);
        assert!(!second_story.paragraph(&mut current, 1, true));
        assert_eq!(current.style.space_before, Some(12.0));
        assert!(!first_story.paragraph(&mut current, 0, false));
        assert_eq!(current.style.space_before, Some(12.0));
    }

    #[test]
    fn after_uses_current_flag_with_noncontextual_following_cell() {
        let mut story = Story::after(Some("Normal"));
        story.paragraph(&mut para(Some("A"), false), 3, true);
        let mut current = para(Some("A"), true);
        current.style.after_lines = Some(100.0);
        assert!(story.paragraph(&mut current, 2, true));
        assert_eq!(current.style.space_after, Some(0.0));
        assert_eq!(current.style.after_lines, None);
        assert_eq!(current.style.space_before, Some(12.0));
        let mut prior = para(Some("A"), false);
        assert!(!story.paragraph(&mut prior, 1, true));
        assert_eq!(prior.style.space_after, Some(6.0));
    }

    #[test]
    fn reverse_walk_uses_next_cells_first_paragraph() {
        let mut blocks = vec![table(vec![
            cell(vec![Block::Paragraph(para(Some("A"), true))]),
            cell(vec![
                Block::Paragraph(para(Some("B"), false)),
                Block::Paragraph(para(Some("A"), false)),
            ]),
        ])];
        Story::after(Some("Normal")).blocks(&mut blocks, 0, false);
        let Block::Table(table) = &blocks[0] else {
            panic!()
        };
        let Block::Paragraph(first) = &table.rows[0].cells[0].blocks[0] else {
            panic!()
        };
        assert_eq!(first.style.space_after, Some(6.0));
    }

    #[test]
    fn after_matches_implicit_default_and_stops_at_story_end() {
        let mut story = Story::after(Some("Normal"));
        let mut last = para(None, true);
        assert!(!story.paragraph(&mut last, 2, true));
        assert_eq!(last.style.space_after, Some(6.0));
        let mut first = para(Some("Normal"), true);
        assert!(story.paragraph(&mut first, 1, true));
        let mut isolated = Story::after(Some("Normal"));
        first = para(Some("Normal"), true);
        assert!(!isolated.paragraph(&mut first, 1, true));
        assert_eq!(first.style.space_after, Some(6.0));
    }

    #[test]
    fn new_row_uses_end_mark_style_instead_of_source_or_above_cell() {
        let mut blocks = vec![table(vec![cell(vec![Block::Paragraph(para(
            Some("A"),
            true,
        ))])])];
        let Block::Table(table) = &mut blocks[0] else {
            panic!()
        };
        let row: TableRow = serde_json::from_value(serde_json::json!({"cells": [cell(vec![
            Block::Paragraph(para(Some("A"), true)),
        ])]}))
        .unwrap();
        table.rows.push(row);
        Story::before(Some("A"), Some("B")).blocks(&mut blocks, 0, false);
        let Block::Table(table) = &blocks[0] else {
            panic!()
        };
        let Block::Paragraph(second) = &table.rows[1].cells[0].blocks[0] else {
            panic!()
        };
        assert_eq!(second.style.space_before, Some(12.0));
        let mut blocks = vec![Block::Table(table.clone())];
        Story::before(Some("A"), Some("A")).blocks(&mut blocks, 0, false);
        let Block::Table(table) = &blocks[0] else {
            panic!()
        };
        let Block::Paragraph(second) = &table.rows[1].cells[0].blocks[0] else {
            panic!()
        };
        assert_eq!(second.style.space_before, Some(0.0));
    }

    #[test]
    fn after_does_not_cross_row_or_nested_table_boundaries() {
        let mut blocks = vec![table(vec![cell(vec![Block::Paragraph(para(
            Some("A"),
            true,
        ))])])];
        let Block::Table(first_table) = &mut blocks[0] else {
            panic!()
        };
        first_table.rows.push(
            serde_json::from_value(serde_json::json!({"cells": [cell(vec![
                Block::Paragraph(para(Some("A"), false)),
            ])]}))
            .unwrap(),
        );
        Story::after(Some("A")).blocks(&mut blocks, 0, false);
        let Block::Table(outer) = &blocks[0] else {
            panic!()
        };
        let Block::Paragraph(first) = &outer.rows[0].cells[0].blocks[0] else {
            panic!()
        };
        assert_eq!(first.style.space_after, Some(6.0));
        let mut nested = vec![
            Block::Paragraph(para(Some("A"), true)),
            table(vec![cell(vec![Block::Paragraph(para(Some("A"), false))])]),
        ];
        Story::after(Some("A")).blocks(&mut nested, 1, true);
        let Block::Paragraph(prior) = &nested[0] else {
            panic!()
        };
        assert_eq!(prior.style.space_after, Some(6.0));
    }

    #[test]
    fn nested_exit_compares_resolved_end_style_even_when_last_cell_differs() {
        let mut blocks = vec![
            table(vec![cell(vec![Block::Paragraph(para(
                Some("Other"),
                false,
            ))])]),
            Block::Paragraph(para(Some("RenamedEndStyle"), true)),
        ];
        Story::before(Some("OtherDefault"), Some("RenamedEndStyle")).blocks(&mut blocks, 1, true);
        let Block::Paragraph(after) = &blocks[1] else {
            panic!()
        };
        assert_eq!(after.style.space_before, Some(0.0));
        assert_eq!(after.style.space_after, Some(6.0));
    }

    #[test]
    fn normal_middle_cell_keeps_after_for_different_next_style_but_row_end_does_not() {
        let mut blocks = vec![table(vec![
            cell(vec![Block::Paragraph(para(Some("EndStyle"), true))]),
            cell(vec![Block::Paragraph(para(Some("Other"), false))]),
        ])];
        Story::after_for(Some("CustomDefault"), Some("EndStyle")).blocks(&mut blocks, 0, false);
        let Block::Table(table) = &blocks[0] else {
            panic!()
        };
        let Block::Paragraph(first) = &table.rows[0].cells[0].blocks[0] else {
            panic!()
        };
        assert_eq!(first.style.space_after, Some(6.0));
        let mut row_end = vec![Block::Paragraph(para(Some("EndStyle"), true))];
        let mut story = Story::after_for(Some("CustomDefault"), Some("EndStyle"));
        story.previous = Some((usize::MAX, Some("EndStyle".to_owned())));
        story.blocks(&mut row_end, 1, true);
        let Block::Paragraph(last) = &row_end[0] else {
            panic!()
        };
        assert_eq!(last.style.space_after, Some(0.0));
        assert_eq!(last.style.space_before, Some(12.0));
    }
}
