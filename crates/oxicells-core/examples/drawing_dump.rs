// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! What the parser makes of a workbook's drawings.
//!
//! Points it at a directory and it lists every workbook whose first sheet
//! draws anything, with the first few anchors: which cell each hangs from,
//! how far into it, and whether it came out as a picture, a shape or a chart.
//!
//! ```text
//! cargo run --release -p oxicells-core --example drawing_dump
//! cargo run --release -p oxicells-core --example drawing_dump -- path\to\book.xlsx
//! ```

use oxicells_core::ir::DrawingKind;
use std::path::Path;

fn main() {
    let target = std::env::args().nth(1).unwrap_or_else(|| {
        r"tools\golden-test\documents\xlsx".to_string()
    });
    let path = Path::new(&target);
    let books: Vec<_> = if path.is_dir() {
        let mut held: Vec<_> = std::fs::read_dir(path)
            .expect("the directory can be read")
            .flatten()
            .map(|entry| entry.path())
            .filter(|path| path.extension().is_some_and(|kind| kind == "xlsx"))
            .collect();
        held.sort();
        held
    } else {
        vec![path.to_path_buf()]
    };

    for book in books {
        let Ok(bytes) = std::fs::read(&book) else { continue };
        let Ok(workbook) = oxicells_core::parser::parse_xlsx_preserving_values(&bytes) else {
            continue;
        };
        let Some(sheet) = workbook.sheets.first() else { continue };
        if sheet.drawings.is_empty() {
            continue;
        }
        let count = |wanted: fn(&DrawingKind) -> bool| {
            sheet.drawings.iter().filter(|drawn| wanted(&drawn.kind)).count()
        };
        println!(
            "{:<46} {:>3} drawings ({} pictures, {} shapes, {} charts)",
            book.file_stem().unwrap_or_default().to_string_lossy(),
            sheet.drawings.len(),
            count(|kind| matches!(kind, DrawingKind::Picture { .. })),
            count(|kind| matches!(kind, DrawingKind::Shape(_))),
            count(|kind| matches!(kind, DrawingKind::Chart(_))),
        );
        for drawn in sheet.drawings.iter().take(3) {
            let what = match &drawn.kind {
                DrawingKind::Picture { bytes } => format!("picture, {} bytes", bytes.len()),
                DrawingKind::Shape(shape) => format!(
                    "{} filled {:?} ruled {:?}",
                    shape.geometry,
                    shape.fill,
                    shape.line.as_ref().map(|line| &line.color)
                ),
                DrawingKind::Chart(chart) => {
                    let mut said = format!(
                        "{} chart, {} series over {} categories, plot {:?}",
                        chart.kind,
                        chart.series.len(),
                        chart.categories.len(),
                        chart.plot.map(|frame| (frame.x, frame.y, frame.w, frame.h)),
                    );
                    for series in &chart.series {
                        said.push_str(&format!(
                            "\n        {:<10} {:>3} points, ruled {:?} {:?}, {} marked, {} labelled",
                            series.name,
                            series.values.len(),
                            series.line.as_ref().map(|line| &line.color),
                            series.line.as_ref().and_then(|line| line.dash.as_deref()),
                            series.points.len(),
                            series.labels.len(),
                        ));
                    }
                    if let Some(axis) = &chart.value_axis {
                        said.push_str(&format!(
                            "\n        value axis {:?}..{:?} by {:?}, ticks {}, {:?}",
                            axis.min, axis.max, axis.major_unit, axis.major_tick,
                            axis.number_format,
                        ));
                    }
                    said
                }
                other => format!("{other:?}"),
            };
            println!(
                "    from ({}, {}) + ({}, {}) to {:?} ext {:?}  {what}",
                drawn.from.col,
                drawn.from.row,
                drawn.from.col_off,
                drawn.from.row_off,
                drawn.to.map(|anchor| (anchor.col, anchor.row)),
                drawn.extent
            );
        }
    }
}
