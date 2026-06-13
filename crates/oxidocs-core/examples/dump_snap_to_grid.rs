// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// S344 diagnostic: dump every paragraph's snap_to_grid + text prefix.
// Used to verify whether b35123 paragraphs other than i=89 actually have
// snap_to_grid=false in Oxi's parsed view (vs XML inspection showing
// default true).

use oxidocs_core::{ir::Block, parse_docx};

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: dump_snap_to_grid <docx_file>");
        std::process::exit(1);
    }
    let data = std::fs::read(&args[1]).expect("failed to read file");
    let doc = parse_docx(&data).expect("parse failed");

    let mut idx = 0usize;
    for (pi, page) in doc.pages.iter().enumerate() {
        // Walk blocks. Recurse into table cells.
        for block in &page.blocks {
            match block {
                Block::Paragraph(p) => {
                    idx += 1;
                    let preview: String = p.runs.iter()
                        .flat_map(|r| r.text.chars())
                        .take(30)
                        .collect();
                    println!(
                        "p_idx={} page={} snap_to_grid={} in_table=false text={:?}",
                        idx, pi + 1, p.style.snap_to_grid, preview
                    );
                }
                Block::Table(t) => {
                    for (ri, row) in t.rows.iter().enumerate() {
                        for (ci, cell) in row.cells.iter().enumerate() {
                            for cb in &cell.blocks {
                                if let Block::Paragraph(p) = cb {
                                    idx += 1;
                                    let preview: String = p.runs.iter()
                                        .flat_map(|r| r.text.chars())
                                        .take(30)
                                        .collect();
                                    // S344 diagnostic: show ALL runs' font_size
                                    let all_fs: Vec<String> = p.runs.iter()
                                        .map(|r| format!("{:?}", r.style.font_size))
                                        .collect();
                                    let ppr_fs = p.style.ppr_rpr.as_ref().and_then(|r| r.font_size);
                                    println!(
                                        "p_idx={} page={} snap_to_grid={} in_table=true row={} col={} all_fs={:?} ppr_fs={:?} text={:?}",
                                        idx, pi + 1, p.style.snap_to_grid, ri, ci, all_fs, ppr_fs, preview
                                    );
                                }
                            }
                        }
                    }
                }
                _ => {}
            }
        }
    }
}
