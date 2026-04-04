use oxidocs_core::{parse_docx, layout};
use std::io::Write;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: layout_json <docx_file> [--structure]");
        std::process::exit(1);
    }
    let structure_mode = args.iter().any(|a| a == "--structure");

    let data = std::fs::read(&args[1]).expect("failed to read file");
    let doc = parse_docx(&data).expect("parse failed");
    let engine = layout::LayoutEngine::for_document(&doc);
    let result = engine.layout(&doc);

    let stdout = std::io::stdout();
    let mut out = stdout.lock();

    if structure_mode {
        // Structural output: paragraphs with line info, table rows
        output_structure(&result, &mut out);
    } else {
        // Legacy flat output for backward compatibility
        output_flat(&result, &mut out);
    }
}

fn output_flat(result: &layout::LayoutResult, out: &mut impl Write) {
    for (pi, page) in result.pages.iter().enumerate() {
        writeln!(out, "PAGE\t{}\t{}\t{}", pi, page.width, page.height).unwrap();
        for elem in &page.elements {
            match &elem.content {
                layout::LayoutContent::Text { text, font_size, font_family, bold, italic, underline, strikethrough, color, highlight, .. } => {
                    let ff = font_family.as_deref().unwrap_or("Calibri");
                    let col = color.as_deref().unwrap_or("#000000");
                    let hl = highlight.as_deref().unwrap_or("");
                    writeln!(out, "TEXT\t{:.3}\t{:.3}\t{:.1}\t{:.3}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}",
                        elem.x, elem.y, elem.width, elem.height,
                        font_size, ff, *bold as u8, *italic as u8, *underline as u8, *strikethrough as u8,
                        col, hl
                    ).unwrap();
                    writeln!(out, "T\t{}", text).unwrap();
                }
                layout::LayoutContent::TableBorder { x1, y1, x2, y2, ref color, width } => {
                    let col = color.as_deref().unwrap_or("#000000");
                    writeln!(out, "BORDER\t{:.3}\t{:.3}\t{:.3}\t{:.3}\t{}\t{}", x1, y1, x2, y2, col, width).unwrap();
                }
                layout::LayoutContent::CellShading { ref color } => {
                    writeln!(out, "BG\t{:.3}\t{:.3}\t{:.1}\t{:.3}\t{}", elem.x, elem.y, elem.width, elem.height, color).unwrap();
                }
                layout::LayoutContent::Image { .. } => {
                    writeln!(out, "IMG\t{:.3}\t{:.3}\t{:.1}\t{:.3}", elem.x, elem.y, elem.width, elem.height).unwrap();
                }
                layout::LayoutContent::BoxRect { ref fill, ref stroke_color, stroke_width, corner_radius } => {
                    let f = fill.as_deref().unwrap_or("");
                    let s = stroke_color.as_deref().unwrap_or("");
                    writeln!(out, "BOX\t{:.3}\t{:.3}\t{:.1}\t{:.3}\t{}\t{}\t{}\t{}", elem.x, elem.y, elem.width, elem.height, f, s, stroke_width, corner_radius).unwrap();
                }
                layout::LayoutContent::ClipStart | layout::LayoutContent::ClipEnd => {}
                layout::LayoutContent::PresetShape { .. } => {}
            }
        }
    }
}

fn output_structure(result: &layout::LayoutResult, out: &mut impl Write) {
    // Reconstruct paragraph/table structure from layout elements.
    // Uses paragraph_index metadata + Y positions to identify:
    //  - Paragraph boundaries (paragraph_index changes)
    //  - Lines within paragraphs (Y position changes within same paragraph_index)
    //  - Table regions (elements without paragraph_index, bounded by TableBorder elements)
    for (pi, page) in result.pages.iter().enumerate() {
        writeln!(out, "PAGE\t{}\t{}\t{}", pi, page.width, page.height).unwrap();

        // Group text elements by paragraph_index
        let mut current_para: Option<usize> = None;
        let mut current_line_y: Option<f32> = None;
        let mut line_chars: usize = 0;
        let mut para_y: f32 = 0.0;
        let mut para_line_count: usize = 0;
        let mut in_table = false;
        let mut table_rows: Vec<f32> = Vec::new(); // unique Y positions of table row borders

        for elem in &page.elements {
            match &elem.content {
                layout::LayoutContent::Text { ref text, .. } => {
                    let para_idx = elem.paragraph_index;
                    let ey = (elem.y * 2.0).round() / 2.0; // round to 0.5pt

                    // Detect paragraph boundary
                    if para_idx != current_para {
                        // Flush previous paragraph
                        if current_para.is_some() && line_chars > 0 {
                            writeln!(out, "  LINE\ty={:.2}\tchars={}", current_line_y.unwrap_or(0.0), line_chars).unwrap();
                        }
                        // End table if a body paragraph appears after table content
                        if in_table && para_idx.is_some() && elem.paragraph_index != current_para {
                            // Emit collected table rows
                            if table_rows.len() >= 2 {
                                table_rows.sort_by(|a, b| a.partial_cmp(b).unwrap());
                                table_rows.dedup();
                                writeln!(out, "TABLE_ROWS\t{}", table_rows.len() - 1).unwrap();
                                for i in 0..table_rows.len() - 1 {
                                    writeln!(out, "  ROW\t{}\ty={:.2}\th={:.2}", i, table_rows[i], table_rows[i + 1] - table_rows[i]).unwrap();
                                }
                            }
                            in_table = false;
                            table_rows.clear();
                        }
                        // Start new paragraph
                        if let Some(idx) = para_idx {
                            if !in_table {
                                writeln!(out, "PARA\t{}\ty={:.2}", idx, elem.y).unwrap();
                            }
                        }
                        current_para = para_idx;
                        current_line_y = Some(ey);
                        line_chars = 0;
                        para_y = elem.y;
                        para_line_count = 0;
                    }

                    // Detect line boundary within paragraph
                    // Use threshold > text_y_offset variation (different fonts on same line
                    // can have ~2pt Y difference), but < smallest line height (~8pt).
                    if let Some(ly) = current_line_y {
                        if (ey - ly).abs() > 4.0 {
                            writeln!(out, "  LINE\ty={:.2}\tchars={}", ly, line_chars).unwrap();
                            line_chars = 0;
                            para_line_count += 1;
                            current_line_y = Some(ey);
                        }
                    } else {
                        current_line_y = Some(ey);
                    }

                    line_chars += text.chars().count();
                }

                layout::LayoutContent::TableBorder { y1, y2, x1, x2, .. } => {
                    // Horizontal border = table row boundary
                    if (y1 - y2).abs() < 0.1 {
                        let by = (*y1 * 2.0).round() / 2.0;
                        if !table_rows.contains(&by) {
                            table_rows.push(by);
                        }
                    }
                    if !in_table {
                        // Flush any pending paragraph
                        if current_para.is_some() && line_chars > 0 {
                            writeln!(out, "  LINE\ty={:.2}\tchars={}", current_line_y.unwrap_or(0.0), line_chars).unwrap();
                            line_chars = 0;
                        }
                        in_table = true;
                        table_rows.clear();
                        let by = (*y1 * 2.0).round() / 2.0;
                        table_rows.push(by);
                        writeln!(out, "TABLE_START\ty={:.2}", y1).unwrap();
                        current_para = None;
                    }
                }

                _ => {
                    // Non-text, non-border: if we were in a table and now see
                    // a paragraph_index text element, the table ended.
                }
            }
        }

        // Flush last paragraph
        if line_chars > 0 {
            writeln!(out, "  LINE\ty={:.2}\tchars={}", current_line_y.unwrap_or(0.0), line_chars).unwrap();
        }

        // Output table rows from collected border Ys
        if in_table && table_rows.len() >= 2 {
            table_rows.sort_by(|a, b| a.partial_cmp(b).unwrap());
            table_rows.dedup();
            writeln!(out, "TABLE_ROWS\t{}", table_rows.len() - 1).unwrap();
            for i in 0..table_rows.len() - 1 {
                let row_y = table_rows[i];
                let row_h = table_rows[i + 1] - table_rows[i];
                writeln!(out, "  ROW\t{}\ty={:.2}\th={:.2}", i, row_y, row_h).unwrap();
            }
        }
    }
}
