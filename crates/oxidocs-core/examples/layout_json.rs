use oxidocs_core::{parse_docx, layout};
use std::io::Write;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: layout_json <docx_file>");
        std::process::exit(1);
    }
    let data = std::fs::read(&args[1]).expect("failed to read file");
    let doc = parse_docx(&data).expect("parse failed");
    let engine = layout::LayoutEngine::for_document(&doc);
    let result = engine.layout(&doc);

    let stdout = std::io::stdout();
    let mut out = stdout.lock();

    // Output as JSON-lines format for easy Python parsing
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
                    // Text content on next line (may contain tabs)
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
