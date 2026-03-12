use oxidocs_core::{parse_docx, ir, layout};

fn main() {
    let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
    let doc = parse_docx(data).expect("parse failed");

    println!("=== PARSED IR ===");
    println!("Pages: {}", doc.pages.len());
    for (pi, page) in doc.pages.iter().enumerate() {
        println!("\n--- Page {} ---", pi + 1);
        println!("Size: {:.1} x {:.1} pt", page.size.width, page.size.height);
        println!("Margins: T={:.0} R={:.0} B={:.0} L={:.0}", 
            page.margin.top, page.margin.right, page.margin.bottom, page.margin.left);
        println!("Blocks: {}", page.blocks.len());
        for (bi, block) in page.blocks.iter().enumerate() {
            match block {
                ir::Block::Paragraph(p) => {
                    let heading = p.style.heading_level
                        .map(|l| format!(" [H{}]", l))
                        .unwrap_or_default();
                    let align = format!("{:?}", p.alignment);
                    print!("  [{}] Paragraph{} align={}", bi, heading, align);
                    for run in &p.runs {
                        let mut flags = Vec::new();
                        if run.style.bold { flags.push("B"); }
                        if run.style.italic { flags.push("I"); }
                        if run.style.underline { flags.push("U"); }
                        let f = if flags.is_empty() { String::new() } else { format!(" [{}]", flags.join("")) };
                        let size = run.style.font_size.map(|s| format!(" {:.0}pt", s)).unwrap_or_default();
                        print!("\n    Run{}{}: \"{}\"", f, size, run.text);
                    }
                    println!();
                }
                ir::Block::Table(t) => {
                    println!("  [{}] Table {}x{} border={}", 
                        bi, t.rows.len(), 
                        t.rows.first().map_or(0, |r| r.cells.len()),
                        t.style.border);
                    for (ri, row) in t.rows.iter().enumerate() {
                        print!("    Row {}: ", ri);
                        for cell in &row.cells {
                            for block in &cell.blocks {
                                if let ir::Block::Paragraph(p) = block {
                                    let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
                                    print!("| {} ", text);
                                }
                            }
                        }
                        println!("|");
                    }
                }
                ir::Block::Image(img) => {
                    println!("  [{}] Image {:.0}x{:.0}pt ({} bytes) alt={:?}", 
                        bi, img.width, img.height, img.data.len(), img.alt_text);
                }
            }
        }
    }

    println!("\n=== STYLES ===");
    for (name, style) in &doc.styles.styles {
        println!("  {} -> before={:?} after={:?} line_spacing={:?}", 
            name, style.space_before, style.space_after, style.line_spacing);
    }

    println!("\n=== LAYOUT ===");
    let engine = layout::LayoutEngine::new();
    let result = engine.layout(&doc);
    for (pi, page) in result.pages.iter().enumerate() {
        println!("\n--- Layout Page {} ({:.0}x{:.0}) ---", pi + 1, page.width, page.height);
        println!("Elements: {}", page.elements.len());
        for (ei, elem) in page.elements.iter().enumerate() {
            match &elem.content {
                layout::LayoutContent::Text { text, font_size, bold, italic, .. } => {
                    let mut flags = String::new();
                    if *bold { flags.push('B'); }
                    if *italic { flags.push('I'); }
                    if !flags.is_empty() { flags = format!(" [{}]", flags); }
                    println!("  [{:2}] TEXT ({:6.1}, {:6.1}) w={:6.1} h={:5.1} {:.0}pt{} \"{}\"", 
                        ei, elem.x, elem.y, elem.width, elem.height, font_size, flags, text);
                }
                layout::LayoutContent::Image { data, .. } => {
                    println!("  [{:2}] IMG  ({:6.1}, {:6.1}) w={:6.1} h={:5.1} {} bytes", 
                        ei, elem.x, elem.y, elem.width, elem.height, data.len());
                }
                layout::LayoutContent::TableBorder { x1, y1, x2, y2 } => {
                    println!("  [{:2}] BORDER ({:.1},{:.1})->({:.1},{:.1})", ei, x1, y1, x2, y2);
                }
            }
        }
    }
}
