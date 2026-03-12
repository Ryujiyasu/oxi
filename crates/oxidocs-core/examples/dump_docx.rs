use oxidocs_core::{parse_docx, ir, layout};

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let data: Vec<u8> = if args.len() > 1 {
        std::fs::read(&args[1]).expect("Failed to read file")
    } else {
        include_bytes!("../../../tests/fixtures/basic_test.docx").to_vec()
    };
    let doc = parse_docx(&data).expect("parse failed");

    println!("=== PARSED IR ===");
    println!("Pages: {}", doc.pages.len());
    for (pi, page) in doc.pages.iter().enumerate() {
        println!("\n--- Page {} ---", pi + 1);
        println!("Size: {:.1} x {:.1} pt", page.size.width, page.size.height);
        println!("Margins: T={:.0} R={:.0} B={:.0} L={:.0}", 
            page.margin.top, page.margin.right, page.margin.bottom, page.margin.left);
        println!("Blocks: {} | TextBoxes: {} | Shapes: {} | FloatingImages: {}",
            page.blocks.len(), page.text_boxes.len(), page.shapes.len(), page.floating_images.len());
        for (ti, tb) in page.text_boxes.iter().enumerate() {
            println!("  TextBox[{}]: {:.1}x{:.1} anchor={} fill={:?} cr={:?} blocks={}",
                ti, tb.width, tb.height, tb.anchor_block_index, tb.fill, tb.corner_radius, tb.blocks.len());
            for (tbi, tblock) in tb.blocks.iter().enumerate() {
                if let ir::Block::Paragraph(tp) = tblock {
                    let ls_info = format!("ls={:?}/{:?} snap={}", tp.style.line_spacing, tp.style.line_spacing_rule, tp.style.snap_to_grid);
                    let sb = tp.style.space_before.map(|v| format!(" sb={:.1}", v)).unwrap_or_default();
                    print!("    TB[{}].P{} {}{}", ti, tbi, ls_info, sb);
                    for run in &tp.runs {
                        let sz = run.style.font_size.map(|s| format!(" {:.0}pt", s)).unwrap_or_default();
                        let col = run.style.color.as_deref().map(|c| format!(" c={}", c)).unwrap_or_default();
                        print!(" |{}{} \"{}\"", sz, col, &run.text[..run.text.len().min(15)]);
                    }
                    println!();
                }
            }
        }
        for (bi, block) in page.blocks.iter().enumerate() {
            match block {
                ir::Block::Paragraph(p) => {
                    let heading = p.style.heading_level
                        .map(|l| format!(" [H{}]", l))
                        .unwrap_or_default();
                    let align = format!("{:?}", p.alignment);
                    let marker = p.style.list_marker.as_deref().map(|m| format!(" marker=\"{}\"", m)).unwrap_or_default();
                    let sid = p.style.style_id.as_deref().map(|s| format!(" style={}", s)).unwrap_or_default();
                    let ls_info = format!(" ls={:?}/{:?} snap={}", p.style.line_spacing, p.style.line_spacing_rule, p.style.snap_to_grid);
                    print!("  [{}] Paragraph{} align={}{}{}{}", bi, heading, align, sid, marker, ls_info);
                    for run in &p.runs {
                        let mut flags = Vec::new();
                        if run.style.bold { flags.push("B"); }
                        if run.style.italic { flags.push("I"); }
                        if run.style.underline { flags.push("U"); }
                        let f = if flags.is_empty() { String::new() } else { format!(" [{}]", flags.join("")) };
                        let size = run.style.font_size.map(|s| format!(" {:.0}pt", s)).unwrap_or_default();
                        let cs = run.style.character_spacing.map(|s| format!(" cs={:.1}", s)).unwrap_or_default();
                        let ff = run.style.font_family.as_deref().or(run.style.font_family_east_asia.as_deref()).map(|f| format!(" font={}", f)).unwrap_or_default();
                        print!("\n    Run{}{}{}{}: \"{}\"", f, size, cs, ff, run.text);
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
                                match block {
                                    ir::Block::Paragraph(p) => {
                                        let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
                                        print!("| {} ", text);
                                    }
                                    ir::Block::Table(nt) => {
                                        print!("| [NESTED TABLE {}x{} border={} style={:?}] ",
                                            nt.rows.len(),
                                            nt.rows.first().map_or(0, |r| r.cells.len()),
                                            nt.style.border,
                                            nt.style.style_id);
                                    }
                                    _ => {}
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
                ir::Block::UnsupportedElement(u) => {
                    println!("  [{}] Unsupported: {}", bi, u.element_type);
                }
            }
        }
    }

    println!("\n=== STYLES ===");
    for (name, def) in &doc.styles.styles {
        let style = &def.paragraph;
        let num = style.num_id.as_deref().map(|n| format!(" numId={} ilvl={}", n, style.num_ilvl)).unwrap_or_default();
        println!("  {} -> before={:?} after={:?} line_spacing={:?} basedOn={:?}{}",
            name, style.space_before, style.space_after, style.line_spacing, def.based_on, num);
    }

    println!("\n=== LAYOUT ===");
    let engine = layout::LayoutEngine::for_document(&doc);
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
                layout::LayoutContent::TableBorder { x1, y1, x2, y2, ref color, width } => {
                    let col = color.as_deref().unwrap_or("#000");
                    println!("  [{:2}] BORDER ({:.1},{:.1})->({:.1},{:.1}) color={} w={:.1}", ei, x1, y1, x2, y2, col, width);
                }
                layout::LayoutContent::CellShading { ref color } => {
                    println!("  [{:2}] SHADING ({:.1},{:.1}) w={:.1} h={:.1} color={}", ei, elem.x, elem.y, elem.width, elem.height, color);
                }
                layout::LayoutContent::BoxRect { ref fill, ref stroke_color, corner_radius, .. } => {
                    println!("  [{:2}] BOX ({:.1},{:.1}) w={:.1} h={:.1} fill={:?} stroke={:?} cr={:.1}",
                        ei, elem.x, elem.y, elem.width, elem.height, fill, stroke_color, corner_radius);
                }
                layout::LayoutContent::ClipStart => {
                    println!("  [{:2}] CLIP_START ({:.1},{:.1}) w={:.1} h={:.1}", ei, elem.x, elem.y, elem.width, elem.height);
                }
                layout::LayoutContent::ClipEnd => {
                    println!("  [{:2}] CLIP_END", ei);
                }
            }
        }
    }
}
