// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use oxidocs_core::layout::LayoutContent;
use oxidocs_core::parse_docx;

fn main() {
    let path = std::env::args().nth(1).expect("usage: dump_underline <file.docx>");
    let max_page: usize = std::env::args()
        .nth(2)
        .and_then(|s| s.parse().ok())
        .unwrap_or(1);

    let data = std::fs::read(&path).expect("read failed");
    let doc = parse_docx(&data).expect("parse failed");
    let engine = oxidocs_core::layout::LayoutEngine::for_document(&doc);
    let result = engine.layout(&doc);

    let mut total_text = 0usize;
    let mut total_underline = 0usize;
    let mut total_strike = 0usize;
    let mut samples: Vec<(usize, f32, f32, f32, f32, f32, String, String)> = Vec::new();

    for (pi, page) in result.pages.iter().enumerate() {
        if pi >= max_page {
            break;
        }
        for el in &page.elements {
            if let LayoutContent::Text {
                text,
                font_size,
                font_family,
                underline,
                strikethrough,
                ..
            } = &el.content
            {
                total_text += 1;
                if *underline {
                    total_underline += 1;
                    if samples.len() < 60 {
                        samples.push((
                            pi + 1,
                            el.x,
                            el.y,
                            el.width,
                            el.height,
                            *font_size,
                            font_family.clone().unwrap_or_else(|| "?".into()),
                            text.clone(),
                        ));
                    }
                }
                if *strikethrough {
                    total_strike += 1;
                }
            }
        }
    }

    println!(
        "doc={} pages={} (limited to first {})",
        path,
        result.pages.len(),
        max_page
    );
    println!(
        "totals: text={} underline={} strike={}",
        total_text, total_underline, total_strike
    );
    println!();
    println!("first {} underlined Text elements:", samples.len());
    println!("page    x       y     w      h    fs   font            text");
    for (p, x, y, w, h, fs, fam, txt) in &samples {
        let snippet: String = txt.chars().take(40).collect();
        println!(
            "  {}  {:6.2} {:6.2} {:6.2} {:5.2} {:4.1} {:14}  {:?}",
            p, x, y, w, h, fs, fam, snippet
        );
    }
}
