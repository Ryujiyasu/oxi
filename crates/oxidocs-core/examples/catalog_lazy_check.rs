// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// Lay a document out and print the full layout (Debug form) to stdout.
//
// Built with `--features external-font-catalog`, it does what the browser does:
// lay out, take the catalog members the layout asked for and did not have,
// register them from font_catalog_metrics.gz, and lay out again until nothing
// is missing. Built without the feature, it is the bundled reference. The two
// outputs must be byte-identical; stderr reports the rounds and members used.
//
//   cargo run --example catalog_lazy_check [--features external-font-catalog] -- <docx> [catalog.gz]

use oxidocs_core::{font, layout, parse_docx};

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let data = std::fs::read(&args[1]).expect("failed to read the document");
    let catalog_path = args.get(2).cloned().unwrap_or_else(|| {
        concat!(env!("CARGO_MANIFEST_DIR"), "/src/font/data/font_catalog_metrics.gz").to_string()
    });
    let doc = parse_docx(&data).expect("parse failed");
    let mut rounds = 0;
    let mut loaded = 0;
    let result = loop {
        rounds += 1;
        let engine = layout::LayoutEngine::for_document(&doc);
        let result = engine.layout(&doc);
        let missing = font::take_missing_catalog_members();
        if missing.is_empty() || rounds > 8 {
            break result;
        }
        let blob = std::fs::read(&catalog_path).expect("failed to read the catalog");
        for m in missing {
            let bytes = blob[m.offset..m.offset + m.length].to_vec();
            assert!(font::register_catalog_member(&m.key, bytes), "rejected member {}", m.key);
            loaded += 1;
        }
    };
    eprintln!("rounds {rounds} members_loaded {loaded} pages {}", result.pages.len());
    // Every element's box, with the f32s printed in full, and its text and face:
    // any metric that differs between the two builds moves one of these.
    for (pi, page) in result.pages.iter().enumerate() {
        println!("PAGE {pi} {:?} {:?} {}", page.width, page.height, page.elements.len());
        for e in &page.elements {
            let (text, face) = match &e.content {
                layout::LayoutContent::Text { text, font_size, font_family, bold, italic, .. } =>
                    (text.as_str(), format!("{font_family:?} {font_size:?} {bold} {italic}")),
                _ => ("", String::from("-")),
            };
            println!("{:?} {:?} {:?} {:?} {:?} {:?} {face} {text}", e.x, e.y, e.width, e.height, e.text_y_off, e.baseline_offset);
        }
    }
}
