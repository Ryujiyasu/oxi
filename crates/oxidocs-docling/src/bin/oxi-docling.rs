// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! oxi-docling <input.docx> [output.json]
//! Writes a DoclingDocument JSON with page/bbox provenance (stdout when no output path).
fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let Some(input) = args.first() else {
        eprintln!("usage: oxi-docling <input.docx> [output.json]");
        std::process::exit(2);
    };
    let data = std::fs::read(input).unwrap_or_else(|e| { eprintln!("{input}: {e}"); std::process::exit(1) });
    let (doc, s) = oxidocs_docling::docx_to_docling(&data, input).unwrap_or_else(|e| { eprintln!("{e}"); std::process::exit(1) });
    let out = serde_json::to_string_pretty(&doc).unwrap();
    match args.get(1) {
        Some(p) => std::fs::write(p, out).unwrap(),
        None => println!("{out}"),
    }
    eprintln!("{input}: {} pages, {} items, {} with provenance, {} unplaced fragments",
        s.pages, s.items, s.items_with_prov, s.unplaced_fragments);
}
