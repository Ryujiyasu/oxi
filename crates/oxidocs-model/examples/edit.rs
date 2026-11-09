// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Apply position edits to a .docx, or list its paragraphs as positions
//! count them.
//!
//!     cargo run -p oxidocs-model --release --example edit -- in.docx edits.json out.docx
//!     cargo run -p oxidocs-model --release --example edit -- --text in.docx [story.json]
//!
//! `edits.json` is an array of edits (see `oxidocs_model::edit::Edit`). With
//! `--text`, prints a JSON array of the story's paragraph texts, one
//! character per Word character position.

use oxidocs_model::edit::{self, Edit, Story};
use oxidocs_model::DocxModel;

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.first().map(|s| s.as_str()) == Some("--text") {
        let model = DocxModel::read(&std::fs::read(&args[1]).expect("read docx")).expect("model");
        let story: Story = args.get(2).map(|s| serde_json::from_str(s).expect("story")).unwrap_or_default();
        let texts = edit::paragraph_texts(&model, &story).expect("story");
        println!("{}", serde_json::to_string(&texts).unwrap());
        return;
    }
    let [input, edits, output] = &args[..] else {
        eprintln!("usage: edit in.docx edits.json out.docx | edit --text in.docx [story]");
        std::process::exit(2);
    };
    let mut model = DocxModel::read(&std::fs::read(input).expect("read docx")).expect("model");
    let edits: Vec<Edit> = serde_json::from_str(&std::fs::read_to_string(edits).expect("read edits")).expect("edits");
    if let Err(e) = edit::apply_all(&mut model, &edits) {
        eprintln!("{e}");
        std::process::exit(1);
    }
    std::fs::write(output, model.write(None).expect("write")).expect("write docx");
}
