// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Apply position edits to a .docx, or list its paragraphs as positions
//! count them.
//!
//!     cargo run -p oxidocs-model --release --example edit -- in.docx edits.json out.docx
//!     cargo run -p oxidocs-model --release --example edit -- --text in.docx [story.json]
//!     cargo run -p oxidocs-model --release --example edit -- --word in.docx out.docx COMMAND ARGS...
//!
//! `--word` runs one of Word's commands as edits on body positions:
//! `table PARA OFFSET ROWS COLS`, `link PARA FROM TO URL`,
//! `footnote PARA OFFSET TEXT`, `comment PARA FROM TO TEXT`.
//!
//! `edits.json` is an array of edits (see `oxidocs_model::edit::Edit`). With
//! `--text`, prints a JSON array of the story's paragraph texts, one
//! character per Word character position.

use oxidocs_model::edit::{self, Edit, Pos, Story};
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
    if args.first().map(|s| s.as_str()) == Some("--word") {
        let mut model = DocxModel::read(&std::fs::read(&args[1]).expect("read docx")).expect("model");
        let n = |i: usize| -> usize { args[i].parse().expect("number") };
        let edits = match args[3].as_str() {
            "table" => edit::table_edits(&model, &Pos::body(n(4), n(5)), n(6), n(7)).map(|r| r.0),
            "link" => edit::link_edits(&model, &Pos::body(n(4), n(5)), &Pos::body(n(4), n(6)), &args[7]).map(|r| r.0),
            "footnote" => edit::footnote_edits(&model, &Pos::body(n(4), n(5)), &args[6]).map(|r| r.0),
            "comment" => edit::comment_edits(&model, &Pos::body(n(4), n(5)), &Pos::body(n(4), n(6)), &args[7], "Oxi", "O", "2026-11-09T00:00:00Z").map(|r| r.0),
            other => panic!("no command {other}"),
        };
        let edits = edits.unwrap_or_else(|e| {
            eprintln!("{e}");
            std::process::exit(1)
        });
        edit::apply_all(&mut model, &edits).expect("apply");
        std::fs::write(&args[2], model.write(None).expect("write")).expect("write docx");
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
