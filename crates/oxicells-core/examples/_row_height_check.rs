// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! What the measured table answers for a face and a size.
//!
//! The browser grid now asks the engine this question instead of declining to
//! answer it, and `web/row-model.test.mjs` gates the rules that USE the answer
//! against a handful of values read off Excel. This prints the engine's answer
//! for the same handful, so the two tables can be held against each other
//! rather than assumed to agree.
//!
//!     cargo run --release -p oxicells-core --example _row_height_check

fn main() {
    let asking: Vec<(String, f32)> = if std::env::args().count() > 2 {
        let mut args = std::env::args().skip(1);
        let face = args.next().unwrap_or_default();
        let size = args.next().and_then(|s| s.parse().ok()).unwrap_or(11.0);
        vec![(face, size)]
    } else {
        vec![
            ("ＭＳ Ｐゴシック".to_string(), 11.0),
            ("ＭＳ Ｐゴシック".to_string(), 18.0),
            ("游ゴシック".to_string(), 11.0),
            ("游ゴシック".to_string(), 8.0),
            ("Calibri".to_string(), 11.0),
            ("Aptos Narrow".to_string(), 11.0),
        ]
    };
    for (face, size) in asking {
        match oxicells_core::row_defaults::font_default_row_px(&face, size) {
            Some(px) => println!("{face} {size}pt -> {px}px"),
            None => println!("{face} {size}pt -> never measured"),
        }
    }
}
