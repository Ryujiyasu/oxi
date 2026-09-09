// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Opens one presentation and says what was in it, as a line of JSON.
//!
//! The companion of `oxicells-core`'s `_corpus_open`: for sweeping a set of
//! decks nothing has been fitted to, and for checking that a deck this repo
//! generates is one the engine can actually read.

fn main() {
    let Some(path) = std::env::args().nth(1) else {
        eprintln!("give it a presentation");
        std::process::exit(2);
    };
    let data = match std::fs::read(&path) {
        Ok(data) => data,
        Err(error) => {
            eprintln!("{error}");
            std::process::exit(1);
        }
    };
    let deck = match oxislides_core::parser::parse_pptx(&data) {
        Ok(deck) => deck,
        Err(error) => {
            eprintln!("{error}");
            std::process::exit(1);
        }
    };

    let shapes: usize = deck.slides.iter().map(|slide| slide.shapes.len()).sum();
    let name = std::path::Path::new(&path)
        .file_name()
        .map(|name| name.to_string_lossy().replace('"', "'"))
        .unwrap_or_default();
    println!(
        "{{\"file\":\"{name}\",\"slides\":{},\"shapes\":{shapes},\
         \"width\":{},\"height\":{}}}",
        deck.slides.len(),
        deck.slide_width,
        deck.slide_height,
    );
}
