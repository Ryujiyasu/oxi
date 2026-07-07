// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Oxidocs — a Word-compatible `.docx` layout and rendering engine.
//!
//! Oxidocs parses OOXML WordprocessingML documents into a language-agnostic
//! intermediate representation, lays them out with a line/page model
//! reverse-engineered against Microsoft Word's rendering (Japanese
//! typography — kinsoku, docGrid, ruby — included), and renders the result.
//!
//! Layout fidelity is measured continuously against Word itself: per-page
//! pixel comparison (SSIM) and per-paragraph pagination equality over a
//! corpus of real-world Japanese and English documents.
//!
//! This crate is the top-level entry point; the engine lives in
//! [`oxidocs_core`], re-exported here as [`core`]. The API is early and
//! unstable (0.x) — development happens at
//! <https://gitlab.com/Ryujiyasu/oxi>.
//!
//! ```no_run
//! let bytes = std::fs::read("example.docx").unwrap();
//! let doc = oxidocs::open(&bytes).unwrap();
//! println!("pages: {}", doc.pages.len());
//! ```

pub use oxidocs_core as core;

pub use oxidocs_core::ir::Document;
pub use oxidocs_core::parser::ParseError;

/// Parse a `.docx` file (as bytes) into the document IR.
///
/// Convenience alias for [`oxidocs_core::parser::parse_docx`].
pub fn open(data: &[u8]) -> Result<Document, ParseError> {
    oxidocs_core::parser::parse_docx(data)
}

#[cfg(test)]
mod tests {
    #[test]
    fn open_blank_docx_roundtrip() {
        let bytes = crate::core::create_blank_docx();
        let doc = crate::open(&bytes).expect("blank docx parses");
        assert!(!doc.pages.is_empty());
    }
}
