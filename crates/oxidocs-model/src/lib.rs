// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A WordprocessingML document model that loses nothing it accepts.
//!
//! [`DocxModel`] is a whole document as plain data: blocks, paragraphs, runs,
//! tables, sections, headers, footers and notes, with every property bag
//! checked against the ECMA-376 schema and written in schema order. Anything
//! the model cannot express is refused with the path of the feature
//! ([`Unsupported`]), never passed through silently.
//!
//! The model is what the editor's creation API speaks: a caller can build a
//! document as JSON, and [`DocxModel::write`] turns it into a .docx, on its
//! own or over a template that supplies the document-wide parts.

pub mod canon;
pub mod content;
pub mod package;
pub mod props;
pub mod schema;
pub mod xml;

pub use content::{Block, Cell, CellItem, Inline, Paragraph, Row, RowItem, Run, RunItem, Table};
pub use package::{DocxModel, Error, Package};
pub use props::{Props, Unsupported};

/// Read a document into the model, pass it through JSON (so only plain data
/// survives), and write it again. With `template`, the original's
/// document-wide parts are reused; without, the model must say everything.
pub fn rebuild(bytes: &[u8], template: bool) -> Result<Vec<u8>, Error> {
    let pkg = Package::read(bytes)?;
    let model = DocxModel::from_package(&pkg)?;
    let json = serde_json::to_string(&model).map_err(|e| Error::Xml(e.to_string()))?;
    let model: DocxModel = serde_json::from_str(&json).map_err(|e| Error::Xml(e.to_string()))?;
    model.write(template.then_some(&pkg))
}
