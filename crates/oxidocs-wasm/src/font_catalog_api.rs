// SPDX-License-Identifier: MIT OR Apache-2.0

//! The host side of the external font catalog (`external-font-catalog`).
//!
//! Built with that feature, the per-face metric catalog is not compiled into
//! the module. After a layout the host asks which catalog members it needed and
//! did not have, fetches exactly those byte ranges of `font_catalog_metrics.gz`,
//! registers them, and lays the document out again. Bundled builds always
//! report nothing missing, so the same page code runs against either build.

use serde::Serialize;
use wasm_bindgen::prelude::*;

#[derive(Serialize)]
struct MissingMember {
    key: String,
    offset: usize,
    length: usize,
}

/// The catalog members requested since the last call that were not
/// registered: `[{ key, offset, length }]`, byte ranges in
/// `font_catalog_metrics.gz`. Empty for a bundled build.
#[wasm_bindgen]
pub fn font_catalog_missing() -> Result<JsValue, JsError> {
    let missing: Vec<MissingMember> = oxidocs_core::font::take_missing_catalog_members()
        .into_iter()
        .map(|m| MissingMember { key: m.key, offset: m.offset, length: m.length })
        .collect();
    serde_wasm_bindgen::to_value(&missing).map_err(|e| JsError::new(&e.to_string()))
}

/// Register one member's bytes (exactly its range in the catalog file).
/// Returns false when the key is unknown or the length does not match.
#[wasm_bindgen]
pub fn font_catalog_register(key: &str, bytes: &[u8]) -> bool {
    oxidocs_core::font::register_catalog_member(key, bytes.to_vec())
}
