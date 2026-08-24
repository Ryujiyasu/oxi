// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Says whether the golden-corpus workbooks are present.
//!
//! `tools/golden-test/documents/` holds real published workbooks that are not
//! redistributed with the source, so the tests that `include_bytes!` them can
//! only be compiled where that corpus exists. Without this flag a clean clone
//! (and CI) fails to build the test target at all, which is what happened on
//! the first CI run after the gates were turned on.

use std::path::Path;

fn main() {
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-env-changed=OXI_NO_GOLDEN_CORPUS");
    println!("cargo:rustc-check-cfg=cfg(golden_corpus)");
    // Setting OXI_NO_GOLDEN_CORPUS reproduces a clean clone on a machine that
    // does have the corpus, so the CI arm can be checked before pushing.
    let forced_off = std::env::var_os("OXI_NO_GOLDEN_CORPUS").is_some();
    if !forced_off && Path::new("../../tools/golden-test/documents/xlsx").is_dir() {
        println!("cargo:rustc-cfg=golden_corpus");
    }
}
