// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// An XLL finds Excel's callback by asking the process's own executable for
// `MdCallBack12` (the Excel SDK's xlcall.cpp does
// GetProcAddress(GetModuleHandle(NULL), "MdCallBack12")), so the host binary
// has to export it.
fn main() {
    if std::env::var("CARGO_CFG_TARGET_OS").as_deref() == Ok("windows") {
        println!("cargo:rustc-link-arg-bins=/EXPORT:MdCallBack12");
    }
}
