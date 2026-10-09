// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::sync::OnceLock;
use oxidocs_font_provider::{InstalledFontPrograms, windows_font_roots};

/// Read only real font programs from the host font sources. Web content supplies
/// a family/style, never a path. Data is returned directly through binary IPC.
#[tauri::command]
pub async fn load_font_program(family: String, bold: bool, italic: bool) -> Result<tauri::ipc::Response, String> {
    let bytes = tauri::async_runtime::spawn_blocking(move || {
        static PROGRAMS: OnceLock<InstalledFontPrograms> = OnceLock::new();
        let registry = PROGRAMS.get_or_init(|| InstalledFontPrograms::from_roots(windows_font_roots()));
        registry.resolve(&family, bold, italic).map(|font| font.into_bytes())
            .map_err(|error| format!("Font unavailable: {family} ({error:?})"))
    }).await.map_err(|error| error.to_string())??;
    Ok(tauri::ipc::Response::new(bytes))
}
