// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use tauri_plugin_dialog::{DialogExt, MessageDialogButtons};
use tauri_plugin_updater::UpdaterExt;

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_dialog::init())
        .setup(|app| {
            let handle = app.handle().clone();
            tauri::async_runtime::spawn(async move {
                check_for_updates(handle).await;
            });
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("error while running Oxi Desktop");
}

/// Startup update check: signed manifest from GitHub Releases (see
/// `plugins.updater` in tauri.conf.json). Silent when up to date or offline;
/// asks before downloading and before restarting.
async fn check_for_updates(app: tauri::AppHandle) {
    let Ok(updater) = app.updater() else { return };
    let Ok(Some(update)) = updater.check().await else { return };
    let version = update.version.clone();

    let install = app
        .dialog()
        .message(format!(
            "新しいバージョン {version} が利用可能です。今すぐ更新しますか？\n\
             A new version ({version}) is available. Update now?"
        ))
        .title("Oxi")
        .buttons(MessageDialogButtons::OkCancelCustom(
            "Update / 更新".into(),
            "Later / 後で".into(),
        ))
        .blocking_show();
    if !install {
        return;
    }

    match update.download_and_install(|_, _| {}, || {}).await {
        Ok(()) => {
            let restart = app
                .dialog()
                .message(
                    "更新をインストールしました。再起動して適用しますか？\n\
                     The update is installed. Restart Oxi to apply it?",
                )
                .title("Oxi")
                .buttons(MessageDialogButtons::OkCancelCustom(
                    "Restart / 再起動".into(),
                    "Later / 後で".into(),
                ))
                .blocking_show();
            if restart {
                app.restart();
            }
        }
        Err(err) => {
            app.dialog()
                .message(format!(
                    "更新に失敗しました。次回起動時に再試行します。\n\
                     The update failed and will be retried on next launch.\n\n{err}"
                ))
                .title("Oxi")
                .blocking_show();
        }
    }
}
