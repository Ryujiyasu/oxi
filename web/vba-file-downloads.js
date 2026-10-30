// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

/// Keep download links for the files a macro produced. A later invocation
/// may update or delete them, or leave them available while reading them.
export function bindVbaFileDownloads(container) {
    const downloads = new Map();
    const key = path => path.trim().replaceAll('/', '\\').toLowerCase();
    const remove = path => {
        const file = downloads.get(key(path));
        if (file) URL.revokeObjectURL(file.url);
        downloads.delete(key(path));
    };
    return {
        clear() {
            for (const file of downloads.values()) URL.revokeObjectURL(file.url);
            downloads.clear();
            container.replaceChildren();
        },
        show(execution) {
            for (const path of execution.deleted_files || []) remove(path);
            for (const file of execution.files || []) {
                remove(file.path);
                const url = URL.createObjectURL(new Blob([new Uint8Array(file.bytes)], { type: 'application/octet-stream' }));
                const link = document.createElement('a');
                link.href = url;
                link.download = file.path.split(/[\\/]/).at(-1) || 'workbook.xlsx';
                link.textContent = link.download;
                link.title = file.path;
                const item = document.createElement('div');
                item.append(link);
                downloads.set(key(file.path), { url, item });
            }
            container.replaceChildren(...Array.from(downloads.values(), file => file.item));
        },
    };
}
