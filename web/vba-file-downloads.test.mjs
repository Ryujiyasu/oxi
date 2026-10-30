// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {bindVbaFileDownloads} from './vba-file-downloads.js';

const element = () => ({children: [], append(child) { this.children.push(child); }, replaceChildren(...children) { this.children = children; }});
const previousDocument = globalThis.document;
globalThis.document = {createElement: element};
const container = element(), downloads = bindVbaFileDownloads(container);
try {
    const bytes = new Uint8Array(await readFile(new URL('../tests/fixtures/vba/row_height_state.xlsm', import.meta.url)));
    const path = 'C:\\copies\\コピー.xlsm';
    downloads.show({files: [{path, bytes: Array.from(bytes)}]});
    assert.equal(container.children.length, 1);
    const first = container.children[0].children[0];
    assert.equal(first.download, 'コピー.xlsm');
    assert.equal(first.textContent, first.download);
    assert.equal(first.title, path);
    assert.deepEqual(new Uint8Array(await (await fetch(first.href)).arrayBuffer()), bytes);
    downloads.show({result: {type: 'long', value: bytes.length}});
    assert.equal(container.children[0].children[0].href, first.href);
    const replacement = Uint8Array.of(80, 75, 0, 255);
    downloads.show({files: [{path: 'c:/COPIES/コピー.xlsm', bytes: replacement}]});
    assert.equal(container.children.length, 1);
    await assert.rejects(fetch(first.href));
    const second = container.children[0].children[0];
    assert.deepEqual(new Uint8Array(await (await fetch(second.href)).arrayBuffer()), replacement);
    downloads.show({deleted_files: [path]});
    assert.equal(container.children.length, 0);
    await assert.rejects(fetch(second.href));
    downloads.show({files: [{path, bytes}, {path: 'another.xlsm', bytes: replacement}]});
    const urls = container.children.map(child => child.children[0].href);
    downloads.clear();
    assert.equal(container.children.length, 0);
    for (const url of urls) await assert.rejects(fetch(url));
    console.log(JSON.stringify({passed: true, binary_download: true, retained_after_read: true, overwrite: true, deletion: true, revoked_urls: true}));
} finally {
    downloads.clear();
    globalThis.document = previousDocument;
}
