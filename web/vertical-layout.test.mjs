// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

// Run after a WASM build: node web/vertical-layout.test.mjs <wasm package directory>
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const pkg = path.resolve(process.argv[2] || path.join(root, 'crates/oxidocs-wasm/pkg'));
const source = await readFile(path.join(pkg, 'oxidocs_wasm.js'), 'utf8');
const wasm = await import('data:text/javascript;base64,' + Buffer.from(source).toString('base64'));
await wasm.default({ module_or_path: await readFile(path.join(pkg, 'oxidocs_wasm_bg.wasm')) });

for (const [fixture, vertical] of [['vertical_text.docx', true], ['basic_test.docx', false]]) {
  const initial = wasm.layout_document(await readFile(path.join(root, 'tests/fixtures', fixture)));
  const initialText = initial.pages.flatMap(p => p.elements).filter(e => e.kind === 'text');
  assert.ok(initialText.length > 0, fixture);
  assert.ok(initialText.every(e => e.is_vertical === vertical), `${fixture}: initial orientation`);
  if (vertical) assert.equal(initialText.map(e => e.text).join(''), '日本語の縦書き');
  const updated = wasm.edit_text_and_relayout(0, 0, '日本語の修正');
  const updatedText = updated.pages.flatMap(p => p.elements).filter(e => e.kind === 'text');
  assert.ok(updatedText.length > 0, fixture);
  assert.ok(updatedText.every(e => e.is_vertical === vertical), `${fixture}: edited orientation`);
  assert.ok(updatedText.map(e => e.text).join('').includes('日本語の修正'), `${fixture}: edited text`);
  assert.equal(updated.pages.length, initial.pages.length, `${fixture}: pagination`);
  console.log(`${fixture}: initial and edited orientation OK (${initial.pages.length} page)`);
}
