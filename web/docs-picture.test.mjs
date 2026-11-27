// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Insert › Picture and the comment pane in the docx editor, through the
//! engine the page uses.
//!
//! - A picture goes in at the caret as Word's inline picture, laid out beside
//!   the text before it, and what is typed after it stays visible. Before,
//!   the picture and the text around it shared one run and the layout drew
//!   the picture only.
//! - A comment made in the editor has the paragraph id Word threads replies
//!   and resolved state by, so the comment pane can reply to it and resolve
//!   it.
//!
//! Run with `node web/docs-picture.test.mjs [path/to/pkg]` (default: web/).

import assert from 'node:assert/strict';
import test from 'node:test';
import { readFile } from 'node:fs/promises';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { dirname, join, resolve } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const pkg = resolve(process.argv[2] || here);
const wasm = await import(pathToFileURL(join(pkg, 'oxidocs_wasm.js')).href);
await wasm.default({ module_or_path: await readFile(join(pkg, 'oxidocs_wasm_bg.wasm')) });
const { layout_document, DocxDocument, create_blank_docx, parse_document, update_docx_comments } = wasm;

// A 2x1 PNG.
const PNG = Uint8Array.from(atob(
  'iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAD0lEQVR4nGM8ISfHwMAAAAXMAQZFiFnUAAAAAElFTkSuQmCC'), c => c.charCodeAt(0));

function blank(text) {
  const doc = new DocxDocument(new Uint8Array(create_blank_docx()));
  doc.apply(JSON.stringify([{ op: 'insert_text', at: { para: 0, offset: 0 }, text }]));
  return doc;
}
const shown = bytes => layout_document(bytes).pages[0].elements
  .filter(e => e.kind === 'image' || (e.kind === 'text' && e.text))
  .map(e => e.kind === 'image' ? '[pic]' : e.text).join('');

test('a picture goes in at the caret, beside the text before it', () => {
  const doc = blank('Logo: ');
  const end = JSON.parse(doc.insert_picture(JSON.stringify({ para: 0, offset: 6 }), PNG, 'image/png', 144 * 12700, 72 * 12700));
  assert.equal(end.offset, 7, 'the picture is one position');
  const layout = layout_document(new Uint8Array(doc.save()));
  const pic = layout.pages[0].elements.find(e => e.kind === 'image');
  assert.ok(pic, 'no picture on the page');
  assert.equal(Math.round(pic.width), 144);
  assert.equal(Math.round(pic.height), 72);
  assert.equal(shown(new Uint8Array(doc.save())).replace(/ /g, ''), 'Logo:[pic]');
});

test('text typed after a picture stays on the page', () => {
  const doc = blank('Logo: ');
  const end = JSON.parse(doc.insert_picture(JSON.stringify({ para: 0, offset: 6 }), PNG, 'image/png', 72 * 12700, 36 * 12700));
  doc.apply(JSON.stringify([{ op: 'insert_text', at: end, text: ' after' }]));
  assert.equal(shown(new Uint8Array(doc.save())).replace(/ /g, ''), 'Logo:[pic]after');
});

test('only pictures are taken', () => {
  const doc = blank('x');
  assert.throws(() => doc.insert_picture(JSON.stringify({ para: 0, offset: 0 }), PNG, 'image/webp', 100, 100));
});

test('a comment made in the editor can be replied to and resolved', () => {
  const doc = blank('Please check this.');
  const plan = JSON.parse(doc.comment_edits(JSON.stringify({ para: 0, offset: 7 }), JSON.stringify({ para: 0, offset: 12 }),
    'Is it?', 'Ann', 'A', '2026-11-27T00:00:00Z'));
  doc.apply(JSON.stringify(plan.edits));
  let bytes = new Uint8Array(doc.save());
  const [c] = parse_document(bytes).comments;
  assert.ok(c.para_id, 'the comment has no paragraph id to thread by');
  bytes = new Uint8Array(update_docx_comments(bytes, {
    add: [{ author: 'Bob', text: 'Yes.', date: '2026-11-27T00:01:00Z', paragraph_index: 0, char_start: 7, char_end: 12, parent_para_id: c.para_id }],
    set_resolved: [{ para_id: c.para_id, done: true }],
  }));
  const after = parse_document(bytes).comments;
  assert.equal(after.length, 2);
  const reply = after.find(x => x.author === 'Bob');
  assert.equal(reply.parent_para_id, c.para_id, 'the reply is not threaded under the comment');
  assert.ok(after.find(x => x.author === 'Ann').resolved, 'the comment is not resolved');
});
