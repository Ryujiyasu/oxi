// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Where the caret can stand in the docx editor around list numbers.
//!
//! A list number ("3.") is laid out as text with no paragraph behind it. The
//! editor used to let a click or an arrow put the caret on it: the caret sat
//! after the number as if it could be edited, but it reached no position in
//! the document, so Backspace and typing did nothing. An empty list item was
//! the worst case, since the number was the only text on its line.
//!
//! This runs the page's own hit testing and caret movement against the real
//! engine's layout of tests/fixtures/numbered_list_empty_item.docx
//! (1. 2. 3.<empty> 4. 5.) and the real model behind it.
//!
//! Run with `node --test web/docs-caret.test.mjs [path/to/pkg]` (default: web/).

import assert from 'node:assert/strict';
import test from 'node:test';
import { readFile } from 'node:fs/promises';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { dirname, join, resolve } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const pkg = resolve(process.argv[2] || here);
const wasm = await import(pathToFileURL(join(pkg, 'oxidocs_wasm.js')).href);
await wasm.default({ module_or_path: await readFile(join(pkg, 'oxidocs_wasm_bg.wasm')) });
const { layout_document, DocxDocument } = wasm;

const html = await readFile(join(here, 'docs.html'), 'utf8');
function fn(name) {
  const start = html.indexOf(`function ${name}(`);
  assert.ok(start >= 0, `docs.html has no ${name}()`);
  return html.slice(start, html.indexOf('\n}', start) + 2);
}

// The page's functions over a stand-in canvas: 5pt per character.
function editor(layoutResult, modelDoc) {
  const env = { layoutResult, modelDoc, cursorPos: null };
  const document = { createElement: () => ({ getContext: () => ({ measureText: t => ({ width: t.length * 5 }) }) }) };
  const body = ['caretStop', 'hitTestText', 'getCharOffset', 'moveCursor', 'elCells', 'modelPosAt']
    .map(fn).join('\n') + `
    return {
      hit: (x, y) => hitTestText({ clientX: x, clientY: y, target: canvas }, 0),
      move(pos, dir) { cursorPos = { ...pos }; moveCursor(dir); return cursorPos; },
      modelPos: pos => modelPosAt(layoutResult.pages[pos.pageIndex].elements[pos.elemIndex], pos.charOffset),
    };`;
  const canvas = { width: 1000, height: 1000, getBoundingClientRect: () => ({ left: 0, top: 0, width: 1000, height: 1000 }) };
  return new Function('layoutResult', 'modelDoc', 'document', 'canvas', 'currentScale', 'startCursorBlink',
    `let cursorPos = null;\n${body}`)(layoutResult, modelDoc, document, canvas, 1, () => {});
}

const bytes = new Uint8Array(await readFile(join(here, '../tests/fixtures/numbered_list_empty_item.docx')));
const els = () => layout_document(bytes).pages[0].elements;
const marker = (list, label) => list.find(e => e.kind === 'text' && e.text === label && e.paragraph_index == null);
const mid = e => e.y + (e.height || e.font_size || 11) / 2;

test('the numbers are laid out without a paragraph', () => {
  const list = els();
  for (const n of ['1.', '2.', '3.', '4.', '5.']) assert.ok(marker(list, n), `no ${n} marker`);
});

test('a click on an empty item\'s number puts the caret in the empty paragraph', () => {
  const layout = layout_document(bytes);
  const ed = editor(layout, new DocxDocument(bytes));
  const three = marker(layout.pages[0].elements, '3.');
  const hit = ed.hit(three.x + three.width, mid(three));
  const el = layout.pages[0].elements[hit.elemIndex];
  assert.notEqual(el.paragraph_index, null, 'the caret is on the number');
  assert.equal(el.text, '');
  assert.ok(el.x > three.x + three.width, 'the caret sits at the text indent, past the number');
  assert.ok(ed.modelPos(hit), 'the caret has a document position');
});

test('Backspace there takes the number off', () => {
  const layout = layout_document(bytes);
  const model = new DocxDocument(bytes);
  const ed = editor(layout, model);
  const three = marker(layout.pages[0].elements, '3.');
  const at = ed.modelPos(ed.hit(three.x + 2, mid(three)));
  assert.equal(at.offset, 0);
  const edits = JSON.parse(model.unnumber_edits(JSON.stringify(at)));
  assert.ok(edits.length, 'nothing to take off');
  model.apply(JSON.stringify(edits));
  const after = layout_document(new Uint8Array(model.save())).pages[0].elements;
  assert.ok(!marker(after, '5.'), 'the list still has five numbers');
  assert.ok(marker(after, '4.'));
});

test('a click on a number puts the caret before its item\'s text', () => {
  const layout = layout_document(bytes);
  const ed = editor(layout, null);
  const four = marker(layout.pages[0].elements, '4.');
  const hit = ed.hit(four.x + four.width, mid(four));
  const el = layout.pages[0].elements[hit.elemIndex];
  assert.equal(el.text, 'Time');
  assert.equal(hit.charOffset, 0);
});

test('arrows step over numbers and stop in an empty item', () => {
  const layout = layout_document(bytes);
  const list = layout.pages[0].elements;
  const ed = editor(layout, null);
  const time = list.findIndex(e => e.text === 'Time');
  const left = ed.move({ pageIndex: 0, elemIndex: time, charOffset: 0 }, -1);
  assert.equal(list[left.elemIndex].paragraph_index, list[time].paragraph_index - 1);
  assert.equal(list[left.elemIndex].text, '');
  const right = ed.move(left, 1);
  assert.equal(right.elemIndex, time);
  const back = ed.move(left, -1);
  assert.equal(list[back.elemIndex].text, 'stand.');
  assert.equal(back.charOffset, 'stand.'.length);
});
