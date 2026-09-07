// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

import assert from 'node:assert/strict';
import test from 'node:test';
import { readFile } from 'node:fs/promises';
import { renderVerticalText } from '../docs/vertical-text.js';

function canvas() {
  const calls = [];
  const stack = [];
  const ctx = { calls, letterSpacing: '3px', textBaseline: 'top', textAlign: 'center' };
  ctx.save = () => stack.push({ ...ctx });
  ctx.restore = () => {
    const saved = stack.pop();
    for (const key of Object.keys(ctx)) delete ctx[key];
    Object.assign(ctx, saved);
  };
  for (const method of ['fillText', 'fillRect', 'beginPath', 'moveTo', 'lineTo', 'stroke']) {
    ctx[method] = (...args) => calls.push([method, ...args]);
  }
  return ctx;
}

const column = { is_vertical: true, x: 20, y: 30, width: 18, height: 48, font_size: 10 };

test('horizontal and older layout elements leave the canvas untouched', () => {
  const ctx = canvas();
  const before = { ...ctx };
  assert.equal(renderVerticalText(ctx, { ...column, is_vertical: false, text: 'abc' }), false);
  assert.equal(renderVerticalText(ctx, { text: 'abc' }), false);
  assert.deepEqual(ctx, before);
  assert.deepEqual(ctx.calls, []);
});

test('upright glyphs advance vertically, including spaces and supplementary characters', () => {
  const ctx = canvas();
  assert.equal(renderVerticalText(ctx, { ...column, text: '日 𠮷本', character_spacing: 2 }), true);
  assert.deepEqual(ctx.calls, [
    ['fillText', '日', 24, 38.5],
    ['fillText', '𠮷', 24, 62.5],
    ['fillText', '本', 24, 74.5],
  ]);
  assert.equal(ctx.letterSpacing, '3px');
  assert.equal(ctx.textBaseline, 'top');
  assert.equal(ctx.textAlign, 'center');
});

test('highlight, double underline and strike follow the vertical column', () => {
  const ctx = canvas();
  renderVerticalText(ctx, { ...column, text: '日本', highlight: '#ffff00', underline: true,
    underline_style: 'double', strikethrough: true });
  assert.deepEqual(ctx.calls, [
    ['fillRect', 20, 30, 18, 48],
    ['fillText', '日', 24, 38.5], ['fillText', '本', 24, 48.5],
    ['beginPath'],
    ['moveTo', 34, 30], ['lineTo', 34, 78],
    ['moveTo', 35, 30], ['lineTo', 35, 78],
    ['moveTo', 29, 30], ['lineTo', 29, 78],
    ['stroke'],
  ]);
});

test('narrow columns do not shift glyphs left and negative spacing stays on the vertical axis', () => {
  const ctx = canvas();
  renderVerticalText(ctx, { ...column, width: 8, text: '日本', character_spacing: -1 });
  assert.deepEqual(ctx.calls, [['fillText', '日', 20, 38.5], ['fillText', '本', 20, 47.5]]);
});

// Exercise the actual page entry points, including the two inline exporters.
// A helper-only test would miss a page that still draws the whole run sideways.
for (const [file, entry] of [
  ['docs.html', 'renderText'], ['print-preview.html', 'renderText'],
  ['../docs/docs.html', 'renderPageToCanvas'],
  ['docs.html', 'inline'], ['render-headless.html', 'inline'],
]) {
  test(`${file}: ${entry} routes vertical text and preserves horizontal drawing`, async () => {
    const html = await readFile(new URL(file, import.meta.url), 'utf8');
    let body;
    if (entry === 'inline') {
      const start = html.indexOf("if (kind === 'text') {");
      const end = html.indexOf("} else if (kind === 'shading')", start);
      assert.ok(start >= 0 && end > start);
      body = `for (const el of [elem]) { const kind = el.kind; ${html.slice(start, end)} } }`;
    } else {
      const start = html.indexOf(`function ${entry}(`);
      const end = html.indexOf('\n}', start) + 2;
      assert.ok(start >= 0 && end > start);
      body = html.slice(start, end) + (entry === 'renderText'
        ? '\nrenderText(ctx, elem);' : '\nrenderPageToCanvas(ctx, { elements: [elem] });');
    }
    const draw = new Function('ctx', 'elem', 'renderVerticalText', body);
    const ctx = canvas();
    draw(ctx, { ...column, kind: 'text', text: '日本' }, renderVerticalText);
    assert.deepEqual(ctx.calls, [['fillText', '日', 24, 38.5], ['fillText', '本', 24, 48.5]]);
    ctx.calls.length = 0;
    draw(ctx, { ...column, kind: 'text', is_vertical: false, text: '日本' }, renderVerticalText);
    assert.deepEqual(ctx.calls, [['fillText', '日本', 20, 38.5]]);
  });
}
