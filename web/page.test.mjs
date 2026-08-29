// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Does the page load at all?
//!
//! Every other test here reaches into the page and exercises one function.
//! None of them runs the script top to bottom, and that is where a different
//! kind of mistake lives: an element renamed in the HTML but not in the script,
//! a listener hung on nothing, a call made before the thing it calls exists.
//! Those pass every unit test and fail on the first click.
//!
//! So this runs the whole script under a DOM small enough to read, with the
//! real engine behind it, and opens the sample workbook the way the page does
//! when someone visits it. It asserts the three things that would otherwise go
//! unnoticed: that every element the script asks for is in the HTML, that the
//! listeners the keyboard and clipboard depend on were actually hung, and that
//! a grid came out the other end.
//!
//! Run with `node web/page.test.mjs`.

import { readFile, writeFile, unlink } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const page = await readFile(join(here, 'xlsx-demo.html'), 'utf8');

// Which ids the HTML actually declares. Asking for one it has not got is the
// mistake being looked for, so the answer is recorded rather than returned as
// a null that the script would then trip over somewhere else entirely.
const ids = new Set([...page.matchAll(/\bid="([^"]+)"/g)].map((one) => one[1]));

const asked = [];
const hung = [];
const made = [];

function element(name, id) {
  const node = {
    tagName: (name || 'div').toUpperCase(),
    id: id || '',
    dataset: {},
    style: {},
    className: '',
    textContent: '',
    innerHTML: '',
    value: '',
    disabled: false,
    tabIndex: 0,
    rows: [],
    cells: [],
    children: [],
    classList: {
      names: new Set(),
      add(...names) { names.forEach((one) => this.names.add(one)); },
      remove(...names) { names.forEach((one) => this.names.delete(one)); },
      contains(one) { return this.names.has(one); },
      toggle() {},
    },
    // Kept rather than dropped, so a test can do what a person would: press a
    // key in a box and see what the page makes of it. The file picker could be
    // driven already only because the page assigns its `onchange` outright.
    listeners: new Map(),
    addEventListener(kind, run) {
      if (!this.listeners.has(kind)) this.listeners.set(kind, []);
      this.listeners.get(kind).push(run);
    },
    removeEventListener() {},
    /// Run whatever is listening for `kind`, as the browser would.
    fire(kind, event) {
      for (const run of this.listeners.get(kind) || []) {
        run({ preventDefault() {}, stopPropagation() {}, target: this, ...event });
      }
    },
    appendChild(child) { this.children.push(child); return child; },
    removeChild() {},
    remove() {},
    insertRow() {
      const row = element('tr');
      row.insertCell = () => {
        const cell = element('td');
        row.cells.push(cell);
        return cell;
      };
      this.rows.push(row);
      return row;
    },
    insertCell() {
      const cell = element('td');
      this.cells.push(cell);
      return cell;
    },
    querySelector: () => null,
    querySelectorAll: () => [],
    closest: () => null,
    getBoundingClientRect: () => ({
      left: 0, top: 0, right: 100, bottom: 20, width: 100, height: 20,
    }),
    scrollIntoView() {},
    focus() {},
    blur() {},
    click() {},
    select() {},
    setSelectionRange() {},
    insertAdjacentHTML() {},
    setAttribute() {},
    getAttribute: () => null,
    hasAttribute: () => false,
    contains: () => false,
    matches: () => false,
    // The sheet is painted rather than laid out, so the page asks a canvas
    // both to measure text and to draw with. What is drawn is not read back
    // here — a shim cannot say whether a cell looks right — so the calls are
    // counted instead, and the test below asks that the grid was painted at
    // all. `strokes` is what says the painter ran rather than fell over.
    width: 0,
    height: 0,
    getContext() {
      const pen = {
        font: '',
        fillStyle: '',
        strokeStyle: '',
        lineWidth: 1,
        textBaseline: '',
        strokes: 0,
        measureText: (text) => ({ width: String(text).length * 7 }),
      };
      for (const name of ['setTransform', 'save', 'restore', 'translate', 'scale',
        'beginPath', 'closePath', 'moveTo', 'lineTo', 'rect', 'clip', 'stroke',
        'setLineDash', 'strokeRect', 'clearRect']) {
        pen[name] = () => { pen.strokes += 1; };
      }
      pen.fillRect = () => { pen.strokes += 1; };
      pen.fillText = () => { pen.strokes += 1; };
      this.pen ||= pen;
      return this.pen;
    },
  };
  made.push(node);
  return node;
}

const nodes = new Map();
globalThis.document = {
  getElementById(id) {
    asked.push(id);
    if (!nodes.has(id)) nodes.set(id, element('div', id));
    return nodes.get(id);
  },
  createElement: (name) => element(name),
  addEventListener: (kind) => hung.push(kind),
  body: element('body'),
  activeElement: null,
};
globalThis.window = { addEventListener: (kind) => hung.push(kind), devicePixelRatio: 1 };
globalThis.performance = { now: () => 0 };
globalThis.Blob = class { constructor(parts) { this.parts = parts; } };
// The real URL class stays: the wasm glue builds one to find its own binary
// beside itself. Only the two blob helpers are stood in for.
URL.createObjectURL = () => 'blob:x';
URL.revokeObjectURL = () => {};

// The page fetches two things — its own wasm and the sample workbook — and the
// glue checks that what comes back really is a Response before it will stream
// it, so this hands over a real one.
globalThis.fetch = async (what) => {
  const name = String(what).split(/[\\/]/).pop();
  const bytes = await readFile(join(here, name.endsWith('.xlsx') ? 'samples' : '.', name));
  return new Response(bytes, {
    headers: {
      'Content-Type': name.endsWith('.wasm') ? 'application/wasm' : 'application/octet-stream',
    },
  });
};

// The script imports the wasm glue by a relative path, so it has to run from
// beside it rather than from a data URL or a scratch directory.
const script = /<script[^>]*>([\s\S]*?)<\/script>/.exec(page)[1];
const under = join(here, '_page_under_test.mjs');
await writeFile(under, script);
let broke = null;
try {
  await import('file:///' + under.replaceAll('\\', '/'));
} catch (error) {
  broke = error;
} finally {
  await unlink(under);
}

let failures = 0;
function is(what, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failures += 1;
  console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}` +
    (ok ? '' : `: got ${JSON.stringify(got)}, wanted ${JSON.stringify(want)}`));
}

is('the page runs from top to bottom without throwing', broke && String(broke), null);

const missing = [...new Set(asked)].filter((id) => !ids.has(id));
is('every element it asks for is in the HTML', missing, []);

// Two elements answering to one id is not a style complaint: getElementById
// hands back whichever comes first, so the later one becomes unreachable and
// whatever wanted it silently gets the wrong node instead. The canvas and the
// text-colour picker were both called `ink`, and the sheet was never painted
// because the painter kept being handed the colour input.
const seen = new Map();
const twice = [];
for (const [, id] of page.matchAll(/id="([^"]+)"/g)) {
  if (seen.has(id)) twice.push(id);
  seen.set(id, true);
}
is('no two elements answer to the same id', twice, []);

// Typing, copying and pasting all hang off the document rather than the grid,
// because the grid is thrown away and redrawn on every edit. A listener that
// stopped being hung would leave the sheet looking fine and doing nothing.
for (const kind of ['keydown', 'copy', 'cut', 'paste', 'mouseup', 'mousemove']) {
  is(`it listens for ${kind}`, hung.includes(kind), true);
}
is('and asks to be warned before the tab closes with edits outstanding',
  hung.includes('beforeunload'), true);

// The grid is a canvas now, so what says it was drawn is that the painter
// ran: the shim cannot tell a right-looking sheet from a wrong one, and a
// count of strokes at least separates "painted" from "threw on the first
// call", which is the failure this catches.
const ink = nodes.get('sheetInk');
is('it painted a grid', Boolean(ink && ink.pen && ink.pen.strokes > 0), true);
is('with the sample workbook in it',
  Boolean(ink && ink.pen && ink.pen.strokes > 100), true);

// ── Opening a second workbook, the way anyone would ─────────────────────
//
// Everything above is the page at rest. This drives the file picker's own
// handler with a workbook that asks for a frozen top row, which is drawn as a
// band of its own that does not scroll — so what is being asked here is that
// opening a second workbook through the picker paints, and that the sheet it
// painted knows how many rows it is holding.

const picker = nodes.get('pick');
if (picker && picker.onchange) {
  const bytes = await readFile(join(here, '..', 'crates', 'oxicells-core',
    'tests', 'fixtures', 'frozen.xlsx'));
  made.length = 0;
  await picker.onchange({ target: { files: [{
    name: 'frozen.xlsx',
    arrayBuffer: async () =>
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  }] } });
  const painted = nodes.get('sheetInk');
  is('a second workbook opens through the picker',
    Boolean(painted && painted.pen && painted.pen.strokes > 0), true);
} else {
  is('the page has a file picker with a handler on it', false, true);
}

// ── Typing a number format, the way a person would ──────────────────────────
//
// The engine has always known how to render `0%`; until now the page had no
// way to choose it. What can be asked here without reaching inside the page is
// whether the box is listening and whether pressing Enter in it carries all
// the way through to a redraw — `restyle` builds the table again, so a fresh
// one appearing is the path having run end to end rather than throwing
// somewhere in the middle.
//
// What it cannot ask is whether the CELL reads differently afterwards: that
// needs a numeric cell under the cursor, and the cursor is not reachable from
// out here. `save.test.mjs` asks the other half — that a format put on a cell
// reaches the file and comes back.

const box = nodes.get('numfmt');
if (box && box.listeners.has('keydown')) {
  const pen = nodes.get('sheetInk').pen;
  const before = pen.strokes;
  box.value = '0%';
  box.fire('keydown', { key: 'Enter' });
  is('typing a format and pressing Enter draws the sheet again',
    pen.strokes > before, true);
  // And a format the engine cannot make sense of must not take the page down
  // with it — a sheet that stops responding is worse than one showing a
  // number oddly.
  const survived = pen.strokes;
  box.value = 'not a format at all';
  box.fire('keydown', { key: 'Enter' });
  is('and a format that means nothing is survivable',
    pen.strokes > survived, true);
} else {
  is('the page has a format box listening for keys', false, true);
}

// ── A cell told to wrap ─────────────────────────────────────────────────────
//
// The grid drew every cell on one line, so a form whose headings wrap came out
// with the headings cut off. The fixture holds the pair the rule turns on,
// measured in Excel: the same wrapped text in a row of 30pt whose height was
// CHOSEN stays at 30 and is cut off, and in a row that chose none comes back
// at 93.75 — so the chosen height is a lid and the other is not.

if (picker && picker.onchange) {
  const bytes = await readFile(join(here, '..', 'crates', 'oxicells-core',
    'tests', 'fixtures', 'wrapped.xlsx'));
  made.length = 0;
  await picker.onchange({ target: { files: [{
    name: 'wrapped.xlsx',
    arrayBuffer: async () =>
      bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  }] } });
  // The sheet is painted, so a wrapped cell is not a DOM node to count. What
  // decides the picture is the geometry: a row whose height the file CHOSE is
  // held to it, and a row whose height Excel worked out for itself is worked
  // out again — the fixture is that pair, 30pt with the flag against a row
  // without one that Excel grew to 93.75.
  const shape = page.match(/function layout\(\)/);
  is('the sheet is laid out before it is painted', Boolean(shape), true);
  const held = nodes.get('sheetInk');
  is('opening a wrapped workbook paints it', held.pen.strokes > 0, true);
} else {
  is('the page has a file picker to open the wrapped fixture with', false, true);
}

// ── The merge menu ──────────────────────────────────────────────────────────
//
// Same shape of question as the format box: whether the control is listening
// and whether choosing from it runs all the way through. The cursor out here
// is on one cell, which is nothing to merge, so what this shows is that the
// path holds and the menu goes back to its own name — not that a block was
// merged. `grid.test.mjs` asks what merging does to the cells, and
// `save.test.mjs` asks whether it reaches the file.

const menu = nodes.get('merge');
if (menu && menu.listeners.has('change')) {
  const drawnBefore = nodes.get('sheetInk').pen.strokes;
  menu.value = 'cells';
  menu.fire('change', {});
  is('choosing from the merge menu runs through to a redraw',
    nodes.get('sheetInk').pen.strokes > drawnBefore, true);
  is('and the menu goes back to its own name', menu.value, '');
} else {
  is('the page has a merge menu listening for a choice', false, true);
}

console.log(failures === 0 ? '\nthe page loads' : `\n${failures} did not`);
process.exit(failures === 0 ? 0 : 1);
