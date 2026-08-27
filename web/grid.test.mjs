// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The sheet's selection and clipboard, against what Excel does.
//!
//! The parts of a grid people notice are the ones with arithmetic behind them:
//! where Enter lands when a block is selected, where a pasted block goes, what
//! the status bar says a column comes to, whether a cell holding a tab survives
//! being copied. Each is a small sum that is easy to get subtly wrong and hard
//! to spot by looking — the cursor goes somewhere plausible, just not where
//! Excel puts it, and a copied block comes back the wrong shape.
//!
//! The page has no build step, so the functions are lifted out of the HTML and
//! given a sheet made of a Map. That way this tests the code the page actually
//! runs, and cannot go on passing against a version it no longer has. The
//! price is that the harness below has to name everything those functions
//! reach for: a ReferenceError here means the page grew a dependency, and the
//! fix is to lift or stub the name it asks for, not to work around it.
//! Run with `node web/grid.test.mjs`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const page = await readFile(
  join(dirname(fileURLToPath(import.meta.url)), 'xlsx-demo.html'), 'utf8');

/** The source of one top-level function declaration, braces balanced. */
function lift(name) {
  const at = page.indexOf(`function ${name}(`);
  if (at < 0) throw new Error(`the page no longer has ${name}()`);
  let depth = 0;
  let open = false;
  for (let i = page.indexOf('{', at); i < page.length; i++) {
    if (page[i] === '{') { depth += 1; open = true; }
    else if (page[i] === '}') depth -= 1;
    if (open && depth === 0) return page.slice(at, i + 1);
  }
  throw new Error(`${name}() is never closed`);
}

const lifted = ['region', 'spread', 'eachInRegion', 'trim', 'tally', 'stepInside',
  'pasteGrid', 'clearSelection', 'goTo', 'contentOf', 'selectionText', 'fieldOf',
  'asGrid', 'change', 'restore', 'undo', 'redo', 'remember', 'edited',
  'markSaved', 'markSteps', 'put', 'heldAt'].map(lift).join('\n');

// Everything the lifted code reaches for that lives in the page's DOM. The
// sheet is the same shape the parser hands back — rows holding cells — so the
// page's own put() runs here unaltered, which is the point: writing a cell is
// where the undo step and the save count are recorded, and stubbing it out
// would leave exactly that untested.
const grid = await import('data:text/javascript,' + encodeURIComponent(`
let at = { row: 1, col: 0 };
let anchor = { row: 1, col: 0 };
let reach = { row: 1, col: 0 };
let tabHome = null;
let dragging = false;
const sheet = { rows: [] };
const sheetNow = () => sheet;
const countBox = { textContent: '' };
const whereBox = { textContent: '' };
const formulaBox = { value: '' };
const cellBox = { innerHTML: '' };
const draw = () => {};
const paint = () => {};
let shown = 0;
const changes = new Map();
const pristine = new Map();
const undone = [];
const redone = [];
let step = null;
const DEPTH = 100;
const describe = () => {};
const saveButton = { disabled: true, textContent: '' };
const backButton = { disabled: true };
const onButton = { disabled: true };
let anyFormulas = false;
// Working out formulas is the engine's job and has its own test; here it only
// has to not happen, so that what a cell holds is exactly what was typed.
const recompute = () => {};
const cellAt = () => null;
const tableNow = () => null;
const describeCell = () => {};
const columnName = (n) => String.fromCharCode(65 + n);
const extent = () => ({ rows: 500, cols: 50 });
${lifted}
// Put a value in as the parser would, without it counting as an edit: this is
// what the file already held when it was opened.
const seed = (row, col, text) => {
  const was = step;
  step = null;
  put(row, col, text);
  step = was;
  changes.clear();
  pristine.clear();
};
// Select a block the way a drag does: anchor at the corner it started from,
// reach at the one under the pointer, and the active cell left on the anchor.
const select = (r1, c1, r2, c2) => {
  at = { row: r1, col: c1 };
  anchor = { row: r1, col: c1 };
  reach = { row: r2, col: c2 };
};
const seat = () => ({ row: at.row, col: at.col });
const box = () => region();
// Only the cells that hold something: clearing one leaves the cell in place
// with nothing in it, which is not the same as never having written there.
const cells = () => {
  const out = [];
  for (const line of sheet.rows) {
    for (const cell of line.cells) {
      const text = contentOf(cell);
      if (text !== '') out.push([line.index + ',' + cell.col, text]);
    }
  }
  return out.sort();
};
const unsaved = () => changes.size;
// Start over as if a fresh file had been opened.
const forget = () => {
  sheet.rows.length = 0;
  changes.clear(); pristine.clear();
  undone.length = 0; redone.length = 0;
};
export { select, seat, box, cells, seed, put, tally, countBox, stepInside,
         pasteGrid, clearSelection, selectionText, asGrid, fieldOf,
         change, undo, redo, unsaved, forget };
`));

let failures = 0;
function is(what, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failures += 1;
  console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}` +
    (ok ? '' : `: got ${JSON.stringify(got)}, wanted ${JSON.stringify(want)}`));
}

// ── Walking a selected block ────────────────────────────────────────────────
//
// Select B2:C4 — three rows, two columns — and the active cell starts on the
// anchor. Enter walks it down the first column then jumps to the top of the
// second; Tab walks across the first row then drops to the start of the next.
// Both wrap round to the beginning rather than leaving the block, which is
// what makes a range the thing you select before filling it in.

grid.select(2, 1, 4, 2);
const down = [];
for (let i = 0; i < 7; i++) { grid.stepInside(1, 0); down.push(grid.seat()); }
is('Enter walks down the column, then to the top of the next', down, [
  { row: 3, col: 1 }, { row: 4, col: 1 },
  { row: 2, col: 2 }, { row: 3, col: 2 }, { row: 4, col: 2 },
  { row: 2, col: 1 }, { row: 3, col: 1 },
]);
is('and the block itself does not move', grid.box(),
  { top: 2, left: 1, bottom: 4, right: 2 });

grid.select(2, 1, 4, 2);
const across = [];
for (let i = 0; i < 7; i++) { grid.stepInside(0, 1); across.push(grid.seat()); }
is('Tab walks across the row, then to the start of the next', across, [
  { row: 2, col: 2 },
  { row: 3, col: 1 }, { row: 3, col: 2 },
  { row: 4, col: 1 }, { row: 4, col: 2 },
  { row: 2, col: 1 }, { row: 2, col: 2 },
]);

grid.select(2, 1, 4, 2);
grid.stepInside(-1, 0);
is('Shift+Enter from the first cell wraps to the last',
  grid.seat(), { row: 4, col: 2 });

// Dragging from the far corner selects the same rectangle.
grid.select(4, 2, 2, 1);
is('a block dragged out backwards is the same block', grid.box(),
  { top: 2, left: 1, bottom: 4, right: 2 });

// ── Where a pasted block lands ──────────────────────────────────────────────

grid.forget();
grid.select(5, 2, 5, 2);
grid.pasteGrid([['a', 'b'], ['c', 'd']]);
is('a block lands from the active cell', grid.cells(),
  [['5,2', 'a'], ['5,3', 'b'], ['6,2', 'c'], ['6,3', 'd']]);
is('and is left selected', grid.box(), { top: 5, left: 2, bottom: 6, right: 3 });

grid.forget();
grid.select(1, 0, 3, 1);
grid.pasteGrid([['x']]);
is('one value fills the whole selected block', grid.cells(),
  [['1,0', 'x'], ['1,1', 'x'], ['2,0', 'x'],
   ['2,1', 'x'], ['3,0', 'x'], ['3,1', 'x']]);

grid.forget();
grid.select(2, 0, 2, 0);
grid.pasteGrid([['x']]);
is('one value into one cell goes in once', grid.cells(), [['2,0', 'x']]);

// A ragged block writes the cells it has and leaves the rest alone.
grid.forget();
grid.select(1, 0, 1, 0);
grid.pasteGrid([['a', 'b', 'c'], ['d']]);
is('a ragged block writes only what it holds', grid.cells(),
  [['1,0', 'a'], ['1,1', 'b'], ['1,2', 'c'], ['2,0', 'd']]);

// ── Clearing ────────────────────────────────────────────────────────────────

grid.forget();
for (let row = 1; row <= 4; row++) grid.seed(row, 0, String(row));
grid.seed(9, 0, 'kept');
grid.select(1, 0, 3, 0);
grid.clearSelection();
is('Delete clears the block and nothing outside it', grid.cells(),
  [['4,0', '4'], ['9,0', 'kept']]);

// ── What the selection comes to ─────────────────────────────────────────────
//
// Excel's status bar: how many cells hold something, and — when any of them
// are numbers — what those come to. It is how a column gets checked without
// writing a formula anywhere.

grid.forget();
[10, 20, 30].forEach((n, i) => grid.seed(i + 1, 0, String(n)));
grid.seed(4, 0, 'text');
grid.select(1, 0, 5, 0);
grid.tally();
is('numbers are summed and averaged, text only counted',
  grid.countBox.textContent, 'Average 20   Sum 60   Count 4');

grid.forget();
grid.seed(1, 0, 'a');
grid.seed(2, 0, 'b');
grid.select(1, 0, 2, 0);
grid.tally();
is('with no numbers there is nothing to total',
  grid.countBox.textContent, 'Count 2');

grid.select(1, 0, 1, 0);
grid.tally();
is('one cell says nothing at all', grid.countBox.textContent, '');

// A total that is plainly a round number is shown as one, rather than with the
// tail of float noise that 0.1 + 0.2 + 0.3 leaves behind.
grid.forget();
['0.1', '0.2', '0.3'].forEach((n, i) => grid.seed(i + 1, 0, n));
grid.select(1, 0, 3, 0);
grid.tally();
is('a round total is shown as round',
  grid.countBox.textContent, 'Average 0.2   Sum 0.6   Count 3');

// ── The clipboard ───────────────────────────────────────────────────────────
//
// Copying cells out and pasting them back is a round trip through one flat
// string. The only thing keeping a cell that itself holds a tab or a newline
// from arriving as three cells is the quoting, and that fails silently: the
// block simply comes back the wrong shape.

grid.forget();
grid.seed(1, 0, 'a'); grid.seed(1, 1, 'b'); grid.seed(1, 2, 'outside');
grid.seed(2, 0, 'c'); grid.seed(2, 1, 'd');
grid.select(1, 0, 2, 1);
is('a copied block holds only what is inside it', grid.selectionText(), 'a\tb\nc\td');
grid.seed(2, 0, '');
is('an empty cell inside it is an empty field', grid.selectionText(), 'a\tb\n\td');

const out = (block) =>
  block.map((line) => line.map(grid.fieldOf).join('\t')).join('\n');
const trip = [
  [[['a', 'b'], ['c', 'd']], 'a plain block'],
  [[['']], 'one empty cell'],
  [[['a', ''], ['', 'd']], 'gaps inside a block'],
  [[['has\ttab', 'plain']], 'a cell holding a tab'],
  [[['two\nlines', 'plain']], 'a cell holding a newline'],
  [[['a\r\nb']], 'a cell holding a CRLF'],
  [[['say "hi"', 'plain']], 'a cell holding quotes'],
  [[['"lead', 'trail"']], 'a cell starting or ending with a quote'],
  [[['12', '-3.5'], ['0', '1e6']], 'numbers'],
  [[['=SUM(A1:A2)']], 'a formula'],
  [[['日本語', 'ＭＳ ゴシック']], 'wide characters'],
  [[['a'], ['b'], ['c']], 'a single column'],
];
for (const [block, what] of trip) is(`${what} survives the clipboard`, grid.asGrid(out(block)), block);

// What arrives from elsewhere is not always what we wrote. Excel ends every
// block with a newline, and plenty of sources use CRLF throughout.
is('Excel’s CRLF and trailing newline read as one block',
  grid.asGrid('a\tb\r\nc\td\r\n'), [['a', 'b'], ['c', 'd']]);
is('a trailing newline is a line ending, not an empty row',
  grid.asGrid('a\tb\n'), [['a', 'b']]);
is('and a row without one is still a row', grid.asGrid('a\tb'), [['a', 'b']]);
is('a blank row in the middle stays blank',
  grid.asGrid('a\n\nb'), [['a'], [''], ['b']]);
is('nothing at all pastes nothing', grid.asGrid(''), [['']]);

// ── Taking it back ──────────────────────────────────────────
//
// Undo has to hand back a whole action, not a cell: pasting a block and then
// undoing must clear all of it, in one press rather than four.

grid.forget();
grid.select(1, 0, 1, 0);
grid.change(() => grid.put(1, 0, 'first'));
grid.change(() => grid.put(1, 0, 'second'));
grid.undo();
is('undo puts back the previous value', grid.cells(), [['1,0', 'first']]);
grid.undo();
is('and again leaves the cell empty', grid.cells(), []);
grid.undo();
is('undo with nothing left to undo does nothing', grid.cells(), []);
grid.redo();
grid.redo();
is('redo walks back up the same path', grid.cells(), [['1,0', 'second']]);

grid.forget();
grid.select(2, 1, 2, 1);
grid.pasteGrid([['a', 'b'], ['c', 'd']]);
grid.undo();
is('undoing a paste clears the whole block in one press', grid.cells(), []);
grid.redo();
is('and redoing puts the whole block back', grid.cells(),
  [['2,1', 'a'], ['2,2', 'b'], ['3,1', 'c'], ['3,2', 'd']]);
is('what an undo touched is left selected, so it is clear what moved',
  grid.box(), { top: 2, left: 1, bottom: 3, right: 2 });

grid.forget();
for (let row = 1; row <= 3; row++) grid.seed(row, 0, String(row));
grid.select(1, 0, 3, 0);
grid.clearSelection();
is('a block delete is one step', grid.cells(), []);
grid.undo();
is('and comes back whole', grid.cells(),
  [['1,0', '1'], ['2,0', '2'], ['3,0', '3']]);

// A change made after undoing forks: what was undone is no longer ahead.
grid.forget();
grid.select(1, 0, 1, 0);
grid.change(() => grid.put(1, 0, 'a'));
grid.undo();
grid.change(() => grid.put(1, 0, 'b'));
grid.redo();
is('a change after undo discards what was ahead', grid.cells(), [['1,0', 'b']]);

// ── What the save button counts ─────────────────────────────
//
// Cells that differ from the file, not keystrokes. Typing a value and taking
// it back leaves the file with nothing to write, and saying otherwise claims
// an edit that is not there.

grid.forget();
grid.select(1, 0, 1, 0);
grid.change(() => grid.put(1, 0, 'typed'));
is('a typed cell is one thing to save', grid.unsaved(), 1);
grid.undo();
is('taking it back leaves nothing to save', grid.unsaved(), 0);
grid.redo();
is('and putting it back leaves one again', grid.unsaved(), 1);

grid.forget();
grid.seed(4, 0, 'from the file');
grid.select(4, 0, 4, 0);
grid.change(() => grid.put(4, 0, 'changed'));
grid.change(() => grid.put(4, 0, 'from the file'));
is('typing a cell back to what the file held is not a change', grid.unsaved(), 0);

grid.forget();
grid.select(1, 0, 1, 0);
grid.change(() => grid.put(1, 0, 'x'));
grid.change(() => grid.put(1, 0, 'y'));
is('editing one cell twice is one thing to save', grid.unsaved(), 1);

console.log(failures === 0
  ? '\nthe grid behaves'
  : `\n${failures} did not`);
process.exit(failures === 0 ? 0 : 1);
