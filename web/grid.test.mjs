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

/** The source of one top-level `const NAME = ...;`, brackets balanced. */
function liftBinding(name) {
  const at = page.indexOf(`const ${name} = `);
  if (at < 0) throw new Error(`the page no longer has ${name}`);
  let depth = 0;
  for (let i = at; i < page.length; i++) {
    const ch = page[i];
    if ('([{'.includes(ch)) depth += 1;
    else if (')]}'.includes(ch)) depth -= 1;
    else if (ch === ';' && depth === 0) return page.slice(at, i + 1);
  }
  throw new Error(`${name} is never closed`);
}

// The fill's tables come across as they are written, not as a copy: the eleven
// lists were read out of Excel, and a test holding its own copy of them would
// go on passing after the page's had drifted.
const tables = ['LISTS', 'TAIL'].map(liftBinding).join('\n');

const lifted = ['region', 'spread', 'eachInRegion', 'trim', 'tally', 'stepInside',
  'pasteGrid', 'clearSelection', 'goTo', 'contentOf', 'selectionText', 'fieldOf',
  'asGrid', 'change', 'restore', 'undo', 'redo', 'remember', 'edited',
  'markSaved', 'markSteps', 'put', 'heldAt',
  'takeColumn', 'takeRow', 'takeAll',
  'seriesOf', 'numberOf', 'fitLine', 'kindOf', 'runsOf', 'inList', 'planRun',
  'runValue', 'fillLine', 'fillTo', 'wrapped'].map(lift).join('\n');

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
${tables}
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
let filling = null;
// Moving a formula's references is the engine's job and has its own test.
// Here it only has to be visible when a formula is carried across.
const translate_formula = (text, rows, cols) =>
  text.slice(1) + '[' + rows + ',' + cols + ']';
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
         takeColumn, takeRow, takeAll, fillTo, fitLine, wrapped,
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

// ── Taking a whole column or row ──────────────────────────────
//
// Clicking a header takes the line it heads, end to end, and leaves the active
// cell at the near end — so typing starts at the top of the column you just
// clicked, not at the bottom of the sheet. The harness reports 500 rows and 50
// columns, which is what the sheet on screen would be.

grid.forget();
grid.select(9, 9, 9, 9);
grid.takeColumn(3, false);
is('a column header takes the column, top to bottom', grid.box(),
  { top: 1, left: 3, bottom: 500, right: 3 });
is('and leaves the cursor at the top of it', grid.seat(), { row: 1, col: 3 });

grid.takeColumn(5, true);
is('shift-clicking another header takes the run between them', grid.box(),
  { top: 1, left: 3, bottom: 500, right: 5 });
is('and the cursor stays where the run began', grid.seat(), { row: 1, col: 3 });

grid.takeColumn(1, true);
is('a run can be extended backwards past where it started', grid.box(),
  { top: 1, left: 1, bottom: 500, right: 3 });

grid.forget();
grid.select(9, 9, 9, 9);
grid.takeRow(4, false);
is('a row header takes the row, end to end', grid.box(),
  { top: 4, left: 0, bottom: 4, right: 50 });
is('and leaves the cursor at the near end', grid.seat(), { row: 4, col: 0 });

grid.takeRow(6, true);
is('rows extend the same way', grid.box(),
  { top: 4, left: 0, bottom: 6, right: 50 });

grid.takeAll();
is('the corner takes the sheet', grid.box(),
  { top: 1, left: 0, bottom: 500, right: 50 });
is('with the cursor at A1', grid.seat(), { row: 1, col: 0 });

// ── Pulling the corner ─────────────────────────────────────
//
// Every case below was read off Excel by `tools/metrics/_xlsx_fill_series.py`,
// which pulls the handle down a real sheet and reports what landed. None of it
// is reasoned out here: the fitted line, and the way a lone number behaves
// differently inside a block, both look wrong until they are measured.

// The Japanese seeds, named so the cases below stay readable.
const _JP_SUN = '日';
const _JP_MON = '月';
const _JP_TUE = '火';
const _JP_WED = '水';
const _JP_THU = '木';
const _JP_SAT = '土';
const _JP_M1 = '睦月';
const _JP_M2 = '如月';
const _JP_M3 = '弥生';
const _JP_Q1 = '第1四半期';
const _JP_Q2 = '第2四半期';
const _JP_Q3 = '第3四半期';
const _JP_Z1 = '子';
const _JP_Z2 = '丑';
const _JP_Z3 = '寅';
const _JP_I = 'い';

const pull = (r1, c1, r2, c2, to) => { grid.select(r1, c1, r2, c2); grid.fillTo(to, c1); };
const column = (col, from, to) => {
  const out = [];
  for (let row = from; row <= to; row++) {
    const found = grid.cells().find(([key]) => key === row + ',' + col);
    out.push(found ? found[1] : '');
  }
  return out;
};
/** Seed a column downwards, pull it `over` further, and read the lot back. */
const filled = (seed, over) => {
  grid.forget();
  seed.forEach((one, i) => grid.seed(i + 1, 0, one));
  pull(1, 0, seed.length, 0, seed.length + over);
  return column(0, 1, seed.length + over);
};

is('a lone number repeats', filled(['5'], 3), ['5', '5', '5', '5']);
is('two numbers set the step', filled(['1', '2'], 4),
  ['1', '2', '3', '4', '5', '6']);
is('the step can be any size', filled(['10', '20', '30'], 2),
  ['10', '20', '30', '40', '50']);
is('and can go down', filled(['10', '8'], 3), ['10', '8', '6', '4', '2']);

// Excel fits a line through the values rather than repeating the last gap. It
// shows 5.333333 in a narrow column; the value behind it is sixteen thirds.
is('unevenly spaced numbers follow the fitted line', filled(['1', '2', '4'], 2),
  ['1', '2', '4', '5.33333333333333', '6.83333333333333']);
is('and so do these', filled(['1', '4', '5'], 2),
  ['1', '4', '5', '7.33333333333333', '9.33333333333333']);

is('text repeats', filled(['total'], 2), ['total', 'total', 'total']);
is('text ending in digits counts up', filled(['Item 1'], 3),
  ['Item 1', 'Item 2', 'Item 3', 'Item 4']);
is('and keeps the width it was written with', filled(['A001'], 2),
  ['A001', 'A002', 'A003']);
is('two numbered texts carry the step in their digits',
  filled(['Item 1', 'Item 2'], 4),
  ['Item 1', 'Item 2', 'Item 3', 'Item 4', 'Item 5', 'Item 6']);
is('a block with no series in it repeats round', filled(['a', 'b'], 4),
  ['a', 'b', 'a', 'b', 'a', 'b']);

// A list Excel knows continues rather than repeating. All eleven were read out
// of Excel itself with GetCustomListContents.
is('an English weekday continues', filled(['Sun'], 3),
  ['Sun', 'Mon', 'Tue', 'Wed']);
is('so does the long form', filled(['Sunday'], 3),
  ['Sunday', 'Monday', 'Tuesday', 'Wednesday']);
is('and a month', filled(['Jan'], 3), ['Jan', 'Feb', 'Mar', 'Apr']);
is('a Japanese weekday continues', filled([_JP_MON], 3),
  [_JP_MON, _JP_TUE, _JP_WED, _JP_THU]);
is('and wraps round the end of the list', filled([_JP_SAT], 2),
  [_JP_SAT, _JP_SUN, _JP_MON]);
is('the old month names continue', filled([_JP_M1], 2),
  [_JP_M1, _JP_M2, _JP_M3]);
is('so do the quarters', filled([_JP_Q1], 2), [_JP_Q1, _JP_Q2, _JP_Q3]);
is('and the zodiac', filled([_JP_Z1], 2), [_JP_Z1, _JP_Z2, _JP_Z3]);
is('a word Excel has no list for just repeats', filled([_JP_I], 2),
  [_JP_I, _JP_I, _JP_I]);
is('and neither has it one for a bare digit as text', filled(['1'], 2),
  ['1', '1', '1']);

// The rule behind all of it: a mixed block is split into runs of neighbours of
// the same kind, and each run carries on by itself. A number that repeats when
// selected alone counts up when it is a run inside a block.
is('a number and a word each follow their own rule', filled(['1', 'a'], 3),
  ['1', 'a', '2', 'a', '3']);
is('whichever order they are in', filled(['a', '1'], 4),
  ['a', '1', 'a', '2', 'a', '3']);
is('a lone number in a block counts up by one, not by itself',
  filled(['2', 'a'], 4), ['2', 'a', '3', 'a', '4', 'a']);
is('however large it is', filled(['100', 'a'], 4),
  ['100', 'a', '101', 'a', '102', 'a']);
is('two numbers beside each other are one run with their own step',
  filled(['2', '4', 'a'], 4), ['2', '4', 'a', '6', '8', 'a', '10']);
is('numbers split by text are two runs, each counting by one',
  filled(['1', 'a', '9'], 4), ['1', 'a', '9', '2', 'a', '10', '3']);
is('neighbouring words are one run and repeat together',
  filled(['1', 'a', 'b'], 4), ['1', 'a', 'b', '2', 'a', 'b', '3']);
is('a list and a number interleave', filled(['Mon', '1'], 4),
  ['Mon', '1', 'Tue', '2', 'Wed', '3']);

// Two members of one list are a run, and carry the stride between them.
is('two neighbouring weekdays carry on', filled(['Sun', 'Mon'], 4),
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri']);
is('two days apart keeps that gap', filled(['Mon', 'Wed'], 4),
  ['Mon', 'Wed', 'Fri', 'Sun', 'Tue', 'Thu']);
is('and a gap that runs backwards wraps round the list',
  filled(['Wed', 'Mon'], 4),
  ['Wed', 'Mon', 'Sat', 'Thu', 'Tue', 'Sun']);
is('but two different lists are two runs', filled(['Mon', 'Jan'], 4),
  ['Mon', 'Jan', 'Tue', 'Feb', 'Wed', 'Mar']);
is('as are two numbered texts with different prefixes',
  filled(['Item 1', 'Row 5'], 4),
  ['Item 1', 'Row 5', 'Item 2', 'Row 6', 'Item 3', 'Row 7']);
is('and numbered text keeps counting inside a block',
  filled(['Item 1', 'a'], 4), ['Item 1', 'a', 'Item 2', 'a', 'Item 3', 'a']);

// Pulling upwards runs every rule the other way.
grid.forget();
grid.seed(5, 0, '10');
grid.seed(6, 0, '20');
grid.select(5, 0, 6, 0);
grid.fillTo(3, 0);
is('pulling upwards counts the other way', column(0, 3, 6),
  ['-10', '0', '10', '20']);

// A formula is carried across with its references moved by how far it went.
grid.forget();
grid.seed(1, 0, '=B1+C1');
pull(1, 0, 1, 0, 3);
is('a formula moves with the cell it lands in', column(0, 1, 3),
  ['=B1+C1', '=B1+C1[1,0]', '=B1+C1[2,0]']);

// What a fill wrote comes back in one press, like any other action.
grid.forget();
grid.seed(1, 0, '1');
grid.seed(2, 0, '2');
pull(1, 0, 2, 0, 8);
grid.undo();
is('a fill is one step', column(0, 1, 8),
  ['1', '2', '', '', '', '', '', '']);

// The pieces the rules are built from.
is('a line through evenly spaced values has that step',
  grid.fitLine([2, 4, 6]), { slope: 2, base: 2 });
is('a line through one value is flat', grid.fitLine([7]), { slope: 0, base: 7 });
is('a line through equal values is flat', grid.fitLine([3, 3, 3]),
  { slope: 0, base: 3 });
is('an index off the end of a list comes round', grid.wrapped(8, 7), 1);
is('and off the start too', grid.wrapped(-1, 7), 6);

console.log(failures === 0
  ? '\nthe grid behaves'
  : `\n${failures} did not`);
process.exit(failures === 0 ? 0 : 1);
