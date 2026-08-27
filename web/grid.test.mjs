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
const tables = ['LISTS', 'TAIL', 'INDENT', 'DAY', 'EPOCH'].map(liftBinding).join('\n');

const lifted = ['region', 'spread', 'eachInRegion', 'trim', 'tally', 'stepInside',
  'pasteGrid', 'clearSelection', 'goTo', 'contentOf', 'selectionText', 'fieldOf',
  'asGrid', 'change', 'restore', 'undo', 'redo', 'remember', 'edited',
  'markSaved', 'markSteps', 'put', 'heldAt',
  'takeColumn', 'takeRow', 'takeAll',
  'seriesOf', 'numberOf', 'fitLine', 'kindOf', 'runsOf', 'inList', 'planRun',
  'runValue', 'fillLine', 'fillTo', 'wrapped', 'unitOf', 'sameStyle',
  'widthForPixels', 'setWidth', 'dropCut',
  'padFor', 'fitWidth', 'fitColumn', 'shownText',
  'restyle', 'allWear', 'toggle', 'alignTo', 'setHeight',
  'holdPanes', 'freezeHere',
  'monthSeries', 'monthAt'].map(lift).join('\n');

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
const sheet = { rows: [], col_widths: [], merge_cells: [] };
const book = { default_style: { font_size: 11 } };
// Every character is ten pixels wide at 11pt and scales with the size, so the
// arithmetic under test is visible in the answer rather than buried in a font.
let inkPer = 10;
function inkOf(text, style) {
  return text.length * inkPer * ((style.font_size || 11) / 11);
}
const ink = (per) => { inkPer = per; };
const format_cell_number = (value) => String(value);
const sheetNow = () => sheet;
let digitWidth = 8;
let sizing = null;
let cutFrom = null;
// The toolbar's Insert and Delete follow the selection, so painting it asks
// whether whole rows or columns are picked. The page grew this dependency when
// bands were added.
const sayBands = () => {};
const showWidth = () => {};
const showHeight = () => {};
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
let beyondValues = false;
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
// Put a value in wearing a format, as the parser would have read it.
const wear = (row, col, text, format) => {
  const was = step;
  step = null;
  put(row, col, text, { number_format: format });
  step = was;
  changes.clear();
  pristine.clear();
  beyondValues = false;
};
const widthAt = (col) => sheet.col_widths[col];
// Mark a block for a cut, the way the cut handler does: remember where it is
// and what it wears, and leave it exactly where it stands.
const mark = (r1, c1, r2, c2) => {
  select(r1, c1, r2, c2);
  const box = region();
  const styles = new Map();
  for (let row = box.top; row <= box.bottom; row++) {
    for (let col = box.left; col <= box.right; col++) {
      const one = heldAt(row, col);
      if (one) styles.set((row - box.top) + ',' + (col - box.left), { ...one.style });
    }
  }
  cutFrom = { sheet: shown, ...box, styles };
};
// What a paste would move, as the paste handler works it out.
const marked = () => cutFrom;
const asDate = (serial) => new Date(EPOCH + serial * DAY);
const asSerial = (year, month, day) =>
  Math.round((Date.UTC(year, month, day) - EPOCH) / DAY);
const monthLength = (year, month) => new Date(Date.UTC(year, month + 1, 0)).getUTCDate();
const sheetOf = () => sheet;
const frozenAt = () => [sheet.frozen_rows || 0, sheet.frozen_cols || 0];
const markFrozen = () => {};

// A table of known geometry: the header strip 20 tall, the row-label column 40
// wide, every row 25 tall and every column 100 wide. Standing one up by hand
// is what makes the arithmetic readable in the answers below.
function aTable(rows, cols) {
  const cell = (extra) => ({
    dataset: {},
    style: {},
    classList: {
      names: new Set(),
      add(...names) { names.forEach((one) => this.names.add(one)); },
      remove(...names) { names.forEach((one) => this.names.delete(one)); },
      contains(one) { return this.names.has(one); },
    },
    getBoundingClientRect: () => extra,
    ...{},
  });
  const table = { rows: [] };
  const head = { children: [], getBoundingClientRect: () => ({ height: 20, width: 0 }) };
  head.children.push(cell({ height: 20, width: 40 }));
  for (let col = 0; col < cols; col++) {
    const one = cell({ height: 20, width: 100 });
    one.dataset.headCol = String(col);
    head.children.push(one);
  }
  table.rows.push(head);
  for (let row = 1; row <= rows; row++) {
    const line = { children: [], getBoundingClientRect: () => ({ height: 25, width: 0 }) };
    const label = cell({ height: 25, width: 40 });
    label.dataset.headRow = String(row);
    line.children.push(label);
    for (let col = 0; col < cols; col++) {
      const one = cell({ height: 25, width: 100 });
      one.dataset.row = String(row);
      one.dataset.col = String(col);
      line.children.push(one);
    }
    table.rows.push(line);
  }
  return table;
}
const lineAt = (row) => sheet.rows.find((one) => one.index === row);
const heightAt = (row) => { const one = lineAt(row); return one ? one.height : undefined; };
const chosenAt = (row) => { const one = lineAt(row); return one ? one.custom_height : undefined; };
const valueAt = (row, col) => {
  const one = heldAt(row, col);
  return one ? one.value : undefined;
};
const beyond = () => beyondValues;
const depth = () => undone.length;
const styleAt = (row, col) => {
  const cell = heldAt(row, col);
  return cell ? cell.style : null;
};
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
  cutFrom = null;
  sheet.frozen_rows = 0;
  sheet.frozen_cols = 0;
  sheet.merge_cells.length = 0;
  sheet.rows.length = 0;
  sheet.col_widths.length = 0;
  beyondValues = false;
  changes.clear(); pristine.clear();
  undone.length = 0; redone.length = 0;
};
export { select, seat, box, cells, seed, put, tally, countBox, stepInside,
         takeColumn, takeRow, takeAll, fillTo, fitLine, wrapped, unitOf,
         wear, styleAt, sameStyle, widthForPixels, setWidth, widthAt,
         beyond, depth, dropCut, mark, marked,
         padFor, fitWidth, fitColumn, ink, sheetOf,
         toggle, alignTo, allWear, valueAt, setHeight, heightAt, chosenAt,
         lineAt, holdPanes, aTable, freezeHere, frozenAt,
         monthSeries, monthAt, asDate, asSerial,
         pasteGrid, clearSelection, selectionText, asGrid, fieldOf, region,
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

// ── Dates ─────────────────────────────────────────────────
//
// A date is a number wearing a format, and that format is the only thing
// saying so. Measured off Excel with the serials read back rather than the
// display, which truncates to hashes in a narrow column:
//
//     2026/01/30 alone -> 46053, 46054, 46055     one day at a time
//     10:30 alone      -> 11:30, 12:30            one hour at a time
//     05 Jan, 12 Jan   -> 46041, 46048, 46055     the gap between them
//
// The first of those is the one that matters: a plain number selected alone
// repeats, so without the format being read a dragged date would sit there
// unchanged, which is not what anyone drags a date for.

is('a date format is a day', grid.unitOf({ style: { number_format: 'mm-dd-yy' } }), 1);
is('a long one too', grid.unitOf({ style: { number_format: 'yyyy年m月d日' } }), 1);
is('a time format is an hour',
  grid.unitOf({ style: { number_format: 'h:mm' } }), 1 / 24);
is('elapsed hours are still hours',
  grid.unitOf({ style: { number_format: '[h]:mm:ss' } }), 1 / 24);
is('a date and time together is a day',
  grid.unitOf({ style: { number_format: 'm/d/yy h:mm' } }), 1);
is('a plain number is neither', grid.unitOf({ style: { number_format: '0.00' } }), null);
is('and nor is one with no format at all', grid.unitOf({ style: {} }), null);
// A format can hold letters inside quotes that mean nothing about its kind.
is('letters in quotes do not make a date',
  grid.unitOf({ style: { number_format: '0.0"days"' } }), null);
is('nor does an escaped one',
  grid.unitOf({ style: { number_format: '0.0\\d' } }), null);

grid.forget();
grid.wear(1, 0, '46052', 'mm-dd-yy');
pull(1, 0, 1, 0, 4);
is('a lone date moves a day at a time', column(0, 1, 4),
  ['46052', '46053', '46054', '46055']);
is('and the cells it filled wear its format', grid.styleAt(3, 0),
  { number_format: 'mm-dd-yy' });

grid.forget();
grid.wear(1, 0, '0.4375', 'h:mm');
pull(1, 0, 1, 0, 3);
is('a lone time moves an hour at a time', column(0, 1, 3),
  ['0.4375', '0.479166666666667', '0.520833333333333']);

grid.forget();
grid.wear(1, 0, '46027', 'mm-dd-yy');
grid.wear(2, 0, '46034', 'mm-dd-yy');
pull(1, 0, 2, 0, 5);
is('two dates carry the gap between them', column(0, 1, 5),
  ['46027', '46034', '46041', '46048', '46055']);

// The format follows every kind of run, not only dates: filling is a copy of
// the cell, and half a cell arriving is worse than none.
grid.forget();
grid.wear(1, 0, 'total', '@');
pull(1, 0, 1, 0, 3);
is('repeated text carries its format too', grid.styleAt(3, 0),
  { number_format: '@' });

// Undo has to put back the formatting as well as the value, or taking back a
// fill leaves the cells looking like something they no longer hold.
grid.forget();
grid.wear(1, 0, '46052', 'mm-dd-yy');
pull(1, 0, 1, 0, 4);
grid.undo();
is('undoing a fill takes the values back', column(0, 1, 4),
  ['46052', '', '', '']);
is('and the formatting with them', grid.styleAt(3, 0), {});

is('two cells with no style wear the same one', grid.sameStyle(null, {}), true);
is('and two with different formats do not',
  grid.sameStyle({ number_format: 'a' }, { number_format: 'b' }), false);

// ── Dragging a column wider ─────────────────────────────────
//
// A stored width times the digit width IS the pixel width: Excel stores 10.625
// for a column someone typed 10 into, and draws it 85 pixels wide. That is the
// whole conversion, and it is exact rather than approximate — the gutter either
// side of the text is already counted inside the stored number, which is why
// the stored number is not the one Excel's own box shows.

is('eighty-five pixels is the width Excel stores for a column typed as ten',
  grid.widthForPixels(85), 10.625);
is('and a hundred and seventeen is what the sample holds',
  grid.widthForPixels(117), 14.625);
is('a column cannot be dragged past nothing', grid.widthForPixels(-20), 0);

grid.forget();
grid.setWidth(2, 240);
is('dragging a column writes its width', grid.widthAt(2), 30);
// Columns before it are padded with zero, which is how the parser records
// "this column had no <col> entry of its own" — the same as never having been
// mentioned, and not a width of nothing.
is('and columns before it are left saying nothing', grid.widthAt(1), 0);

// A width is not something a list of cell values can carry, so saving it has
// to take the path that sends the whole workbook.
is('resizing needs more than a list of values', grid.beyond(), true);

grid.undo();
is('undo puts the column back', grid.widthAt(2), 0);
grid.redo();
is('and redo widens it again', grid.widthAt(2), 30);

// Dragging to the width it already has is not a change, so it does not fill
// the undo stack with nothing.
grid.forget();
grid.setWidth(1, 96);
const deep = grid.depth();
grid.setWidth(1, 96);
is('dragging to the width it already has changes nothing', grid.depth(), deep);

// ── Cutting ─────────────────────────────────────────────
//
// A cut takes nothing away. It says which cells are going to move; they move
// when they are pasted, and Escape calls the whole thing off. Emptying them at
// the moment of the cut — which is what this used to do — loses the work of
// anyone who cuts a column and then changes their mind.

grid.forget();
grid.seed(1, 0, 'a');
grid.seed(2, 0, 'b');
grid.mark(1, 0, 2, 0);
is('a cut leaves the cells where they are', grid.cells(),
  [['1,0', 'a'], ['2,0', 'b']]);
is('and remembers which ones it marked',
  [grid.marked().top, grid.marked().bottom], [1, 2]);

grid.dropCut();
is('Escape calls it off', grid.marked(), null);
is('with everything still there', grid.cells(), [['1,0', 'a'], ['2,0', 'b']]);

// Pasting after a cut moves the block: it arrives, and the cells it came from
// empty, as one action.
grid.forget();
grid.seed(1, 0, 'a');
grid.seed(2, 0, 'b');
grid.mark(1, 0, 2, 0);
const moving = grid.marked();
grid.select(5, 3, 5, 3);
grid.change(() => {
  for (let row = moving.top; row <= moving.bottom; row++) grid.put(row, moving.left, '');
  grid.pasteGrid([['a'], ['b']], moving);
});
is('the block arrives where it was pasted', grid.cells(),
  [['5,3', 'a'], ['6,3', 'b']]);
grid.undo();
is('and one undo puts the whole move back', grid.cells(),
  [['1,0', 'a'], ['2,0', 'b']]);

// The formatting travels with the cells, or a column of dates arrives as a
// column of five-figure numbers.
grid.forget();
grid.wear(1, 0, '46052', 'mm-dd-yy');
grid.mark(1, 0, 1, 0);
const dated = grid.marked();
grid.select(4, 2, 4, 2);
grid.change(() => {
  grid.put(1, 0, '');
  grid.pasteGrid([['46052']], dated);
});
is('a cut date arrives still wearing its format', grid.styleAt(4, 2),
  { number_format: 'mm-dd-yy' });
is('and the cell it came from is empty', grid.cells(), [['4,2', '46052']]);

// Writing to the sheet at all makes the mark stale, because the cells the
// clipboard holds are no longer the cells that are there.
grid.forget();
grid.seed(1, 0, 'a');
grid.mark(1, 0, 1, 0);
grid.change(() => grid.put(9, 9, 'something else'));
is('editing anything drops the mark', grid.marked(), null);

// ── Fitting a column to what is in it ──────────────────────
//
// Double-clicking a column's edge makes it as wide as the widest thing it has
// to show. The padding either side was read off Excel: 'M', 'MM', 'MMM' come
// out 26, 40 and 54 pixels — fourteen a letter and twelve left over — and 'i'
// and a kanji agree on the same twelve from quite different letter widths.
// It grows with the type, stepping by two for every seven pixels of em.

is('twelve pixels at the size nearly everything uses', grid.padFor(11), 12);
is('ten at the smallest', grid.padFor(6), 10);
is('still twelve just below the step', grid.padFor(11.5), 12);
is('fourteen just above it', grid.padFor(12), 14);
is('fourteen up to the next step', grid.padFor(16.5), 14);
is('sixteen above that', grid.padFor(17), 16);
is('and eighteen further up', grid.padFor(24), 18);
is('a size of nothing is read as the ordinary one', grid.padFor(0), 12);

// ── what the width is measured from ─────────────────────────────────────────

grid.forget();
grid.ink(10);
grid.seed(1, 0, 'abcde');
is('the widest text plus the padding', grid.fitWidth(0), 50 + 12);

grid.seed(5, 0, 'abcdefghij');
is('the whole column, not the first row of it', grid.fitWidth(0), 100 + 12);

grid.seed(3, 1, 'x');
is('and each column answers for itself', grid.fitWidth(1), 10 + 12);

is('a column with nothing in it has no answer', grid.fitWidth(9), null);

// A cell states its own size, and the padding follows that cell rather than
// the workbook's.
grid.forget();
grid.wear(1, 0, 'ab', null);
grid.styleAt(1, 0).font_size = 24;
is('a bigger cell is measured, and padded, at its own size',
  grid.fitWidth(0), 2 * 10 * (24 / 11) + 18);

// ── what it leaves out ──────────────────────────────────────────────────────
//
// Merged cells are ignored entirely, which is why autofit so often looks like
// it did nothing. A wrapped cell is left out too: Excel picks a width there to
// suit the row's height, and the same words in the same font came out 55px in
// one sheet and 79px in another.

grid.forget();
grid.ink(10);
grid.seed(1, 0, 'a very long piece of text');
grid.sheetOf().merge_cells.push({ start_row: 1, start_col: 0, end_row: 1, end_col: 1 });
is('a merged cell is not what the column is fitted to', grid.fitWidth(0), null);

grid.forget();
grid.seed(1, 0, 'short');
grid.seed(2, 0, 'a very long piece of text');
grid.styleAt(2, 0).wrap_text = true;
is('and neither is a wrapped one', grid.fitWidth(0), 5 * 10 + 12);

// An indent pushes the text in, and the column makes room for it: measured as
// twelve pixels a level, so 'Hello' goes from 47px to 83px at three.
grid.forget();
grid.seed(1, 0, 'abc');
grid.styleAt(1, 0).indent = 3;
is('an indent is made room for', grid.fitWidth(0), 30 + 3 * 12 + 12);

// ── and what it then does ───────────────────────────────────────────────────

grid.forget();
grid.ink(10);
grid.seed(1, 0, 'abcde');
grid.fitColumn(0);
is('fitting writes the width, rounded up to a whole pixel',
  grid.widthAt(0), Math.ceil(62) / 8);
is('and it is a change like any other, to be taken back', grid.beyond(), true);
grid.undo();
is('undo puts the column back', grid.widthAt(0), 0);

// A column of nothing but merged cells is left exactly as it is rather than
// closed up to nothing.
grid.forget();
grid.seed(1, 0, 'text');
grid.sheetOf().merge_cells.push({ start_row: 1, start_col: 0, end_row: 1, end_col: 1 });
grid.setWidth(0, 200);
const held = grid.widthAt(0);
grid.fitColumn(0);
is('a column with nothing to fit is left alone', grid.widthAt(0), held);

// ── Bold, italic and lining up ────────────────────────────────
//
// Ctrl+B is muscle memory, and a grid where it does nothing feels broken in a
// way that is hard to name. Excel's rule for a block that is only partly bold
// was measured rather than assumed: such a block reports its boldness as
// neither true nor false, and Excel sets it to the opposite of THAT — so one
// press turns the whole block on, and the next turns it all off.

grid.forget();
grid.seed(1, 0, 'a');
grid.select(1, 0, 1, 0);
grid.toggle('bold');
is('one cell goes bold', grid.styleAt(1, 0).bold, true);
grid.toggle('bold');
is('and comes back', grid.styleAt(1, 0).bold, false);

grid.forget();
grid.seed(1, 0, 'a');
grid.seed(2, 0, 'b');
grid.seed(3, 0, 'c');
grid.select(1, 0, 1, 0);
grid.toggle('bold');
grid.select(1, 0, 3, 0);
is('a block only partly bold does not count as bold',
  grid.allWear((style) => Boolean(style.bold)), false);
grid.toggle('bold');
is('so one press takes the whole block bold',
  [1, 2, 3].map((row) => grid.styleAt(row, 0).bold), [true, true, true]);
grid.toggle('bold');
is('and the next takes it all off',
  [1, 2, 3].map((row) => grid.styleAt(row, 0).bold), [false, false, false]);

grid.forget();
grid.seed(1, 0, 'a');
grid.select(1, 0, 1, 0);
grid.alignTo('center');
is('a cell can be centred', grid.styleAt(1, 0).horizontal_align, 'center');
grid.alignTo('right');
is('and moved to the right', grid.styleAt(1, 0).horizontal_align, 'right');
grid.alignTo('right');
is('and asking for the same again takes the alignment away',
  grid.styleAt(1, 0).horizontal_align, undefined);

// ── What must NOT change ────────────────────────────────────
//
// Changing how a cell looks used to send its value round through its text, and
// the text of a cell holding the string "007" is "007", which reads back as
// the number seven. Bolding a part number would have quietly renumbered it.

grid.forget();
grid.put(1, 0, 'x');
grid.sheetOf().rows[0].cells[0].value = { String: '007' };
grid.select(1, 0, 1, 0);
grid.toggle('bold');
is('bolding a cell does not re-read what it holds',
  grid.valueAt(1, 0), { String: '007' });
is('it only changes how it looks', grid.styleAt(1, 0).bold, true);

// A formula is not turned into its own text either.
grid.forget();
grid.seed(1, 0, '=A9+1');
grid.select(1, 0, 1, 0);
grid.toggle('italic');
is('and a formula stays a formula', grid.cells(), [['1,0', '=A9+1']]);

// ── and it is a change like any other ──────────────────────────

grid.forget();
grid.seed(1, 0, 'a');
grid.seed(2, 0, 'b');
grid.select(1, 0, 2, 0);
grid.toggle('bold');
is('changing a look needs the wider way of saving', grid.beyond(), true);
grid.undo();
is('and one undo unbolds the whole block',
  [1, 2].map((row) => grid.styleAt(row, 0).bold), [undefined, undefined]);
grid.redo();
is('redo bolds it again',
  [1, 2].map((row) => grid.styleAt(row, 0).bold), [true, true]);

// ── Dragging a row taller ────────────────────────────────────
//
// The file states a row's height in points where the drag is in pixels, and it
// carries a flag beside the number saying whether the height was CHOSEN or
// worked out from the contents. A drag is a choice, and the flag is not
// decoration: a row given 33 points without it came back from Excel at 18.75,
// because Excel threw the number away and worked the height out again.

grid.forget();
grid.seed(3, 0, 'a');
grid.setHeight(3, 40);
is('a dragged row is stated in points', grid.heightAt(3), 30);
is('and marked as chosen rather than worked out', grid.chosenAt(3), true);
is('a height is not something a list of values can carry', grid.beyond(), true);

grid.undo();
is('undo puts the height back', grid.heightAt(3), null);
is('and unmarks it', grid.chosenAt(3), false);
grid.redo();
is('redo makes it tall again', grid.heightAt(3), 30);

// A row nobody has written to can still be dragged.
grid.forget();
grid.setHeight(9, 96);
is('a row with nothing in it can be dragged', grid.heightAt(9), 72);

// Dragging to the height it already has, once chosen, is not a change.
grid.forget();
grid.setHeight(2, 40);
const settled = grid.depth();
grid.setHeight(2, 40);
is('dragging to the height it already has changes nothing', grid.depth(), settled);

// A row whose height was worked out rather than chosen IS changed by a drag to
// the same number, because the drag is what makes it stick.
grid.forget();
grid.seed(4, 0, 'a');
grid.lineAt(4).height = 30;
grid.lineAt(4).custom_height = false;
grid.setHeight(4, 40);
is('dragging a computed height to the same number makes it chosen',
  grid.chosenAt(4), true);

// ── Holding the frozen rows and columns in view ─────────────
//
// A workbook says `<pane ySplit="1" state="frozen"/>` to mean "keep the top
// row while the rest scrolls", counting in cells. The header strip is already
// pinned, so a held row sits below it and each held row below the one before.
// The offsets are added up from what the browser actually laid out, because
// the sizes a file asks for and the sizes it gets are not the same thing.
//
// The table below is stood up by hand: header 20 tall, row labels 40 wide,
// rows 25 tall, columns 100 wide.

const pinned = (table, row, col) => {
  const line = table.rows[row];
  const cell = line.children[col + 1];
  return {
    top: cell.style.top,
    left: cell.style.left,
    pinned: cell.classList.contains('held'),
    both: cell.classList.contains('both'),
  };
};

let table = grid.aTable(6, 4);
grid.holdPanes(table, { frozen_rows: 1, frozen_cols: 0 });
is('one held row sits just under the header strip',
  pinned(table, 1, 0).top, '20px');
is('and is pinned', pinned(table, 1, 0).pinned, true);
is('the row under it is not', pinned(table, 2, 0).pinned, false);

table = grid.aTable(6, 4);
grid.holdPanes(table, { frozen_rows: 3, frozen_cols: 0 });
is('three held rows stack up under the header',
  [1, 2, 3].map((row) => pinned(table, row, 0).top), ['20px', '45px', '70px']);
is('and the fourth is loose', pinned(table, 4, 0).pinned, false);

table = grid.aTable(6, 4);
grid.holdPanes(table, { frozen_rows: 0, frozen_cols: 1 });
is('one held column sits beside the row labels',
  pinned(table, 1, 0).left, '40px');
is('the column beside it is loose', pinned(table, 1, 1).pinned, false);
is('and every row of the held column is pinned, not just the first',
  [1, 2, 5].every((row) => pinned(table, row, 0).pinned), true);

table = grid.aTable(6, 4);
grid.holdPanes(table, { frozen_rows: 0, frozen_cols: 2 });
is('two held columns stack up beside the labels',
  [0, 1].map((col) => pinned(table, 1, col).left), ['40px', '140px']);

// The corner: a cell held both ways has to stay put in both directions, and
// has to pass over the cells held in only one.
table = grid.aTable(6, 4);
grid.holdPanes(table, { frozen_rows: 3, frozen_cols: 2 });
const corner = pinned(table, 2, 1);
is('a cell in a held row AND a held column is held both ways',
  [corner.top, corner.left], ['45px', '140px']);
is('and outranks the ones held one way only', corner.both, true);
is('a cell in a held row alone is not', pinned(table, 2, 3).both, false);
is('nor is one in a held column alone', pinned(table, 5, 0).both, false);

// A sheet that asks for nothing is left entirely alone.
table = grid.aTable(6, 4);
grid.holdPanes(table, { frozen_rows: 0, frozen_cols: 0 });
is('a sheet with no frozen panes is not touched',
  [1, 3].every((row) => !pinned(table, row, 0).pinned), true);
is('and nothing is given a position', pinned(table, 1, 0).top, undefined);

// The row labels are held across as well, or they slide out from under the
// rows they are numbering.
table = grid.aTable(6, 4);
grid.holdPanes(table, { frozen_rows: 1, frozen_cols: 0 });
is('the row labels are pinned across even with no held columns',
  table.rows[2].children[0].classList.contains('held'), true);

// ── Setting the freeze ─────────────────────────────────────
//
// Excel freezes at the active cell rather than at a line you pick: standing on
// C4 and freezing holds everything above and left of it, which is three rows
// and two columns. The same command undoes it once a sheet is frozen, which is
// what makes it one control rather than two.

grid.forget();
grid.select(4, 2, 4, 2);
grid.freezeHere();
is('freezing at C4 holds three rows and two columns', grid.frozenAt(), [3, 2]);

grid.freezeHere();
is('and doing it again lets everything go', grid.frozenAt(), [0, 0]);

grid.forget();
grid.select(2, 0, 2, 0);
grid.freezeHere();
is('at A2 it holds the top row and no columns', grid.frozenAt(), [1, 0]);

grid.forget();
grid.select(1, 1, 1, 1);
grid.freezeHere();
is('at B1 it holds the first column and no rows', grid.frozenAt(), [0, 1]);

// At A1 there is nothing above or left, so there is nothing to hold.
grid.forget();
grid.select(1, 0, 1, 0);
grid.freezeHere();
is('at A1 there is nothing to hold', grid.frozenAt(), [0, 0]);

// It is a change like any other.
grid.forget();
grid.select(3, 1, 3, 1);
grid.freezeHere();
is('a freeze needs the wider way of saving', grid.beyond(), true);
grid.undo();
is('undo lets it go', grid.frozenAt(), [0, 0]);
grid.redo();
is('redo holds it again', grid.frozenAt(), [2, 1]);

// Unfreezing a sheet that was frozen goes back to what it was, not to nothing.
grid.forget();
grid.sheetOf().frozen_rows = 5;
grid.sheetOf().frozen_cols = 3;
grid.select(2, 0, 2, 0);
grid.freezeHere();
is('a frozen sheet is unfrozen wherever the cursor is', grid.frozenAt(), [0, 0]);
grid.undo();
is('and undo puts back the freeze it had', grid.frozenAt(), [5, 3]);

// ── Dates that step by calendar months ─────────────────────
//
// Two dates set a step the way two numbers do: 5 Jan and 12 Jan continue a
// week at a time. But 31 Jan and 28 Feb continue 31 Mar, 30 Apr, 31 May —
// Excel has seen a MONTH between them and is stepping by the calendar. A line
// through those two serials gives 28-day steps and lands on 28 Mar instead.
//
// Every row below was read off Excel by `tools/metrics/_xlsx_fill_dates.py`.

// Readable both ways, so a failure names a date rather than a five-figure
// number nobody can check by eye.
const dayOf = (serial) => grid.asDate(serial).toISOString().slice(0, 10);
const serialOf = (text) => {
  const [year, month, day] = text.split('-').map(Number);
  return grid.asSerial(year, month - 1, day);
};

const carry = (seed, over) => {
  const found = grid.monthSeries(seed.map(serialOf));
  if (!found) return null;
  const out = [];
  for (let at = seed.length; at < seed.length + over; at++) {
    out.push(dayOf(grid.monthAt(found, at)));
  }
  return out;
};

is('two month ends carry the month, and the day is clamped as it goes',
  carry(['2026-01-31', '2026-02-28'], 4),
  ['2026-03-31', '2026-04-30', '2026-05-31', '2026-06-30']);
is('the same day of the month carries too',
  carry(['2026-01-15', '2026-02-15'], 2), ['2026-03-15', '2026-04-15']);
is('so does the first of the month',
  carry(['2026-01-01', '2026-02-01'], 2), ['2026-03-01', '2026-04-01']);
is('three months at a time', carry(['2026-01-01', '2026-04-01'], 3),
  ['2026-07-01', '2026-10-01', '2027-01-01']);
is('two months at a time, clamping where the month is short',
  carry(['2026-01-31', '2026-03-31'], 3),
  ['2026-05-31', '2026-07-31', '2026-09-30']);
is('and it runs backwards', carry(['2026-03-31', '2026-02-28'], 3),
  ['2026-01-31', '2025-12-31', '2025-11-30']);
is('three dates a month apart carry on the same way',
  carry(['2026-01-31', '2026-02-28', '2026-03-31'], 3),
  ['2026-04-30', '2026-05-31', '2026-06-30']);
is('a year is twelve months', carry(['2026-01-01', '2027-01-01'], 3),
  ['2028-01-01', '2029-01-01', '2030-01-01']);
is('and a year from a leap day finds the next one',
  carry(['2024-02-29', '2025-02-28'], 3),
  ['2026-02-28', '2027-02-28', '2028-02-29']);
is('the new year is not a boundary', carry(['2026-12-31', '2027-01-31'], 3),
  ['2027-02-28', '2027-03-31', '2027-04-30']);

// ── and what is NOT a month ─────────────────────────────────────────────────
//
// These are the rows that make the rule a test rather than a guess. A gap that
// happens to land on a month end is not a month, and reading one as a month
// would drag a fortnightly rota onto the wrong days.

is('twenty-nine days ending on a month end is not a month',
  grid.monthSeries(['2026-01-30', '2026-02-28'].map(serialOf)), null);
is('nor is thirty days ending on one',
  grid.monthSeries(['2026-01-29', '2026-02-28'].map(serialOf)), null);
is('nor thirty days that miss the month end',
  grid.monthSeries(['2026-01-15', '2026-02-14'].map(serialOf)), null);
is('nor a month apart on different days',
  grid.monthSeries(['2026-01-10', '2026-02-20'].map(serialOf)), null);
is('a week is a week', grid.monthSeries(['2026-01-05', '2026-01-12'].map(serialOf)), null);
is('and two dates in one month have no months between them',
  grid.monthSeries(['2026-01-05', '2026-01-06'].map(serialOf)), null);
is('one date on its own is not a series at all',
  grid.monthSeries([serialOf('2026-01-31')]), null);
is('and neither is one carrying a time',
  grid.monthSeries([serialOf('2026-01-15') + 0.5, serialOf('2026-02-15') + 0.5]), null);

// The three dates have to be evenly spaced in months, not merely all on the
// same day.
is('an uneven run of months is not a series',
  grid.monthSeries(['2026-01-15', '2026-02-15', '2026-05-15'].map(serialOf)), null);

// ── through the fill itself ─────────────────────────────────────────────────
//
// The rule above only matters if a date run reaches it, which means the cells
// have to be recognised as dates by the format they wear.

grid.forget();
grid.wear(1, 0, String(serialOf('2026-01-31')), 'mm-dd-yy');
grid.wear(2, 0, String(serialOf('2026-02-28')), 'mm-dd-yy');
pull(1, 0, 2, 0, 5);
is('pulling two month ends down fills calendar months',
  column(0, 3, 5).map((one) => dayOf(Number(one))),
  ['2026-03-31', '2026-04-30', '2026-05-31']);
is('and the cells keep their date format', grid.styleAt(4, 0),
  { number_format: 'mm-dd-yy' });

// The same two numbers with no date format on them are just numbers.
grid.forget();
grid.seed(1, 0, String(serialOf('2026-01-31')));
grid.seed(2, 0, String(serialOf('2026-02-28')));
pull(1, 0, 2, 0, 3);
is('without a date format they are numbers, and step by the gap',
  column(0, 3, 3), [String(serialOf('2026-02-28') + 28)]);

console.log(failures === 0
  ? '\nthe grid behaves'
  : `\n${failures} did not`);
process.exit(failures === 0 ? 0 : 1);
