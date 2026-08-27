// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Typing a formula into the sheet, against the engine that reads the corpus.
//!
//! A grid where `=B2+C2` shows nothing until the file is saved and reopened is
//! not one anyone would use, so the page asks the engine for the answer after
//! every edit. Two things have to hold for that to be worth doing. The answer
//! must be the same one the file would have when Excel opens it — a browser
//! that computes its own answers is worse than one that computes none. And the
//! asking must not stall the sheet: recalculation rebuilds the whole
//! dependency graph, which on the largest workbook in the corpus takes
//! thirty-eight seconds, so the page has to decide beforehand whether it can
//! afford to ask.
//!
//! Run with `node web/formula.test.mjs`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const wasm = await import('file://' + join(here, 'oxidocs_wasm.js').replaceAll('\\', '/'));
await wasm.default({ module_or_path: await readFile(join(here, 'oxidocs_wasm_bg.wasm')) });
const { parse_spreadsheet, recalculate_spreadsheet, edit_xlsx } = wasm;

const page = await readFile(join(here, 'xlsx-demo.html'), 'utf8');

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

const bits = await import('data:text/javascript,' + encodeURIComponent(`
let book = null;
let anyFormulas = false;
let liveCalc = true;
let weight = 0;
const PATIENCE = 300;
const calcNote = null;
const recalculate_spreadsheet = null;
const format_cell_number = (value) => String(value);
${['calcCost', 'sayCalc', 'contentOf', 'shownText'].map(lift).join('\n')}
const use = (one) => { book = one; };
const affordable = () => { const cost = calcCost(); return { cost, formulas: anyFormulas }; };
export { use, affordable, contentOf, shownText };
`));

let failures = 0;
function is(what, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failures += 1;
  console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}` +
    (ok ? '' : `: got ${JSON.stringify(got)}, wanted ${JSON.stringify(want)}`));
}

const sample = new Uint8Array(await readFile(join(here, 'samples', 'quarterly-sales.xlsx')));

/** Write `text` into a cell the way the page does, then recalculate. */
function typeInto(book, row, col, text) {
  const sheet = book.sheets[0];
  let line = sheet.rows.find((one) => one.index === row);
  if (!line) { line = { index: row, cells: [], height: null }; sheet.rows.push(line); }
  let cell = line.cells.find((one) => one.col === col);
  if (!cell) { cell = { col, value: 'Empty', style: {}, formula: null, runs: [] }; line.cells.push(cell); }
  cell.formula = text.startsWith('=') ? text.slice(1) : null;
  if (cell.formula) cell.value = { String: '' };
  else if (text === '') cell.value = 'Empty';
  else if (text.trim() !== '' && Number.isFinite(Number(text))) cell.value = { Number: Number(text) };
  else cell.value = { String: text };
  return recalculate_spreadsheet(book);
}

const cellAt = (book, row, col) => {
  const line = book.sheets[0].rows.find((one) => one.index === row);
  return line && line.cells.find((one) => one.col === col);
};
/** What the grid draws in the cell — the answer, not the formula. */
const shows = (book, row, col) => {
  const cell = cellAt(book, row, col);
  return cell ? bits.shownText(cell) : '';
};
/** What the formula bar shows — the formula, when there is one. */
const behind = (book, row, col) => {
  const cell = cellAt(book, row, col);
  return cell ? bits.contentOf(cell) : '';
};

// ── The answer appears, and it is the right one ─────────────────────────────
//
// A free row well below anything the sample holds, so nothing is disturbed.

let book = parse_spreadsheet(sample.slice());
const free = Math.max(...book.sheets[0].rows.map((r) => r.index)) + 3;

book = typeInto(book, free, 0, '10');
book = typeInto(book, free, 1, '32');
book = typeInto(book, free, 2, '=A' + free + '+B' + free);
is('a typed sum shows its answer', shows(book, free, 2), '42');
is('and the formula bar still shows the formula',
  behind(book, free, 2), '=A' + free + '+B' + free);

book = typeInto(book, free + 1, 2, '=SUM(A' + free + ':B' + free + ')');
is('so does a function over a range', shows(book, free + 1, 2), '42');

book = typeInto(book, free + 2, 2, '=C' + free + '*2');
is('a formula that reads another formula sees its value, not its text',
  shows(book, free + 2, 2), '84');

// Changing an input has to move everything downstream of it, not just the
// cell that was typed into — the whole reason to recalculate rather than
// evaluate one cell.
book = typeInto(book, free, 0, '20');
is('changing an input moves what depends on it', shows(book, free, 2), '52');
is('and what depends on that', shows(book, free + 2, 2), '104');

book = typeInto(book, free + 3, 2, '=1/0');
is('a formula that cannot work says so', shows(book, free + 3, 2), '#DIV/0!');

book = typeInto(book, free + 4, 2, '=IF(A' + free + '>10,"over","under")');
is('a comparison works', shows(book, free + 4, 2), 'over');

// ── The browser and the file agree ──────────────────────────────────────────
//
// The answer shown while typing must be the answer the file holds afterwards.
// If these ever disagree, the browser is computing something of its own, which
// is worse than computing nothing.

const written = edit_xlsx(sample.slice(), [
  { sheet_index: 0, row: free, col: 0, new_value: '20', value_type: 'number' },
  { sheet_index: 0, row: free, col: 1, new_value: '32', value_type: 'number' },
  { sheet_index: 0, row: free, col: 2, new_value: `=A${free}+B${free}`, value_type: 'formula' },
]);
const reopened = parse_spreadsheet(written);
is('what the sheet showed is what the saved file holds',
  shows(reopened, free, 2), '52');
is('and the file kept the formula, not just its answer',
  behind(reopened, free, 2), '=A' + free + '+B' + free);

// ── Knowing when not to ask ─────────────────────────────────────────────────

bits.use(parse_spreadsheet(sample.slice()));
const small = bits.affordable();
is('the sample has formulas', small.formulas, true);
is('and is cheap enough to work out as you type', small.cost < 300, true);

// A workbook with no formulas is never worth recalculating, however large.
const bare = { sheets: [{ name: 'S', rows: [{ index: 1, cells: [
  { col: 0, value: { String: 'x' }, style: {}, formula: null, runs: [] },
] }] }] };
bits.use(bare);
is('a sheet with no formulas has nothing to work out', bits.affordable().formulas, false);

// The estimate has to refuse the workbooks that would hang. Half a million
// cells and twenty thousand formulas is the shape of the corpus's largest,
// which takes thirty-eight seconds.
const heavy = { sheets: [{ name: 'S', rows: [] }] };
for (let row = 1; row <= 400; row++) {
  const cells = [];
  for (let col = 0; col < 1250; col++) {
    cells.push({ col, value: { Number: 1 }, style: {},
      formula: col % 20 === 0 ? 'A1' : null, runs: [] });
  }
  heavy.sheets[0].rows.push({ index: row, cells, height: null });
}
bits.use(heavy);
const big = bits.affordable();
is('a workbook that would hang is refused before it is asked', big.cost > 300, true);

console.log(failures === 0
  ? '\nformulas are worked out, and only when they can be'
  : `\n${failures} did not`);
process.exit(failures === 0 ? 0 : 1);
