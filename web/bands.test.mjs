//! Putting a row in, and taking one out, through the engine the page uses.
//!
//! This is the edit that touches most of a workbook: every formula naming a
//! row below the change has to follow it, the merges over it have to stretch
//! or slide, the frozen rows have to stay the rows they were. Each of those
//! fails quietly — a sum that still adds the right four numbers today and the
//! wrong four tomorrow — so they are asked about here against the real wasm
//! rather than only in Rust.
//!
//! Running it against the real bundle is also what catches the things that
//! only go wrong in a browser. The engine reached for `SystemTime::now` so
//! `TODAY()` could answer, which on wasm does not fail politely but panics,
//! and the first number anyone typed took the page down. Rust's own tests were
//! all green.
//!
//! Run with `node web/bands.test.mjs`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const wasm = await import('file://' + join(here, 'oxidocs_wasm.js').replaceAll('\\', '/'));
await wasm.default({ module_or_path: await readFile(join(here, 'oxidocs_wasm_bg.wasm')) });
const { parse_spreadsheet, recalculate_spreadsheet, shift_band } = wasm;

const sample = await readFile(join(here, 'samples', 'quarterly-sales.xlsx'));

let failures = 0;
function is(what, got, want) {
  const same = JSON.stringify(got) === JSON.stringify(want);
  if (!same) failures += 1;
  console.log(`${same ? 'ok' : 'NOT OK'} — ${what}`);
  if (!same) console.log(`   got ${JSON.stringify(got)}\n   want ${JSON.stringify(want)}`);
}

const cellAt = (book, row, col) => {
  const line = book.sheets[0].rows.find((one) => one.index === row);
  return line && line.cells.find((one) => one.col === col);
};
const formulaAt = (book, row, col) => {
  const cell = cellAt(book, row, col);
  return cell && cell.formula ? cell.formula : null;
};
const numberAt = (book, row, col) => {
  const cell = cellAt(book, row, col);
  return cell && cell.value && typeof cell.value === 'object' && 'Number' in cell.value
    ? cell.value.Number
    : null;
};

// A fixed moment, so nothing here depends on the day it is run.
const MOMENT = 45297.5;

// ── What the sample holds ───────────────────────────────────────────────────
//
// Rows 4 to 8 are the figures, row 9 sums each column, and column F sums each
// row. So a row put in at 5 has to be picked up by BOTH: the column totals
// below it, and nothing else.

let book = parse_spreadsheet(sample.slice());
const sheet = book.sheets[0].name;
is('the sample sums a column in row 9', formulaAt(book, 9, 1), 'SUM(B4:B8)');
is('and a row in column F', formulaAt(book, 4, 5), 'SUM(B4:E4)');
const wasTotal = numberAt(book, 9, 1);

// ── Putting one in ──────────────────────────────────────────────────────────

book = shift_band(book, sheet, true, 5, 1);
is('the row below the insert came down', formulaAt(book, 10, 1), 'SUM(B4:B9)');
is('and the range it sums grew to hold the new row',
  formulaAt(book, 10, 1), 'SUM(B4:B9)');
is('a row total moved without its columns changing',
  formulaAt(book, 4, 5), 'SUM(B4:E4)');
is('the one that was on row 5 is now on row 6',
  formulaAt(book, 6, 5), 'SUM(B6:E6)');
is('and row 5 is empty', cellAt(book, 5, 1), undefined);

book = recalculate_spreadsheet(book, MOMENT);
is('the column total still comes to what it did — an empty row adds nothing',
  numberAt(book, 10, 1), wasTotal);

// ── Taking it out again ─────────────────────────────────────────────────────

book = shift_band(book, sheet, false, 0, 0);   // a no-op, to be sure one is
is('asking for nothing changes nothing', formulaAt(book, 10, 1), 'SUM(B4:B9)');

book = shift_band(book, sheet, true, 5, -1);
is('taking the row out put everything back', formulaAt(book, 9, 1), 'SUM(B4:B8)');
is('including the row totals', formulaAt(book, 5, 5), 'SUM(B5:E5)');
book = recalculate_spreadsheet(book, MOMENT);
is('and the answer is the one it started with', numberAt(book, 9, 1), wasTotal);

// ── Taking out a row something was counting on ──────────────────────────────
//
// Excel writes `#REF!` into the formula itself rather than sliding the
// reference along, because sliding it would answer confidently with somebody
// else's number.

let broken = parse_spreadsheet(sample.slice());
broken = shift_band(broken, sheet, true, 4, -5);  // every figure row, gone
is('a sum whose rows all went says so', formulaAt(broken, 4, 1), 'SUM(#REF!)');

// ── A column, the same way ──────────────────────────────────────────────────

let sideways = parse_spreadsheet(sample.slice());
sideways = shift_band(sideways, sheet, false, 1, 1);
is('a column put in at B moved the row total right',
  formulaAt(sideways, 4, 6), 'SUM(C4:F4)');
is('and the column totals moved with their columns',
  formulaAt(sideways, 9, 2), 'SUM(C4:C8)');

console.log(failures === 0 ? '\nrows and columns move' : `\n${failures} did not`);
process.exit(failures === 0 ? 0 : 1);
