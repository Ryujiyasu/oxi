// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! What the sheet was changed to has to be what the file ends up holding.
//!
//! The browser edits an IR and the editor writes that back into the original
//! archive. Between those two there is room for a change to look right on
//! screen and arrive wrong — or not arrive at all — and every case here is one
//! that did. A date that loses its format is a five-figure serial number; a
//! column width that never reaches the file is a resize that undoes itself the
//! next time the workbook is opened.
//!
//! There are two ways back out. `edit_xlsx` takes a list of cell values and
//! patches only the cells it names, which is cheap. It has nowhere to say what
//! a cell should look like or how wide a column is, so anything past a value
//! goes by `edit_xlsx_from_workbook`, which sends the whole workbook.
//!
//! Run with `node web/save.test.mjs`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const wasm = await import('file://' + join(here, 'oxidocs_wasm.js').replaceAll('\\', '/'));
await wasm.default({ module_or_path: await readFile(join(here, 'oxidocs_wasm_bg.wasm')) });
const { parse_spreadsheet, edit_xlsx, edit_xlsx_from_workbook } = wasm;

let failures = 0;
function is(what, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failures += 1;
  console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}` +
    (ok ? '' : `: got ${JSON.stringify(got)}, wanted ${JSON.stringify(want)}`));
}

const read = async (name) =>
  new Uint8Array(await readFile(join(here, 'samples', name)));
const cellAt = (book, row, col) => {
  const line = book.sheets[0].rows.find((one) => one.index === row);
  return line && line.cells.find((one) => one.col === col);
};

// ── A filled date has to arrive as a date ───────────────────────────────────
//
// A date is a number wearing a format. Filling one down carries the format
// with it, and if the file cannot hold that, the cell arrives correct
// underneath and wrong on the face of it — which is the worst way for this to
// fail, because nothing looks broken until someone reads the number.

const dated = await read('dates-and-times.xlsx');
const opened = parse_spreadsheet(dated.slice());
const format = cellAt(opened, 1, 0).style.number_format;
is('the sample holds a date under a date format', format !== undefined, true);

// What the page does when the fill handle is pulled: the next day, wearing the
// format of the cell it came from.
const rows = opened.sheets[0].rows;
rows.push({ index: 6, height: null, cells: [{
  col: 0, value: { Number: 46053 }, formula: null, runs: [],
  style: { ...cellAt(opened, 1, 0).style },
}] });
rows.sort((one, two) => one.index - two.index);

const narrow = parse_spreadsheet(edit_xlsx(dated.slice(), [
  { sheet_index: 0, row: 6, col: 0, new_value: '46053', value_type: 'number' },
]));
is('saved as a value alone it arrives as a bare number',
  cellAt(narrow, 6, 0).style.number_format, undefined);

const wide = parse_spreadsheet(edit_xlsx_from_workbook(dated.slice(), opened));
is('saved as a workbook it arrives as a date',
  cellAt(wide, 6, 0).style.number_format, format);
is('holding the day it was filled with',
  JSON.stringify(cellAt(wide, 6, 0).value), JSON.stringify({ Number: 46053 }));
is('while the cell it was filled from is untouched',
  JSON.stringify(cellAt(wide, 1, 0)), JSON.stringify(cellAt(opened, 1, 0)));

// ── A dragged column has to stay dragged ────────────────────────────────────
//
// Widths live in a `<cols>` element of their own, which a sheet nobody has
// resized does not have at all. Excel writes 10.625 for a column typed as 10:
// the gutter either side of the text is inside the stored number.

const sample = await read('quarterly-sales.xlsx');
const book = parse_spreadsheet(sample.slice());
const widths = book.sheets[0].col_widths.slice(0, 6);
book.sheets[0].col_widths[2] = 30.625;
book.sheets[0].col_widths[4] = 4.625;
const resized = parse_spreadsheet(edit_xlsx_from_workbook(sample.slice(), book));
const after = resized.sheets[0].col_widths.slice(0, 6);
is('a widened column comes back widened', after[2], 30.625);
is('a narrowed one comes back narrowed', after[4], 4.625);
is('and the columns either side are as they were',
  [after[0], after[1], after[3], after[5]],
  [widths[0], widths[1], widths[3], widths[5]]);

// A cell value written at the same time still lands.
const both = parse_spreadsheet(edit_xlsx_from_workbook(sample.slice(), (() => {
  const one = parse_spreadsheet(sample.slice());
  one.sheets[0].col_widths[1] = 22.625;
  const line = one.sheets[0].rows[0];
  line.cells[0].value = { String: 'changed' };
  return one;
})()));
is('a width and a value travel together',
  [both.sheets[0].col_widths[1], cellAt(both, both.sheets[0].rows[0].index, 0).value],
  [22.625, { String: 'changed' }]);

// ── A dragged row has to stay dragged ───────────────────────────
//
// A row's height comes with a flag saying whether it was chosen or worked out.
// Measured: a row given 33 points WITHOUT the flag came back from Excel at
// 18.75, because Excel threw the number away and worked the height out again
// from what was in the row. So the flag is what makes a drag stick, and
// writing a height without it writes nothing at all.

const tall = parse_spreadsheet(sample.slice());
const lines = tall.sheets[0].rows;
const line = (n) => lines.find((one) => one.index === n);
const wasTall = line(1).height;
line(4).height = 40;
line(4).custom_height = true;
line(5).height = 33;
line(5).custom_height = false;
line(1).height = null;
line(1).custom_height = false;
const sized = parse_spreadsheet(edit_xlsx_from_workbook(sample.slice(), tall));
const back = (n) => sized.sheets[0].rows.find((one) => one.index === n);
is('a chosen height comes back, and comes back chosen',
  [back(4).height, back(4).custom_height], [40, true]);
is('a worked-out one comes back without the flag',
  [back(5).height, Boolean(back(5).custom_height)], [33, false]);
// A row with no height of its own comes back as `undefined` rather than
// `null`: that is how the wasm bridge spells an absent Option, and both mean
// the same thing here.
is('and a row put back on automatic has no height at all',
  back(1).height ?? null, null);
is('which is a change from what it had', wasTall !== null, true);

console.log(failures === 0
  ? '\nwhat was changed is what the file holds'
  : `\n${failures} did not`);
process.exit(failures === 0 ? 0 : 1);
