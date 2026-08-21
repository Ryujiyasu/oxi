// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The row model, against the cases measured off Excel.
//!
//! Every number below came from a COM or PDF measurement of Excel itself, not
//! from reading this code back to itself. Run with `node web/row-model.test.mjs`.

import { countLines, mayBreak, explainRow, sheetDefault } from './row-model.js';

let failures = 0;
function is(what, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failures += 1;
  console.log(`${ok ? 'ok  ' : 'FAIL'} ${what}: got ${JSON.stringify(got)}` +
    (ok ? '' : `, wanted ${JSON.stringify(want)}`));
}

// Excel breaks a line of fullwidth glyphs wherever it likes, so a box that
// holds eighteen of them holds eighteen.
const em = 16;
const box = 18 * em;
const glyphs = n => Array.from({ length: n }, () => 'あ');
const flat = letters => letters.map(() => em);

is('sixty glyphs take four lines of eighteen',
  countLines(glyphs(60), flat(glyphs(60)), box), 4);

// Measured from Excel's PDF: a line that would start with a forbidden
// character pushes its neighbour down rather than hanging it past the edge.
const touten = [...glyphs(18), '、', ...glyphs(30)];
is('a 、 one past the edge pushes one character down',
  countLines(touten, flat(touten), box), 3);

const two = [...glyphs(17), '、', '、', ...glyphs(30)];
is('two forbidden characters push two down',
  countLines(two, flat(two), box), 3);

const opening = [...glyphs(17), '「', ...glyphs(30)];
is('an opening bracket may not end a line',
  countLines(opening, flat(opening), box), 3);

is('a break is allowed between two ideographs', mayBreak('あ', 'あ'), true);
is('but not before a 。', mayBreak('あ', '。'), false);
is('nor after a 「', mayBreak('「', 'あ'), false);
is('never inside a web address', mayBreak(':', '/'), false);
is('but after a space', mayBreak(' ', 'w'), true);

// eb5538c draws a 94-character address on two lines; breaking after its
// colon would draw twenty.
const address = [...'https://docs.google.com/spreadsheets/d/1a2b3c4d5e6f7g8h9i0j'];
is('a run with nowhere to break is cut at the edge',
  countLines(address, address.map(() => 8), 80),
  Math.ceil(address.length * 8 / 80));

// A sheet shaped the way the IR delivers one: column A in 18pt, the rest in
// 11pt, and a merge down column A across rows 52 to 56.
const heights = {
  'ＭＳ Ｐゴシック|44': 18, 'ＭＳ Ｐゴシック|72': 28, '游ゴシック|44': 25, '游ゴシック|32': 17,
};
const heightOf = (face, size) => heights[`${face}|${Math.round(size * 4)}`];
const advance = () => em;
const sheet = {
  name: 'S',
  col_count: 4,
  col_fonts: [[0, 0, 'ＭＳ Ｐゴシック', 18], [1, 16383, 'ＭＳ Ｐゴシック', 11]],
  normal_font: ['ＭＳ Ｐゴシック', 11],
  default_row_height: 21,
  default_row_custom: false,
  merge_cells: [{ start_row: 52, end_row: 56, start_col: 0, end_col: 0 }],
  rows: [],
};
const columns = [0, 100, 200, 300, 400];

is('a row the file does not record takes the tallest column font',
  sheetDefault(sheet, heightOf), 28);

// The quantiser: 0.05pt of grace, then floored to the pixel.
is('a pinned 14.95pt row draws twenty pixels',
  explainRow({
    sheet, index: 1, columns, heightOf, advance,
    row: { index: 1, height: 14.95, custom_height: true, cells: [] },
  }).px, 20);
is('and 14.93 draws nineteen',
  explainRow({
    sheet, index: 1, columns, heightOf, advance,
    row: { index: 1, height: 14.93, custom_height: true, cells: [] },
  }).px, 19);

// 00876's shape: the 18pt column is swallowed by a merge across these rows,
// so the row draws at the height of its 11pt cells.
is('a merge across rows drops the tall column',
  explainRow({
    sheet, index: 56, columns, heightOf, advance,
    row: {
      index: 56, height: 13.2, custom_height: false,
      cells: [
        { col: 0, style: { font_name: 'ＭＳ Ｐゴシック', font_size: 18, wrap_text: true }, value: null },
        { col: 1, style: { font_name: 'ＭＳ Ｐゴシック', font_size: 11 }, value: { String: '1.0' } },
      ],
    },
  }).px, 18);

// Away from that merge, column A is on show and lends its 18pt line; the
// thick rule adds its pixel on top.
is('a thick bottom adds a pixel to a height Excel works out',
  explainRow({
    sheet, index: 5, columns, heightOf, advance,
    row: {
      index: 5, custom_height: false, thick_bottom: true,
      cells: [{ col: 1, style: { font_name: '游ゴシック', font_size: 11 }, value: { String: 'x' } }],
    },
  }).px, 29);

// 119a4's shape: a bare cell in the tall column is the Normal format, and
// the column no longer lends its own font.
is('a recorded column is spoken for by its cell',
  explainRow({
    sheet, index: 5, columns, heightOf, advance,
    row: {
      index: 5, custom_height: false,
      cells: [
        { col: 0, style: { font_name: 'ＭＳ Ｐゴシック', font_size: 11 }, value: null },
        { col: 1, style: { font_name: 'ＭＳ Ｐゴシック', font_size: 11 }, value: { String: 'x' } },
      ],
    },
  }).px, 18);

console.log(failures ? `\n${failures} failed` : '\nall good');
process.exit(failures ? 1 : 0);
