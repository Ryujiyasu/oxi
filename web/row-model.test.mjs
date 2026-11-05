// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The row model, against the cases measured off Excel.
//!
//! Every number below came from a COM or PDF measurement of Excel itself, not
//! from reading this code back to itself. Run with `node web/row-model.test.mjs`.

import { countLines, mayBreak, explainRow, sheetDefault, rememberRowHeight } from './row-model.js';

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

const themedSheet = {
  normal_font: ['游ゴシック', 11], column_width_font: 'Aptos Narrow',
  default_row_height: 15, rows: [], merge_cells: [],
};
const asked = [];
const contextualHeight = (face, size, normal, bold, italic) => {
  if (normal) {
    asked.push([face, size, normal, bold, italic]);
    return 21;
  }
  return face === 'Aptos Narrow' ? 20 : face === 'Arial' ? 20 : 25;
};
const themedRow = { index: 1, cells: [{ col: 0, value: {String: 'M'},
  style: { font_name: 'Arial', font_size: 12 } }] };
const contextual = explainRow({ sheet: themedSheet, index: 1, row: themedRow,
  columns: [0, 100], heightOf: contextualHeight, advance: () => 13 });
is('a mixed-font row uses the themed Latin Normal font', contextual.px, 21);
is('the shared line calculation receives the effective Normal and style', asked,
  [['Arial', 12, ['Aptos Narrow', 11], false, false]]);
rememberRowHeight(themedSheet, themedRow, contextual);
is('automatic height is available to save and VBA', themedRow.height, 15.75);
rememberRowHeight(themedSheet, themedRow, {px: 20});
is('returning to the default removes the automatic height cache', themedRow.height, undefined);
themedRow.height = 30; themedRow.custom_height = true;
rememberRowHeight(themedSheet, themedRow, {px: 21});
is('a computed height leaves a manually chosen height alone', themedRow.height, 30);
themedRow.custom_height = false;
rememberRowHeight(themedSheet, themedRow, {px: 21, note: 'unmeasured font'});
is('an unmeasured font keeps its stored height', themedRow.height, 30);
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


// Raised and lowered characters share a row's ascent and descent.
const positionedSheet = {
  normal_font: ['Arial', 8], column_width_font: 'Arial',
  default_row_height: 11.25, rows: [], merge_cells: [],
};
const positionRequests = [];
const positionHeight = (face, size) => size === 8 ? 15 : 20;
const mixedPositionRow = {
  index: 1, cells: [{
    col: 0, value: { String: 'Mg' }, style: { font_name: 'Arial', font_size: 12 },
    runs: [
      { text: 'M', dressed: true, font: 'Arial', size: 12, vert_align: 'superscript' },
      { text: 'g', dressed: true, font: 'Arial', size: 12, vert_align: 'subscript' },
    ],
  }],
};
const mixedPosition = explainRow({
  sheet: positionedSheet, index: 1, row: mixedPositionRow, columns: [0, 100],
  heightOf: positionHeight, advance: () => 13,
  fontPositionRowHeightOf: (face, size, fonts) => {
    positionRequests.push([face, size, fonts]);
    return 27;
  },
});
is('a raised and lowered run share the measured Excel row', mixedPosition.px, 27);
is('the row calculation receives both effective run positions', positionRequests, [
  ['Arial', 8, [
    ['Arial', 12, false, false, false, 'superscript', 1],
    ['Arial', 12, false, false, false, 'subscript', 1],
  ]],
]);
rememberRowHeight(positionedSheet, mixedPositionRow, mixedPosition);
is('the mixed-position row retains the measured height for save', mixedPositionRow.height, 20.25);

const wholePositionRow = {
  index: 2, cells: [{
    col: 0, value: { String: 'M' },
    style: { font_name: 'Arial', font_size: 12, font_position: 'superscript' },
  }],
};
const wholeRequests = [];
const wholePosition = explainRow({
  sheet: positionedSheet, index: 2, row: wholePositionRow, columns: [0, 100],
  heightOf: positionHeight, advance: () => 13,
  fontPositionRowHeightOf: (face, size, fonts) => {
    wholeRequests.push(fonts);
    return 24;
  },
});
is('a whole-cell raised font uses the shared row calculation', wholePosition.px, 24);
is('a plain run inherits the whole-cell position', wholeRequests, [
  [['Arial', 12, false, false, false, 'superscript', 1]],
]);


const wrappedPositionSheet = {
  normal_font: ['Arial', 20], column_width_font: 'Arial',
  default_row_height: 25.5, rows: [], merge_cells: [],
};
const wrappedPositionRow = {
  index: 1, cells: [{
    col: 0, value: { String: 'M\nM' },
    style: { font_name: 'Arial', font_size: 12, font_position: 'superscript', wrap_text: true },
  }],
};
const wrappedRequests = [];
const wrappedPosition = explainRow({
  sheet: wrappedPositionSheet, index: 1, row: wrappedPositionRow, columns: [0, 100],
  heightOf: (face, size) => size === 20 ? 34 : 20, advance: () => 13,
  fontPositionRowHeightOf: (face, size, fonts) => {
    wrappedRequests.push(fonts);
    return 51;
  },
});
is('a wrapped raised font shares only its last baseline with Normal', wrappedPosition.px, 51);
is('the row calculation receives the actual line count', wrappedRequests, [
  [['Arial', 12, false, false, false, 'superscript', 2]],
]);
rememberRowHeight(wrappedPositionSheet, wrappedPositionRow, wrappedPosition);
is('the wrapped positioned height is retained for save', wrappedPositionRow.height, 38.25);

let richWrappedAsked = false;
const richWrappedRow = {
  index: 1, height: 36, cells: [{
    col: 0, value: { String: 'M\nM' },
    style: { font_name: 'Arial', font_size: 12, wrap_text: true },
    runs: [{ text: 'M\nM', dressed: true, font: 'Arial', size: 12, vert_align: 'superscript' }],
  }],
};
const richWrapped = explainRow({
  sheet: positionedSheet, index: 1, row: richWrappedRow, columns: [0, 100],
  heightOf: positionHeight, advance: () => 13,
  fontPositionRowHeightOf: () => { richWrappedAsked = true; return 20; },
});
is('an unmodeled wrapped rich position keeps its measured height', richWrapped.px, 48);
is('wrapped rich positions do not use a plain-text line calculation', richWrappedAsked, false);

console.log(failures ? `\n${failures} failed` : '\nall good');
process.exit(failures ? 1 : 0);
