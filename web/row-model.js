// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The height Excel gives a row, worked out from what the file says.
//!
//! The same rules `tools/oxi-xlsx-renderer` draws with, derived from Excel by
//! measurement (see the SX15–SX19 notes): a row that pins its height keeps the
//! stated number, any other stated height is a cache Excel recomputes, and the
//! height it recomputes is the tallest of what the row's cells ask for and what
//! the blank of the row asks for, plus a pixel for each thick rule.
//!
//! Nothing here touches the DOM. Text is measured through the `advance`
//! function the caller passes in, so the browser can hand it a canvas and a
//! test can hand it a table of known widths.

/// Excel will not draw a row taller than 409.5pt, however much it holds.
export const CEILING = 546;
/// The gutter Excel keeps either side of a wrapped line, in pixels.
export const GUTTER = 5;

/// Characters that may not start a line: the closing half of a pair, the
/// small kana, the sound mark, and the punctuation that ends a phrase.
const NEVER_STARTS = '、。，．・：；？！゛゜ゝゞヽヾー々〆ぁぃぅぇぉっゃゅょゎ' +
  'ァィゥェォッャュョヮ）］｝〉》」』】〕〙〗”’';
/// Characters that may not end one: the opening half of a pair.
const NEVER_ENDS = '（［｛〈《「『【〔〘〖“‘￥＄';

/// Text written on the body of the em, which breaks between any two
/// characters the kinsoku rules allow.
export function ideographic(letter) {
  const c = letter.codePointAt(0);
  return (c >= 0x1100 && c <= 0x115F) || (c >= 0x2E80 && c <= 0x303E) ||
    (c >= 0x3041 && c <= 0x33FF) || (c >= 0x3400 && c <= 0x4DBF) ||
    (c >= 0x4E00 && c <= 0x9FFF) || (c >= 0xA000 && c <= 0xA4CF) ||
    (c >= 0xAC00 && c <= 0xD7A3) || (c >= 0xF900 && c <= 0xFAFF) ||
    (c >= 0xFE30 && c <= 0xFE6F) || (c >= 0xFF01 && c <= 0xFF60) ||
    (c >= 0xFFE0 && c <= 0xFFE6);
}

/// Where Excel is willing to end a line. Measured from its own PDF: a line
/// that would start with a forbidden character does not hang it past the
/// edge — the break moves back a character at a time until it is allowed.
export function mayBreak(before, after) {
  if (before === ' ' || before === '　' || before === '\t') return true;
  if (NEVER_STARTS.includes(after) || NEVER_ENDS.includes(before)) return false;
  // Anything that is not written on the em travels as one run to the next
  // space — a Latin word, and a web address with it.
  return ideographic(before) || ideographic(after);
}

/// How many lines a paragraph takes in a box this wide, given what each of
/// its characters advances.
export function countLines(letters, advances, width) {
  if (!letters.length) return 1;
  const box = Math.max(1, width);
  let lines = 0;
  let start = 0;
  while (start < letters.length) {
    let take = 0;
    let run = 0;
    while (start + take < letters.length) {
      const next = run + advances[start + take];
      if (take > 0 && next > box) break;
      run = next;
      take += 1;
    }
    // Give characters back until the break is one Excel would make. A run
    // with nowhere to break — a long web address — is cut where it stops
    // fitting rather than left to fill a line a character at a time.
    const fill = take;
    while (start + take < letters.length && take > 1 &&
           !mayBreak(letters[start + take - 1], letters[start + take])) {
      take -= 1;
    }
    if (take <= 1 && fill > 1) take = fill;
    lines += 1;
    start += Math.max(take, 1);
  }
  return Math.max(lines, 1);
}

/// The text a cell shows, as the IR carries it.
export function cellText(cell) {
  const value = cell?.value;
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'object') {
    for (const key of ['String', 'Number', 'Boolean', 'Error']) {
      if (key in value) return String(value[key]);
    }
    return '';
  }
  return String(value);
}

function columnName(index) {
  let name = '';
  let n = index;
  do {
    name = String.fromCharCode(65 + (n % 26)) + name;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return name;
}

/// Which stretches of columns a merge swallows in this row.
function mergedColumns(sheet, index) {
  return (sheet.merge_cells || [])
    .filter(m => m.start_row <= index && index <= m.end_row)
    .map(m => [m.start_col, m.end_col]);
}

function mergeAt(sheet, index, column) {
  for (const m of sheet.merge_cells || []) {
    if (m.start_row <= index && index <= m.end_row &&
        m.start_col <= column && column <= m.end_col) {
      return {
        anchor: m.start_row === index && m.start_col === column,
        rows: m.end_row - m.start_row,
      };
    }
  }
  return null;
}

/// The height the blank of a row asks for: the tallest font worn by a column
/// that has room to show it. A column swallowed by a merge shows nothing, and
/// one the row records a cell in is spoken for by that cell.
export function blankHeight(sheet, heightOf, merged, recorded) {
  const free = (first, last) => {
    for (let column = first; column <= last; column += 1) {
      const swallowed = merged.some(([a, b]) => a <= column && column <= b);
      if (!swallowed && !recorded.has(column)) return true;
    }
    return false;
  };
  let tallest;
  let unknown = null;
  const raise = (face, size) => {
    const px = heightOf(face, size);
    if (px === undefined) { unknown ||= `${face} ${size}`; return; }
    tallest = Math.max(tallest ?? 0, px);
  };
  const dressed = [];
  for (const [first, last, face, size] of sheet.col_fonts || []) {
    dressed.push([first, last]);
    if (free(first, last)) raise(face, size);
  }
  dressed.sort((a, b) => a[0] - b[0]);
  let next = 0;
  const bare = [];
  for (const [first, last] of dressed) {
    if (first > next) bare.push([next, first - 1]);
    next = Math.max(next, last + 1);
  }
  if (next < 16384) bare.push([next, 16383]);
  if (sheet.normal_font && bare.some(([a, b]) => free(a, b))) {
    raise(sheet.normal_font[0], sheet.normal_font[1]);
  }
  return { px: tallest, unknown };
}

/// The height of a row the file records nothing for.
export function sheetDefault(sheet, heightOf) {
  if (sheet.default_row_custom && sheet.default_row_height > 0) {
    return Math.floor((sheet.default_row_height + 0.05) / 0.75);
  }
  const blank = blankHeight(sheet, heightOf, [], new Set());
  if (blank.px !== undefined) return blank.px;
  return Math.ceil((sheet.default_row_height || 18.75) / 0.75);
}

/// Everything that asked this row for height, what each got, and which won.
///
/// `heightOf(face, size)` gives a font's own row height in pixels, or
/// undefined for one that has never been measured. `advance(face, size, bold,
/// italic, letter)` gives one character's width. `columns` holds the left edge
/// of every column, so a wrapped cell knows the box it breaks inside.
export function explainRow({ sheet, index, row, columns, heightOf, advance }) {
  const parts = [];
  const fallback = () => row?.height
    ? {
      px: Math.min(Math.floor((row.height + 0.05) / 0.75), CEILING),
      note: 'a font this page has never measured, so the height the file remembers stands in',
    }
    : { px: sheetDefault(sheet, heightOf), note: 'a font this page has never measured' };

  if (!row) {
    const px = sheetDefault(sheet, heightOf);
    parts.push({
      kind: 'blank', won: true, px,
      what: "the sheet's own default — the file records nothing for this row",
    });
    return { px, parts };
  }

  if (row.custom_height && row.height) {
    const px = Math.min(Math.floor((row.height + 0.05) / 0.75), CEILING);
    parts.push({
      kind: 'pin', won: true, px,
      what: `pinned at ${row.height}pt (customHeight), floored to the pixel`,
    });
    return { px, parts };
  }

  const merged = mergedColumns(sheet, index);
  const recorded = new Set((row.cells || []).map(cell => cell.col));
  let base = 0;
  if (row.style_font) {
    const px = heightOf(row.style_font[0], row.style_font[1]);
    if (px === undefined) return { ...fallback(), parts };
    base = px;
    parts.push({
      kind: 'blank', px,
      what: `the row's own format wears ${row.style_font[0]} ${row.style_font[1]}`,
    });
  } else {
    const blank = blankHeight(sheet, heightOf, merged, recorded);
    if (blank.px === undefined) return { ...fallback(), parts };
    base = blank.px;
    parts.push({
      kind: 'blank', px: blank.px,
      what: 'the blank of the row — the tallest font on a column with room to show it',
    });
  }

  let raise = 0;
  for (const cell of row.cells || []) {
    const face = cell.style?.font_name || 'Calibri';
    const size = cell.style?.font_size || 11;
    const px = heightOf(face, size);
    if (px === undefined) return { ...fallback(), parts };
    const where = mergeAt(sheet, index, cell.col);
    const text = cellText(cell);
    let lines = 1;
    let what;
    if (where && (where.rows > 0 || !where.anchor)) {
      parts.push({
        kind: 'cell', px: 0, column: cell.col,
        what: `a merge across rows holds ${columnName(cell.col)} — it asks for nothing`,
      });
      continue;
    } else if (where) {
      what = `${columnName(cell.col)} is merged across its row: one line of ${face} ${size}`;
    } else if (!text) {
      what = `${columnName(cell.col)} is empty but dressed in ${face} ${size}: one line`;
    } else if (cell.style?.wrap_text) {
      const width = (columns[cell.col + 1] ?? 0) - (columns[cell.col] ?? 0);
      lines = 0;
      for (const paragraph of String(text).replace(/\r\n/g, '\n').split('\n')) {
        const letters = [...paragraph];
        const advances = letters.map(letter =>
          advance(face, size, !!cell.style?.bold, !!cell.style?.italic, letter));
        lines += countLines(letters, advances, width - GUTTER);
      }
      lines = Math.max(lines, 1);
      what = `${columnName(cell.col)} wraps ${lines} line${lines > 1 ? 's' : ''} of ${face} ${size}`;
    } else {
      what = `${columnName(cell.col)} holds one line of ${face} ${size}`;
    }
    const asked = lines * px;
    parts.push({ kind: 'cell', px: asked, column: cell.col, what });
    raise = Math.max(raise, asked);
  }

  const thick = (row.thick_top ? 1 : 0) + (row.thick_bottom ? 1 : 0);
  if (thick) {
    parts.push({
      kind: 'edge', px: thick,
      what: `a thick rule along the row's ${
        row.thick_top && row.thick_bottom ? 'top and bottom'
          : row.thick_top ? 'top' : 'bottom'} — a pixel of room each`,
    });
  }
  const measured = Math.max(base, raise);
  const px = Math.min((measured || sheetDefault(sheet, heightOf)) + thick, CEILING);
  for (const part of parts) {
    part.won = part.px === measured && part.kind !== 'edge';
  }
  return { px, parts };
}

export { columnName };
