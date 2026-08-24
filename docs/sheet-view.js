/// Draw a worksheet onto a canvas, by the rules Excel was measured for.
///
/// The rules themselves are not invented here. Every one of them — the width a
/// stored column really takes, the room kept either side of a cell's text, the
/// height Excel gives a row, where a line is allowed to break, what a cell
/// clips itself to — was measured against Excel and is carried by
/// `row-model.js` and `row-heights.json` beside this file. This module only
/// puts them on a canvas, so the viewer page and the in-repo measuring page
/// draw from one copy rather than two that can drift apart.
///
/// `drawSheet` returns the canvas; the caller decides where to put it.

import { format_cell_number } from './oxidocs_wasm.js';
import { explainRow, cellText, mayBreak, GUTTER } from './row-model.js';

/// Excel's own default column width for a sheet that states none.
const DEFAULT_CHARACTERS = 8.43;
/// A pixel of a 96-dpi screen, in EMU.
const EMU = 9525;

const advances = new Map();

/// One character's advance. Excel measures each on its own rather than as
/// part of a run, so a line is the sum of single characters here too.
function advance(measurer, font, letter) {
  const key = `${font}|${letter}`;
  let held = advances.get(key);
  if (held === undefined) {
    measurer.font = font;
    held = Math.round(measurer.measureText(letter).width);
    advances.set(key, held);
  }
  return held;
}

/// The face a cell wears, written the way a canvas wants it.
function fontOf(style, size) {
  const points = size ?? style?.font_size ?? 11;
  const face = style?.font_name || 'Calibri';
  return `${style?.italic ? 'italic ' : ''}${style?.bold ? '700 ' : ''}` +
    `${Math.round(points * 96 / 72)}px "${face}", "Yu Gothic", sans-serif`;
}

/// What a cell shows: a number wears the format the cell states, which is the
/// same reading the rest of the engine uses rather than the browser's own.
function shownText(cell) {
  const value = cell?.value;
  const format = cell?.style?.number_format;
  if (value && typeof value === 'object' && 'Number' in value && format) {
    try {
      return format_cell_number(value.Number, format);
    } catch {
      return cellText(cell);
    }
  }
  return cellText(cell);
}

function widthOf(measurer, font, text) {
  let total = 0;
  for (const letter of text) total += advance(measurer, font, letter);
  return total;
}

function measureDigit(name, points) {
  const canvas = measureDigit.canvas ||= document.createElement('canvas');
  const ink = canvas.getContext('2d');
  ink.font = `${(points || 11) * 96 / 72}px ${name ? `"${name}"` : 'Calibri'}, Calibri, sans-serif`;
  return Math.round(ink.measureText('0').width) || 7;
}

/// A stored column width already carries the room either side of the text.
function columnPixels(width, digitWidth) {
  const padding = Math.trunc(128 / digitWidth);
  return Math.trunc(((256 * width + padding) / 256) * digitWidth);
}

function columnEdges(sheet, digitWidth) {
  const edges = [0];
  const hidden = new Set(sheet.hidden_cols || []);
  for (let column = 0; column < Math.max(sheet.col_count || 0, 1); column += 1) {
    const stated = (sheet.col_widths || [])[column];
    const width = hidden.has(column) ? 0
      : stated > 0 ? columnPixels(stated, digitWidth)
        : sheet.default_col_width > 0 ? columnPixels(sheet.default_col_width, digitWidth)
          : Math.trunc(DEFAULT_CHARACTERS * digitWidth + GUTTER);
    edges.push(edges[edges.length - 1] + width);
  }
  return edges;
}

/// The room Excel keeps either side of a cell's text: the cell font's own
/// digit, in bands. Measured on Excel — see the gutters note in the renderer.
function gutters(measurer, font) {
  measurer.font = font;
  const digit = Math.round(measurer.measureText('0').width) || 7;
  if (digit <= 8) return [3, 2];
  if (digit <= 12) return [4, 3];
  if (digit <= 16) return [5, 4];
  return [6, 5];
}

/// One level of indent: three spaces of the workbook's first font, which is
/// not the font the Normal style points at in every workbook.
function indentStep(measurer, sheet) {
  const [face, size] = sheet.first_font || sheet.normal_font || ['Calibri', 11];
  measurer.font = `${Math.round((size || 11) * 96 / 72)}px "${face}", sans-serif`;
  return 3 * Math.round(measurer.measureText(' ').width || 4);
}

/// Where a line breaks in the room it is given, one advance at a time, held
/// back from any mark that may not start a line.
function breakLines(measurer, font, text, room) {
  const held = [];
  for (const paragraph of String(text).split('\n')) {
    const letters = [...paragraph];
    if (!room || !letters.length) { held.push(paragraph); continue; }
    let at = 0;
    while (at < letters.length) {
      let take = 0;
      let run = 0;
      while (at + take < letters.length) {
        const next = run + advance(measurer, font, letters[at + take]);
        if (take > 0 && next > room) break;
        run = next;
        take += 1;
      }
      while (take > 1 && at + take < letters.length &&
             !mayBreak(letters[at + take - 1], letters[at + take])) take -= 1;
      held.push(letters.slice(at, at + take).join(''));
      at += take;
    }
  }
  return held;
}

function colour(hex, fallback) {
  if (!hex) return fallback;
  const held = String(hex).replace('#', '');
  return `#${held.length === 8 ? held.slice(2) : held}`;
}

/// How wide a rule of this kind is drawn, and whether it is broken.
function ruleOf(border) {
  if (!border) return null;
  const kind = border.style || 'thin';
  if (kind === 'none') return null;
  const width = kind === 'thick' ? 3 : kind === 'medium' || kind === 'mediumDashed' ? 2 : 1;
  const dash = kind.includes('dot') || kind === 'hair' ? [1, 1]
    : kind.toLowerCase().includes('dash') ? [3, 1] : [];
  return { width, dash, colour: colour(border.color, '#000000'), double: kind === 'double' };
}

function mergeMap(sheet) {
  const covered = new Map();
  for (const held of sheet.merge_cells || []) {
    for (let row = held.start_row; row <= held.end_row; row += 1) {
      for (let col = held.start_col; col <= held.end_col; col += 1) {
        covered.set(`${row},${col}`, {
          anchor: row === held.start_row && col === held.start_col,
          span: held,
        });
      }
    }
  }
  return covered;
}

/// Where a drawing hangs, from the cells its anchors name.
function anchorAt(anchor, columns, tops, first) {
  const left = (columns[anchor.col] ?? columns[columns.length - 1] ?? 0) +
    anchor.col_off / EMU;
  const top = (tops[anchor.row + 1 - first] ?? tops[tops.length - 1] ?? 0) +
    anchor.row_off / EMU;
  return [left, top];
}

/// What a shape says, inside the box it says it in.
function saysShape(ink, said, box, measurer) {
  const inset = held => (held ?? 0) / EMU;
  const area = {
    left: box.left + inset(said.insets?.[0] ?? 91440),
    top: box.top + inset(said.insets?.[1] ?? 45720),
    right: box.right - inset(said.insets?.[2] ?? 91440),
    bottom: box.bottom - inset(said.insets?.[3] ?? 45720),
  };
  const room = area.right - area.left;
  if (room <= 0) return;

  // Every line, with the paragraph it belongs to: an empty paragraph spends a
  // line of its own, whether it sits first, last or between two blocks.
  const lines = [];
  for (const paragraph of said.paragraphs || []) {
    const points = paragraph.size || 18;
    const font = `${paragraph.italic ? 'italic ' : ''}${paragraph.bold ? '700 ' : ''}` +
      `${Math.round(points * 96 / 72)}px "${paragraph.face || 'ＭＳ Ｐゴシック'}", sans-serif`;
    const pitch = (paragraph.line_pitch ? paragraph.line_pitch * 96 / 72
      : Math.round(points * 96 / 72) * 1.3) * (paragraph.line_scale || 1);
    for (const held of breakLines(measurer, font, paragraph.text || '',
      said.wrap === false ? 0 : room)) {
      lines.push({ text: held, font, pitch, paragraph });
    }
  }
  if (!lines.length) return;

  const block = lines.reduce((sum, held) => sum + held.pitch, 0);
  const slack = (area.bottom - area.top) - block;
  const anchor = said.anchor || 't';
  let at = area.top + (anchor === 'ctr' ? Math.floor(slack / 2)
    : anchor === 'b' ? slack : 0);

  ink.save();
  ink.beginPath();
  ink.rect(box.left, box.top, box.right - box.left, box.bottom - box.top);
  ink.clip();
  for (const held of lines) {
    ink.font = held.font;
    ink.fillStyle = colour(held.paragraph.color, '#000000');
    const wide = widthOf(measurer, held.font, held.text);
    const align = held.paragraph.align;
    const left = align === 'ctr' ? area.left + (room - wide) / 2
      : align === 'r' ? area.right - wide : area.left;
    ink.fillText(held.text, left, at + held.pitch * .78);
    at += held.pitch;
  }
  ink.restore();
}

function drawShapes(ink, sheet, columns, tops, first, measurer) {
  for (const drawn of sheet.drawings || []) {
    const [left, top] = anchorAt(drawn.from, columns, tops, first);
    let right, bottom;
    if (drawn.to) {
      [right, bottom] = anchorAt(drawn.to, columns, tops, first);
    } else if (drawn.extent) {
      right = left + drawn.extent[0] / EMU;
      bottom = top + drawn.extent[1] / EMU;
    } else {
      continue;
    }
    const kind = drawn.kind;
    const shape = kind?.Shape;
    if (!shape) {
      // A picture's bytes are left out of the serialised IR, and a chart is
      // drawn from a part of its own: both are shown as the room they take.
      ink.strokeStyle = '#C9D1D9';
      ink.setLineDash([4, 3]);
      ink.strokeRect(left + .5, top + .5, right - left, bottom - top);
      ink.setLineDash([]);
      ink.fillStyle = '#8C959F';
      ink.font = '11px "Segoe UI", sans-serif';
      ink.fillText(kind?.Chart ? 'chart' : kind === 'Other' ? 'group' : 'picture',
        left + 6, top + 16);
      continue;
    }
    const round = (shape.geometry || '').toLowerCase().includes('round');
    const oval = /ellipse|oval|flowchartconnector/.test((shape.geometry || '').toLowerCase());
    const line = /^(line|straightconnector|bentconnector)/i.test(shape.geometry || '');
    ink.beginPath();
    if (line) {
      ink.moveTo(left, top);
      ink.lineTo(right, bottom);
    } else if (oval) {
      ink.ellipse((left + right) / 2, (top + bottom) / 2,
        Math.abs(right - left) / 2, Math.abs(bottom - top) / 2, 0, 0, Math.PI * 2);
    } else if (round) {
      const radius = Math.min(16, Math.abs(right - left) / 4, Math.abs(bottom - top) / 4);
      ink.roundRect(left, top, right - left, bottom - top, radius);
    } else {
      ink.rect(left, top, right - left, bottom - top);
    }
    if (shape.fill && !line) {
      ink.fillStyle = colour(shape.fill, '#FFFFFF');
      ink.fill();
    }
    if (shape.line) {
      ink.strokeStyle = colour(shape.line.color, '#000000');
      ink.lineWidth = Math.max(1, Math.round((shape.line.width || 9525) / EMU));
      ink.setLineDash(shape.line.dash ? [4, 3] : []);
      ink.stroke();
      ink.setLineDash([]);
    }
    if (shape.text) saysShape(ink, shape.text, { left, top, right, bottom }, measurer);
  }
}

/// Draw one sheet of a parsed workbook.
///
/// `heights` is the table shipped in `row-heights.json`: the height Excel
/// gives a row of a given face and size, measured from Excel itself. Without
/// it the row model falls back to the device's own metrics, which is close
/// but not the same.
///
/// Returns `{ canvas, width, height, rows, columns }`.
export function drawSheet({ book, index = 0, heights = {}, grid = false, shapes = true }) {
  const sheet = book.sheets[index];
  const digitWidth = measureDigit(book.default_style?.font_name, book.default_style?.font_size);
  const heightOf = (face, size) =>
    (!face || !size) ? undefined : heights[`${face}|${Math.round(size * 4)}`];
  const columns = columnEdges(sheet, digitWidth);
  const rows = new Map((sheet.rows || []).map(row => [row.index, row]));
  const indices = (sheet.rows || []).map(row => row.index);
  const first = indices.length ? Math.min(...indices) : 1;
  const last = indices.length ? Math.max(...indices) : 1;

  const canvas = document.createElement('canvas');
  const ink = canvas.getContext('2d');
  const measurer = document.createElement('canvas').getContext('2d');
  const step = (face, points, bold, italic, letter) =>
    advance(measurer, `${italic ? 'italic ' : ''}${bold ? '700 ' : ''}` +
      `${Math.round(points * 96 / 72)}px "${face}", sans-serif`, letter);

  // Every row's height, worked out the way Excel works it out.
  const tops = [0];
  const heightsOfRows = [];
  for (let at = first; at <= last; at += 1) {
    const row = rows.get(at);
    const told = explainRow({ sheet, index: at, row, columns, heightOf, advance: step });
    const px = row?.hidden ? 0 : told.px;
    heightsOfRows.push(px);
    tops.push(tops[tops.length - 1] + px);
  }

  let width = columns[columns.length - 1] || 1;
  let height = tops[tops.length - 1] || 1;
  if (shapes) {
    for (const drawn of sheet.drawings || []) {
      const [left, top] = anchorAt(drawn.from, columns, tops, first);
      const right = drawn.to ? anchorAt(drawn.to, columns, tops, first)[0]
        : left + (drawn.extent?.[0] ?? 0) / EMU;
      const bottom = drawn.to ? anchorAt(drawn.to, columns, tops, first)[1]
        : top + (drawn.extent?.[1] ?? 0) / EMU;
      width = Math.max(width, Math.ceil(right) + 1);
      height = Math.max(height, Math.ceil(bottom) + 1);
    }
  }
  const ratio = window.devicePixelRatio || 1;
  canvas.width = Math.max(1, Math.round(width * ratio));
  canvas.height = Math.max(1, Math.round(height * ratio));
  canvas.style.width = `${width}px`;
  canvas.style.height = `${height}px`;
  ink.scale(ratio, ratio);
  ink.fillStyle = '#FFFFFF';
  ink.fillRect(0, 0, width, height);
  ink.textBaseline = 'alphabetic';

  const covered = mergeMap(sheet);
  const boxOf = (at, cell) => {
    const held = covered.get(`${at},${cell.col}`);
    const span = held?.span;
    const lastRow = span ? span.end_row : at;
    const lastCol = span ? span.end_col : cell.col;
    return {
      left: columns[cell.col] ?? 0,
      right: columns[Math.min(lastCol + 1, columns.length - 1)] ?? 0,
      top: tops[at - first] ?? 0,
      bottom: tops[Math.min(lastRow - first + 1, tops.length - 1)] ?? 0,
    };
  };

  // The grounds, all of them, before any of the words: a letter that hangs
  // over its neighbour must not be wiped by the next cell's fill.
  for (const row of sheet.rows || []) {
    for (const cell of row.cells || []) {
      const shade = cell.style?.bg_color;
      if (!shade) continue;
      const box = boxOf(row.index, cell);
      ink.fillStyle = colour(shade, '#FFFFFF');
      ink.fillRect(box.left, box.top, box.right - box.left, box.bottom - box.top);
    }
  }

  // The grid Excel shows behind a sheet, which is not part of the sheet.
  if (grid) {
    ink.strokeStyle = '#D4D4D4';
    ink.lineWidth = 1;
    ink.beginPath();
    for (const edge of columns) {
      ink.moveTo(edge + .5, 0);
      ink.lineTo(edge + .5, height);
    }
    for (const edge of tops) {
      ink.moveTo(0, edge + .5);
      ink.lineTo(width, edge + .5);
    }
    ink.stroke();
  }

  // The rules a cell states, over the grid and under the words.
  for (const row of sheet.rows || []) {
    for (const cell of row.cells || []) {
      const box = boxOf(row.index, cell);
      const edges = [
        ['border_top', box.left, box.top, box.right, box.top],
        ['border_bottom', box.left, box.bottom, box.right, box.bottom],
        ['border_left', box.left, box.top, box.left, box.bottom],
        ['border_right', box.right, box.top, box.right, box.bottom],
      ];
      for (const [name, x1, y1, x2, y2] of edges) {
        const rule = ruleOf(cell.style?.[name]);
        if (!rule) continue;
        ink.strokeStyle = rule.colour;
        ink.lineWidth = rule.width;
        ink.setLineDash(rule.dash);
        const nudge = rule.width % 2 ? .5 : 0;
        ink.beginPath();
        ink.moveTo(x1 + (x1 === x2 ? nudge : 0), y1 + (y1 === y2 ? nudge : 0));
        ink.lineTo(x2 + (x1 === x2 ? nudge : 0), y2 + (y1 === y2 ? nudge : 0));
        ink.stroke();
        if (rule.double) {
          ink.beginPath();
          const away = 2;
          ink.moveTo(x1 + (x1 === x2 ? nudge + away : 0), y1 + (y1 === y2 ? nudge + away : 0));
          ink.lineTo(x2 + (x1 === x2 ? nudge + away : 0), y2 + (y1 === y2 ? nudge + away : 0));
          ink.stroke();
        }
        ink.setLineDash([]);
      }
      // A cell a form strikes out corner to corner.
      const across = ruleOf(cell.style?.border_diagonal);
      if (across && (cell.style?.diagonal_down || cell.style?.diagonal_up)) {
        ink.strokeStyle = across.colour;
        ink.lineWidth = across.width;
        ink.beginPath();
        if (cell.style.diagonal_down) {
          ink.moveTo(box.left, box.top);
          ink.lineTo(box.right, box.bottom);
        }
        if (cell.style.diagonal_up) {
          ink.moveTo(box.left, box.bottom);
          ink.lineTo(box.right, box.top);
        }
        ink.stroke();
      }
    }
  }

  // The words.
  const indent = indentStep(measurer, sheet);
  for (const row of sheet.rows || []) {
    if (row.hidden) continue;
    for (const cell of row.cells || []) {
      const held = covered.get(`${row.index},${cell.col}`);
      if (held && !held.anchor) continue;
      const text = shownText(cell);
      if (!text) continue;
      const style = cell.style || {};
      const font = fontOf(style);
      const box = boxOf(row.index, cell);
      const [before, after] = gutters(measurer, font);
      const pushed = indent * (style.indent || 0);
      const numeric = typeof cell.value === 'object' && cell.value && 'Number' in cell.value;
      const placed = style.horizontal_align
        || (numeric ? 'right' : style.stacked_text ? 'center' : 'left');

      const left = box.left + before + (placed === 'left' ? pushed : 0);
      const right = box.right - after - (placed === 'right' ? pushed : 0);
      const room = right - left;
      // A cell that does not wrap takes no notice of the breaks inside it:
      // Excel runs the pieces together on one line (`_xlsx_cell_break.py`).
      const lines = style.wrap_text
        ? breakLines(measurer, font, text, room)
        : [text.split('\n').join('')];

      ink.font = font;
      ink.fillStyle = colour(style.font_color, '#000000');
      const size = Math.round((style.font_size ?? 11) * 96 / 72);
      const line = Math.round(size * 1.28);
      const block = line * lines.length;
      const slack = (box.bottom - box.top) - block;
      const place = style.vertical_align || 'bottom';
      const down = box.top + (place === 'top' ? 0
        : place === 'center' || place === 'centre' ? Math.floor(slack / 2) : slack);

      // A cell's text is clipped to what it may cover: a wrapped or merged
      // one keeps to itself, a plain one runs on over an empty neighbour.
      ink.save();
      const runsOn = !style.wrap_text && !held && !numeric;
      ink.beginPath();
      ink.rect(runsOn ? 0 : box.left, box.top,
        runsOn ? width : box.right - box.left, box.bottom - box.top);
      ink.clip();

      if (style.stacked_text) {
        // A stacked cell sets its characters one under the next, and a letter
        // or a digit stands upright rather than lying on its side.
        let at = down + size;
        for (const letter of text) {
          const wide = widthOf(measurer, font, letter);
          ink.fillText(letter, box.left + (box.right - box.left - wide) / 2, at);
          at += line;
        }
      } else {
        lines.forEach((piece, at) => {
          const wide = widthOf(measurer, font, piece);
          const x = placed === 'right' ? right - wide
            : placed === 'center' || placed === 'centre' || placed === 'centerContinuous'
              ? left + (room - wide) / 2
              : left;
          ink.fillText(piece, x, down + line * at + Math.round(size * .92));
        });
      }
      ink.restore();
    }
  }

  if (shapes) drawShapes(ink, sheet, columns, tops, first, measurer);

  return { canvas, width, height, rows: heightsOfRows, columns };
}
