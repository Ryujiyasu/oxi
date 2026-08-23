// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Oxi XLSX Renderer — draws a worksheet for pixel comparison against Excel.
//!
//! Excel's truth comes from `ExportAsFixedFormat` (a PDF), rasterised
//! alongside this. What is drawn here is the printed sheet: the used range,
//! cell by cell, with the text each cell shows under its number format.
//!
//! Usage:
//!   oxi-xlsx-renderer <input.xlsx> <output.png> [dpi]
//!
//! Geometry, measured from Excel:
//!   - A column's width is stated in characters; Excel renders the default
//!     8.38 characters as 54 pixels at 96 DPI.
//!   - A row's height is stated in points, 18.75 by default.

use oxicells_core::ir::{Cell, CellStyle, CellValue, Row, Sheet, Workbook};
use oxicells_core::parser::parse_xlsx;

mod graph;
mod row_defaults;

/// A stored column width already carries the gutter either side of a cell's
/// text, so it is not the width a person types into Excel: typing 10 stores
/// 10.625. Pixels come back out of it by OOXML's own rule, which a ruled
/// worksheet confirmed to the pixel.
fn column_pixels(width: f32, digit: f32) -> f32 {
    let padding = (128.0 / digit).trunc();
    (((256.0 * width + padding) / 256.0) * digit).trunc()
}
/// What a digit measures when the standard font could not be asked: Calibri 11,
/// which is what a workbook that names no font gets.
const FALLBACK_DIGIT_WIDTH: f32 = 7.0;
const DEFAULT_ROW_POINTS: f32 = 18.75;
/// Excel's own default column width for a sheet that states none: eight
/// digits of the standard font and eight pixels. Measured across digits of
/// 6, 7, 8, 9, 10 and 12 pixels — 56, 64, 72, 80, 88 and 104. The 8.43
/// characters Excel reports is that width read back through the padding, not
/// the number it starts from, and reading it forwards is a pixel out on a
/// wide digit.
const DEFAULT_DIGITS: f32 = 8.0;
const DEFAULT_PADDING: f32 = 8.0;

pub(crate) struct Geometry {
    /// Left edge of each column, and the sheet's full width.
    columns: Vec<f32>,
    /// Top edge of each row, and the sheet's full height.
    rows: Vec<f32>,
    /// Left edge of each column before the drawn range, as an offset back
    /// from its left edge — negative, and only needed by what hangs over the
    /// grid rather than sits in it.
    before_columns: Vec<f32>,
    /// The same for the rows above the range, indexed from row 1.
    before_rows: Vec<f32>,
    /// Left edge of each column past the drawn range, starting one column
    /// beyond the range's own right edge. A drawing whose far corner hangs
    /// out there has a real place, and clamping it to the edge of the picture
    /// would squeeze everything inside it — a chart cut off by the edge would
    /// be drawn whole in the room that is left instead of running past it.
    after_columns: Vec<f32>,
    /// The same for the rows below the range.
    after_rows: Vec<f32>,
    first_column: u32,
    first_row: u32,
}

/// The range to draw, when the caller states one: `OXI_XLSX_RANGE="2,1,140,100"`
/// is rows 2 to 140 and columns 1 to 100, counted the way a sheet counts them.
///
/// Excel's own `UsedRange` is part content and part cache — it leaves out a
/// row of empty cells in a font of their own in one workbook and keeps the
/// same thing in another — so a comparison against a picture of it has to be
/// told which rectangle that was, rather than working it out again and
/// differing for reasons that are nothing to do with drawing.
fn stated_extent() -> Option<(u32, u32, u32, u32)> {
    let held = std::env::var("OXI_XLSX_RANGE").ok()?;
    let numbers: Vec<u32> = held
        .split(',')
        .filter_map(|part| part.trim().parse().ok())
        .collect();
    match numbers[..] {
        [first_row, first_column, last_row, last_column] if first_column >= 1 => Some((
            first_row,
            first_column - 1,
            last_row,
            last_column.saturating_sub(1),
        )),
        _ => None,
    }
}

/// Has an empty cell been dressed in anything Excel would remember?
///
/// The range Excel hands over counts an empty cell that was given a face or
/// something to show, and passes over one that carries no more than an
/// alignment. `20211210_mousikomi` writes 35 such cells across seven columns
/// — the workbook's own font, no fill, no border, `vertical="center"` — and
/// Excel's range stops seven columns short of them; `h2daa202601s` does the
/// same over sixteen rows, its style even saying `applyFill="1"` over a fill
/// of `patternType="none"`. A cell wearing a face of its own is remembered,
/// though: `fies_t2`'s first column and `zuhyo`'s first two are empty cells in
/// a font that is not the workbook's, and Excel keeps every one of them.
fn dressed(style: &CellStyle, normal: Option<&(String, f32)>) -> bool {
    let named = style.font_name.as_deref();
    let sized = style.font_size;
    let own_face = match normal {
        Some((face, size)) => {
            named.is_some_and(|name| name != face) || sized.is_some_and(|points| points != *size)
        }
        None => named.is_some() || sized.is_some(),
    };
    own_face
        || style.bold
        || style.italic
        || style.underline
        || style.font_color.is_some()
        || style.bg_color.is_some()
        || style.number_format.is_some()
        || style.border_top.is_some()
        || style.border_bottom.is_some()
        || style.border_left.is_some()
        || style.border_right.is_some()
}

fn used_extent(sheet: &Sheet) -> (u32, u32, u32, u32) {
    if let Some(stated) = stated_extent() {
        return stated;
    }
    let mut first_row = u32::MAX;
    let mut last_row = 0;
    let mut first_column = u32::MAX;
    let mut last_column = 0;
    for row in &sheet.rows {
        // A row with nothing in it but a height of its own, or a format of
        // its own, is still part of the range Excel hands over: the 8.5pt
        // spacer at the top of a procurement list is a row like any other.
        if row.custom_height || row.style_font.is_some() {
            first_row = first_row.min(row.index);
            last_row = last_row.max(row.index);
        }
        for cell in &row.cells {
            // A cell counts as used when it holds something, or when it has
            // something to show: a ruled or filled but empty cell is part of
            // the range Excel hands over, and leaving it out shifts the whole
            // sheet against Excel's own picture. A font or a number format is
            // not something to show — data_A28's first row is a hundred empty
            // cells in ＭＳ 明朝 11 with a text format, and Excel's own range
            // starts at the row below it.
            let empty = matches!(cell.value, CellValue::Empty) && cell.formula.is_none();
            if empty && !dressed(&cell.style, sheet.normal_font.as_ref()) {
                continue;
            }
            first_row = first_row.min(row.index);
            last_row = last_row.max(row.index);
            first_column = first_column.min(cell.col);
            last_column = last_column.max(cell.col);
        }
    }
    // A sheet says how far it reaches, and that can be further than its last
    // filled cell. Excel hands over the declared range when asked for a
    // picture, so text running on past the last cell has somewhere to go.
    //
    // Only the far end of it, though. The near end is a cache like the row
    // heights are: data_A28 declares a range from row 1 while Excel's own
    // picture starts at row 2, and following the declaration puts an extra
    // row above the sheet and every row below it out of place.
    if let Some((start_row, start_column, end_row, end_column)) = sheet.declared_range {
        if first_row == u32::MAX {
            return (start_row, start_column, end_row, end_column);
        }
        return (
            first_row,
            first_column,
            last_row.max(end_row),
            last_column.max(end_column),
        );
    }
    if first_row == u32::MAX {
        return (1, 1, 0, 0);
    }
    (first_row, first_column, last_row, last_column)
}

/// The height the blank of a row asks for, in 96-dpi pixels: the tallest
/// font worn by a column that has room to show it. Every column is written
/// in its `<col>` style's font, or the workbook's Normal font where no
/// `<col>` dresses it, and a blank cell still stands one line of that font
/// tall. A column whose cell in this row is swallowed by a merge shows
/// nothing and so asks for nothing — which is why 00876's rows of 11pt
/// cells draw 18px even though its column A wears an 18pt font: A is inside
/// a merge that spans those rows.
///
/// A column the row records a cell in speaks for itself and is not counted
/// here, whatever its `<col>` says: 119a4's column D wears 14pt while its
/// rows hold a bare `<c r="D29"/>`, which is the Normal format, and Excel
/// draws those rows at Normal's height.
///
/// `None` means a font the measured table does not know, so the caller must
/// fall back rather than guess.
fn blank_row_px(
    sheet: &Sheet,
    merged_columns: &[(u32, u32)],
    recorded_columns: &[u32],
) -> Option<u16> {
    // Where neither a merge nor a cell of its own reaches, the sheet's
    // columns are what is on show.
    let free = |first: u32, last: u32| {
        (first..=last).any(|column| {
            !merged_columns
                .iter()
                .any(|(from, to)| *from <= column && column <= *to)
                && !recorded_columns.contains(&column)
        })
    };
    let mut tallest: Option<u16> = None;
    let mut raise = |face: &str, size: f32| -> Option<()> {
        let px = row_defaults::font_default_row_px(face, size)?;
        tallest = Some(tallest.unwrap_or(0).max(px));
        Some(())
    };
    let mut dressed: Vec<(u32, u32)> = Vec::new();
    for (first, last, face, size) in &sheet.col_fonts {
        dressed.push((*first, *last));
        if free(*first, *last) {
            raise(face, *size)?;
        }
    }
    // The columns no <col> covers wear Normal. Excel's sheets run to 16384
    // columns whether or not a file says so.
    dressed.sort_unstable();
    let mut next = 0u32;
    let mut bare = Vec::new();
    for (first, last) in dressed {
        if first > next {
            bare.push((next, first - 1));
        }
        next = next.max(last + 1);
    }
    if next < 16384 {
        bare.push((next, 16383));
    }
    if let Some((face, size)) = &sheet.normal_font {
        if bare.iter().any(|(first, last)| free(*first, *last)) {
            raise(face, *size)?;
        }
    }
    tallest
}

/// The height of a row the sheet says nothing about. A sheet that pins its
/// default (customHeight) gets the stated number, given 0.05pt of grace and
/// floored to the 96-dpi pixel: 17.18 draws 22px but 17.2 draws 23px. A sheet
/// that does not pin it has the number thrown away — Excel derives the height
/// from the fonts on show instead, and the tallest one wins. Both rules were
/// measured on Excel itself (2026-08-21); the stated number only *looks*
/// honoured in most files because the author's Excel derived it from the same
/// fonts this rule sees.
fn default_row_points(sheet: &Sheet) -> f32 {
    if sheet.default_row_custom && sheet.default_row_height > 0.0 {
        return ((sheet.default_row_height + 0.05) / 0.75).floor() * 0.75;
    }
    match blank_row_px(sheet, &[], &[]) {
        Some(px) => px as f32 * 0.75,
        None => fallback_row_points(sheet),
    }
}

/// The height of one row, in 96-dpi pixels. A row that pins its height
/// (customHeight) gets the stated number, +0.05pt of grace, floored to the
/// pixel — 14.93 draws 19px and 14.95 draws 20px. Any other row is measured
/// afresh: the stored ht is only a cache from the machine that wrote the
/// file, and Excel derives the height from the row's own content — the
/// tallest of `lines × the cell font's default height` across the cells,
/// floored at the sheet default. An empty cell still counts one line of its
/// font, which is how a tall-styled empty cell props a row open.
#[allow(clippy::too_many_arguments)]
fn row_pixels(
    held: Option<&Row>,
    sheet: &Sheet,
    merged_columns: &[(u32, u32)],
    default_px: f32,
    columns: &[f32],
    first_column: u32,
    scale: f32,
    counter: Option<&LineCounter>,
    merged: &std::collections::HashMap<(u32, u32), Merged>,
) -> f32 {
    // Excel will not draw a row taller than 409.5pt, however much its
    // content asks for — three corpus rows carrying 85 embedded lines all
    // sit at exactly these 546 pixels.
    const CEILING_PX: f32 = 546.0;
    let Some(row) = held else { return default_px };
    if row.custom_height {
        if let Some(ht) = row.height {
            return (((ht + 0.05) / 0.75).floor()).min(CEILING_PX);
        }
    }
    // What the row's blank asks for. A row with a format of its own
    // (customFormat) wears that font from end to end, whatever its columns
    // say; otherwise the columns are on show, minus the ones a merge
    // swallows in this row.
    let stated = |row: &Row| match row.height {
        Some(ht) => (((ht + 0.05) / 0.75).floor()).min(CEILING_PX),
        None => default_px,
    };
    let base = match &row.style_font {
        Some((face, size)) => match row_defaults::font_default_row_px(face, *size) {
            Some(px) => px as f32,
            None => return stated(row),
        },
        None => {
            let recorded: Vec<u32> = row.cells.iter().map(|cell| cell.col).collect();
            match blank_row_px(sheet, merged_columns, &recorded) {
                Some(px) => px as f32,
                None => return stated(row),
            }
        }
    };
    let mut raise: f32 = 0.0;
    for cell in &row.cells {
        let mut face = cell.style.font_name.as_deref().unwrap_or("Calibri");
        let mut size = cell.style.font_size.unwrap_or(11.0);
        // A cell whose text is dressed in stretches stands as tall as its
        // tallest stretch: 6dca80's B43 is 10pt with an 11pt run inside it,
        // and Excel gives the row the 11pt line.
        for run in &cell.runs {
            let run_face = run.font.as_deref().unwrap_or(face);
            let run_size = run.size.unwrap_or(size);
            let taller = match (
                row_defaults::font_default_row_px(run_face, run_size),
                row_defaults::font_default_row_px(face, size),
            ) {
                (Some(theirs), Some(ours)) => theirs > ours,
                _ => false,
            };
            if taller {
                face = run_face;
                size = run_size;
            }
        }
        let Some(font_px) = row_defaults::font_default_row_px(face, size) else {
            // A font the table has never measured: trust the cached height
            // the way the pre-derivation renderer did.
            return match row.height {
                Some(ht) => ((ht + 0.05) / 0.75).floor(),
                None => default_px,
            };
        };
        let text = cell_text(&cell.value, &cell.style).replace("\r\n", "\n");
        let text = text.as_str();
        if std::env::var("OXI_XLSX_DUMP_LINES").is_ok() {
            eprintln!(
                "row {} col {} font {} {} ({}px) wrap {} merge {} chars {}",
                row.index, cell.col, face, size, font_px, cell.style.wrap_text,
                match merged.get(&(row.index, cell.col)) {
                    Some(Merged::Anchor { rows, columns }) =>
                        format!("anchor {rows}x{columns}"),
                    Some(Merged::Covered) => "covered".to_string(),
                    None => "-".to_string(),
                },
                text.chars().count()
            );
        }
        match merged.get(&(row.index, cell.col)) {
            // A cell whose merge crosses rows holds no single row open:
            // an 18pt two-row title leaves both its rows at the default.
            Some(Merged::Anchor { rows, .. }) if *rows > 0 => continue,
            Some(Merged::Covered) => continue,
            // A one-row merge does not grow its row line by line (measured:
            // one and three wrapped lines sit in the same box); close to one
            // line of its font, and it never pulls the row down.
            Some(Merged::Anchor { .. }) => {
                raise = raise.max(font_px as f32);
                continue;
            }
            None => {}
        }
        // Stacked text spends a whole line of its font on every character,
        // and Excel gives the row two pixels beyond them: measured across
        // three faces and stacks of one to eight characters, a row of N
        // stands at N lines and two, unless the row's own font wants more.
        if cell.style.stacked_text && !text.is_empty() {
            let letters = text.chars().filter(|letter| *letter != '\n').count();
            raise = raise.max(letters as f32 * font_px as f32 + 2.0);
            continue;
        }
        // A raised or lowered run reaches past the line its font would
        // otherwise need, and Excel grows the row for it: `h2dee1989kre`
        // row 5 is 游ゴシック 11, whose line is 25 pixels, and Excel fits it
        // to 27 because one of its cells says `10³㎡`. Measured per face and
        // size by _xlsx_raised_extra.py.
        let font_px = if cell.runs.iter().any(|run| run.vert_align.is_some()) {
            row_defaults::raised_row_px(face, size).unwrap_or(font_px).max(font_px)
        } else {
            font_px
        };
        // A cell that does not wrap is one line however many breaks it
        // holds: Excel shows "あ\n\nあ" on a single line and leaves the row
        // at one line's height. Only a wrapping cell spends its newlines,
        // and it spends every one of them — a trailing break makes an
        // empty line that counts like any other.
        if text.is_empty() || !cell.style.wrap_text {
            raise = raise.max(font_px as f32);
            continue;
        }
        let lines = {
            let (from, to) = centred_across(row, cell, 0);
            let column = from.saturating_sub(first_column) as usize;
            let beyond = to.saturating_sub(first_column) as usize + 1;
            let width = match (columns.get(column), columns.get(beyond)) {
                (Some(left), Some(right)) => (right - left) / scale,
                _ => 0.0,
            };
            // A line is measured against the column minus the room Excel
            // keeps either side of it, which is the cell font's own — five
            // pixels for an eight-pixel digit, more for a bigger one.
            let (left_room, right_room) = gutters(face, size, cell.style.bold, cell.style.italic);
            let (before, after) = indent_room(&cell.style, indent_level(sheet));
            let width = width - left_room - right_room - before - after;
            counter
                .and_then(|counter| {
                    counter.lines(
                        face,
                        size,
                        cell.style.bold,
                        cell.style.italic,
                        text,
                        width,
                    )
                })
                .unwrap_or_else(|| text.matches('\n').count() as u32 + 1)
        };
        if std::env::var("OXI_XLSX_DUMP_LINES").is_ok() {
            let column = cell.col.saturating_sub(first_column) as usize;
            let box_px = match (columns.get(column), columns.get(column + 1)) {
                (Some(left), Some(right)) => (right - left) / scale,
                _ => 0.0,
            };
            eprintln!(
                "row {} col {} font {} {} wrap {} column {:.0}px chars {} lines {}",
                row.index, cell.col, face, size, cell.style.wrap_text, box_px,
                text.chars().count(), lines
            );
            if cell.style.wrap_text {
                for (index, para) in text.split('\n').enumerate() {
                    let count = counter.and_then(|counter| {
                        counter.lines(
                            face,
                            size,
                            cell.style.bold,
                            cell.style.italic,
                            para,
                            {
                                let (left, right) =
                                    gutters(face, size, cell.style.bold, cell.style.italic);
                                box_px - left - right
                            },
                        )
                    });
                    eprintln!(
                        "    para {} chars {} lines {:?} {:?}",
                        index,
                        para.chars().count(),
                        count,
                        para.chars().take(30).collect::<String>()
                    );
                }
            }
        }
        let contribution = (lines * font_px as u32) as f32;
        raise = raise.max(contribution);
    }
    // Every column swallowed by a merge and nothing written in the row: the
    // sheet's own default stands in.
    let measured = base.max(raise);
    if measured == 0.0 {
        return default_px;
    }
    // A thick rule along an edge is drawn in room the fitter keeps for it:
    // a pixel per edge, and only on a height Excel works out. A pinned row
    // measures the same with the rule as without.
    let thick = u32::from(row.thick_top) + u32::from(row.thick_bottom);
    // ★A row's stored height is a cache, and Excel recomputes it. Both ways of
    // trusting it were measured against the corpus and both are worse: taking
    // it whenever it is stated gives 0.9625 of the rows, taking it as a floor
    // under the computed height gives 0.9843, against 0.9995 for recomputing.
    // Excel shrinks a row below its stored height as readily as it grows one.
    (measured + thick as f32).min(CEILING_PX)
}

/// Characters that may not start a line: the closing half of a pair, the
/// small kana, the sound mark, and the punctuation that ends a phrase.
const NEVER_STARTS: &str = "、。，．・：；？！゛゜ゝゞヽヾー々〆\u{3005}\
    ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮ）］｝〉》」』】〕〙〗”’";
/// Characters that may not end one: the opening half of a pair.
const NEVER_ENDS: &str = "（［｛〈《「『【〔〘〖“‘￥＄";

/// Text written on the body of the em, which breaks between any two
/// characters the kinsoku rules allow.
fn ideographic(letter: char) -> bool {
    matches!(letter as u32,
        0x1100..=0x115F | 0x2E80..=0x303E | 0x3041..=0x33FF | 0x3400..=0x4DBF
        | 0x4E00..=0x9FFF | 0xA000..=0xA4CF | 0xAC00..=0xD7A3 | 0xF900..=0xFAFF
        | 0xFE30..=0xFE6F | 0xFF01..=0xFF60 | 0xFFE0..=0xFFE6
        | 0x20000..=0x2FA1F)
}

/// Where Excel is willing to end a line. Measured from its own PDF: a line
/// that would start with a forbidden character does not hang it past the
/// edge — the break moves back a character at a time until it is allowed, so
/// the line simply holds fewer glyphs.
fn may_break(before: char, after: char) -> bool {
    if before == ' ' || before == '\u{3000}' || before == '\t' {
        return true;
    }
    // A line may also end where a space begins — the space stays with the
    // line it ends and the next one starts past it, which is what Excel
    // draws: "The quick" then "brown fox", not "The" then "quick brown".
    if after == ' ' || after == '\u{3000}' || after == '\t' {
        return true;
    }
    if NEVER_STARTS.contains(after) || NEVER_ENDS.contains(before) {
        return false;
    }
    // A hyphen ends a line the way a space does, but only inside a word:
    // Excel breaks a long web address after one and leaves a minus alone.
    if before == '-' && !after.is_ascii_digit() {
        return true;
    }
    // Anything that is not written on the em travels as one run to the next
    // space — a Latin word, and a web address with it. Excel puts a 94
    // character URL on two lines rather than breaking it after its colon.
    ideographic(before) || ideographic(after)
}

/// The pieces a distributed cell is spread by: a run of plain Latin letters
/// travels whole, and everything else stands on its own.
///
/// Measured on `_xlsx_distributed.py`: `(1-2)` sits in the middle of its cell
/// as a single piece, brackets and hyphen and all, while `（①－②）` is pulled
/// apart into five and `あ、い` into three. So this is *not* the line-breaking
/// rule — Excel spreads a comma away from the character it may not follow —
/// but the plainer one of whether both characters are ASCII. A space breaks a
/// piece even between two of them: `A B` is spread to the cell's two edges.
/// Returns how many characters are in each piece.
pub(crate) fn distribution(line: &str) -> Vec<usize> {
    let clustered = |before: char, after: char| {
        before.is_ascii()
            && after.is_ascii()
            && !before.is_ascii_whitespace()
            && !after.is_ascii_whitespace()
    };
    let mut pieces = Vec::new();
    let mut held = 0usize;
    let mut before: Option<char> = None;
    for letter in line.chars() {
        if let Some(before) = before {
            if !clustered(before, letter) {
                pieces.push(held);
                held = 0;
            }
        }
        held += 1;
        before = Some(letter);
    }
    if held > 0 {
        pieces.push(held);
    }
    pieces
}

/// How far a run of text reaches, measured the way Excel measures it: each
/// character's own advance, added up. A text engine's run measurement is not
/// the same number — GDI compresses neighbouring Japanese punctuation and
/// DirectWrite runs a shade narrow — and the difference is what decides
/// whether the last letter of a spilling cell is drawn or clipped.
///
/// `None` when the platform cannot be asked, so the caller falls back.
pub(crate) fn run_width(
    face: &str,
    points: f32,
    bold: bool,
    italic: bool,
    text: &str,
) -> Option<f32> {
    advances(face, points, bold, italic, text)
        .map(|advances| advances.iter().sum::<i32>() as f32)
}

/// What each character of this text advances, in whole pixels.
///
/// Excel puts every character at the running total of these, and the total is
/// not what a text engine laying the same run out arrives at: DirectWrite
/// packs a Latin run one or two pixels tighter over a web address of fifty
/// characters. Drawing character by character at these positions is what
/// makes a long line of Latin text land where Excel's does.
pub(crate) fn advances(
    face: &str,
    points: f32,
    bold: bool,
    italic: bool,
    text: &str,
) -> Option<Vec<i32>> {
    held(|counter| {
        let letters: Vec<char> = text.chars().collect();
        counter.advances_of(face, points, bold, italic, &letters)
    })
}

/// Where each character of a shape's text lands, from the left of the line.
///
/// The advances are the font's own, scaled by the exact em rather than the
/// whole-pixel one the device draws at, and accumulated before they are
/// rounded — which is how Excel steps a shape's text.
#[cfg(windows)]
/// The same advances, before they are put on whole pixels.
///
/// A shape is *drawn* on whole pixels — a glyph lands where `shape_run` puts
/// it — but the break is decided on the fractions. `_xlsx_shape_room.py`
/// sweeps a box a quarter of a pixel at a time under two lines: ten
/// full-width kana, whose advances are 160.00 exactly, turn at a room of
/// 160.00, and `sanko_tool`'s own thirty-seven-character line turns at
/// **557.75** — which is that line's fractional sum, not the 558 its
/// rounded-per-glyph positions add up to. Excel breaks on the first, draws by
/// the second.
pub(crate) fn shape_widths(
    face: &str,
    points: f32,
    bold: bool,
    italic: bool,
    text: &str,
) -> Option<Vec<f32>> {
    held(|counter| {
        let letters: Vec<char> = text.chars().collect();
        let shares = counter.design_advances(face, bold, italic, &letters)?;
        let em = points * 96.0 / 72.0;
        Some(shares.into_iter().map(|share| share * em).collect())
    })
}

pub(crate) fn shape_run(
    face: &str,
    points: f32,
    bold: bool,
    italic: bool,
    text: &str,
) -> Option<Vec<i32>> {
    held(|counter| {
        let letters: Vec<char> = text.chars().collect();
        let shares = counter.design_advances(face, bold, italic, &letters)?;
        let em = points * 96.0 / 72.0;
        let mut at = 0.0f32;
        let mut held = Vec::with_capacity(shares.len());
        let mut was = 0;
        for share in shares {
            at += share * em;
            let now = at.round() as i32;
            held.push(now - was);
            was = now;
        }
        Some(held)
    })
}

/// The advances of a line whose runs are not all worn the same way.
///
/// Only weight varies inside a shape's paragraph in the corpus — never the
/// size or the face — so the shares are gathered run by run and then added up
/// ONCE. Measuring each run on its own and laying them end to end would start
/// the cumulative rounding again at every boundary, and a line of four bold
/// letters among Japanese text has four of those.
pub(crate) fn shape_run_worn(
    face: &str,
    points: f32,
    italic: bool,
    worn: &[(bool, String)],
) -> Option<Vec<i32>> {
    held(|counter| {
        let em = points * 96.0 / 72.0;
        let mut at = 0.0f32;
        let mut was = 0;
        let mut steps = Vec::new();
        for (bold, text) in worn {
            let letters: Vec<char> = text.chars().collect();
            let shares = counter.design_advances(face, *bold, italic, &letters)?;
            for (letter, share) in letters.iter().zip(shares) {
                at += share * em;
                let now = at.round() as i32;
                // One `dx` a UTF-16 unit, as `ExtTextOutW` wants them.
                for unit in 0..letter.len_utf16() {
                    steps.push(if unit == 0 { now - was } else { 0 });
                }
                was = now;
            }
        }
        Some(steps)
    })
}

/// One counter for the life of the program, holding its device context and
/// remembering every character it has measured. The renderer draws on one
/// thread, so it lives there.
fn held<T>(ask: impl FnOnce(&LineCounter) -> Option<T>) -> Option<T> {
    thread_local! {
        static HELD: Option<LineCounter> = LineCounter::new();
    }
    HELD.with(|counter| ask(counter.as_ref()?))
}

/// The size Excel draws a cell at when it is told to shrink the text to fit.
///
/// Not a scaling of the size the cell asks for: measured against Excel's own
/// picture across three faces and fifteen lengths, the em comes down a whole
/// pixel at a time and stops at the first size whose text fits the room. At
/// 96dpi a pixel of em is three quarters of a point, which is why the sizes
/// Excel settles on step 11, 10.25, 9.5, 8.75 and so on.
pub(crate) fn shrunk_to_fit(
    face: &str,
    points: f32,
    bold: bool,
    italic: bool,
    text: &str,
    room: f32,
) -> f32 {
    let fits = |points: f32| {
        run_width(face, points, bold, italic, text).is_some_and(|width| width <= room)
    };
    if room <= 0.0 || fits(points) {
        return points;
    }
    let natural = (points * 96.0 / 72.0).round() as i32;
    (1..natural)
        .rev()
        .map(|em| em as f32 * 72.0 / 96.0)
        .find(|smaller| fits(*smaller))
        .unwrap_or(points)
}

/// The columns a cell's text is laid out across: its own — and the ones a
/// merge takes with it — plus the run of neighbours that carry the same
/// `centerContinuous` alignment and hold nothing themselves.
///
/// This is how a heading is put over a group of columns without merging them,
/// and Excel lays the text out across the whole run: it centres it there, and
/// it *wraps* it there. Measuring the wrap against the cell's own column
/// alone makes a one-line heading two lines and the row twice as tall, which
/// is what h2dee1989kre's second row was.
pub(crate) fn centred_across(row: &Row, cell: &Cell, spans_columns: u32) -> (u32, u32) {
    if cell.style.horizontal_align.as_deref() != Some("centerContinuous") {
        return (cell.col, cell.col + spans_columns);
    }
    let joins = |column: u32| {
        row.cells.iter().any(|other| {
            other.col == column
                && other.style.horizontal_align.as_deref() == Some("centerContinuous")
                && matches!(other.value, CellValue::Empty)
        })
    };
    // Rightwards only. A heading is written in the leftmost cell of its
    // group, and reaching left as well would take in the cells that belong to
    // the heading before it: h2dee1989kre's 板ガラス and 安全ガラス then centre
    // over ranges that overlap, and neither lands where Excel puts it.
    let first = cell.col;
    let mut last = cell.col + spans_columns;
    while joins(last + 1) {
        last += 1;
    }
    (first, last)
}

/// The room Excel keeps either side of a cell's text, in pixels.
///
/// Not a constant. Measured by narrowing a column a pixel at a time until the
/// text takes a second line — across four faces, ten sizes and both weights,
/// 23 readings — the pair grows with the cell font's own digit: five pixels
/// together up to an eight-pixel digit, seven to twelve, nine to sixteen,
/// eleven beyond. The left keeps one more than the right, which is the three
/// and two that the small sizes show.
/// EMU to a pixel at 96 dpi: 914400 to the inch.
const EMU: f32 = 9525.0;

/// The characters a number format reserves room for without showing them.
///
/// `_x` asks for the width of x and draws nothing there. The text a cell
/// shows carries a space instead — that is what Excel's own `Range.Text`
/// gives — so a line measured from the text is short by the difference
/// between that space and the character it stands for. Excel's built-in
/// format 38, which every accounting column in the corpus wears, ends
/// `_)`: two pixels wider than the space in its place at ＭＳ 11, which is
/// exactly how far the corpus's right-aligned numbers sat from Excel's.
///
/// Returns what is reserved before the number and after it.
pub(crate) fn reserved_room(format: &str, negative: bool) -> (Vec<char>, Vec<char>) {
    let sections: Vec<&str> = format.split(';').collect();
    let section = match (negative, sections.len()) {
        (true, 2..) => sections[1],
        _ => sections[0],
    };
    let (mut before, mut after) = (Vec::new(), Vec::new());
    let mut seen_digit = false;
    let mut quoted = false;
    let mut characters = section.chars().peekable();
    while let Some(character) = characters.next() {
        match character {
            '"' => quoted = !quoted,
            _ if quoted => {}
            '0' | '#' | '?' | '.' | ',' => seen_digit = true,
            '\\' => {
                characters.next();
            }
            '_' => {
                if let Some(held) = characters.next() {
                    if seen_digit {
                        after.push(held);
                    } else {
                        before.push(held);
                    }
                }
            }
            '*' => {
                characters.next();
            }
            '[' => {
                for held in characters.by_ref() {
                    if held == ']' {
                        break;
                    }
                }
            }
            _ => {}
        }
    }
    (before, after)
}

/// How far apart Excel sets the lines of a shape's text.
///
/// Not the line box a cell uses: measured on `_xlsx_shape_text.py` over ten
/// faces and sixteen sizes, a shape's pitch is the **font's own** line height
/// — its unrounded `tmHeight`, scaled to the size — and a Japanese face gets
/// **1.3** of it where a Latin one gets all of it and no more. メイリオ 20pt
/// comes out at 40px of line height and 52px of pitch; Calibri 18pt at 29 and
/// 29. The pitch follows the face, not the letters: メイリオ set with `Hxpq`
/// spaces its lines exactly as it does with 国国国国.
#[cfg(windows)]
pub(crate) fn shape_line(face: &str, points: f32, bold: bool, italic: bool) -> Option<(f32, f32)> {
    use std::sync::Mutex;
    use windows::Win32::Graphics::Gdi::*;

    // The font's line height per em, and whether it is an East Asian face.
    static KNOWN: Mutex<Option<std::collections::HashMap<(String, bool, bool), (f32, bool)>>> =
        Mutex::new(None);
    const MEASURED_AT: i32 = 2048;

    let key = (face.to_string(), bold, italic);
    let mut held = KNOWN.lock().ok()?;
    let known = held.get_or_insert_with(std::collections::HashMap::new);
    let (per_em, japanese) = match known.get(&key) {
        Some(found) => *found,
        None => unsafe {
            let named: Vec<u16> = face.encode_utf16().chain(std::iter::once(0)).collect();
            let screen = GetDC(None);
            let font = CreateFontW(
                -MEASURED_AT,
                0,
                0,
                0,
                if bold { 700 } else { 400 },
                u32::from(italic),
                0,
                0,
                DEFAULT_CHARSET.0 as u32,
                OUT_DEFAULT_PRECIS.0 as u32,
                CLIP_DEFAULT_PRECIS.0 as u32,
                DEFAULT_QUALITY.0 as u32,
                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                windows::core::PCWSTR(named.as_ptr()),
            );
            let previous = SelectObject(screen, font);
            let mut metrics = TEXTMETRICW::default();
            let asked = GetTextMetricsW(screen, &mut metrics).as_bool();
            SelectObject(screen, previous);
            let _ = DeleteObject(font);
            ReleaseDC(None, screen);
            if !asked {
                return None;
            }
            let found = (
                metrics.tmHeight as f32 / MEASURED_AT as f32,
                metrics.tmCharSet == SHIFTJIS_CHARSET.0 as u8,
            );
            known.insert(key, found);
            found
        },
    };
    let em = points * 96.0 / 72.0;
    let natural = per_em * em;
    Some((natural * if japanese { 1.3 } else { 1.0 }, natural))
}

/// Where a drawing lands on the sheet, in the picture's own pixels.
///
/// A drawing hangs from a cell and an offset into it, so its place follows
/// the columns and rows the renderer has already worked out. `None` when it
/// hangs from a cell outside what is being drawn.
#[cfg(windows)]
pub(crate) fn drawing_box(
    drawn: &oxicells_core::ir::Drawing,
    layout: &Geometry,
    scale: f32,
) -> Option<windows::Win32::Foundation::RECT> {
    anchored_box(&drawn.from, drawn.to.as_ref(), drawn.extent, layout, scale)
}

/// The box between two anchors, or between one and a stated size.
#[cfg(windows)]
pub(crate) fn anchored_box(
    from: &oxicells_core::ir::Anchor,
    to: Option<&oxicells_core::ir::Anchor>,
    extent: Option<(i64, i64)>,
    layout: &Geometry,
    scale: f32,
) -> Option<windows::Win32::Foundation::RECT> {
    let at = |anchor: &oxicells_core::ir::Anchor| -> Option<(i32, i32)> {
        // The drawing part counts both from zero; the layout counts columns
        // from zero and rows from one, the way the sheet states them. A cell
        // before the range has its own place, back from the left edge.
        let left = match anchor.col.checked_sub(layout.first_column) {
            Some(column) => match layout.columns.get(column as usize) {
                Some(edge) => *edge,
                // Past the right of the range, where the picture stops but
                // the sheet does not.
                None => *layout
                    .after_columns
                    .get(column as usize - layout.columns.len())?,
            },
            None => *layout.before_columns.get(anchor.col as usize)?,
        };
        let top = match (anchor.row + 1).checked_sub(layout.first_row) {
            Some(row) => match layout.rows.get(row as usize) {
                Some(edge) => *edge,
                None => *layout.after_rows.get(row as usize - layout.rows.len())?,
            },
            None => *layout.before_rows.get(anchor.row as usize)?,
        };
        Some((
            (left + anchor.col_off as f32 / EMU * scale).round() as i32,
            (top + anchor.row_off as f32 / EMU * scale).round() as i32,
        ))
    };
    let (left, top) = at(from)?;
    let (right, bottom) = match (to, extent) {
        // A corner past the drawn range falls off the picture, which is where
        // the sheet's own edge is: keep the box and let the drawing be cut.
        (Some(to), _) => at(to).unwrap_or((
            *layout.columns.last().unwrap_or(&0.0) as i32,
            *layout.rows.last().unwrap_or(&0.0) as i32,
        )),
        (None, Some((cx, cy))) => (
            left + (cx as f32 / EMU * scale).round() as i32,
            top + (cy as f32 / EMU * scale).round() as i32,
        ),
        (None, None) => return None,
    };
    Some(windows::Win32::Foundation::RECT { left, top, right, bottom })
}

/// How wide the box is before its edges are put on whole pixels.
///
/// A shape's anchors carry EMU, and a box that starts 134.4 pixels into one
/// column and ends 65.6 into another is 0.8 of a pixel narrower or wider than
/// the rectangle it is drawn in. Excel breaks a line when the run is longer
/// than that unrounded room and not before: `_xlsx_shape_room.py` sweeps a box
/// a quarter of a pixel at a time and the turn is exactly at
/// `room = run` — which is what puts `sanko_tool`'s last character on a line
/// of its own here and not there.
#[cfg(windows)]
pub(crate) fn drawing_room(
    drawn: &oxicells_core::ir::Drawing,
    layout: &Geometry,
    scale: f32,
) -> Option<f32> {
    anchored_room(&drawn.from, drawn.to.as_ref(), drawn.extent, layout, scale)
}

#[cfg(windows)]
pub(crate) fn anchored_room(
    from: &oxicells_core::ir::Anchor,
    to: Option<&oxicells_core::ir::Anchor>,
    extent: Option<(i64, i64)>,
    layout: &Geometry,
    scale: f32,
) -> Option<f32> {
    let at = |anchor: &oxicells_core::ir::Anchor| -> Option<f32> {
        let left = match anchor.col.checked_sub(layout.first_column) {
            Some(column) => match layout.columns.get(column as usize) {
                Some(edge) => *edge,
                None => *layout
                    .after_columns
                    .get(column as usize - layout.columns.len())?,
            },
            None => *layout.before_columns.get(anchor.col as usize)?,
        };
        Some(left + anchor.col_off as f32 / EMU * scale)
    };
    let left = at(from)?;
    let right = match (to, extent) {
        (Some(to), _) => at(to).unwrap_or(*layout.columns.last().unwrap_or(&0.0)),
        (None, Some((cx, _))) => left + cx as f32 / EMU * scale,
        (None, None) => return None,
    };
    Some(right - left)
}

/// The face Excel draws when the workbook asks for one this machine has not.
///
/// Not the name, and not the PANOSE the file carries: `_xlsx_missing_face_map.py`
/// gives every arm a workbook of its own — a name resolves once per document,
/// so two arms sharing a name in one book share its answer — and sweeps them
/// separately. Six names, from the two the corpus asks for to one that never
/// existed, all answer the same way, and only the run's charset moves it:
/// Japanese (-128) draws in 游ゴシック, everything else in ＭＳ ゴシック,
/// whatever the pitchFamily says. GDI's own mapper answers ＭＳ Ｐゴシック to
/// both, which is neither.
pub(crate) fn face_in_place(face: &str, charset: Option<i32>) -> String {
    if face.is_empty() || installed(face) {
        return face.to_string();
    }
    // -128 is SHIFT_JIS as a signed byte; a file can also write it as 128.
    match charset {
        Some(-128) | Some(128) => "游ゴシック".to_string(),
        _ => "ＭＳ ゴシック".to_string(),
    }
}

/// Whether the device has this face, or is quietly drawing something else.
///
/// GDI has no way to say "I have not got that": it hands back a face of its
/// own choosing. So the question is asked twice — once for the face, once for
/// a name nothing can have — and a face that answers the way the impossible
/// one does is a face the device has not got.
#[cfg(windows)]
fn installed(face: &str) -> bool {
    thread_local! {
        static KNOWN: std::cell::RefCell<std::collections::HashMap<String, bool>> =
            std::cell::RefCell::new(std::collections::HashMap::new());
    }
    if let Some(held) = KNOWN.with(|known| known.borrow().get(face).copied()) {
        return held;
    }
    const NOTHING: &str = "Nonesuch Face ZZQ";
    let answer = physical_face(face);
    let fallback = physical_face(NOTHING);
    let held = match (answer, fallback) {
        (Some(answer), Some(fallback)) => answer != fallback || answer == face,
        _ => true,
    };
    KNOWN.with(|known| known.borrow_mut().insert(face.to_string(), held));
    held
}

#[cfg(not(windows))]
fn installed(_face: &str) -> bool {
    true
}

/// The face the device actually draws when asked for this one.
///
/// `GetTextFace` answers with the name that was asked for, whether or not
/// anything answered to it; the outline metrics carry the name of the face
/// that was actually realised, which is the one worth comparing.
#[cfg(windows)]
fn physical_face(face: &str) -> Option<String> {
    use windows::core::PCWSTR;
    use windows::Win32::Graphics::Gdi::*;
    unsafe {
        let screen = GetDC(None);
        if screen.is_invalid() {
            return None;
        }
        let dc = CreateCompatibleDC(screen);
        let name: Vec<u16> = face.encode_utf16().chain(Some(0)).collect();
        let font = CreateFontW(
            -16,
            0,
            0,
            0,
            400,
            0,
            0,
            0,
            DEFAULT_CHARSET.0 as u32,
            OUT_DEFAULT_PRECIS.0 as u32,
            CLIP_DEFAULT_PRECIS.0 as u32,
            DEFAULT_QUALITY.0 as u32,
            (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
            PCWSTR(name.as_ptr()),
        );
        let previous = SelectObject(dc, font);
        let mut answer = None;
        let size = GetOutlineTextMetricsW(dc, 0, None);
        if size > 0 {
            let mut held: Vec<u8> = vec![0; size as usize];
            let metrics = held.as_mut_ptr() as *mut OUTLINETEXTMETRICW;
            if GetOutlineTextMetricsW(dc, size, Some(metrics)) != 0 {
                // The name is stored inside the same block, as an offset in
                // bytes from its start.
                let at = (*metrics).otmpFaceName.0 as usize;
                if at > 0 && at < held.len() {
                    let letters = held.as_ptr().add(at) as *const u16;
                    let mut end = 0usize;
                    while at + (end + 1) * 2 <= held.len() && *letters.add(end) != 0 {
                        end += 1;
                    }
                    answer = Some(String::from_utf16_lossy(
                        std::slice::from_raw_parts(letters, end),
                    ));
                }
            }
        }
        SelectObject(dc, previous);
        let _ = DeleteObject(font);
        let _ = DeleteDC(dc);
        ReleaseDC(None, screen);
        answer
    }
}

/// One level of indent, in 96-dpi pixels, when the workbook's own font
/// cannot be measured.
pub(crate) const INDENT: f32 = 15.0;

/// One level of indent: three spaces of the workbook's first font.
///
/// `_xlsx_indent.py` reads fifteen pixels a level whatever the cell wears,
/// and called it a constant; the `h2daa*dendeba_kmc` trio reads twelve.
/// `_xlsx_indent_bisect.py` puts one property of that workbook back at a
/// time: only the first font in the list moves it, and moving it moves the
/// level with the width of that font's space — ＭＳ Ｐゴシック 11 five and
/// fifteen, ＭＳ ゴシック 11 eight and twenty-four, 游ゴシック 11 four and
/// twelve, Calibri 11 three and nine, at 8 and 14 point likewise, in twelve
/// arms. The probe read a constant because every book it wrote resolves its
/// first font to the same ＭＳ Ｐゴシック through the theme.
///
/// It is the first font in the list, not the one the Normal style points at:
/// pointing Normal elsewhere leaves the level where it was.
pub(crate) fn indent_level(sheet: &oxicells_core::ir::Sheet) -> f32 {
    sheet
        .first_font
        .as_ref()
        .and_then(|(face, points)| {
            let spaces = advances(face, *points, false, false, " ")?;
            spaces.first().map(|space| 3.0 * *space as f32)
        })
        .unwrap_or(INDENT)
}

/// How far this cell's text is pushed in from the edge it sits against.
pub(crate) fn indent_px(style: &CellStyle, level: f32) -> f32 {
    level * style.indent as f32
}

/// How much room an indent takes from a cell, before and after the text.
///
/// Against the left or the right edge the indent comes off that edge alone.
/// A distributed cell gives it up at **both** edges: at `indent="1"` its text
/// runs from 18 to 139 of a column whose room is 3 to 154, which is the
/// fifteen pixels off each end. A centred cell loses the same total from its
/// right, which is why its text sits 15px left of the middle a level rather
/// than staying in the middle — all measured on `_xlsx_distributed.py` and
/// `_xlsx_indent.py`.
///
/// The wrapped line is broken in what is left, and both kinds do wrap tighter
/// with an indent: 4, 4, 5, 7 lines at indents 0 to 3 against the left edge,
/// 4, 5, 10, 20 centred (`_xlsx_indent_wrap.py`).
pub(crate) fn indent_room(style: &CellStyle, level: f32) -> (f32, f32) {
    let indent = indent_px(style, level);
    match style.horizontal_align.as_deref() {
        Some("distributed") | Some("justify") => (indent, indent),
        Some("center") | Some("centre") | Some("centerContinuous") => (0.0, indent * 2.0),
        Some("right") => (0.0, indent),
        _ => (indent, 0.0),
    }
}

pub(crate) fn gutters(face: &str, points: f32, bold: bool, italic: bool) -> (f32, f32) {
    let digit = advances(face, points, bold, italic, "0")
        .and_then(|held| held.first().copied())
        .unwrap_or(7) as f32;
    let extra = (((digit - 5.0) / 4.0).floor()).max(0.0);
    (3.0 + extra, 2.0 + extra)
}

/// The box Excel lays one line of this font in, and how far down that box its
/// baseline sits.
///
/// The box is the row a sheet of nothing but this font would have — the height
/// the renderer already carries as a measured table, and the same height a
/// wrapped cell spends per line. Where the baseline sits in it was measured
/// against Excel's own picture across twenty faces: it is the device's descent
/// above the bottom of the box for the ＭＳ faces, Calibri, Arial and Consolas,
/// and a pixel or two higher for the faces with a large internal leading
/// (Meiryo, the 游 family, Segoe UI, Times New Roman), which is the part of
/// this that is not yet derived.
///
/// `None` when the platform cannot be asked, so the caller draws the way it
/// did before.
pub(crate) fn line_box(face: &str, points: f32, bold: bool, italic: bool) -> Option<(f32, f32)> {
    match row_defaults::font_line_box(face, points, bold) {
        Some((box_px, baseline)) => Some((box_px as f32, baseline as f32)),
        // A face the table has never seen: the device's own line, with the
        // baseline where the descent puts it, which is where two thirds of the
        // measured table sits anyway.
        _ => {
            let (_, descent, height) =
                held(|counter| counter.shape_of(face, points, bold, italic))?;
            Some((height, height - descent))
        }
    }
}

/// Where one paragraph breaks in a box this wide, given what each character
/// advances: the first character of each line after the first. The line holds
/// characters until the next would not fit, then gives them back one at a time
/// until the break is one Excel would make.
pub(crate) fn line_breaks(letters: &[char], advances: &[i32], width: f32) -> Vec<usize> {
    let held: Vec<f32> = advances.iter().map(|advance| *advance as f32).collect();
    // A cell measures its room in whole pixels, and always has, and does not
    // hang its punctuation (`_xlsx_cell_hang.py`).
    broken_at(letters, &held, width.max(1.0).trunc(), false)
}

/// The two characters a shape lets hang past the end of its line.
///
/// `_xlsx_shape_hang.py` sweeps a box's room across the point where a line's
/// last character stops fitting, over eight last characters. 「。」 and 「、」
/// hold the line together all the way down to the room the body alone needs —
/// they hang their whole em — and 「」」「）」「！」「ゃ」「あ」 break the moment
/// the room is two pixels short. (「)」 looks like it hangs and does not: it is
/// half-width, so it simply fits for longer.)
///
/// A CELL does not do this. `_xlsx_cell_hang.py` asks the same question of a
/// wrapping cell — same face, same body, the column swept through the same
/// crossing — and every one of 。、」あ takes two lines. So this is a shape's
/// rule, not the breaker's, and it is passed in rather than assumed.
fn hangs(letter: char) -> bool {
    matches!(letter, '。' | '、')
}

fn broken_at(letters: &[char], advances: &[f32], width: f32, hang: bool) -> Vec<usize> {
    let mut breaks = Vec::new();
    if letters.is_empty() {
        return breaks;
    }
    let room = width.max(1.0);
    let mut start = 0usize;
    while start < letters.len() {
        let mut take = 0usize;
        let mut run = 0.0f32;
        while start + take < letters.len() {
            let next = run + advances[start + take];
            // A hair of slack: the fractions are worked out from the font's
            // own design units and Excel's arithmetic is not this one's, so a
            // line whose sum lands on the room to a thousandth is a line that
            // fits.
            if take > 0 && next > room + 0.01 {
                // A 句読点 that will not fit hangs past the end instead of
                // taking the character before it down to the next line —
                // which is what kinsoku would otherwise do, since it may not
                // start one. `001`'s panel ends a line on 「ください。」 and
                // Excel keeps all of it.
                if hang && hangs(letters[start + take]) {
                    take += 1;
                }
                break;
            }
            run = next;
            take += 1;
        }
        // Give characters back until the break is one Excel would make. A
        // run with nowhere to break — a long web address — is cut where it
        // stops fitting rather than left to fill a line a character at a
        // time, which is what Excel draws.
        let fill = take;
        while start + take < letters.len()
            && take > 1
            && !may_break(letters[start + take - 1], letters[start + take])
        {
            take -= 1;
        }
        // A line of one character is a line: Excel leaves 者 on its own
        // rather than end the line before （. Only a run with no break in it
        // anywhere — a web address — is cut where it stops fitting instead.
        let nowhere_to_break = take == 1
            && start + 1 < letters.len()
            && !may_break(letters[start], letters[start + 1]);
        if nowhere_to_break && fill > 1 {
            take = fill;
        }
        // The spaces at a break belong to the line they end. Excel starts the
        // next line at the first character past them, however many there are.
        while start + take < letters.len() && letters[start + take] == ' ' {
            take += 1;
        }
        start += take.max(1);
        if start < letters.len() {
            breaks.push(start);
        }
    }
    breaks
}

/// How many lines one paragraph takes in a box this wide.
fn count_lines(letters: &[char], advances: &[i32], width: f32) -> u32 {
    line_breaks(letters, advances, width).len() as u32 + 1
}

/// The lines a cell's text is drawn as: its own breaks, and the wrapping ones
/// where it wraps. The same rule that gives the row its height, so what is
/// drawn and what the row was measured for cannot disagree.
pub(crate) fn wrapped_lines(
    face: &str,
    points: f32,
    bold: bool,
    italic: bool,
    text: &str,
    width: Option<f32>,
) -> Vec<String> {
    broken_lines(face, points, bold, italic, text, width, false)
}

/// The same, for a shape, which breaks its lines where its own run of glyphs
/// runs out of room.
///
/// A cell measures a break in whole-pixel advances; a shape sets its text by
/// the font's own at the exact em (`shape_run`), and a break has to be
/// measured in the advances the text is actually set in. The two part company
/// by a fraction of a pixel a character, which over `sanko_tool`'s
/// thirty-three-character line is enough to push its last letter onto a line
/// of its own — and a block one line too tall hangs its first line above a
/// middle-anchored box. `_xlsx_shape_advance.py` walks that very line a
/// character at a time and finds the shape's own advances land on Excel's ink
/// at all twenty-six lengths, so they are what a shape's break is measured in.
pub(crate) fn wrapped_shape_lines(
    face: &str,
    points: f32,
    bold: bool,
    italic: bool,
    text: &str,
    width: Option<f32>,
) -> Vec<String> {
    broken_lines(face, points, bold, italic, text, width, true)
}

fn broken_lines(
    face: &str,
    points: f32,
    bold: bool,
    italic: bool,
    text: &str,
    width: Option<f32>,
    shape: bool,
) -> Vec<String> {
    let mut held = Vec::new();
    for paragraph in text.split('\n') {
        let letters: Vec<char> = paragraph.chars().collect();
        let breaks = match width {
            Some(width) if shape => shape_widths(face, points, bold, italic, paragraph)
                .map(|advances| broken_at(&letters, &advances, width, true))
                .unwrap_or_default(),
            Some(width) => advances(face, points, bold, italic, paragraph)
                .map(|advances| line_breaks(&letters, &advances, width))
                .unwrap_or_default(),
            None => Vec::new(),
        };
        let mut at = 0usize;
        // The last stretch runs to the end of the paragraph, and an empty
        // paragraph is one empty line rather than none.
        for stop in breaks.iter().copied().chain(std::iter::once(letters.len())) {
            held.push(letters[at..stop].iter().collect());
            at = stop;
        }
    }
    held
}

/// Counts the lines a cell's text wraps into, measuring the way Excel does:
/// each character advances what the font gives it on its own, and the line
/// is the sum of those. A text engine will not do: DirectWrite's advances
/// run a fraction under GDI's, and GDI's own run measurement compresses
/// neighbouring Japanese punctuation, which Excel's PDF shows it does not —
/// 49ac46's D16 measures 589px as a run and 592px character by character,
/// and Excel wraps it as 592.
/// A face at a size, in the weight and slant it is asked for.
#[cfg(windows)]
type FontKey = (String, u32, bool, bool);

/// What every character measured so far advances, per font.
#[cfg(windows)]
type AdvanceCache =
    std::collections::HashMap<FontKey, std::collections::HashMap<char, i32>>;

/// How the device lays a line of a font out: ascent, descent, and the height
/// of the two together.
#[cfg(windows)]
type LineShape = (f32, f32, f32);

#[cfg(windows)]
struct LineCounter {
    dc: windows::Win32::Graphics::Gdi::HDC,
    /// One character's advance, kept per font so a sheet of one typeface
    /// asks the device about each character once.
    advances: std::cell::RefCell<AdvanceCache>,
    /// What the device says about each font as a whole.
    shapes: std::cell::RefCell<std::collections::HashMap<FontKey, LineShape>>,
    /// What each character advances as a share of the em, measured once per
    /// face at a size large enough that the device's own rounding is lost in
    /// it. A shape's text is set by these rather than by the whole-pixel
    /// advances a cell's is: see `design_advances`.
    design: std::cell::RefCell<
        std::collections::HashMap<(String, bool, bool), std::collections::HashMap<char, f32>>,
    >,
}

#[cfg(windows)]
impl LineCounter {
    fn new() -> Option<Self> {
        unsafe {
            let dc = windows::Win32::Graphics::Gdi::GetDC(None);
            if dc.is_invalid() {
                return None;
            }
            Some(Self {
                dc,
                advances: std::cell::RefCell::new(std::collections::HashMap::new()),
                shapes: std::cell::RefCell::new(std::collections::HashMap::new()),
                design: std::cell::RefCell::new(std::collections::HashMap::new()),
            })
        }
    }

    /// What the device makes of this font: how far its ascent reaches above
    /// the baseline, how far its descent falls below, and the two together.
    fn shape_of(&self, face: &str, points: f32, bold: bool, italic: bool) -> Option<LineShape> {
        use windows::core::PCWSTR;
        use windows::Win32::Graphics::Gdi::*;
        let pixels = (points * 96.0 / 72.0).round();
        let key = (face.to_string(), pixels as u32, bold, italic);
        if let Some(held) = self.shapes.borrow().get(&key) {
            return Some(*held);
        }
        let shape = unsafe {
            let name: Vec<u16> = face.encode_utf16().chain(Some(0)).collect();
            let font = CreateFontW(
                -(pixels as i32),
                0,
                0,
                0,
                if bold { 700 } else { 400 },
                u32::from(italic),
                0,
                0,
                DEFAULT_CHARSET.0 as u32,
                OUT_DEFAULT_PRECIS.0 as u32,
                CLIP_DEFAULT_PRECIS.0 as u32,
                ANTIALIASED_QUALITY.0 as u32,
                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                PCWSTR(name.as_ptr()),
            );
            if font.is_invalid() {
                return None;
            }
            let previous = SelectObject(self.dc, font);
            let mut measured = TEXTMETRICW::default();
            let ok = GetTextMetricsW(self.dc, &mut measured).as_bool();
            SelectObject(self.dc, previous);
            let _ = DeleteObject(font);
            if !ok {
                return None;
            }
            (
                measured.tmAscent as f32,
                measured.tmDescent as f32,
                measured.tmHeight as f32,
            )
        };
        self.shapes.borrow_mut().insert(key, shape);
        Some(shape)
    }

    /// What each of these characters advances as a share of the em.
    ///
    /// A cell's text is set by whole-pixel advances, which is what Excel
    /// measures a wrap against. A shape's is not: `311e2f9c271e_zuhyo`'s
    /// footnote is ＭＳ 明朝 at 11 point, whose glyphs are 13 pixels of ink in
    /// both pictures, but Excel steps 14.67 pixels a character where the
    /// device's own advance at that size is 16 — three quarters of a pixel a
    /// character, which is forty by the end of the line. So the shape is
    /// measured once at a large size, where the device's rounding is a
    /// thousandth of the answer, and the exact em does the rest.
    fn design_advances(
        &self,
        face: &str,
        bold: bool,
        italic: bool,
        letters: &[char],
    ) -> Option<Vec<f32>> {
        use windows::core::PCWSTR;
        use windows::Win32::Foundation::SIZE;
        use windows::Win32::Graphics::Gdi::*;
        /// Big enough that a whole-pixel advance is a rounding of a thousandth.
        const EM: i32 = 2048;
        let key = (face.to_string(), bold, italic);
        let mut held = self.design.borrow_mut();
        let known = held.entry(key).or_default();
        let wanted: Vec<char> = letters
            .iter()
            .copied()
            .filter(|letter| !known.contains_key(letter))
            .collect();
        if !wanted.is_empty() {
            unsafe {
                let name: Vec<u16> = face.encode_utf16().chain(Some(0)).collect();
                let font = CreateFontW(
                    -EM,
                    0,
                    0,
                    0,
                    if bold { 700 } else { 400 },
                    u32::from(italic),
                    0,
                    0,
                    DEFAULT_CHARSET.0 as u32,
                    OUT_DEFAULT_PRECIS.0 as u32,
                    CLIP_DEFAULT_PRECIS.0 as u32,
                    ANTIALIASED_QUALITY.0 as u32,
                    (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                    PCWSTR(name.as_ptr()),
                );
                if font.is_invalid() {
                    return None;
                }
                let previous = SelectObject(self.dc, font);
                for letter in wanted {
                    let mut measured = SIZE::default();
                    let one: Vec<u16> = letter.encode_utf16(&mut [0; 2]).to_vec();
                    let ok = GetTextExtentPoint32W(self.dc, &one, &mut measured).as_bool();
                    known.insert(letter, if ok { measured.cx as f32 / EM as f32 } else { 0.0 });
                }
                SelectObject(self.dc, previous);
                let _ = DeleteObject(font);
            }
        }
        Some(letters.iter().map(|letter| known[letter]).collect())
    }

    /// What each of these characters advances, in whole pixels.
    fn advances_of(
        &self,
        face: &str,
        points: f32,
        bold: bool,
        italic: bool,
        letters: &[char],
    ) -> Option<Vec<i32>> {
        use windows::core::PCWSTR;
        use windows::Win32::Foundation::SIZE;
        use windows::Win32::Graphics::Gdi::*;
        let pixels = (points * 96.0 / 72.0).round();
        let key = (face.to_string(), pixels as u32, bold, italic);
        let mut held = self.advances.borrow_mut();
        let known = held.entry(key).or_default();
        let wanted: Vec<char> = letters
            .iter()
            .copied()
            .filter(|letter| !known.contains_key(letter))
            .collect();
        if !wanted.is_empty() {
            unsafe {
                let name: Vec<u16> = face.encode_utf16().chain(Some(0)).collect();
                let font = CreateFontW(
                    -(pixels as i32),
                    0,
                    0,
                    0,
                    if bold { 700 } else { 400 },
                    u32::from(italic),
                    0,
                    0,
                    DEFAULT_CHARSET.0 as u32,
                    OUT_DEFAULT_PRECIS.0 as u32,
                    CLIP_DEFAULT_PRECIS.0 as u32,
                    ANTIALIASED_QUALITY.0 as u32,
                    (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                    PCWSTR(name.as_ptr()),
                );
                if font.is_invalid() {
                    return None;
                }
                let previous = SelectObject(self.dc, font);
                for letter in wanted {
                    let one: Vec<u16> = letter.to_string().encode_utf16().collect();
                    let mut size = SIZE::default();
                    let ok = GetTextExtentPoint32W(self.dc, &one, &mut size).as_bool();
                    known.insert(letter, if ok { size.cx } else { 0 });
                }
                SelectObject(self.dc, previous);
                let _ = DeleteObject(font);
            }
        }
        Some(letters.iter().map(|letter| known[letter]).collect())
    }

    fn lines(
        &self,
        face: &str,
        points: f32,
        bold: bool,
        italic: bool,
        text: &str,
        width: f32,
    ) -> Option<u32> {
        // Each paragraph is broken on its own.
        let mut total = 0u32;
        for paragraph in text.split('\n') {
            let letters: Vec<char> = paragraph.chars().collect();
            let advances = self.advances_of(face, points, bold, italic, &letters)?;
            if std::env::var("OXI_XLSX_DUMP_ADVANCES").is_ok() {
                eprintln!(
                    "      advances in {:.0}px: sum {} first {:?}",
                    width,
                    advances.iter().sum::<i32>(),
                    advances.iter().take(8).collect::<Vec<_>>()
                );
            }
            total += count_lines(&letters, &advances, width);
        }
        Some(total.max(1))
    }
}

#[cfg(windows)]
impl Drop for LineCounter {
    fn drop(&mut self) {
        unsafe {
            windows::Win32::Graphics::Gdi::ReleaseDC(None, self.dc);
        }
    }
}

#[cfg(not(windows))]
struct LineCounter;

#[cfg(not(windows))]
impl LineCounter {
    fn new() -> Option<Self> {
        None
    }

    fn lines(
        &self,
        _face: &str,
        _points: f32,
        _bold: bool,
        _italic: bool,
        _text: &str,
        _width: f32,
    ) -> Option<u32> {
        None
    }

    fn advances_of(
        &self,
        _face: &str,
        _points: f32,
        _bold: bool,
        _italic: bool,
        _letters: &[char],
    ) -> Option<Vec<i32>> {
        None
    }

    fn shape_of(
        &self,
        _face: &str,
        _points: f32,
        _bold: bool,
        _italic: bool,
    ) -> Option<(f32, f32, f32)> {
        None
    }
}

/// The pre-derivation reading of the stated default, kept for sheets whose
/// fonts the table does not know: rounded UP to the next 0.75, which matched
/// Excel on most Excel-authored files because their stated number was already
/// the font-derived one.
fn fallback_row_points(sheet: &Sheet) -> f32 {
    let stated = if sheet.default_row_height > 0.0 {
        sheet.default_row_height
    } else {
        DEFAULT_ROW_POINTS
    };
    (stated / 0.75).ceil() * 0.75
}

fn geometry(sheet: &Sheet, scale: f32, digit_width: f32) -> Geometry {
    let (first_row, first_column, last_row, last_column) = used_extent(sheet);

    let column_width = |column: u32| -> f32 {
        if sheet.hidden_cols.contains(&column) {
            return 0.0;
        }
        let stated = sheet
            .col_widths
            .get(column as usize)
            .copied()
            .filter(|width| *width > 0.0);
        match stated {
            // A stated width already carries the gutter either side of the
            // cell's text.
            Some(width) => column_pixels(width, digit_width) * scale,
            // A default the sheet states is measured the same way. Excel's
            // own default, for a sheet that states none, is a plain count
            // of characters, so there the gutter is added here instead.
            None if sheet.default_col_width > 0.0 => {
                column_pixels(sheet.default_col_width, digit_width) * scale
            }
            None => (DEFAULT_DIGITS * digit_width + DEFAULT_PADDING).trunc() * scale,
        }
    };

    let mut columns = vec![0.0];
    for column in first_column..=last_column {
        columns.push(columns.last().unwrap() + column_width(column));
    }

    // Where the columns before the range sit, as offsets back from its left
    // edge: a drawing can hang from a cell outside the picture and still
    // reach into it, which is how `002` lays a banner across the top of a
    // sheet whose used range starts two columns in.
    let mut before_columns = vec![0.0; first_column as usize];
    let mut back = 0.0;
    for column in (0..first_column).rev() {
        back -= column_width(column);
        before_columns[column as usize] = back;
    }

    let default_px = default_row_points(sheet) / 0.75;
    if std::env::var("OXI_XLSX_DUMP_LINES").is_ok() {
        eprintln!(
            "sheet default {:.0}px from stated {} custom {} normal {:?} columns {:?}",
            default_px, sheet.default_row_height, sheet.default_row_custom,
            sheet.normal_font, sheet.col_fonts
        );
    }
    let counter = LineCounter::new();
    let merged = merges(sheet);
    let mut rows = vec![0.0];
    let row_height = |index: u32, columns: &[f32]| -> f32 {
        let held = sheet.rows.iter().find(|row| row.index == index);
        let hidden = held.is_some_and(|row| row.hidden);
        if hidden {
            0.0
        } else {
            // The stretches of columns a merge swallows in this row.
            let merged_columns: Vec<(u32, u32)> = sheet
                .merge_cells
                .iter()
                .filter(|merge| merge.start_row <= index && index <= merge.end_row)
                .map(|merge| (merge.start_col, merge.end_col))
                .collect();
            let px = row_pixels(
                held,
                sheet,
                &merged_columns,
                default_px,
                &columns,
                first_column,
                scale,
                counter.as_ref(),
                &merged,
            );
            (px * scale).trunc()
        }
    };
    for index in first_row..=last_row {
        rows.push(rows.last().unwrap() + row_height(index, &columns));
    }
    // The rows above the range, the same way round as the columns before it.
    let mut before_rows = vec![0.0; first_row.saturating_sub(1) as usize];
    let mut back = 0.0;
    for index in (1..first_row).rev() {
        back -= row_height(index, &columns);
        before_rows[(index - 1) as usize] = back;
    }

    // How far past the range anything hangs. Only what is asked for is
    // measured: a sheet is a million rows wide of nothing much.
    let (reach_column, reach_row) = sheet
        .drawings
        .iter()
        .filter_map(|drawn| drawn.to.as_ref())
        .fold((0, 0), |(column, row), anchor| {
            (column.max(anchor.col), row.max(anchor.row + 1))
        });
    let mut after_columns = Vec::new();
    let mut ahead = *columns.last().unwrap_or(&0.0);
    for column in (last_column + 2)..=reach_column.min(last_column + 1024) {
        ahead += column_width(column - 1);
        after_columns.push(ahead);
    }
    let mut after_rows = Vec::new();
    let mut below = *rows.last().unwrap_or(&0.0);
    for index in (last_row + 2)..=reach_row.min(last_row + 4096) {
        below += row_height(index - 1, &columns);
        after_rows.push(below);
    }

    Geometry {
        columns,
        rows,
        before_columns,
        before_rows,
        after_columns,
        after_rows,
        first_column,
        first_row,
    }
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() < 3 {
        eprintln!("Usage: {} <input.xlsx> <output.png> [dpi]", args[0]);
        std::process::exit(1);
    }
    let dpi: f32 = args.get(3).and_then(|value| value.parse().ok()).unwrap_or(96.0);
    let scale = dpi / 96.0;

    let data = match std::fs::read(&args[1]) {
        Ok(data) => data,
        Err(error) => {
            eprintln!("cannot read {}: {error}", args[1]);
            std::process::exit(1);
        }
    };
    let workbook: Workbook = match parse_xlsx(&data) {
        Ok(workbook) => workbook,
        Err(error) => {
            eprintln!("cannot parse {}: {error}", args[1]);
            std::process::exit(1);
        }
    };
    let Some(sheet) = workbook.sheets.first() else {
        eprintln!("the workbook holds no sheets");
        std::process::exit(1);
    };

    // A column is measured in digits of the workbook's standard font, so the
    // sheet cannot be laid out until that digit has been measured.
    let digit_width = digit_width(&workbook.default_style).unwrap_or(FALLBACK_DIGIT_WIDTH);
    let layout = geometry(sheet, scale, digit_width);
    // One pixel past the last edge, because a rule on that edge is drawn there
    // and Excel's own picture is that pixel wider and taller.
    let width = *layout.columns.last().unwrap_or(&0.0) as u32 + 1;
    let height = *layout.rows.last().unwrap_or(&0.0) as u32 + 1;
    if width == 0 || height == 0 {
        eprintln!("the sheet has nothing in it to draw");
        std::process::exit(1);
    }

    // Row-by-row geometry, for holding the model against Excel's answers.
    if std::env::var("OXI_XLSX_DUMP_ROWS").is_ok() {
        for (step, pair) in layout.rows.windows(2).enumerate() {
            println!("row {} px {}", layout.first_row as usize + step, pair[1] - pair[0]);
        }
    }
    if std::env::var("OXI_XLSX_DUMP_COLUMNS").is_ok() {
        for (step, pair) in layout.columns.windows(2).enumerate() {
            println!(
                "column {} px {}",
                layout.first_column as usize + step,
                pair[1] - pair[0]
            );
        }
    }

    let canvas = draw(sheet, &layout, width, height, scale);
    if let Err(error) = canvas.save(&args[2]) {
        eprintln!("cannot write {}: {error}", args[2]);
        std::process::exit(1);
    }
    println!("{} {}x{} @{}dpi", args[2], width, height, dpi);
}

/// How wide a digit is in the font a workbook treats as standard.
#[cfg(windows)]
fn digit_width(style: &CellStyle) -> Option<f32> {
    windows_draw::digit_width(style)
}

#[cfg(not(windows))]
fn digit_width(_style: &CellStyle) -> Option<f32> {
    None
}

#[cfg(windows)]
fn draw(
    sheet: &Sheet,
    layout: &Geometry,
    width: u32,
    height: u32,
    scale: f32,
) -> image::RgbImage {
    // Excel's sheet text is GDI's own ClearType — fringe colour for fringe
    // colour, measured against its picture — and Direct2D cannot be asked for
    // the same: over a WIC bitmap it ignores its text rendering parameters
    // entirely. So GDI draws, and DirectWrite stays reachable for comparison
    // (0.9478 against 0.9456 over the 285-workbook corpus, and ahead on 128
    // of them to DirectWrite's 89).
    if std::env::var("OXI_XLSX_DWRITE").is_ok() {
        match dwrite_draw::draw(sheet, layout, width, height, scale) {
            Ok(canvas) => return canvas,
            Err(error) => eprintln!("DirectWrite could not draw the sheet: {error}"),
        }
    }
    windows_draw::draw(sheet, layout, width, height, scale)
}

#[cfg(not(windows))]
fn draw(
    _sheet: &Sheet,
    _layout: &Geometry,
    width: u32,
    height: u32,
    _scale: f32,
) -> image::RgbImage {
    // Text needs the platform's own rasteriser to match Excel, so elsewhere
    // this only reports the geometry it would have drawn into.
    image::RgbImage::from_pixel(width, height, image::Rgb([255, 255, 255]))
}

/// The text a cell shows, under whatever number format it wears.
pub(crate) fn cell_text(value: &CellValue, style: &CellStyle) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::String(text) => text.clone(),
        CellValue::Boolean(true) => "TRUE".to_string(),
        CellValue::Boolean(false) => "FALSE".to_string(),
        CellValue::Error(text) => text.clone(),
        CellValue::Number(number) => oxicells_core::format_number(
            *number,
            style.number_format.as_deref().unwrap_or("General"),
        ),
    }
}

/// A stacked cell's text, one character to a line.
///
/// `textRotation="255"` stands the characters one above the next rather than
/// turning them on their side, which is how a Japanese form labels a narrow
/// column. Each character takes a whole line of the font, so the cell is the
/// same thing as text with a break after every character — and everything
/// downstream, from the row's height to where each line is drawn, follows.
pub(crate) fn stacked_text(text: &str) -> String {
    let mut held = String::with_capacity(text.len() * 2);
    for letter in text.chars().filter(|letter| *letter != '\n') {
        if !held.is_empty() {
            held.push('\n');
        }
        held.push(letter);
    }
    held
}

/// Where the text sits across a cell. Excel puts numbers to the right and text
/// to the left unless the cell says otherwise.
/// How a table dresses one of its cells: Excel paints this, and no cell inside
/// the range carries any of it in its own style.
pub(crate) struct Dressed {
    fill: Option<String>,
    /// A header row is written in white on the accent colour, and in bold:
    /// Excel's own table styles set the header heavier than the body, and no
    /// cell inside the range says so in its own style.
    white_text: bool,
    bold: bool,
}

/// A filtered column carries a button in its heading. Measured against Excel
/// across headings of 20, 30, 45 and 90 pixels and columns of two widths: 17
/// by 17 pixels, its right edge a pixel in from the cell's, and its foot two
/// pixels up from the cell's foot — it hangs from the bottom of the heading,
/// wherever the top of it is. A pale face inside a grey outline, with a
/// seven-wide triangle narrowing to a point four rows down.
pub(crate) const FILTER_BUTTON: i32 = 17;
/// How far the button's foot sits above the cell's.
pub(crate) const FILTER_BUTTON_FOOT: i32 = 2;

/// Whether the cell at this spot carries a filter button.
///
/// A table's header row carries one in every column, and so does the heading
/// row of a sheet's own `<autoFilter>` — the procurement lists filter that
/// way, with no table in sight.
pub(crate) fn has_filter_button(sheet: &Sheet, row: u32, column: u32) -> bool {
    let in_a_table = sheet.tables.iter().any(|table| {
        table.header_rows > 0
            && row >= table.start_row
            && row < table.start_row + table.header_rows
            && column >= table.start_col
            && column <= table.end_col
    });
    let filtered = sheet.auto_filter.as_ref().is_some_and(|filter| {
        row == filter.start_row && column >= filter.start_col && column <= filter.end_col
    });
    in_a_table || filtered
}

pub(crate) fn dressed_by_table(sheet: &Sheet, row: u32, column: u32) -> Option<Dressed> {
    let table = sheet.tables.iter().find(|table| {
        row >= table.start_row
            && row <= table.end_row
            && column >= table.start_col
            && column <= table.end_col
    })?;
    if row < table.start_row + table.header_rows {
        return Some(Dressed {
            fill: table.accent.clone(),
            white_text: true,
            bold: true,
        });
    }
    if !table.banded_rows {
        return None;
    }
    // The first row below the header is banded, then every other one.
    let below = row - (table.start_row + table.header_rows);
    Some(Dressed {
        fill: if below % 2 == 0 { table.band.clone() } else { None },
        white_text: false,
        bold: false,
    })
}

/// What a merge does to a cell: the top-left one is drawn across the whole
/// block, and the rest are not drawn at all.
pub(crate) enum Merged {
    /// Drawn across this many further columns and rows.
    Anchor { columns: u32, rows: u32 },
    Covered,
}

pub(crate) fn merges(sheet: &Sheet) -> std::collections::HashMap<(u32, u32), Merged> {
    let mut held = std::collections::HashMap::new();
    for merge in &sheet.merge_cells {
        for row in merge.start_row..=merge.end_row {
            for column in merge.start_col..=merge.end_col {
                let entry = if row == merge.start_row && column == merge.start_col {
                    Merged::Anchor {
                        columns: merge.end_col - merge.start_col,
                        rows: merge.end_row - merge.start_row,
                    }
                } else {
                    Merged::Covered
                };
                held.insert((row, column), entry);
            }
        }
    }
    held
}

/// Whether a cell holds nothing, so a neighbour's text may run over it.
fn is_free(
    row: &Row,
    column: u32,
    merged: &std::collections::HashMap<(u32, u32), Merged>,
) -> bool {
    if merged.contains_key(&(row.index, column)) {
        return false;
    }
    !row.cells.iter().any(|cell| {
        cell.col == column && !cell_text(&cell.value, &cell.style).is_empty()
    })
}

/// How far left of its own cell a run-on may reach, in pixels, stopping at the
/// first neighbour that holds something.
pub(crate) fn room_before(
    layout: &Geometry,
    row: &Row,
    column: u32,
    wanted: i32,
    merged: &std::collections::HashMap<(u32, u32), Merged>,
) -> i32 {
    let mut room = 0;
    let mut at = column;
    while room < wanted && at > layout.first_column {
        at -= 1;
        if !is_free(row, at, merged) {
            break;
        }
        let index = (at - layout.first_column) as usize;
        let (Some(left), Some(right)) =
            (layout.columns.get(index), layout.columns.get(index + 1))
        else {
            break;
        };
        room += (right - left) as i32;
    }
    room.min(wanted)
}

/// The same, rightward.
pub(crate) fn room_after(
    layout: &Geometry,
    row: &Row,
    column: u32,
    wanted: i32,
    merged: &std::collections::HashMap<(u32, u32), Merged>,
) -> i32 {
    let mut room = 0;
    let mut at = column;
    loop {
        if room >= wanted {
            break;
        }
        at += 1;
        let index = (at - layout.first_column) as usize;
        let (Some(left), Some(right)) =
            (layout.columns.get(index), layout.columns.get(index + 1))
        else {
            break;
        };
        if !is_free(row, at, merged) {
            break;
        }
        room += (right - left) as i32;
    }
    room.min(wanted)
}

/// Which pixels of a broken rule Excel inks.
///
/// Measured by `tools\metrics\_xlsx_border_pattern.py` and
/// `_xlsx_border_phase.py` on a grid of boxes at deliberately mixed odd and
/// even boundaries — 912 pixels a style, no exception. The short patterns turn
/// out to be halftones pinned to the picture: a hair rule is inked wherever
/// `x + y` is even, so the pattern runs on through the cells and two cells
/// sharing a boundary paint it the same way round. The long ones are counted
/// along the rule instead, from the picture's edge, and a two-pixel rule
/// carries the same phase on both of its rows.
#[derive(Clone, Copy)]
pub(crate) enum Broken {
    /// Every pixel of the rule.
    Whole,
    /// Inked where the picture coordinates say so.
    Halftone(fn(i32, i32) -> bool),
    /// Inked by the distance along the rule, modulo a period.
    Along { period: i32, inked: fn(i32) -> bool },
}

impl Broken {
    /// Is the pixel at `(x, y)`, `along` pixels down the rule, inked?
    pub(crate) fn inked(self, x: i32, y: i32, along: i32) -> bool {
        match self {
            Broken::Whole => true,
            Broken::Halftone(pattern) => pattern(x, y),
            Broken::Along { period, inked } => inked(along.rem_euclid(period)),
        }
    }
}

/// Which side of the dotted halftone a coordinate falls on: `0` and `3` of
/// every four go together, `1` and `2` go together, and a pixel is inked where
/// the two coordinates agree.
fn dotted_half(coordinate: i32) -> i32 {
    i32::from(matches!(coordinate.rem_euclid(4), 1 | 2))
}

/// How Excel draws one kind of rule, measured off a worksheet holding one of
/// each: how far the ink reaches either side of the boundary, and which pixels
/// along it are inked.
pub(crate) struct Rule {
    /// Pixels before the boundary the ink starts at, and after it ends.
    before: i32,
    after: i32,
    /// The pixel at the boundary itself is skipped by a double rule.
    hollow: bool,
    pub(crate) broken: Broken,
}

pub(crate) fn rule_for(kind: &str) -> Rule {
    let solid = |before, after| Rule { before, after, hollow: false, broken: Broken::Whole };
    let broken = |before, broken| Rule { before, after: 0, hollow: false, broken };
    match kind {
        "medium" => solid(1, 0),
        "thick" => solid(1, 1),
        "double" => Rule { before: 1, after: 1, hollow: true, broken: Broken::Whole },
        // A checkerboard.
        "hair" => broken(0, Broken::Halftone(|x, y| (x + y) % 2 == 0)),
        "dotted" => broken(0, Broken::Halftone(|x, y| dotted_half(x) == dotted_half(y))),
        // Three of every four, on the diagonal like the shorter patterns.
        "dashed" => broken(0, Broken::Halftone(|x, y| (x + y).rem_euclid(4) != 3)),
        // Nine on, three off, and the dots are three long as well.
        "mediumDashed" => broken(1, Broken::Along { period: 12, inked: |at| at < 9 }),
        "dashDot" | "slantDashDot" => {
            broken(0, Broken::Along { period: 18, inked: |at| at < 9 || (12..15).contains(&at) })
        }
        "mediumDashDot" => {
            broken(1, Broken::Along { period: 18, inked: |at| at < 9 || (12..15).contains(&at) })
        }
        "dashDotDot" => broken(0, Broken::Along {
            period: 24,
            inked: |at| at < 9 || (12..15).contains(&at) || (18..21).contains(&at),
        }),
        "mediumDashDotDot" => broken(1, Broken::Along {
            period: 24,
            inked: |at| at < 9 || (12..15).contains(&at) || (18..21).contains(&at),
        }),
        // "thin" and anything unfamiliar: a single pixel on the boundary.
        _ => solid(0, 0),
    }
}

pub(crate) fn alignment(style: &CellStyle, value: &CellValue) -> Align {
    match style.horizontal_align.as_deref() {
        Some("center") => Align::Centre,
        Some("right") => Align::Right,
        Some("left") => Align::Left,
        // A heading centred across its neighbours is centred within the cell
        // that holds the text.
        Some("centerContinuous") => Align::Centre,
        // Spread out to fill the cell, which is how Japanese workbooks set a
        // title: 第 ３ 表, not 第３表.
        Some("distributed") | Some("justify") => Align::Spread,
        _ => match value {
            CellValue::Number(_) => Align::Right,
            CellValue::Boolean(_) => Align::Centre,
            _ => Align::Left,
        },
    }
}

#[derive(Clone, Copy, PartialEq)]
pub(crate) enum Align {
    Left,
    Centre,
    Right,
    /// Spread across the width of the cell, letters and all.
    Spread,
}

#[cfg(windows)]
mod dwrite_draw;

#[cfg(windows)]
mod windows_draw {
    use super::{alignment, cell_text, Align, Geometry, Merged};
    use oxicells_core::ir::{BorderLine, CellStyle, CellValue, DrawingKind, Sheet};
    use windows::core::PCWSTR;
    use windows::Win32::Foundation::{COLORREF, POINT, RECT, SIZE};
    use windows::Win32::Graphics::Gdi::*;

    fn wide(text: &str) -> Vec<u16> {
        text.encode_utf16().chain(std::iter::once(0)).collect()
    }

    fn colour(hex: Option<&str>, fallback: u32) -> COLORREF {
        let value = hex
            .and_then(|hex| u32::from_str_radix(hex.trim_start_matches('#'), 16).ok())
            .unwrap_or(fallback);
        // A sheet writes RRGGBB; GDI wants BGR.
        let (r, g, b) = ((value >> 16) & 0xFF, (value >> 8) & 0xFF, value & 0xFF);
        COLORREF(b << 16 | g << 8 | r)
    }

    /// Measures a digit of the workbook's standard font, which is the unit a
    /// column width is stated in.
    pub fn digit_width(style: &CellStyle) -> Option<f32> {
        unsafe {
            let dc = CreateCompatibleDC(None);
            if dc.is_invalid() {
                return None;
            }
            let points = style.font_size.unwrap_or(11.0);
            let face = wide(style.font_name.as_deref().unwrap_or("Calibri"));
            let font = CreateFontW(
                -((points * 96.0 / 72.0).round() as i32),
                0,
                0,
                0,
                400,
                0,
                0,
                0,
                DEFAULT_CHARSET.0 as u32,
                OUT_DEFAULT_PRECIS.0 as u32,
                CLIP_DEFAULT_PRECIS.0 as u32,
                ANTIALIASED_QUALITY.0 as u32,
                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                PCWSTR(face.as_ptr()),
            );
            let previous = SelectObject(dc, font);
            let mut size = SIZE::default();
            let zero = wide("0");
            let measured = GetTextExtentPoint32W(dc, &zero[..1], &mut size).as_bool();
            SelectObject(dc, previous);
            let _ = DeleteObject(font);
            let _ = DeleteDC(dc);
            measured.then_some(size.cx as f32)
        }
    }

    /// Draw a preset shape into the box its anchors give it.
    ///
    /// The corpus's 2245 shapes are 1641 lines, 453 rectangles and 82 rounded
    /// ones; the rest — braces, arrows, flowchart boxes — are left undrawn
    /// rather than drawn as something they are not.
    unsafe fn shape(
        dc: HDC,
        shape: &oxicells_core::ir::Shape,
        box_: RECT,
        // The box's width before its edges were rounded, which is what a line
        // is broken against (see `drawing_room`).
        room: Option<f32>,
        scale: f32,
        normal: Option<&(String, f32)>,
    ) {
        let rule = shape.line.as_ref().map(|line| {
            let width = ((line.width as f32 / super::EMU) * scale).round().max(1.0) as i32;
            let shade = colour(Some(&line.color), 0x0000_0000);
            // A rule wider than a pixel can only be broken by a geometric
            // pen, and the pattern is stated rather than left to GDI: OOXML's
            // presets are runs of the line's own width — a `dash` is four on
            // and three off, a `sysDot` one and one — which is not what any of
            // GDI's own dash styles draw.
            let pattern: &[u32] = match line.dash.as_deref() {
                Some("dot") => &[1, 3],
                Some("dash") => &[4, 3],
                Some("lgDash") => &[8, 3],
                Some("dashDot") => &[4, 3, 1, 3],
                Some("lgDashDot") => &[8, 3, 1, 3],
                Some("lgDashDotDot") => &[8, 3, 1, 3, 1, 3],
                Some("sysDash") => &[3, 1],
                Some("sysDot") => &[1, 1],
                Some("sysDashDot") => &[3, 1, 1, 1],
                Some("sysDashDotDot") => &[3, 1, 1, 1, 1, 1],
                _ => &[],
            };
            let pen = if pattern.is_empty() {
                CreatePen(PS_SOLID, width, shade)
            } else {
                let brush = LOGBRUSH { lbStyle: BS_SOLID, lbColor: shade, lbHatch: 0 };
                let runs: Vec<u32> = pattern
                    .iter()
                    .map(|part| (part * width.max(1) as u32).max(1))
                    .collect();
                let held = ExtCreatePen(
                    PEN_STYLE(PS_GEOMETRIC.0 | PS_USERSTYLE.0 | PS_ENDCAP_FLAT.0),
                    width.max(1) as u32,
                    &brush,
                    Some(&runs),
                );
                if held.is_invalid() {
                    CreatePen(PS_SOLID, width, shade)
                } else {
                    held
                }
            };
            (pen, width)
        });
        let held = rule.map(|(pen, _)| SelectObject(dc, pen));

        // A shape that carries an outline of its own is drawn from it rather
        // than from a preset. The points are stated in a space the path names
        // — `002`'s brace is 799 by 1377 — and are mapped onto the box the
        // anchors give the shape. Sixteen shapes across four workbooks are
        // drawn this way, one of them the brace beside `002`'s notes, and
        // before this they were drawn as nothing at all.
        if let Some(drawn) = &shape.path {
            let across = box_.right - box_.left;
            let down = box_.bottom - box_.top;
            let at = |x: i64, y: i64| POINT {
                x: box_.left + (x as f64 / drawn.across as f64 * across as f64).round() as i32,
                y: box_.top + (y as f64 / drawn.down as f64 * down as f64).round() as i32,
            };
            let brush = shape
                .fill
                .as_deref()
                .map(|fill| CreateSolidBrush(colour(Some(fill), 0x00FF_FFFF)));
            let held_brush = brush.map(|brush| SelectObject(dc, brush));
            let _ = BeginPath(dc);
            for step in &drawn.steps {
                match *step {
                    oxicells_core::ir::PathStep::MoveTo(x, y) => {
                        let point = at(x, y);
                        let _ = MoveToEx(dc, point.x, point.y, None);
                    }
                    oxicells_core::ir::PathStep::LineTo(x, y) => {
                        let point = at(x, y);
                        let _ = LineTo(dc, point.x, point.y);
                    }
                    oxicells_core::ir::PathStep::CurveTo(ax, ay, bx, by, cx, cy) => {
                        let held = [at(ax, ay), at(bx, by), at(cx, cy)];
                        let _ = PolyBezierTo(dc, &held);
                    }
                    oxicells_core::ir::PathStep::Close => {
                        let _ = CloseFigure(dc);
                    }
                }
            }
            let _ = EndPath(dc);
            // Painted, ruled, or both. A path with neither is nothing to draw
            // rather than a black outline, which is what `FillPath` on its
            // own would give it.
            match (brush.is_some(), rule.is_some()) {
                (true, true) => {
                    let _ = StrokeAndFillPath(dc);
                }
                (true, false) => {
                    let _ = FillPath(dc);
                }
                (false, true) => {
                    let _ = StrokePath(dc);
                }
                (false, false) => {
                    let _ = AbortPath(dc);
                }
            }
            if let Some(held_brush) = held_brush {
                SelectObject(dc, held_brush);
            }
            if let Some(brush) = brush {
                let _ = DeleteObject(brush);
            }
        } else {

        match shape.geometry.as_str() {
            // A line runs from one corner of its box to the other, and a
            // flipped one from the corners the other way round.
            "line" | "straightConnector1" => {
                if rule.is_some() {
                    let (from_x, to_x) = if shape.flip_h {
                        (box_.right, box_.left)
                    } else {
                        (box_.left, box_.right)
                    };
                    let (from_y, to_y) = if shape.flip_v {
                        (box_.bottom, box_.top)
                    } else {
                        (box_.top, box_.bottom)
                    };
                    let _ = MoveToEx(dc, from_x, from_y, None);
                    let _ = LineTo(dc, to_x, to_y);
                }
            }
            "rect" | "roundRect" => {
                // A rounded rectangle's corner has a radius of a sixth of its
                // shorter side, which is the adjustment OOXML leaves at its
                // default. GDI asks for the whole ellipse, so twice that.
                let round = if shape.geometry == "roundRect" {
                    (box_.right - box_.left).min(box_.bottom - box_.top) / 3
                } else {
                    0
                };
                let brush = shape
                    .fill
                    .as_deref()
                    .map(|fill| CreateSolidBrush(colour(Some(fill), 0x00FF_FFFF)));
                if round == 0 {
                    if let Some(brush) = brush {
                        FillRect(dc, &box_, brush);
                    }
                    if rule.is_some() {
                        let hollow = SelectObject(dc, GetStockObject(NULL_BRUSH));
                        let _ = Rectangle(dc, box_.left, box_.top, box_.right, box_.bottom);
                        SelectObject(dc, hollow);
                    }
                } else if brush.is_some() || rule.is_some() {
                    // One call paints and rules a rounded box, so the fill
                    // stops where the rule runs rather than under it.
                    let held = SelectObject(dc, brush.map_or(GetStockObject(NULL_BRUSH), |b| b.into()));
                    let pen = if rule.is_none() {
                        Some(SelectObject(dc, GetStockObject(NULL_PEN)))
                    } else {
                        None
                    };
                    let _ = RoundRect(
                        dc, box_.left, box_.top, box_.right, box_.bottom, round, round,
                    );
                    if let Some(pen) = pen {
                        SelectObject(dc, pen);
                    }
                    SelectObject(dc, held);
                }
                if let Some(brush) = brush {
                    let _ = DeleteObject(brush);
                }
            }
            // A flowchart's decision is a diamond in its box, and Excel's
            // `ellipse` is the box's inscribed one.
            "flowChartDecision" | "diamond" | "ellipse" | "flowChartConnector" => {
                let brush = shape
                    .fill
                    .as_deref()
                    .map(|fill| CreateSolidBrush(colour(Some(fill), 0x00FF_FFFF)));
                let held_brush =
                    SelectObject(dc, brush.map_or(GetStockObject(NULL_BRUSH), |b| b.into()));
                let pen = if rule.is_none() {
                    Some(SelectObject(dc, GetStockObject(NULL_PEN)))
                } else {
                    None
                };
                if shape.geometry.starts_with("ellipse") || shape.geometry.ends_with("Connector") {
                    let _ = Ellipse(dc, box_.left, box_.top, box_.right, box_.bottom);
                } else {
                    let middle = |low: i32, high: i32| low + (high - low) / 2;
                    let points = [
                        POINT { x: middle(box_.left, box_.right), y: box_.top },
                        POINT { x: box_.right, y: middle(box_.top, box_.bottom) },
                        POINT { x: middle(box_.left, box_.right), y: box_.bottom },
                        POINT { x: box_.left, y: middle(box_.top, box_.bottom) },
                    ];
                    let _ = Polygon(dc, &points);
                }
                if let Some(pen) = pen {
                    SelectObject(dc, pen);
                }
                SelectObject(dc, held_brush);
                if let Some(brush) = brush {
                    let _ = DeleteObject(brush);
                }
            }
            _ => {}
        }

        }

        if let (Some((pen, _)), Some(held)) = (rule, held) {
            SelectObject(dc, held);
            let _ = DeleteObject(pen);
        }

        // What a shape says is drawn. It was held behind a flag for as long as
        // it cost the corpus more than it paid: a shape that says something is
        // usually a shape this did not draw properly yet, and text landing on
        // geometry that is missing or misplaced reads worse than no text at
        // all. Four laws closed that gap — a clipped box draws only the lines
        // it has room for and anchors those, a rounded box pulls its text
        // rectangle in by the corner, a pinned pitch puts the baseline at
        // three quarters of itself less the descent that overruns a quarter
        // em, and a line's last 句読点 hangs past the end — and the corpus
        // went from 6 improved against 15 regressed to 18 against 2, 0.9859
        // -> 0.9862. The two that still lost were `tb_r8_jizensoudan` and
        // `tb_r8_youshiki`, and they were not losses of this at all: their
        // runs name the theme's font rather than a face, which went to the
        // device as a face name nothing has. Resolving it put both above
        // where they stood before the text was drawn, and nothing loses now.
        // `OXI_XLSX_NO_SHAPE_TEXT` puts it back.
        if let Some(said) = &shape.text {
            if std::env::var("OXI_XLSX_NO_SHAPE_TEXT").is_err() {
                says(
                    dc,
                    said,
                    Frame {
                        box_,
                        exact: room,
                        pull: preset_pull(&shape.geometry, box_),
                    },
                    scale,
                    normal,
                    false,
                );
            }
        }
    }

    /// A pen that draws a rule of the width and pattern OOXML states.
    ///
    /// GDI's own dash styles are not OOXML's, and a cosmetic pen is a pixel
    /// wide whatever it is told, so a broken rule wider than that is drawn
    /// with a geometric pen whose pattern is spelled out in multiples of its
    /// own width — the ratios OOXML's presets stand for.
    unsafe fn ruling_pen(shade: COLORREF, width: i32, dash: Option<&str>) -> HPEN {
        let pattern: &[u32] = match dash {
            Some("dot") => &[1, 3],
            Some("dash") => &[4, 3],
            Some("lgDash") => &[8, 3],
            Some("dashDot") => &[4, 3, 1, 3],
            Some("lgDashDot") => &[8, 3, 1, 3],
            Some("lgDashDotDot") => &[8, 3, 1, 3, 1, 3],
            Some("sysDash") => &[3, 1],
            Some("sysDot") => &[1, 1],
            Some("sysDashDot") => &[3, 1, 1, 1],
            Some("sysDashDotDot") => &[3, 1, 1, 1, 1, 1],
            _ => &[],
        };
        if pattern.is_empty() {
            return CreatePen(PS_SOLID, width, shade);
        }
        let brush = LOGBRUSH { lbStyle: BS_SOLID, lbColor: shade, lbHatch: 0 };
        let runs: Vec<u32> = pattern
            .iter()
            .map(|part| (part * width.max(1) as u32).max(1))
            .collect();
        let held = ExtCreatePen(
            PEN_STYLE(PS_GEOMETRIC.0 | PS_USERSTYLE.0 | PS_ENDCAP_FLAT.0),
            width.max(1) as u32,
            &brush,
            Some(&runs),
        );
        if held.is_invalid() {
            CreatePen(PS_SOLID, width, shade)
        } else {
            held
        }
    }

    /// Draw a graph into the box its anchors give it.
    ///
    /// Only a line chart, and only one that pins its own plot rectangle: the
    /// corpus's five charts are all of that kind, and where Excel places a
    /// plot itself is a separate thing to work out. Everything plotted is in
    /// the part — a chart caches the numbers beside the cells it read them
    /// from — so the picture is drawn without going back to the sheet.
    unsafe fn graph(
        dc: HDC,
        chart: &oxicells_core::ir::Chart,
        box_: RECT,
        scale: f32,
        normal: Option<&(String, f32)>,
    ) {
        let (across, down) = (box_.right - box_.left, box_.bottom - box_.top);
        if chart.kind != "line" || chart.series.is_empty() || across <= 0 || down <= 0 {
            return;
        }
        let Some(frame) = chart.plot else { return };

        // The chart's own ground, then the plot's.
        if let Some(fill) = &chart.fill {
            let brush = CreateSolidBrush(colour(Some(fill), 0x00FF_FFFF));
            FillRect(dc, &box_, brush);
            let _ = DeleteObject(brush);
        }
        // The fractions are cut, not rounded: measured against Excel's own
        // picture, all four edges of `311e2f9c271e_zuhyo`'s plot land a pixel
        // out when rounded and exactly when truncated.
        let plot = RECT {
            left: box_.left + (frame.x * across as f64) as i32,
            top: box_.top + (frame.y * down as f64) as i32,
            right: box_.left + ((frame.x + frame.w) * across as f64) as i32,
            bottom: box_.top + ((frame.y + frame.h) * down as f64) as i32,
        };
        if plot.right <= plot.left || plot.bottom <= plot.top {
            return;
        }
        if let Some(fill) = &chart.plot_fill {
            let brush = CreateSolidBrush(colour(Some(fill), 0x00FF_FFFF));
            FillRect(dc, &plot, brush);
            let _ = DeleteObject(brush);
        }

        let up_axis = chart.value_axis.clone().unwrap_or_default();
        let along_axis = chart.category_axis.clone().unwrap_or_default();
        let numbers: Vec<f64> = chart
            .series
            .iter()
            .flat_map(|series| series.values.iter().flatten().copied())
            .collect();
        let tall_points = (plot.bottom - plot.top) as f64 / scale as f64 * 72.0 / 96.0;
        let label_size = if up_axis.size > 0.0 { up_axis.size } else { 10.0 };
        let up = super::graph::scale(
            &numbers,
            (up_axis.min, up_axis.max, up_axis.major_unit),
            tall_points,
            label_size,
        );

        // Where each category stands. `midCat` puts the first point on the
        // axis itself and the last against the far edge; anything else puts
        // every point in the middle of a step.
        let count = chart
            .series
            .iter()
            .map(|series| series.values.len())
            .max()
            .unwrap_or(0)
            .max(chart.categories.len());
        let room = (plot.right - plot.left) as f64;
        // `crossBetween` is stated on the value axis — it says where that
        // axis crosses the other one — but what it decides is where the
        // categories stand: `midCat` puts the first on the axis itself.
        let mid_cat = up_axis
            .cross_between
            .as_deref()
            .or(along_axis.cross_between.as_deref())
            == Some("midCat");
        let across_at = |index: usize| -> i32 {
            let step = if mid_cat {
                if count > 1 {
                    room * index as f64 / (count - 1) as f64
                } else {
                    room / 2.0
                }
            } else {
                room * (index as f64 + 0.5) / count.max(1) as f64
            };
            plot.left + step.round() as i32
        };
        let up_at = |value: f64| -> i32 {
            plot.bottom - (up.at(value) * (plot.bottom - plot.top) as f64).round() as i32
        };

        // A gridline across the plot at every tick, when the axis wants one.
        if let Some(line) = &up_axis.major_gridline {
            let width = ((line.width as f32 / super::EMU) * scale).round().max(1.0) as i32;
            let pen = ruling_pen(colour(Some(&line.color), 0x00D9_D9D9), width, line.dash.as_deref());
            let held = SelectObject(dc, pen);
            for step in 0..=steps(&up) {
                let at = up_at(up.low + step as f64 * up.unit);
                let _ = MoveToEx(dc, plot.left, at, None);
                let _ = LineTo(dc, plot.right, at);
            }
            SelectObject(dc, held);
            let _ = DeleteObject(pen);
        }

        // Every series, in the order the chart states them.
        for series in &chart.series {
            let line = series.line.clone().unwrap_or(oxicells_core::ir::ShapeLine {
                color: "000000".into(),
                width: 9525,
                dash: None,
            });
            let width = ((line.width as f32 / super::EMU) * scale).round().max(1.0) as i32;
            let pen = ruling_pen(
                colour(Some(&line.color), 0x0000_0000),
                width,
                line.dash.as_deref(),
            );
            let held = SelectObject(dc, pen);
            // A gap in the data breaks the line rather than being read as a
            // zero, which is what `dispBlanksAs="gap"` asks for.
            let mut drawing = false;
            for (index, value) in series.values.iter().enumerate() {
                match value {
                    Some(value) => {
                        let (x, y) = (across_at(index), up_at(*value));
                        if drawing {
                            let _ = LineTo(dc, x, y);
                        } else {
                            let _ = MoveToEx(dc, x, y, None);
                            drawing = true;
                        }
                    }
                    None => drawing = false,
                }
            }
            SelectObject(dc, held);
            let _ = DeleteObject(pen);

            // What the series wears at every point, and what single points
            // wear instead.
            for (index, value) in series.values.iter().enumerate() {
                let Some(value) = value else { continue };
                let own = series
                    .points
                    .iter()
                    .find(|point| point.index as usize == index)
                    .and_then(|point| point.marker.as_ref());
                let Some(marker) = own.or(series.marker.as_ref()) else {
                    continue;
                };
                if marker.symbol.is_empty() || marker.symbol == "none" {
                    continue;
                }
                mark(dc, marker, across_at(index), up_at(*value), scale);
            }
        }

        // The axes, over the plot and under nothing.
        let axis_pen = |line: &Option<oxicells_core::ir::ShapeLine>| {
            let stated = line.clone().unwrap_or(oxicells_core::ir::ShapeLine {
                color: "000000".into(),
                width: 3175,
                dash: None,
            });
            let width = ((stated.width as f32 / super::EMU) * scale).round().max(1.0) as i32;
            (
                ruling_pen(colour(Some(&stated.color), 0x0000_0000), width, None),
                width,
            )
        };
        // Excel's tick marks are four pixels long at 96 dpi.
        let tick = (4.0 * scale).round().max(1.0) as i32;
        let foot = up_at(up.low.max(0.0).min(up.high));

        let (pen, _) = axis_pen(&along_axis.line);
        let mut held = SelectObject(dc, pen);
        if !along_axis.deleted {
            let _ = MoveToEx(dc, plot.left, foot, None);
            let _ = LineTo(dc, plot.right, foot);
            if along_axis.major_tick != "none" && !along_axis.major_tick.is_empty() {
                for index in 0..count {
                    // A tick stands between two categories when the points
                    // do, and under the point itself when they do not.
                    let at = if mid_cat {
                        across_at(index)
                    } else {
                        plot.left + (room * index as f64 / count as f64).round() as i32
                    };
                    let (from, to) = match along_axis.major_tick.as_str() {
                        "in" => (foot - tick, foot),
                        "out" => (foot, foot + tick),
                        _ => (foot - tick, foot + tick),
                    };
                    let _ = MoveToEx(dc, at, from, None);
                    let _ = LineTo(dc, at, to);
                }
            }
        }
        SelectObject(dc, held);
        let _ = DeleteObject(pen);

        let (pen, _) = axis_pen(&up_axis.line);
        held = SelectObject(dc, pen);
        if !up_axis.deleted {
            let _ = MoveToEx(dc, plot.left, plot.top, None);
            let _ = LineTo(dc, plot.left, plot.bottom);
            if up_axis.major_tick != "none" && !up_axis.major_tick.is_empty() {
                for step in 0..=steps(&up) {
                    let at = up_at(up.low + step as f64 * up.unit);
                    let (from, to) = match up_axis.major_tick.as_str() {
                        "in" => (plot.left, plot.left + tick),
                        "out" => (plot.left - tick, plot.left),
                        _ => (plot.left - tick, plot.left + tick),
                    };
                    let _ = MoveToEx(dc, from, at, None);
                    let _ = LineTo(dc, to, at);
                }
            }
        }
        SelectObject(dc, held);
        let _ = DeleteObject(pen);

        // What the axes are labelled with. A value's label is set against the
        // axis itself — the room between them is the glyph's own bearing —
        // and a category's stands eight pixels below it. Measured off Excel's
        // picture of `311e2f9c271e_zuhyo` and `744b4e4a4cfd_zuhyo`.
        let gap = (8.0 * scale).round() as i32;
        let face = |named: &Option<String>| {
            named
                .clone()
                .or_else(|| normal.map(|(face, _)| face.clone()))
                .unwrap_or_else(|| "ＭＳ Ｐゴシック".to_string())
        };
        if !up_axis.deleted && up_axis.tick_labels != "none" {
            let font = chart_font(&face(&up_axis.face), label_size, scale);
            let held = SelectObject(dc, font);
            SetTextAlign(dc, TA_TOP | TA_RIGHT);
            for step in 0..=steps(&up) {
                let value = up.low + step as f64 * up.unit;
                let said = oxicells_core::format_number(
                    value,
                    up_axis.number_format.as_deref().unwrap_or("General"),
                );
                let letters = wide(said.trim());
                let letters = &letters[..letters.len() - 1];
                if letters.is_empty() {
                    continue;
                }
                let mut measured = SIZE::default();
                let _ = GetTextExtentPoint32W(dc, letters, &mut measured);
                let _ = TextOutW(dc, plot.left, up_at(value) - measured.cy / 2, letters);
            }
            SelectObject(dc, held);
            let _ = DeleteObject(font);
        }
        if !along_axis.deleted && along_axis.tick_labels != "none" {
            let size = if along_axis.size > 0.0 { along_axis.size } else { 10.0 };
            let named = face(&along_axis.face);
            let font = chart_font(&named, size, scale);
            let held = SelectObject(dc, font);
            SetTextAlign(dc, TA_TOP | TA_CENTER);
            // A label wider than the step it stands under is broken to fit,
            // which is what stacks `昭和51` as three lines under the first
            // category of the corpus's charts.
            let step = if count > 1 {
                (room / if mid_cat { (count - 1) as f64 } else { count as f64 }) as f32
            } else {
                room as f32
            };
            let pitch = super::line_box(&named, size, false, false)
                .map(|(tall, _)| tall * scale)
                .unwrap_or(size * scale * 96.0 / 72.0 * 1.3);
            for (index, said) in chart.categories.iter().enumerate() {
                let mut at = foot + gap;
                for line in super::wrapped_lines(&named, size, false, false, said, Some(step / scale))
                {
                    let letters = wide(&line);
                    let letters = &letters[..letters.len() - 1];
                    if !letters.is_empty() {
                        let _ = TextOutW(dc, across_at(index), at, letters);
                    }
                    at += pitch.round() as i32;
                }
            }
            SelectObject(dc, held);
            let _ = DeleteObject(font);
        }

        // A number written beside the point it belongs to. Which side it is
        // set against is `dLblPos`; how far it is then moved is a fraction of
        // the chart's own box. Measured on `744b4e4a4cfd_zuhyo`'s one visible
        // label: set to the right, it is level with its point and clear of it
        // by the marker's own half-width and four points more.
        SetTextAlign(dc, TA_TOP | TA_LEFT);
        for series in &chart.series {
            for label in &series.labels {
                let Some(Some(value)) = series.values.get(label.index as usize) else {
                    continue;
                };
                let said = label.text.clone().unwrap_or_else(|| {
                    oxicells_core::format_number(
                        *value,
                        label.number_format.as_deref().unwrap_or("General"),
                    )
                });
                let letters = wide(said.trim());
                let letters = &letters[..letters.len() - 1];
                if letters.is_empty() {
                    continue;
                }
                let size = [label.size, series.label_size, 10.0]
                    .into_iter()
                    .find(|points| *points > 0.0)
                    .unwrap_or(10.0);
                let named = match (&label.face, &series.label_face) {
                    (Some(own), _) => face(&Some(own.clone())),
                    (None, held) => face(held),
                };
                let font = chart_font(&named, size, scale);
                let held = SelectObject(dc, font);
                let mut measured = SIZE::default();
                let _ = GetTextExtentPoint32W(dc, letters, &mut measured);

                let (x, y) = (across_at(label.index as usize), up_at(*value));
                // Half the mark the point wears, so the label clears it.
                let worn = series
                    .points
                    .iter()
                    .find(|point| point.index as usize == label.index as usize)
                    .and_then(|point| point.marker.as_ref())
                    .or(series.marker.as_ref())
                    .filter(|marker| !marker.symbol.is_empty() && marker.symbol != "none")
                    .map_or(0.0, |marker| marker.size as f32 * 96.0 / 72.0 / 2.0);
                let clear = ((worn + 4.0 * 96.0 / 72.0) * scale).round() as i32;
                let (mut left, mut top) = match label
                    .position
                    .as_deref()
                    .or(series.label_pos.as_deref())
                    .unwrap_or("t")
                {
                    "r" => (x + clear, y - measured.cy / 2),
                    "l" => (x - clear - measured.cx, y - measured.cy / 2),
                    "b" => (x - measured.cx / 2, y + clear),
                    "ctr" => (x - measured.cx / 2, y - measured.cy / 2),
                    _ => (x - measured.cx / 2, y - clear - measured.cy),
                };
                let nudge = label.offset.unwrap_or((0.0, 0.0));
                left += (nudge.0 * across as f64).round() as i32;
                top += (nudge.1 * down as f64).round() as i32;
                let _ = TextOutW(dc, left, top, letters);
                SelectObject(dc, held);
                let _ = DeleteObject(font);
            }
        }

        // The legend: a sample of each series' rule with its name beside it,
        // one under the next down the box the chart gives them.
        if let Some(legend) = &chart.legend {
            if let Some(frame) = legend.frame {
                let box_of = RECT {
                    left: box_.left + (frame.x * across as f64).round() as i32,
                    top: box_.top + (frame.y * down as f64).round() as i32,
                    right: box_.left + ((frame.x + frame.w) * across as f64).round() as i32,
                    bottom: box_.top + ((frame.y + frame.h) * down as f64).round() as i32,
                };
                let size = if legend.size > 0.0 { legend.size } else { 10.0 };
                let named = face(&legend.face);
                let font = chart_font(&named, size, scale);
                let held = SelectObject(dc, font);
                SetTextAlign(dc, TA_TOP | TA_LEFT);
                let line_box = super::line_box(&named, size, false, false)
                    .map(|(tall, _)| tall * scale)
                    .unwrap_or(size * scale * 96.0 / 72.0 * 1.3);

                // A sample of the rule is 26.75pt long whatever the chart's
                // size, and what an entry says follows it with nothing
                // between. What is left over in the box is split evenly: one
                // share after each entry, and a leading share three and three
                // quarter points wider before the first. Measured through
                // `LegendEntry` in `_xlsx_chart_legend.py` — box widths of
                // 151, 181 and 289 points all give back a 26.75pt sample.
                let sample = (26.75 * scale * 96.0 / 72.0).round() as i32;
                let mut said: Vec<(String, i32)> = Vec::new();
                for series in &chart.series {
                    let letters = wide(&series.name);
                    let mut measured = SIZE::default();
                    if letters.len() > 1 {
                        let _ = GetTextExtentPoint32W(
                            dc,
                            &letters[..letters.len() - 1],
                            &mut measured,
                        );
                    }
                    said.push((series.name.clone(), measured.cx));
                }
                // As many entries to a row as their samples and names will
                // fit, and the rows spread evenly down the box.
                let room = box_of.right - box_of.left;
                let mut rows: Vec<Vec<usize>> = vec![Vec::new()];
                let mut used = 0;
                for (index, (_, width)) in said.iter().enumerate() {
                    let wanted = sample + width;
                    if used + wanted > room && !rows.last().unwrap().is_empty() {
                        rows.push(Vec::new());
                        used = 0;
                    }
                    rows.last_mut().unwrap().push(index);
                    used += wanted;
                }
                let pitch = (box_of.bottom - box_of.top) as f32 / rows.len() as f32;
                let lead = (3.75 * scale * 96.0 / 72.0).round() as i32;
                for (row, entries) in rows.iter().enumerate() {
                    let middle =
                        box_of.top + (pitch * (row as f32 + 0.5)).round() as i32;
                    let taken: i32 = entries
                        .iter()
                        .map(|index| sample + said[*index].1)
                        .sum::<i32>();
                    let share = ((room - lead - taken) / (entries.len() as i32 + 1)).max(0);
                    let mut at = box_of.left + share + lead;
                    for index in entries {
                        let series = &chart.series[*index];
                        if let Some(line) = &series.line {
                            let width = ((line.width as f32 / super::EMU) * scale)
                                .round()
                                .max(1.0) as i32;
                            let pen = ruling_pen(
                                colour(Some(&line.color), 0x0000_0000),
                                width,
                                line.dash.as_deref(),
                            );
                            let held = SelectObject(dc, pen);
                            let _ = MoveToEx(dc, at, middle, None);
                            let _ = LineTo(dc, at + sample, middle);
                            SelectObject(dc, held);
                            let _ = DeleteObject(pen);
                        }
                        let letters = wide(&said[*index].0);
                        let letters = &letters[..letters.len() - 1];
                        if !letters.is_empty() {
                            let _ = TextOutW(
                                dc,
                                at + sample,
                                middle - (line_box / 2.0).round() as i32,
                                letters,
                            );
                        }
                        at += sample + said[*index].1 + share;
                    }
                }
                SelectObject(dc, held);
                let _ = DeleteObject(font);
            }
        }
        SetTextAlign(dc, TA_TOP | TA_LEFT);

        // What the chart is annotated with, over everything it plots. These
        // are text boxes in a part of their own — the corpus's charts keep
        // their footnotes there rather than in a cell — so what they say is
        // drawn whether or not a sheet's own shapes are.
        for drawn in &chart.shapes {
            let (Some(frame), oxicells_core::ir::DrawingKind::Shape(held)) =
                (drawn.frame, &drawn.kind)
            else {
                continue;
            };
            let over = RECT {
                left: box_.left + (frame.x * across as f64).round() as i32,
                top: box_.top + (frame.y * down as f64).round() as i32,
                right: box_.left + ((frame.x + frame.w) * across as f64).round() as i32,
                bottom: box_.top + ((frame.y + frame.h) * down as f64).round() as i32,
            };
            shape(dc, held, over, None, scale, normal);
            if let Some(said) = &held.text {
                says(
                    dc,
                    said,
                    Frame { box_: over, exact: None, pull: 0.0 },
                    scale,
                    normal,
                    false,
                );
            }
        }
    }

    /// How many ticks stand between the ends of a scale.
    fn steps(scale: &super::graph::Scale) -> i32 {
        if scale.unit <= 0.0 {
            return 0;
        }
        (((scale.high - scale.low) / scale.unit).floor() as i32).clamp(0, 1000)
    }

    unsafe fn chart_font(face: &str, points: f32, scale: f32) -> HFONT {
        let named = wide(face);
        CreateFontW(
            -((points * scale * 96.0 / 72.0).round() as i32),
            0,
            0,
            0,
            400,
            0,
            0,
            0,
            DEFAULT_CHARSET.0 as u32,
            OUT_DEFAULT_PRECIS.0 as u32,
            CLIP_DEFAULT_PRECIS.0 as u32,
            CLEARTYPE_QUALITY.0 as u32,
            (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
            PCWSTR(named.as_ptr()),
        )
    }

    /// The mark a series wears at one of its points.
    unsafe fn mark(
        dc: HDC,
        marker: &oxicells_core::ir::Marker,
        x: i32,
        y: i32,
        scale: f32,
    ) {
        // A marker's size is stated in points across.
        let half = ((marker.size.max(2) as f32 * scale * 96.0 / 72.0) / 2.0).round() as i32;
        let brush = CreateSolidBrush(colour(marker.fill.as_deref(), 0x00FF_FFFF));
        let pen = CreatePen(PS_SOLID, scale.round().max(1.0) as i32,
            colour(marker.line.as_deref().or(marker.fill.as_deref()), 0x0000_0000));
        let held_brush = SelectObject(dc, brush);
        let held_pen = SelectObject(dc, pen);
        match marker.symbol.as_str() {
            "circle" | "dot" => {
                let _ = Ellipse(dc, x - half, y - half, x + half, y + half);
            }
            "diamond" => {
                let points = [
                    POINT { x, y: y - half },
                    POINT { x: x + half, y },
                    POINT { x, y: y + half },
                    POINT { x: x - half, y },
                ];
                let _ = Polygon(dc, &points);
            }
            "triangle" => {
                let points = [
                    POINT { x, y: y - half },
                    POINT { x: x + half, y: y + half },
                    POINT { x: x - half, y: y + half },
                ];
                let _ = Polygon(dc, &points);
            }
            // A square, and anything else Excel draws as one.
            _ => {
                let _ = Rectangle(dc, x - half, y - half, x + half, y + half);
            }
        }
        SelectObject(dc, held_pen);
        SelectObject(dc, held_brush);
        let _ = DeleteObject(pen);
        let _ = DeleteObject(brush);
    }

    /// Write what a shape says inside it.
    ///
    /// The text sits in the shape's box less the insets the body states —
    /// Excel's own are a tenth of an inch either side and a twentieth above
    /// and below — wrapped in what is left, and the block of lines is put
    /// against the top, the middle or the foot by the body's anchor.
    /// What a preset's own text rectangle pulls in from each edge of the box.
    ///
    /// A rounded rectangle's text does not start at its edge: the preset's
    /// `rect` is inset by `(1 - cos 45°)` of the corner radius, which is
    /// `min(w, h) × adj / 100000` with `adj` 16667 when the file leaves
    /// `avLst` empty — every workbook in the corpus does. Measured against a
    /// plain rectangle at nine box heights by `_xlsx_shape_overflow.py`: the
    /// two differ by 9, 8, 7, 6, 5, 5, 4, 3, 3 pixels where this arithmetic
    /// says 9.11, 7.81, 6.51, 5.86, 5.21, 4.56, 3.91, 3.25, 2.60.
    fn preset_pull(geometry: &str, box_: RECT) -> f32 {
        if geometry != "roundRect" {
            return 0.0;
        }
        let across = (box_.right - box_.left) as f32;
        let down = (box_.bottom - box_.top) as f32;
        across.min(down) * 0.166_67 * 0.292_89
    }

    /// The box a shape's text is set in: the rectangle it is drawn against,
    /// the exact width a line breaks against before that rectangle was put on
    /// whole pixels, and what the preset's own text rectangle pulls in from
    /// every edge.
    struct Frame {
        box_: RECT,
        exact: Option<f32>,
        pull: f32,
    }

    unsafe fn says(
        dc: HDC,
        said: &oxicells_core::ir::ShapeText,
        frame: Frame,
        scale: f32,
        normal: Option<&(String, f32)>,
        // A note is laid out by the engine that lays out cells, not the one
        // that lays out shapes: its lines are a cell's line box apart —
        // メイリオ 14pt comes out 30.5px in `002`'s note where the same face
        // and size in a shape is 36.5.
        note: bool,
    ) {
        let Frame { box_, exact, pull } = frame;
        let inset = |emu: i64| (emu as f32 / super::EMU * scale).round() as i32;
        let pulled = pull.round() as i32;
        let area = RECT {
            left: box_.left + inset(said.insets.0) + pulled,
            top: box_.top + inset(said.insets.1) + pulled,
            right: box_.right - inset(said.insets.2) - pulled,
            bottom: box_.bottom - inset(said.insets.3) - pulled,
        };
        if area.right <= area.left {
            return;
        }
        // The insets keep their own fractions here as well: Excel states them
        // in points (7.2 of them either side is 9.6 pixels, not 10), and the
        // room is what is left of the exact box after the exact insets.
        let room = match exact {
            Some(exact) => {
                (exact
                    - (said.insets.0 + said.insets.2) as f32 / super::EMU * scale
                    - 2.0 * pull)
                    / scale
            }
            None => (area.right - area.left) as f32 / scale,
        };

        // Every line, with the paragraph it belongs to, so the block can be
        // measured before any of it is written.
        // Each line, the paragraph it belongs to, and where in that
        // paragraph's characters it starts — which is what says which runs it
        // is written in.
        let mut lines: Vec<(usize, String, usize)> = Vec::new();
        let mut pitch: Vec<f32> = Vec::new();
        let mut leading: Vec<i32> = Vec::new();
        for (index, paragraph) in said.paragraphs.iter().enumerate() {
            // A face this machine has not got is not GDI's business to
            // guess: Excel answers by the run's charset (see `face_in_place`).
            let face = super::face_in_place(
                &paragraph
                    .face
                    .clone()
                    .or_else(|| normal.map(|(face, _)| face.clone()))
                    .unwrap_or_else(|| "ＭＳ Ｐゴシック".to_string()),
                paragraph.charset,
            );
            let broken = if note {
                super::wrapped_lines(
                    &face,
                    paragraph.size,
                    paragraph.bold,
                    paragraph.italic,
                    &paragraph.text,
                    said.wrap.then_some(room),
                )
            } else {
                super::wrapped_shape_lines(
                    &face,
                    paragraph.size,
                    paragraph.bold,
                    paragraph.italic,
                    &paragraph.text,
                    said.wrap.then_some(room),
                )
            };
            // The glyphs sit in the middle of the pitch, not against its top:
            // a Japanese face's extra three tenths of leading falls half above
            // the letters and half below, which is what puts `002`'s panel
            // four pixels down from its inset.
            let measured = if note {
                super::line_box(&face, paragraph.size, paragraph.bold, paragraph.italic)
                    .map(|(tall, _)| (tall, tall))
            } else {
                super::shape_line(&face, paragraph.size, paragraph.bold, paragraph.italic)
            };
            let (tall, natural) = match (paragraph.line_pitch, measured) {
                (Some(points), found) => (
                    points * 96.0 / 72.0,
                    found.map_or(points * 96.0 / 72.0, |(_, natural)| natural),
                ),
                (None, Some((tall, natural))) => (tall, natural),
                (None, None) => {
                    let em = paragraph.size * 96.0 / 72.0;
                    (em * 1.3, em)
                }
            };
            // A paragraph that asks for a share of the font's own pitch takes
            // it from the whole box: `glossary_05`'s flowchart sets every one
            // of its boxes at four fifths, which is 22 pixels a line where
            // Yu Gothic UI's own is 28.
            let tall = tall * paragraph.line_scale.unwrap_or(1.0);
            // A paragraph that pins its pitch outright does NOT centre its
            // glyphs in it: Excel puts the baseline three quarters of the way
            // down the pinned pitch, and then lifts the line by however much
            // the face's descent overruns a quarter of the em.
            //
            //     baseline = line top + 0.75 × pitch - max(0, descent - em/4)
            //
            // `_xlsx_shape_pitch.py` sweeps eight pinned pitches (12 to 33
            // point) over four faces at ten sizes. The slope is 0.750 for
            // every one of them; only the intercept moves, and it is dead
            // constant across all eight pitches of a face (メイリオ at 20
            // point gives 17 six times running). The lift accounts for every
            // intercept measured: メイリオ 5-3, 7-3.75, 8-4.75, 12-6.75 give
            // 2, 3, 3, 5 against the measured -2, -3, -3, -5; 游ゴシック's
            // 1.25 gives 1; Meiryo UI and ＭＳ Ｐゴシック go negative and are
            // held at 0 — which is why the earlier reading, that internal
            // leading was the discriminator, was wrong: Meiryo UI carries
            // leading and does not move. Ten of ten.
            let pinned = paragraph.line_pitch.and_then(|_| {
                super::held(|counter| {
                    counter.shape_of(&face, paragraph.size, paragraph.bold, paragraph.italic)
                })
            });
            if std::env::var("OXI_XLSX_DUMP_LINES").is_ok() {
                for line in &broken {
                    let run = super::shape_run(
                        &face, paragraph.size, paragraph.bold, paragraph.italic, line,
                    )
                    .map(|steps| steps.iter().sum::<i32>());
                    eprintln!(
                        "shape line room {room:.2} run {run:?} chars {} {:?}",
                        line.chars().count(),
                        line.chars().take(14).collect::<String>()
                    );
                }
            }
            // The lines are the paragraph's text with only its newlines
            // taken out, so walking them forward gives each its own start.
            let letters: Vec<char> = paragraph.text.chars().collect();
            let mut cursor = 0usize;
            for line in broken {
                let start = cursor;
                cursor += line.chars().count();
                if letters.get(cursor) == Some(&'\n') {
                    cursor += 1;
                }
                lines.push((index, line, start));
                pitch.push(tall * scale);
                leading.push(match pinned {
                    Some((ascent, descent, _)) => {
                        let em = paragraph.size * 96.0 / 72.0;
                        let lift = (descent - em / 4.0).max(0.0).floor();
                        // The ink of a pinned line may well start above the
                        // line: a pitch smaller than the face asks for is
                        // exactly what a pinned pitch is usually for.
                        ((0.75 * tall - lift - ascent) * scale).round() as i32
                    }
                    None => (((tall - natural) / 2.0) * scale).round().max(0.0) as i32,
                });
            }
        }
        if lines.is_empty() {
            return;
        }

        // A block taller than its box is not centred and left to hang out of
        // both ends: Excel draws the lines that FIT and anchors those, and
        // drops the rest. `_xlsx_shape_overflow.py` sweeps a box from
        // comfortably taller than the text down to half its height, over
        // t/ctr/b and over a plain rectangle as well as a rounded one — 54
        // arms, every one of them the count `floor(room / pitch)` and the
        // shortened block anchored in the usual way. Against a plain
        // rectangle the top-anchored first line does not move at all (145 at
        // every one of nine heights), which is what says the box, not the
        // block, is what gets cut. Centring the whole block instead put
        // `sanko_tool`'s ten lines a whole pitch above Excel's.
        // Only a body that says so, though: `_xlsx_shape_clip.py` draws the
        // same five lines in the same box three times over, with
        // `vertOverflow` written clip, left out, and written overflow. The
        // clipped one holds four lines; the other two hold all five, the
        // fifth hanging below the box. So the dropping is what `clip` means,
        // not what a box does.
        if said.clip {
            let mut fits = 0usize;
            let mut down = 0.0f32;
            let room_down = (area.bottom - area.top) as f32;
            for step in &pitch {
                if fits > 0 && down + step > room_down + 0.01 {
                    break;
                }
                down += step;
                fits += 1;
            }
            let fits = fits.clamp(1, lines.len());
            lines.truncate(fits);
            pitch.truncate(fits);
            leading.truncate(fits);
        }

        let block: i32 = pitch.iter().sum::<f32>().round() as i32;
        let slack = (area.bottom - area.top) - block;
        // The pitch is kept as it is measured and rounded only where a line
        // lands, so a block of many lines does not drift from Excel's.
        let mut at = (area.top
            + match said.anchor.as_deref() {
                Some("ctr") => (slack as f32 / 2.0).floor() as i32,
                Some("b") => slack,
                _ => 0,
            }) as f32;

        // A line of a shape's text is placed by its top, not its baseline.
        // Nothing follows this on the sheet, so the alignment stays as it is
        // left.
        SetTextAlign(dc, TA_TOP | TA_LEFT);
        for (step, (index, line, from)) in lines.iter().enumerate() {
            let paragraph = &said.paragraphs[*index];
            // A face this machine has not got is not GDI's business to
            // guess: Excel answers by the run's charset (see `face_in_place`).
            let face = super::face_in_place(
                &paragraph
                    .face
                    .clone()
                    .or_else(|| normal.map(|(face, _)| face.clone()))
                    .unwrap_or_else(|| "ＭＳ Ｐゴシック".to_string()),
                paragraph.charset,
            );
            let pixels = -((paragraph.size * scale * 96.0 / 72.0).round() as i32);
            let named = wide(&face);
            let dressed = |bold: bool, underline: bool| {
                CreateFontW(
                    pixels,
                    0,
                    0,
                    0,
                    if bold { 700 } else { 400 },
                    u32::from(paragraph.italic),
                    u32::from(underline),
                    0,
                    DEFAULT_CHARSET.0 as u32,
                    OUT_DEFAULT_PRECIS.0 as u32,
                    CLIP_DEFAULT_PRECIS.0 as u32,
                    CLEARTYPE_QUALITY.0 as u32,
                    (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                    PCWSTR(named.as_ptr()),
                )
            };
            // The pieces this line is written in: a paragraph written all one
            // way gives one, and one behaves exactly as the whole line did.
            // `u="sng"` and a weight sit on the RUN, and `sanko_tool` holds an
            // underlined heading, two breaks and then its body inside a single
            // paragraph — so wearing the first run's dressing across the
            // paragraph underlines the body too.
            let mut worn: Vec<(bool, bool, Option<String>, String)> = Vec::new();
            {
                let want = *from..*from + line.chars().count();
                let mut walked = 0usize;
                for run in &paragraph.runs {
                    let len = run.text.chars().count();
                    let start = walked.max(want.start);
                    let stop = (walked + len).min(want.end);
                    if start < stop {
                        let piece: String = run
                            .text
                            .chars()
                            .skip(start - walked)
                            .take(stop - start)
                            .filter(|letter| *letter != '\n')
                            .collect();
                        if !piece.is_empty() {
                            worn.push((
                                run.bold,
                                run.underline,
                                run.color.clone().or_else(|| paragraph.color.clone()),
                                piece,
                            ));
                        }
                    }
                    walked += len;
                }
            }
            if worn.is_empty() {
                worn.push((paragraph.bold, false, paragraph.color.clone(), line.clone()));
            }
            let font = dressed(worn[0].0, worn[0].1);
            let held = SelectObject(dc, font);
            let letters = wide(line);
            let letters = &letters[..letters.len() - 1];
            if !letters.is_empty() {
                // Stepped by the font's own advances at the exact em, not by
                // the whole-pixel ones the device would use — but only for a
                // shape. A note is laid out by the engine that lays out
                // cells, and that one steps in whole pixels.
                let steps = (!note)
                    .then(|| {
                        super::shape_run_worn(
                            &face,
                            paragraph.size,
                            paragraph.italic,
                            &worn
                                .iter()
                                .map(|(bold, _, _, text)| (*bold, text.clone()))
                                .collect::<Vec<_>>(),
                        )
                    })
                    .flatten();
                let mut measured = SIZE::default();
                let _ = GetTextExtentPoint32W(dc, letters, &mut measured);
                let width = match &steps {
                    Some(steps) => steps.iter().sum::<i32>(),
                    None => measured.cx,
                };
                let left = match paragraph.align.as_deref() {
                    Some("ctr") => {
                        area.left + ((area.right - area.left - width) as f32 / 2.0).round() as i32
                    }
                    Some("r") => area.right - width,
                    _ => area.left,
                };
                let down = at.round() as i32 + leading[step];
                let mut x = left;
                let mut taken = 0usize;
                for (bold, underline, painted, piece) in &worn {
                    let held_piece = wide(piece);
                    let shown = &held_piece[..held_piece.len() - 1];
                    if shown.is_empty() {
                        continue;
                    }
                    let piece_font = dressed(*bold, *underline);
                    let previous = SelectObject(dc, piece_font);
                    SetTextColor(dc, colour(painted.as_deref(), 0x0000_0000));
                    let mine = steps
                        .as_ref()
                        .filter(|steps| steps.len() == letters.len())
                        .map(|steps| &steps[taken..taken + shown.len()]);
                    match mine {
                        // One `dx` a UTF-16 unit, so a character outside the
                        // basic plane carries its whole advance on the first
                        // of its two.
                        Some(mine) => {
                            let _ = ExtTextOutW(
                                dc,
                                x,
                                down,
                                ETO_OPTIONS(0),
                                None,
                                PCWSTR(shown.as_ptr()),
                                shown.len() as u32,
                                Some(mine.as_ptr()),
                            );
                            x += mine.iter().sum::<i32>();
                        }
                        None => {
                            let _ = TextOutW(dc, x, down, shown);
                            let mut walked = SIZE::default();
                            let _ = GetTextExtentPoint32W(dc, shown, &mut walked);
                            x += walked.cx;
                        }
                    }
                    taken += shown.len();
                    SelectObject(dc, previous);
                    let _ = DeleteObject(piece_font);
                }
            }
            at += pitch[step];
            SelectObject(dc, held);
            let _ = DeleteObject(font);
        }
        SetTextColor(dc, COLORREF(0x0000_0000));
    }

    /// Lay a picture into the box its anchors give it.
    ///
    /// The bytes are whatever the file holds — PNG, JPEG, GIF, EMF — and only
    /// what the image crate can decode is drawn; an EMF is left out rather
    /// than drawn wrong. GDI does the scaling, over the alpha the picture
    /// carries, which is how a logo with a transparent corner lands on the
    /// sheet's white rather than on a grey box.
    unsafe fn picture(dc: HDC, bytes: &[u8], box_: RECT) {
        let (across, down) = (box_.right - box_.left, box_.bottom - box_.top);
        if across <= 0 || down <= 0 {
            return;
        }
        // An enhanced metafile holds no pixels at all: it is a list of the
        // drawing calls that made it, which Windows can play straight into
        // the picture at whatever size the anchor gives it. Eight of the
        // corpus's workbooks put a graph on the sheet that way. Its header
        // record is a little-endian 1.
        if bytes.starts_with(&[0x01, 0x00, 0x00, 0x00]) {
            let held = SetEnhMetaFileBits(bytes);
            if !held.is_invalid() {
                let _ = PlayEnhMetaFile(dc, held, &box_);
                let _ = DeleteEnhMetaFile(held);
            }
            return;
        }
        let Ok(decoded) = image::load_from_memory(bytes) else {
            return;
        };
        let decoded = decoded.to_rgba8();
        let (wide_px, tall_px) = (decoded.width() as i32, decoded.height() as i32);
        if wide_px <= 0 || tall_px <= 0 {
            return;
        }
        // GDI blends premultiplied BGRA, top row first when the height is
        // stated negative.
        let mut pixels = Vec::with_capacity((wide_px * tall_px * 4) as usize);
        for shade in decoded.pixels() {
            let [red, green, blue, alpha] = shade.0;
            let mix = |part: u8| ((part as u16 * alpha as u16) / 255) as u8;
            pixels.extend_from_slice(&[mix(blue), mix(green), mix(red), alpha]);
        }
        let info = BITMAPINFO {
            bmiHeader: BITMAPINFOHEADER {
                biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                biWidth: wide_px,
                biHeight: -tall_px,
                biPlanes: 1,
                biBitCount: 32,
                biCompression: BI_RGB.0,
                ..Default::default()
            },
            ..Default::default()
        };
        let held = CreateCompatibleDC(dc);
        let mut bits: *mut std::ffi::c_void = std::ptr::null_mut();
        let Ok(bitmap) = CreateDIBSection(held, &info, DIB_RGB_COLORS, &mut bits, None, 0) else {
            let _ = DeleteDC(held);
            return;
        };
        std::ptr::copy_nonoverlapping(pixels.as_ptr(), bits as *mut u8, pixels.len());
        let previous = SelectObject(held, bitmap);
        let _ = AlphaBlend(
            dc,
            box_.left,
            box_.top,
            across,
            down,
            held,
            0,
            0,
            wide_px,
            tall_px,
            BLENDFUNCTION {
                BlendOp: AC_SRC_OVER as u8,
                BlendFlags: 0,
                SourceConstantAlpha: 255,
                AlphaFormat: AC_SRC_ALPHA as u8,
            },
        );
        SelectObject(held, previous);
        let _ = DeleteObject(bitmap);
        let _ = DeleteDC(held);
    }

    pub fn draw(
        sheet: &Sheet,
        layout: &Geometry,
        width: u32,
        height: u32,
        scale: f32,
    ) -> image::RgbImage {
        unsafe {
            let screen = GetDC(None);
            let dc = CreateCompatibleDC(screen);
            let mut bits: *mut std::ffi::c_void = std::ptr::null_mut();
            let info = BITMAPINFO {
                bmiHeader: BITMAPINFOHEADER {
                    biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                    biWidth: width as i32,
                    biHeight: -(height as i32),
                    biPlanes: 1,
                    biBitCount: 32,
                    biCompression: BI_RGB.0,
                    ..Default::default()
                },
                ..Default::default()
            };
            let bitmap =
                CreateDIBSection(dc, &info, DIB_RGB_COLORS, &mut bits, None, 0).unwrap();
            let previous = SelectObject(dc, bitmap);

            let whole = RECT {
                left: 0,
                top: 0,
                right: width as i32,
                bottom: height as i32,
            };
            let white = CreateSolidBrush(COLORREF(0x00FF_FFFF));
            FillRect(dc, &whole, white);
            let _ = DeleteObject(white);

            SetBkMode(dc, TRANSPARENT);

            let merged = super::merges(sheet);

            // Every cell's ground goes down before any of the text, because
            // text that runs on past its own cell is drawn over whatever the
            // cell it runs into is painted with. `data_B01`'s headings sit on
            // a pale green band and spill into the next column, which is
            // painted the same green: drawn cell by cell, that second fill
            // took the spilled half of every heading out again.
            for row in &sheet.rows {
                if row.index < layout.first_row {
                    continue;
                }
                let top_at = (row.index - layout.first_row) as usize;
                let (Some(top), Some(bottom)) =
                    (layout.rows.get(top_at), layout.rows.get(top_at + 1))
                else {
                    continue;
                };
                if bottom <= top {
                    continue;
                }
                for cell in &row.cells {
                    if cell.col < layout.first_column {
                        continue;
                    }
                    let (spans_columns, spans_rows) =
                        match merged.get(&(row.index, cell.col)) {
                            Some(Merged::Covered) => continue,
                            Some(Merged::Anchor { columns, rows }) => (*columns, *rows),
                            None => (0, 0),
                        };
                    let left_at = (cell.col - layout.first_column) as usize;
                    let (Some(left), Some(right)) = (
                        layout.columns.get(left_at),
                        layout.columns.get(left_at + 1 + spans_columns as usize),
                    ) else {
                        continue;
                    };
                    if right <= left {
                        continue;
                    }
                    // A table's own dress goes down first; a cell that names a
                    // fill of its own paints over it.
                    let dress = super::dressed_by_table(sheet, row.index, cell.col);
                    let Some(fill) = cell
                        .style
                        .bg_color
                        .as_deref()
                        .or_else(|| dress.as_ref().and_then(|d| d.fill.as_deref()))
                    else {
                        continue;
                    };
                    let bottom = layout
                        .rows
                        .get(top_at + 1 + spans_rows as usize)
                        .unwrap_or(bottom);
                    let brush = CreateSolidBrush(colour(Some(fill), 0xFFFFFF));
                    FillRect(
                        dc,
                        &RECT {
                            left: *left as i32,
                            top: *top as i32,
                            right: *right as i32,
                            bottom: *bottom as i32,
                        },
                        brush,
                    );
                    let _ = DeleteObject(brush);
                }
            }

            for row in &sheet.rows {
                if row.index < layout.first_row {
                    continue;
                }
                let top_at = (row.index - layout.first_row) as usize;
                let (Some(top), Some(bottom)) =
                    (layout.rows.get(top_at), layout.rows.get(top_at + 1))
                else {
                    continue;
                };
                if bottom <= top {
                    continue; // a hidden row takes no space
                }
                for cell in &row.cells {
                    if cell.col < layout.first_column {
                        continue;
                    }
                    // A merged block belongs to its top-left cell; the cells
                    // it covers are not drawn at all.
                    let (spans_columns, spans_rows) =
                        match merged.get(&(row.index, cell.col)) {
                            Some(Merged::Covered) => continue,
                            Some(Merged::Anchor { columns, rows }) => (*columns, *rows),
                            None => (0, 0),
                        };
                    let left_at = (cell.col - layout.first_column) as usize;
                    let (Some(left), Some(right)) = (
                        layout.columns.get(left_at),
                        layout.columns.get(left_at + 1 + spans_columns as usize),
                    ) else {
                        continue;
                    };
                    if right <= left {
                        continue;
                    }
                    let bottom = layout
                        .rows
                        .get(top_at + 1 + spans_rows as usize)
                        .unwrap_or(bottom);
                    let box_ = RECT {
                        left: *left as i32,
                        top: *top as i32,
                        right: *right as i32,
                        bottom: *bottom as i32,
                    };

                    // The fill went down in the pass above; what a table
                    // dresses this cell with is still needed for its text.
                    let dress = super::dressed_by_table(sheet, row.index, cell.col);

                    // A rule sits ON the boundary, not inside the cell: the
                    // bottom of one row and the top of the next are the same
                    // pixel. Measured against Excel on a box ruled all the way
                    // round, whose edges landed exactly on the column and row
                    // starts.
                    let edges: [(&Option<BorderLine>, bool, i32); 4] = [
                        (&cell.style.border_top, true, box_.top),
                        (&cell.style.border_bottom, true, box_.bottom),
                        (&cell.style.border_left, false, box_.left),
                        (&cell.style.border_right, false, box_.right),
                    ];
                    for (line, horizontal, at) in edges {
                        let Some(line) = line else { continue };
                        let rule = super::rule_for(&line.style);
                        let shade = colour(line.color.as_deref(), 0x000000);
                        let ink = CreateSolidBrush(shade);
                        for step in -rule.before..=rule.after {
                            if rule.hollow && step == 0 {
                                continue;
                            }
                            let edge = if horizontal {
                                RECT { top: at + step, bottom: at + step + 1, ..box_ }
                            } else {
                                RECT { left: at + step, right: at + step + 1, ..box_ }
                            };
                            match rule.broken {
                                super::Broken::Whole => {
                                    FillRect(dc, &edge, ink);
                                }
                                broken => {
                                    // A broken rule is inked pixel by pixel
                                    // against the picture, not run by run from
                                    // the cell: that is what keeps the pattern
                                    // in step across a whole ruled sheet.
                                    let (start, stop) = if horizontal {
                                        (edge.left, edge.right)
                                    } else {
                                        (edge.top, edge.bottom)
                                    };
                                    for along in start..stop {
                                        let (x, y) = if horizontal {
                                            (along, edge.top)
                                        } else {
                                            (edge.left, along)
                                        };
                                        if broken.inked(x, y, along) {
                                            let _ = SetPixelV(dc, x, y, shade);
                                        }
                                    }
                                }
                            }
                        }
                        let _ = DeleteObject(ink);
                    }

                    // A rule corner to corner, which a Japanese form draws to
                    // strike a cell out. Both ways at once make a cross.
                    if let Some(line) = &cell.style.border_diagonal {
                        let rule = super::rule_for(&line.style);
                        let width = (1 + rule.before + rule.after).max(1);
                        let pen = CreatePen(
                            PS_SOLID,
                            width,
                            colour(line.color.as_deref(), 0x0000_0000),
                        );
                        let held = SelectObject(dc, pen);
                        if cell.style.diagonal_down {
                            let _ = MoveToEx(dc, box_.left, box_.top, None);
                            let _ = LineTo(dc, box_.right, box_.bottom);
                        }
                        if cell.style.diagonal_up {
                            let _ = MoveToEx(dc, box_.left, box_.bottom, None);
                            let _ = LineTo(dc, box_.right, box_.top);
                        }
                        SelectObject(dc, held);
                        let _ = DeleteObject(pen);
                    }

                    // A carriage return would otherwise be drawn as a glyph.
                    let filtered = super::has_filter_button(sheet, row.index, cell.col);
                    if filtered {
                        // The button hangs from the foot of the heading, a
                        // pixel in from its right edge.
                        let left = box_.right - super::FILTER_BUTTON - 1;
                        let top = box_.bottom - super::FILTER_BUTTON_FOOT
                            - super::FILTER_BUTTON;
                        let face = RECT {
                            left,
                            top,
                            right: box_.right - 1,
                            bottom: top + super::FILTER_BUTTON,
                        };
                        let outline = CreateSolidBrush(colour(Some("A6ACB3"), 0xA6ACB3));
                        FillRect(dc, &face, outline);
                        let _ = DeleteObject(outline);
                        let pale = CreateSolidBrush(colour(Some("FEFEFE"), 0xFEFEFE));
                        let inside = RECT {
                            left: face.left + 1,
                            top: face.top + 1,
                            right: face.right - 1,
                            bottom: face.bottom - 1,
                        };
                        FillRect(dc, &inside, pale);
                        let _ = DeleteObject(pale);
                        // Seven pixels wide at the top, losing one either side
                        // each row until it is a single point.
                        let arrow = CreateSolidBrush(colour(Some("58595B"), 0x58595B));
                        let middle = (face.left + face.right) / 2;
                        for step in 0..4 {
                            let half = 3 - step;
                            let bar = RECT {
                                left: middle - half,
                                top: face.top + 7 + step,
                                right: middle + half + 1,
                                bottom: face.top + 8 + step,
                            };
                            FillRect(dc, &bar, arrow);
                        }
                        let _ = DeleteObject(arrow);
                    }

                    let text = cell_text(&cell.value, &cell.style)
                        .replace("\r\n", "\n");
                    // A cell that does not wrap takes no notice of the breaks
                    // inside it: Excel runs the pieces together on one line
                    // and gives the row the height of one. `_xlsx_cell_break.py`
                    // puts a break at the front, in the middle and at the end
                    // of a cell's text, wrapped and not, in a row of a stated
                    // height and one Excel works out for itself: with wrapping
                    // every break spends a line, without it none of them do —
                    // 「あA(break)いB」 comes out 46 pixels of ink on the one
                    // line where each piece alone is 23. It is what centred
                    // `r03_seizosangyo_tkh`'s headings nine pixels high here.
                    let text = if cell.style.wrap_text {
                        text
                    } else {
                        text.replace('\n', "")
                    };
                    let text = if cell.style.stacked_text {
                        super::stacked_text(&text)
                    } else {
                        text
                    };
                    if text.is_empty() {
                        continue;
                    }
                    // A cell names its own typeface; Calibri is only the
                    // fallback for one that does not.
                    let name = cell.style.font_name.as_deref().unwrap_or("Calibri");
                    let face = wide(name);
                    // A table's header row is bold, and no cell inside the
                    // range says so in its own style.
                    let bold = cell.style.bold
                        || matches!(&dress, Some(dress) if dress.bold);

                    // Excel keeps more room at the left of a cell than at
                    // the right, and both grow with the font's digit.
                    let (left_room, right_room) =
                        super::gutters(name, cell.style.font_size.unwrap_or(11.0), bold, cell.style.italic);
                    let gutter = (left_room * scale).round() as i32;
                    let mut area = box_;
                    area.left += gutter;
                    area.right -= (right_room * scale).round() as i32;
                    if filtered {
                        area.right -= super::FILTER_BUTTON;
                    }
                    let placed = alignment(&cell.style, &cell.value);
                    // A stacked cell that says nothing about where its text
                    // sits across the cell is centred, not left: measured on
                    // data_B01, whose column headings say only `vertical` and
                    // come out down the middle of their columns.
                    let placed = if cell.style.stacked_text
                        && cell.style.horizontal_align.is_none()
                    {
                        Align::Centre
                    } else {
                        placed
                    };

                    // Text centred *across* cells rather than in one: Excel
                    // spreads the centring over the run of neighbours that
                    // carry the same alignment and hold nothing themselves,
                    // which is how a heading is put over a group of columns
                    // without merging them. 45 of the corpus's 285 workbooks
                    // do it.
                    let (from, to) = super::centred_across(row, cell, spans_columns);
                    if from != cell.col || to != cell.col + spans_columns {
                        if let (Some(left), Some(right)) = (
                            layout.columns.get(from.saturating_sub(layout.first_column) as usize),
                            layout
                                .columns
                                .get(to.saturating_sub(layout.first_column) as usize + 1),
                        ) {
                            area.left = *left as i32 + (left_room * scale).round() as i32;
                            area.right = *right as i32 - (right_room * scale).round() as i32;
                        }
                    }

                    // An indent takes three of the workbook's own spaces a
                    // level off the cell, from whichever edge its alignment says.
                    let (before, after) =
                        super::indent_room(&cell.style, super::indent_level(sheet));
                    area.left += (before * scale).round() as i32;
                    area.right -= (after * scale).round() as i32;

                    // A cell told to shrink to fit is drawn smaller until its
                    // text fits, and the size it settles on is not a scaling
                    // of the one it asks for: measured across three faces and
                    // fifteen lengths, Excel comes down a whole pixel of em at
                    // a time and stops at the first that fits.
                    let asked = cell.style.font_size.unwrap_or(11.0);
                    let points = if cell.style.shrink_to_fit && !cell.style.wrap_text {
                        super::shrunk_to_fit(
                            name,
                            asked,
                            bold,
                            cell.style.italic,
                            &text,
                            (area.right - area.left) as f32 / scale,
                        )
                    } else {
                        asked
                    };
                    let pixels = -((points * scale * 96.0 / 72.0).round() as i32);
                    let font = CreateFontW(
                        pixels,
                        0,
                        0,
                        0,
                        if bold { 700 } else { 400 },
                        u32::from(cell.style.italic),
                        u32::from(cell.style.underline),
                        0,
                        DEFAULT_CHARSET.0 as u32,
                        OUT_DEFAULT_PRECIS.0 as u32,
                        CLIP_DEFAULT_PRECIS.0 as u32,
                        // Excel *prints* greyscale-antialiased glyphs, and the
                        // comparison used to be against a print. It is against
                        // the sheet on screen now, and there Excel's edges are
                        // its own ClearType, which this reproduces fringe
                        // colour for fringe colour where DirectWrite cannot.
                        CLEARTYPE_QUALITY.0 as u32,
                        (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                        PCWSTR(face.as_ptr()),
                    );
                    let previous_font = SelectObject(dc, font);
                    // A table header is written in white on the accent, and
                    // that beats whatever the cell's own format inherited.
                    let header = matches!(&dress, Some(dress) if dress.white_text);
                    SetTextColor(
                        dc,
                        if header {
                            colour(None, 0x00FF_FFFF)
                        } else {
                            colour(cell.style.font_color.as_deref(), 0x0000_0000)
                        },
                    );
                    SetBkMode(dc, TRANSPARENT);

                    // The lines to draw: the cell's own breaks, and where it
                    // wraps, the breaks the row's height was measured with.
                    let lines = super::wrapped_lines(
                        name,
                        points,
                        bold,
                        cell.style.italic,
                        &text,
                        cell.style
                            .wrap_text
                            .then(|| (area.right - area.left) as f32 / scale),
                    );
                    if std::env::var("OXI_XLSX_DUMP_LINES").is_ok() {
                        eprintln!(
                            "drawn row {} col {} lines {:?}",
                            row.index, cell.col, lines
                        );
                    }
                    let width_of = |line: &str| -> i32 {
                        let held = wide(line);
                        let letters = &held[..held.len() - 1];
                        let mut measured = SIZE::default();
                        if !letters.is_empty()
                            && GetTextExtentPoint32W(dc, letters, &mut measured).as_bool()
                        {
                            measured.cx
                        } else {
                            0
                        }
                    };
                    let reach = lines.iter().map(|line| width_of(line)).max().unwrap_or(0);

                    // Text too long for its cell runs on over the neighbours,
                    // as long as they are empty — that is what Excel shows, and
                    // a wrapping cell keeps to itself instead. Only text runs
                    // on: a number that will not fit stays in its cell, where
                    // Excel shows ##### rather than let it spill.
                    let runs_on = !cell.style.wrap_text
                        && placed != Align::Spread
                        && matches!(cell.value, CellValue::String(_));
                    if runs_on && reach > area.right - area.left {
                        let spare = reach - (area.right - area.left);
                        let (leftward, rightward) = match placed {
                            Align::Left | Align::Spread => (0, spare),
                            Align::Right => (spare, 0),
                            Align::Centre => (spare / 2, spare - spare / 2),
                        };
                        area.left -= super::room_before(layout, row, cell.col, leftward, &merged);
                        // A merged block's own columns are already inside the
                        // box, so the search for room starts past them.
                        area.right += super::room_after(
                            layout,
                            row,
                            cell.col + spans_columns,
                            rightward,
                            &merged,
                        );
                    }

                    // A line stands in the box Excel gives its font, with the
                    // baseline where Excel puts it, and the block of lines sits
                    // in the cell by the cell's own rule.
                    let (line_px, baseline) =
                        super::line_box(name, points, bold, cell.style.italic)
                            .unwrap_or(((-pixels) as f32, (-pixels) as f32));
                    let line_px = (line_px * scale).round() as i32;
                    let block = line_px * lines.len() as i32;
                    let slack = (box_.bottom - box_.top) - block;
                    // A merged block carries a pixel of leading under its
                    // text that a plain cell does not, so its text sits a
                    // pixel higher: measured over thirteen row heights and
                    // eight fonts by _xlsx_valign_pixels.py, which is what
                    // put the `h2daa*kre` family a pixel low. A single line
                    // keeps it when centred but not when sat on the bottom,
                    // and several lines with no room to spare lose it again.
                    let merged_block = spans_columns > 0 || spans_rows > 0;
                    let one_line = lines.len() == 1;
                    let top = box_.top
                        + match cell.style.vertical_align.as_deref() {
                            Some("top") => 0,
                            Some("center") | Some("centre") => {
                                // Several lines with no room to spare lose
                                // the pixel again, and are centred as a plain
                                // cell's would be.
                                let leading =
                                    i32::from(merged_block && (one_line || slack > 0));
                                ((slack - leading) as f32 / 2.0).floor() as i32
                            }
                            // Sat on the bottom, only a block of several
                            // lines gives the pixel up.
                            _ => slack - i32::from(merged_block && !one_line),
                        };

                    // Nothing is drawn outside the cell, or outside the room
                    // the text was given to run on into. The head of the ink
                    // is cut a pixel below the row's top edge, not at it, at
                    // every row height there is: measured on both ends by
                    // _xlsx_bleed_threshold.py.
                    //
                    // The foot is not cut at all when the row can hold the
                    // face's ascent and three pixels more. ＭＳ Ｐゴシック at
                    // 11 point runs its last scanline into the row below in a
                    // 16-pixel row and is cut dead at the edge in a 15-pixel
                    // one; the turn is the device's ascent plus three in all
                    // fourteen faces, sizes and weights swept. Only a plain
                    // cell does it — one line, no wrapping, no merge — which
                    // is the same family that runs its text on over an empty
                    // neighbour sideways. Cutting the foot at the row's edge
                    // is what left the tight rows of _xlsx_valign_pixels.py
                    // reading a pixel short of Excel.
                    let spills = one_line
                        && !cell.style.wrap_text
                        && !merged_block
                        && super::held(|counter| {
                            counter.shape_of(name, points, bold, cell.style.italic)
                        })
                        .is_some_and(|(ascent, _, _)| {
                            (box_.bottom - box_.top) as f32 / scale >= ascent + 3.0
                        });
                    let cut = if spills { height as i32 } else { box_.bottom };
                    let clip = CreateRectRgn(area.left, box_.top + 1, area.right, cut);
                    SelectClipRgn(dc, clip);
                    SetTextAlign(dc, TA_BASELINE | TA_LEFT);
                    let mut at = top + (baseline * scale).round() as i32;

                    // Parts of the text dressed differently — an 8pt aside
                    // inside an 11pt cell, a raised footnote marker — are
                    // drawn one after another, each in its own font. Only a
                    // cell that fits on one line: a wrapped one is left to the
                    // plain path below, where the breaking is what matters.
                    // Room a number format asks for and shows nothing in. The
                    // text carries a space where the format says `_x`, so what
                    // is missing is the difference between that space and x.
                    let reserved = match (&cell.value, cell.style.number_format.as_deref()) {
                        (CellValue::Number(value), Some(format)) if format.contains('_') => {
                            let (before, after) = super::reserved_room(format, *value < 0.0);
                            let blank = width_of(" ");
                            let room = |held: Vec<char>| {
                                held.iter()
                                    .map(|letter| width_of(&letter.to_string()) - blank)
                                    .sum::<i32>()
                            };
                            (room(before), room(after))
                        }
                        _ => (0, 0),
                    };
                    let dressed_runs = !cell.runs.is_empty() && lines.len() == 1;
                    if dressed_runs {
                        let piece = |run: &oxicells_core::ir::TextRun| {
                            let raised = run.vert_align.is_some();
                            let size = run.size.unwrap_or(points);
                            let size = if raised { size * 0.65 } else { size };
                            let face = wide(run.font.as_deref().unwrap_or(name));
                            let font = CreateFontW(
                                -((size * scale * 96.0 / 72.0).round() as i32),
                                0,
                                0,
                                0,
                                if run.bold || bold { 700 } else { 400 },
                                u32::from(run.italic || cell.style.italic),
                                u32::from(run.underline || cell.style.underline),
                                0,
                                DEFAULT_CHARSET.0 as u32,
                                OUT_DEFAULT_PRECIS.0 as u32,
                                CLIP_DEFAULT_PRECIS.0 as u32,
                                CLEARTYPE_QUALITY.0 as u32,
                                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                                PCWSTR(face.as_ptr()),
                            );
                            (font, size)
                        };
                        let mut width = 0i32;
                        for run in &cell.runs {
                            let (font, _) = piece(run);
                            let held = SelectObject(dc, font);
                            let letters = wide(&run.text);
                            let mut measured = SIZE::default();
                            if letters.len() > 1
                                && GetTextExtentPoint32W(
                                    dc,
                                    &letters[..letters.len() - 1],
                                    &mut measured,
                                )
                                .as_bool()
                            {
                                width += measured.cx;
                            }
                            SelectObject(dc, held);
                            let _ = DeleteObject(font);
                        }
                        let room = area.right - area.left;
                        let mut x = match placed {
                            Align::Left | Align::Spread => area.left,
                            Align::Right => area.right - width,
                            Align::Centre => {
                                area.left + ((room - width) as f32 / 2.0).ceil() as i32
                            }
                        };
                        for run in &cell.runs {
                            let (font, size) = piece(run);
                            let held = SelectObject(dc, font);
                            // A run can carry its own colour — a red aside in
                            // a black line — and it goes back afterwards.
                            if let Some(shade) = run.color.as_deref() {
                                if !header {
                                    SetTextColor(dc, colour(Some(shade), 0x0000_0000));
                                }
                            }
                            let letters = wide(&run.text);
                            if letters.len() > 1 {
                                let letters = &letters[..letters.len() - 1];
                                // A raised run sits a third of its own size
                                // above the line; a lowered one, below it.
                                let lift = match run.vert_align.as_deref() {
                                    Some("superscript") => {
                                        -((size * scale * 96.0 / 72.0) / 2.2).round() as i32
                                    }
                                    Some("subscript") => {
                                        ((size * scale * 96.0 / 72.0) / 4.0).round() as i32
                                    }
                                    _ => 0,
                                };
                                let _ = TextOutW(dc, x, at + lift, letters);
                                let mut measured = SIZE::default();
                                if GetTextExtentPoint32W(dc, letters, &mut measured).as_bool() {
                                    x += measured.cx;
                                }
                            }
                            SelectObject(dc, held);
                            let _ = DeleteObject(font);
                            if run.color.is_some() && !header {
                                SetTextColor(
                                    dc,
                                    colour(cell.style.font_color.as_deref(), 0x0000_0000),
                                );
                            }
                        }
                    }
                    // A cell dressed in pieces that does not come out on one
                    // line: Excel breaks it with each piece measured in its
                    // own font and gives every line the height of the tallest
                    // piece standing on it. `_xlsx_cell_runs.py`, over seven
                    // dressings in a row of a stated height: a 20-point piece
                    // after an 11-point one puts the first line's ink at 6..29
                    // and the second at 38..61, where an undressed line sits
                    // at 4..16 and 22..34, and a big piece on the last line
                    // grows that line alone. Drawn whole in the cell's own
                    // font — which is what this did — a 20-point title inside
                    // an 11-point cell comes out 11. 58 of the 285 workbooks
                    // hold 817 such cells.
                    let dressed_lines = !dressed_runs
                        && !cell.style.stacked_text
                        && !cell.runs.is_empty()
                        && cell.runs.iter().map(|run| run.text.chars().count()).sum::<usize>()
                            == text.chars().count();
                    if dressed_lines {
                        // Which piece each character belongs to, and what that
                        // piece is worn in.
                        let mut letters: Vec<char> = Vec::new();
                        let mut wearing: Vec<usize> = Vec::new();
                        for (index, run) in cell.runs.iter().enumerate() {
                            for letter in run.text.chars() {
                                letters.push(letter);
                                wearing.push(index);
                            }
                        }
                        let dress = |index: usize| {
                            let run = &cell.runs[index];
                            let raised = run.vert_align.is_some();
                            let size = run.size.unwrap_or(points);
                            let size = if raised { size * 0.65 } else { size };
                            (
                                run.font.clone().unwrap_or_else(|| name.to_string()),
                                size,
                                run.bold || bold,
                                run.italic || cell.style.italic,
                                run.underline || cell.style.underline,
                            )
                        };
                        let font_of = |index: usize| {
                            let (face, size, bold, italic, underline) = dress(index);
                            let named = wide(&face);
                            CreateFontW(
                                -((size * scale * 96.0 / 72.0).round() as i32),
                                0,
                                0,
                                0,
                                if bold { 700 } else { 400 },
                                u32::from(italic),
                                u32::from(underline),
                                0,
                                DEFAULT_CHARSET.0 as u32,
                                OUT_DEFAULT_PRECIS.0 as u32,
                                CLIP_DEFAULT_PRECIS.0 as u32,
                                CLEARTYPE_QUALITY.0 as u32,
                                (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                                PCWSTR(named.as_ptr()),
                            )
                        };
                        // Each character advances what its own piece gives it.
                        let mut steps: Vec<i32> = Vec::with_capacity(letters.len());
                        {
                            let mut at = 0usize;
                            while at < letters.len() {
                                let piece = wearing[at];
                                let font = font_of(piece);
                                let previous = SelectObject(dc, font);
                                while at < letters.len() && wearing[at] == piece {
                                    let held = wide(&letters[at].to_string());
                                    let mut measured = SIZE::default();
                                    let _ = GetTextExtentPoint32W(
                                        dc,
                                        &held[..held.len() - 1],
                                        &mut measured,
                                    );
                                    steps.push(if letters[at] == '\n' {
                                        0
                                    } else {
                                        measured.cx
                                    });
                                    at += 1;
                                }
                                SelectObject(dc, previous);
                                let _ = DeleteObject(font);
                            }
                        }
                        // The breaks: the cell's own, and where it wraps, the
                        // ones the room asks for.
                        let mut breaks: Vec<usize> = Vec::new();
                        let mut start = 0usize;
                        for (index, letter) in letters.iter().enumerate() {
                            if *letter == '\n' {
                                breaks.push(index + 1);
                                start = index + 1;
                            }
                        }
                        let _ = start;
                        if cell.style.wrap_text {
                            let room = (area.right - area.left) as f32 / scale;
                            // Text that ends on a newline ends on an empty
                            // line, and re-breaking the paragraphs must not
                            // swallow the break that made it.
                            let tail = breaks.last() == Some(&letters.len());
                            let mut held: Vec<usize> = Vec::new();
                            let mut from = 0usize;
                            for stop in breaks.iter().copied().chain(std::iter::once(letters.len())) {
                                if from == stop && stop == letters.len() {
                                    continue;
                                }
                                let piece = &letters[from..stop];
                                let inside = super::line_breaks(piece, &steps[from..stop], room);
                                held.extend(inside.iter().map(|at| from + at));
                                if stop < letters.len() {
                                    held.push(stop);
                                }
                                from = stop;
                            }
                            breaks = held;
                            if tail {
                                breaks.push(letters.len());
                            }
                        }
                        // Every line, as a stretch of characters.
                        let mut stretches: Vec<(usize, usize)> = Vec::new();
                        let mut from = 0usize;
                        for stop in breaks.iter().copied().chain(std::iter::once(letters.len())) {
                            stretches.push((from, stop));
                            from = stop;
                        }
                        // A line stands in the box of the tallest piece on it.
                        let boxes: Vec<(f32, f32)> = stretches
                            .iter()
                            .map(|(from, stop)| {
                                let mut tall = (0.0f32, 0.0f32);
                                // A line with no characters of its own — the
                                // empty one text ending in a newline ends on —
                                // is dressed by the newline that made it, which
                                // sits at the end of the line before.
                                let worn: Vec<usize> = if *from == *stop && *from > 0 {
                                    vec![*from - 1]
                                } else {
                                    (*from..*stop).collect()
                                };
                                for at in worn {
                                    let (face, size, bold, italic, _) = dress(wearing[at]);
                                    if let Some((held, base)) =
                                        super::line_box(&face, size, bold, italic)
                                    {
                                        if held > tall.0 {
                                            tall = (held, base);
                                        }
                                    }
                                }
                                if tall.0 == 0.0 {
                                    (line_px as f32 / scale, baseline)
                                } else {
                                    tall
                                }
                            })
                            .collect();
                        let block: i32 = boxes
                            .iter()
                            .map(|(tall, _)| (tall * scale).round() as i32)
                            .sum();
                        let slack = (box_.bottom - box_.top) - block;
                        // Placed by the cell's own rule, the merged block's
                        // pixel of leading included: without it this sat a
                        // pixel below the plain path on every merged cell it
                        // took over, which is what `bunya_taikeizu_point`
                        // lost 0.0199 to.
                        let alone = stretches.len() == 1;
                        let mut at = box_.top
                            + match cell.style.vertical_align.as_deref() {
                                Some("top") => 0,
                                Some("center") | Some("centre") => {
                                    let leading =
                                        i32::from(merged_block && (alone || slack > 0));
                                    ((slack - leading) as f32 / 2.0).floor() as i32
                                }
                                _ => slack - i32::from(merged_block && !alone),
                            };
                        for ((from, stop), (tall, base)) in stretches.iter().zip(&boxes) {
                            let width: i32 = steps[*from..*stop].iter().sum();
                            // A distributed line fills its cell here as it
                            // does when the whole cell is worn in one font:
                            // the same pieces, the same shares. Without it
                            // `kojo`'s headings, which are dressed and
                            // distributed both, came out packed to the left.
                            let room = area.right - area.left;
                            let shown: String = letters[*from..*stop]
                                .iter()
                                .filter(|letter| **letter != '\n')
                                .collect();
                            let pieces = super::distribution(&shown);
                            let spread =
                                placed == Align::Spread && pieces.len() > 1 && room > width;
                            let mut extra: Vec<i32> = vec![0; *stop - *from];
                            if spread {
                                let spare = room - width;
                                let gaps = pieces.len() as i32 - 1;
                                let mut piece = 0usize;
                                let mut left_in_piece = pieces[0];
                                let mut given = 0;
                                for (n, letter) in letters[*from..*stop].iter().enumerate() {
                                    if *letter == '\n' {
                                        continue;
                                    }
                                    left_in_piece -= 1;
                                    if left_in_piece == 0 && piece + 1 < pieces.len() {
                                        piece += 1;
                                        let want = (spare * piece as i32 + gaps - 1) / gaps;
                                        extra[n] = want - given;
                                        given = want;
                                        left_in_piece = pieces[piece];
                                    }
                                }
                            }
                            let middle = area.left + ((room - width) as f32 / 2.0).ceil() as i32;
                            let mut x = match placed {
                                Align::Left => area.left,
                                Align::Right => area.right - width,
                                Align::Centre => middle,
                                Align::Spread if spread => area.left,
                                Align::Spread if room > width => middle,
                                Align::Spread => area.left,
                            };
                            let down = at + (base * scale).round() as i32;
                            let mut walk = *from;
                            while walk < *stop {
                                let piece = wearing[walk];
                                let mut end = walk;
                                while end < *stop && wearing[end] == piece {
                                    end += 1;
                                }
                                let font = font_of(piece);
                                let previous = SelectObject(dc, font);
                                if let Some(shade) = cell.runs[piece].color.as_deref() {
                                    if !header {
                                        SetTextColor(dc, colour(Some(shade), 0x0000_0000));
                                    }
                                }
                                let held: String = letters[walk..end]
                                    .iter()
                                    .filter(|letter| **letter != '\n')
                                    .collect();
                                let wided = wide(&held);
                                if wided.len() > 1 {
                                    // A raised piece sits above the line, a
                                    // lowered one below it, as on one line.
                                    let size = cell.runs[piece].size.unwrap_or(points);
                                    let lift = match cell.runs[piece].vert_align.as_deref() {
                                        Some("superscript") => {
                                            -(((size * 0.65) * scale * 96.0 / 72.0) / 2.2).round()
                                                as i32
                                        }
                                        Some("subscript") => {
                                            (((size * 0.65) * scale * 96.0 / 72.0) / 4.0).round()
                                                as i32
                                        }
                                        _ => 0,
                                    };
                                    if spread {
                                        // The share a character was given is
                                        // written into the step that follows
                                        // it, as the plain path writes it.
                                        let mut given: Vec<i32> = Vec::new();
                                        for at in walk..end {
                                            if letters[at] == '\n' {
                                                continue;
                                            }
                                            let step = steps[at] + extra[at - *from];
                                            for unit in 0..letters[at].len_utf16() {
                                                given.push(if unit == 0 { step } else { 0 });
                                            }
                                        }
                                        let _ = ExtTextOutW(
                                            dc,
                                            x,
                                            down + lift,
                                            ETO_OPTIONS(0),
                                            None,
                                            PCWSTR(wided.as_ptr()),
                                            (wided.len() - 1) as u32,
                                            Some(given.as_ptr()),
                                        );
                                    } else {
                                        let _ =
                                            TextOutW(dc, x, down + lift, &wided[..wided.len() - 1]);
                                    }
                                }
                                x += steps[walk..end].iter().sum::<i32>()
                                    + extra[walk - *from..end - *from].iter().sum::<i32>();
                                if cell.runs[piece].color.is_some() && !header {
                                    SetTextColor(
                                        dc,
                                        colour(cell.style.font_color.as_deref(), 0x0000_0000),
                                    );
                                }
                                SelectObject(dc, previous);
                                let _ = DeleteObject(font);
                                walk = end;
                            }
                            at += (tall * scale).round() as i32;
                        }
                    }

                    // A stacked cell is drawn through the vertical face —
                    // "@ＭＳ ゴシック" turned a quarter turn — because that is
                    // the face Excel takes its shapes from: measured character
                    // by character, ー ｰ ～ （ ） 「 」 【 】 ＝ come out on
                    // their side and everything else upright, exactly as the
                    // "@" face draws them. The character sits at the top of
                    // its own line box, the em's width across.
                    if cell.style.stacked_text && !dressed_runs {
                        let em = -pixels;
                        let turned = wide(&format!("@{name}"));
                        let font = CreateFontW(
                            pixels,
                            0,
                            2700,
                            2700,
                            if bold { 700 } else { 400 },
                            u32::from(cell.style.italic),
                            u32::from(cell.style.underline),
                            0,
                            DEFAULT_CHARSET.0 as u32,
                            OUT_DEFAULT_PRECIS.0 as u32,
                            CLIP_DEFAULT_PRECIS.0 as u32,
                            CLEARTYPE_QUALITY.0 as u32,
                            (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                            PCWSTR(turned.as_ptr()),
                        );
                        // A letter or a digit is not turned at all: Excel
                        // stacks `01糖尿病` as an upright 0 over an upright 1
                        // over the kanji, and the turned face has no rotated
                        // shape for them — drawn through it they come out on
                        // their side. Only what the "@" face itself turns is
                        // drawn through it.
                        let plain = CreateFontW(
                            pixels,
                            0,
                            0,
                            0,
                            if bold { 700 } else { 400 },
                            u32::from(cell.style.italic),
                            u32::from(cell.style.underline),
                            0,
                            DEFAULT_CHARSET.0 as u32,
                            OUT_DEFAULT_PRECIS.0 as u32,
                            CLIP_DEFAULT_PRECIS.0 as u32,
                            CLEARTYPE_QUALITY.0 as u32,
                            (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
                            PCWSTR(face.as_ptr()),
                        );
                        let held = SelectObject(dc, font);
                        // A stacked character is centred in the whole cell,
                        // gutters and all: measured on data_B01, where the
                        // heading of a merged group lands a pixel further
                        // right than centring inside the gutters puts it.
                        let left = match placed {
                            Align::Left | Align::Spread => area.left,
                            Align::Right => area.right - em,
                            Align::Centre => {
                                box_.left + ((box_.right - box_.left - em) as f32 / 2.0).round() as i32
                            }
                        };
                        for (step, line) in lines.iter().enumerate() {
                            let letters = wide(line);
                            let letters = &letters[..letters.len() - 1];
                            if letters.is_empty() {
                                continue;
                            }
                            // The pen of a turned face sits at the top left of
                            // the character, not on a baseline, so the line
                            // box's own padding is what puts it in place.
                            let down = top + step as i32 * line_px + (line_px - em).max(0);
                            if line.chars().all(|letter| letter.is_ascii_graphic()) {
                                SelectObject(dc, plain);
                                SetTextAlign(dc, TA_TOP | TA_LEFT);
                                let mut measured = SIZE::default();
                                let _ = GetTextExtentPoint32W(dc, letters, &mut measured);
                                let _ = TextOutW(
                                    dc,
                                    left + ((em - measured.cx) as f32 / 2.0).round() as i32,
                                    down,
                                    letters,
                                );
                                SetTextAlign(dc, TA_BASELINE | TA_LEFT);
                                SelectObject(dc, font);
                            } else {
                                let _ = TextOutW(dc, left, down, letters);
                            }
                        }
                        SelectObject(dc, held);
                        let _ = DeleteObject(font);
                        let _ = DeleteObject(plain);
                    }
                    for line in lines
                        .iter()
                        .filter(|_| !dressed_runs && !dressed_lines && !cell.style.stacked_text)
                    {
                        let held = wide(line);
                        let letters = &held[..held.len() - 1];
                        if !letters.is_empty() {
                            let width = width_of(line);
                            let room = area.right - area.left;
                            // A distributed cell fills its whole width, which
                            // is how a Japanese sheet sets a heading: 第 ３ 表,
                            // not 第３表. It is spread by the pieces it could
                            // break a line at, so a Latin word travels whole,
                            // the first piece sits against the left edge and
                            // the last against the right, with nothing kept
                            // back at either end. A single piece is centred
                            // instead — measured on _xlsx_distributed.py.
                            let pieces = super::distribution(line);
                            let spread =
                                placed == Align::Spread && pieces.len() > 1 && room > width;
                            let middle = area.left + ((room - width) as f32 / 2.0).ceil() as i32;
                            let left = match placed {
                                Align::Left => area.left + reserved.0,
                                Align::Right => area.right - width - reserved.1,
                                // The odd pixel goes to the left of the text.
                                Align::Centre => middle + (reserved.0 - reserved.1) / 2,
                                Align::Spread if room > width => middle,
                                Align::Spread => area.left,
                            };
                            if spread {
                                let spare = room - width;
                                let gaps = pieces.len() as i32 - 1;
                                let mut steps: Vec<i32> = Vec::new();
                                let mut piece = 0usize;
                                let mut left_in_piece = pieces[0];
                                let mut given = 0;
                                for letter in line.chars() {
                                    let one = wide(&letter.to_string());
                                    let mut measured = SIZE::default();
                                    let advance = if GetTextExtentPoint32W(
                                        dc,
                                        &one[..one.len() - 1],
                                        &mut measured,
                                    )
                                    .as_bool()
                                    {
                                        measured.cx
                                    } else {
                                        0
                                    };
                                    left_in_piece -= 1;
                                    // The gap falls after the last character
                                    // of a piece, and the spare room is shared
                                    // out from the total so it never drifts.
                                    let mut gap = 0;
                                    if left_in_piece == 0 && piece + 1 < pieces.len() {
                                        piece += 1;
                                        let want = (spare * piece as i32 + gaps - 1) / gaps;
                                        gap = want - given;
                                        given = want;
                                        left_in_piece = pieces[piece];
                                    }
                                    // A character written as a surrogate pair
                                    // takes two of the units GDI steps by.
                                    for step in 0..letter.len_utf16() {
                                        steps.push(if step == 0 { advance + gap } else { 0 });
                                    }
                                }
                                let _ = ExtTextOutW(
                                    dc,
                                    area.left,
                                    at,
                                    ETO_OPTIONS(0),
                                    None,
                                    PCWSTR(letters.as_ptr()),
                                    letters.len() as u32,
                                    Some(steps.as_ptr()),
                                );
                            } else {
                                let _ = TextOutW(dc, left, at, letters);
                            }
                        }
                        at += line_px;
                    }
                    SelectClipRgn(dc, None);
                    let _ = DeleteObject(clip);

                    SelectObject(dc, previous_font);
                    let _ = DeleteObject(font);
                }
            }

            // What hangs over the grid is drawn last, so it covers the cells
            // it is laid over rather than the other way round.
            let telling = std::env::var("OXI_XLSX_DUMP_DRAWINGS").is_ok();
            for drawn in &sheet.drawings {
                let Some(box_) = super::drawing_box(drawn, layout, scale) else {
                    if telling {
                        eprintln!("drawing off the picture: from {:?}", drawn.from);
                    }
                    continue;
                };
                if telling {
                    let what = match &drawn.kind {
                        DrawingKind::Picture { bytes } => format!("picture {} bytes", bytes.len()),
                        DrawingKind::Shape(shape) => format!(
                            "{} fill {:?} line {:?} says {:?}",
                            shape.geometry,
                            shape.fill,
                            shape.line.as_ref().map(|line| (&line.color, line.width, &line.dash)),
                            shape.text.as_ref().map(|said| said
                                .paragraphs
                                .iter()
                                .map(|held| held.text.chars().take(12).collect::<String>())
                                .collect::<Vec<_>>()),
                        ),
                        other => format!("{other:?}"),
                    };
                    eprintln!(
                        "drawing {},{} to {},{} room {:?}  {what}",
                        box_.left,
                        box_.top,
                        box_.right,
                        box_.bottom,
                        super::drawing_room(drawn, layout, scale)
                    );
                }
                match &drawn.kind {
                    DrawingKind::Picture { bytes } => picture(dc, bytes, box_),
                    DrawingKind::Shape(held) => {
                        shape(dc, held, box_, super::drawing_room(drawn, layout, scale), scale,
                              sheet.normal_font.as_ref())
                    }
                    DrawingKind::Chart(held) => {
                        graph(dc, held, box_, scale, sheet.normal_font.as_ref())
                    }
                    _ => {}
                }
            }

            // A table is ruled along every one of its rows, over the fills
            // and under whatever hangs above the grid.
            for table in &sheet.tables {
                let Some(shade) = table.rule.as_deref() else {
                    continue;
                };
                let ink = CreateSolidBrush(colour(Some(shade), 0x0000_0000));
                let left = layout
                    .columns
                    .get(table.start_col.saturating_sub(layout.first_column) as usize);
                let right = layout
                    .columns
                    .get((table.end_col + 1).saturating_sub(layout.first_column) as usize);
                let _ = (left, right);
                for row in table.start_row..=table.end_row + 1 {
                    let at = row.checked_sub(layout.first_row).unwrap_or(0) as usize;
                    let Some(top) = layout.rows.get(at) else { continue };
                    for column in table.start_col..=table.end_col {
                        // A cell that rules itself keeps its own rule: the
                        // table's is the ground the cell's format is laid on,
                        // and `procurement_contractor_list` fills and rules
                        // every cell of its table by hand.
                        let ruled = |row: u32| {
                            sheet
                                .rows
                                .iter()
                                .find(|held| held.index == row)
                                .and_then(|held| held.cells.iter().find(|cell| cell.col == column))
                                .is_some_and(|cell| {
                                    cell.style.border_bottom.is_some()
                                        || cell.style.border_top.is_some()
                                })
                        };
                        if ruled(row) || (row > 0 && ruled(row - 1)) {
                            continue;
                        }
                        let (Some(left), Some(right)) = (
                            layout
                                .columns
                                .get(column.saturating_sub(layout.first_column) as usize),
                            layout
                                .columns
                                .get((column + 1).saturating_sub(layout.first_column) as usize),
                        ) else {
                            continue;
                        };
                        let edge = RECT {
                            left: *left as i32,
                            top: *top as i32,
                            right: *right as i32,
                            bottom: *top as i32 + 1,
                        };
                        FillRect(dc, &edge, ink);
                    }
                }
                let _ = DeleteObject(ink);
            }

            // A note the sheet keeps pinned open sits above everything.
            for note in &sheet.comments {
                let extent = (
                    (note.size.0 * super::EMU * 96.0 / 72.0) as i64,
                    (note.size.1 * super::EMU * 96.0 / 72.0) as i64,
                );
                let Some(box_) = super::anchored_box(&note.from, None, Some(extent), layout, scale)
                else {
                    continue;
                };
                if box_.right <= box_.left || box_.bottom <= box_.top {
                    continue;
                }
                let paper = CreateSolidBrush(colour(note.fill.as_deref(), 0x00E1_FFFF));
                FillRect(dc, &box_, paper);
                let _ = DeleteObject(paper);
                let pen = CreatePen(PS_SOLID, 1, COLORREF(0x0000_0000));
                let held = SelectObject(dc, pen);
                let hollow = SelectObject(dc, GetStockObject(NULL_BRUSH));
                let _ = Rectangle(dc, box_.left, box_.top, box_.right, box_.bottom);
                SelectObject(dc, hollow);
                SelectObject(dc, held);
                let _ = DeleteObject(pen);
                // A note shows what fits in its box and no more.
                let clip = CreateRectRgn(box_.left, box_.top, box_.right, box_.bottom);
                SelectClipRgn(dc, clip);
                says(
            dc,
            &note.text,
            Frame { box_, exact: None, pull: 0.0 },
            scale,
            sheet.normal_font.as_ref(),
            true,
        );
                SelectClipRgn(dc, None);
                let _ = DeleteObject(clip);
            }

            // Excel prints no gridlines by default, so none are drawn here.
            let mut canvas = image::RgbImage::new(width, height);
            let pixels =
                std::slice::from_raw_parts(bits as *const u8, (width * height * 4) as usize);
            for y in 0..height {
                for x in 0..width {
                    let at = ((y * width + x) * 4) as usize;
                    canvas.put_pixel(
                        x,
                        y,
                        image::Rgb([pixels[at + 2], pixels[at + 1], pixels[at]]),
                    );
                }
            }

            SelectObject(dc, previous);
            let _ = DeleteObject(bitmap);
            let _ = DeleteDC(dc);
            ReleaseDC(None, screen);
            let _ = CellValue::Empty;
            canvas
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{column_pixels, stacked_text};

    /// The room a cell keeps either side of its text grows with the font's
    /// digit, in steps of two pixels — measured against Excel by narrowing a
    /// column until the text wraps, over four faces, ten sizes and both
    /// weights. The left keeps a pixel more than the right.
    #[test]
    fn the_gutter_grows_with_the_digit() {
        // ＭＳ ゴシック at 11pt has an eight pixel digit, at 18pt a twelve
        // pixel one, at 24pt a sixteen: five, seven and nine together.
        let small = super::gutters("ＭＳ ゴシック", 11.0, false, false);
        let middle = super::gutters("ＭＳ ゴシック", 18.0, false, false);
        let large = super::gutters("ＭＳ ゴシック", 24.0, false, false);
        assert_eq!(small, (3.0, 2.0));
        assert_eq!(middle, (4.0, 3.0));
        assert_eq!(large, (5.0, 4.0));
        // The left always keeps one more than the right.
        for (left, right) in [small, middle, large] {
            assert_eq!(left - right, 1.0);
        }
    }

    /// Where a line of text is allowed to end. Read back out of Excel's own
    /// picture, character by character, over six samples in three faces and
    /// three column widths: a space ends the line it follows *and* the line
    /// it precedes, a hyphen inside a word ends a line, and a minus sign in
    /// front of a number does not.
    #[test]
    fn a_line_ends_at_a_space_or_a_hyphen() {
        assert!(super::may_break(' ', 'q'));
        assert!(super::may_break('k', ' '));
        assert!(super::may_break('-', 'a'));
        assert!(!super::may_break('-', '5'));
        assert!(!super::may_break('c', 'k'));
        // The kinsoku rules still hold: a line does not start with a full
        // stop, nor end with an opening bracket.
        assert!(!super::may_break('あ', '。'));
        assert!(!super::may_break('「', 'あ'));
        assert!(super::may_break('あ', 'い'));
    }

    /// What a cell is drawn as, line by line. A cell that does not wrap keeps
    /// its own breaks and nothing else; an empty stretch between two breaks is
    /// a line of its own, which the row was measured for.
    #[test]
    fn a_cell_is_drawn_as_the_lines_its_row_was_measured_for() {
        let lines = super::wrapped_lines("Calibri", 11.0, false, false, "one\n\ntwo", None);
        assert_eq!(lines, vec!["one", "", "two"]);
        let one = super::wrapped_lines("Calibri", 11.0, false, false, "", None);
        assert_eq!(one, vec![""]);
    }

    /// A stacked cell is the same thing as text with a break after every
    /// character, which is how its row's height and its drawing both come
    /// out right. Breaks already in the text are not counted twice.
    #[test]
    fn stacked_text_puts_every_character_on_its_own_line() {
        assert_eq!(stacked_text("政府統計"), "政\n府\n統\n計");
        assert_eq!(stacked_text("A1"), "A\n1");
        assert_eq!(stacked_text("あ\nい"), "あ\nい");
        assert_eq!(stacked_text(""), "");
    }

    /// Every pair here was read off a worksheet Excel drew: the stored width on
    /// the left, the pixels its column occupied on the right, with a digit of
    /// the standard font measuring 8.
    #[test]
    fn a_stored_width_becomes_the_pixels_excel_drew() {
        assert_eq!(column_pixels(10.625, 8.0), 85.0);
        assert_eq!(column_pixels(14.625, 8.0), 117.0);
        assert_eq!(column_pixels(12.625, 8.0), 101.0);
        assert_eq!(column_pixels(9.625, 8.0), 77.0);
    }

    /// The same rule under Calibri, whose digit measures 7: Excel's own default
    /// column comes to 64 pixels.
    #[test]
    fn the_rule_holds_for_a_narrower_digit() {
        assert_eq!(column_pixels(9.140625, 7.0), 64.0);
    }
}
