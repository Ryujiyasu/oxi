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
        // A cell that names a LATIN face keeps the workbook's Japanese face
        // beside it — xlsx has one slot and a Japanese cell wears two — and
        // the row holds a line built from BOTH: the deeper baseline, and the
        // longer descent under it.
        //
        //     line = max(baseline, その baseline - 1) + max(descent, その descent)
        //
        // `_xlsx_row_companion2.py` reads eleven faces and sizes against
        // `fies_t2`'s own Terminal 14 (row 21, baseline 18) and every one of
        // them lands here: Century 9 and 10 give 21 and 20 where the taller of
        // the two rows alone would give 21 for both, and Century 14 keeps its
        // own 24 because its baseline is the deeper one. The pixel off the
        // Japanese baseline is measured, not derived — a row set in that face
        // ALONE keeps it, and a row that mixes the two does not.
        let font_px = match (sheet.normal_font.as_ref(), line_box_of(face, size, cell.style.bold)) {
            (Some((normal, normal_size)), Some((own, own_base))) if !speaks_japanese(face) => {
                match line_box_of(normal, *normal_size, false) {
                    Some((theirs, their_base)) => {
                        let base = own_base.max(their_base.saturating_sub(1));
                        let under = (own - own_base).max(theirs - their_base);
                        base + under
                    }
                    None => font_px,
                }
            }
            _ => font_px,
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
    ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮ）］｝〉》」』】〕〙〗”’\
    ℃°′″";
/// Characters that may not end one: the opening half of a pair.
const NEVER_ENDS: &str = "（［｛〈《「『【〔〘〖“‘￥＄";

/// Text written on the body of the em, which breaks between any two
/// characters the kinsoku rules allow.
///
/// The symbols a Japanese sheet is full of — ● ○ ▲ △ ■ □ ◆ ◇ ★ ☆ ※ → ← ↑ ↓
/// ① ② Ⅰ ± × ÷ ≦ ≧ ∞ ‰ § — are written on the em as well, and Excel breaks
/// between two of them: `_xlsx_break_symbols.py` sets `ああCCああ` in a column
/// three ideographs wide and reads how many characters the first line holds,
/// and all twenty-six of them fill it where a control of （ holds two. Leaving
/// them out drags a pair down whole — `tb_r8_jizensoudan` sets 「の▲▲審議会」
/// in a seven-character column, and every line under it was a character out.
/// The same sweep found ℃ ° ′ ″ may not START a line, which is why they stand
/// in `NEVER_STARTS`: Excel gives characters back until 「あ」 is alone on the
/// line rather than open one on a unit mark.
fn ideographic(letter: char) -> bool {
    matches!(letter as u32,
        0x00A7 | 0x00B0..=0x00B1 | 0x00D7 | 0x00F7
        | 0x1100..=0x115F | 0x2010..=0x2027 | 0x2030..=0x205E
        | 0x2100..=0x2BFF | 0x2E80..=0x303E | 0x3041..=0x33FF | 0x3400..=0x4DBF
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
    // A run of spaces is ONE piece, the way a Latin word is:
    // `_xlsx_distributed_spaces.py` sets `"  有業人員"` beside `" 有業人員"`
    // and Excel starts the kanji four pixels further along — a space's own
    // advance — where giving the second space a share of the spread would
    // move it twenty.
    let clustered = |before: char, after: char| {
        before.is_ascii()
            && after.is_ascii()
            && ((!before.is_ascii_whitespace() && !after.is_ascii_whitespace())
                || (before == ' ' && after == ' '))
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

/// The room a right-aligned italic line keeps past its own advance.
///
/// A slanted line sits further left in Excel than its advance alone would put
/// it. `_xlsx_italic_right.py` reads a LEFT-aligned arm beside every
/// right-aligned one: the left arms are ink for ink identical, so what differs
/// is the width Excel reserves, not the rasterising. Swept over twelve sizes of
/// ＭＳ Ｐゴシック with one glyph in the cell, the reservation is
///
///     floor(em / 6)      8, 11, 12, 13, 15, 16, 19, 21, 27, 32, 37, 48 px em
///                        1(0), 1, 2, 2, 2, 2, 3, 3, 4, 5, 6, 8
///
/// — every size but the smallest, where Excel keeps none. It is a property of
/// the FONT and not of the last glyph: `R6kessan` sets 「〈386,904,389〉」 whose
/// last character has no right bearing at all, and Excel still keeps the two
/// pixels. (Century, メイリオ and 游ゴシック keep one more at every size; none
/// of them is italic anywhere in the corpus.)
#[cfg(windows)]
pub(crate) fn slant_room(points: f32, italic: bool) -> i32 {
    if !italic {
        return 0;
    }
    ((points * 96.0 / 72.0).round() / 6.0).floor() as i32
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

/// How far Excel lets a shape's run wander from its exact place before it
/// gives a pixel back — ahead of it, and behind it.
///
/// A shape steps by `round(design)` a character, so the run drifts by the
/// rounding a character. Excel does NOT re-round the running total the way a
/// naive accumulator would: it keeps stepping until the drift passes a limit
/// and then spends one pixel. `_xlsx_shape_phase.py` reads the limit off
/// Excel's own picture over ten faces and four sizes, and
/// `_xlsx_shape_phase2.py` fits it arm by arm:
///
///     ＭＳ 明朝 / ＭＳ ゴシック / BIZ UDゴシック   ahead 1.05..1.30  behind 0.05..0.30
///     メイリオ / 游 / ＭＳ Ｐ / Meiryo UI / HGS   ahead 3.35..3.65  behind 5.05..5.30
///
/// The two groups are told apart by the device's own hinted advance: where it
/// is WIDER than `round(design)` — the ＭＳ faces report `round(design) + 1`
/// for every character — the run is held tight; everywhere else it wanders.
fn phase_room(
    counter: &LineCounter,
    // The face whose DEVICE advance decides which pair of caps this is. For
    // an installed face it is the face itself; for one this machine has not
    // got it is whatever the run's own `pitchFamily` and charset map to,
    // which is not the face Excel draws with (SX101).
    face: &str,
    points: f32,
    bold: bool,
    italic: bool,
    letters: &[char],
    shares: &[f32],
) -> (f32, f32, bool) {
    let em = points * 96.0 / 72.0;
    // Read the face on one FULL-WIDTH character: a proportional face reports a
    // wider hinted advance for some of its narrow marks, and that is not what
    // tells the two groups apart.
    let at = shares
        .iter()
        .position(|share| *share >= 0.9)
        .unwrap_or_default();
    let tight = letters
        .get(at)
        .and_then(|letter| counter.advances_of(face, points, bold, italic, &[*letter]))
        .and_then(|whole| whole.first().copied())
        .zip(shares.get(at))
        .is_some_and(|(whole, share)| whole > (share * em).round() as i32);
    // The third is `tight` itself, which also says which advance the run
    // STEPS by. A face whose hinted advance is wider than the rounded design
    // for every character — ＭＳ 明朝 and ＭＳ ゴシック — is stepped by the
    // design: `_xlsx_shape_yakumono.py` reads ＭＳ ゴシック's `A` at 14pt as 9
    // where its device says 10. Every other face is stepped by the DEVICE,
    // which is what ＭＳ Ｐゴシック's marks need — its `（` at the same size
    // designs 9.334 and Excel draws 10.
    if tight {
        (1.2, 0.2, true)
    } else {
        (3.5, 5.2, false)
    }
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
        let (ahead, behind, tight) = phase_room(counter, face, points, bold, italic, &letters, &shares);
        // The step is the DEVICE advance, not the design one rounded. The two
        // are the same for a full-width character, which is what the phase
        // sweeps were made of, and they part company on a MARK: ＭＳ Ｐゴシック
        // at 14pt designs its （ 9.334 wide, which rounds to 9, and the device
        // hints it to 10 — and Excel steps 10. `_xlsx_shape_yakumono.py` reads
        // four arms where the two readings differ and Excel follows the device
        // every time: （ at 14pt (round 9, device 10, Excel 10), Ａ at 10pt
        // (round 10, device 9, Excel 9), ａ at 10pt (8 / 7 / 7), A at 16pt
        // (14 / 13 / 13). The drift is still measured against the exact design
        // sum, which is what makes a long run of kanji give a pixel back.
        let devices = (!tight)
            .then(|| counter.advances_of(face, points, bold, italic, &letters))
            .flatten()
            .unwrap_or_default();
        // A SPACE is set on its exact place, not on the device's rounding of
        // it. `_xlsx_shape_latin_step.py` writes 「あ」, N spaces and 「あ」 in
        // one shape and reads how far the second ideograph stands from the
        // first, over N from 0 to 12: Yu Gothic UI 12pt, ＭＳ Ｐゴシック and
        // ＭＳ Ｐ明朝 11pt, メイリオ, Meiryo UI and 游ゴシック 12pt — six faces
        // whose space designs 4.38, 4.47, 4.47, 5.44, 5.44 and 4.56 pixels
        // against device advances of 4, 5, 5, 5, 5 and 5 — all read the
        // design's running total at every count. A run of ｉ or of １ beside
        // them reads the device, so it is the space and not Latin at large.
        // `glossary_05`'s flowchart is seven of them before 「対応可能な規模」:
        // Excel stands it 31 pixels along, `round(7 x 4.3828)`, where seven
        // device advances make 28. (One arm of the fifty lands on an exact
        // half — Yu Gothic UI's ninth space, at 52.5 — and Excel takes 52
        // where rounding away from zero gives 53. One tie is not a rule, so
        // it is left rounding as everything else here does.)
        let mut exact = 0.0f32;
        let mut drawn = 0i32;
        let mut held = Vec::with_capacity(shares.len());
        let mut was = 0;
        for (at, share) in shares.iter().enumerate() {
            let advance = share * em;
            exact += advance;
            if letters.get(at) == Some(&' ') {
                drawn = exact.round() as i32;
            } else {
                drawn += match devices.get(at) {
                    Some(device) if *device > 0 => *device,
                    _ => advance.round() as i32,
                };
                if drawn as f32 - exact > ahead {
                    drawn -= 1;
                } else if drawn as f32 - exact < -behind {
                    drawn += 1;
                }
            }
            held.push(drawn - was);
            was = drawn;
        }
        if std::env::var("OXI_XLSX_DUMP_RUN").is_ok() {
            eprintln!("run {face} {points} {text:?} devices {devices:?} steps {held:?}");
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
    metrics: &str,
    points: f32,
    italic: bool,
    worn: &[(bool, String)],
) -> Option<Vec<i32>> {
    held(|counter| {
        let em = points * 96.0 / 72.0;
        let mut exact = 0.0f32;
        let mut drawn = 0i32;
        let mut was = 0;
        let mut steps = Vec::new();
        let mut room: Option<(f32, f32, bool)> = None;
        for (bold, text) in worn {
            let letters: Vec<char> = text.chars().collect();
            let shares = counter.design_advances(face, *bold, italic, &letters)?;
            // The line's own limits, read once off its first run.
            let (ahead, behind, tight) = *room.get_or_insert_with(|| {
                phase_room(counter, metrics, points, *bold, italic, &letters, &shares)
            });
            // The step is the DEVICE advance, unless the face is one of the
            // tight ones; see `shape_run`.
            let devices = (!tight)
                .then(|| counter.advances_of(face, points, *bold, italic, &letters))
                .flatten()
                .unwrap_or_default();
            for (at, (letter, share)) in letters.iter().zip(shares).enumerate() {
                let advance = share * em;
                exact += advance;
                // A space stands on its exact place; see `shape_run`.
                if *letter == ' ' {
                    drawn = exact.round() as i32;
                } else {
                    drawn += match devices.get(at) {
                        Some(device) if *device > 0 => *device,
                        _ => advance.round() as i32,
                    };
                    if drawn as f32 - exact > ahead {
                        drawn -= 1;
                    } else if drawn as f32 - exact < -behind {
                        drawn += 1;
                    }
                }
                // One `dx` a UTF-16 unit, as `ExtTextOutW` wants them.
                for unit in 0..letter.len_utf16() {
                    steps.push(if unit == 0 { drawn - was } else { 0 });
                }
                was = drawn;
            }
            if std::env::var("OXI_XLSX_DUMP_RUN").is_ok() {
                eprintln!(
                    "worn {face} {points} tight={tight} {text:?} devices {devices:?} steps {steps:?}"
                );
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
    let told = std::env::var("OXI_XLSX_DUMP_SHRINK").is_ok();
    if room <= 0.0 || fits(points) {
        if told {
            eprintln!(
                "shrink {face} {points} bold={bold} room={room:.2} width={:?} stays",
                run_width(face, points, bold, italic, text)
            );
        }
        return points;
    }
    let natural = (points * 96.0 / 72.0).round() as i32;
    let chosen = (1..natural)
        .rev()
        .map(|em| em as f32 * 72.0 / 96.0)
        .find(|smaller| fits(*smaller))
        .unwrap_or(points);
    if told {
        eprintln!(
            "shrink {face} {points} bold={bold} room={room:.2} asked_width={:?} -> {chosen} width={:?} text={:?}",
            run_width(face, points, bold, italic, text),
            run_width(face, chosen, bold, italic, text),
            text.chars().take(20).collect::<String>()
        );
    }
    chosen
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
    let (per_em, _, japanese) = face_per_em(face, bold, italic)?;
    let em = points * 96.0 / 72.0;
    let natural = per_em * em;
    Some((natural * if japanese { 1.3 } else { 1.0 }, natural))
}

/// How far the baseline sits below the top of a line, unrounded.
///
/// GDI hands out an ascent already rounded to the device, and drawing with
/// `TA_TOP` makes the renderer use that one — but Excel keeps the exact
/// ascent and rounds the BASELINE. Over a block of lines the two answers part
/// company wherever the fraction crosses a half: `tb_r8_jizensoudan`'s third
/// line came out a pixel low for exactly that reason.
#[cfg(windows)]
pub(crate) fn shape_ascent(face: &str, points: f32, bold: bool, italic: bool) -> Option<f32> {
    let (_, up, _) = face_per_em(face, bold, italic)?;
    Some(up * points * 96.0 / 72.0)
}

/// How far the face falls below its baseline, unrounded.
///
/// The device's own descent is already whole, and a pinned line's lift is
/// derived from a number that is not: `_xlsx_shape_lift.py` sweeps four faces
/// at four sizes against six pinned pitches, and the lift the exact descent
/// predicts is right for メイリオ, Yu Gothic UI and ＭＳ Ｐゴシック at every
/// size where the device's is wrong for メイリオ at 16 point.
#[cfg(windows)]
pub(crate) fn shape_descent(face: &str, points: f32, bold: bool, italic: bool) -> Option<f32> {
    let (tall, up, _) = face_per_em(face, bold, italic)?;
    Some((tall - up) * points * 96.0 / 72.0)
}

/// A face's line height and ascent as shares of the em, measured once.
///
/// Taken at 2048 pixels so the device's own rounding is a thousandth of the
/// answer, which is what lets the exact size do the rest.
#[cfg(windows)]
fn face_per_em(face: &str, bold: bool, italic: bool) -> Option<(f32, f32, bool)> {
    use std::sync::Mutex;
    use windows::Win32::Graphics::Gdi::*;

    // The font's line height and ascent per em, and whether it is an East
    // Asian face.
    static KNOWN: Mutex<Option<std::collections::HashMap<(String, bool, bool), (f32, f32, bool)>>> =
        Mutex::new(None);
    const MEASURED_AT: i32 = 2048;

    let key = (face.to_string(), bold, italic);
    let mut held = KNOWN.lock().ok()?;
    let known = held.get_or_insert_with(std::collections::HashMap::new);
    let found = match known.get(&key) {
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
                metrics.tmAscent as f32 / MEASURED_AT as f32,
                metrics.tmCharSet == SHIFTJIS_CHARSET.0 as u8,
            );
            known.insert(key, found);
            found
        },
    };
    Some(found)
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
    // An anchor the FILE states is held inside its cell, the way Excel holds
    // it. One this worked out for itself is not: a shape inside a group hangs
    // from the cell the group hangs from with the whole distance folded into
    // the offset, so it runs past that cell for a reason Excel never sees.
    // Clamping those too cost `glossary_05` 0.0013 — the group's own text
    // re-wrapped a character.
    anchored_box(
        &drawn.from,
        drawn.to.as_ref(),
        drawn.extent,
        layout,
        scale,
        !drawn.grouped,
    )
}

/// The left edge of a column, wherever it sits against the drawn range.
#[cfg(windows)]
fn column_edge(layout: &Geometry, col: u32) -> Option<f32> {
    match col.checked_sub(layout.first_column) {
        Some(column) => match layout.columns.get(column as usize) {
            Some(edge) => Some(*edge),
            // Past the right of the range, where the picture stops but the
            // sheet does not.
            None => layout
                .after_columns
                .get(column as usize - layout.columns.len())
                .copied(),
        },
        None => layout.before_columns.get(col as usize).copied(),
    }
}

/// The top edge of a row. The drawing part counts rows from zero; the layout
/// counts them from one, the way the sheet states them.
#[cfg(windows)]
fn row_edge(layout: &Geometry, row: u32) -> Option<f32> {
    match (row + 1).checked_sub(layout.first_row) {
        Some(index) => match layout.rows.get(index as usize) {
            Some(edge) => Some(*edge),
            None => layout
                .after_rows
                .get(index as usize - layout.rows.len())
                .copied(),
        },
        None => layout.before_rows.get(row as usize).copied(),
    }
}

/// Where an anchor's offset lands, given the cell's own edges.
///
/// An offset is measured INTO a cell and cannot leave it. Asked of Excel by
/// `_xlsx_anchor_overrun.py`, which writes far corners reaching 0 to 100
/// pixels past a twenty-pixel row and 0 to 300 past a seventy-two pixel
/// column and reads back how big each shape came out: the offset stops at the
/// cell's own edge, 9 of 9 arms down and 7 of 7 across, where taking it at its
/// word fits only the arms that never overrun.
///
/// It is not a rare thing to write. `002` — the corpus floor — pins a note
/// whose box ends at row 3 plus 34 pixels where that row is 27 high, and
/// Excel ends the note at the top of row 4.
#[cfg(windows)]
fn along(edge: f32, next: Option<f32>, off: i64, scale: f32, hold: bool) -> f32 {
    let want = off as f32 / EMU * scale;
    match next {
        Some(next) if hold => edge + want.clamp(0.0, (next - edge).max(0.0)),
        _ => edge + want,
    }
}

/// The box between two anchors, or between one and a stated size.
#[cfg(windows)]
pub(crate) fn anchored_box(
    from: &oxicells_core::ir::Anchor,
    to: Option<&oxicells_core::ir::Anchor>,
    extent: Option<(i64, i64)>,
    layout: &Geometry,
    scale: f32,
    hold: bool,
) -> Option<windows::Win32::Foundation::RECT> {
    // The two axes are asked separately, because a corner can be readable in
    // one and not the other: `002`'s note reaches column 92 of a sheet drawn
    // to 93 — off the picture — while its row is right there. Giving up on
    // both because one is missing loses the row that IS the answer.
    let across = |anchor: &oxicells_core::ir::Anchor| -> Option<i32> {
        let left = column_edge(layout, anchor.col)?;
        let next = column_edge(layout, anchor.col + 1);
        Some(along(left, next, anchor.col_off, scale, hold).round() as i32)
    };
    let down = |anchor: &oxicells_core::ir::Anchor| -> Option<i32> {
        let top = row_edge(layout, anchor.row)?;
        let next = row_edge(layout, anchor.row + 1);
        Some(along(top, next, anchor.row_off, scale, hold).round() as i32)
    };
    let at = |anchor: &oxicells_core::ir::Anchor| -> Option<(i32, i32)> {
        Some((across(anchor)?, down(anchor)?))
    };
    let (left, top) = at(from)?;
    let (right, bottom) = match (to, extent) {
        // A corner past the drawn range falls off the picture, which is where
        // the sheet's own edge is: keep the box and let the drawing be cut.
        //
        // Unless the thing also states how big it is. A note's far corner can
        // name a column the picture never reaches — `002`'s reaches column 92
        // of a sheet drawn to 93 — and stretching it to the sheet's edge then
        // loses the note's own bottom rule entirely. Its stated size is the
        // better answer there.
        (Some(to), Some((cx, cy))) => (
            across(to).unwrap_or(left + (cx as f32 / EMU * scale).round() as i32),
            down(to).unwrap_or(top + (cy as f32 / EMU * scale).round() as i32),
        ),
        (Some(to), None) => at(to).unwrap_or((
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
    // The offset is taken at its word here, where the drawn box clamps it to
    // the cell. The two are different numbers on purpose: what Excel was asked
    // was how big it draws the SHAPE, and the room a line breaks in was
    // derived on its own (`_xlsx_shape_room.py`). Clamping here as well moved
    // `glossary_05`'s flowchart text a character and cost 0.0013, which is the
    // measurement saying the two do not share the rule.
    let at = |anchor: &oxicells_core::ir::Anchor| -> Option<f32> {
        let left = column_edge(layout, anchor.col)?;
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

/// Both of a drawing's side edges before they are put on whole pixels.
///
/// Excel adds the inset to the EXACT edge and rounds once: a box whose left
/// falls at 26.667 pixels sets its text from `round(26.667 + 9.6) = 36`, where
/// rounding the box first and the inset second gives 37. Swept a quarter-point
/// at a time on both sides by `_xlsx_shape_origin.py`, with a filled box beside
/// the text so the box's own rounding cancels out of the reading.
#[cfg(windows)]
pub(crate) fn drawing_edges(
    drawn: &oxicells_core::ir::Drawing,
    layout: &Geometry,
    scale: f32,
) -> Option<(f32, f32)> {
    // The offset is taken at its word here, where the drawn box clamps it to
    // the cell. The two are different numbers on purpose: what Excel was asked
    // was how big it draws the SHAPE, and the room a line breaks in was
    // derived on its own (`_xlsx_shape_room.py`). Clamping here as well moved
    // `glossary_05`'s flowchart text a character and cost 0.0013, which is the
    // measurement saying the two do not share the rule.
    let at = |anchor: &oxicells_core::ir::Anchor| -> Option<f32> {
        let left = column_edge(layout, anchor.col)?;
        Some(left + anchor.col_off as f32 / EMU * scale)
    };
    let left = at(&drawn.from)?;
    let right = match (drawn.to.as_ref(), drawn.extent) {
        (Some(to), _) => at(to).unwrap_or(*layout.columns.last().unwrap_or(&0.0)),
        (None, Some((cx, _))) => left + cx as f32 / EMU * scale,
        (None, None) => return None,
    };
    Some((left, right))
}

/// Where a drawing's top edge falls before it was put on a whole pixel.
///
/// The sides are given by `drawing_edges`, and Excel adds the inset to the
/// exact edge and rounds ONCE. The top follows the same rule: swept an eighth
/// of a pixel at a time over five shapes — a single line, a wrapped and
/// clipped block of four paragraphs, a second face and size, and a bordered
/// box — `_xlsx_shape_origin_down.py` reads 8 arms of 8 with the exact top
/// and the exact inset rounded together, where rounding them separately reads
/// 6 of 8.
#[cfg(windows)]
pub(crate) fn drawing_down(
    drawn: &oxicells_core::ir::Drawing,
    layout: &Geometry,
    scale: f32,
) -> Option<(f32, f32)> {
    // Taken at its word, as the sides are: what the clamp in `anchored_box`
    // answers is how big Excel draws the SHAPE, which is a different question
    // from where it sets the text.
    let at = |anchor: &oxicells_core::ir::Anchor| -> Option<f32> {
        let top = row_edge(layout, anchor.row)?;
        Some(top + anchor.row_off as f32 / EMU * scale)
    };
    let top = at(&drawn.from)?;
    // The foot matters as much as the head: a block anchored `ctr` or `b`
    // divides what is left of the box between the two, so rounding one edge
    // and not the other moves the text by half of the difference.
    let bottom = match (drawn.to.as_ref(), drawn.extent) {
        (Some(to), _) => at(to).unwrap_or(*layout.rows.last().unwrap_or(&0.0)),
        (None, Some((_, cy))) => top + cy as f32 / EMU * scale,
        (None, None) => return None,
    };
    Some((top, bottom))
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

/// The face whose metrics a run is laid out with, which is not always the one
/// it is drawn with.
///
/// For an installed face the two are the same. For one this machine has not
/// got, Excel draws the substitute `face_in_place` names but SPACES the line
/// by what the run's own `pitchFamily` and charset map to on the device:
/// `_xlsx_cas_face.py --dress` puts `pitchFamily="49" charset="-128"` on a
/// ruler that had none and its ink goes from 136 pixels to the title's own
/// 134, matching it exactly — while centring, wrapping and the box's width
/// move nothing.
pub(crate) fn metrics_face(
    asked: &str,
    drawn: &str,
    charset: Option<i32>,
    pitch_family: Option<i32>,
) -> String {
    match pitch_family {
        Some(pitch) if asked != drawn => stood_in_by(asked, charset.unwrap_or(0), pitch)
            .unwrap_or_else(|| drawn.to_string()),
        _ => drawn.to_string(),
    }
}

#[cfg(windows)]
fn stood_in_by(face: &str, charset: i32, pitch_family: i32) -> Option<String> {
    physical_face_asked(face, (charset & 0xFF) as u32, (pitch_family.max(0) as u32) & 0xFF)
}

#[cfg(not(windows))]
fn stood_in_by(_face: &str, _charset: i32, _pitch_family: i32) -> Option<String> {
    None
}

/// The face Excel draws for a missing one in a CELL, which is not the answer
/// it gives in a shape.
///
/// `face_in_place` above is the shape rule, and it is a table: whatever the
/// pitchFamily says, a Japanese charset draws 游ゴシック and everything else
/// ＭＳ ゴシック. A cell is answered differently, and not by a table at all —
/// Excel hands the name to the device with the charset and the family the
/// `<font>` record states, and draws what the mapper gives back.
///
/// `_xlsx_cell_missing_face.py` reads twenty arms, a workbook each (a name
/// resolves once per document, so arms cannot share one): the four names the
/// corpus asks for and has not got, one invented, and the dressings swept
/// separately. Eighteen are the mapper's own answer, asked with the file's
/// charset — a charset the file omits counted as ANSI — and `family << 4`,
/// which is where GDI keeps FF_ROMAN, FF_SWISS and FF_MODERN:
///
/// | dressing | Excel draws |
/// |---|---|
/// | family 1 + charset 128 | ＭＳ Ｐ明朝 |
/// | family 2 + charset 128 | ＭＳ Ｐゴシック |
/// | family 3 + charset ±128 | ＭＳ ゴシック |
/// | charset ±128, no family | ＭＳ Ｐゴシック |
/// | family 3, no charset | Courier New |
/// | family 2, no charset | Arial |
///
/// The two the mapper does not account for are the ones where the file states
/// NOTHING — no family, and no Japanese charset. There Excel does not ask: it
/// draws ＭＳ ゴシック, which is what SX54 read for the same case in shapes.
///
/// The name itself makes no difference: an invented one, a vendor face, and
/// the corpus's own near-misses of installed names (`MS P ゴシック` with
/// spaces, `MS　Pゴシック` with an ideographic one) all answer alike.
pub(crate) fn cell_face_in_place(
    face: &str,
    charset: Option<i32>,
    family: Option<i32>,
) -> String {
    if face.is_empty() || known_face(face) {
        return face.to_string();
    }
    let japanese = matches!(charset, Some(-128) | Some(128));
    if family.unwrap_or(0) <= 0 && !japanese {
        return "ＭＳ ゴシック".to_string();
    }
    stood_in(face, charset.unwrap_or(0), family.unwrap_or(0))
}

#[cfg(windows)]
fn stood_in(face: &str, charset: i32, family: i32) -> String {
    physical_face_asked(
        face,
        (charset & 0xFF) as u32,
        ((family.max(0) as u32) & 0x0F) << 4,
    )
    .unwrap_or_else(|| "ＭＳ ゴシック".to_string())
}

#[cfg(not(windows))]
fn stood_in(_face: &str, _charset: i32, _family: i32) -> String {
    "ＭＳ ゴシック".to_string()
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

/// Whether the device knows a face by this name — by enumeration, not by what
/// the mapper hands back.
///
/// `installed()` asks the mapper twice and calls a face missing when its answer
/// is the answer to a name nothing can have. That test cannot see a face whose
/// own answer IS that fallback: this machine answers ＭＳ Ｐゴシック to an
/// impossible name, so `MS PGothic` — the English name of a face it has —
/// reads as missing. Enumeration answers the question the way it was asked,
/// aliases and all.
#[cfg(windows)]
fn known_face(face: &str) -> bool {
    use windows::Win32::Foundation::LPARAM;
    use windows::Win32::Graphics::Gdi::*;
    thread_local! {
        static SEEN: std::cell::RefCell<std::collections::HashMap<String, bool>> =
            std::cell::RefCell::new(std::collections::HashMap::new());
    }
    if let Some(held) = SEEN.with(|seen| seen.borrow().get(face).copied()) {
        return held;
    }
    unsafe extern "system" fn count(
        _font: *const LOGFONTW,
        _metrics: *const TEXTMETRICW,
        _kind: u32,
        held: LPARAM,
    ) -> i32 {
        unsafe { *(held.0 as *mut i32) += 1 };
        0
    }
    let held = unsafe {
        let screen = GetDC(None);
        let dc = CreateCompatibleDC(screen);
        let mut asked = LOGFONTW {
            lfCharSet: DEFAULT_CHARSET,
            ..Default::default()
        };
        for (at, letter) in face.encode_utf16().take(31).enumerate() {
            asked.lfFaceName[at] = letter;
        }
        let mut found: i32 = 0;
        EnumFontFamiliesExW(
            dc,
            &asked,
            Some(count),
            LPARAM(&mut found as *mut i32 as isize),
            0,
        );
        let _ = DeleteDC(dc);
        ReleaseDC(None, screen);
        found > 0
    };
    SEEN.with(|seen| seen.borrow_mut().insert(face.to_string(), held));
    held
}

#[cfg(not(windows))]
fn known_face(_face: &str) -> bool {
    true
}

/// Whether a face can set Japanese at all.
///
/// A Japanese cell carries TWO faces — the 「日本語用」 and the 「英数字用」 of
/// the font dialog — and xlsx has a slot for only one, so a cell that names a
/// Latin face keeps the workbook's Japanese face beside it. Excel settles the
/// row's height with both in hand: `_xlsx_row_companion.py` dresses a row in
/// one face, sets only the Latin one on a cell, and reads the row back — 42
/// arms, and every one is `max(the Latin face's line, the Japanese face's)`.
/// It is what gives `fies_t2`'s Century 9pt notes 21 pixels where Century's own
/// line is 18: that workbook's Normal is Terminal 14, whose line is 21.
#[cfg(windows)]
fn speaks_japanese(face: &str) -> bool {
    use windows::core::PCWSTR;
    use windows::Win32::Graphics::Gdi::*;
    thread_local! {
        static KNOWN: std::cell::RefCell<std::collections::HashMap<String, bool>> =
            std::cell::RefCell::new(std::collections::HashMap::new());
    }
    if let Some(held) = KNOWN.with(|known| known.borrow().get(face).copied()) {
        return held;
    }
    let held = unsafe {
        let screen = GetDC(None);
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
            ANTIALIASED_QUALITY.0 as u32,
            (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
            PCWSTR(name.as_ptr()),
        );
        let previous = SelectObject(dc, font);
        // One kana is enough: a face that has it sets Japanese, and a face
        // that does not leaves the mark GGI_MARK_NONEXISTING_GLYPHS puts there.
        let letters: Vec<u16> = "あ".encode_utf16().chain(Some(0)).collect();
        let mut glyphs = [0u16; 1];
        let read = GetGlyphIndicesW(
            dc,
            PCWSTR(letters.as_ptr()),
            1,
            glyphs.as_mut_ptr(),
            GGI_MARK_NONEXISTING_GLYPHS,
        );
        let answer = read as i64 != GDI_ERROR as i64 && glyphs[0] != 0xFFFF;
        SelectObject(dc, previous);
        let _ = DeleteObject(font);
        let _ = DeleteDC(dc);
        ReleaseDC(None, screen);
        answer
    };
    KNOWN.with(|known| known.borrow_mut().insert(face.to_string(), held));
    held
}

#[cfg(not(windows))]
fn speaks_japanese(_face: &str) -> bool {
    true
}

/// The face the device actually draws when asked for this one.
///
/// `GetTextFace` answers with the name that was asked for, whether or not
/// anything answered to it; the outline metrics carry the name of the face
/// that was actually realised, which is the one worth comparing.
#[cfg(windows)]
fn physical_face(face: &str) -> Option<String> {
    use windows::Win32::Graphics::Gdi::*;
    physical_face_asked(
        face,
        DEFAULT_CHARSET.0 as u32,
        (DEFAULT_PITCH.0 | FF_DONTCARE.0) as u32,
    )
}

/// The same question, with the charset and the family the file states.
///
/// GDI's mapper takes both: a name it cannot match is answered from the
/// charset (which alphabet) and the family (serif, sans, fixed). Excel asks
/// it the same way, so this is how its answer is read rather than tabulated.
#[cfg(windows)]
fn physical_face_asked(face: &str, charset: u32, pitch: u32) -> Option<String> {
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
            charset,
            OUT_DEFAULT_PRECIS.0 as u32,
            CLIP_DEFAULT_PRECIS.0 as u32,
            DEFAULT_QUALITY.0 as u32,
            pitch,
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

/// The dash pattern in whole pixels, as Excel lays it down.
///
/// The preset ratios are the format's own — a `dash` is four on and three off,
/// in multiples of the rule's width — and ours were already right. What was
/// missing is what Excel does when the rule is an ODD number of pixels wide:
/// every stretch of ink gains a pixel and the gap behind it loses one. The
/// period is untouched, so the error is invisible at the first dash and total
/// by the tenth — `glossary_05`'s flowchart borders scored as if the lines
/// were absent. `_xlsx_shape_dash.py` reads eleven presets at six widths, and
/// the rule holds at every one.
///
/// A ROUND-capped rule is the exception, and the only exception: `dot cap="rnd"`
/// and `dot cap="sq"` are the same preset told apart by the cap alone, and
/// they are drawn differently.
///
/// A gap that falls to nothing joins the ink on either side of it, which is
/// why Excel draws a hairline `sysDot` — one on, one off — as a solid line.
pub(crate) fn dash_runs(pattern: &[u32], width: i32, cap: Option<&str>) -> Vec<u32> {
    if pattern.is_empty() {
        return Vec::new();
    }
    let stretch = width.max(1) as u32;
    let lengthen = width % 2 == 1 && cap != Some("rnd");
    let mut runs: Vec<i64> = Vec::with_capacity(pattern.len());
    for (at, part) in pattern.iter().enumerate() {
        let held = (*part * stretch) as i64;
        runs.push(match (lengthen, at % 2 == 0) {
            (true, true) => held + 1,
            (true, false) => held - 1,
            _ => held,
        });
    }
    // Join across any gap that has closed. What is left is ink, gap, ink, gap;
    // a single stretch of ink means the rule is solid.
    let mut joined: Vec<i64> = Vec::with_capacity(runs.len());
    let mut at = 0;
    while at < runs.len() {
        let mut ink = runs[at];
        let mut gap = runs.get(at + 1).copied().unwrap_or(0);
        at += 2;
        while gap <= 0 && at < runs.len() {
            ink += gap + runs[at];
            gap = runs.get(at + 1).copied().unwrap_or(0);
            at += 2;
        }
        joined.push(ink.max(1));
        if gap > 0 {
            joined.push(gap);
        }
    }
    if joined.len() < 2 {
        return Vec::new();
    }
    joined.into_iter().map(|held| held.max(1) as u32).collect()
}

pub(crate) fn gutters(face: &str, points: f32, bold: bool, italic: bool) -> (f32, f32) {
    let digit = advances(face, points, bold, italic, "0")
        .and_then(|held| held.first().copied())
        .unwrap_or(7) as f32;
    // The step goes BELOW zero for a small digit, and clamping it there was
    // wrong: `_xlsx_gutter_ink.py` swept 5, 6 and 7 point after `barrier_free`
    // — which sets four of its fonts at 6pt — and every face whose digit is
    // four pixels keeps one pixel LESS at each side than a five-pixel one,
    // while every digit of five or more agrees with the old rule. The gutter
    // itself cannot go negative.
    let extra = ((digit - 5.0) / 4.0).floor();
    if std::env::var("OXI_XLSX_DUMP_GUTTER").is_ok() {
        let plain = advances(face, points, false, italic, "0")
            .and_then(|held| held.first().copied())
            .unwrap_or(7);
        eprintln!(
            "gutter {face} {points} bold={bold} digit={digit} plain={plain} left={} right={}",
            3.0 + extra,
            2.0 + extra
        );
    }
    ((3.0 + extra).max(0.0), (2.0 + extra).max(0.0))
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

/// The row a face asks for, and how far down it its baseline sits.
fn line_box_of(face: &str, points: f32, bold: bool) -> Option<(u16, u16)> {
    row_defaults::font_line_box(face, points, bold)
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
        // next line at the first character past them, however many there are
        // — and that means every space `may_break` already counts as one, not
        // the ASCII space alone. `_xlsx_break_space.py` breaks a wrapped cell
        // on each kind in turn and reads where the second line's ink starts:
        // U+0020, U+3000, U+00A0 and the first two doubled all leave the new
        // line flush with the first. `data_A22` wraps a 791-character merged
        // cell whose separators are every one U+3000, and carrying one down
        // indented a line by 12px — the book's worst tile.
        while start + take < letters.len()
            && matches!(letters[start + take], ' ' | '\u{3000}' | '\t')
        {
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
    //
    // A NOTE hangs past it as readily as a drawing does, and counting only
    // the drawings left `002`'s open notes without a far edge to stand on:
    // their anchor names column 92 of a sheet drawn to 87, `anchored_box`
    // could not place it, and the note fell back to the width the file states
    // for it — 454 pixels where the anchor gives 486. Thirty pixels is a
    // character and a half of メイリオ 14pt, and it is what put 「資」 on the
    // second line where Excel keeps it on the first.
    let (reach_column, reach_row) = sheet
        .drawings
        .iter()
        .filter_map(|drawn| drawn.to.as_ref())
        .chain(sheet.comments.iter().filter_map(|note| note.to.as_ref()))
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

/// The marks Excel lays on their side inside a stacked cell.
///
/// Everything else it leaves standing, and draws through the UPRIGHT face —
/// which matters, because ＭＳ 明朝 and ＭＳ ゴシック carry embedded bitmaps at
/// the sizes a sheet uses and the turned face cannot reach them, so the two put
/// down visibly different ink at the same pixel size.
///
/// Derived by `_xlsx_stack_class.py`, which predicts the ink the upright face
/// would leave at Excel's own pen — `top + baseline - ascent`, read from a
/// reference character in the same column — and asks whether Excel's picture is
/// that, to the pixel. Over the CJK punctuation block, the full-width forms, the
/// half-width katakana and the dashes, 205 characters were the upright
/// prediction exactly and these 50 were not, identically for both faces.
///
/// Note what is NOT here: 〜 U+301C, ： ； ＜ ＞ ／ ＼ － and every dash but
/// U+2010 and U+2015 stand, though the turned face has a rotated shape for
/// several of them — Excel's class is its own, not the font's.
pub(crate) fn turned_in_a_stack(letter: char) -> bool {
    matches!(
        letter,
        '\u{2010}' | '\u{2015}' | '\u{2025}' | '\u{2026}'      // ‐ ― ‥ …
        | '\u{3001}' | '\u{3002}'                              // 、 。
        | '\u{3008}'..='\u{3011}'                              // 〈〉《》「」『』【】
        | '\u{3013}'..='\u{3017}'                              // 〓〔〕〖〗
        | '\u{3021}'..='\u{3029}'                              // 〡..〩
        | '\u{302E}' | '\u{302F}'
        | '\u{3038}'..='\u{303A}'                              // 〸〹〺
        | '\u{303E}' | '\u{303F}'
        | '\u{30FC}'                                           // ー
        | '\u{FF08}' | '\u{FF09}'                              // （ ）
        | '\u{FF1D}'                                           // ＝
        | '\u{FF3B}' | '\u{FF3D}' | '\u{FF3F}'                 // ［ ］ ＿
        | '\u{FF5B}'..='\u{FF5E}'                              // ｛ ｜ ｝ ～
        | '\u{FF62}' | '\u{FF63}'                              // ｢ ｣
        | '\u{FF70}'                                           // ｰ
    )
}

/// Half of a centred line's leftover, rounded Excel's way.
///
/// The odd pixel goes to the LEFT of the text — and to the RIGHT when the cell
/// wraps. `_xlsx_center_across.py` walks a column a pixel at a time so the
/// leftover runs through odd and even: 14 widths in ＭＳ Ｐゴシック 12pt with
/// 「一般競争入札（総合評価）」 and 14 in ＭＳ 明朝 11pt with 「契約の方法」, and
/// the wrapping arms step a pixel before the plain ones every time. It is what
/// put the right-hand column of `procurement-plan_outline_01` — every row of
/// it — one pixel out.
pub(crate) fn halfway(spare: i32, wraps: bool) -> i32 {
    let half = spare as f32 / 2.0;
    if wraps {
        half.floor() as i32
    } else {
        half.ceil() as i32
    }
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
    /// A drawing's box before its edges were put on whole pixels.
    ///
    /// The rounded rectangle is what gets drawn; these are what the TEXT
    /// inside it is measured and placed from, and they are carried together
    /// because Excel adds the insets to them and rounds once.
    #[derive(Clone, Copy)]
    struct Exact {
        /// The width a line is broken against (see `drawing_room`).
        room: Option<f32>,
        /// The side edges, which is where the text starts from
        /// (see `drawing_edges`).
        sides: Option<(f32, f32)>,
        /// And the top and bottom edges (see `drawing_down`).
        down: Option<(f32, f32)>,
    }

    unsafe fn shape(
        dc: HDC,
        shape: &oxicells_core::ir::Shape,
        box_: RECT,
        // The box before its edges were put on whole pixels.
        exact: Exact,
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
            let runs = super::dash_runs(pattern, width.max(1), line.cap.as_deref());
            let pen = if runs.len() < 2 {
                // `CreatePen` rounds the ends of anything wider than a pixel,
                // which lengthens a rule by half its width at each end: our
                // 2.25pt line ran 643px where Excel's ran 641, and started
                // one pixel early. Excel cuts its rules flat and squares its
                // corners, so the pen has to say so.
                //
                // Only above a pixel. A geometric pen one pixel wide is not
                // the cosmetic pen we have always drawn hairlines with, and
                // swapping it moved eight workbooks the wrong way by a
                // hair each — the measurement was of LINES, and a hairline
                // has no cap worth the name.
                if width <= 1 {
                    return (CreatePen(PS_SOLID, width, shade), width);
                }
                let brush = LOGBRUSH { lbStyle: BS_SOLID, lbColor: shade, lbHatch: 0 };
                let held = ExtCreatePen(
                    PEN_STYLE(
                        PS_GEOMETRIC.0 | PS_SOLID.0 | PS_ENDCAP_FLAT.0 | PS_JOIN_MITER.0,
                    ),
                    width.max(1) as u32,
                    &brush,
                    None,
                );
                if held.is_invalid() {
                    CreatePen(PS_SOLID, width, shade)
                } else {
                    held
                }
            } else {
                let brush = LOGBRUSH { lbStyle: BS_SOLID, lbColor: shade, lbHatch: 0 };
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
            // flipped one from the corners the other way round. An elbow
            // connector runs along one whole side and then down the next; the
            // corner it turns at is whichever of the box's four the flips and
            // the turn put it at, which `laid` works out.
            "line" | "straightConnector1" | "bentConnector2" => {
                if rule.is_some() {
                    let unit: &[(f32, f32)] = if shape.geometry == "bentConnector2" {
                        &[(0.0, 0.0), (1.0, 0.0), (1.0, 1.0)]
                    } else {
                        &[(0.0, 0.0), (1.0, 1.0)]
                    };
                    let path = laid(unit, shape, box_);
                    let (from_x, from_y) = (path[0].x, path[0].y);
                    let (to_x, to_y) = (path[path.len() - 1].x, path[path.len() - 1].y);
                    // A rule an odd number of pixels wide reaches one pixel
                    // further than its path, exactly as each dash of a broken
                    // one does — a solid rule is a single stretch of ink, and
                    // the same rule governs it.
                    let stretch = shape
                        .line
                        .as_ref()
                        .filter(|held| {
                            // Only a rule that is solid because it was ASKED
                            // to be. A `sysDot` a pixel wide is solid because
                            // its gap closed, and the pixel that closed it is
                            // the same pixel — giving it another would spend
                            // it twice.
                            //
                            // And only a rule that ends bare. The pixel goes
                            // on past where the path stops, which is where an
                            // arrow's tip is: `glossary_05` draws its whole
                            // flowchart out of headed connectors, and the
                            // extra pixel poked through every one of them.
                            held.head_end.is_none()
                                && held.tail_end.is_none()
                                && matches!(held.dash.as_deref(), None | Some("solid"))
                                && ((held.width as f32 / super::EMU) * scale).round().max(1.0)
                                    as i32
                                    % 2
                                    == 1
                                && held.cap.as_deref() != Some("rnd")
                        })
                        .map_or(0, |_| 1);
                    // The pixel goes on past the end of the LAST leg, which
                    // for a bare line is the whole of it.
                    let before = path[path.len() - 2];
                    let (reach_x, reach_y) = (
                        to_x + (to_x - before.x).signum() * stretch,
                        to_y + (to_y - before.y).signum() * stretch,
                    );
                    let _ = MoveToEx(dc, path[0].x, path[0].y, None);
                    for step in 1..path.len() - 1 {
                        let _ = LineTo(dc, path[step].x, path[step].y);
                    }
                    let _ = LineTo(dc, reach_x, reach_y);
                    // And whatever the rule wears at its ends. Measured off
                    // Excel's own picture by `_xlsx_arrow_head.py`, which
                    // reads the head as the ink the same line has WITH one
                    // and has not without — a head is thinnest at its tip and
                    // thickest at its base, so every attempt to find it
                    // inside a single picture measured something else.
                    //
                    // At the width the corpus actually draws them — 1296 of
                    // its 1301 ruled ends are 0.75pt, which is one pixel — a
                    // `triangle` is 7 long and 7 across and an `arrow` is 8
                    // and 10, and 1pt gives the same. Wider rules were swept
                    // too and grow more slowly than the rule does, which is
                    // what the halves below say; the corpus does not live
                    // there, so that part is a fit rather than a law.
                    if let Some(ruled) = shape.line.as_ref() {
                        let thick = ((ruled.width as f32 / super::EMU) * scale).max(1.0);
                        // A head points along the leg it sits on, not at
                        // the far end of the path: an elbow's two legs run at
                        // right angles, and aiming across the corner would
                        // put the arrow at forty-five degrees to both.
                        let after = path[1];
                        for (worn, tip, tail) in [
                            (ruled.tail_end.as_deref(), (to_x, to_y), (before.x, before.y)),
                            (ruled.head_end.as_deref(), (from_x, from_y), (after.x, after.y)),
                        ] {
                            let Some(worn) = worn else { continue };
                            let (long, wide) = match worn {
                                "arrow" => (3.0 * thick + 5.0, 4.0 * thick + 6.0),
                                "stealth" => (2.0 * thick + 4.0, 2.0 * thick + 5.0),
                                // triangle, diamond, oval and the rest.
                                _ => (2.0 * thick + 5.0, 2.0 * thick + 5.0),
                            };
                            let (dx, dy) = ((tip.0 - tail.0) as f32, (tip.1 - tail.1) as f32);
                            let span = (dx * dx + dy * dy).sqrt();
                            if span < 1.0 {
                                continue;
                            }
                            let (ux, uy) = (dx / span, dy / span);
                            let base = (tip.0 as f32 - ux * long, tip.1 as f32 - uy * long);
                            let half = wide / 2.0;
                            let corner = |sign: f32| POINT {
                                x: (base.0 - uy * half * sign).round() as i32,
                                y: (base.1 + ux * half * sign).round() as i32,
                            };
                            let points = [
                                POINT { x: tip.0, y: tip.1 },
                                corner(1.0),
                                corner(-1.0),
                            ];
                            if worn == "arrow" {
                                // An open V, drawn in the rule's own pen.
                                let _ = MoveToEx(dc, points[1].x, points[1].y, None);
                                let _ = LineTo(dc, points[0].x, points[0].y);
                                let _ = LineTo(dc, points[2].x, points[2].y);
                            } else {
                                let paint = CreateSolidBrush(colour(
                                    Some(&ruled.color),
                                    0x0000_0000,
                                ));
                                let held_brush = SelectObject(dc, paint);
                                let _ = Polygon(dc, &points);
                                SelectObject(dc, held_brush);
                                let _ = DeleteObject(paint);
                            }
                        }
                    }
                }
            }
            // A curly brace: two arms reaching the left edge, a body up the
            // middle, and a point at the right. Measured off Excel's own
            // picture by `_xlsx_brace_shape.py` and `_xlsx_brace_adjust.py`
            // rather than taken from a remembered preset definition:
            //   * the point sits at `h x adj2/100000` — six arms, exact;
            //   * `adj1` is capped at `min(a2, 100000-a2)/2 x h/ss` and the
            //     corner's y-radius is `ss x a1/100000`. The capped arm is
            //     what shows the cap is real: adj1 58333 with adj2 11152 comes
            //     to 22.3 capped where the bare arithmetic says 30.9, and the
            //     fitted radius is 23.1.
            //   * the corner's x-radius is half the box, so the arm leaves the
            //     left edge horizontally and meets the body vertically.
            // 24 braces across three workbooks, and one of them is the brace
            // beside `002`'s notes.
            "rightBrace" => {
                let across = (box_.right - box_.left) as f32;
                let down = (box_.bottom - box_.top) as f32;
                let smaller = across.min(down);
                let adjust = |name: &str, fallback: f32| {
                    shape
                        .adjusts
                        .iter()
                        .find(|(held, _)| held == name)
                        .map_or(fallback, |(_, value)| *value as f32)
                };
                let a2 = adjust("adj2", 50_000.0).clamp(0.0, 100_000.0);
                let cap = if smaller > 0.0 {
                    (100_000.0 - a2).min(a2) / 2.0 * down / smaller
                } else {
                    0.0
                };
                let a1 = adjust("adj1", 8_333.0).clamp(0.0, cap.max(0.0));
                let corner = smaller * a1 / 100_000.0;
                let point = down * a2 / 100_000.0;
                // The body has to have somewhere to run, however the adjusts
                // are set.
                let corner = corner.min(point).min(down - point).max(0.0);
                let half = across / 2.0;
                // A quarter ellipse drawn as a curve: the control points sit
                // this far along, which is the usual approximation.
                const PULL: f32 = 0.552_284_8;
                let left = box_.left as f32;
                let top = box_.top as f32;
                let at = |x: f32, y: f32| POINT { x: (left + x).round() as i32, y: (top + y).round() as i32 };
                let _ = BeginPath(dc);
                let start = at(0.0, 0.0);
                let _ = MoveToEx(dc, start.x, start.y, None);
                let curve = |one: (f32, f32), two: (f32, f32), end: (f32, f32)| {
                    let held = [at(one.0, one.1), at(two.0, two.1), at(end.0, end.1)];
                    let _ = PolyBezierTo(dc, &held);
                };
                // Down from the top-left arm to the body.
                curve(
                    (half * PULL, 0.0),
                    (half, corner * (1.0 - PULL)),
                    (half, corner),
                );
                let body = at(half, point - corner);
                let _ = LineTo(dc, body.x, body.y);
                // Out to the point and back again.
                curve(
                    (half, point - corner + corner * PULL),
                    (across - half * PULL, point),
                    (across, point),
                );
                curve(
                    (across - half * PULL, point),
                    (half, point + corner * (1.0 - PULL)),
                    (half, point + corner),
                );
                let foot = at(half, down - corner);
                let _ = LineTo(dc, foot.x, foot.y);
                // And down to the bottom-left arm.
                curve(
                    (half, down - corner + corner * PULL),
                    (half * PULL, down),
                    (0.0, down),
                );
                let _ = EndPath(dc);
                if rule.is_some() {
                    let _ = StrokePath(dc);
                } else {
                    let _ = AbortPath(dc);
                }
            }
            // A pair of square brackets: two strokes, each a straight side
            // with a quarter-circle hook at the top and the bottom. Read off
            // Excel's picture by `_xlsx_bracket_shape.py` over five arms —
            // the hook's radius is `min(w,h) x adj/100000` (default 16667),
            // and it is a CIRCLE, not an ellipse stretched to the box: the
            // 80x267 arm settles that, where adj 25000 gives 20 pixels and
            // not the 67 that scaling by the height would give.
            "bracketPair" => {
                let across = (box_.right - box_.left) as f32;
                let down = (box_.bottom - box_.top) as f32;
                let adjust = shape
                    .adjusts
                    .iter()
                    .find(|(held, _)| held == "adj")
                    .map_or(16_667.0, |(_, value)| *value as f32);
                let hook = (across.min(down) * adjust.clamp(0.0, 50_000.0) / 100_000.0)
                    .min(across / 2.0)
                    .min(down / 2.0)
                    .max(0.0);
                const PULL: f32 = 0.552_284_8;
                let left = box_.left as f32;
                let top = box_.top as f32;
                let at = |x: f32, y: f32| POINT {
                    x: (left + x).round() as i32,
                    y: (top + y).round() as i32,
                };
                let curve = |one: (f32, f32), two: (f32, f32), end: (f32, f32)| {
                    let held = [at(one.0, one.1), at(two.0, two.1), at(end.0, end.1)];
                    let _ = PolyBezierTo(dc, &held);
                };
                let _ = BeginPath(dc);
                // The left bracket, top hook down to the bottom one.
                let start = at(hook, 0.0);
                let _ = MoveToEx(dc, start.x, start.y, None);
                curve(
                    (hook - hook * PULL, 0.0),
                    (0.0, hook - hook * PULL),
                    (0.0, hook),
                );
                let foot = at(0.0, down - hook);
                let _ = LineTo(dc, foot.x, foot.y);
                curve(
                    (0.0, down - hook + hook * PULL),
                    (hook - hook * PULL, down),
                    (hook, down),
                );
                // The right one, which is the same the other way about.
                let start = at(across - hook, 0.0);
                let _ = MoveToEx(dc, start.x, start.y, None);
                curve(
                    (across - hook + hook * PULL, 0.0),
                    (across, hook - hook * PULL),
                    (across, hook),
                );
                let foot = at(across, down - hook);
                let _ = LineTo(dc, foot.x, foot.y);
                curve(
                    (across, down - hook + hook * PULL),
                    (across - hook + hook * PULL, down),
                    (across - hook, down),
                );
                let _ = EndPath(dc);
                if rule.is_some() {
                    let _ = StrokePath(dc);
                } else {
                    let _ = AbortPath(dc);
                }
            }
            // A bevelled box: one face in the middle and four sloped ones
            // around it, each the fill under a fixed lightening or darkening.
            // Read off `bunya_taikeizu_point`'s own headings, whose fill is
            // C0504D: the left face is 217,150,148 — two fifths of the way to
            // white — the foot 154,64,62, four fifths of the fill, and the
            // right 115,48,46, three fifths. Those three land on the unit. The
            // top reads 204,114,112 where a fifth of the way to white gives
            // 205,115,113, one out on every channel and no fill that rounds to
            // C0504D explains all four; a unit of colour is nothing to the
            // picture, so it is left at the fifth its siblings are stated in.
            //
            // The slope is an eighth of the shorter side — the adjustment
            // OOXML leaves at its default — and the rule runs round the outer
            // box and the inner one both.
            "bevel" => {
                let (wide, high) = (box_.right - box_.left, box_.bottom - box_.top);
                let slope = ((wide.min(high) as f32) * 0.125).round().max(1.0) as i32;
                let toward = |shade: u32, target: u32, share: f32| -> u32 {
                    let part = |at: u32| {
                        let (from, to) = (((shade >> at) & 0xFF) as f32, ((target >> at) & 0xFF) as f32);
                        ((from + (to - from) * share).round() as u32).min(255) << at
                    };
                    part(0) | part(8) | part(16)
                };
                if let Some(fill) = shape.fill.as_deref() {
                    let shade = colour(Some(fill), 0x00FF_FFFF).0;
                    let (left, top) = (box_.left, box_.top);
                    let (right, foot) = (box_.right + 1, box_.bottom + 1);
                    let (inner_left, inner_top) = (left + slope, top + slope);
                    let (inner_right, inner_foot) = (right - slope, foot - slope);
                    let faces: [(u32, [POINT; 4]); 5] = [
                        (toward(shade, 0x00FF_FFFF, 0.2), [
                            POINT { x: left, y: top },
                            POINT { x: right, y: top },
                            POINT { x: inner_right, y: inner_top },
                            POINT { x: inner_left, y: inner_top },
                        ]),
                        (toward(shade, 0x00FF_FFFF, 0.4), [
                            POINT { x: left, y: top },
                            POINT { x: inner_left, y: inner_top },
                            POINT { x: inner_left, y: inner_foot },
                            POINT { x: left, y: foot },
                        ]),
                        (toward(shade, 0, 0.2), [
                            POINT { x: left, y: foot },
                            POINT { x: inner_left, y: inner_foot },
                            POINT { x: inner_right, y: inner_foot },
                            POINT { x: right, y: foot },
                        ]),
                        (toward(shade, 0, 0.4), [
                            POINT { x: right, y: top },
                            POINT { x: inner_right, y: inner_top },
                            POINT { x: inner_right, y: inner_foot },
                            POINT { x: right, y: foot },
                        ]),
                        (shade, [
                            POINT { x: inner_left, y: inner_top },
                            POINT { x: inner_right, y: inner_top },
                            POINT { x: inner_right, y: inner_foot },
                            POINT { x: inner_left, y: inner_foot },
                        ]),
                    ];
                    let hollow = SelectObject(dc, GetStockObject(NULL_PEN));
                    for (paint, corners) in faces {
                        let brush = CreateSolidBrush(COLORREF(paint));
                        let held = SelectObject(dc, brush);
                        let _ = Polygon(dc, &corners);
                        SelectObject(dc, held);
                        let _ = DeleteObject(brush);
                    }
                    SelectObject(dc, hollow);
                }
                if rule.is_some() {
                    let empty = SelectObject(dc, GetStockObject(NULL_BRUSH));
                    let _ = Rectangle(dc, box_.left, box_.top, box_.right + 1, box_.bottom + 1);
                    let _ = Rectangle(
                        dc,
                        box_.left + slope,
                        box_.top + slope,
                        box_.right + 1 - slope,
                        box_.bottom + 1 - slope,
                    );
                    SelectObject(dc, empty);
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
                        // GDI rules a rectangle up to `right - 1` and
                        // `bottom - 1`, so a box asked for at its own edges
                        // comes out a pixel short on two sides. Excel rules
                        // the whole box: `application_B`'s 204.7-pixel panel
                        // with a 3-pixel pen measures 208 across in Excel's
                        // picture — the box and the pen — and 207 here, with
                        // its foot a row high as well.
                        let hollow = SelectObject(dc, GetStockObject(NULL_BRUSH));
                        let _ = Rectangle(
                            dc,
                            box_.left,
                            box_.top,
                            box_.right + 1,
                            box_.bottom + 1,
                        );
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
                        dc,
                        box_.left,
                        box_.top,
                        box_.right + 1,
                        box_.bottom + 1,
                        round,
                        round,
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
                        exact: exact.room,
                        pull: preset_pull(&shape.geometry, box_),
                        edges: exact.sides,
                        down: exact.down,
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
    /// Start GDI+ once per process, and say whether it is running.
    fn started() -> bool {
        use windows::Win32::Graphics::GdiPlus::*;
        static STARTED: std::sync::OnceLock<bool> = std::sync::OnceLock::new();
        *STARTED.get_or_init(|| unsafe {
            let mut token: usize = 0;
            let input = GdiplusStartupInput {
                GdiplusVersion: 1,
                ..Default::default()
            };
            let mut output = GdiplusStartupOutput::default();
            GdiplusStartup(&mut token, &input, &mut output) == Status(0)
        })
    }

    /// Draw a chart's curve the way Excel draws it: softened.
    ///
    /// Excel rules a grid line and an axis with a hard pen — a single black
    /// row, which is what GDI already gives — but it draws the SERIES curve
    /// softened. Read a column across `9fd461bf494a_zuhyo`'s curve out of
    /// Excel's own picture and it reads 151, 0, 56: one dark row with a pale
    /// row either side. The same column from GDI reads 0, 0, 0 — three hard
    /// rows, because a steep polyline steps. The geometry already agrees; it
    /// is only the edge that does not, and across the five `zuhyo` books
    /// Excel carries 0.17 to 1.16 of soft ink for every solid pixel while we
    /// carry none at all.
    ///
    /// GDI has no way to soften a line, so the curve alone goes through
    /// GDI+, which does. Everything else on the chart stays on the hard pen.
    /// Returns false if GDI+ will not start, and the caller then rules it the
    /// old way.
    unsafe fn softened(
        dc: HDC,
        points: &[POINT],
        shade: COLORREF,
        width: f32,
        dash: Option<&str>,
    ) -> bool {
        use windows::Win32::Graphics::GdiPlus::*;
        if points.len() < 2 {
            return false;
        }
        if !started() {
            return false;
        }
        let mut graphics: *mut GpGraphics = std::ptr::null_mut();
        if GdipCreateFromHDC(dc, &mut graphics) != Status(0) || graphics.is_null() {
            return false;
        }
        let mut pen: *mut GpPen = std::ptr::null_mut();
        // COLORREF is 0x00BBGGRR; GDI+ wants 0xAARRGGBB.
        let raw = shade.0;
        let argb = 0xFF00_0000
            | ((raw & 0x0000_00FF) << 16)
            | (raw & 0x0000_FF00)
            | ((raw & 0x00FF_0000) >> 16);
        let made = GdipCreatePen1(argb, width.max(1.0), Unit(2), &mut pen) == Status(0)
            && !pen.is_null();
        let mut drew = false;
        if made {
            // The same ratios the hard pen uses, in multiples of the width.
            let pattern: &[f32] = match dash {
                Some("dot") => &[1.0, 3.0],
                Some("dash") => &[4.0, 3.0],
                Some("lgDash") => &[8.0, 3.0],
                Some("dashDot") => &[4.0, 3.0, 1.0, 3.0],
                Some("lgDashDot") => &[8.0, 3.0, 1.0, 3.0],
                Some("lgDashDotDot") => &[8.0, 3.0, 1.0, 3.0, 1.0, 3.0],
                Some("sysDash") => &[3.0, 1.0],
                Some("sysDot") => &[1.0, 1.0],
                Some("sysDashDot") => &[3.0, 1.0, 1.0, 1.0],
                Some("sysDashDotDot") => &[3.0, 1.0, 1.0, 1.0, 1.0, 1.0],
                _ => &[],
            };
            if !pattern.is_empty() {
                let _ = GdipSetPenDashArray(pen, pattern.as_ptr(), pattern.len() as i32);
            }
            let _ = GdipSetSmoothingMode(graphics, SmoothingMode(4));
            let held: Vec<Point> = points.iter().map(|at| Point { X: at.x, Y: at.y }).collect();
            drew =
                GdipDrawLinesI(graphics, pen, held.as_ptr(), held.len() as i32) == Status(0);
            let _ = GdipDeletePen(pen);
        }
        let _ = GdipDeleteGraphics(graphics);
        drew
    }

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
        // The box before its edges were put on whole pixels.
        exact: Exact,
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
        // The plot is a fraction of the chart's box, and the box has a
        // fraction of its own: `a08feeb4a00b_zuhyo`'s runs from 16.98 to
        // 1038.14 pixels, so rounding it first loses an eighth of a pixel of
        // plot and, over forty-eight categories, a whole pixel of stride.
        // Taken exactly, the plot spans 86.778 to 988.111 and every one of the
        // thirty-one ticks Excel draws lands on `floor` of it — 31 of 31,
        // where the rounded box read 28.
        //
        // The fractions are cut, not rounded: measured against Excel's own
        // picture, all four edges of `311e2f9c271e_zuhyo`'s plot land a pixel
        // out when rounded and exactly when truncated.
        let (left_edge, right_edge) = exact
            .sides
            .map(|(left, right)| (left as f64, right as f64))
            .unwrap_or((box_.left as f64, box_.right as f64));
        let (top_edge, foot_edge) = exact
            .down
            .map(|(top, foot)| (top as f64, foot as f64))
            .unwrap_or((box_.top as f64, box_.bottom as f64));
        let (span, height) = (right_edge - left_edge, foot_edge - top_edge);
        let (plot_left, plot_right) =
            (left_edge + frame.x * span, left_edge + (frame.x + frame.w) * span);
        let (plot_top, plot_foot) =
            (top_edge + frame.y * height, top_edge + (frame.y + frame.h) * height);
        let plot = RECT {
            left: plot_left.floor() as i32,
            top: plot_top.floor() as i32,
            right: plot_right.floor() as i32,
            bottom: plot_foot.floor() as i32,
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
        let room = plot_right - plot_left;
        // `crossBetween` is stated on the value axis — it says where that
        // axis crosses the other one — but what it decides is where the
        // categories stand: `midCat` puts the first on the axis itself.
        let mid_cat = up_axis
            .cross_between
            .as_deref()
            .or(along_axis.cross_between.as_deref())
            == Some("midCat");
        let stands_at = |index: usize| -> f64 {
            let step = if mid_cat {
                if count > 1 {
                    room * index as f64 / (count - 1) as f64
                } else {
                    room / 2.0
                }
            } else {
                room * (index as f64 + 0.5) / count.max(1) as f64
            };
            plot_left + step
        };
        // A category falls between two pixels, and its tick is drawn in the
        // one it falls in while its label is centred on the next. Read off
        // `a08feeb4a00b_zuhyo`, whose stride is 169/9 of a pixel so that every
        // ninth category lands on a whole one: the tick matches `floor` at all
        // 31 Excel draws, and the label — measured as the shift that aligns
        // Excel's ink with ours, which cancels the glyph's own bearing —
        // matches `ceil` at 27 of 29, the two misses being single digits five
        // pixels wide whose alignment is a pixel ambiguous either way. Where a
        // category lands on a whole pixel the two agree, and Excel draws them
        // agreeing.
        let across_at = |index: usize| -> i32 { stands_at(index).floor() as i32 };
        let label_at = |index: usize| -> i32 { stands_at(index).ceil() as i32 };
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
                cap: None,
                head_end: None,
                tail_end: None,
            });
            let width = ((line.width as f32 / super::EMU) * scale).round().max(1.0) as i32;
            let pen = ruling_pen(
                colour(Some(&line.color), 0x0000_0000),
                width,
                line.dash.as_deref(),
            );
            // A gap in the data breaks the line rather than being read as a
            // zero, which is what `dispBlanksAs="gap"` asks for, so the curve
            // comes out as one run of points per unbroken stretch.
            let mut runs: Vec<Vec<POINT>> = Vec::new();
            for (index, value) in series.values.iter().enumerate() {
                match value {
                    Some(value) => {
                        let at = POINT { x: across_at(index), y: up_at(*value) };
                        match runs.last_mut() {
                            Some(run) if !run.is_empty() => run.push(at),
                            _ => runs.push(vec![at]),
                        }
                    }
                    None => runs.push(Vec::new()),
                }
            }
            let soften = std::env::var("OXI_XLSX_HARD_CURVE").is_err();
            let held = SelectObject(dc, pen);
            for run in runs.iter().filter(|run| run.len() > 1) {
                let softly = soften
                    && softened(
                        dc,
                        run,
                        colour(Some(&line.color), 0x0000_0000),
                        (line.width as f32 / super::EMU) * scale,
                        line.dash.as_deref(),
                    );
                if !softly {
                    let _ = MoveToEx(dc, run[0].x, run[0].y, None);
                    for at in &run[1..] {
                        let _ = LineTo(dc, at.x, at.y);
                    }
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
        // An axis that states no line of its own is drawn in `898989`, not in
        // black. Asked of `a08feeb4a00b_zuhyo`'s own chart, whose value axis
        // states nothing and comes out 137,137,137: given an explicit black it
        // comes out 0, given `898989` it comes out 137 again, and given red it
        // comes out red — so the grey is a colour and not a line too thin to
        // fill its pixel. `_xlsx_chart_axis.py`. Six axes of the corpus state
        // no line, all in the five `zuhyo` workbooks.
        let axis_pen = |line: &Option<oxicells_core::ir::ShapeLine>| {
            let stated = line.clone().unwrap_or(oxicells_core::ir::ShapeLine {
                color: "898989".into(),
                width: 3175,
                dash: None,
                cap: None,
                head_end: None,
                tail_end: None,
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

        // The two axes are drawn value first, category over it: where they
        // meet, `a08feeb4a00b_zuhyo` shows Excel's black category tick
        // standing on the grey value-axis line and its black axis line
        // covering the grey value tick, which is the order the other way
        // round from ours.
        let (pen, _) = axis_pen(&up_axis.line);
        let mut held = SelectObject(dc, pen);
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

        let (pen, _) = axis_pen(&along_axis.line);
        held = SelectObject(dc, pen);
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
                        (plot_left + room * index as f64 / count as f64).floor() as i32
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
            let named = face(&up_axis.face);
            let font = chart_font(&named, label_size, scale);
            let held = SelectObject(dc, font);
            SetTextAlign(dc, TA_TOP | TA_LEFT);
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
                // A value-axis label stands off from the axis, and the stand-off
                // grows with the text. Measured on `a08feeb4a00b_zuhyo`'s own
                // chart at twelve sizes (`_xlsx_chart_label_gap.py`): the ink
                // stops 10, 13, 15, 17, 18, 19, 22, 24, 29, 32, 36 and 44 pixels
                // short of the axis at 6 to 24 point, which is four thirds of
                // the em to within a pixel and a third at every one of them.
                // Ours stopped 2 short at every size — the label was set flush
                // against the axis and only the glyph's own bearing stood in
                // the way. Which way the tick marks point changes nothing:
                // `in`, `out`, `none` and `cross` all read 18 at 10 point.
                // The size is in points and `scale` is the device's own, so
                // the em in pixels carries the 96-over-72 as well.
                let em = label_size * scale * 96.0 / 72.0;
                // The label's far edge stands an em and a quarter from the
                // plot's EXACT left, and the string is then set from
                // `floor(edge - its own exact width)` with every glyph on
                // `round(origin + the exact running total)`. That is three
                // things at once, and each of them is a pixel:
                //
                // * the edge is measured from the fractional plot left, so
                //   two charts of the same size and different anchors set
                //   their labels differently — `a08feeb4a00b_zuhyo` (86.805)
                //   and `311e2f9c271e_zuhyo` (69.529) part company by one;
                // * the width is the design's, not the device's. Three digits
                //   of ＭＳ 明朝 at 10 point design 20.0 and hint to 21, and
                //   Excel's three stand 13 apart where ours stood 14;
                // * so the run steps 7, 6 rather than 7, 7.
                //
                // Read off every value label of the four `zuhyo` charts —
                // three different fractional lefts, one to three digits — and
                // off `_xlsx_chart_label_gap.py`'s own pictures at 8, 9, 10,
                // 11, 12, 14, 16 and 18 point. Every one of them lands.
                // The old reading, four thirds of an em less two pixels, was
                // this measured through the device's own rounding.
                let edge = plot_left - 1.25 * em as f64;
                let widths = super::shape_widths(&named, label_size * scale, false, false, said.trim());
                let mut steps: Vec<i32> = Vec::new();
                let mut origin = plot.left - measured.cx;
                if let Some(widths) = &widths {
                    origin = (edge - widths.iter().sum::<f32>() as f64).floor() as i32;
                    let (mut walked, mut was) = (0.0f32, 0);
                    for (letter, width) in said.trim().chars().zip(widths) {
                        walked += width;
                        let next = walked.round() as i32;
                        for unit in 0..letter.len_utf16() {
                            steps.push(if unit == 0 { next - was } else { 0 });
                        }
                        was = next;
                    }
                }
                let down = up_at(value) - measured.cy / 2;
                if steps.len() == letters.len() {
                    let _ = ExtTextOutW(
                        dc,
                        origin,
                        down,
                        ETO_OPTIONS(0),
                        None,
                        PCWSTR(letters.as_ptr()),
                        letters.len() as u32,
                        Some(steps.as_ptr()),
                    );
                } else {
                    let _ = TextOutW(dc, origin, down, letters);
                }
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
            // A chart sets its label lines three tenths of an em apart,
            // whatever the face's own line box comes to, and keeps the
            // fraction: `_xlsx_chart_line_pitch.py` stacks one character to a
            // line under the first category and sweeps the size from 6 to 13
            // point, solving for the one stride a single origin can round to
            // all six lines. The intervals close on 1.2981 to 1.3077 of the em
            // — 13 point alone gives that pair — and every other size agrees.
            // The face's line box reads 16 pixels at 10 point where Excel is
            // 17.33, which is the two, four and five pixels the `zuhyo`
            // family's stacked 昭和51 loses down its three lines.
            let pitch = size * scale * 96.0 / 72.0 * 1.3;
            for (index, said) in chart.categories.iter().enumerate() {
                let head = (foot + gap) as f32;
                for (step, line) in super::wrapped_lines(&named, size, false, false, said, Some(step / scale))
                    .into_iter()
                    .enumerate()
                {
                    let letters = wide(&line);
                    let letters = &letters[..letters.len() - 1];
                    if !letters.is_empty() {
                        let at = (head + step as f32 * pitch).round() as i32;
                        let _ = TextOutW(dc, label_at(index), at, letters);
                    }
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
            // A shape inside a group is placed by the group's own fractions,
            // which have already been put on whole pixels here; it has no
            // exact edge of its own to add an inset to.
            shape(dc, held, over, Exact { room: None, sides: None, down: None }, scale, normal);
            if let Some(said) = &held.text {
                says(
                    dc,
                    said,
                    Frame { box_: over, exact: None, pull: 0.0, edges: None, down: None },
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
    /// Carry a preset's path into the box its anchor gives it.
    ///
    /// The path is stated in fractions of the shape's OWN box, which is
    /// mirrored before it is turned. The anchor holds the box the turn LEAVES
    /// the shape in — a tall shape turned a quarter hangs from an anchor as
    /// wide as it was tall — so a point is placed in the unit square, mirrored,
    /// turned, and only then given the anchor's own width and height.
    ///
    /// Measured, `_xlsx_bent_connector.py`: sixteen elbow connectors, every
    /// quarter turn against every pair of flips, read out of Excel's own
    /// picture. The corner the elbow turns at lands on all four corners of the
    /// box in the order this produces, and the ink spans the anchor's box in
    /// every one of them — Excel does not grow the box to hold a turned shape.
    ///
    /// Only quarter turns. The corpus states five turns and all five are one;
    /// anything else is left lying where it was written rather than guessed at.
    pub(super) fn laid(
        points: &[(f32, f32)],
        shape: &oxicells_core::ir::Shape,
        box_: RECT,
    ) -> Vec<POINT> {
        let quarters = {
            let round = shape.rotation.rem_euclid(21_600_000);
            (round % 5_400_000 == 0).then_some(round / 5_400_000)
        };
        let wide = (box_.right - box_.left) as f32;
        let tall = (box_.bottom - box_.top) as f32;
        points
            .iter()
            .map(|&(mut across, mut down)| {
                if shape.flip_h {
                    across = 1.0 - across;
                }
                if shape.flip_v {
                    down = 1.0 - down;
                }
                let (across, down) = match quarters {
                    Some(1) => (1.0 - down, across),
                    Some(2) => (1.0 - across, 1.0 - down),
                    Some(3) => (down, 1.0 - across),
                    _ => (across, down),
                };
                POINT {
                    x: box_.left + (across * wide).round() as i32,
                    y: box_.top + (down * tall).round() as i32,
                }
            })
            .collect()
    }

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
        /// The box's own side edges before they were put on whole pixels.
        /// Excel adds the inset to these and rounds ONCE (`drawing_edges`).
        edges: Option<(f32, f32)>,
        /// And its top and bottom edges, which follow the same rule
        /// (`drawing_down`).
        down: Option<(f32, f32)>,
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
        let Frame { box_, exact, pull, edges, down } = frame;
        let inset = |emu: i64| (emu as f32 / super::EMU * scale).round() as i32;
        let room_of = |emu: i64| emu as f32 / super::EMU * scale;
        let pulled = pull.round() as i32;
        // Excel snaps a text edge to a SIXTEENTH of a pixel before putting it
        // on a whole one, which moves the boundary down by a thirty-second:
        // 0.46875 rather than 0.5. `_xlsx_shape_top_boundary.py` and
        // `_xlsx_shape_left_boundary.py` step the edge a thousandth of a pixel
        // at a time across it — the top steps between 0.468 and 0.469, the
        // left between 0.4683 and 0.4693, and both brackets hold 15/32 and
        // exclude a twip's 0.4667. Two lanes with insets 0.6 of a pixel apart
        // step 0.4 of a pixel apart, which is what says the SUM is what gets
        // rounded rather than the edge on its own.
        let sixteenth = |value: f32| ((value * 16.0).round() / 16.0).round() as i32;
        // Excel sets the text from the EXACT edge plus the exact inset, put on
        // a whole pixel once. Rounding the box first and the inset second
        // costs a pixel wherever the box's own edge falls two thirds of the way
        // across one — `_xlsx_shape_origin.py`, both sides, eight lefts.
        // Only where the preset pulls nothing in: a rounded box's own text
        // rectangle carries a third fraction (`preset_pull`), and which of the
        // three Excel rounds together is not measured yet — `002`'s pink
        // roundRect wants the old arithmetic where its plain rectangles want
        // this one.
        let (from_left, from_right) = match edges.filter(|_| pull == 0.0) {
            Some((left, right)) => (
                sixteenth(left + room_of(said.insets.0)) + pulled,
                sixteenth(right - room_of(said.insets.2)) - pulled,
            ),
            None => (
                box_.left + inset(said.insets.0) + pulled,
                box_.right - inset(said.insets.2) - pulled,
            ),
        };
        if std::env::var("OXI_XLSX_DUMP_TEXTBOX").is_ok() {
            eprintln!(
                "textbox box {},{} to {},{}  exact {:?}  insets {:?}  inset_px {}  pull {pull}  says {:?}",
                box_.left,
                box_.top,
                box_.right,
                box_.bottom,
                down.map(|(top, foot)| ((top * 100.0).round() / 100.0,
                                        (foot * 100.0).round() / 100.0)),
                said.insets,
                inset(said.insets.1),
                said.paragraphs
                    .iter()
                    .map(|held| (
                        held.face.clone(),
                        held.size,
                        held.text.chars().take(6).collect::<String>(),
                    ))
                    .collect::<Vec<_>>(),
            );
        }
        // The top is the exact edge plus the exact inset, put on a pixel once,
        // the same way the sides are — but the boundary it turns on is not the
        // half. `_xlsx_shape_top_boundary.py` steps the edge a thousandth of a
        // pixel at a time across it, with the inset written 0, 3.6 and 7.2
        // points in three lanes: all three step at the same SUM fraction, so
        // it is the sum being put on a pixel, and they step between 0.468 and
        // 0.469 — which brackets 15/32. So Excel snaps the sum to a SIXTEENTH
        // of a pixel first and rounds that, moving the boundary down by a
        // thirty-second. A twip's worth of quantising (0.4667) is ruled out by
        // the same reading.
        //
        // `tb_r8_jizensoudan`'s panel is what this is worth: its top is 8.6865
        // and its inset 4.8, and 13.4865 rounds to 13 where Excel draws 14 —
        // the sixteenth takes it to 13.5 and up.
        // The FOOT does not turn over where the head does. SX123 gave it the
        // head's rule on the assumption that a box is symmetric; asking
        // instead — `_xlsx_shape_foot_boundary.py` hangs a block from the foot
        // and sweeps the box's height a hundredth of a pixel at a time, in two
        // lanes whose bottom insets differ by 0.6 of a pixel — the two lanes
        // step 0.60 apart (so it is the SUM being put on a pixel, as at the
        // head) but they step at a fraction of 0.24, not the head's 15/32.
        // The old rule disagrees with Excel in 46 of 100 sub-pixel positions.
        let foot_of = |value: f32| (value + 1.0 - 0.235).floor() as i32;
        let (top, foot) = match down.filter(|_| pull == 0.0) {
            Some((top, bottom)) => (
                sixteenth(top + room_of(said.insets.1)) + pulled,
                foot_of(bottom - room_of(said.insets.3)) - pulled,
            ),
            None => (
                box_.top + inset(said.insets.1) + pulled,
                box_.bottom - inset(said.insets.3) - pulled,
            ),
        };
        let area = RECT { left: from_left, top, right: from_right, bottom: foot };
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
        // How far below each line's top its BASELINE sits, unrounded.
        let mut leading: Vec<f32> = Vec::new();
        for (index, paragraph) in said.paragraphs.iter().enumerate() {
            // A face this machine has not got is not GDI's business to
            // guess: Excel answers by the run's charset (see `face_in_place`).
            let asked = paragraph
                .face
                .clone()
                .or_else(|| normal.map(|(face, _)| face.clone()))
                .unwrap_or_else(|| "ＭＳ Ｐゴシック".to_string());
            let face = super::face_in_place(&asked, paragraph.charset);
            // …but the METRICS it lays that face out with are a different
            // question: they follow the run's own `pitchFamily`, and Excel
            // asks the device with it. `cas-r*`'s title asks for a missing
            // face at `pitchFamily="49"`, which maps to ＭＳ ゴシック — a face
            // whose device advance is wider than the rounded design, so the
            // run is a TIGHT one and gives a pixel back at the fourth glyph
            // and the seventh. Drawn with 游ゴシック's own loose caps it gives
            // none, and the line comes out two pixels wide (SX101).
            let _metrics = super::metrics_face(&asked, &face, paragraph.charset,
                                              paragraph.pitch_family);
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
            let own = tall;
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
            // Excel keeps the face's EXACT ascent and rounds the baseline it
            // lands on; the device hands out an ascent already rounded, and
            // drawing by the top of the line would use that one. Over
            // twenty-four lines, four faces and eight sizes
            // (`_xlsx_shape_pitch_size.py` solving the pair the tops imply,
            // `_xlsx_block_start_law.py` holding each candidate against it)
            // the start of the block is the half-leading plus this ascent,
            // and nothing else: 31 arms of 31, where the next best candidate
            // holds 13 and a start on a whole pixel holds 8.
            // A note is laid out by the engine that lays out cells, and that
            // one does not ask the device where the baseline is either: it
            // reads the measured table, whose second column is how far down
            // the line box the baseline sits. `line_box` hands both back and
            // this path was throwing the second away, leaving GDI's own
            // rounded ascent to stand in for it — メイリオ at 14 point is 20
            // there where the table says 21, which is `002`'s two notes a
            // pixel high apiece.
            let up = if note {
                super::line_box(&face, paragraph.size, paragraph.bold, paragraph.italic)
                    .map_or(0.0, |(_, baseline)| baseline)
            } else {
                super::shape_ascent(&face, paragraph.size, paragraph.bold, paragraph.italic)
                    .unwrap_or(0.0)
            };
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
                    Some((ascent, device, _)) => {
                        let em = paragraph.size * 96.0 / 72.0;
                        // The face's EXACT descent against a WHOLE quarter of
                        // the em. `_xlsx_shape_lift.py` reads the lift off
                        // Excel over four faces, four sizes and six pinned
                        // pitches — 96 arms, and the lift holds still across
                        // the pitch in 14 of the 16 face-and-size pairs, so
                        // three quarters is the slope and this is the whole
                        // of the rest. What changes is the DESCENT: the exact
                        // one, not the device's already-rounded one, which is
                        // the same thing SX118 found for the ascent. It gets
                        // メイリオ right at 10, 12, 14, 16 and 20 point where
                        // the device's descent misses 16 — `002`'s title
                        // pins its fourth line there — and Yu Gothic UI and
                        // ＭＳ Ｐゴシック at every size.
                        //
                        // …and a three eighths of a pixel that is FITTED, not
                        // derived. With the exact descent alone the floor is a
                        // pixel short at 11 point (`dc4fcff7f5f8_001` pins
                        // three panels there) and right at 16; with the
                        // device's it is the other way about. Over メイリオ at
                        // 9, 10, 11, 12, 14, 16 and 20 point the two bracket
                        // the answer, and `floor(exact - em/4 + c)` reproduces
                        // all seven for any c between 0.214 and 0.454.
                        //
                        // 游ゴシック still wants one more pixel than this at 16
                        // and 20 point, and no arithmetic over the em, the
                        // ascent, the descent or the line box gives it — a
                        // search over 840 candidate formulas tops out at 14 of
                        // 16 and every one of those misses メイリオ at 14.
                        // Left wrong and written down rather than carved out
                        // per face.
                        let descent = super::shape_descent(
                            &face, paragraph.size, paragraph.bold, paragraph.italic)
                            .unwrap_or(device);
                        let lift = (descent - em / 4.0 + 0.375).max(0.0).floor();
                        // The ink of a pinned line may well start above the
                        // line: a pitch smaller than the face asks for is
                        // exactly what a pinned pitch is usually for.
                        let _ = ascent;
                        (0.75 * tall - lift) * scale
                    }
                    // A paragraph that asks for a SHARE of the font's own
                    // pitch moves its baseline three quarters of the CHANGE —
                    // the same slope a pinned pitch has. Over six percentages
                    // in four faces `_xlsx_shape_pitch.py` reads Excel's first
                    // line against 0.75 x pitch and the remainder is dead
                    // constant per face (メイリオ 19.9, ＭＳ Ｐゴシック 21.9,
                    // 游ゴシック 20.9). Leaving the line where the font's own
                    // pitch puts it — which is what centring in the scaled box
                    // does below 100% — is what put `glossary_05`'s flowchart
                    // four pixels out at 80%.
                    None => {
                        let settled = ((own - natural) / 2.0).max(0.0);
                        (settled + 0.75 * (tall - own) + up) * scale
                    }
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

        // The pitch is kept as it is measured and rounded only where a line
        // lands, so a block of many lines does not drift from Excel's — and
        // the leftover the anchor divides is kept the same way. Rounding the
        // block to whole pixels first and halving THAT throws away the half
        // pixel that decides which side the odd one falls: it read a pixel
        // high on `001`'s five-line note and a pixel low on `zuhyo`'s
        // three-line one, which no single rounding of a whole-pixel block
        // can satisfy at once.
        let block: f32 = pitch.iter().sum();
        let slack = (area.bottom - area.top) as f32 - block;
        let mut at = (area.top as f32
            + match said.anchor.as_deref() {
                Some("ctr") => slack / 2.0,
                Some("b") => slack,
                _ => 0.0,
            })
        // The block STARTS on a whole pixel, and its lines walk the exact
        // pitch from there. This is not the same as rounding the block or its
        // pitch — both of those were tried and both went backwards. Excel's
        // own gaps are uneven (ＭＳ 明朝 at 9pt gives 16, 15, 16 and at 14pt
        // 24, 24, 25), which is what accumulating a fraction and rounding
        // each line looks like; the only thing it does differently is begin
        // at a whole number. `zuhyo`'s note starts at 529.4 here, and 19.067
        // from there rounds to 19 then 20 where Excel has 19 and 19.
        .round();
        if std::env::var("OXI_XLSX_DUMP_BLOCK").is_ok() {
            eprintln!(
                "block area {}..{} anchor={:?} block={block:.3} slack={slack:.3} at={at:.3} pitch={:?}",
                area.top,
                area.bottom,
                said.anchor.as_deref(),
                pitch.iter().map(|one| (one * 1000.0).round() / 1000.0).collect::<Vec<_>>(),
            );
        }

        // A line is placed by its BASELINE, which is the number Excel rounds
        // — for a note as well as a shape, the two differing only in where
        // the baseline sits below the line's top. Nothing follows this on the
        // sheet, so the alignment stays as it is left.
        SetTextAlign(dc, TA_BASELINE | TA_LEFT);
        for (step, (index, line, from)) in lines.iter().enumerate() {
            let paragraph = &said.paragraphs[*index];
            // A face this machine has not got is not GDI's business to
            // guess: Excel answers by the run's charset (see `face_in_place`).
            let asked = paragraph
                .face
                .clone()
                .or_else(|| normal.map(|(face, _)| face.clone()))
                .unwrap_or_else(|| "ＭＳ Ｐゴシック".to_string());
            let face = super::face_in_place(&asked, paragraph.charset);
            // …but the METRICS it lays that face out with are a different
            // question: they follow the run's own `pitchFamily`, and Excel
            // asks the device with it. `cas-r*`'s title asks for a missing
            // face at `pitchFamily="49"`, which maps to ＭＳ ゴシック — a face
            // whose device advance is wider than the rounded design, so the
            // run is a TIGHT one and gives a pixel back at the fourth glyph
            // and the seventh. Drawn with 游ゴシック's own loose caps it gives
            // none, and the line comes out two pixels wide (SX101).
            let metrics = super::metrics_face(&asked, &face, paragraph.charset,
                                              paragraph.pitch_family);
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
                            &metrics,
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
                let down = (at + leading[step]).round() as i32;
                if std::env::var("OXI_XLSX_DUMP_BLOCK").is_ok() {
                    eprintln!(
                        "  line {step} at={at:.3} down={down} baseline_off={:.3} face={face} size={} text={:?}",
                        leading[step],
                        paragraph.size,
                        line.chars().take(8).collect::<String>()
                    );
                }
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

    /// Replay an enhanced metafile with its edges softened.
    ///
    /// GDI+ takes an HENHMETAFILE and walks the same records, but through its
    /// own pipeline, so a sloped line comes out anti-aliased. Returns false if
    /// GDI+ will not take it, and the caller then plays it through GDI.
    unsafe fn replayed(dc: HDC, held: HENHMETAFILE, box_: RECT) -> bool {
        use windows::Win32::Graphics::GdiPlus::*;
        if !started() {
            return false;
        }
        let mut graphics: *mut GpGraphics = std::ptr::null_mut();
        if GdipCreateFromHDC(dc, &mut graphics) != Status(0) || graphics.is_null() {
            return false;
        }
        // GDI+ takes ownership of the handle it is given, so it gets a copy:
        // the caller deletes its own either way.
        let mut picture: *mut GpMetafile = std::ptr::null_mut();
        let copied = CopyEnhMetaFileW(held, None);
        let mut drew = false;
        if !copied.is_invalid()
            && GdipCreateMetafileFromEmf(copied, true, &mut picture) == Status(0)
            && !picture.is_null()
        {
            let _ = GdipSetSmoothingMode(graphics, SmoothingMode(4));
            let _ = GdipSetInterpolationMode(graphics, InterpolationMode(7));
            let where_ = RectF {
                X: box_.left as f32,
                Y: box_.top as f32,
                Width: (box_.right - box_.left) as f32,
                Height: (box_.bottom - box_.top) as f32,
            };
            drew = GdipDrawImageRectI(
                graphics,
                picture.cast(),
                where_.X as i32,
                where_.Y as i32,
                where_.Width as i32,
                where_.Height as i32,
            ) == Status(0);
            let _ = GdipDisposeImage(picture.cast());
        } else if !copied.is_invalid() {
            let _ = DeleteEnhMetaFile(copied);
        }
        let _ = GdipDeleteGraphics(graphics);
        drew
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
                // Played straight, GDI rules every line in the metafile hard.
                // Excel does not: a column across the graph in
                // `9fd461bf494a_zuhyo` — which is an EMF, not a chart part —
                // reads 151, 0, 56 out of Excel's picture and 0, 0, 0 out of
                // ours. GDI+ replays the same records with the edges softened,
                // which is what Excel shows, so it is asked first and GDI only
                // catches what it will not take.
                if !replayed(dc, held, box_) {
                    let _ = PlayEnhMetaFile(dc, held, &box_);
                }
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
                    // A cell's paint covers BOTH of its boundaries —
                    // `[left, right]` and `[top, bottom]` — and the cells go
                    // down in order, so a neighbour below or to the right
                    // takes the shared pixel back. Where there is no such
                    // neighbour, or it carries no paint of its own, the pixel
                    // stays. The picture's first row and column are never
                    // painted.
                    //
                    // Two generated books settle it between them, and neither
                    // could alone. `gen_styled` stacks six painted rows: its
                    // red band shows 1..17 because green paints 18 back.
                    // `gen2_000` paints only its header row, and its band
                    // shows 1..18 — the same rule with nothing below to take
                    // the boundary. Reading either on its own gives a rule
                    // that the other contradicts.
                    FillRect(
                        dc,
                        &RECT {
                            left: (*left as i32).max(1),
                            top: (*top as i32).max(1),
                            right: *right as i32 + 1,
                            bottom: *bottom as i32 + 1,
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
                    // A merged block takes each edge from the cell ON that
                    // edge, not from the one it is anchored to. `B4:B5` in
                    // `glossary_05` is anchored on B4, which states a top
                    // rule and no bottom; the bottom medium rule is B5's, and
                    // Excel draws it — 745 pixels of it, right under the
                    // banner, which this drew as nothing at all.
                    let beside = |row_at: u32, column: u32| {
                        sheet
                            .rows
                            .iter()
                            .find(|held| held.index == row_at)
                            .and_then(|held| held.cells.iter().find(|held| held.col == column))
                    };
                    let foot = if spans_rows > 0 {
                        beside(row.index + spans_rows, cell.col)
                            .map_or(&cell.style.border_bottom, |held| &held.style.border_bottom)
                    } else {
                        &cell.style.border_bottom
                    };
                    let far = if spans_columns > 0 {
                        beside(row.index, cell.col + spans_columns)
                            .map_or(&cell.style.border_right, |held| &held.style.border_right)
                    } else {
                        &cell.style.border_right
                    };
                    // Two cells sharing an edge both state a rule for it, and
                    // Excel draws ONE: the one belonging to the cell below, or
                    // to the right. `_xlsx_border_contest.py` stacks two cells
                    // and sweeps the style each states at the edge between
                    // them — thin, medium, thick, double, dashed, dotted,
                    // hair — and in all forty-two pairs, both ways round and
                    // in both directions, what is drawn is the lower or
                    // righthand cell's rule, whatever the two styles are. It
                    // is not the heavier that wins; it is the later.
                    //
                    // Drawing both is only visible when one of them is hollow:
                    // `R6kessan` stacks `bottom thin` on `top double`, and the
                    // thin fills the double's white gap so a two-line rule
                    // reads as a solid three-pixel one.
                    //
                    // A neighbour outside the drawn range draws no rule of
                    // its own, though the file still holds one, so there is
                    // nothing there to give way to either.
                    let inside = |row_at: u32, column: u32| {
                        let down = (row_at as usize).checked_sub(layout.first_row as usize);
                        let across = (column as usize).checked_sub(layout.first_column as usize);
                        matches!(down, Some(down) if down + 1 < layout.rows.len())
                            && matches!(across, Some(across) if across + 1 < layout.columns.len())
                    };
                    // Which cell's rule is actually drawn along an edge, for
                    // the neighbour that shares it. Usually that cell itself.
                    // When a merge COVERS it, the block draws one rule along
                    // its own top (or left) edge instead — so the rule to give
                    // way to is the ANCHOR's, and only when the block BEGINS
                    // at this boundary. A block that straddles the boundary
                    // draws nothing along it, and there is nothing to give way
                    // to.
                    let holder = |row_at: u32, column: u32, horizontal: bool| {
                        if !inside(row_at, column) {
                            return None;
                        }
                        match merged.get(&(row_at, column)) {
                            Some(super::Merged::Covered) => {
                                let held = sheet.merge_cells.iter().find(|merge| {
                                    merge.start_row <= row_at
                                        && row_at <= merge.end_row
                                        && merge.start_col <= column
                                        && column <= merge.end_col
                                })?;
                                let begins = if horizontal {
                                    held.start_row == row_at
                                } else {
                                    held.start_col == column
                                };
                                begins.then_some((held.start_row, held.start_col))
                            }
                            _ => Some((row_at, column)),
                        }
                    };
                    // Giving way is only WORTH it where drawing both can be
                    // seen. For a solid rule the winner is laid down after the
                    // loser and covers it, so the picture is already right;
                    // standing aside as well only moves single-pixel lines
                    // about, and across the corpus that cost as much as it
                    // paid (15 improved against 10 regressed, no net change).
                    // A hollow rule cannot cover anything — its middle is the
                    // gap that makes it a double — so there, and only there,
                    // the loser has to be held back.
                    let hollow = |line: &Option<BorderLine>| {
                        line.as_ref().is_some_and(|line| super::rule_for(&line.style).hollow)
                    };
                    let below = holder(row.index + spans_rows + 1, cell.col, true)
                        .and_then(|(row_at, column)| beside(row_at, column))
                        .is_some_and(|held| hollow(&held.style.border_top));
                    let after = holder(row.index, cell.col + spans_columns + 1, false)
                        .and_then(|(row_at, column)| beside(row_at, column))
                        .is_some_and(|held| hollow(&held.style.border_left));
                    if std::env::var("OXI_XLSX_DUMP_EDGES").is_ok() {
                        // `left`/`right` here are the CELL's own; a merged
                        // block draws its far edges from `foot` and `far`,
                        // which belong to other members.
                        eprintln!(
                            "edge row {} col {} box x {}..{} y {}..{} below={below} after={after} left={:?} right={:?} foot={:?} far={:?}",
                            row.index, cell.col, box_.left, box_.right, box_.top, box_.bottom,
                            cell.style.border_left.as_ref().map(|line| line.style.clone()),
                            cell.style.border_right.as_ref().map(|line| line.style.clone()),
                            foot.as_ref().map(|line| line.style.clone()),
                            far.as_ref().map(|line| line.style.clone()),
                        );
                    }
                    let edges: [(&Option<BorderLine>, bool, i32); 4] = [
                        (&cell.style.border_top, true, box_.top),
                        (if below { &None } else { foot }, true, box_.bottom),
                        (&cell.style.border_left, false, box_.left),
                        (if after { &None } else { far }, false, box_.right),
                    ];
                    // A vertical rule does not show inside a horizontal
                    // double's gap. `_xlsx_double_gap.py` rules a pair of
                    // rows on all four sides and puts a double on the
                    // boundary between them: where the verticals cross it the
                    // gap holds the cell's own fill — yellow on a filled
                    // arm, so the gap is the BACKGROUND and not merely
                    // unpainted — and the vertical is not there. A cell's own
                    // rect already stops short of its foot, so the row to
                    // leave out is its head.
                    let gap_at_top = hollow(&cell.style.border_top)
                        || row
                            .index
                            .checked_sub(1)
                            .and_then(|row_at| beside(row_at, cell.col))
                            .is_some_and(|held| hollow(&held.style.border_bottom));
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
                                RECT {
                                    left: at + step,
                                    right: at + step + 1,
                                    top: box_.top + i32::from(gap_at_top),
                                    bottom: box_.bottom,
                                }
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
                    // fallback for one that does not. A face this machine has
                    // not got is answered the way Excel answers it, from the
                    // charset and the family the font record states
                    // (`cell_face_in_place`) — `sanko_tool` asks for
                    // AR P丸ゴシック体E with family 3, and Excel draws it in
                    // fixed-pitch ＭＳ ゴシック where the device's own default
                    // mapper hands back a proportional face whose full-width
                    // marks are narrow.
                    let asked = cell.style.font_name.as_deref().unwrap_or("Calibri");
                    let held = super::cell_face_in_place(
                        asked,
                        cell.style.font_charset,
                        cell.style.font_family,
                    );
                    let name = held.as_str();
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
                    let points_before = asked;
                    let points = if cell.style.shrink_to_fit && !cell.style.wrap_text {
                        // The room the test is made in is the room the line is
                        // actually PLACED in — which for a slanted line is
                        // short by the lean Excel keeps for it (SX87). Testing
                        // against the whole area lets a line that exactly
                        // fills it stand at its asked size: `R6kessan`'s
                        // 「の内数」 measures 48 in ＭＳ Ｐゴシック 11pt bold
                        // italic against an area of 48, so we leave it alone
                        // where Excel comes down to 10.25 and draws it 45 wide.
                        let lean = super::slant_room(asked, cell.style.italic) as f32;
                        super::shrunk_to_fit(
                            name,
                            asked,
                            bold,
                            cell.style.italic,
                            &text,
                            ((area.right - area.left) as f32 - lean) / scale,
                        )
                    } else {
                        asked
                    };
                    // A cell that came down a size keeps the room its DRAWN
                    // size asks for, not the room it asked for before it
                    // shrank: `barrier_free`'s addresses are 11pt cells that
                    // Excel draws at about six, and its gutter is the small
                    // font's two where ours was the asked font's three. The
                    // area was gutted above with the asked size, so only the
                    // difference is taken back here.
                    if points < points_before {
                        let (was_left, was_right) =
                            super::gutters(name, points_before, bold, cell.style.italic);
                        let (now_left, now_right) =
                            super::gutters(name, points, bold, cell.style.italic);
                        area.left -= ((was_left - now_left) * scale).round() as i32;
                        area.right += ((was_right - now_right) * scale).round() as i32;
                    }
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
                    // A row too short for the whole block does not hold a
                    // clipped one: Excel drops the lines that will not start
                    // inside the row, and centres what is left.
                    // `_xlsx_wrap_center_round.py` reads the count off the
                    // picture — a two-line cell shows ONE line until the row
                    // reaches 19 pixels and a three-line one shows two until
                    // it reaches 37, where the line box is 18 — which is one
                    // line for every whole line box the row holds, counting
                    // from its first pixel.
                    let shown = if line_px > 0 {
                        (((box_.bottom - box_.top - 1) / line_px) + 1)
                            .clamp(1, lines.len() as i32) as usize
                    } else {
                        lines.len()
                    };
                    let block = line_px * shown as i32;
                    let slack = (box_.bottom - box_.top) - block;
                    if std::env::var("OXI_XLSX_DUMP_LINES").is_ok() {
                        eprintln!(
                            "   row {} col {} box {}..{} line_px {line_px} lines {} shown {shown} block {block} slack {slack}",
                            row.index, cell.col, box_.top, box_.bottom, lines.len()
                        );
                    }
                    // A merged block carries a pixel of leading under its
                    // text that a plain cell does not, so its text sits a
                    // pixel higher: measured over thirteen row heights and
                    // eight fonts by _xlsx_valign_pixels.py, which is what
                    // put the `h2daa*kre` family a pixel low. A single line
                    // keeps it when centred but not when sat on the bottom,
                    // and several lines with no room to spare lose it again.
                    let merged_block = spans_columns > 0 || spans_rows > 0;
                    let one_line = shown == 1;
                    // WHERE that pixel is taken from depends on how the text
                    // is placed across the cell. A spread line — distributed
                    // or justified — takes it from the FOOT, so a merged one
                    // sat on the bottom rises by it and a merged one centred
                    // does not move; every other alignment takes it from the
                    // block, which is the opposite. 90 arms in
                    // `_xlsx_align_baseline.py` (three row heights, three
                    // vertical alignments, five horizontal ones, merged and
                    // not) part exactly along that line, and it is the pixel
                    // every merged label in `fies_t2`'s worst column stands
                    // below Excel's.
                    let spread_foot = merged_block && placed == Align::Spread;
                    // How far the cell's own bottom rule reaches inside it.
                    let sunk_foot = i32::from(
                        cell.style
                            .border_bottom
                            .as_ref()
                            .is_some_and(|line| super::rule_for(&line.style).before > 0),
                    );
                    let top = box_.top
                        + match cell.style.vertical_align.as_deref() {
                            Some("top") => 0,
                            // A cell told to spread its lines down the box has
                            // only one to spread when it holds one, and Excel
                            // centres it: `shosai_R2`'s `vertical="distributed"`
                            // heading sits at 73 where sitting it on the foot
                            // puts it at 78, and centring gives exactly 73.
                            // Several lines are a different rule and are left
                            // where they were.
                            Some("distributed") | Some("justify") if one_line => {
                                ((slack - i32::from(merged_block)) as f32 / 2.0).floor() as i32
                            }
                            Some("center") | Some("centre") => {
                                // Several lines with no room to spare lose
                                // the pixel again, and are centred as a plain
                                // cell's would be.
                                // The spread cell's pixel is NOT taken out of
                                // the centred block. It reads that way on a
                                // single unwrapped line — `_xlsx_align_baseline.py`
                                // has Excel a pixel below us there — but every
                                // workbook in the corpus that holds centred
                                // distributed cells moves the wrong way for it
                                // (`28C037_2` alone by 0.0041), and
                                // `_xlsx_spread_wrapped.py` agrees on only 8 of
                                // 48 arms once the text wraps. The centred case
                                // is its own question; only the foot is settled.
                                let leading =
                                    i32::from(merged_block && (one_line || slack > 0));
                                // A PLAIN cell halves its leftover toward
                                // zero, not down, and the two only part
                                // company when the block outgrows its row and
                                // the leftover goes below zero.
                                // `_xlsx_wrap_center_round.py` walks the row
                                // a pixel at a time under a wrapped two-line
                                // cell — rows 30 to 51, twenty-two arms — and
                                // every row that cannot hold the block reads
                                // a pixel lower than flooring gives. It is
                                // what puts the `h2daa*_dendeba_kmc` trio's
                                // `C29`, two lines of 11pt in a 25pt row, a
                                // pixel high.
                                //
                                // A cell that does NOT wrap keeps the floor,
                                // and so does a merged block. It is the same
                                // split SX85 found across the cell: wrapping
                                // rounds one way and not wrapping the other.
                                // `001904852/3` are rows of 13pt holding an
                                // 11pt line — a leftover of -1, no wrapping —
                                // and truncating there costs them 0.0259
                                // each; the merged block's own pixel of
                                // leading was measured on that path
                                // (`_xlsx_valign_pixels.py`) and truncating
                                // costs the `h2daa*kre` trio 0.036.
                                if merged_block || !cell.style.wrap_text {
                                    ((slack - leading) as f32 / 2.0).floor() as i32
                                } else {
                                    slack / 2
                                }
                            }
                            // Sat on the bottom, only a block of several
                            // lines gives the pixel up — and a heavy bottom
                            // rule takes one more. `_xlsx_border_room.py`
                            // sweeps the eight styles under a cell sat on its
                            // foot: `medium`, `thick` and `double`, the three
                            // whose ink reaches a pixel INSIDE the cell, lift
                            // the text one pixel; `hair`, `thin`, `dotted`
                            // and `dashed`, which sit on the boundary alone,
                            // do not. The same sweep centred moves for none
                            // of them, so the rule takes the pixel from the
                            // FOOT and not from the box. It is what puts the
                            // `1c*zbd` nine — figures against the foot of
                            // rows ruled with a double — a pixel low.
                            _ => {
                                slack
                                    - i32::from((merged_block && !one_line) || spread_foot)
                                    - sunk_foot
                            }
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
                    // A stacked character is not kept inside the gutters: in
                    // `data_B01` a 15-pixel em stands in a 17-pixel column and
                    // Excel lets its first stroke sit a pixel inside the border,
                    // where the left gutter is three. It is the cell's own edges
                    // that cut it.
                    let (walled, walls) = if cell.style.stacked_text {
                        (box_.left, box_.right)
                    } else {
                        (area.left, area.right)
                    };
                    let clip = CreateRectRgn(walled, box_.top + 1, walls, cut);
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
                        // Not when the cell wraps: the room is trimmed away
                        // with the line's trailing space, so there is nothing
                        // left to make the difference up on.
                        (CellValue::Number(value), Some(format))
                            if format.contains('_') && !cell.style.wrap_text =>
                        {
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
                    // A run's `<rPr>` REPLACES the cell's font; it does not
                    // override parts of it. `_xlsx_run_props.py` puts a bold
                    // 20pt cell behind five hand-written second runs and reads
                    // the ink: no `<rPr>` comes out bold at 20pt like the cell,
                    // an `<rPr>` naming a size and a face but not `<b/>` comes
                    // out REGULAR at 20pt, and an `<rPr>` holding `<b/>` alone
                    // comes out bold at the DEFAULT size rather than the
                    // cell's. So a dressed run wears only what it names.
                    // 1161 of the corpus's 1211 dressed runs leave `<b/>` out
                    // while sitting in a bold cell — `glossary_05` is nine
                    // strings of them — and every one of the 1211 names both
                    // its size and its face, so the fall-back for those two is
                    // not reachable from any workbook here and is left on the
                    // cell rather than guessed at.
                    let worn = |run: &oxicells_core::ir::TextRun| {
                        (
                            run.bold || (bold && !run.dressed),
                            run.italic || (cell.style.italic && !run.dressed),
                            run.underline || (cell.style.underline && !run.dressed),
                        )
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
                                if worn(run).0 { 700 } else { 400 },
                                u32::from(worn(run).1),
                                u32::from(worn(run).2),
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
                            let (heavy, leaning, ruled) = worn(run);
                            (
                                run.font.clone().unwrap_or_else(|| name.to_string()),
                                size,
                                heavy,
                                leaning,
                                ruled,
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
                                // One line to spread is a centred one; see the
                                // plain path above.
                                Some("distributed") | Some("justify") if alone => {
                                    ((slack - i32::from(merged_block)) as f32 / 2.0).floor() as i32
                                }
                                Some("center") | Some("centre") => {
                                    let leading =
                                        i32::from(merged_block && (alone || slack > 0));
                                    // Toward zero for a wrapping cell, down
                                    // for one that does not wrap and for a
                                    // merged block, as above.
                                    if merged_block || !cell.style.wrap_text {
                                        ((slack - leading) as f32 / 2.0).floor() as i32
                                    } else {
                                        slack / 2
                                    }
                                }
                                _ => {
                                    let sunk = i32::from(
                                        cell.style.border_bottom.as_ref().is_some_and(|line| {
                                            super::rule_for(&line.style).before > 0
                                        }),
                                    );
                                    // The spread line takes its merged pixel
                                    // from the foot, as on the plain path.
                                    slack
                                        - i32::from(
                                            (merged_block && !alone)
                                                || (merged_block && placed == Align::Spread),
                                        )
                                        - sunk
                                }
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
                            // The spaces at a distributed line's end are
                            // not spread; see the plain path below.
                            let pieces = super::distribution(if placed == Align::Spread {
                                shown.trim_end_matches(' ')
                            } else {
                                shown.as_str()
                            });
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
                                        // The running total is rounded UP, so
                                        // the spare pixels fall in the EARLY
                                        // gaps: `_xlsx_distributed_round.py`
                                        // spreads three, four and five
                                        // identical glyphs over sixteen cell
                                        // widths and Excel reads [56,55],
                                        // [37,37,36], [28,28,28,27]. (Inferred
                                        // from a real workbook's mixed glyphs
                                        // it looks like the opposite: each
                                        // glyph's own side bearing is in the
                                        // gap you measure that way.)
                                        let want = (spare * piece as i32 + gaps - 1) / gaps;
                                        extra[n] = want - given;
                                        given = want;
                                        left_in_piece = pieces[piece];
                                    }
                                }
                            }
                            let middle = area.left + super::halfway(room - width, cell.style.wrap_text);
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

                    // A stacked cell is drawn through TWO faces. The marks
                    // Excel lays on their side — ー ～ （ ） 「 」 【 】 ＝ 、
                    // 。 and the rest of `turned_in_a_stack` — go through the
                    // vertical face, "@ＭＳ 明朝" turned a quarter turn, which
                    // is where their turned shapes live. Everything standing
                    // goes through the UPRIGHT face, because that is what
                    // Excel draws it with: ＭＳ 明朝 carries an embedded bitmap
                    // at the sizes a sheet uses and a turned face cannot reach
                    // it, so the same 相 comes out 13 pixels across upright and
                    // 10 through the turned face — which is what left every
                    // stacked heading in the `data_*` family with the wrong
                    // glyph, two pixels left and one high, its leading stroke
                    // clipped away at the gutter (`_xlsx_stack_place.py`).
                    //
                    // A standing character sits on the row's own baseline,
                    // `top + baseline - ascent`, measured over seven sizes from
                    // 6 to 36 point by `_xlsx_stack_pen.py`. That is the line
                    // box's padding at every size but 6 point, where the box is
                    // a pixel deeper than the baseline asks for.
                    if cell.style.stacked_text && !dressed_runs {
                        let em = -pixels;
                        // Where each of the two boxes begins. The standing one
                        // starts at the row's baseline less the ascent, which
                        // is to say the character sits on the row's own
                        // baseline. The turned one starts the face's DESCENT
                        // higher — `top + baseline - em`, since a face's ascent
                        // and descent together are its em.
                        //
                        // `_xlsx_stack_turnpen.py` swept 25 sizes in both faces
                        // with the standing character and the mark in ONE cell,
                        // so the cell's centring cancels: every arm read the
                        // descent up, for （ ー and 、 alike. ACROSS is not
                        // settled — Excel's turned pen is one to four pixels
                        // right of the standing one and neither the em less the
                        // turned ascent nor either descent predicts all of it,
                        // so it is left where it was rather than tabulated.
                        let (sit, lay) = super::held(|counter| {
                            counter.shape_of(name, points, bold, cell.style.italic)
                        })
                        .map_or(
                            ((line_px - em).max(0), (line_px - em).max(0)),
                            |(ascent, descent, _)| {
                                (
                                    ((baseline - ascent) * scale).round() as i32,
                                    ((baseline - ascent - descent) * scale).round() as i32,
                                )
                            },
                        );
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
                        // The face everything standing is drawn through —
                        // every kanji and kana, the letters and digits Excel
                        // stacks upright in `01糖尿病`, and the marks its own
                        // class leaves standing even where the turned face has
                        // a rotated shape for them.
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
                        //
                        // The room being halved is a whole number of pixels,
                        // and half an odd one goes to the LEFT. Swept a pixel
                        // at a time from 21 to 49 wide by
                        // `_xlsx_stack_parity.py`: in ＭＳ 明朝 11pt and
                        // ＭＳ ゴシック 9pt the two of us agree on every even
                        // room and ours stands a pixel right on every odd one,
                        // 25 of 26 and 26 of 26. `r03_syukei2`'s headings are
                        // that pixel, in every column it has.
                        let left = match placed {
                            Align::Left | Align::Spread => area.left,
                            Align::Right => area.right - em,
                            Align::Centre => {
                                box_.left + (box_.right - box_.left - em).div_euclid(2)
                            }
                        };
                        if std::env::var("OXI_XLSX_DUMP_STACK").is_ok() {
                            eprintln!(
                                "stack row {} col {} box {}..{} em {em} sit {sit} lay {lay} left {left} lines {lines:?}",
                                row.index, cell.col, box_.left, box_.right
                            );
                        }
                        for (step, line) in lines.iter().enumerate() {
                            let letters = wide(line);
                            let letters = &letters[..letters.len() - 1];
                            if letters.is_empty() {
                                continue;
                            }
                            let head = top + step as i32 * line_px;
                            let stands =
                                line.chars().all(|letter| !super::turned_in_a_stack(letter));
                            if stands {
                                SelectObject(dc, plain);
                                SetTextAlign(dc, TA_TOP | TA_LEFT);
                                let mut measured = SIZE::default();
                                let _ = GetTextExtentPoint32W(dc, letters, &mut measured);
                                // A character narrower than the em is centred
                                // on it — a half-width kana, a letter, a digit.
                                // A full-width one advances a pixel more than
                                // the em at these sizes and is not moved at all.
                                let across =
                                    ((em - measured.cx).max(0) as f32 / 2.0).round() as i32;
                                let _ = TextOutW(dc, left + across, head + sit, letters);
                                SetTextAlign(dc, TA_BASELINE | TA_LEFT);
                                SelectObject(dc, font);
                            } else {
                                let _ = TextOutW(dc, left, head + lay, letters);
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
                            // A distributed line drops the spaces at its end
                            // before it is spread: Excel puts the last glyph
                            // of `"有業人員  "` against the right edge, where
                            // `"有業人員"` puts it. They go from the WIDTH as
                            // well as from the pieces, or the spread comes out
                            // short by their advance
                            // (`_xlsx_distributed_spaces.py`).
                            //
                            // A WRAPPED line drops them too, whatever its
                            // alignment. That is what a number format's
                            // reserved room turns into — `0_)` writes the
                            // blank as a space — so a wrapped `0_)` sits
                            // exactly where a wrapped `0` sits, while an
                            // unwrapped one keeps the room. 42 pairs of arms
                            // in `_xlsx_reserved_align.py` say so across two
                            // faces, seven formats and the three alignments,
                            // and it is the five pixels every company number
                            // in `procurement_contractor_list_02` stands left
                            // of Excel's.
                            let spread_line = if placed == Align::Spread || cell.style.wrap_text {
                                line.trim_end_matches(' ')
                            } else {
                                line.as_str()
                            };
                            let width = width_of(spread_line);
                            let room = area.right - area.left;
                            // A distributed cell fills its whole width, which
                            // is how a Japanese sheet sets a heading: 第 ３ 表,
                            // not 第３表. It is spread by the pieces it could
                            // break a line at, so a Latin word travels whole,
                            // the first piece sits against the left edge and
                            // the last against the right, with nothing kept
                            // back at either end. A single piece is centred
                            // instead — measured on _xlsx_distributed.py.
                            let pieces = super::distribution(spread_line);
                            let spread =
                                placed == Align::Spread && pieces.len() > 1 && room > width;
                            if spread && std::env::var("OXI_XLSX_DUMP_SPREAD").is_ok() {
                                eprintln!(
                                    "spread row {} col {} room {room} width {width} spare {} pieces {:?}",
                                    row.index,
                                    cell.col,
                                    room - width,
                                    pieces
                                );
                            }
                            let middle = area.left + super::halfway(room - width, cell.style.wrap_text);
                            // An italic line leans past its advance, and Excel
                            // keeps room for the lean when the line is put
                            // against the right edge.
                            let lean = super::slant_room(points, cell.style.italic);
                            let left = match placed {
                                Align::Left => area.left + reserved.0,
                                Align::Right => area.right - width - lean - reserved.1,
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
                                        // The running total is rounded UP, so
                                        // the spare pixels fall in the EARLY
                                        // gaps: `_xlsx_distributed_round.py`
                                        // spreads three, four and five
                                        // identical glyphs over sixteen cell
                                        // widths and Excel reads [56,55],
                                        // [37,37,36], [28,28,28,27]. (Inferred
                                        // from a real workbook's mixed glyphs
                                        // it looks like the opposite: each
                                        // glyph's own side bearing is in the
                                        // gap you measure that way.)
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

            // A table rules a line round the whole of itself, and states it
            // as an index into the workbook's differential formats rather
            // than as a border on any cell. Its bottom edge is the only one
            // that usually shows — the other three sit on cell borders that
            // are drawn anyway — and it is long: thirteen hundred pixels in
            // every one of the fifteen `procurement_contractor` workbooks,
            // under the last row of the table, mentioned by no cell.
            for table in &sheet.tables {
                let Some(line) = table.outline.as_ref() else { continue };
                let (Some(left), Some(right)) = (
                    layout.columns.get(
                        (table.start_col as usize).saturating_sub(layout.first_column as usize),
                    ),
                    layout.columns.get(
                        (table.end_col as usize + 1).saturating_sub(layout.first_column as usize),
                    ),
                ) else {
                    continue;
                };
                let (Some(top), Some(bottom)) = (
                    layout.rows.get(
                        (table.start_row as usize).saturating_sub(layout.first_row as usize),
                    ),
                    layout.rows.get(
                        (table.end_row as usize + 1).saturating_sub(layout.first_row as usize),
                    ),
                ) else {
                    continue;
                };
                let box_ = RECT {
                    left: *left as i32,
                    top: *top as i32,
                    right: *right as i32,
                    bottom: *bottom as i32,
                };
                let rule = super::rule_for(&line.style);
                let ink = CreateSolidBrush(colour(line.color.as_deref(), 0x0000_0000));
                for (horizontal, at) in [
                    (true, box_.top),
                    (true, box_.bottom),
                    (false, box_.left),
                    (false, box_.right),
                ] {
                    for step in -rule.before..=rule.after {
                        if rule.hollow && step == 0 {
                            continue;
                        }
                        let edge = if horizontal {
                            RECT { top: at + step, bottom: at + step + 1, ..box_ }
                        } else {
                            RECT { left: at + step, right: at + step + 1, ..box_ }
                        };
                        FillRect(dc, &edge, ink);
                    }
                }
                let _ = DeleteObject(ink);
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
                        shape(
                            dc,
                            held,
                            box_,
                            Exact {
                                room: super::drawing_room(drawn, layout, scale),
                                sides: super::drawing_edges(drawn, layout, scale),
                                down: super::drawing_down(drawn, layout, scale),
                            },
                            scale,
                            sheet.normal_font.as_ref(),
                        )
                    }
                    DrawingKind::Chart(held) => graph(
                        dc,
                        held,
                        box_,
                        Exact {
                            room: super::drawing_room(drawn, layout, scale),
                            sides: super::drawing_edges(drawn, layout, scale),
                            down: super::drawing_down(drawn, layout, scale),
                        },
                        scale,
                        sheet.normal_font.as_ref(),
                    ),
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
                // A note's anchor is written in the file as a cell and a count
                // of pixels into it, and Excel holds that count inside the
                // cell: `_xlsx_anchor_overrun.py` reads back 9 of 9 heights
                // and 7 of 7 widths that stop at the cell's own edge. `002`
                // pins a note ending at row 3 plus 34 pixels where that row is
                // 27 high, and Excel ends it at the top of row 4.
                let Some(box_) = super::anchored_box(
                    &note.from,
                    note.to.as_ref(),
                    Some(extent),
                    layout,
                    scale,
                    true,
                ) else {
                    continue;
                };
                // Excel's border sits OUTSIDE the fill, ours is drawn over the
                // fill's edge, so the box has to be a pixel bigger each way to
                // leave the same fill behind: `_xlsx_note_box.py` reads the
                // fill off Excel's picture at twelve heights and it is always
                // one more than ours was.
                let mut box_ = box_;
                box_.right += 1;
                box_.bottom += 1;
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
                // The insets keep their fractions here as they do in a
                // shape: 2.5mm either side is 9.449 pixels, not 9. `002`'s
                // note is 487 pixels of box, and 469 of room lets 「資」 on to
                // the first line where Excel ends it there and starts the
                // second on 「産」; 468.1 does not.
                says(
                    dc,
                    &note.text,
                    Frame {
                        box_,
                        exact: Some((box_.right - box_.left) as f32),
                        pull: 0.0,
                        edges: None,
                        down: None,
                    },
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

    /// The corner an elbow connector turns at, once the flips and the turn
    /// have been applied. Measured against Excel, `_xlsx_bent_connector.py`:
    /// sixteen arms, every quarter turn against every pair of flips, and the
    /// corner lands on all four corners of the box in this order.
    ///
    /// The box itself does not change shape. A quarter turn about the centre
    /// would swap a box's sides, but the anchor already holds the box the turn
    /// LEAVES the shape in — Excel reports a turned connector's own width and
    /// height the other way round from the anchor it hangs in — so the path is
    /// placed in the anchor's box and nothing is swapped here.
    #[test]
    fn a_turned_elbow_meets_at_the_corner_excel_puts_it() {
        use oxicells_core::ir::Shape;
        let box_ = windows::Win32::Foundation::RECT {
            left: 0, top: 0, right: 128, bottom: 80,
        };
        let elbow = [(0.0, 0.0), (1.0, 0.0), (1.0, 1.0)];
        let bent = |rotation: i32, flip_h: bool, flip_v: bool| {
            let shape = Shape {
                geometry: "bentConnector2".to_string(),
                fill: None,
                line: None,
                adjusts: Vec::new(),
                path: None,
                flip_h,
                flip_v,
                rotation,
                text: None,
            };
            let laid = super::windows_draw::laid(&elbow, &shape, box_);
            (laid[1].x, laid[1].y)
        };
        // Unturned, the bend is at the top right: the path runs along the top
        // and then down the far side.
        assert_eq!(bent(0, false, false), (128, 0));
        assert_eq!(bent(0, true, false), (0, 0));
        assert_eq!(bent(0, false, true), (128, 80));
        assert_eq!(bent(0, true, true), (0, 80));
        // A quarter turn clockwise carries that corner round with it.
        assert_eq!(bent(5_400_000, false, false), (128, 80));
        assert_eq!(bent(10_800_000, false, false), (0, 80));
        assert_eq!(bent(16_200_000, false, false), (0, 0));
        // And the flips happen first: mirrored, then turned.
        assert_eq!(bent(5_400_000, true, false), (128, 0));
        assert_eq!(bent(16_200_000, true, false), (0, 80));
        // A turn that is not a quarter is left where it was written rather
        // than guessed at — the corpus states five turns and all five are one.
        assert_eq!(bent(2_700_000, false, false), (128, 0));
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

    /// Which way a centred line's leftover pixel falls. A wrapping cell keeps
    /// it on the right, a plain one on the left — 28 arms in two faces walked
    /// a pixel at a time by `_xlsx_center_across.py`.
    #[test]
    fn a_wrapping_cell_keeps_the_odd_pixel_on_the_right() {
        assert_eq!(super::halfway(4, false), 2);
        assert_eq!(super::halfway(4, true), 2);
        assert_eq!(super::halfway(5, false), 3);
        assert_eq!(super::halfway(5, true), 2);
        // A line wider than its room hangs the same way round.
        assert_eq!(super::halfway(-5, false), -2);
        assert_eq!(super::halfway(-5, true), -3);
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

    /// Which of a stack's characters Excel lays on their side, and which it
    /// leaves standing. Every character named here was read out of Excel's own
    /// picture by `_xlsx_stack_class.py`, in ＭＳ 明朝 and ＭＳ ゴシック alike:
    /// a standing one is the upright face's ink at the upright face's pen, to
    /// the pixel, and a turned one is not.
    #[test]
    fn a_stack_turns_the_brackets_and_leaves_the_rest_standing() {
        for letter in "、。「」『』（）【】〔〕〈〉《》ー～＝｜［］｛｝＿｢｣ｰ‐―‥…".chars() {
            assert!(super::turned_in_a_stack(letter), "{letter} is turned");
        }
        // Marks the turned face has a rotated shape for, which Excel stands up
        // anyway — the class is Excel's, not the font's.
        for letter in "〜：；＜＞／＼－ｱ".chars() {
            assert!(!super::turned_in_a_stack(letter), "{letter} stands");
        }
        for letter in "相談あウ一二・！？＋Ａａ01①⑧ⅠⅡ々〆".chars() {
            assert!(!super::turned_in_a_stack(letter), "{letter} stands");
        }
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
