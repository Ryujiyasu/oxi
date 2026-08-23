// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use thiserror::Error;

use oxidocs_common::archive::OoxmlArchive;
use oxidocs_common::relationships::parse_relationships;

/// Turns an OOXML custom-filter operator into the prefix VBA states it with.
fn filter_operator(operator: &str) -> &'static str {
    match operator {
        "greaterThan" => ">",
        "greaterThanOrEqual" => ">=",
        "lessThan" => "<",
        "lessThanOrEqual" => "<=",
        "notEqual" => "<>",
        _ => "",
    }
}

/// Records a worksheet feature this parser walks past, so a caller can say
/// what it could not show. The part itself is left alone when the workbook is
/// saved, so nothing is lost — it simply does not reach the IR.
fn note_unsupported(name: &str, noted: &mut Vec<String>) {
    let feature = match name {
        "conditionalFormatting" => "Conditional formatting",
        "dataValidation" | "dataValidations" => "Data validation",
        "hyperlink" | "hyperlinks" => "Hyperlinks",
        "pane" => "Frozen panes",
        "sheetProtection" => "Sheet protection",
        "drawing" => "Drawings",
        "legacyDrawing" => "Comments",
        "tableParts" => "Tables",
        "pivotSelection" => "Pivot tables",
        _ => return,
    };
    if !noted.iter().any(|held| held == feature) {
        noted.push(feature.to_string());
    }
}

/// OOXML writes a boolean attribute as `1`/`0` or `true`/`false`.
fn is_true(value: Option<&str>) -> bool {
    matches!(value, Some("1") | Some("true"))
}

/// Whether a cell format applies a part of itself, rather than taking that
/// part from the named style it is built on.
///
/// Absent is not the same as `0` here. Excel leaves the flag off when the
/// format's own value is the one to use, and writes `applyFont="0"` when it
/// is not — a cell that draws its blue underline from the Hyperlink style
/// says so explicitly. Writers that are not Excel commonly leave every flag
/// off while still naming a font on each cell, and reading absent as "do not
/// apply" throws that font away and draws the whole sheet in the workbook's
/// default face.
fn unless_denied(value: Option<&str>) -> bool {
    !matches!(value, Some("0") | Some("false"))
}
use oxidocs_common::xml_utils::{get_attr, local_name};

use crate::ir::{BorderLine, Cell, CellStyle, CellValue, MergeCell, Row, Sheet, Workbook};

#[derive(Error, Debug)]
pub enum XlsxError {
    #[error("Archive error: {0}")]
    Archive(#[from] oxidocs_common::OxiError),

    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("Invalid cell reference: {0}")]
    InvalidCellRef(String),

    #[error("Invalid data: {0}")]
    InvalidData(String),
}

/// Parse a cell reference like "A1" into (col, row) as 0-based indices.
/// "A1" -> (0, 0), "B2" -> (1, 1), "AA1" -> (26, 0), "AZ3" -> (51, 2)
pub fn parse_cell_ref(s: &str) -> (u32, u32) {
    let mut col: u32 = 0;
    let mut row_str = String::new();
    let mut found_digit = false;

    for ch in s.chars() {
        if ch.is_ascii_alphabetic() && !found_digit {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - b'A' as u32 + 1);
        } else {
            found_digit = true;
            row_str.push(ch);
        }
    }

    let col = if col > 0 { col - 1 } else { 0 };
    let row = row_str.parse::<u32>().unwrap_or(1).saturating_sub(1);

    (col, row)
}

/// Parse a range reference like "A1:C3" into (start_col, start_row, end_col, end_row).
/// Columns are 0-based, rows are 1-based.
fn parse_range_ref(s: &str) -> Option<(u32, u32, u32, u32)> {
    let parts: Vec<&str> = s.split(':').collect();
    if parts.len() != 2 {
        return None;
    }
    let (start_col, start_row_0) = parse_cell_ref(parts[0]);
    let (end_col, end_row_0) = parse_cell_ref(parts[1]);
    // Convert to 1-based rows for MergeCell
    Some((start_col, start_row_0 + 1, end_col, end_row_0 + 1))
}

/// Parse the shared strings table (xl/sharedStrings.xml).
/// Returns a Vec of strings indexed by position.
/// A shared string, and how its parts are dressed when they are not all dressed
/// alike. `text` always holds the whole of it.
#[derive(Debug, Clone, Default)]
struct SharedString {
    text: String,
    runs: Vec<crate::ir::TextRun>,
}

fn parse_shared_strings(xml: &str) -> Result<Vec<SharedString>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut strings: Vec<SharedString> = Vec::new();
    let mut current = SharedString::default();
    let mut in_si = false;
    let mut in_t = false;
    // A phonetic guide (furigana) is stored as an <rPh> element containing its
    // own <t>. It is not part of the cell's text: Excel shows "区分", not
    // "区分クブン". Japanese workbooks carry these on names and addresses
    // constantly, so failing to skip them corrupts a large share of real files.
    let mut in_phonetic = false;
    // A string can be built from <r> runs, each optionally carrying its own
    // <rPr>: an 8pt aside inside an 11pt cell, a raised footnote marker. The
    // run's text goes into the whole string as well, so a reader that ignores
    // the dressing still sees everything.
    let mut in_run = false;
    let mut in_run_props = false;
    let mut current_run = crate::ir::TextRun::default();

    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "si" => {
                        in_si = true;
                        in_phonetic = false;
                        in_run = false;
                        current = SharedString::default();
                    }
                    "rPh" => in_phonetic = true,
                    "r" if in_si && !in_phonetic => {
                        in_run = true;
                        current_run = crate::ir::TextRun::default();
                    }
                    "rPr" if in_run => in_run_props = true,
                    "sz" if in_run_props => {
                        current_run.size = get_attr(&e, "val").and_then(|v| v.parse().ok());
                    }
                    "rFont" | "font" if in_run_props => {
                        current_run.font = get_attr(&e, "val");
                    }
                    "b" if in_run_props => {
                        current_run.bold = get_attr(&e, "val").as_deref() != Some("0");
                    }
                    "i" if in_run_props => {
                        current_run.italic = get_attr(&e, "val").as_deref() != Some("0");
                    }
                    "u" if in_run_props => {
                        current_run.underline = !matches!(
                            get_attr(&e, "val").as_deref(),
                            Some("0") | Some("none")
                        );
                    }
                    "vertAlign" if in_run_props => {
                        current_run.vert_align = get_attr(&e, "val");
                    }
                    "t" if in_si => in_t = true,
                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "si" => {
                        in_si = false;
                        // A string whose runs are all dressed the same as the
                        // cell is no richer than a plain one.
                        if current.runs.iter().all(|run| {
                            run.size.is_none()
                                && run.font.is_none()
                                && run.vert_align.is_none()
                                && !run.bold
                                && !run.italic
                                && !run.underline
                                && run.color.is_none()
                        }) {
                            current.runs.clear();
                        }
                        strings.push(std::mem::take(&mut current));
                    }
                    "rPh" => in_phonetic = false,
                    "rPr" => in_run_props = false,
                    "r" if in_run => {
                        in_run = false;
                        if !current_run.text.is_empty() {
                            current.runs.push(std::mem::take(&mut current_run));
                        }
                    }
                    "t" => in_t = false,
                    _ => {}
                }
            }
            Event::Text(e) => {
                if in_t && in_si && !in_phonetic {
                    // A break inside a cell is written as a pair of
                    // characters by some writers and one by others; the IR
                    // holds one, so the pieces of a string add up to the
                    // string. They did not before: `001290291`'s title came
                    // out 85 characters in its five pieces against 62 in the
                    // cell, and every one of the 23 was the other half of a
                    // break.
                    let text = e.unescape()?.replace("\r\n", "\n");
                    current.text.push_str(&text);
                    if in_run {
                        current_run.text.push_str(&text);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(strings)
}

/// Information about a sheet from workbook.xml
struct SheetInfo {
    name: String,
    r_id: String,
}

/// Parse workbook.xml to extract sheet names and their relationship IDs.
fn parse_workbook_sheets(xml: &str) -> Result<Vec<SheetInfo>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut sheets = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                if name == "sheet" {
                    let sheet_name = get_attr(&e, "name").unwrap_or_default();
                    // r:id attribute — try both namespaced and raw forms
                    let r_id = get_attr(&e, "id")
                        .or_else(|| {
                            // Try raw attribute key "r:id"
                            for attr in e.attributes().flatten() {
                                let key =
                                    std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                                if key == "r:id" {
                                    return Some(
                                        String::from_utf8_lossy(&attr.value).to_string(),
                                    );
                                }
                            }
                            None
                        })
                        .unwrap_or_default();

                    sheets.push(SheetInfo {
                        name: sheet_name,
                        r_id,
                    });
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(sheets)
}

// =====================================================================
// styles.xml parsing
// =====================================================================

#[derive(Debug, Clone, Default)]
struct FontInfo {
    bold: bool,
    italic: bool,
    underline: bool,
    size: Option<f32>,
    color: Option<String>,
    name: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct FillInfo {
    bg_color: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct BorderInfo {
    left: Option<BorderLine>,
    right: Option<BorderLine>,
    top: Option<BorderLine>,
    bottom: Option<BorderLine>,
    diagonal: Option<BorderLine>,
    up: bool,
    down: bool,
}

#[derive(Debug, Clone, Default)]
struct XfRecord {
    num_fmt_id: u32,
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    /// The named style this format is built on, and whether it overrides each
    /// part of it. A cell format that does not apply its own font wears the
    /// font of the style it names — that is how Excel dresses a hyperlink.
    style_id: Option<usize>,
    applies_font: bool,
    applies_fill: bool,
    applies_border: bool,
    applies_number_format: bool,
    horizontal_align: Option<String>,
    vertical_align: Option<String>,
    indent: u32,
    wrap_text: bool,
    stacked_text: bool,
    shrink_to_fit: bool,
}

#[derive(Debug, Clone, Default)]
struct StyleSheet {
    num_fmts: HashMap<u32, String>,
    fonts: Vec<FontInfo>,
    fills: Vec<FillInfo>,
    borders: Vec<BorderInfo>,
    cell_xfs: Vec<XfRecord>,
    cell_style_xfs: Vec<XfRecord>,
}

/// Built-in number format strings for well-known IDs.
/// The formats OOXML gives a number rather than spelling them out, from
/// ECMA-376 §18.8.30. A workbook that uses one of these writes only its id:
/// the accounting formats 37 to 40 are how a Japanese statistical table asks
/// for thousands separators, and reading them as "General" prints 24493 where
/// Excel shows 24,493.
fn builtin_number_format(id: u32) -> Option<&'static str> {
    match id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        5 => Some("$#,##0_);($#,##0)"),
        6 => Some("$#,##0_);[Red]($#,##0)"),
        7 => Some("$#,##0.00_);($#,##0.00)"),
        8 => Some("$#,##0.00_);[Red]($#,##0.00)"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("mm-dd-yy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        // Excel's own `NumberFormat` for these four, asked of a cell that
        // wears one: the room after the number is the width of a bracket,
        // not of a space, which is two pixels in ＭＳ 11 and shows up as a
        // right-aligned number sitting two pixels off.
        37 => Some("#,##0_);(#,##0)"),
        38 => Some("#,##0_);[Red](#,##0)"),
        39 => Some("#,##0.00_);(#,##0.00)"),
        40 => Some("#,##0.00_);[Red](#,##0.00)"),
        // The Japanese locale fills in 27 to 36 and 50 to 58, which is what a
        // government workbook's dates are written with. The era forms — ggge,
        // 令和 — are given their Gregorian equivalent here, because the
        // formatter has no calendar to name an era from; a date in the right
        // shape is nearer Excel's ink than the serial number that reading
        // none of them at all leaves behind.
        27 | 36 | 50 | 57 => Some("yyyy\".\"m\".\"d"),
        28 | 29 | 51 | 54 | 58 => Some("yyyy\"年\"m\"月\"d\"日\""),
        30 => Some("m/d/yy"),
        31 => Some("yyyy\"年\"m\"月\"d\"日\""),
        32 => Some("h\"時\"mm\"分\""),
        33 => Some("h\"時\"mm\"分\"ss\"秒\""),
        34 | 52 | 55 => Some("yyyy\"年\"m\"月\""),
        35 | 53 | 56 => Some("m\"月\"d\"日\""),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mmss.0"),
        48 => Some("##0.0E+0"),
        49 => Some("@"),
        _ => None,
    }
}

/// The colours a workbook's theme names, in the order the theme states them:
/// dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink.
#[derive(Debug, Clone, Default)]
pub(crate) struct Theme {
    colours: Vec<String>,
    /// The palette the workbook states in place of Excel's own, if it does.
    /// It lives in the styles part rather than the theme, and is read before
    /// the styles that use it.
    indexed: Vec<String>,
    /// The faces the theme's major and minor schemes name for this script.
    /// A font that says `<scheme val="minor"/>` wears one of these and its
    /// own `<name>` counts for nothing.
    major_face: Option<String>,
    minor_face: Option<String>,
}

impl Theme {
    /// A `theme="N"` is not an index into that order. Excel counts the first
    /// two the other way round, and the second two as well, so 0 is lt1 and 1
    /// is dk1.
    fn colour(&self, index: usize) -> Option<&str> {
        const ORDER: [usize; 12] = [1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11];
        let slot = *ORDER.get(index)?;
        self.colours.get(slot).map(String::as_str)
    }
}

/// The script whose face a theme font resolves to. Excel picks it by the
/// language it is running as; measured against a Japanese Excel, which is
/// what the corpus is compared to.
const THEME_SCRIPT: &str = "Jpan";

fn parse_theme_xml(xml: &str) -> Theme {
    let mut reader = Reader::from_str(xml);
    let mut theme = Theme::default();
    let mut in_scheme = false;
    // Which of the two font schemes is being read, and how good the face
    // held for each is: the entry for this script beats the East Asian
    // fallback, which beats the Latin one.
    let mut in_font_scheme: Option<bool> = None;
    let mut face_rank = [0u8; 2];
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()).as_str() {
                "clrScheme" => in_scheme = true,
                "majorFont" => in_font_scheme = Some(true),
                "minorFont" => in_font_scheme = Some(false),
                _ => {}
            },
            Ok(Event::End(e)) => match local_name(e.name().as_ref()).as_str() {
                // The colour scheme comes first; reading on gathers the
                // fonts that follow it.
                "clrScheme" => in_scheme = false,
                "majorFont" | "minorFont" => in_font_scheme = None,
                "theme" | "themeElements" => break,
                _ => {}
            },
            Ok(Event::Empty(e)) if in_font_scheme.is_some() => {
                let major = in_font_scheme == Some(true);
                let name = local_name(e.name().as_ref());
                // Rank: the script's own entry, then <a:ea>, then <a:latin>.
                let rank = match name.as_str() {
                    "font" if get_attr(&e, "script").as_deref() == Some(THEME_SCRIPT) => 3,
                    "ea" => 2,
                    "latin" => 1,
                    _ => 0,
                };
                let face = get_attr(&e, "typeface").filter(|face| !face.is_empty());
                let held = &mut face_rank[usize::from(major)];
                if let (Some(face), true) = (face, rank > *held) {
                    *held = rank;
                    if major {
                        theme.major_face = Some(face);
                    } else {
                        theme.minor_face = Some(face);
                    }
                }
            }
            Ok(Event::Empty(e)) if in_scheme => {
                // A scheme colour is either stated outright or taken from the
                // system, which records what it last resolved to.
                let colour = match local_name(e.name().as_ref()).as_str() {
                    "srgbClr" => get_attr(&e, "val"),
                    "sysClr" => get_attr(&e, "lastClr"),
                    _ => None,
                };
                if let Some(colour) = colour {
                    theme.colours.push(colour);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    theme
}

/// Moves a colour toward white or black, the way a theme tint does. Only how
/// light it is changes; its hue and how saturated it is do not, which is why
/// this goes through HSL rather than nudging each channel.
fn tinted(hex: &str, tint: f32) -> String {
    let Ok(value) = u32::from_str_radix(hex, 16) else {
        return hex.to_string();
    };
    let channel = |shift: u32| ((value >> shift) & 0xFF) as f32 / 255.0;
    let (red, green, blue) = (channel(16), channel(8), channel(0));
    let high = red.max(green).max(blue);
    let low = red.min(green).min(blue);
    let lum = (high + low) / 2.0;
    let span = high - low;
    let sat = if span == 0.0 {
        0.0
    } else if lum < 0.5 {
        span / (high + low)
    } else {
        span / (2.0 - high - low)
    };
    let hue = if span == 0.0 {
        0.0
    } else if high == red {
        (((green - blue) / span) % 6.0 + 6.0) % 6.0
    } else if high == green {
        (blue - red) / span + 2.0
    } else {
        (red - green) / span + 4.0
    };

    let moved = if tint < 0.0 {
        lum * (1.0 + tint)
    } else {
        lum * (1.0 - tint) + tint
    };

    let chroma = (1.0 - (2.0 * moved - 1.0).abs()) * sat;
    let second = chroma * (1.0 - ((hue % 2.0) - 1.0).abs());
    let (red, green, blue) = match hue as u32 {
        0 => (chroma, second, 0.0),
        1 => (second, chroma, 0.0),
        2 => (0.0, chroma, second),
        3 => (0.0, second, chroma),
        4 => (second, 0.0, chroma),
        _ => (chroma, 0.0, second),
    };
    let base = moved - chroma / 2.0;
    let byte = |part: f32| (((part + base).clamp(0.0, 1.0) * 255.0).round() as u32).min(255);
    format!("{:02X}{:02X}{:02X}", byte(red), byte(green), byte(blue))
}

/// The colours a workbook may name by number rather than by value.
///
/// Excel kept a palette of 56 from the days when a file could hold no more,
/// and files still write `<color indexed="12"/>` for blue. A workbook may
/// state its own in `<indexedColors>`, which is why `Theme` carries one.
/// 64 and 65 are not in the palette at all: they mean the system's own
/// foreground and background, which the caller decides.
const PALETTE: [&str; 56] = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
];

/// The colour a number stands for, from the workbook's own palette where it
/// states one and Excel's otherwise.
fn indexed_colour(index: usize, theme: &Theme) -> Option<String> {
    if let Some(own) = theme.indexed.get(index) {
        return Some(own.clone());
    }
    PALETTE.get(index).map(|hex| hex.to_string())
}

fn parse_color_attr(e: &quick_xml::events::BytesStart, theme: &Theme) -> Option<String> {
    if let Some(rgb) = get_attr(e, "rgb") {
        // Strip leading alpha if 8-char hex
        let hex = if rgb.len() == 8 { &rgb[2..] } else { &rgb };
        return Some(hex.to_string());
    }
    if let Some(index) = get_attr(e, "indexed").and_then(|at| at.parse::<usize>().ok()) {
        return indexed_colour(index, theme);
    }
    let index = get_attr(e, "theme")?.parse::<usize>().ok()?;
    let named = theme.colour(index)?;
    let tint = get_attr(e, "tint")
        .and_then(|value| value.parse::<f32>().ok())
        .unwrap_or(0.0);
    Some(if tint == 0.0 {
        named.to_string()
    } else {
        tinted(named, tint)
    })
}

/// The palette a workbook states in place of Excel's own, empty when it
/// states none. `<rgbColor rgb="00RRGGBB"/>`, in order from index zero.
fn indexed_palette(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut inside = false;
    let mut held = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                match local_name(e.name().as_ref()).as_str() {
                    "indexedColors" => inside = true,
                    "rgbColor" if inside => {
                        if let Some(rgb) = get_attr(e, "rgb") {
                            let hex = if rgb.len() == 8 { &rgb[2..] } else { &rgb };
                            held.push(hex.to_string());
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                if local_name(e.name().as_ref()) == "indexedColors" {
                    break;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    held
}

fn parse_styles_xml(xml: &str, theme: &Theme) -> Result<StyleSheet, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut ss = StyleSheet::default();

    // Parsing state
    #[derive(PartialEq)]
    enum Section {
        None,
        NumFmts,
        Fonts,
        Fills,
        Borders,
        CellXfs,
        CellStyleXfs,
    }
    let mut section = Section::None;
    let mut in_font = false;
    let mut current_font = FontInfo::default();
    let mut current_font_scheme: Option<String> = None;
    let mut in_fill = false;
    let mut current_fill = FillInfo::default();
    let mut in_border = false;
    let mut current_border = BorderInfo::default();
    let mut in_xf = false;
    let mut current_xf = XfRecord::default();

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "numFmts" => section = Section::NumFmts,
                    "fonts" => section = Section::Fonts,
                    "fills" => section = Section::Fills,
                    "borders" => section = Section::Borders,
                    "cellXfs" => section = Section::CellXfs,
                    "cellStyleXfs" => section = Section::CellStyleXfs,

                    "font" if section == Section::Fonts => {
                        in_font = true;
                        current_font = FontInfo::default();
                    }
                    "fill" if section == Section::Fills => {
                        in_fill = true;
                        current_fill = FillInfo::default();
                    }
                    "border" if section == Section::Borders => {
                        in_border = true;
                        current_border = BorderInfo::default();
                        // Which way the corner-to-corner rule runs is stated
                        // on the border itself, and a cell can carry both.
                        current_border.up = is_true(get_attr(&e, "diagonalUp").as_deref());
                        current_border.down = is_true(get_attr(&e, "diagonalDown").as_deref());
                    }
                    "xf" if section == Section::CellXfs
                        || section == Section::CellStyleXfs =>
                    {
                        in_xf = true;
                        current_xf = XfRecord {
                            num_fmt_id: get_attr(&e, "numFmtId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            font_id: get_attr(&e, "fontId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            fill_id: get_attr(&e, "fillId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            border_id: get_attr(&e, "borderId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            horizontal_align: None,
                            indent: 0,
                            vertical_align: None,
                            wrap_text: false,
                            stacked_text: false,
                            shrink_to_fit: false,
                            style_id: get_attr(&e, "xfId").and_then(|v| v.parse().ok()),
                            applies_font: unless_denied(
                                get_attr(&e, "applyFont").as_deref(),
                            ),
                            applies_fill: unless_denied(
                                get_attr(&e, "applyFill").as_deref(),
                            ),
                            applies_border: unless_denied(
                                get_attr(&e, "applyBorder").as_deref(),
                            ),
                            applies_number_format: unless_denied(
                                get_attr(&e, "applyNumberFormat").as_deref(),
                            ),
                        };
                    }
                    "alignment" if in_xf => {
                        current_xf.horizontal_align = get_attr(&e, "horizontal");
                        current_xf.vertical_align = get_attr(&e, "vertical");
                        current_xf.indent = get_attr(&e, "indent")
                            .and_then(|level| level.parse().ok())
                            .unwrap_or(0);
                        current_xf.wrap_text =
                            matches!(get_attr(&e, "wrapText").as_deref(), Some("1") | Some("true"));
                        // 255 is not an angle: it is Excel's way of saying the
                        // characters stand one above the next.
                        current_xf.stacked_text =
                            get_attr(&e, "textRotation").as_deref() == Some("255");
                        current_xf.shrink_to_fit =
                            is_true(get_attr(&e, "shrinkToFit").as_deref());
                    }

                    // Inside a border element, parse child elements with style attr
                    "left" if in_border => {
                        current_border.left = border_line(&e);
                    }
                    "right" if in_border => {
                        current_border.right = border_line(&e);
                    }
                    "top" if in_border => {
                        current_border.top = border_line(&e);
                    }
                    "bottom" if in_border => {
                        current_border.bottom = border_line(&e);
                    }
                    // A diagonal that states a colour has a child, so it
                    // arrives here and not among the self-closing sides.
                    "diagonal" if in_border => {
                        current_border.diagonal = border_line(&e);
                    }

                    // Font color
                    "color" if in_font => {
                        if let Some(c) = parse_color_attr(&e, theme) {
                            current_font.color = Some(c);
                        }
                    }

                    // Fill color — look for fgColor inside patternFill
                    "fgColor" if in_fill => {
                        if let Some(c) = parse_color_attr(&e, theme) {
                            current_fill.bg_color = Some(c);
                        }
                    }

                    _ => {}
                }
            }
            Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "numFmt" if section == Section::NumFmts => {
                        if let (Some(id_str), Some(code)) =
                            (get_attr(&e, "numFmtId"), get_attr(&e, "formatCode"))
                        {
                            if let Ok(id) = id_str.parse::<u32>() {
                                ss.num_fmts.insert(id, code);
                            }
                        }
                    }
                    // Self-closing <b/>, <i/>, <sz val="..."/>
                    "b" if in_font => {
                        // <b/> means bold=true, <b val="0"/> means false
                        let val = get_attr(&e, "val");
                        current_font.bold = val.as_deref() != Some("0");
                    }
                    "i" if in_font => {
                        let val = get_attr(&e, "val");
                        current_font.italic = val.as_deref() != Some("0");
                    }
                    "u" if in_font => {
                        // <u/> underlines; <u val="none"/> does not.
                        let val = get_attr(&e, "val");
                        current_font.underline = !matches!(
                            val.as_deref(),
                            Some("0") | Some("none")
                        );
                    }
                    "sz" if in_font => {
                        current_font.size =
                            get_attr(&e, "val").and_then(|v| v.parse().ok());
                    }
                    "name" | "rFont" if in_font => {
                        current_font.name = get_attr(&e, "val");
                    }
                    "scheme" if in_font => {
                        current_font_scheme = get_attr(&e, "val");
                    }
                    "color" if in_font => {
                        if let Some(c) = parse_color_attr(&e, theme) {
                            current_font.color = Some(c);
                        }
                    }
                    "fgColor" if in_fill => {
                        if let Some(c) = parse_color_attr(&e, theme) {
                            current_fill.bg_color = Some(c);
                        }
                    }
                    // Self-closing border sides: <left style="thin"/>
                    "left" if in_border => {
                        current_border.left = border_line(&e);
                    }
                    "right" if in_border => {
                        current_border.right = border_line(&e);
                    }
                    "top" if in_border => {
                        current_border.top = border_line(&e);
                    }
                    "bottom" if in_border => {
                        current_border.bottom = border_line(&e);
                    }
                    "diagonal" if in_border => {
                        current_border.diagonal = border_line(&e);
                    }
                    "alignment" if in_xf => {
                        current_xf.horizontal_align = get_attr(&e, "horizontal");
                        current_xf.vertical_align = get_attr(&e, "vertical");
                        current_xf.indent = get_attr(&e, "indent")
                            .and_then(|level| level.parse().ok())
                            .unwrap_or(0);
                        current_xf.wrap_text =
                            matches!(get_attr(&e, "wrapText").as_deref(), Some("1") | Some("true"));
                        // 255 is not an angle: it is Excel's way of saying the
                        // characters stand one above the next.
                        current_xf.stacked_text =
                            get_attr(&e, "textRotation").as_deref() == Some("255");
                        current_xf.shrink_to_fit =
                            is_true(get_attr(&e, "shrinkToFit").as_deref());
                    }
                    "xf" if section == Section::CellXfs
                        || section == Section::CellStyleXfs =>
                    {
                        // Self-closing <xf ... />
                        let xf = XfRecord {
                            num_fmt_id: get_attr(&e, "numFmtId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            font_id: get_attr(&e, "fontId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            fill_id: get_attr(&e, "fillId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            border_id: get_attr(&e, "borderId")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(0),
                            horizontal_align: None,
                            indent: 0,
                            vertical_align: None,
                            wrap_text: false,
                            stacked_text: false,
                            shrink_to_fit: false,
                            style_id: get_attr(&e, "xfId").and_then(|v| v.parse().ok()),
                            applies_font: unless_denied(
                                get_attr(&e, "applyFont").as_deref(),
                            ),
                            applies_fill: unless_denied(
                                get_attr(&e, "applyFill").as_deref(),
                            ),
                            applies_border: unless_denied(
                                get_attr(&e, "applyBorder").as_deref(),
                            ),
                            applies_number_format: unless_denied(
                                get_attr(&e, "applyNumberFormat").as_deref(),
                            ),
                        };
                        if section == Section::CellStyleXfs {
                            ss.cell_style_xfs.push(xf);
                        } else {
                            ss.cell_xfs.push(xf);
                        }
                    }

                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "numFmts" | "fonts" | "fills" | "borders" | "cellXfs"
                    | "cellStyleXfs" => {
                        section = Section::None;
                    }
                    "font" if in_font => {
                        // A font that names a theme scheme wears the face
                        // that scheme states, and its own <name> counts for
                        // nothing: openpyxl writes Calibri with
                        // `scheme="minor"` and Excel opens the workbook in
                        // the theme's ＭＳ Ｐゴシック, rows and all.
                        let face = match current_font_scheme.take().as_deref() {
                            Some("major") => theme.major_face.clone(),
                            Some("minor") => theme.minor_face.clone(),
                            _ => None,
                        };
                        if let Some(face) = face {
                            current_font.name = Some(face);
                        }
                        ss.fonts.push(std::mem::take(&mut current_font));
                        in_font = false;
                    }
                    "fill" if in_fill => {
                        ss.fills.push(std::mem::take(&mut current_fill));
                        in_fill = false;
                    }
                    "border" if in_border => {
                        ss.borders.push(std::mem::take(&mut current_border));
                        in_border = false;
                    }
                    "xf" if in_xf
                        && (section == Section::CellXfs
                            || section == Section::CellStyleXfs) =>
                    {
                        let xf = std::mem::take(&mut current_xf);
                        if section == Section::CellStyleXfs {
                            ss.cell_style_xfs.push(xf);
                        } else {
                            ss.cell_xfs.push(xf);
                        }
                        in_xf = false;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(ss)
}

/// An edge that names no style is not drawn; one that does carries how.
fn border_line(e: &quick_xml::events::BytesStart) -> Option<BorderLine> {
    let style = get_attr(e, "style")?;
    if style == "none" {
        return None;
    }
    Some(BorderLine { style, color: None })
}

/// Build a CellStyle from a style index referencing the StyleSheet.
fn resolve_cell_style(style_index: usize, stylesheet: &StyleSheet) -> CellStyle {
    let xf = match stylesheet.cell_xfs.get(style_index) {
        Some(xf) => xf,
        None => return CellStyle::default(),
    };

    // A cell format that does not apply a part of itself wears that part from
    // the named style it is built on. This is how a hyperlink gets its blue
    // underline: the cell's own format names no font at all.
    let parent = xf
        .style_id
        .and_then(|id| stylesheet.cell_style_xfs.get(id));
    let inherited = |applies: bool, own: usize, from: fn(&XfRecord) -> usize| {
        match (applies, parent) {
            (false, Some(parent)) => from(parent),
            _ => own,
        }
    };
    let font_id = inherited(xf.applies_font, xf.font_id, |xf| xf.font_id);
    let fill_id = inherited(xf.applies_fill, xf.fill_id, |xf| xf.fill_id);
    let border_id = inherited(xf.applies_border, xf.border_id, |xf| xf.border_id);

    let mut font = stylesheet.fonts.get(font_id).cloned().unwrap_or_default();
    // A font record that names no face, or no size, is not asking for the
    // reader's own default: it is saying "the workbook's font, in bold" —
    // which is how a generated header row is written. The parts it leaves out
    // come from the Normal style's font.
    if font.name.is_none() || font.size.is_none() {
        let normal = stylesheet
            .cell_style_xfs
            .first()
            .map(|xf| xf.font_id)
            .and_then(|id| stylesheet.fonts.get(id));
        if let Some(normal) = normal {
            font.name = font.name.or_else(|| normal.name.clone());
            font.size = font.size.or(normal.size);
        }
    }
    let fill = stylesheet.fills.get(fill_id).cloned().unwrap_or_default();
    let border = stylesheet
        .borders
        .get(border_id)
        .cloned()
        .unwrap_or_default();

    // Resolve number format
    let num_fmt_id = match (xf.applies_number_format, parent) {
        (false, Some(parent)) => parent.num_fmt_id,
        _ => xf.num_fmt_id,
    };
    let number_format = if num_fmt_id == 0 {
        None // General — no explicit format needed
    } else if let Some(custom) = stylesheet.num_fmts.get(&num_fmt_id) {
        Some(custom.clone())
    } else {
        builtin_number_format(num_fmt_id).map(|s| s.to_string())
    };

    CellStyle {
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        font_size: font.size,
        font_name: font.name.clone(),
        font_color: font.color,
        bg_color: fill.bg_color,
        number_format,
        horizontal_align: xf.horizontal_align.clone(),
        indent: xf.indent,
        vertical_align: xf.vertical_align.clone(),
        wrap_text: xf.wrap_text,
        stacked_text: xf.stacked_text,
        shrink_to_fit: xf.shrink_to_fit,
        border_top: border.top.clone(),
        border_bottom: border.bottom.clone(),
        border_left: border.left.clone(),
        border_right: border.right.clone(),
        border_diagonal: border.diagonal.clone(),
        diagonal_up: border.up,
        diagonal_down: border.down,
    }
}

/// The font a row's own format puts on it, when the row carries one
/// (`s=` with customFormat="1").
fn row_style_font(
    e: &quick_xml::events::BytesStart,
    stylesheet: &StyleSheet,
) -> Option<(String, f32)> {
    if !is_true(get_attr(e, "customFormat").as_deref()) {
        return None;
    }
    let si = get_attr(e, "s")?.parse::<usize>().ok()?;
    let style = resolve_cell_style(si, stylesheet);
    Some((style.font_name?, style.font_size?))
}

/// Parse a single worksheet XML into a Sheet.
/// Where a relationship target lands, given the part that names it. Targets
/// are written relative to the naming part's own folder, so `../tables/t.xml`
/// beside `xl/worksheets/sheet1.xml` is `xl/tables/t.xml`.
fn part_beside(from: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.trim_start_matches('/').to_string();
    }
    let mut parts: Vec<&str> = from.rsplit_once('/').map_or("", |(dir, _)| dir).split('/').collect();
    parts.retain(|part| !part.is_empty());
    for step in target.split('/') {
        match step {
            "." | "" => {}
            ".." => {
                parts.pop();
            }
            step => parts.push(step),
        }
    }
    parts.join("/")
}

/// Reads one `xl/tables/*.xml` part into the range and dress it describes.
/// What a sheet's drawing part holds, with each picture's relationship id.
///
/// The anchors come in three kinds: hung from two cells, hung from one with a
/// size of its own, or placed at a fixed spot on the sheet. Only the first two
/// are common in the corpus, and an absolute one is read as a one-cell anchor
/// at the top-left, which is where its offsets are measured from anyway.
fn parse_drawing_xml(xml: &str, theme: &Theme) -> Vec<(crate::ir::Drawing, Option<String>)> {
    use crate::ir::{
        Anchor, Drawing, DrawingKind, Shape, ShapeLine, ShapeParagraph, ShapeText,
    };

    /// Which part of the shape a colour being read belongs to.
    #[derive(Clone, Copy, PartialEq)]
    enum Paints {
        Fill,
        Line,
        Text,
    }

    let mut reader = Reader::from_str(xml);
    let mut buf = Vec::new();
    let mut found: Vec<(Drawing, Option<String>)> = Vec::new();
    // Where a chart's own shape sits, as fractions of the chart's box.
    let mut frame: Option<(f64, f64)> = None;
    let mut frame_to: Option<(f64, f64)> = None;

    let blank = Anchor { col: 0, col_off: 0, row: 0, row_off: 0 };
    let mut from = blank;
    let mut to: Option<Anchor> = None;
    let mut extent: Option<(i64, i64)> = None;
    let mut kind: Option<DrawingKind> = None;
    let blank_shape = Shape {
        geometry: String::new(),
        fill: None,
        line: None,
        flip_h: false,
        flip_v: false,
        text: None,
    };
    let mut shape = blank_shape.clone();
    let mut line_width: i64 = 9525;
    let mut dash: Option<String> = None;
    let mut embed: Option<String> = None;
    // Which corner the col/row elements belong to, and whether we are inside
    // a shape — where `<ext>` is the shape's own size, not the anchor's.
    let mut corner: Option<bool> = None;
    let mut depth_in_shape = 0usize;
    let mut number = String::new();
    // Where a colour is being read from, and what is being read into.
    let (mut in_sp_pr, mut in_ln, mut in_style) = (false, false, false);
    // An extension list holds what a shape would be painted with if it were
    // painted at all: Excel writes the fill a shape had before it was set to
    // none into an `a14:hiddenFill`, and reading that paints white boxes over
    // the sheet. Nothing inside one is a shape's own paint.
    let mut in_ext_lst = 0usize;
    // Whether the shape itself said what it is painted and ruled with, in
    // which case its style's references are not consulted.
    let (mut fill_stated, mut line_stated) = (false, false);
    let mut paints: Option<Paints> = None;
    let mut colour: Option<(String, Vec<(String, f32)>)> = None;
    // What the shape says: the body's own settings, the paragraph being read,
    // and whether a run's text is what the reader is collecting.
    let mut said = ShapeText {
        paragraphs: Vec::new(),
        anchor: None,
        insets: (91440, 45720, 91440, 45720),
        wrap: true,
        clip: false,
    };
    let mut paragraph: Option<ShapeParagraph> = None;
    let mut in_run_props = false;
    let mut in_text = false;
    let mut in_line_spacing = false;
    let mut text = String::new();
    // A group states where it sits and what its children's coordinates mean;
    // a child inside one is placed by mapping its own box through that.
    let mut group: Option<((i64, i64), (i64, i64), (i64, i64), (i64, i64))> = None;
    let mut in_group_props = false;
    let mut own_off: Option<(i64, i64)> = None;
    let mut own_ext: Option<(i64, i64)> = None;

    loop {
        let event = reader.read_event_into(&mut buf);
        let (start, empty) = match &event {
            Ok(Event::Start(e)) => (Some(e), false),
            Ok(Event::Empty(e)) => (Some(e), true),
            _ => (None, false),
        };
        if let Some(e) = start {
            let name = local_name(e.name().as_ref());
            match name.as_str() {
                "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor"
                // A shape that hangs on a chart rather than on the grid
                // is anchored by fractions of the chart's own box.
                | "relSizeAnchor" | "absSizeAnchor" => {
                    frame = None;
                    frame_to = None;
                    from = blank;
                    to = None;
                    extent = None;
                    kind = None;
                    embed = None;
                    depth_in_shape = 0;
                    shape = blank_shape.clone();
                    line_width = 9525;
                    dash = None;
                    fill_stated = false;
                    line_stated = false;
                    said = ShapeText {
                        paragraphs: Vec::new(),
                        anchor: None,
                        insets: (91440, 45720, 91440, 45720),
                        wrap: true,
                        clip: false,
                    };
                    paragraph = None;
                    group = None;
                    own_off = None;
                    own_ext = None;
                }
                "from" => corner = Some(true),
                "to" => {
                    corner = Some(false);
                    to = Some(blank);
                }
                "col" | "colOff" | "row" | "rowOff" => number.clear(),
                "x" | "y" if corner.is_some() => number.clear(),
                "ext" => {
                    let cx = get_attr(e, "cx").and_then(|v| v.parse().ok()).unwrap_or(0);
                    let cy = get_attr(e, "cy").and_then(|v| v.parse().ok()).unwrap_or(0);
                    if depth_in_shape == 0 && !in_group_props {
                        extent = Some((cx, cy));
                    } else if in_group_props {
                        group.get_or_insert(((0, 0), (0, 0), (0, 0), (0, 0))).1 = (cx, cy);
                    } else if in_sp_pr {
                        own_ext = Some((cx, cy));
                    }
                }
                "pic" | "sp" | "cxnSp" => {
                    depth_in_shape += 1;
                    // A child of a group starts its own shape; one that is
                    // not carries the anchor's.
                    if group.is_some() {
                        shape = blank_shape.clone();
                        said.paragraphs.clear();
                        paragraph = None;
                        fill_stated = false;
                        line_stated = false;
                        own_off = None;
                        own_ext = None;
                        embed = None;
                    }
                    if name == "pic" {
                        kind = Some(DrawingKind::Picture { bytes: Vec::new() });
                    } else if kind.is_none() || group.is_some() {
                        kind = Some(DrawingKind::Shape(shape.clone()));
                    }
                }
                "graphicFrame" => {
                    depth_in_shape += 1;
                    kind = Some(DrawingKind::Chart(Default::default()));
                }
                // The frame holds no picture of its own: it names the part
                // that does, the way a picture names its bytes.
                "chart" if embed.is_none() => {
                    embed = get_attr(e, "id");
                }
                "grpSp" => {
                    depth_in_shape += 1;
                    if kind.is_none() {
                        kind = Some(DrawingKind::Other);
                    }
                }
                "grpSpPr" => in_group_props = true,
                "chOff" | "chExt" | "off" if in_group_props || in_sp_pr => {
                    let x = get_attr(e, "x").or_else(|| get_attr(e, "cx"));
                    let y = get_attr(e, "y").or_else(|| get_attr(e, "cy"));
                    let pair = (
                        x.and_then(|v| v.parse().ok()).unwrap_or(0),
                        y.and_then(|v| v.parse().ok()).unwrap_or(0),
                    );
                    if in_group_props {
                        let held = group.get_or_insert(((0, 0), (0, 0), (0, 0), (0, 0)));
                        match name.as_str() {
                            "off" => held.0 = pair,
                            "chOff" => held.2 = pair,
                            _ => held.3 = pair,
                        }
                    } else if name == "off" {
                        own_off = Some(pair);
                    }
                }
                "spPr" => in_sp_pr = true,
                "style" => in_style = true,
                "extLst" => in_ext_lst += 1,
                "ln" if in_sp_pr && in_ext_lst == 0 => {
                    in_ln = true;
                    line_width = get_attr(e, "w").and_then(|w| w.parse().ok()).unwrap_or(9525);
                }
                "xfrm" if in_sp_pr => {
                    shape.flip_h = is_true(get_attr(e, "flipH").as_deref());
                    shape.flip_v = is_true(get_attr(e, "flipV").as_deref());
                }
                "prstGeom" => {
                    if let Some(preset) = get_attr(e, "prst") {
                        shape.geometry = preset;
                    }
                }
                "prstDash" if in_ln => dash = get_attr(e, "val"),
                "noFill" if in_sp_pr && in_ext_lst == 0 => {
                    if in_ln {
                        line_stated = true;
                        shape.line = None;
                    } else {
                        fill_stated = true;
                        shape.fill = None;
                    }
                }
                // A shape that states no fill or line of its own wears the
                // one its style names, which is a theme colour.
                "solidFill" if in_sp_pr && in_ext_lst == 0 => {
                    if in_ln {
                        line_stated = true;
                        paints = Some(Paints::Line);
                    } else {
                        fill_stated = true;
                        paints = Some(Paints::Fill);
                    }
                }
                "fillRef" if in_style && !fill_stated => paints = Some(Paints::Fill),
                "lnRef" if in_style && !line_stated => paints = Some(Paints::Line),
                "solidFill" if in_run_props => paints = Some(Paints::Text),
                "srgbClr" | "schemeClr" | "sysClr" if paints.is_some() => {
                    let named = match name.as_str() {
                        "sysClr" => get_attr(e, "lastClr"),
                        "srgbClr" => get_attr(e, "val"),
                        _ => get_attr(e, "val").and_then(|val| scheme_colour(&val, theme)),
                    };
                    colour = named.map(|hex| (hex, Vec::new()));
                }
                "lumMod" | "lumOff" | "shade" | "tint" if colour.is_some() => {
                    if let (Some((_, mods)), Some(value)) = (
                        colour.as_mut(),
                        get_attr(e, "val").and_then(|v| v.parse::<f32>().ok()),
                    ) {
                        mods.push((name.clone(), value / 100_000.0));
                    }
                }
                // What the shape says, and how it is laid in its box.
                "bodyPr" => {
                    let emu = |name: &str, fallback: i64| {
                        get_attr(e, name).and_then(|v| v.parse().ok()).unwrap_or(fallback)
                    };
                    said.insets = (
                        emu("lIns", 91440),
                        emu("tIns", 45720),
                        emu("rIns", 91440),
                        emu("bIns", 45720),
                    );
                    said.anchor = get_attr(e, "anchor");
                    said.wrap = get_attr(e, "wrap").as_deref() != Some("none");
                    said.clip = get_attr(e, "vertOverflow").as_deref() == Some("clip");
                }
                "p" => {
                    paragraph = Some(ShapeParagraph {
                        text: String::new(),
                        align: None,
                        size: 18.0,
                        bold: false,
                        italic: false,
                        face: None,
                        charset: None,
                        color: None,
                        line_pitch: None,
                        line_scale: None,
                    });
                }
                "pPr" if paragraph.is_some() => {
                    if let (Some(held), Some(align)) = (paragraph.as_mut(), get_attr(e, "algn")) {
                        held.align = Some(align);
                    }
                }
                // A break inside a paragraph starts a line without starting a
                // paragraph, which is how `sanko_tool` sets a heading over the
                // text that follows it.
                "br" => {
                    if let Some(held) = paragraph.as_mut() {
                        held.text.push('\n');
                    }
                }
                "lnSpc" => in_line_spacing = true,
                // A paragraph can pin its line pitch outright, or ask for a
                // share of the font's own. Both are stated the same way
                // inside `<a:lnSpc>` as they are inside the space before and
                // after a paragraph, so only what is inside that counts.
                "spcPts" if in_line_spacing => {
                    if let (Some(held), Some(points)) = (
                        paragraph.as_mut(),
                        get_attr(e, "val").and_then(|v| v.parse::<f32>().ok()),
                    ) {
                        if held.line_pitch.is_none() && points > 0.0 {
                            held.line_pitch = Some(points / 100.0);
                        }
                    }
                }
                "spcPct" if in_line_spacing => {
                    if let (Some(held), Some(share)) = (
                        paragraph.as_mut(),
                        get_attr(e, "val").and_then(|v| v.parse::<f32>().ok()),
                    ) {
                        if held.line_scale.is_none() && share > 0.0 {
                            held.line_scale = Some(share / 100_000.0);
                        }
                    }
                }
                // Only the first run of a paragraph dresses it: the corpus
                // dresses every run of a shape's paragraph alike.
                // An empty paragraph states its own size in `endParaRPr`
                // rather than `rPr` — it has no run to dress — and Excel gives
                // it a line of that size. Read as one and the same thing:
                // measured on `_xlsx_shape_block.py`, where an empty paragraph
                // between two blocks is worth exactly one line of the size it
                // states, not of the size a paragraph defaults to.
                "rPr" | "endParaRPr" => {
                    in_run_props = true;
                    if let Some(held) = paragraph.as_mut() {
                        if held.text.is_empty() {
                            if let Some(size) = get_attr(e, "sz").and_then(|v| v.parse::<f32>().ok())
                            {
                                held.size = size / 100.0;
                            }
                            held.bold = is_true(get_attr(e, "b").as_deref());
                            held.italic = is_true(get_attr(e, "i").as_deref());
                        }
                    }
                }
                // The East Asian face wins over the Latin one for the sheets
                // this is measured against.
                "latin" | "ea" if in_run_props => {
                    if let Some(held) = paragraph.as_mut() {
                        if held.text.is_empty() {
                            let face = get_attr(e, "typeface").filter(|face| !face.is_empty());
                            if face.is_some() && (name == "ea" || held.face.is_none()) {
                                held.face = face;
                                // The charset travels with the face it is
                                // written beside: it is what Excel falls back
                                // on when the face is not installed.
                                held.charset = get_attr(e, "charset")
                                    .and_then(|held| held.parse::<i32>().ok());
                            }
                        }
                    }
                }
                "t" => {
                    in_text = true;
                    text.clear();
                }
                // The picture's own part is named by a relationship.
                "blip" if embed.is_none() => {
                    embed = e
                        .attributes()
                        .flatten()
                        .find(|attr| local_name(attr.key.as_ref()) == "embed")
                        .and_then(|attr| String::from_utf8(attr.value.to_vec()).ok());
                }
                _ => {}
            }
            // An element with no children ends where it starts.
            if empty && matches!(name.as_str(), "srgbClr" | "schemeClr" | "sysClr") {
                if let (Some((hex, mods)), Some(part)) = (colour.take(), paints) {
                    let painted = shaded(&hex, &mods);
                    match part {
                        Paints::Fill => shape.fill = Some(painted),
                        Paints::Line => {
                            shape.line = Some(ShapeLine {
                                color: painted,
                                width: line_width,
                                dash: dash.clone(),
                            })
                        }
                        Paints::Text => {
                            if let Some(held) = paragraph.as_mut() {
                                if held.text.is_empty() {
                                    held.color = Some(painted);
                                }
                            }
                        }
                    }
                }
            }
        }
        match &event {
            Ok(Event::Text(e)) => {
                if let Ok(held) = e.unescape() {
                    if in_text {
                        text.push_str(&held);
                    } else {
                        number.push_str(&held);
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = local_name(e.name().as_ref());
                let value = || number.trim().parse::<i64>().unwrap_or(0);
                match name.as_str() {
                    "col" | "colOff" | "row" | "rowOff" => {
                        if let Some(which) = corner {
                            let corner = if which { &mut from } else { to.get_or_insert(blank) };
                            match name.as_str() {
                                "col" => corner.col = value().max(0) as u32,
                                "colOff" => corner.col_off = value(),
                                "row" => corner.row = value().max(0) as u32,
                                _ => corner.row_off = value(),
                            }
                        }
                    }
                    "x" | "y" if corner.is_some() => {
                        let held = number.trim().parse::<f64>().unwrap_or(0.0);
                        let corner = if corner == Some(true) {
                            frame.get_or_insert((0.0, 0.0))
                        } else {
                            frame_to.get_or_insert((0.0, 0.0))
                        };
                        if name == "x" {
                            corner.0 = held;
                        } else {
                            corner.1 = held;
                        }
                    }
                    "from" | "to" => corner = None,
                    "srgbClr" | "schemeClr" | "sysClr" => {
                        if let (Some((hex, mods)), Some(part)) = (colour.take(), paints) {
                            let painted = shaded(&hex, &mods);
                            match part {
                                Paints::Fill => shape.fill = Some(painted),
                                Paints::Line => {
                                    shape.line = Some(ShapeLine {
                                        color: painted,
                                        width: line_width,
                                        dash: dash.clone(),
                                    })
                                }
                                Paints::Text => {
                                    if let Some(held) = paragraph.as_mut() {
                                        if held.text.is_empty() {
                                            held.color = Some(painted);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    "t" => {
                        in_text = false;
                        if let Some(held) = paragraph.as_mut() {
                            held.text.push_str(&text);
                        }
                        text.clear();
                    }
                    "lnSpc" => in_line_spacing = false,
                    "rPr" | "endParaRPr" => in_run_props = false,
                    "p" => {
                        // A paragraph with nothing in it still spends a line,
                        // wherever it sits: `_xlsx_shape_block.py` anchors the
                        // same four paragraphs to the top, the middle and the
                        // foot of a box, and an empty one at the front, at the
                        // back or in the middle moves the ink by exactly one
                        // line every time. Both ends used to be dropped.
                        if let Some(held) = paragraph.take() {
                            said.paragraphs.push(held);
                        }
                    }
                    "txBody" => {
                        if said.paragraphs.iter().any(|held| !held.text.is_empty()) {
                            shape.text = Some(said.clone());
                        }
                    }
                    "solidFill" | "fillRef" | "lnRef" => paints = None,
                    // A rule states its colour first and how it is broken
                    // after, so the dash can only be put on the line once the
                    // whole of `<a:ln>` has been read.
                    "ln" => {
                        in_ln = false;
                        if let Some(line) = shape.line.as_mut() {
                            line.width = line_width;
                            line.dash = dash.take();
                        }
                    }
                    "spPr" => in_sp_pr = false,
                    "style" => in_style = false,
                    "extLst" => in_ext_lst = in_ext_lst.saturating_sub(1),
                    "grpSpPr" => in_group_props = false,
                    "pic" | "sp" | "cxnSp" | "grpSp" | "graphicFrame" => {
                        depth_in_shape = depth_in_shape.saturating_sub(1);
                        // A child of a group is placed by mapping its own box
                        // through the group's transform, and hangs from the
                        // anchor the group hangs from with the difference as
                        // its offset. `002`'s callout is one of these.
                        if let (Some((off, ext, ch_off, ch_ext)), Some(own), Some(size), true) =
                            (group, own_off, own_ext, name != "grpSp")
                        {
                            let across = if ch_ext.0 != 0 {
                                ext.0 as f64 / ch_ext.0 as f64
                            } else {
                                1.0
                            };
                            let down = if ch_ext.1 != 0 {
                                ext.1 as f64 / ch_ext.1 as f64
                            } else {
                                1.0
                            };
                            let left = ((own.0 - ch_off.0) as f64 * across) as i64;
                            let top = ((own.1 - ch_off.1) as f64 * down) as i64;
                            let mut kind = kind.take().unwrap_or(DrawingKind::Other);
                            if let DrawingKind::Shape(held) = &mut kind {
                                *held = shape.clone();
                            }
                            let picture = matches!(kind, DrawingKind::Picture { .. });
                            found.push((
                                Drawing {
                                    from: Anchor {
                                        col: from.col,
                                        col_off: from.col_off + left,
                                        row: from.row,
                                        row_off: from.row_off + top,
                                    },
                                    to: None,
                                    extent: Some((
                                        (size.0 as f64 * across) as i64,
                                        (size.1 as f64 * down) as i64,
                                    )),
                                    kind,
                                    frame: None,
                                },
                                if picture { embed.take() } else { None },
                            ));
                            let _ = (off, ch_off);
                        }
                    }
                    "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor"
                    | "relSizeAnchor" | "absSizeAnchor" => {
                        // A group's children have been pushed one by one; the
                        // group itself draws nothing.
                        if group.is_none() {
                            let mut kind = kind.take().unwrap_or(DrawingKind::Other);
                            if let DrawingKind::Shape(held) = &mut kind {
                                *held = shape.clone();
                            }
                            let named = matches!(
                                kind,
                                DrawingKind::Picture { .. } | DrawingKind::Chart(_)
                            );
                            let held = frame.take().map(|(x, y)| {
                                let (to_x, to_y) = frame_to.take().unwrap_or((x, y));
                                crate::ir::Frame {
                                    x,
                                    y,
                                    w: (to_x - x).max(0.0),
                                    h: (to_y - y).max(0.0),
                                }
                            });
                            found.push((
                                Drawing {
                                    from,
                                    to: to.take(),
                                    extent: extent.take(),
                                    kind,
                                    frame: held,
                                },
                                if named { embed.take() } else { None },
                            ));
                        }
                        embed = None;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    found
}

/// The notes a sheet keeps pinned open, from the pair of parts that hold
/// them: the text in `xl/comments{n}.xml`, keyed by cell, and the box in the
/// VML beside it, which states it in points from the sheet's corner and says
/// whether Excel shows it.
fn parse_comments(comments_xml: &str, vml: &str) -> Vec<crate::ir::Comment> {
    use crate::ir::{Anchor, Comment, ShapeParagraph, ShapeText};

    // The text of each note, by the cell it belongs to.
    let mut said: HashMap<(u32, u32), Vec<ShapeParagraph>> = HashMap::new();
    let mut reader = Reader::from_str(comments_xml);
    let mut buf = Vec::new();
    let mut at: Option<(u32, u32)> = None;
    let mut run = ShapeParagraph {
        text: String::new(),
        align: None,
        size: 9.0,
        bold: false,
        italic: false,
        face: None,
        charset: None,
        color: None,
        line_pitch: None,
        line_scale: None,
    };
    let blank_run = run.clone();
    let (mut in_run_props, mut in_text, mut first) = (false, false, true);
    let mut text = String::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                match local_name(e.name().as_ref()).as_str() {
                    "comment" => {
                        let cell = get_attr(e, "ref").map(|held| parse_cell_ref(&held));
                        at = cell.map(|(column, row)| (row, column));
                        run = blank_run.clone();
                        first = true;
                    }
                    "rPr" => in_run_props = true,
                    "sz" if in_run_props && first => {
                        if let Some(points) = get_attr(e, "val").and_then(|v| v.parse().ok()) {
                            run.size = points;
                        }
                    }
                    "b" if in_run_props && first => run.bold = true,
                    "i" if in_run_props && first => run.italic = true,
                    "rFont" if in_run_props && first => run.face = get_attr(e, "val"),
                    "t" => {
                        in_text = true;
                        text.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) if in_text => {
                if let Ok(held) = e.unescape() {
                    text.push_str(&held);
                }
            }
            Ok(Event::End(ref e)) => match local_name(e.name().as_ref()).as_str() {
                "rPr" => in_run_props = false,
                "t" => {
                    in_text = false;
                    run.text.push_str(&text);
                    first = false;
                }
                "comment" => {
                    if let Some(key) = at.take() {
                        // A note is one run of text with newlines in it; the
                        // paragraphs are what the drawing splits it into.
                        let paragraphs = run
                            .text
                            .split('\n')
                            .map(|line| ShapeParagraph {
                                text: line.trim_end_matches('\r').to_string(),
                                ..run.clone()
                            })
                            .collect();
                        said.insert(key, paragraphs);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // The box of each note, and whether it is shown at all.
    let mut held = Vec::new();
    for shape in vml.split("<v:shape").skip(1) {
        let shape = shape.split("</v:shape>").next().unwrap_or(shape);
        if !shape.contains("<x:Visible/>") {
            continue;
        }
        let tagged = |name: &str| -> Option<String> {
            let open = format!("<x:{name}>");
            let at = shape.find(&open)? + open.len();
            let rest = &shape[at..];
            Some(rest[..rest.find('<')?].trim().to_string())
        };
        let (Some(row), Some(column)) = (
            tagged("Row").and_then(|v| v.parse::<u32>().ok()),
            tagged("Column").and_then(|v| v.parse::<u32>().ok()),
        ) else {
            continue;
        };
        // `<x:Row>` and `<x:Column>` count from zero, the way the cell
        // reference in the comment part does once it is parsed.
        let Some(paragraphs) = said.get(&(row, column)) else {
            continue;
        };
        // The anchor is eight numbers: the cell each corner hangs from and
        // how far into it, in pixels at 96 dpi. Measured against Excel's own
        // picture of `002`, whose note lands at column 64 plus 12 pixels —
        // where the margin the same shape states is 13 pixels short.
        let Some(anchor) = tagged("Anchor") else { continue };
        let numbers: Vec<i64> = anchor
            .split(',')
            .filter_map(|part| part.trim().parse().ok())
            .collect();
        let [left, dx, top, dy, right, dx2, bottom, dy2] = numbers[..] else {
            continue;
        };
        let corner = |column: i64, x: i64, row: i64, y: i64| Anchor {
            col: column.max(0) as u32,
            col_off: x * 9525,
            row: row.max(0) as u32,
            row_off: y * 9525,
        };
        let _ = (right, dx2, bottom, dy2);
        // How big the box is comes from the style, in whatever unit it names:
        // Excel sizes a note to its text and writes the answer there.
        let measure = |name: &str| -> Option<f32> {
            let at = shape.find(name)? + name.len();
            let rest = &shape[at..];
            let end = rest
                .find(|c: char| !(c.is_ascii_digit() || c == '.' || c == '-'))
                .unwrap_or(rest.len());
            let number: f32 = rest[..end].parse().ok()?;
            Some(match rest[end..].chars().take(2).collect::<String>().as_str() {
                held if held.starts_with("in") => number * 72.0,
                held if held.starts_with("mm") => number * 72.0 / 25.4,
                held if held.starts_with("cm") => number * 72.0 / 2.54,
                held if held.starts_with("px") => number * 72.0 / 96.0,
                _ => number,
            })
        };
        let (Some(wide), Some(tall)) = (measure("width:"), measure("height:")) else {
            continue;
        };
        // A note is `#ffffe1` unless the shape says otherwise; the three-digit
        // form is the one Excel writes.
        let fill = shape
            .split("fillcolor=\"#")
            .nth(1)
            .and_then(|rest| rest.split('"').next())
            .map(|held| held.split_whitespace().next().unwrap_or("").to_string())
            .filter(|held| held.len() == 3 || held.len() == 6)
            .map(|held| {
                if held.len() == 3 {
                    held.chars().flat_map(|part| [part, part]).collect()
                } else {
                    held
                }
            })
            .unwrap_or_else(|| "FFFFE1".to_string());
        held.push(Comment {
            from: corner(left, dx, top, dy),
            size: (wide, tall),
            text: ShapeText {
                paragraphs: paragraphs.clone(),
                anchor: Some("t".to_string()),
                // 2.5mm and 2.3mm, which is what the VML states.
                insets: (90000, 82800, 90000, 82800),
                wrap: true,
                clip: false,
            },
            fill: Some(fill.to_uppercase()),
        });
    }
    held
}

/// The theme slot a DrawingML colour names. `tx1` and `bg1` are the same two
/// colours as `dk1` and `lt1` under other names.
pub(crate) fn scheme_colour(name: &str, theme: &Theme) -> Option<String> {
    let slot = match name {
        "dk1" | "tx1" => 0,
        "lt1" | "bg1" => 1,
        "dk2" | "tx2" => 2,
        "lt2" | "bg2" => 3,
        "accent1" => 4,
        "accent2" => 5,
        "accent3" => 6,
        "accent4" => 7,
        "accent5" => 8,
        "accent6" => 9,
        "hlink" => 10,
        "folHlink" => 11,
        _ => return None,
    };
    theme.colours.get(slot).cloned()
}

/// A colour with the modifiers DrawingML hangs off it: `shade` and `tint`
/// move every channel toward black or white, `lumMod` and `lumOff` scale and
/// raise how light it is.
pub(crate) fn shaded(hex: &str, mods: &[(String, f32)]) -> String {
    let Ok(value) = u32::from_str_radix(hex, 16) else {
        return hex.to_string();
    };
    let mut channels = [
        ((value >> 16) & 0xFF) as f32 / 255.0,
        ((value >> 8) & 0xFF) as f32 / 255.0,
        (value & 0xFF) as f32 / 255.0,
    ];
    for (name, amount) in mods {
        match name.as_str() {
            "shade" => channels.iter_mut().for_each(|part| *part *= amount),
            "tint" => channels
                .iter_mut()
                .for_each(|part| *part = *part * amount + (1.0 - amount)),
            // Lightness is HSL's, and only it moves: the hue and the
            // saturation stay where they are. Measured on `002`'s banner,
            // whose fill is accent1 under `lumMod 20% lumOff 80%` — 5B9BD5
            // becomes DEEBF7 exactly, where shifting the bytes by hand lands
            // at CCEDFF.
            "lumMod" | "lumOff" => {
                let high = channels[0].max(channels[1]).max(channels[2]);
                let low = channels[0].min(channels[1]).min(channels[2]);
                let lum = (high + low) / 2.0;
                let span = high - low;
                let wanted = if name == "lumMod" {
                    lum * amount
                } else {
                    (lum + amount).min(1.0)
                };
                if span <= f32::EPSILON {
                    channels = [wanted; 3];
                } else {
                    let sat = span / (1.0 - (2.0 * lum - 1.0).abs()).max(f32::EPSILON);
                    let reach = (1.0 - (2.0 * wanted - 1.0).abs()) * sat.min(1.0);
                    let scale = reach / span;
                    for part in channels.iter_mut() {
                        *part = (wanted + (*part - lum) * scale).clamp(0.0, 1.0);
                    }
                }
            }
            _ => {}
        }
    }
    format!(
        "{:02X}{:02X}{:02X}",
        (channels[0] * 255.0).round() as u8,
        (channels[1] * 255.0).round() as u8,
        (channels[2] * 255.0).round() as u8
    )
}

fn parse_table_xml(xml: &str, theme: &Theme) -> Option<crate::ir::Table> {
    let mut reader = Reader::from_str(xml);
    let mut range = None;
    let mut header_rows = 1;
    let mut style = None;
    let mut banded_rows = false;
    let mut buf = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buf);
        match event {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                match local_name(e.name().as_ref()).as_str() {
                    "table" => {
                        range = get_attr(e, "ref").as_deref().and_then(parse_range_ref);
                        if let Some(count) = get_attr(e, "headerRowCount") {
                            header_rows = count.parse().unwrap_or(1);
                        }
                    }
                    "tableStyleInfo" => {
                        style = get_attr(e, "name");
                        banded_rows = is_true(get_attr(e, "showRowStripes").as_deref());
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    let (start_col, start_row, end_col, end_row) = range?;
    // A built-in style is named for the theme colour it uses: TableStyleMedium2
    // dresses the table in accent1, Medium7 in accent6. Measured on a worksheet
    // holding one table per style, header and band both exact.
    let accent = style.as_deref().and_then(|name| {
        let number: u32 = name
            .strip_prefix("TableStyleMedium")
            .or_else(|| name.strip_prefix("TableStyleLight"))
            .or_else(|| name.strip_prefix("TableStyleDark"))?
            .parse()
            .ok()?;
        // 1 is the greyscale one; 2 onward walk accent1..accent6 and repeat.
        if number < 2 {
            return None;
        }
        theme.colour(4 + ((number - 2) % 6) as usize).map(str::to_string)
    });
    let band = accent.as_deref().map(|colour| tinted(colour, 0.8));
    // A Medium table is ruled along every row in the accent under a lighter
    // tint: `doi-list` wears TableStyleMedium7 over accent6 4EA72E, and Excel
    // draws its rules 8ED973 — which is that colour with its lightness moved
    // to `0.6 L + 0.4`, the tint below. The Light and Dark styles rule
    // themselves differently and are left alone until they are measured.
    let rule = accent
        .as_deref()
        .filter(|_| style.as_deref().is_some_and(|name| name.starts_with("TableStyleMedium")))
        .map(|colour| tinted(colour, 0.4));
    Some(crate::ir::Table {
        start_row,
        start_col,
        end_row,
        end_col,
        style,
        header_rows,
        banded_rows,
        accent,
        band,
        rule,
    })
}

fn parse_worksheet(
    xml: &str,
    sheet_name: &str,
    shared_strings: &[SharedString],
    stylesheet: &StyleSheet,
) -> Result<Sheet, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut rows: Vec<Row> = Vec::new();
    let mut max_col: u32 = 0;

    // Column widths: index is 0-based col number
    let mut col_widths: Vec<f32> = Vec::new();
    let mut hidden_cols: Vec<u32> = Vec::new();
    let mut auto_filter: Option<crate::ir::AutoFilter> = None;
    let mut declared_range: Option<(u32, u32, u32, u32)> = None;
    let mut unsupported: Vec<String> = Vec::new();
    let mut filter_field: Option<u32> = None;
    let mut filter_criteria: Vec<String> = Vec::new();
    let mut filter_either = false;
    // Zero until the sheet states one. A stated default is measured the same
    // way a <col> width is; Excel's own default, for a sheet that states none,
    // is a plain count of characters instead, so the two cannot share a value.
    let mut default_col_width: f32 = 0.0;
    let mut default_row_height: f32 = 15.0;
    let mut default_row_custom = false;
    // The fonts <col> styles put on their columns. A row's height is
    // measured from these, not from the number the sheet or the row states.
    let mut col_fonts: Vec<(u32, u32, String, f32)> = Vec::new();
    let mut merge_cells: Vec<MergeCell> = Vec::new();

    // State tracking
    let mut current_row_index: u32 = 0;
    let mut current_row_height: Option<f32> = None;
    let mut current_row_custom = false;
    let mut current_row_font: Option<(String, f32)> = None;
    let mut current_row_thick_top = false;
    let mut current_row_thick_bottom = false;
    let mut current_row_hidden = false;
    let mut current_cells: Vec<Cell> = Vec::new();
    let mut in_row = false;

    // Cell state
    let mut cell_col: u32 = 0;
    let mut cell_type: Option<String> = None;
    let mut cell_style_index: Option<usize> = None;
    let mut in_cell = false;
    let mut in_value = false;
    let mut value_text = String::new();
    let mut in_formula = false;
    let mut formula_text = String::new();
    // A cell may carry its text itself, in an <is>, instead of pointing into
    // the shared table. Anything that writes a sheet without building that
    // table does it this way. Its text sits in <t> elements, and its phonetic
    // guides in <rPh><t>, which are no more part of the cell's text here than
    // they are in a shared string.
    let mut in_inline = false;
    let mut in_inline_text = false;
    let mut in_inline_phonetic = false;

    // Section tracking
    let mut in_merge_cells = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref());
                note_unsupported(&name, &mut unsupported);
                match name.as_str() {
                    "dimension" => {
                        declared_range = get_attr(&e, "ref")
                            .as_deref()
                            .and_then(parse_range_ref)
                            .map(|(start_col, start_row, end_col, end_row)| {
                                (start_row, start_col, end_row, end_col)
                            });
                    }
                    "autoFilter" => {
                        if let Some(reference) = get_attr(&e, "ref") {
                            if let Some((start_col, start_row, end_col, end_row)) =
                                parse_range_ref(&reference)
                            {
                                auto_filter = Some(crate::ir::AutoFilter {
                                    start_row,
                                    start_col,
                                    end_row,
                                    end_col,
                                    columns: Vec::new(),
                                });
                            }
                        }
                    }
                    "filterColumn" => {
                        filter_field = get_attr(&e, "colId")
                            .and_then(|value| value.parse::<u32>().ok())
                            .map(|col| col + 1);
                        filter_criteria.clear();
                        filter_either = false;
                    }
                    "filters" => {
                        filter_either = true;
                    }
                    "customFilters" => {
                        filter_either = get_attr(&e, "and").as_deref() != Some("1");
                    }
                    "row" => {
                        in_row = true;
                        current_cells.clear();
                        let row_num = get_attr(&e, "r")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(current_row_index + 1);
                        current_row_index = row_num;

                        // A row states its height whenever it is not the
                        // default one. customHeight says whether that number
                        // is pinned; without it the number is only a cache
                        // from the machine that wrote the file, and Excel
                        // works the height out again from the row's content.
                        current_row_height =
                            get_attr(&e, "ht").and_then(|v| v.parse::<f32>().ok());
                        current_row_custom =
                            is_true(get_attr(&e, "customHeight").as_deref());
                        current_row_font = row_style_font(&e, stylesheet);
                        current_row_thick_top =
                            is_true(get_attr(&e, "thickTop").as_deref());
                        current_row_thick_bottom =
                            is_true(get_attr(&e, "thickBot").as_deref());
                        current_row_hidden = is_true(get_attr(&e, "hidden").as_deref());
                    }
                    "c" if in_row => {
                        in_cell = true;
                        value_text.clear();
                        formula_text.clear();
                        in_formula = false;
                        cell_type = get_attr(&e, "t");
                        cell_style_index =
                            get_attr(&e, "s").and_then(|v| v.parse::<usize>().ok());
                        let cell_ref = get_attr(&e, "r").unwrap_or_default();
                        let (col, _) = parse_cell_ref(&cell_ref);
                        cell_col = col;
                        if col + 1 > max_col {
                            max_col = col + 1;
                        }
                    }
                    "f" if in_cell => {
                        in_formula = true;
                        formula_text.clear();
                    }
                    "v" if in_cell => {
                        in_value = true;
                        value_text.clear();
                    }
                    "is" if in_cell => {
                        in_inline = true;
                        in_inline_phonetic = false;
                        value_text.clear();
                    }
                    "rPh" if in_inline => in_inline_phonetic = true,
                    "t" if in_inline && !in_inline_phonetic => in_inline_text = true,
                    "cols" => {
                        // We'll handle col elements inside
                    }
                    "mergeCells" => {
                        in_merge_cells = true;
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "filterColumn" => {
                        if let (Some(field), Some(filter)) =
                            (filter_field.take(), auto_filter.as_mut())
                        {
                            if !filter_criteria.is_empty() {
                                filter.columns.push(crate::ir::AutoFilterColumn {
                                    field,
                                    criteria: std::mem::take(&mut filter_criteria),
                                    either: filter_either,
                                });
                            }
                        }
                    }
                    "row" => {
                        in_row = false;
                        rows.push(Row {
                            index: current_row_index,
                            cells: std::mem::take(&mut current_cells),
                            height: current_row_height,
                            custom_height: current_row_custom,
                            style_font: current_row_font.take(),
                            thick_top: current_row_thick_top,
                            thick_bottom: current_row_thick_bottom,
                            hidden: current_row_hidden,
                        });
                        current_row_height = None;
                        current_row_custom = false;
                        current_row_thick_top = false;
                        current_row_thick_bottom = false;
                        current_row_hidden = false;
                    }
                    "c" => {
                        if in_cell {
                            let cell_value =
                                resolve_cell_value(&value_text, &cell_type, shared_strings);
                            // A cell without an s wears the workbook's default
                            // format, which is cellXfs[0] — not no format at all.
                            let style =
                                resolve_cell_style(cell_style_index.unwrap_or(0), stylesheet);
                            let formula = if formula_text.is_empty() {
                                None
                            } else {
                                Some(formula_text.clone())
                            };
                            current_cells.push(Cell {
                                col: cell_col,
                                value: cell_value,
                                style,
                                formula,
                                runs: runs_of(&value_text, &cell_type, shared_strings),
                            });
                            in_cell = false;
                            in_formula = false;
                            cell_type = None;
                            cell_style_index = None;
                        }
                    }
                    "f" => {
                        in_formula = false;
                    }
                    "v" => {
                        in_value = false;
                    }
                    "is" => {
                        in_inline = false;
                        in_inline_text = false;
                    }
                    "rPh" => in_inline_phonetic = false,
                    "t" => in_inline_text = false,
                    "mergeCells" => {
                        in_merge_cells = false;
                    }
                    _ => {}
                }
            }
            Event::Text(e) => {
                if in_formula {
                    let text = e.unescape()?.to_string();
                    formula_text.push_str(&text);
                } else if in_value || in_inline_text {
                    let text = e.unescape()?.to_string();
                    value_text.push_str(&text);
                }
            }
            Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                note_unsupported(&name, &mut unsupported);
                match name.as_str() {
                    // Handle self-closing <c .../> (cell with no value)
                    "c" if in_row => {
                        let cell_ref = get_attr(&e, "r").unwrap_or_default();
                        let (col, _) = parse_cell_ref(&cell_ref);
                        if col + 1 > max_col {
                            max_col = col + 1;
                        }
                        let si =
                            get_attr(&e, "s").and_then(|v| v.parse::<usize>().ok());
                        let style = resolve_cell_style(si.unwrap_or(0), stylesheet);
                        current_cells.push(Cell {
                            col,
                            value: CellValue::Empty,
                            style,
                            formula: None,
                            runs: Vec::new(),
                        });
                    }

                    // <sheetFormatPr defaultRowHeight="15" defaultColWidth="8.43" ... />
                    "sheetFormatPr" => {
                        if let Some(v) = get_attr(&e, "defaultRowHeight") {
                            if let Ok(h) = v.parse::<f32>() {
                                default_row_height = h;
                            }
                        }
                        default_row_custom =
                            is_true(get_attr(&e, "customHeight").as_deref());
                        if let Some(v) = get_attr(&e, "defaultColWidth") {
                            if let Ok(w) = v.parse::<f32>() {
                                default_col_width = w;
                            }
                        }
                    }

                    // <col min="1" max="3" width="12.5" ... />
                    "dimension" => {
                        declared_range = get_attr(&e, "ref")
                            .as_deref()
                            .and_then(parse_range_ref)
                            .map(|(start_col, start_row, end_col, end_row)| {
                                (start_row, start_col, end_row, end_col)
                            });
                    }
                    "autoFilter" => {
                        if let Some(reference) = get_attr(&e, "ref") {
                            if let Some((start_col, start_row, end_col, end_row)) =
                                parse_range_ref(&reference)
                            {
                                auto_filter = Some(crate::ir::AutoFilter {
                                    start_row,
                                    start_col,
                                    end_row,
                                    end_col,
                                    columns: Vec::new(),
                                });
                            }
                        }
                    }
                    "filterColumn" => {
                        // colId counts from zero within the filtered range;
                        // Field counts from one.
                        filter_field = get_attr(&e, "colId")
                            .and_then(|value| value.parse::<u32>().ok())
                            .map(|col| col + 1);
                        filter_criteria.clear();
                        filter_either = false;
                    }
                    "filters" => {
                        // A list of values is an "is one of these" test.
                        filter_either = true;
                    }
                    "filter" => {
                        if let Some(value) = get_attr(&e, "val") {
                            filter_criteria.push(value);
                        }
                    }
                    "customFilter" => {
                        let operator = get_attr(&e, "operator").unwrap_or_default();
                        let value = get_attr(&e, "val").unwrap_or_default();
                        filter_criteria.push(format!("{}{value}", filter_operator(&operator)));
                    }
                    "customFilters" => {
                        filter_either = get_attr(&e, "and").as_deref() != Some("1");
                    }
                    "col" => {
                        let min_col = get_attr(&e, "min")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(1);
                        let max_col_attr = get_attr(&e, "max")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(min_col);
                        let width = get_attr(&e, "width")
                            .and_then(|v| v.parse::<f32>().ok())
                            .unwrap_or(default_col_width);

                        // Ensure col_widths vec is large enough (0-based)
                        let needed = max_col_attr as usize;
                        if col_widths.len() < needed {
                            col_widths.resize(needed, 0.0);
                        }
                        for c in min_col..=max_col_attr {
                            col_widths[(c - 1) as usize] = width;
                        }
                        if is_true(get_attr(&e, "hidden").as_deref()) {
                            for c in min_col..=max_col_attr {
                                hidden_cols.push(c - 1);
                            }
                        }
                        // A column's style puts a font on every cell in it,
                        // blank ones included, and that is what a row's
                        // height is measured from when nothing in the row
                        // is taller.
                        if let Some(si) =
                            get_attr(&e, "style").and_then(|v| v.parse::<usize>().ok())
                        {
                            let style = resolve_cell_style(si, stylesheet);
                            if let (Some(name), Some(size)) =
                                (style.font_name, style.font_size)
                            {
                                col_fonts.push((min_col - 1, max_col_attr - 1, name, size));
                            }
                        }
                    }

                    // <mergeCell ref="A1:C3"/>
                    "mergeCell" if in_merge_cells => {
                        if let Some(ref_str) = get_attr(&e, "ref") {
                            if let Some((sc, sr, ec, er)) = parse_range_ref(&ref_str) {
                                merge_cells.push(MergeCell {
                                    start_row: sr,
                                    start_col: sc,
                                    end_row: er,
                                    end_col: ec,
                                });
                            }
                        }
                    }

                    // Self-closing <row ... /> (empty row with attributes)
                    "row" => {
                        let row_num = get_attr(&e, "r")
                            .and_then(|v| v.parse::<u32>().ok())
                            .unwrap_or(current_row_index + 1);
                        current_row_index = row_num;
                        rows.push(Row {
                            index: row_num,
                            cells: Vec::new(),
                            height: get_attr(&e, "ht")
                                .and_then(|v| v.parse::<f32>().ok()),
                            custom_height:
                                is_true(get_attr(&e, "customHeight").as_deref()),
                            style_font: row_style_font(&e, stylesheet),
                            thick_top: is_true(get_attr(&e, "thickTop").as_deref()),
                            thick_bottom: is_true(get_attr(&e, "thickBot").as_deref()),
                            hidden: is_true(get_attr(&e, "hidden").as_deref()),
                        });
                    }

                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // Ensure col_widths covers all columns
    let col_count = max_col as usize;
    if col_widths.len() < col_count {
        // A column with no <col> entry of its own keeps a width of zero. The
        // sheet's default is measured in plain characters while a stated width
        // carries the cell's gutter, so a reader has to be able to tell them
        // apart.
        col_widths.resize(col_count, 0.0);
    }

    // The Normal font, which every column no <col> dresses is written in.
    let normal_font = {
        let font_id = stylesheet
            .cell_style_xfs
            .first()
            .map(|xf| xf.font_id)
            .unwrap_or(0);
        stylesheet.fonts.get(font_id).and_then(|font| {
            match (font.name.clone(), font.size) {
                (Some(name), Some(size)) => Some((name, size)),
                _ => None,
            }
        })
    };

    // The first font in the list, which is what an indent is measured in.
    let first_font = stylesheet.fonts.first().and_then(|font| {
        match (font.name.clone(), font.size) {
            (Some(name), Some(size)) => Some((name, size)),
            _ => None,
        }
    });

    Ok(Sheet {
        tables: Vec::new(),
        drawings: Vec::new(),
        comments: Vec::new(),
        name: sheet_name.to_string(),
        rows,
        col_count,
        col_widths,
        default_col_width,
        default_row_height,
        default_row_custom,
        col_fonts,
        normal_font,
        first_font,
        merge_cells,
        auto_filter,
        declared_range,
        hidden_cols: {
            // The order columns appear in says nothing; keep the list tidy.
            let mut hidden_cols = hidden_cols;
            hidden_cols.sort_unstable();
            hidden_cols.dedup();
            hidden_cols
        },
        unsupported_elements: unsupported,
    })
}

/// Resolve a cell's raw value text + type attribute into a CellValue.
/// The dressing of the shared string a cell holds, if it holds one that is not
/// all dressed alike.
fn runs_of(
    value_text: &str,
    cell_type: &Option<String>,
    shared_strings: &[SharedString],
) -> Vec<crate::ir::TextRun> {
    if cell_type.as_deref() != Some("s") {
        return Vec::new();
    }
    value_text
        .parse::<usize>()
        .ok()
        .and_then(|at| shared_strings.get(at))
        .map(|held| held.runs.clone())
        .unwrap_or_default()
}

fn resolve_cell_value(
    value_text: &str,
    cell_type: &Option<String>,
    shared_strings: &[SharedString],
) -> CellValue {
    if value_text.is_empty() && cell_type.is_none() {
        return CellValue::Empty;
    }

    match cell_type.as_deref() {
        Some("s") => {
            // Shared string index
            if let Ok(idx) = value_text.parse::<usize>() {
                if idx < shared_strings.len() {
                    CellValue::String(shared_strings[idx].text.clone())
                } else {
                    CellValue::Error(format!("Invalid SST index: {}", idx))
                }
            } else {
                CellValue::Error(format!("Non-numeric SST index: {}", value_text))
            }
        }
        Some("b") => {
            CellValue::Boolean(value_text == "1" || value_text.eq_ignore_ascii_case("true"))
        }
        Some("e") => CellValue::Error(value_text.to_string()),
        Some("str") | Some("inlineStr") => {
            // Inline string or formula string result
            CellValue::String(value_text.to_string())
        }
        _ => {
            // No type attribute means number
            if value_text.is_empty() {
                CellValue::Empty
            } else if let Ok(n) = value_text.parse::<f64>() {
                CellValue::Number(n)
            } else {
                CellValue::String(value_text.to_string())
            }
        }
    }
}

/// Parse an .xlsx file from raw bytes into a Workbook IR.
///
/// The values Excel cached in the file are **kept**. For display they are the
/// correct answer — recalculating a freshly loaded workbook can only introduce
/// divergence from what the user sees in Excel. Only formula cells that arrive
/// without a cached value are computed.
///
/// Use [`crate::formula::evaluate_workbook_formulas`] after editing a workbook,
/// and [`parse_xlsx_preserving_values`] when even the gap filling is unwanted.
pub fn parse_xlsx(data: &[u8]) -> Result<Workbook, XlsxError> {
    let mut workbook = parse_xlsx_preserving_values(data)?;
    crate::formula::fill_missing_formula_values(&mut workbook);
    Ok(workbook)
}

/// Parse an .xlsx file without recalculating anything.
///
/// Every formula cell keeps the value Excel last computed for it, alongside its
/// formula text. That pair is what makes a workbook usable as a test oracle.
pub fn parse_xlsx_preserving_values(data: &[u8]) -> Result<Workbook, XlsxError> {
    let mut archive = OoxmlArchive::new(data)?;

    // 1. Parse shared strings (optional — some xlsx files have none)
    let shared_strings = match archive.try_read_part("xl/sharedStrings.xml")? {
        Some(xml) => parse_shared_strings(&xml)?,
        None => Vec::new(),
    };

    // 2. Parse styles.xml (optional — some simple xlsx have none)
    // The theme has to be read first: a style may name one of its colours
    // rather than state one of its own.
    let mut theme = match archive.try_read_part("xl/theme/theme1.xml")? {
        Some(xml) => parse_theme_xml(&xml),
        None => Theme::default(),
    };
    let stylesheet = match archive.try_read_part("xl/styles.xml")? {
        Some(xml) => {
            // A hundred of the corpus's workbooks state a palette of their
            // own, and it sits at the foot of the same part as the styles
            // that name it — so it is read in a pass of its own first.
            theme.indexed = indexed_palette(&xml);
            parse_styles_xml(&xml, &theme)?
        }
        None => StyleSheet::default(),
    };

    // 3. Parse workbook.xml to get sheet names and rIds
    let workbook_xml = archive.read_part("xl/workbook.xml")?;
    let sheet_infos = parse_workbook_sheets(&workbook_xml)?;

    // 4. Parse workbook relationships to map rIds to sheet file paths
    let rels_xml = archive.read_part("xl/_rels/workbook.xml.rels")?;
    let rels = parse_relationships(&rels_xml)?;

    // Build rId -> target path map
    let rid_to_path: HashMap<String, String> = rels
        .into_iter()
        .map(|(id, rel)| (id, rel.target))
        .collect();

    // 5. Parse each worksheet
    let mut sheets = Vec::new();
    for info in &sheet_infos {
        let sheet_path = match rid_to_path.get(&info.r_id) {
            Some(target) => {
                // Target is relative to xl/, e.g. "worksheets/sheet1.xml"
                if target.starts_with('/') {
                    // Absolute path within archive (strip leading /)
                    target.trim_start_matches('/').to_string()
                } else {
                    format!("xl/{}", target)
                }
            }
            None => {
                log::warn!(
                    "No relationship found for sheet '{}' (rId={}), skipping",
                    info.name,
                    info.r_id
                );
                continue;
            }
        };

        match archive.try_read_part(&sheet_path)? {
            Some(sheet_xml) => {
                let mut sheet =
                    parse_worksheet(&sheet_xml, &info.name, &shared_strings, &stylesheet)?;
                // A table lives in its own part, named by the sheet's own
                // relationships. Excel dresses the range from there, so no cell
                // inside carries the header fill or the banding.
                let rels_path = sheet_path
                    .rsplit_once('/')
                    .map(|(dir, file)| format!("{dir}/_rels/{file}.rels"))
                    .unwrap_or_default();
                if let Some(rels_xml) = archive.try_read_part(&rels_path)? {
                    let rels = parse_relationships(&rels_xml)?;
                    // A note is two parts: its text, and the VML that says
                    // where its box is and whether the sheet shows it.
                    let beside = |ending: &str| {
                        rels.values()
                            .find(|rel| rel.rel_type.ends_with(ending))
                            .map(|rel| part_beside(&sheet_path, &rel.target))
                    };
                    if let (Some(notes), Some(vml)) = (beside("/comments"), beside("Drawing")) {
                        if let (Some(notes), Some(vml)) = (
                            archive.try_read_part(&notes)?,
                            archive.try_read_part(&vml)?,
                        ) {
                            sheet.comments = parse_comments(&notes, &vml);
                        }
                    }
                    for rel in rels.values() {
                        if rel.rel_type.ends_with("/table") {
                            let part = part_beside(&sheet_path, &rel.target);
                            if let Some(table_xml) = archive.try_read_part(&part)? {
                                if let Some(table) = parse_table_xml(&table_xml, &theme) {
                                    sheet.tables.push(table);
                                }
                            }
                        }
                        // What is drawn over the grid lives in a part of its
                        // own, and each picture inside names its bytes through
                        // that part's own relationships.
                        if rel.rel_type.ends_with("/drawing") {
                            let part = part_beside(&sheet_path, &rel.target);
                            let Some(drawing_xml) = archive.try_read_part(&part)? else {
                                continue;
                            };
                            let inside = part
                                .rsplit_once('/')
                                .map(|(dir, file)| format!("{dir}/_rels/{file}.rels"))
                                .unwrap_or_default();
                            let media = match archive.try_read_part(&inside)? {
                                Some(xml) => parse_relationships(&xml)?,
                                None => Default::default(),
                            };
                            for (mut drawn, named) in parse_drawing_xml(&drawing_xml, &theme) {
                                let beside = named
                                    .and_then(|named| media.get(&named))
                                    .map(|rel| part_beside(&part, &rel.target));
                                match (&mut drawn.kind, beside) {
                                    (
                                        crate::ir::DrawingKind::Picture { bytes },
                                        Some(image),
                                    ) => {
                                        if let Some(found) = archive.try_read_bytes(&image)? {
                                            *bytes = found;
                                        }
                                    }
                                    (crate::ir::DrawingKind::Chart(chart), Some(graph)) => {
                                        match archive
                                            .try_read_part(&graph)?
                                            .as_deref()
                                            .and_then(|xml| {
                                                crate::chart::parse_chart_xml(xml, &theme)
                                            }) {
                                            Some(read) => *chart = read,
                                            // A chart nothing can draw is left
                                            // out rather than half-drawn.
                                            None => continue,
                                        }
                                        // The boxes a chart is annotated with
                                        // hang from a part of its own.
                                        let beside = graph
                                            .rsplit_once('/')
                                            .map(|(dir, file)| {
                                                format!("{dir}/_rels/{file}.rels")
                                            })
                                            .unwrap_or_default();
                                        if let Some(xml) = archive.try_read_part(&beside)? {
                                            for rel in parse_relationships(&xml)?.values() {
                                                if !rel.rel_type.ends_with("/chartUserShapes") {
                                                    continue;
                                                }
                                                let held = part_beside(&graph, &rel.target);
                                                if let Some(body) =
                                                    archive.try_read_part(&held)?
                                                {
                                                    chart.shapes.extend(
                                                        parse_drawing_xml(&body, &theme)
                                                            .into_iter()
                                                            .map(|(drawn, _)| drawn)
                                                            .filter(|drawn| {
                                                                drawn.frame.is_some()
                                                            }),
                                                    );
                                                }
                                            }
                                        }
                                    }
                                    _ => {}
                                }
                                sheet.drawings.push(drawn);
                            }
                        }
                    }
                }
                sheets.push(sheet);
            }
            None => {
                log::warn!("Sheet file '{}' not found in archive, skipping", sheet_path);
            }
        }
    }

    Ok(Workbook {
        sheets,
        default_style: resolve_cell_style(0, &stylesheet),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_ref_simple() {
        assert_eq!(parse_cell_ref("A1"), (0, 0));
        assert_eq!(parse_cell_ref("B2"), (1, 1));
        assert_eq!(parse_cell_ref("Z1"), (25, 0));
    }

    /// Furigana lives in `<rPh>` and is not part of the cell's text. Excel shows
    /// "区分"; appending the reading would give "区分クブン" and silently corrupt
    /// most Japanese workbooks, which carry phonetic guides on names and
    /// addresses as a matter of course.
    #[test]
    fn shared_strings_exclude_phonetic_guides() {
        let xml = r#"<?xml version="1.0"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3">
  <si><t>区分</t><rPh sb="0" eb="2"><t>クブン</t></rPh><phoneticPr fontId="1"/></si>
  <si><r><t>山田</t></r><r><t>太郎</t></r><rPh sb="0" eb="2"><t>ヤマダ</t></rPh><rPh sb="2" eb="4"><t>タロウ</t></rPh></si>
  <si><t>plain</t></si>
</sst>"#;
        let strings = parse_shared_strings(xml).expect("should parse");
        let held: Vec<&str> = strings.iter().map(|one| one.text.as_str()).collect();
        assert_eq!(held, vec!["区分", "山田太郎", "plain"]);
    }

    /// A cell can hold its own text instead of pointing into the shared table.
    /// Whole sheets are written that way by anything that streams a workbook
    /// out without building the table first, and Excel shows them like any
    /// other text — so a reader that only looks in <v> draws an empty sheet.
    #[test]
    fn inline_strings_are_read_from_the_cell() {
        let xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Ag亜</t></is></c>
      <c r="B1" t="inlineStr"><is><r><t>half </t></r><r><t>and half</t></r></is></c>
      <c r="C1" t="inlineStr"><is><t>区分</t><rPh sb="0" eb="2"><t>クブン</t></rPh></is></c>
      <c r="D1"><v>42</v></c>
    </row>
  </sheetData>
</worksheet>"#;
        let sheet = parse_worksheet(xml, "probe", &[], &StyleSheet::default())
            .expect("should parse");
        let cells = &sheet.rows[0].cells;
        let text_of = |cell: &Cell| match &cell.value {
            CellValue::String(held) => held.clone(),
            other => panic!("wanted text, found {other:?}"),
        };
        assert_eq!(text_of(&cells[0]), "Ag亜");
        assert_eq!(text_of(&cells[1]), "half and half");
        // A phonetic guide is no more part of the text here than in a shared
        // string: Excel shows 区分, not 区分クブン.
        assert_eq!(text_of(&cells[2]), "区分");
        assert!(matches!(cells[3].value, CellValue::Number(n) if n == 42.0));
    }

    /// Shared strings that are all dressed alike, which is what most tests want.
    fn shared(texts: &[&str]) -> Vec<SharedString> {
        texts
            .iter()
            .map(|text| SharedString {
                text: (*text).to_string(),
                runs: Vec::new(),
            })
            .collect()
    }

    #[test]
    fn test_parse_cell_ref_multi_letter() {
        assert_eq!(parse_cell_ref("AA1"), (26, 0));
        assert_eq!(parse_cell_ref("AB1"), (27, 0));
        assert_eq!(parse_cell_ref("AZ3"), (51, 2));
    }

    #[test]
    fn test_parse_cell_ref_large_row() {
        assert_eq!(parse_cell_ref("A100"), (0, 99));
        assert_eq!(parse_cell_ref("C65536"), (2, 65535));
    }

    #[test]
    fn test_resolve_cell_value_number() {
        let sst: Vec<SharedString> = vec![];
        assert!(matches!(
            resolve_cell_value("42", &None, &sst),
            CellValue::Number(n) if (n - 42.0).abs() < f64::EPSILON
        ));
    }

    #[test]
    fn test_resolve_cell_value_shared_string() {
        let sst = shared(&["Hello", "World"]);
        let t = Some("s".to_string());
        assert!(matches!(
            resolve_cell_value("0", &t, &sst),
            CellValue::String(ref s) if s == "Hello"
        ));
        assert!(matches!(
            resolve_cell_value("1", &t, &sst),
            CellValue::String(ref s) if s == "World"
        ));
    }

    #[test]
    fn test_resolve_cell_value_boolean() {
        let sst: Vec<SharedString> = vec![];
        let t = Some("b".to_string());
        assert!(matches!(
            resolve_cell_value("1", &t, &sst),
            CellValue::Boolean(true)
        ));
        assert!(matches!(
            resolve_cell_value("0", &t, &sst),
            CellValue::Boolean(false)
        ));
    }

    #[test]
    fn test_resolve_cell_value_error() {
        let sst: Vec<SharedString> = vec![];
        let t = Some("e".to_string());
        assert!(matches!(
            resolve_cell_value("#REF!", &t, &sst),
            CellValue::Error(ref s) if s == "#REF!"
        ));
    }

    #[test]
    fn test_resolve_cell_value_empty() {
        let sst: Vec<SharedString> = vec![];
        assert!(matches!(
            resolve_cell_value("", &None, &sst),
            CellValue::Empty
        ));
    }

    #[test]
    fn test_cell_value_display() {
        assert_eq!(CellValue::Empty.display(), "");
        assert_eq!(CellValue::String("hello".into()).display(), "hello");
        assert_eq!(CellValue::Number(42.0).display(), "42");
        assert_eq!(CellValue::Number(2.75).display(), "2.75");
        assert_eq!(CellValue::Boolean(true).display(), "TRUE");
        assert_eq!(CellValue::Boolean(false).display(), "FALSE");
        assert_eq!(CellValue::Error("#N/A".into()).display(), "#N/A");
    }

    #[test]
    fn test_parse_shared_strings() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
  <si><r><t>Rich</t></r><r><t> Text</t></r></si>
</sst>"#;
        let result = parse_shared_strings(xml).unwrap();
        assert_eq!(result.len(), 3);
        assert_eq!(result[0].text, "Hello");
        assert_eq!(result[1].text, "World");
        assert_eq!(result[2].text, "Rich Text");
    }

    #[test]
    fn test_parse_workbook_sheets() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Data" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"#;
        let result = parse_workbook_sheets(xml).unwrap();
        assert_eq!(result.len(), 2);
        assert_eq!(result[0].name, "Sheet1");
        assert_eq!(result[0].r_id, "rId1");
        assert_eq!(result[1].name, "Data");
        assert_eq!(result[1].r_id, "rId2");
    }

    #[test]
    fn test_parse_range_ref() {
        assert_eq!(parse_range_ref("A1:C3"), Some((0, 1, 2, 3)));
        assert_eq!(parse_range_ref("B2:D5"), Some((1, 2, 3, 5)));
        assert_eq!(parse_range_ref("A1"), None);
    }

    /// A cell format that names a font and says nothing about applying it is
    /// using that font. Only an explicit `applyFont="0"` sends the reader to
    /// the named style the format is built on — which is how a hyperlink keeps
    /// its blue underline. Files written by anything other than Excel
    /// routinely leave every flag off, so reading absent as "do not apply"
    /// draws whole workbooks in the wrong face.
    #[test]
    fn a_font_is_applied_unless_the_format_denies_it() {
        let xml = r##"<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><sz val="10"/><name val="ＭＳ ゴシック"/></font>
    <font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/></font>
  </fonts>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1" applyFont="0"/>
  </cellXfs>
</styleSheet>"##;
        let sheet = parse_styles_xml(xml, &Theme::default()).unwrap();
        let named = resolve_cell_style(1, &sheet);
        assert_eq!(named.font_name.as_deref(), Some("ＭＳ ゴシック"));
        assert_eq!(named.font_size, Some(10.0));
        let denied = resolve_cell_style(2, &sheet);
        assert!(denied.underline, "should wear the style's own font");
    }

    /// A generated header row is written as `<font><b/></font>` — bold, and
    /// nothing else. The face and size it leaves out are the workbook's, not
    /// the reader's idea of a default: read otherwise, every such row is drawn
    /// in the wrong face and stands the wrong height.
    #[test]
    fn a_font_that_names_no_face_wears_the_workbook_s() {
        let xml = r##"<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><name val="ＭＳ Ｐゴシック"/><sz val="9"/></font>
    <font><b val="1"/><color rgb="00FFFFFF"/></font>
  </fonts>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>"##;
        let sheet = parse_styles_xml(xml, &Theme::default()).unwrap();
        let header = resolve_cell_style(1, &sheet);
        assert!(header.bold);
        assert_eq!(header.font_name.as_deref(), Some("ＭＳ Ｐゴシック"));
        assert_eq!(header.font_size, Some(9.0));
    }

    /// A drawing part states where a shape hangs, what it is painted with,
    /// and — for a picture — which part holds its bytes.
    #[test]
    fn a_drawing_part_gives_up_its_anchors_and_its_paint() {
        let xml = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>5</xdr:col><xdr:colOff>12699</xdr:colOff>
      <xdr:row>8</xdr:row><xdr:rowOff>85725</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>7</xdr:col><xdr:colOff>590550</xdr:colOff>
      <xdr:row>9</xdr:row><xdr:rowOff>123825</xdr:rowOff></xdr:to>
    <xdr:sp macro="" textlink="">
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
        <a:ln w="28575"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="952500" cy="476250"/>
    <xdr:pic>
      <xdr:blipFill><a:blip r:embed="rId7"/></xdr:blipFill>
      <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>"##;
        let found = parse_drawing_xml(xml, &Theme::default());
        assert_eq!(found.len(), 2);

        let (rectangle, embed) = &found[0];
        assert_eq!(embed.as_deref(), None);
        assert_eq!(rectangle.from.col, 5);
        assert_eq!(rectangle.from.col_off, 12699);
        assert_eq!(rectangle.to.map(|to| (to.col, to.row)), Some((7, 9)));
        let crate::ir::DrawingKind::Shape(shape) = &rectangle.kind else {
            panic!("the first anchor holds a shape");
        };
        assert_eq!(shape.geometry, "rect");
        assert_eq!(shape.fill, None, "a:noFill leaves the shape unpainted");
        let line = shape.line.as_ref().expect("the shape is ruled");
        assert_eq!(line.color, "FF0000");
        assert_eq!(line.width, 28575, "three pixels at 96 dpi");
        assert_eq!(line.dash, None);

        let (picture, embed) = &found[1];
        assert_eq!(embed.as_deref(), Some("rId7"));
        assert_eq!(picture.extent, Some((952500, 476250)));
        assert!(matches!(picture.kind, crate::ir::DrawingKind::Picture { .. }));
    }

    /// A shape's own text, with the dressing of the run that starts each
    /// paragraph and the insets its body states.
    #[test]
    fn a_shape_gives_up_what_it_says() {
        let xml = r##"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
      <xdr:txBody>
        <a:bodyPr lIns="288000" tIns="72000" rIns="288000" bIns="72000" anchor="ctr"/>
        <a:p>
          <a:pPr algn="ctr"><a:lnSpc><a:spcPts val="3000"/></a:lnSpc></a:pPr>
          <a:r>
            <a:rPr sz="2000" b="1">
              <a:solidFill><a:srgbClr val="203864"/></a:solidFill>
              <a:latin typeface="Calibri"/><a:ea typeface="メイリオ"/>
            </a:rPr>
            <a:t>この計算書は、</a:t>
          </a:r>
          <a:r><a:rPr sz="2000" b="1"/><a:t>所得を計算するものです。</a:t></a:r>
        </a:p>
        <a:p><a:endParaRPr sz="2000"/></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"##;
        let found = parse_drawing_xml(xml, &Theme::default());
        let crate::ir::DrawingKind::Shape(shape) = &found[0].0.kind else {
            panic!("the anchor holds a shape");
        };
        let said = shape.text.as_ref().expect("the shape says something");
        assert_eq!(said.anchor.as_deref(), Some("ctr"));
        assert_eq!(said.insets, (288000, 72000, 288000, 72000));
        // The trailing empty paragraph is a line, not a marker: Excel moves
        // the ink of a middle-anchored block half a line up and a
        // bottom-anchored one a whole line up when it is there
        // (`_xlsx_shape_block.py`). It states its own size in `endParaRPr`.
        assert_eq!(said.paragraphs.len(), 2);
        assert_eq!(said.paragraphs[1].text, "");
        assert_eq!(said.paragraphs[1].size, 20.0);
        let first = &said.paragraphs[0];
        assert_eq!(first.text, "この計算書は、所得を計算するものです。");
        assert_eq!(first.align.as_deref(), Some("ctr"));
        assert_eq!(first.size, 20.0);
        assert!(first.bold);
        assert_eq!(first.face.as_deref(), Some("メイリオ"), "the East Asian face wins");
        assert_eq!(first.color.as_deref(), Some("203864"));
        assert_eq!(first.line_pitch, Some(30.0));
    }

    /// 162 of the corpus's workbooks name a colour by number rather than by
    /// value, and a hundred of them state a palette of their own.
    #[test]
    fn a_colour_named_by_number_comes_from_the_palette() {
        let xml = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><u/><sz val="11"/><color indexed="12"/><name val="Calibri"/></font>
    <font><sz val="11"/><color indexed="64"/><name val="Calibri"/></font>
  </fonts>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
</styleSheet>"##;
        let sheet = parse_styles_xml(xml, &Theme::default()).unwrap();
        // 12 is the blue Excel has kept since it had only 56 colours to keep.
        assert_eq!(resolve_cell_style(1, &sheet).font_color.as_deref(), Some("0000FF"));
        // 64 is the system's own foreground, which is not in the palette at
        // all: the caller decides, as it did before any of this was read.
        assert_eq!(resolve_cell_style(2, &sheet).font_color, None);

        // A workbook that states its own palette is taken at its word.
        let own = xml.replace(
            "</styleSheet>",
            "<colors><indexedColors>\
             <rgbColor rgb=\"00000000\"/><rgbColor rgb=\"00FFFFFF\"/>\
             <rgbColor rgb=\"00FF0000\"/><rgbColor rgb=\"0000FF00\"/>\
             <rgbColor rgb=\"000000FF\"/><rgbColor rgb=\"00FFFF00\"/>\
             <rgbColor rgb=\"00FF00FF\"/><rgbColor rgb=\"0000FFFF\"/>\
             <rgbColor rgb=\"00000000\"/><rgbColor rgb=\"00FFFFFF\"/>\
             <rgbColor rgb=\"00FF0000\"/><rgbColor rgb=\"0000FF00\"/>\
             <rgbColor rgb=\"00123456\"/></indexedColors></colors></styleSheet>",
        );
        let mut theme = Theme::default();
        theme.indexed = indexed_palette(&own);
        assert_eq!(theme.indexed.len(), 13);
        let sheet = parse_styles_xml(&own, &theme).unwrap();
        assert_eq!(resolve_cell_style(1, &sheet).font_color.as_deref(), Some("123456"));
    }

    /// `<a:br/>` starts a line without starting a paragraph. Without it
    /// `sanko_tool`'s heading ran into the sentence under it — one line where
    /// Excel draws two.
    #[test]
    fn a_break_inside_a_paragraph_starts_a_line() {
        let xml = r##"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr sz="1200" u="sng"/><a:t>イメージ</a:t></a:r>
          <a:br><a:rPr sz="1200"/></a:br>
          <a:br><a:rPr sz="1200"/></a:br>
          <a:r><a:rPr sz="1200"/><a:t>１．確認したい品目</a:t></a:r>
        </a:p>
        <a:p><a:r><a:rPr sz="1200"/><a:t>２．リンク係数</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"##;
        let found = parse_drawing_xml(xml, &Theme::default());
        let crate::ir::DrawingKind::Shape(shape) = &found[0].0.kind else {
            panic!("the anchor holds a shape");
        };
        let said = shape.text.as_ref().expect("the shape says something");
        assert_eq!(said.paragraphs.len(), 2);
        assert_eq!(said.paragraphs[0].text, "イメージ\n\n１．確認したい品目");
        assert_eq!(said.paragraphs[1].text, "２．リンク係数");
    }

    /// A rule states its colour before it says how it is broken, so the dash
    /// has to be put on the line at the end of `<a:ln>` rather than when the
    /// colour is read — `glossary_05`'s dashed frame came out solid until it
    /// was.
    #[test]
    fn a_dash_stated_after_the_colour_still_reaches_the_line() {
        let xml = r##"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
        <a:ln w="19050">
          <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
          <a:prstDash val="dash"/>
        </a:ln>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"##;
        let found = parse_drawing_xml(xml, &Theme::default());
        let crate::ir::DrawingKind::Shape(shape) = &found[0].0.kind else {
            panic!("the anchor holds a shape");
        };
        let line = shape.line.as_ref().expect("the shape is ruled");
        assert_eq!(line.dash.as_deref(), Some("dash"));
        assert_eq!(line.width, 19050);
    }

    /// A shape can name a theme colour and shade it rather than state one.
    #[test]
    fn a_scheme_colour_is_resolved_and_shaded() {
        let theme = Theme {
            colours: vec![
                "000000".into(), "FFFFFF".into(), "44546A".into(), "E7E6E6".into(),
                "4472C4".into(),
            ],
            ..Theme::default()
        };
        let xml = r##"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:spPr>
        <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
        <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
        <a:ln><a:solidFill><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:ln>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"##;
        let found = parse_drawing_xml(xml, &theme);
        let crate::ir::DrawingKind::Shape(shape) = &found[0].0.kind else {
            panic!("the anchor holds a shape");
        };
        assert_eq!(shape.fill.as_deref(), Some("4472C4"));
        assert_eq!(
            shape.line.as_ref().map(|line| line.color.as_str()),
            Some("223962"),
            "a 50% shade halves every channel"
        );
    }

    /// How light a colour is moves through HSL, and Excel's own picture is
    /// what says so: `002`'s banner is accent1 — 5B9BD5 in that workbook's
    /// theme — under `lumMod 20% lumOff 80%`, and Excel paints it DEEBF7.
    #[test]
    fn lightness_is_moved_the_way_excel_moves_it() {
        assert_eq!(
            shaded("5B9BD5", &[("lumMod".into(), 0.2), ("lumOff".into(), 0.8)]),
            "DEEBF7"
        );
        // tx1 at 65% with 35% added is the grey Excel rules a table with.
        assert_eq!(
            shaded("000000", &[("lumMod".into(), 0.65), ("lumOff".into(), 0.35)]),
            "595959"
        );
    }

    /// An indent is a level, not a measurement, and it survives whichever way
    /// the alignment element is written: quick-xml hands an empty element to a
    /// different arm than one with children, and only one of them used to be
    /// read.
    #[test]
    fn an_indent_is_read_either_way_the_alignment_is_written() {
        let xml = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
      <alignment horizontal="left" indent="2"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
      <alignment horizontal="distributed" indent="3" vertical="center"/>
    </xf>
  </cellXfs>
</styleSheet>"##;
        let sheet = parse_styles_xml(xml, &Theme::default()).unwrap();
        assert_eq!(resolve_cell_style(0, &sheet).indent, 0);
        assert_eq!(resolve_cell_style(1, &sheet).indent, 2);
        assert_eq!(resolve_cell_style(2, &sheet).indent, 3);
    }

    /// The font an indent is measured in is the first one in the list, not
    /// the one the Normal style points at. The two are the same in a workbook
    /// Excel writes and different in one openpyxl does, and Excel's indent
    /// follows the first entry in both (`_xlsx_indent_bisect.py`).
    #[test]
    fn the_first_font_is_kept_apart_from_the_normal_style_font() {
        let styles = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><sz val="9"/><name val="Meiryo UI"/></font>
  </fonts>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>"##;
        let stylesheet = parse_styles_xml(styles, &Theme::default()).unwrap();
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>"#;
        let sheet = parse_worksheet(xml, "probe", &[], &stylesheet).expect("should parse");
        assert_eq!(sheet.first_font, Some(("Calibri".to_string(), 11.0)));
        assert_eq!(sheet.normal_font, Some(("Meiryo UI".to_string(), 9.0)));
    }

    /// The pieces a shared string is dressed in add up to the string: a
    /// reader that draws them one after another must not be handed more text
    /// than the cell holds. `f1b851d0a096_001290291`'s title is five pieces
    /// of 23, 2, 3, 19 and 15 characters — 62 in all, which is what its cell
    /// shows.
    #[test]
    fn the_pieces_of_a_shared_string_add_up_to_it() {
        let xml = concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
            r#"<si><r><rPr><sz val="20"/><rFont val="ＭＳ 明朝"/></rPr>"#,
            "<t xml:space=\"preserve\">\n\nはじめ</t></r>",
            r#"<r><rPr><sz val="14"/><rFont val="ＭＳ 明朝"/></rPr>"#,
            "<t xml:space=\"preserve\">\n\n</t></r>",
            r#"<r><rPr><sz val="18"/><rFont val="ＭＳ 明朝"/></rPr><t>集計表</t></r>"#,
            r#"<rPh sb="25" eb="28"><t>シュウケイヒョウ</t></rPh>"#,
            r#"<phoneticPr fontId="9"/></si></sst>"#,
        );
        let held = parse_shared_strings(xml).expect("should parse");
        let string = &held[0];
        let pieces: usize = string.runs.iter().map(|run| run.text.chars().count()).sum();
        assert_eq!(string.text.chars().count(), pieces, "text {:?} runs {:?}",
                   string.text, string.runs.iter().map(|r| &r.text).collect::<Vec<_>>());
        assert_eq!(string.runs.len(), 3);
        // The phonetic guide is in neither.
        assert!(!string.text.contains('シ'));
    }

    #[test]
    fn a_diagonal_is_read_whether_or_not_it_states_a_colour() {
        // A `<diagonal/>` that names a colour has a child and so arrives as a
        // different event than one that closes itself. Both spellings rule
        // the cell corner to corner, so both are pinned here.
        let xml = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border diagonalUp="1">
      <left/><right/><top/><bottom/>
      <diagonal style="thin"><color indexed="64"/></diagonal>
    </border>
    <border diagonalDown="1">
      <left/><right/><top/><bottom/><diagonal style="medium"/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
  </cellXfs>
</styleSheet>"##;
        let sheet = parse_styles_xml(xml, &Theme::default()).unwrap();

        let plain = resolve_cell_style(0, &sheet);
        assert!(plain.border_diagonal.is_none());

        let up = resolve_cell_style(1, &sheet);
        assert_eq!(up.border_diagonal.as_ref().map(|l| l.style.as_str()), Some("thin"));
        assert!(up.diagonal_up);
        assert!(!up.diagonal_down);

        let down = resolve_cell_style(2, &sheet);
        assert_eq!(down.border_diagonal.as_ref().map(|l| l.style.as_str()), Some("medium"));
        assert!(down.diagonal_down);
        assert!(!down.diagonal_up);
    }

    #[test]
    fn test_parse_styles_xml() {
        let xml = r##"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00_ "/>
  </numFmts>
  <fonts count="2">
    <font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FFFF0000"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/></border>
    <border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border>
  </borders>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="1" fillId="1" borderId="1"><alignment horizontal="center"/></xf>
  </cellXfs>
</styleSheet>"##;
        let ss = parse_styles_xml(xml, &Theme::default()).unwrap();
        assert_eq!(ss.num_fmts.len(), 1);
        assert_eq!(ss.num_fmts.get(&164).unwrap(), "#,##0.00_ ");
        assert_eq!(ss.fonts.len(), 2);
        assert!(ss.fonts[1].bold);
        assert_eq!(ss.fonts[1].size, Some(14.0));
        assert_eq!(ss.fonts[1].color.as_deref(), Some("FF0000"));
        assert_eq!(ss.fills.len(), 2);
        assert_eq!(ss.fills[1].bg_color.as_deref(), Some("FFFF00"));
        assert_eq!(ss.borders.len(), 2);
        assert!(ss.borders[0].left.is_none());
        assert!(ss.borders[1].left.is_some());
        assert!(ss.borders[1].right.is_some());
        assert!(ss.borders[1].top.is_some());
        assert!(ss.borders[1].bottom.is_some());
        assert_eq!(ss.cell_xfs.len(), 2);
        assert_eq!(ss.cell_xfs[1].num_fmt_id, 164);
        assert_eq!(
            ss.cell_xfs[1].horizontal_align.as_deref(),
            Some("center")
        );

        // Test resolve_cell_style
        let style = resolve_cell_style(1, &ss);
        assert!(style.bold);
        assert_eq!(style.font_color.as_deref(), Some("FF0000"));
        assert_eq!(style.bg_color.as_deref(), Some("FFFF00"));
        assert_eq!(style.number_format.as_deref(), Some("#,##0.00_ "));
        assert_eq!(style.horizontal_align.as_deref(), Some("center"));
        assert!(style.border_top.is_some());
        assert!(style.border_bottom.is_some());
        assert!(style.border_left.is_some());
        assert!(style.border_right.is_some());
    }

    #[test]
    fn test_builtin_number_formats() {
        assert_eq!(builtin_number_format(0), Some("General"));
        assert_eq!(builtin_number_format(3), Some("#,##0"));
        assert_eq!(builtin_number_format(14), Some("mm-dd-yy"));
        assert_eq!(builtin_number_format(99), None);
    }
}

#[cfg(test)]
mod theme_tints {
    use super::tinted;

    fn close_to(got: &str, want: &str) {
        let byte = |hex: &str, at: usize| {
            u8::from_str_radix(&hex[at..at + 2], 16).expect("two hex digits") as i32
        };
        for at in [0, 2, 4] {
            let apart = (byte(got, at) - byte(want, at)).abs();
            assert!(apart <= 1, "{got} is not within a shade of {want}");
        }
    }

    /// Both colours were read off a worksheet Excel drew: a table's banded row
    /// is its header colour under a tint of 0.8. Rounding puts us within one
    /// step of Excel on each channel.
    #[test]
    fn a_tint_lightens_without_changing_the_hue() {
        close_to(&tinted("156082", 0.8), "C0E6F5");
        close_to(&tinted("4EA72E", 0.8), "DAF2D0");
    }

    #[test]
    fn a_negative_tint_darkens() {
        close_to(&tinted("FFFFFF", -0.5), "808080");
        assert_eq!(tinted("156082", 0.0), "156082");
    }
}
