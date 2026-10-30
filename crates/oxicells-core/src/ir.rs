// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use serde::{Deserialize, Serialize};

pub type DefinedNameAttributes = std::collections::BTreeMap<String, std::collections::BTreeMap<String, String>>;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
    /// The DrawingML theme palette in OOXML order: dark1, light1, dark2,
    /// light2, accents 1-6, hyperlink and followed hyperlink.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub theme_colors: Vec<String>,
    /// The workbook's indexed colour palette in OOXML order (`indexedColors`).
    /// VBA ColorIndex reads and writes depend on this palette; an empty list
    /// means the Excel default palette.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub indexed_colors: Vec<String>,
    #[serde(default)]
    pub calculation: CalculationSettings,
    /// In-memory calculation work carried between editor/API invocations.
    /// It is not written as a private part in an OOXML package.
    #[serde(default, skip_serializing_if = "RecalculationState::is_empty")]
    pub recalculation: RecalculationState,
    /// The format a cell wears when it names none of its own: the first entry
    /// of the workbook's cell formats. A column width is stated in characters
    /// of this font, so drawing a sheet needs it.
    #[serde(default)]
    pub default_style: CellStyle,
    /// The names the workbook gives to things, as `(name, what it refers to)`.
    /// A formula saying `SUM(sales)` means one of these, and without them it
    /// means nothing at all.
    ///
    /// Sheet-local names include their quoted sheet prefix (`'My Sheet'!rate`),
    /// so different sheets can define the same leaf without losing its scope.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub defined_names: Vec<(String, String)>,
    /// Attributes of each defined name, keyed by its complete scoped name.
    /// `name` and `localSheetId` describe the imported spelling; saving derives
    /// the current name and scope from `defined_names` and the sheet order.
    #[serde(default, skip_serializing_if = "std::collections::BTreeMap::is_empty")]
    pub defined_name_attributes: DefinedNameAttributes,
    /// In-session visibility of compatibility identifiers exposed by formula
    /// syntax. These are not definedName elements and are not saved to OOXML.
    #[serde(default, skip_serializing_if = "std::collections::BTreeMap::is_empty")]
    pub compatibility_names: std::collections::BTreeMap<String, bool>,
    /// The workbooks this one only links to, in the order `[1]`, `[2]` 窶ｦ name
    /// them, each carrying the values it was last seen holding.
    ///
    /// Excel keeps that copy in the file precisely so a formula reading
    /// another workbook still has an answer when the other workbook is not
    /// open 窶・which, in a browser, it never is. A link whose copy is empty
    /// (Excel writes `refreshError="1"` for one it could not refresh) has no
    /// answer to give, and the cell keeps whatever it was saved showing.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub external_books: Vec<ExternalBook>,
}

/// A sheet-qualified defined name, using the same quoting as a formula.
pub fn scoped_defined_name(sheet: &str, name: &str) -> String {
    let plain = !sheet.is_empty()
        && !sheet.chars().next().is_some_and(|first| first.is_ascii_digit())
        && sheet.chars().all(|ch| ch.is_alphanumeric() || ch == '_' || ch == '.');
    let scope = if plain { sheet.to_string() } else { format!("'{}'", sheet.replace('\'', "''")) };
    format!("{scope}!{name}")
}

/// The owner and leaf of a sheet-local name; apostrophes belong to the owner.
pub fn defined_name_scope(name: &str) -> Option<(String, &str)> {
    let (scope, leaf) = name.rsplit_once('!')?;
    let scope = scope.strip_prefix('\'').and_then(|s| s.strip_suffix('\''))
        .map(|s| s.replace("''", "'")).unwrap_or_else(|| scope.to_string());
    Some((scope, leaf))
}

/// Formula coordinates use zero-based columns and rows, as the calculator does.
pub type FormulaLocation = (String, (u32, u32));

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct RecalculationState {
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub order: Vec<FormulaLocation>,
    /// Traversal seed retained across consecutive forced calculations. A spill
    /// consumer discovered in a later phase stays in that phase next time;
    /// normal and explicit range calculations establish a fresh seed.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub forced_order: Option<Vec<FormulaLocation>>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub blocked_spills: Vec<(FormulaLocation, (u32, u32))>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub changed: Vec<(String, (u32, u32, u32, u32))>,
    #[serde(default, skip_serializing_if = "std::collections::BTreeSet::is_empty")]
    pub dirty: std::collections::BTreeSet<FormulaLocation>,
    /// Cells whose UDFs most recently registered Application.Volatile True.
    #[serde(default, skip_serializing_if = "std::collections::BTreeSet::is_empty")]
    pub volatile: std::collections::BTreeSet<FormulaLocation>,
    /// Explicit False also overrides volatility from built-in functions.
    /// Retain the formula text so replacing a formula cannot inherit a stale flag.
    #[serde(default, skip_serializing_if = "std::collections::BTreeSet::is_empty")]
    pub nonvolatile: std::collections::BTreeSet<(FormulaLocation, String)>,
}

impl RecalculationState {
    pub fn is_empty(&self) -> bool { self.order.is_empty() && self.forced_order.is_none() && self.blocked_spills.is_empty() && self.changed.is_empty() && self.dirty.is_empty() && self.volatile.is_empty() && self.nonvolatile.is_empty() }
    pub fn pending(&self) -> bool { !self.changed.is_empty() || !self.dirty.is_empty() }
    pub fn change(&mut self, sheet: String, range: (u32, u32, u32, u32)) {
        if let Some((last_sheet, last)) = self.changed.last_mut() {
            if *last_sheet == sheet {
                if last.0 == range.0 && last.2 == range.2 && range.1 <= last.3.saturating_add(1) && last.1 <= range.3.saturating_add(1) {
                    last.1 = last.1.min(range.1); last.3 = last.3.max(range.3); return;
                }
                if last.1 == range.1 && last.3 == range.3 && range.0 <= last.2.saturating_add(1) && last.0 <= range.2.saturating_add(1) {
                    last.0 = last.0.min(range.0); last.2 = last.2.max(range.2); return;
                }
            }
        }
        self.changed.push((sheet, range));
    }
}

/// Workbook calculation settings recorded in OOXML's `calcPr` element.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub enum CalculationMode {
    #[default]
    #[serde(rename = "auto")]
    Automatic,
    #[serde(rename = "manual")]
    Manual,
    #[serde(rename = "autoNoTable")]
    AutomaticExceptTables,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
pub struct CalculationSettings {
    /// Calculation-host language context; it is not an OOXML workbook setting.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub dbcs_language: Option<bool>,
    /// Workbook calculation compatibility: absent OOXML metadata means version 1.
    /// Version 2 changes selected worksheet functions' surrogate-pair indexing.
    #[serde(skip_serializing_if = "legacy_compatibility")]
    pub compatibility_version: u8,
    pub mode: CalculationMode,
    pub on_save: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub completed: Option<bool>,
}

impl Default for CalculationSettings {
    fn default() -> Self {
        Self { mode: CalculationMode::Automatic, on_save: true, completed: None, compatibility_version: 1, dbcs_language: None }
    }
}

fn legacy_compatibility(version: &u8) -> bool { *version == 1 }

/// A workbook this one links to, as much of it as this file remembers.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ExternalBook {
    /// The sheets it names, in the order the file lists them 窶・a cached cell
    /// says which one it belongs to by that position, not by name.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub sheets: Vec<ExternalSheet>,
}

/// One sheet of a linked workbook, and the cells this file remembers of it.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ExternalSheet {
    pub name: String,
    /// `(row, column)` 窶・one-based row, zero-based column, as the rest of the
    /// IR counts them 窶・and the value that cell held.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cells: Vec<(u32, u32, CellValue)>,
}

/// Whether a sheet is shown in the tab strip.
///
/// A file says nothing for a sheet that is shown, `state="hidden"` for one a
/// person can bring back, and `state="veryHidden"` for one only a macro can.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum Visibility {
    #[default]
    Visible,
    Hidden,
    VeryHidden,
}

impl Visibility {
    pub fn is_shown(&self) -> bool {
        matches!(self, Visibility::Visible)
    }

    /// What the file calls it, or nothing where the file says nothing.
    pub fn stated(&self) -> Option<&'static str> {
        match self {
            Visibility::Visible => None,
            Visibility::Hidden => Some("hidden"),
            Visibility::VeryHidden => Some("veryHidden"),
        }
    }

    /// The other way round, from what the file says.
    pub fn of(stated: &str) -> Self {
        match stated {
            "hidden" => Visibility::Hidden,
            "veryHidden" => Visibility::VeryHidden,
            _ => Visibility::Visible,
        }
    }
}

/// The document hosted by a workbook tab. Chart tabs do not contain cells.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum SheetKind {
    #[default]
    Worksheet,
    Chart,
}

impl SheetKind {
    pub fn is_worksheet(&self) -> bool { matches!(self, Self::Worksheet) }
}

/// A dynamic-array anchor and its extent, in cells.
pub type CachedSpill = ((u32, u32), (u32, u32));

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Sheet {
    pub name: String,
    #[serde(default, skip_serializing_if = "SheetKind::is_worksheet")]
    pub kind: SheetKind,
    /// Whether the sheet is shown, and how firmly it is hidden when it is not.
    #[serde(default, skip_serializing_if = "Visibility::is_shown")]
    pub visibility: Visibility,
    pub rows: Vec<Row>,
    /// Whole-sheet, row, and column formatting also applies to cells that do
    /// not yet exist. Rows are one-based; columns are zero-based.
    #[serde(default, skip_serializing_if = "BandStyles::is_empty")]
    pub band_styles: BandStyles,
    /// Last calculated dynamic-array owners and their (width, height).
    /// Anchors use zero-based (column, row). None is an older serialized
    /// workbook without ownership; Some([]) is a known empty spill layout.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cached_spills: Option<Vec<CachedSpill>>,
    /// Previous used-cell bounds retained after clearing cells. Presence
    /// queries keep this extent until an explicit used-range query trims it.
    /// Zero-based (first column, first row, last column, last row).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pending_used_bounds: Option<(u32, u32, u32, u32)>,
    pub col_count: usize,
    pub col_widths: Vec<f32>,
    /// Explicit OOXML customWidth state, independent of the numeric width.
    /// Missing entries retain compatibility with older serialized workbooks.
    #[serde(default, skip_serializing_if = "std::collections::BTreeMap::is_empty")]
    pub col_custom_widths: std::collections::BTreeMap<u32, bool>,
    /// The Normal style's Latin face, used to measure digits in column widths.
    /// Text can still use the theme's East Asian face through normal_font.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub column_width_font: Option<String>,
    pub default_col_width: f32,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub default_col_width_explicit: bool,
    pub default_row_height: f32,
    /// True when the sheet pins its default row height (customHeight="1").
    /// Without the pin Excel throws the stated number away and derives the
    /// height from the fonts the columns wear.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub default_row_custom: bool,
    /// The font each stretch of columns wears, as `(first, last, name, size)`
    /// with columns zero-based and both ends inside. A cell that states no
    /// format of its own is written in its column's font, and so is the
    /// blank space in every row 窶・which is what a row's height is measured
    /// from when nothing in the row is taller.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub col_fonts: Vec<(u32, u32, String, f32)>,
    /// The workbook's Normal font, worn by every column no `<col>` dresses.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub normal_font: Option<(String, f32)>,
    /// The first font in the workbook's list, with a theme scheme resolved
    /// the way every other font's is. One level of a cell's indent is three
    /// of this font's spaces 窶・this one, and not the font the Normal style
    /// points at, which is a different entry in books openpyxl writes.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_font: Option<(String, f32)>,
    /// How many rows are held at the top of the view while the rest scrolls,
    /// and how many columns at the left. Both zero unless the sheet says
    /// `<pane state="frozen">`, which counts them in cells; a plain split
    /// states its position in twips and is a different thing entirely.
    #[serde(default, skip_serializing_if = "is_zero")]
    pub frozen_rows: u32,
    #[serde(default, skip_serializing_if = "is_zero")]
    pub frozen_cols: u32,
    pub merge_cells: Vec<MergeCell>,
    /// Zero-based indices of the columns that are hidden. Columns have no
    /// record of their own, so this sits beside `col_widths`.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub hidden_cols: Vec<u32>,
    /// The range the sheet declares it occupies, as `(start_row, start_col,
    /// end_row, end_col)` 窶・rows one-based, columns zero-based. Excel hands
    /// this range over when asked for a picture of the sheet, and it can reach
    /// past the last cell that holds anything.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub declared_range: Option<(u32, u32, u32, u32)>,
    /// The tables on the sheet, each dressed by a named style.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tables: Vec<Table>,
    /// What is drawn over the grid rather than in it: pictures and shapes.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub drawings: Vec<Drawing>,
    /// The notes a sheet keeps pinned open. A comment Excel hides is not part
    /// of the picture; one the workbook pins is, and two of the conformance
    /// corpus's workbooks pin ninety between them.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub comments: Vec<Comment>,
    /// Threaded comments and replies. `cell` is zero-based (row, column), as
    /// in legacy comments; replies share the cell and name their parent id.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub threaded_comments: Vec<ThreadedComment>,
    /// The filter a sheet is under, if any.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<AutoFilter>,
    /// The links its cells carry, each already resolved to where it goes.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub hyperlinks: Vec<Hyperlink>,
    /// Whether the sheet is protected 窶・`<sheetProtection sheet="1">`, and not
    /// merely the presence of that element, which Excel also writes to record
    /// the options of a sheet whose protection is off. Six of the seven
    /// protected-looking workbooks in the sweep are the second kind.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub protected: bool,
    /// The buttons and tick boxes on the grid, from the same legacy drawing
    /// that carries the notes.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub form_controls: Vec<FormControl>,
    /// What its ranges will accept.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub validations: Vec<DataValidation>,
    /// The rules that change how its ranges look, lowest priority first.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub conditional_rules: Vec<ConditionalRule>,
    /// Unsupported elements found in this sheet (e.g. "Chart", "PivotTable", "Drawing")
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub unsupported_elements: Vec<String>,
    /// The colour of the sheet's tab, as a resolved `RRGGBB` hex string, where
    /// the file gives one in `<sheetPr><tabColor>`. Excel paints the tab this
    /// colour; a sheet with no `<tabColor>` leaves it `None`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<String>,
    /// A stable identity the parser stamps (the sheet's position in the file),
    /// so an editor round-trip can tell a rename (same id, new name) from an
    /// add-and-delete that happens to leave the count unchanged. `None` on a
    /// sheet the editor added, and on a workbook not built by the parser.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub origin_id: Option<u32>,
}

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct BandStyles {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet: Option<CellStyle>,
    #[serde(default, skip_serializing_if = "std::collections::BTreeMap::is_empty", with = "band_style_entries")]
    pub rows: std::collections::BTreeMap<u32, CellStyle>,
    #[serde(default, skip_serializing_if = "std::collections::BTreeMap::is_empty", with = "band_style_entries")]
    pub columns: std::collections::BTreeMap<u32, CellStyle>,
}

// JavaScript Map values disappear during JSON.stringify. Use entry arrays so
// workbook snapshots retain band styles across browser and worker boundaries.
mod band_style_entries {
    use super::CellStyle;
    use serde::{Deserialize, Deserializer, Serialize, Serializer};
    use std::collections::BTreeMap;

    pub fn serialize<S: Serializer>(styles: &BTreeMap<u32, CellStyle>, serializer: S) -> Result<S::Ok, S::Error> {
        styles.iter().collect::<Vec<_>>().serialize(serializer)
    }

    pub fn deserialize<'de, D: Deserializer<'de>>(deserializer: D) -> Result<BTreeMap<u32, CellStyle>, D::Error> {
        Ok(Vec::<(u32, CellStyle)>::deserialize(deserializer)?.into_iter().collect())
    }
}

impl BandStyles {
    pub fn is_empty(&self) -> bool { self.sheet.is_none() && self.rows.is_empty() && self.columns.is_empty() }
}

/// A filter over a range, and what each filtered column is testing for.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct AutoFilter {
    pub start_row: u32, // 1-based
    pub start_col: u32, // 0-based
    pub end_row: u32,   // 1-based
    pub end_col: u32,   // 0-based
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<AutoFilterColumn>,
}

/// One column's test, written the way VBA states it: `"apple"`, `">15"`,
/// `"<>banana"`. Two entries are joined by `either`.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct AutoFilterColumn {
    /// One-based column within the filtered range.
    pub field: u32,
    pub criteria: Vec<String>,
    /// True when the criteria are joined by xlOr rather than xlAnd.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub either: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MergeCell {
    pub start_row: u32, // 1-based
    pub start_col: u32, // 0-based
    pub end_row: u32,   // 1-based
    pub end_col: u32,   // 0-based
}

impl Sheet {
    pub fn has_default_column_width(&self) -> bool { self.default_col_width_explicit || self.default_col_width > 0.0 }
    pub fn column_has_custom_width(&self, column: u32) -> bool {
        self.col_custom_widths.get(&column).copied().unwrap_or_else(|| {
            self.col_widths.get(column as usize).is_some_and(|width| *width > 0.0)
        })
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Row {
    pub index: u32,
    pub cells: Vec<Cell>,
    pub height: Option<f32>,
    /// True when the row's height is pinned (customHeight="1"). Without the
    /// pin the stored height is only a cache from the machine that wrote the
    /// file 窶・Excel recomputes the height from the row's content on open.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub custom_height: bool,
    /// The font of the row's own format (`s=` with customFormat="1"). The
    /// row wears it across all its columns, and its default height becomes
    /// the row's base 窶・which is how a row sinks below the sheet default
    /// without any cell saying a word.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_font: Option<(String, f32)>,
    /// The row says a thick rule runs along its top or bottom edge. Excel
    /// keeps a pixel of room for each when it works a row's height out.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub thick_top: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub thick_bottom: bool,
    /// A hidden row keeps everything it holds; only the display is affected.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub hidden: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Cell {
    pub col: u32,
    pub value: CellValue,
    pub style: CellStyle,
    /// Original formula string (e.g. "=SUM(A1:A3)"), if any
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    /// Exact formula code units when the source contains unpaired surrogates.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula_utf16: Option<Vec<u16>>,
    /// The stretches the text is dressed in, when it is not all dressed alike.
    /// `value` still holds the whole of the text, so a reader that does not
    /// care about the dressing needs no change.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub runs: Vec<TextRun>,
    /// The block of cells this one shares an ARRAY formula with, as
    /// `(start_row, start_col, end_row, end_col)` 窶・rows one-based, columns
    /// zero-based, the way a merge is written. One formula is worked out
    /// once for the block and its answers dealt across the cells; every
    /// member holds the formula's text in `formula` and its own share of
    /// the answer in `value`. The file writes the formula on the top-left
    /// cell alone (`<f t="array" ref="窶ｦ">`) and only values on the rest.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub array_block: Option<(u32, u32, u32, u32)>,
    /// A dynamic-array formula's part in a spill: the cell whose formula
    /// spills, or one of the cells its answer spilled into.
    #[serde(default, skip_serializing_if = "Spill::is_none")]
    pub spill: Spill,
}

impl Cell {
    /// Exact code units only while they still describe this cell's formula.
    pub fn formula_units(&self) -> Option<&[u16]> {
        let units = self.formula_utf16.as_deref()?;
        let formula = self.formula.as_deref()?;
        let display = String::from_utf16_lossy(units);
        (display.strip_prefix('=').unwrap_or(&display) == formula.strip_prefix('=').unwrap_or(formula)).then_some(units)
    }

    /// Rewrite references while retaining unpaired units in formula literals.
    /// A failed rewrite leaves both source representations untouched.
    pub fn rewrite_formula(&mut self, transform: impl FnOnce(&str) -> Result<String, String>) -> Result<(), String> {
        let Some(formula) = self.formula.as_deref() else { return Ok(()) };
        if let Some(units) = self.formula_units() {
            let rewritten = oxicells_calc::parser::transform_utf16_formula(units, transform)?;
            let text = String::from_utf16_lossy(&rewritten);
            let exact = String::from_utf16(&rewritten).err().map(|_| rewritten);
            self.formula = Some(text);
            self.formula_utf16 = exact;
        } else {
            let rewritten = transform(formula)?;
            self.formula = Some(rewritten);
            self.formula_utf16 = None;
        }
        Ok(())
    }
}

/// Where a cell stands in a dynamic array's spill.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub enum Spill {
    /// An ordinary cell.
    #[default]
    None,
    /// Its formula was written as a dynamic array (`.Formula2`) and spills
    /// a block of answers across the cells beside it.
    Dynamic,
    /// Its value is a share of the spill of the formula above or left of it;
    /// it holds no formula of its own and goes when the spill does.
    Member,
}

impl Spill {
    pub fn is_none(&self) -> bool {
        *self == Spill::None
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum CellValue {
    Empty,
    String(String),
    /// UTF-16 text containing unpaired surrogate code units.
    Utf16(Vec<u16>),
    Number(f64),
    Boolean(bool),
    Error(String),
}

impl CellValue {
    pub fn from_utf16(units: Vec<u16>) -> Self {
        match String::from_utf16(&units) {
            Ok(text) => Self::String(text),
            Err(_) => Self::Utf16(units),
        }
    }

    pub fn string_units(&self) -> Option<std::borrow::Cow<'_, [u16]>> {
        match self {
            Self::String(text) => Some(text.encode_utf16().collect::<Vec<_>>().into()),
            Self::Utf16(units) => Some(units.as_slice().into()),
            _ => None,
        }
    }

    pub fn is_text(&self) -> bool {
        matches!(self, Self::String(_) | Self::Utf16(_))
    }

    pub fn display(&self) -> String {
        match self {
            CellValue::Empty => String::new(),
            CellValue::String(s) => s.clone(),
            // Rendering may use a replacement glyph; storage must keep the units.
            CellValue::Utf16(units) => String::from_utf16_lossy(units),
            CellValue::Number(n) => {
                if *n == (*n as i64) as f64 {
                    format!("{}", *n as i64)
                } else {
                    format!("{}", n)
                }
            }
            CellValue::Boolean(b) => {
                if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            CellValue::Error(e) => e.clone(),
        }
    }
}

/// A range a sheet treats as a table, and the style it wears. Excel dresses
/// these itself: no cell inside carries the header's fill or the banding.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Table {
    /// Source package part, used to preserve this table through simultaneous edits.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub origin_part: Option<String>,
    /// What a formula calls it. `tblNomina[[#This Row],[DATE]]` names this.
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub name: String,
    /// The heading of each column, left to right, so that a formula naming one
    /// can be told which column of `start_col..=end_col` it meant.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<String>,
    pub start_row: u32, // 1-based
    pub start_col: u32, // 0-based
    pub end_row: u32,   // 1-based
    pub end_col: u32,   // 0-based
    /// The built-in style's name, e.g. "TableStyleMedium7".
    pub style: Option<String>,
    /// How many rows at the top are the header. Excel writes 1 unless told.
    pub header_rows: u32,
    /// How many rows at the bottom are the totals row: 0 or 1.
    #[serde(default, skip_serializing_if = "is_zero")]
    pub totals_rows: u32,
    /// Per-column totals settings, retained while the totals row is hidden.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub totals: Vec<TableTotals>,
    /// The table has an insertion row instead of a data body.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub insert_row: bool,
    /// Whether every other row is shaded.
    pub banded_rows: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub banded_columns: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub emphasize_first_column: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub emphasize_last_column: bool,
    #[serde(default = "yes")]
    pub show_auto_filter: bool,
    /// Complete filter state, including criteria not interpreted by the engine.
    /// None creates a fresh filter when filter buttons are enabled.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_filter_xml: Option<String>,
    /// Zero-based columns whose AutoFilter button is hidden.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub hidden_filter_buttons: Vec<u32>,
    /// The colour the style dresses the table in, as six hex digits. A built-in
    /// style takes it from the workbook's theme, so it is resolved on the way
    /// in rather than left as a name to look up later.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accent: Option<String>,
    /// The banded rows' fill, which is the accent under a tint.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub band: Option<String>,
    /// The rule Excel draws along every row of a `TableStyleMedium` table: the
    /// accent again, under a lighter tint than the banding.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rule: Option<String>,
    /// The rule the table draws round its whole self, which it states as an
    /// index into the workbook's differential formats rather than as a border
    /// on any cell. Fifteen of the corpus's workbooks carry one 窶・the whole
    /// `procurement_contractor` family 窶・and its bottom edge is a rule
    /// thirteen hundred pixels long that no cell mentions.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline: Option<BorderLine>,
}

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct TableTotals {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub function: Option<String>,
    /// Label and custom formula text as lossless UTF-16 code units.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label: Option<Vec<u16>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<Vec<u16>>,
}

impl TableTotals {
    pub fn is_empty(&self) -> bool { self.function.is_none() && self.label.is_none() && self.formula.is_none() }
}

impl Table {
    /// Whether the enabled table filter has criteria, rather than just buttons.
    pub fn has_filter_criteria(&self) -> bool {
        if !self.show_auto_filter || self.header_rows == 0 { return false; }
        let Some(xml) = self.auto_filter_xml.as_deref() else { return false; };
        let mut reader = quick_xml::Reader::from_str(xml);
        loop {
            match reader.read_event() {
                Ok(quick_xml::events::Event::Start(element) | quick_xml::events::Event::Empty(element)) => {
                    if matches!(element.name().as_ref().rsplit(|byte| *byte == b':').next(),
                        Some(b"filters" | b"customFilters" | b"top10" | b"dynamicFilter" | b"colorFilter" | b"iconFilter")) { return true; }
                }
                Ok(quick_xml::events::Event::Eof) | Err(_) => return false,
                _ => {}
            }
        }
    }
}

/// A stretch of a cell's text that is dressed differently from the rest of it.
/// A cell holds none of these when all of its text is dressed the same; when it
/// does hold them they cover the whole of the text, in order, and each field
/// left empty means "as the cell itself says".
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct TextRun {
    pub text: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub italic: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub underline: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// Whether the run carried an `<rPr>` of its own.
    ///
    /// It is not the same question as whether any of the fields above are
    /// set. An `<rPr>` REPLACES the cell's font rather than overriding parts
    /// of it, so a run that carries one and does not say `<b/>` is NOT bold
    /// even in a bold cell 窶・while a run with no `<rPr>` at all wears the
    /// cell's font whole. Both look identical in the fields above, so the
    /// fact that the element was there has to be kept.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub dressed: bool,
    /// "superscript" or "subscript" when the run is raised or lowered.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_align: Option<String>,
}

/// One edge of a cell, and how it is drawn.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BorderLine {
    /// The kind OOXML names: "thin", "medium", "thick", "hair", "dotted",
    /// "dashed", "double", and the rest. Each is drawn differently 窶・a thick
    /// rule is three pixels where a thin one is a single pixel.
    pub style: String,
    /// The colour as six hex digits, when the edge names one.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
}

fn is_zero(value: &u32) -> bool {
    *value == 0
}

fn is_zero_i32(value: &i32) -> bool {
    *value == 0
}

/// A corner of a drawing: the cell it hangs from, and how far into that cell
/// it sits. The offsets are EMU 窶・914400 to the inch, so 9525 to a pixel at
/// 96 dpi.
#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub struct Anchor {
    /// Zero-based, like `Cell::col`.
    pub col: u32,
    pub col_off: i64,
    /// Zero-based, unlike `Row::index`, which is how the drawing part states it.
    pub row: u32,
    pub row_off: i64,
}

/// Something drawn over the grid rather than in it.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Drawing {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// ProgID of a worksheet-embedded OLE object, joined through `shapeId`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_prog_id: Option<String>,
    /// Source OLE drawing name when this object was duplicated by the VBA host.
    /// The editor uses it to copy the embedded package and worksheet anchor.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_source_name: Option<String>,
    /// Point position of a newly added or duplicated OLE object for its VML fallback.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_clone_position: Option<(f64, f64)>,
    /// Complete CFB package for an OLE object created from a VBA file input.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_embedded_package: Option<Vec<u8>>,
    /// Excel `Package` OLE server instantiated from `ClassType` without a file.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_package_is_empty: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_clone_visible: Option<bool>,
    /// Excel `OLEObject.Placement`: MoveAndSize=1, Move=2, FreeFloating=3.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_placement: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_type: Option<i64>,
    /// Whether Excel stores an OLE icon aspect (`dvAspect="DVASPECT_ICON"`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_display_as_icon: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_auto_load: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_enabled: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_locked: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_print_object: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_shadow: Option<bool>,
    /// Excel `OLEObject.OnAction`, stored in `oleObject/objectPr/@macro`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_on_action: Option<String>,
    /// Excel `OLEObject.Border.Color`, retained from the VML fallback shape.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_border_color: Option<String>,
    /// Excel `OLEObject.Border.Weight` enum, retained from the VML fallback shape.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_border_weight: Option<i64>,
    /// Excel `OLEObject.Interior.Color`, retained from the VML fallback shape.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ole_interior_color: Option<String>,
    /// Where its top-left corner hangs.
    pub from: Anchor,
    /// Where its bottom-right corner hangs, when the anchor names a second
    /// cell. A drawing anchored to one cell states its size instead.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub to: Option<Anchor>,
    /// The size the anchor states outright, in EMU.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub extent: Option<(i64, i64)>,
    pub kind: DrawingKind,
    /// Where it sits as fractions of its parent's box, for a shape that hangs
    /// on a chart rather than on the grid: a chart keeps the boxes it is
    /// annotated with in a part of its own, anchored that way.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame: Option<Frame>,
    /// True when the anchor above was WORKED OUT rather than read: a shape
    /// inside a group is placed by mapping its own box through the group's
    /// transform, and it then hangs from the cell the GROUP hangs from with
    /// the whole distance folded into the offset. Such an offset routinely
    /// runs past its cell, which a file's own never does 窶・and Excel holds a
    /// file's offset inside its cell, so the two cannot be treated alike.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub grouped: bool,
    /// Name of the parent DrawingML group when this is one of its children.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub group_name: Option<String>,
    /// Rotation of the parent DrawingML group, in sixtieth-thousandths of a degree.
    #[serde(default, skip_serializing_if = "is_zero_i32")]
    pub group_rotation: i32,
    /// DrawingML coordinate systems for the parent group, shared by its leaves.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub group_transform: Option<GroupTransform>,
    /// DrawingML transform of this leaf when it belongs to a group being created.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub group_child_transform: Option<(i64, i64, i64, i64)>,
    /// True when a leaf came through at least one nested DrawingML group.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub group_nested: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub struct GroupTransform {
    pub offset: (i64, i64),
    pub extent: (i64, i64),
    pub child_offset: (i64, i64),
    pub child_extent: (i64, i64),
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct PictureStyle {
    /// Physical source extent for vector-backed pictures whose raster preview
    /// has different dimensions. Kept in points for Excel's
    /// `RelativeToOriginalSize` shape operations; it is metadata, not a style.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub original_size_points: Option<(f64, f64)>,
    /// DrawingML crop margins in 1/1000 percent units.
    #[serde(default)]
    pub crop: (i32, i32, i32, i32),
    /// DrawingML `a:lum` adjustments, in 1/1000 percent units.
    #[serde(default)]
    pub brightness: i32,
    #[serde(default)]
    pub contrast: i32,
    /// MsoPictureColorType: automatic, grayscale, black-and-white, watermark.
    #[serde(default = "default_picture_color_type")]
    pub color_type: i64,
    #[serde(default)]
    pub transparency_color: Option<String>,
    #[serde(default = "default_transparent_background")]
    pub transparent_background: bool,
}

fn default_picture_color_type() -> i64 { 1 }
fn default_transparent_background() -> bool { true }

impl Default for PictureStyle {
    fn default() -> Self {
        Self {
            original_size_points: None,
            crop: (0, 0, 0, 0),
            brightness: 0,
            contrast: 0,
            color_type: 1,
            transparency_color: None,
            transparent_background: true,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum DrawingKind {
    /// A picture, holding the bytes of the part it names. They are left out of
    /// the serialised form: what reads the IR back is an editor that has the
    /// file beside it, and a base64 image in every sheet would dwarf the rest.
    Picture {
        #[serde(skip)]
        bytes: Vec<u8>,
        #[serde(default)]
        format: PictureStyle,
    },
    /// A chart, whose picture is drawn from a part of its own.
    Chart(Chart),
    /// A shape Excel fills and rules. The corpus draws 2176 of its 2245
    /// shapes as a line, a rectangle or a rounded one.
    Shape(Shape),
    /// Anything else the drawing part holds 窶・a group, a text box.
    Other,
}

/// A note pinned open on the sheet.
///
/// Its box hangs from two cells, the way a drawing's does. The VML the note
/// lives in states that twice over 窶・once as a margin in points, cached from
/// wherever the file was last written, and once as `<x:Anchor>`, which is what
/// Excel lays it out from and what these corners hold.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Comment {
    /// Name in the worksheet comment part's shared author table.
    #[serde(default)]
    pub author: String,
    /// Original UTF-16 when the note contains unpaired surrogate units.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub utf16_text: Option<Vec<u16>>,
    pub from: Anchor,
    /// How big the box is, in points, as the style states it. It is what is
    /// used when the anchor's far corner names a cell the picture does not
    /// reach.
    pub size: (f32, f32),
    /// The anchor's far corner. Excel draws a note between its two corners:
    /// `002`'s note says `height:58pt` and Excel draws 78 pixels, which is
    /// what the anchor's rows come to. A note Excel wrote itself agrees 窶・    /// its far corner's offset is the width in pixels exactly.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub to: Option<Anchor>,
    /// The note's text, dressed the way a shape's is.
    pub text: ShapeText,
    /// Six hex digits; Excel's own note is FFFFE1.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    /// Whether the VML note fill is drawn. Excel keeps the fill color even
    /// when it writes `filled="f"`.
    #[serde(default = "yes")]
    pub fill_visible: bool,
    /// Fill opacity as VBA transparency, where 0 is opaque and 1 is clear.
    #[serde(default)]
    pub fill_transparency: f32,
    /// RGB color of the VML note outline. Excel's default outline is black.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_color: Option<String>,
    /// Whether the VML note outline is drawn.
    #[serde(default = "yes")]
    pub line_visible: bool,
    /// Outline width in points. Excel creates note outlines at 0.75 pt.
    #[serde(default = "default_comment_line_weight")]
    pub line_weight: f32,
    /// Outline transparency as VBA transparency, where 0 is opaque and 1 is clear.
    #[serde(default)]
    pub line_transparency: f32,
    /// The cell the note is ABOUT, counted from zero, which is not where the
    /// box sits: Excel draws a hairline from the box back to it, and that line
    /// is the only thing on the sheet that says which cell a note belongs to
    /// once the box has been dragged away.
    #[serde(default)]
    pub cell: (u32, u32),
    /// Whether the workbook pins the note open. A pinned note is drawn as its
    /// box; one left closed is drawn only as the red corner Excel puts on the
    /// cell, which is all Excel itself shows until the pointer arrives.
    #[serde(default = "yes")]
    pub visible: bool,
}

fn default_comment_line_weight() -> f32 {
    0.75
}

/// One message in an Excel threaded-comment conversation.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ThreadedComment {
    /// Zero-based (row, column) of the cell carrying the conversation.
    pub cell: (u32, u32),
    pub id: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub parent_id: Option<String>,
    pub person_id: String,
    #[serde(default)]
    pub person_name: String,
    #[serde(default)]
    pub person_provider_id: String,
    #[serde(default)]
    pub person_user_id: String,
    pub date: String,
    pub text: String,
    /// Original UTF-16 code units when text contains unpaired surrogates.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub utf16_text: Option<Vec<u16>>,
    #[serde(default)]
    pub resolved: bool,
}

/// A note is pinned open unless something says otherwise, which is what every
/// `Comment` built before this field existed meant.
fn yes() -> bool {
    true
}

/// The parts of a look that a conditional format puts over a cell's own.
///
/// A differential format states only what it changes, which is why every field
/// is optional: a rule that turns text red says nothing about the fill, and the
/// cell keeps the fill it already had.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct DiffStyle {
    /// Original differential format, retained so edits preserve fields the IR does not model.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub source_xml: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bg_color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
}

/// A rule that changes how a range looks when what it tests comes out true.
///
/// Excel keeps these on the sheet rather than on the cells, and applies them in
/// `priority` order over whatever the cell already wears.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ConditionalRule {
    /// Excel's own word: `expression`, `cellIs`, `containsText`, `beginsWith`,
    /// `duplicateValues`, and the rest.
    pub kind: String,
    /// For `cellIs`: `equal`, `between`, `greaterThan` and so on.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub operator: Option<String>,
    /// What `containsText` and its relatives look for.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    /// The rule's own formulas, in the order the file writes them. They are
    /// written for the top-left cell of the first range and are relative to it.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub formulas: Vec<String>,
    /// Number of criteria in a colour scale, when modelled.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_scale_points: Option<u8>,
    /// The visual rule body, retained so copying preserves its settings.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub visual_xml: Option<String>,
    /// Linked extended visual settings, when the file has them.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub visual_extension_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub visual_extension_xml: Option<String>,
    /// The rule exists only in the worksheet's x14 extension, without a base rule.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub visual_extension_only: bool,
    /// How many a `top10` rule takes: the 10 of "Top 10 Items", or the
    /// percentage when `percent` is set.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rank: Option<u32>,
    /// A `top10` rule counting from the bottom instead of the top.
    #[serde(default)]
    pub bottom: bool,
    /// A `top10` rule whose `rank` is a percentage of the range rather than a
    /// count of cells.
    #[serde(default)]
    pub percent: bool,
    /// Settings of an above-average conditional formatting rule.
    #[serde(default = "yes")]
    pub above_average: bool,
    #[serde(default)]
    pub equal_average: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub standard_deviations: Option<u32>,
    /// Lower runs first. Excel writes it on every rule.
    #[serde(default)]
    pub priority: i32,
    /// Original file priority, used to identify a rule when its priority changes.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub origin_priority: Option<i32>,
    /// Whether a match stops the rules below it from being tried.
    #[serde(default)]
    pub stop_if_true: bool,
    /// The ranges the rule covers, `(start_row, start_col, end_row, end_col)`
    /// counted from zero.
    pub ranges: Vec<(u32, u32, u32, u32)>,
    /// The look it puts on, already resolved from the `dxfId` it names.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<DiffStyle>,
}

/// A rule about what a range of cells will accept.
///
/// Only one kind of rule is drawn: a `list` puts an arrow on the cell the
/// cursor is in, and that arrow is the whole of what Excel shows. The rest
/// decide what a typed value is allowed to be, and show nothing until it is
/// not.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct DataValidation {
    /// Excel's own word: `list`, `whole`, `decimal`, `date`, `time`,
    /// `textLength`, `custom`. Absent in the file means `any`, which allows
    /// everything and is kept as `any` rather than as nothing.
    pub kind: String,
    /// `between`, `notBetween`, `equal`, `greaterThan` and the rest. Only the
    /// numeric and date kinds use it.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub operator: Option<String>,
    /// The ranges the rule covers, each `(start_row, start_col, end_row,
    /// end_col)` counted from zero.
    pub ranges: Vec<(u32, u32, u32, u32)>,
    /// For a list, either the items spelled out 窶・`"Yes,No"` 窶・or a reference
    /// to the cells holding them. For the others, the bound being tested.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula1: Option<String>,
    /// The second bound, where the operator takes two.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula2: Option<String>,
    /// Whether an empty cell passes the rule.
    #[serde(default)]
    pub allow_blank: bool,
    /// Whether the in-cell arrow is drawn for a list.
    ///
    /// The file's own attribute is `showDropDown`, and it means the opposite
    /// of what it reads like: Excel writes `1` when the person turned the
    /// in-cell dropdown OFF. This field is already the right way round.
    #[serde(default = "yes")]
    pub in_cell_dropdown: bool,
    /// Display and title attributes are kept separately so VBA and the
    /// browser can reproduce the rule rather than only its formula.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error_style: Option<String>,
    #[serde(default)]
    pub show_input: bool,
    #[serde(default)]
    pub show_error: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prompt_title: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error_title: Option<String>,
    /// What Excel says when a value is refused.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error: Option<String>,
    /// What it says while the cell is selected.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prompt: Option<String>,
}

/// A button, tick box or other control sitting on the grid.
///
/// These share a part with a sheet's notes 窶・the legacy VML drawing 窶・which is
/// why a sheet can hold a hundred of them and not a single note. Excel draws
/// them over the cells the way it draws a shape.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FormControl {
    /// Excel's own word for what it is: `Button`, `Checkbox`, `Radio`, `Drop`,
    /// `List`, `Spin`, `Scroll`, `GBox`, `Label`. Kept as the file spells it,
    /// because the file's vocabulary is the one that will grow.
    pub kind: String,
    pub from: Anchor,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub to: Option<Anchor>,
    /// The caption written on it, where it has one. A tick box put beside a
    /// cell usually has none, and takes its meaning from the cell.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    /// Whether a tick box or radio button is set.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub checked: bool,
}

/// A link a cell carries.
///
/// How a link looks is already settled before this: Excel dresses the cell
/// through the Hyperlink named style, which the cell holds like any other
/// style, so it is blue and underlined without anything here. What a style
/// cannot say is where the link goes, and that is the whole of this.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Hyperlink {
    /// The first cell the link spans, counted from zero, as `(row, col)`.
    pub cell: (u32, u32),
    /// The last cell in a range-spanning anchor, also zero-based `(row, col)`.
    /// `None` means the anchor is the single cell in `cell`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub end_cell: Option<(u32, u32)>,
    /// A URL, or 窶・when `internal` 窶・a place in this workbook, spelled the way
    /// the file spells it: `Sheet2!A1`, or a defined name.
    pub target: String,
    /// Whether the target is somewhere else in this same workbook.
    #[serde(default)]
    pub internal: bool,
    /// An optional in-workbook location paired with an external address.
    /// OOXML stores this separately from the relationship target.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sub_address: Option<String>,
    /// What Excel shows while the pointer rests on the cell.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,
}

/// One step of a shape's own outline, in the space the path states.
#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub enum PathStep {
    MoveTo(i64, i64),
    LineTo(i64, i64),
    /// Two control points and the point the curve ends on.
    CurveTo(i64, i64, i64, i64, i64, i64),
    Close,
}

/// A shape drawn from its own outline rather than from a preset.
///
/// The corpus holds sixteen of them across four workbooks 窶・one path each,
/// 160 straight segments, 18 curves and 10 closes 窶・and one of them is the
/// curly brace beside `002`'s notes, in the lowest-scoring workbook there is.
/// The points are stated in a space of the path's own, which is mapped onto
/// the box the anchors give the shape.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ShapePath {
    pub across: i64,
    pub down: i64,
    pub steps: Vec<PathStep>,
}

/// A preset shape, with what it is painted and ruled with.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Shape {
    /// The preset geometry OOXML names: "rect", "line", "roundRect", 窶ｦ
    pub geometry: String,
    /// True for an Office text box (`cNvSpPr txBox="1"`), rather than a
    /// regular preset shape that happens to contain text.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub text_box: bool,
    /// Six hex digits, when the shape is filled.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    /// Office ColorFormat metadata for a DrawingML theme fill.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_theme: Option<usize>,
    /// Excel's backward-compatible DrawingML color index for this fill.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_legacy_color_index: Option<u8>,
    /// Office ColorFormat brightness encoded by DrawingML luminance modifiers.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill_brightness: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<ShapeLine>,
    /// What a preset's adjust handles are set to, by name 窶・`adj`, `adj1`,
    /// `adj2` 窶・in the hundred-thousandths a preset states them in. A preset
    /// left alone carries none and takes its own defaults.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub adjusts: Vec<(String, i64)>,
    /// The outline the shape draws itself with, when it has one of its own
    /// rather than a preset's.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub path: Option<ShapePath>,
    /// A shape whose box is flipped: a line drawn from the other corner.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub flip_h: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub flip_v: bool,
    /// How far round the shape is turned, in sixtieth-thousandths of a degree,
    /// clockwise. The turn is about the box's centre, and the anchor states
    /// the box the turn LEAVES it in: a tall shape turned a quarter is hung
    /// from an anchor as wide as the shape was tall.
    #[serde(default, skip_serializing_if = "is_zero_i32")]
    pub rotation: i32,
    /// What the shape says, when it says anything. 198 of the corpus's 2202
    /// shapes hold text, and they are the large ones: a banner across the top
    /// of a sheet, a heading over a table.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<ShapeText>,
}

/// The text a shape holds, and how it sits in the shape's box.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ShapeText {
    pub paragraphs: Vec<ShapeParagraph>,
    /// "t", "ctr" or "b" 窶・where the block of lines sits in the box. Excel
    /// leaves a shape's text at the top.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub anchor: Option<String>,
    /// Left, top, right and bottom, in EMU. Excel's defaults are 91440 and
    /// 45720 窶・a tenth and a twentieth of an inch.
    pub insets: (i64, i64, i64, i64),
    /// True when a line too long for the box breaks rather than running on.
    pub wrap: bool,
    /// True when the body says `vertOverflow="clip"`. Excel then draws only
    /// the lines that fit the box and drops the rest; with the attribute
    /// absent, or saying `overflow`, every line is drawn and the block hangs
    /// out of the box. 256 of the corpus's shapes clip, 322 say nothing and
    /// 7 say overflow.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub clip: bool,
}

/// One piece of a shape's paragraph, when the paragraph is not written all
/// one way.
///
/// Runs retain the font properties that vary within a paragraph. Missing run
/// values inherit the paragraph's corresponding property. The texts of the runs
/// add up to the paragraph's own text, including each `<a:br/>` line break.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ShapeRun {
    pub text: String,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub underline: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<f32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// DrawingML theme slot used by the run's text fill, in VBA's 1窶・2 order.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_theme: Option<usize>,
    /// DrawingML tint/shade brightness on a theme text fill.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_brightness: Option<f32>,
    /// Office's legacy spreadsheet scheme index carried by a14 on sRGB fills.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy_color_index: Option<u8>,
}

/// One paragraph of a shape's text, dressed the way its first run is.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ShapeParagraph {
    pub text: String,
    /// The pieces the text is written in. Empty for a paragraph written all
    /// one way, which is the common case.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub runs: Vec<ShapeRun>,
    /// "l", "ctr", "r" or "just".
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    /// Points. A run states hundredths of one.
    pub size: f32,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub italic: bool,
    // A shape's underline is NOT here on purpose. `u="sng"` sits on a run,
    // and a shape paragraph is dressed by its FIRST run 窶・in `sanko_tool`
    // one paragraph holds an underlined heading, two `<a:br/>`s and then the
    // body, so wearing the heading's underline across the paragraph underlines
    // all of it and costs that workbook 0.0092. It needs per-run dressing of a
    // shape's line, the way a cell's runs are dressed, not a paragraph flag.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub face: Option<String>,
    /// The charset the run states beside the face, when it states one.
    /// Excel's answer for a face this machine has not got turns on this and
    /// on nothing else 窶・not the name, not the PANOSE the file carries.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub charset: Option<i32>,
    /// The `pitchFamily` beside it, which chooses the METRICS Excel lays the
    /// substitute out with 窶・a different question from which face it draws
    /// (SX101). It travels with the face the same way the charset does.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pitch_family: Option<i32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// The line pitch the paragraph pins, in points, when it pins one.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_pitch: Option<f32>,
    /// The share of the font's own pitch the paragraph asks for, when it
    /// states one as a percentage rather than outright 窶・`glossary_05`'s
    /// flowchart sets every box at four fifths.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_scale: Option<f32>,
}

/// A graph drawn over the grid, as its own part states it.
///
/// Everything the picture needs is in the file: a chart caches the values it
/// plots beside the reference it plots them from, so nothing here waits on a
/// formula.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct Chart {
    /// Stable identity within the source document, retained across edits.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub origin_id: Option<u32>,
    /// An enabled title, including an explicitly empty title.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    /// "line", "bar", "pie" 窶・the first plot the chart holds.
    pub kind: String,
    /// The rectangle the axes enclose, as fractions of the chart's own box,
    /// when the chart pins one. Left out, Excel places the plot itself.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot: Option<Frame>,
    pub series: Vec<ChartSeries>,
    /// What the category axis is labelled with, shared by every series.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub categories: Vec<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value_axis: Option<ChartAxis>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub category_axis: Option<ChartAxis>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legend: Option<Legend>,
    /// What the chart is annotated with: text boxes and rules kept in a part
    /// of their own and placed as fractions of the chart's box. The corpus's
    /// charts carry their footnotes there rather than in a cell.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub shapes: Vec<Drawing>,
    /// The chart's own background, and the plot rectangle's.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub plot_fill: Option<String>,
}

/// A box stated as fractions of its parent's width and height.
#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize, Default)]
pub struct Frame {
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct ChartSeries {
    pub name: String,
    /// One per category. A gap in the data is a `None`, and the line either
    /// breaks across it or steps over it, as the chart says.
    pub values: Vec<Option<f64>>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<ShapeLine>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<Marker>,
    /// The points that dress differently from the rest of their series.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub points: Vec<ChartPoint>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub labels: Vec<DataLabel>,
    /// How the series dresses its labels, where a label does not say for
    /// itself: points, a face, and "t", "r", "b", "l" or "ctr" for the side
    /// of the point the label is set against.
    pub label_size: f32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub label_pos: Option<String>,
}

/// One point of a series, where it is drawn unlike its neighbours.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct ChartPoint {
    pub index: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub marker: Option<Marker>,
    /// How the line reaching this point is drawn, when the point restates it.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<ShapeLine>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct Marker {
    /// "circle", "diamond", "square", "none", 窶ｦ
    pub symbol: String,
    /// Points across, as the chart states it. Excel's own default is 7.
    pub size: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<String>,
}

/// A number written beside the point it belongs to.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct DataLabel {
    pub index: u32,
    /// How far the label is nudged from where it would otherwise sit, as
    /// fractions of the chart's box.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub offset: Option<(f64, f64)>,
    /// What it says, when the label says something other than the value.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    pub size: f32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub face: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    /// Which side of the point the label is set against, when this one
    /// differs from the rest of its series.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct ChartAxis {
    /// "b", "l", "r" or "t".
    pub position: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub min: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub max: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_unit: Option<f64>,
    /// "in", "out", "cross" or "none".
    pub major_tick: String,
    /// "nextTo", "low", "high" or "none".
    pub tick_labels: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<ShapeLine>,
    /// A gridline drawn across the plot at every major tick, when the axis
    /// asks for one.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub major_gridline: Option<ShapeLine>,
    /// Points.
    pub size: f32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub face: Option<String>,
    /// An axis the chart states but does not show.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub deleted: bool,
    /// "midCat" puts the first point on the axis itself; "between" puts it a
    /// half-step in.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cross_between: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct Legend {
    /// "r", "l", "t", "b" or "tr".
    pub position: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame: Option<Frame>,
    pub size: f32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub face: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ShapeLine {
    /// What the rule wears at each end 窶・"triangle", "arrow", "stealth",
    /// "oval", "diamond" 窶・when it wears anything. 76 of the corpus's rules
    /// carry one, and `glossary_05`'s flowchart is drawn almost entirely out
    /// of them.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub head_end: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tail_end: Option<String>,
    /// DrawingML headEnd len and w values (sm, med, lg).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub head_end_length: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub head_end_width: Option<String>,
    /// DrawingML tailEnd len and w values (sm, med, lg).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tail_end_length: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tail_end_width: Option<String>,
    /// Six hex digits.
    pub color: String,
    /// Office ColorFormat theme index for the line paint.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_theme: Option<usize>,
    /// Excel's backward-compatible DrawingML color index for the line paint.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub legacy_color_index: Option<u8>,
    /// Office ColorFormat brightness encoded by DrawingML luminance modifiers.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_brightness: Option<f32>,
    /// DrawingML shade applied by a theme style reference, independent of VBA brightness.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color_shade: Option<f32>,
    /// DrawingML patterned stroke fill, when the line uses an Office pattern.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pattern: Option<String>,
    /// Six hex digits used behind a patterned stroke.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub back_color: Option<String>,
    /// Office ColorFormat theme index for the patterned stroke background.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub back_color_theme: Option<usize>,
    /// Excel legacy color index for the patterned stroke background.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub back_legacy_color_index: Option<u8>,
    /// Office ColorFormat brightness of the patterned line background.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub back_color_brightness: Option<f32>,
    /// DrawingML shade applied to the patterned line background.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub back_color_shade: Option<f32>,
    /// Width in EMU. Excel draws a line of its own accord at 9525 窶・one pixel
    /// at 96 dpi 窶・when the shape states no width.
    pub width: i64,
    /// "dash", "dashDot", "sysDot" and the rest, when the rule is broken.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dash: Option<String>,
    /// What the rule's ends are cut with: "rnd", "sq", or flat when it says
    /// nothing. It is not decoration 窶・Excel lengthens every dash of a
    /// FLAT-capped or square-capped rule by a pixel when the rule is an odd
    /// number of pixels wide, and leaves a round-capped one alone. The two
    /// `dot` rules a workbook can hold are told apart by this and nothing
    /// else.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cap: Option<String>,
}

/// How a cell is drawn.
///
/// `#[serde(default)]` on the whole struct is load-bearing for the browser:
/// the editor builds a cell the moment someone types into an empty one, and it
/// has no style to give it. Requiring every field would mean the editor had to
/// know and restate the default of each 窶・which is to say, hold a second copy
/// of this struct that would drift.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
#[serde(default)]
pub struct CellStyle {
    pub bold: bool,
    pub italic: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub underline: bool,
    /// A line drawn THROUGH the writing, which a form uses to strike an entry
    /// out. 154 of the fonts in the 285 conformance workbooks ask for one.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub strikethrough: bool,
    pub font_size: Option<f32>,
    /// The typeface a cell asks for, when it names one.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_name: Option<String>,
    /// The charset the font record states beside the face, and the family.
    /// A workbook that asks for a face this machine has not got is answered
    /// from these two and not from the name, so both have to survive as far
    /// as the renderer (`cell_face_in_place`).
    #[serde(default)]
    pub font_charset: Option<i32>,
    #[serde(default)]
    pub font_family: Option<i32>,
    pub font_color: Option<String>,
    pub bg_color: Option<String>,
    pub number_format: Option<String>,
    pub horizontal_align: Option<String>,
    /// How many levels the text is pushed in from the edge it is aligned to.
    /// 71 of the conformance corpus's 285 workbooks use one.
    #[serde(default, skip_serializing_if = "is_zero")]
    pub indent: u32,
    /// Where the text sits within the cell's height: "top", "center", "bottom".
    /// Excel leaves a cell at the bottom when it says nothing.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_align: Option<String>,
    /// True when the text breaks onto further lines rather than running past
    /// the cell's right edge.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub wrap_text: bool,
    /// True when the text is made smaller until it fits the cell's width
    /// rather than running on or being clipped 窶・85 of the conformance
    /// corpus's 285 workbooks ask for it.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub shrink_to_fit: bool,
    /// True when the cell's characters are stacked one above the next rather
    /// than set in a line 窶・`textRotation="255"`, which is how a Japanese
    /// form labels a narrow column. 771 of the 774 rotations in the
    /// conformance corpus are this one.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub stacked_text: bool,
    /// True when an apostrophe typed first told Excel to keep the entry as
    /// text -- `quotePrefix="1"` on the cell's format, which
    /// `Range.PrefixCharacter` answers as `'`.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub quote_prefix: bool,
    pub border_top: Option<BorderLine>,
    pub border_bottom: Option<BorderLine>,
    pub border_left: Option<BorderLine>,
    pub border_right: Option<BorderLine>,
    /// The rule drawn corner to corner, which a Japanese form uses to strike
    /// a cell out. The border says how it is drawn; the two flags say which
    /// way, and a cell can carry both.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_diagonal: Option<BorderLine>,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub diagonal_up: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub diagonal_down: bool,
    /// Whether the cell may still be typed into once its sheet is protected.
    ///
    /// Excel locks every cell by default and a form unlocks the few it wants
    /// filled in, so the negative is what is held here: a style built from its
    /// own default is a locked one, which is what Excel means by silence.
    /// It decides nothing on an unprotected sheet.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub unlocked: bool,
}
