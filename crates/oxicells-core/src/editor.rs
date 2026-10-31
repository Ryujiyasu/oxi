// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Round-trip xlsx editor.
//!
//! Preserves the original ZIP archive. Patches cell values in worksheet XML,
//! inserting missing cells and rows when necessary. Cells edited to string
//! values are written as inline strings (t="str") to avoid rewriting the
//! shared-string table.

use std::collections::{BTreeMap, HashMap};
use std::io::{Cursor, Read, Write};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::ir::{AutoFilter, BorderLine, CalculationMode, CalculationSettings, CellStyle, ConditionalRule, DataValidation, DiffStyle, Hyperlink, MergeCell, Workbook};
use crate::parser::{parse_xlsx, XlsxError};
use oxicells_calc::rename_sheet_in_formula;
use oxidocs_common::archive::OoxmlArchive;
use oxidocs_common::relationships::parse_relationships;
use oxidocs_common::xml_utils::{get_attr, local_name};

/// A cell edit operation.
#[derive(Debug, Clone)]
pub struct CellEdit {
    /// 0-based sheet index
    pub sheet_index: usize,
    /// 1-based row number (as in OOXML)
    pub row: u32,
    /// 0-based column index
    pub col: u32,
    /// New display value (written as inline string)
    pub new_value: String,
}

/// The OOXML value type to write for a cell edit.
#[derive(Debug, Clone)]
pub enum CellEditValue {
    String(String),
    Utf16(Vec<u16>),
    RichText { value: crate::ir::CellValue, runs: Vec<crate::ir::TextRun> },
    Number(f64),
    Boolean(bool),
    Error(String),
    Formula(String),
    /// An ARRAY formula, written on the top-left cell of `reach` (an A1
    /// range such as `C1:C2`); the other cells of the block hold values only.
    ArrayFormula { formula: String, reach: String },
    /// A formula and its current IR value, including an intentionally stale
    /// value left by a macro running with manual calculation.
    CachedFormula {
        formula: String,
        reach: Option<String>,
        value: crate::ir::CellValue,
    },
    CachedUtf16Formula {
        formula: Vec<u16>,
        reach: Option<String>,
        value: crate::ir::CellValue,
    },
    Empty,
}

/// Round-trip xlsx editor.
pub struct XlsxEditor {
    original_data: Vec<u8>,
    workbook: Workbook,
    edits: HashMap<(usize, u32, u32), CellEditValue>, // (sheet_idx, row, col) -> value
    /// (sheet_idx, 1-based row) -> hidden
    row_hidden: HashMap<(usize, u32), bool>,
    /// (sheet_idx, 1-based row) -> its height in points, or `None` to put the
    /// row back on the sheet's automatic one, and whether it was chosen.
    row_height: HashMap<(usize, u32), (Option<f32>, bool)>,
    /// (sheet_idx, 0-based column) -> hidden
    col_hidden: HashMap<(usize, u32), bool>,
    /// sheet_idx -> how many rows to hold at the top and columns at the left.
    /// Both zero unfreezes: freezing nothing and holding nothing are the same
    /// thing, and Excel writes the same file for either.
    panes: HashMap<usize, (u32, u32)>,
    /// (sheet_idx, 0-based column) -> width, in characters of the workbook's
    /// default font, gutter included. This is the number the file states, not
    /// the one a person types into Excel: typing 10 stores 10.625.
    col_width: HashMap<(usize, u32), f32>,
    col_custom_width: HashMap<(usize, u32), bool>,
    default_col_width: HashMap<usize, f32>,
    /// Sheets whose merged cells are being replaced wholesale. A merge is not
    /// edited in place, so the whole list travels together.
    merges: HashMap<usize, Vec<MergeCell>>,
    /// (sheet_idx, 1-based row, 0-based column) -> the style it should carry.
    styles: HashMap<(usize, u32, u32), CellStyle>,
    /// Whole-row and whole-column styles, including styles of blank cells.
    row_band_styles: HashMap<(usize, u32), Option<CellStyle>>,
    column_band_styles: HashMap<(usize, u32), Option<CellStyle>>,
    /// Sheets whose filter is being replaced. `None` takes the filter away.
    filters: HashMap<usize, Option<AutoFilter>>,
    /// Complete validation rules for sheets whose rules changed.
    validations: HashMap<usize, Vec<DataValidation>>,
    /// Conditional rules whose original worksheet XML must be reconciled.
    conditional_rules: HashMap<usize, Vec<ConditionalRule>>,
    /// Whole-sheet hyperlink sets changed by the caller, keyed by original
    /// worksheet index so additions and deletions are written together.
    hyperlinks: HashMap<usize, Vec<Hyperlink>>,
    /// Sheets whose tab colour is being set. `Some(hex)` paints it, `None`
    /// takes any colour away. Keyed by the sheet's origin index.
    tab_colors: HashMap<usize, Option<String>>,
    /// What the workbook's sheets should end up as, when they are changing.
    sheet_plan: Option<Vec<PlannedSheet>>,
    /// The names the workbook should end up keeping, when they have changed.
    /// Held whole, because a name is not edited in place: the list travels
    /// together so that one taken away really goes.
    name_plan: Option<Vec<(String, String)>>,
    name_attributes_plan: Option<crate::ir::DefinedNameAttributes>,
    calculation_plan: Option<CalculationSettings>,
    /// Replacement of the workbook's legacy 56-color palette, stored in the
    /// stylesheet's `indexedColors` list.
    indexed_colors: Option<Vec<String>>,
    /// `styles.xml` and its package relationships are absent in some small
    /// workbooks; a palette edit creates that part when needed.
    chart_edits: HashMap<u32,(crate::ir::Chart,crate::ir::Chart)>,
    /// Existing drawing shapes whose line paint changed, keyed by original sheet and shape name.
    drawing_line_edits: HashMap<(usize, String), Option<crate::ir::ShapeLine>>,
    drawing_fill_edits: HashMap<(usize, String), (Option<String>, Option<usize>, Option<u8>, Option<f32>)>,
    drawing_text_edits: HashMap<(usize, String), ShapeTextEdit>,
    /// PictureFormat adjustments for existing worksheet pictures.
    drawing_picture_edits: HashMap<(usize, String), PictureStyleEdit>,
    /// Existing drawings whose cell anchors or one-cell extent changed.
    drawing_geometry_edits: HashMap<(usize, String), crate::ir::Drawing>,
    /// Existing drawings in the changed front-to-back order for each sheet.
    drawing_order_edits: HashMap<usize, Vec<String>>,
    /// Current shape geometry for package edits that flatten nested groups.
    drawing_states: HashMap<(usize, String), crate::ir::Drawing>,
    /// Existing preset shapes removed by a VBA Shapes.Delete call.
    drawing_deletions: HashMap<usize, std::collections::HashSet<String>>,
    /// Shapes a run added to sheets that already existed in the source package.
    drawing_additions: HashMap<usize, Vec<crate::ir::Drawing>>,
    /// Existing top-level drawings joined into a newly created group.
    drawing_group_additions: HashMap<usize, Vec<DrawingGroupAddition>>,
    /// (original sheet index, threaded-comment id) -> replacement text.
    threaded_comment_text: HashMap<(usize, String), (String, Option<Vec<u16>>)>,
    /// New threaded comments, keyed by their original sheet index.
    threaded_comment_additions: HashMap<usize, Vec<crate::ir::ThreadedComment>>,
    /// IDs removed from each original sheet.
    threaded_comment_deletions: HashMap<usize, std::collections::HashSet<String>>,
    /// Resolved state changes by original sheet and threaded-comment id.
    threaded_comment_resolved: HashMap<(usize, String), bool>,
    /// Replacement legacy notes for original sheets whose comments changed.
    legacy_comments: HashMap<usize, Vec<crate::ir::Comment>>,
    /// Final sheet-order table definitions, when VBA changed a table.
    table_plan: Option<Vec<Option<Vec<crate::ir::Table>>>>,
    /// Final sheet-order dynamic-array owners, including one-cell and blocked
    /// arrays. Their extents are independent of adjacent spill members.
    dynamic_plan: Option<Vec<Vec<crate::ir::CachedSpill>>>,
}

/// One sheet of a rearranged workbook: either one the file already holds, or a
/// new one to be written from scratch.
enum PlannedSheet {
    /// The sheet at this position in the original file, under this name, and
    /// shown or hidden as the edited workbook says.
    Held {
        origin: usize,
        name: String,
        visibility: crate::ir::Visibility,
    },
    /// A sheet the file has never held.
    Added(crate::ir::Sheet),
}

/// Convert 0-based column to letter reference (0->A, 25->Z, 26->AA).
pub fn col_to_letter(mut col: u32) -> String {
    let mut result = String::new();
    loop {
        result.insert(0, (b'A' + (col % 26) as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    result
}

impl XlsxEditor {
    pub fn new(data: &[u8]) -> Result<Self, XlsxError> {
        let workbook = parse_xlsx(data)?;
        Ok(Self {
            original_data: data.to_vec(),
            dynamic_plan: None,
            workbook,
            edits: HashMap::new(),
            row_hidden: HashMap::new(),
            row_height: HashMap::new(),
            col_hidden: HashMap::new(),
            col_width: HashMap::new(),
            col_custom_width: HashMap::new(),
            default_col_width: HashMap::new(),
            panes: HashMap::new(),
            merges: HashMap::new(),
            styles: HashMap::new(),
            row_band_styles: HashMap::new(),
            column_band_styles: HashMap::new(),
            filters: HashMap::new(),
            validations: HashMap::new(),
            conditional_rules: HashMap::new(),
            hyperlinks: HashMap::new(),
            tab_colors: HashMap::new(),
            sheet_plan: None,
            name_plan: None,
            name_attributes_plan: None,
            calculation_plan: None,
            indexed_colors: None,
            chart_edits: HashMap::new(),
            drawing_line_edits: HashMap::new(),
            drawing_fill_edits: HashMap::new(),
            drawing_text_edits: HashMap::new(),
            drawing_picture_edits: HashMap::new(),
            drawing_geometry_edits: HashMap::new(),
            drawing_order_edits: HashMap::new(),
            drawing_states: HashMap::new(),
            drawing_deletions: HashMap::new(),
            drawing_additions: HashMap::new(),
            drawing_group_additions: HashMap::new(),
            threaded_comment_text: HashMap::new(),
            threaded_comment_additions: HashMap::new(),
            threaded_comment_deletions: HashMap::new(),
            threaded_comment_resolved: HashMap::new(),
            legacy_comments: HashMap::new(),
            table_plan: None,
        })
    }

    pub fn workbook(&self) -> &Workbook {
        &self.workbook
    }

    pub fn set_cell(&mut self, sheet_index: usize, row: u32, col: u32, value: String) {
        self.set_cell_value(sheet_index, row, col, CellEditValue::String(value));
    }

    pub fn set_cell_value(
        &mut self,
        sheet_index: usize,
        row: u32,
        col: u32,
        value: CellEditValue,
    ) {
        self.edits.insert((sheet_index, row, col), value);
    }

    pub fn apply_edits(&mut self, edits: &[CellEdit]) {
        for e in edits {
            self.set_cell(e.sheet_index, e.row, e.col, e.new_value.clone());
        }
    }

    /// Works out what changed between the workbook this editor opened and the
    /// one it is handed, and records those changes as edits.
    ///
    /// This is how a VBA run reaches the file: the runtime rewrites the IR in
    /// place, and the difference against the original is what has to be
    /// written back.
    ///
    /// Only what the editor can write is compared: cell values and formulas,
    /// which rows and columns are hidden, which cells are merged, the styling
    /// each cell wears, the filter the sheet is under, and which sheets the
    /// workbook holds.
    ///
    /// Anything the IR does not model  Econditional formatting, data
    /// validation, charts, drawings  Erides along in the original XML
    /// untouched, except for line paint changes on existing named shapes.
    pub fn apply_workbook(&mut self, edited: &Workbook) -> Result<(), XlsxError> {
        if edited.sheets.is_empty() {
            return Err(XlsxError::InvalidData(
                "a workbook must keep at least one sheet".to_string(),
            ));
        }
        self.dynamic_plan = Some(dynamic_array_plan(edited));
        self.indexed_colors = (edited.indexed_colors != self.workbook.indexed_colors)
            .then(|| edited.indexed_colors.clone());
        let originals:HashMap<u32,&crate::ir::Chart>=self.workbook.sheets.iter().flat_map(|sheet|&sheet.drawings)
            .filter_map(|drawing|match &drawing.kind {crate::ir::DrawingKind::Chart(chart)=>chart.origin_id.map(|id|(id,chart)),_=>None}).collect();
        self.chart_edits.clear();
        self.drawing_line_edits.clear();
        self.drawing_fill_edits.clear();
        self.drawing_text_edits.clear();
        self.drawing_picture_edits.clear();
        self.drawing_geometry_edits.clear();
        self.drawing_order_edits.clear();
        self.drawing_states.clear();
        self.drawing_deletions.clear();
        self.drawing_additions.clear();
        self.drawing_group_additions.clear();
        for chart in edited.sheets.iter().flat_map(|sheet|&sheet.drawings).filter_map(|drawing|match &drawing.kind {crate::ir::DrawingKind::Chart(chart)=>Some(chart),_=>None}) {
            if let Some(origin)=chart.origin_id {
                let before=originals.get(&origin).ok_or_else(||XlsxError::InvalidData("chart source identity is not in this workbook".into()))?;
                if before.title!=chart.title || crate::chart_edit::legend_position(before)!=crate::chart_edit::legend_position(chart) {
                    if self.chart_edits.get(&origin).is_some_and(|(_,prior)|prior.title!=chart.title || crate::chart_edit::legend_position(prior)!=crate::chart_edit::legend_position(chart)) {
                        return Err(XlsxError::InvalidData("conflicting edits for a shared chart source".into()));
                    }
                    self.chart_edits.insert(origin,((*before).clone(),chart.clone()));
                }
            }
        }
        for (edited_index, edited_sheet) in edited.sheets.iter().enumerate() {
            let original_index = edited_sheet.origin_id
                .and_then(|id| self.workbook.sheets.iter().position(|sheet| sheet.origin_id == Some(id)))
                .unwrap_or(edited_index);
            let Some(original_sheet) = self.workbook.sheets.get(original_index) else { continue };
            let order = |drawings: &[crate::ir::Drawing]| drawings.iter().filter_map(|drawing| drawing.name.clone()).collect::<Vec<_>>();
            let before_order = order(&original_sheet.drawings);
            let after_order = order(&edited_sheet.drawings);
            if before_order != after_order {
                self.drawing_order_edits.insert(original_index, after_order);
            }
            let mut new_groups = Vec::new();
            for drawing in edited_sheet.drawings.iter().filter(|drawing| drawing.grouped) {
                let Some(group_name) = drawing.group_name.as_ref() else { continue };
                if new_groups.iter().any(|group: &DrawingGroupAddition| &group.name == group_name) { continue; }
                let from = drawing.from;
                let Some(to) = drawing.to else { continue };
                let Some(transform) = drawing.group_transform else { continue };
                let members = edited_sheet.drawings.iter().filter_map(|member| {
                    if !member.grouped || member.group_name.as_ref() != Some(group_name) { return None; }
                    let name = member.name.as_ref()?;
                    let was_ungrouped = original_sheet.drawings.iter()
                        .find(|original| original.name.as_ref() == Some(name))
                        .is_some_and(|original| !original.grouped);
                    was_ungrouped.then_some((name.clone(), member.group_child_transform?))
                }).collect::<Vec<_>>();
                if members.len() > 1 {
                    new_groups.push(DrawingGroupAddition { name: group_name.clone(), from, to, transform, members });
                }
            }
            if !new_groups.is_empty() {
                self.drawing_group_additions.insert(original_index, new_groups);
            }
            for original in &original_sheet.drawings {
                let Some(name) = original.name.as_ref() else { continue };
                if !edited_sheet.drawings.iter().any(|drawing| drawing.name.as_ref() == Some(name)) {
                    self.drawing_deletions.entry(original_index).or_default().insert(name.clone());
                }
            }
            for drawing in &edited_sheet.drawings {
                let Some(name) = drawing.name.as_ref() else { continue };
                if drawing.grouped
                    && self.drawing_deletions.get(&original_index).is_some_and(|names| !names.is_empty())
                {
                    self.drawing_states.insert((original_index, name.clone()), drawing.clone());
                }
                let existed = original_sheet.drawings.iter().any(|original| original.name.as_ref() == Some(name));
                let is_new_ole = drawing.ole_prog_id.is_some() && drawing.ole_embedded_package.is_some()
                    && drawing.ole_source_name.is_none();
                if !existed && (matches!(&drawing.kind, crate::ir::DrawingKind::Shape(_)) || drawing.ole_source_name.is_some() || is_new_ole) {
                    self.drawing_additions.entry(original_index).or_default().push(drawing.clone());
                    continue;
                }
                let matches = original_sheet.drawings.iter().filter(|original| original.name.as_ref() == Some(name)).collect::<Vec<_>>();
                if matches.len() > 1 {
                    return Err(XlsxError::InvalidData(format!("drawing name {name:?} is ambiguous on sheet {}", original_sheet.name)));
                }
                let Some(original) = matches.first() else { continue };
                if (original.grouped && !drawing.grouped)
                    || ((!original.grouped || !drawing.grouped) && drawing_position_changed(original, drawing))
                {
                    let mut geometry = drawing.clone();
                    if original.to.is_none() { geometry.to = None; }
                    self.drawing_geometry_edits.insert((original_index, name.clone()), geometry);
                }
                if let (crate::ir::DrawingKind::Shape(before_shape), crate::ir::DrawingKind::Shape(after_shape)) = (&original.kind, &drawing.kind) {
                    if let (Some(before_text), Some(after_text)) = (before_shape.text.as_ref(), after_shape.text.as_ref()) {
                        if shape_text_formatting_changed(before_text, after_text) {
                            self.drawing_text_edits.insert((original_index, name.clone()), ShapeTextEdit {
                                before: before_text.clone(), after: after_text.clone(),
                            });
                        }
                    }
                    if before_shape.fill != after_shape.fill
                        || before_shape.fill_theme != after_shape.fill_theme
                        || before_shape.fill_legacy_color_index != after_shape.fill_legacy_color_index
                        || before_shape.fill_brightness != after_shape.fill_brightness
                    {
                        self.drawing_fill_edits.insert((original_index, name.clone()), (
                            after_shape.fill.clone(), after_shape.fill_theme, after_shape.fill_legacy_color_index, after_shape.fill_brightness,
                        ));
                    }
                    if before_shape.line != after_shape.line {
                        self.drawing_line_edits.insert((original_index, name.clone()), after_shape.line.clone());
                    }
                }
                if let (crate::ir::DrawingKind::Picture { format: before, .. }, crate::ir::DrawingKind::Picture { format: after, .. }) = (&original.kind, &drawing.kind) {
                    if before != after {
                        self.drawing_picture_edits.insert((original_index, name.clone()), PictureStyleEdit { before: before.clone(), after: after.clone() });
                    }
                }
            }
        }

        // Sheets are matched by position when there are as many as before, so a
        // rename is a rename. Otherwise they are matched by name, and whatever
        // is left over has been added or taken away.
        // A name is a reference the workbook keeps; it changes on its own,
        // without any cell being touched.
        let mut edited_calculation = edited.calculation;
        let mut original_calculation = self.workbook.calculation;
        // Calculation language belongs to the host, not to the OOXML package.
        edited_calculation.dbcs_language = None;
        original_calculation.dbcs_language = None;
        if edited_calculation != original_calculation {
            self.calculation_plan = Some(edited_calculation);
        }

        let same_count = edited.sheets.len() == self.workbook.sheets.len();
        let mut origins: Vec<Option<usize>> = Vec::with_capacity(edited.sheets.len());
        if self.workbook.sheets.iter().all(|held| held.origin_id.is_some()) {
            // The parser stamped every original sheet with a stable id, so the
            // match is by that id: a renamed sheet keeps it, an added sheet has
            // none, and a deleted sheet's id is simply absent from the edited
            // list. This tells a rename apart from an add-and-delete even when
            // the two leave the sheet count unchanged, which position matching
            // could not.
            for sheet in &edited.sheets {
                origins.push(match sheet.origin_id {
                    Some(id) => self
                        .workbook
                        .sheets
                        .iter()
                        .position(|held| held.origin_id == Some(id)),
                    None => None,
                });
            }
        } else {
            // A workbook not built by the parser has no ids; fall back to the
            // old rule -- position when the count is unchanged, name otherwise.
            for (index, sheet) in edited.sheets.iter().enumerate() {
                origins.push(if same_count {
                    Some(index)
                } else {
                    self.workbook
                        .sheets
                        .iter()
                        .position(|held| held.name.eq_ignore_ascii_case(&sheet.name))
                });
            }
        }

        // A rename is a sheet matched to an origin whose name has changed.
        // The formulas the user did not touch still name the old sheet, so
        // every reference to it across the workbook is rewritten to the new
        // name; the rewritten cells then fall out as ordinary edits below.
        let renames: Vec<(String, String)> = edited
            .sheets
            .iter()
            .zip(&origins)
            .filter_map(|(after, origin)| {
                let before = self.workbook.sheets.get((*origin)?)?;
                (before.name != after.name).then(|| (before.name.clone(), after.name.clone()))
            })
            .collect();
        let deleted: Vec<&str> = self.workbook.sheets.iter().enumerate()
            .filter(|(index, _)| !origins.contains(&Some(*index))).map(|(_, s)| s.name.as_str()).collect();
        let rewritten: Workbook;
        let edited: &Workbook = if renames.is_empty() && deleted.is_empty() { edited } else {
            let mut copy = edited.clone();
            for sheet in &mut copy.sheets {
                for row in &mut sheet.rows {
                    for cell in &mut row.cells {
                        for (old, new) in &renames {
                            let _ = cell.rewrite_formula(|formula| Ok(rename_sheet_in_formula(formula, old, new)));
                        }
                    }
                }
            }
            copy.defined_names.retain(|(name, _)| !crate::ir::defined_name_scope(name)
                .is_some_and(|(scope, _)| deleted.iter().any(|gone| gone.eq_ignore_ascii_case(&scope))));
            for (name, refers_to) in &mut copy.defined_names {
                for (old, new) in &renames {
                    if let Some((scope, leaf)) = crate::ir::defined_name_scope(name) {
                        if scope.eq_ignore_ascii_case(old) {
                            let renamed = crate::ir::scoped_defined_name(new, leaf);
                            if let Some(attributes) = copy.defined_name_attributes.remove(name) { copy.defined_name_attributes.insert(renamed.clone(), attributes); }
                            *name = renamed;
                        }
                    }
                    *refers_to = rename_sheet_in_formula(refers_to, old, new);
                }
            }
            copy.defined_name_attributes.retain(|name, _| copy.defined_names.iter().any(|(held, _)| held == name));
            rewritten = copy;
            &rewritten
        };

        if edited.defined_names != self.workbook.defined_names
            || edited.defined_name_attributes != self.workbook.defined_name_attributes
        {
            self.name_plan = Some(edited.defined_names.clone());
            self.name_attributes_plan = Some(edited.defined_name_attributes.clone());
        }

        let rearranged = !same_count
            || origins.iter().enumerate().any(|(index, origin)| {
                *origin != Some(index)
            })
            || edited
                .sheets
                .iter()
                .zip(&self.workbook.sheets)
                .any(|(after, before)| {
                    after.name != before.name || after.visibility != before.visibility
                });
        if rearranged {
            // A localSheetId is an index in the current sheet list, not a
            // persistent sheet id. Reordering sheets rewrites local names too.
            self.name_plan = Some(edited.defined_names.clone());
            self.name_attributes_plan = Some(edited.defined_name_attributes.clone());
            self.sheet_plan = Some(
                edited
                    .sheets
                    .iter()
                    .zip(&origins)
                    .map(|(sheet, origin)| match origin {
                        Some(origin) => PlannedSheet::Held {
                            origin: *origin,
                            name: sheet.name.clone(),
                            visibility: sheet.visibility,
                        },
                        None => PlannedSheet::Added(sheet.clone()),
                    })
                    .collect(),
            );
        }

        let pairs: Vec<(usize, (&crate::ir::Sheet, &crate::ir::Sheet))> = edited
            .sheets
            .iter()
            .zip(&origins)
            .filter_map(|(after, origin)| {
                origin.map(|origin| (origin, (&self.workbook.sheets[origin], after)))
            })
            .collect();

        if self.table_plan.is_some() || edited.sheets.iter().zip(&origins).any(|(after, origin)| {
            origin.map(|index| self.workbook.sheets[index].tables != after.tables).unwrap_or(!after.tables.is_empty())
        }) {
            self.table_plan = Some(edited.sheets.iter().zip(&origins).map(|(sheet, origin)| {
                let changed = origin.map(|index| self.workbook.sheets[index].tables != sheet.tables).unwrap_or(!sheet.tables.is_empty());
                changed.then(|| sheet.tables.clone())
            }).collect());
        }

        for (index, (before, after)) in pairs {
            if before.hyperlinks != after.hyperlinks {
                self.hyperlinks.insert(index, after.hyperlinks.clone());
            }
            let mut after_comments = after.comments.clone();
            for compatibility in &before.comments {
                if compatibility.author.strip_prefix("tc=").is_some_and(|id| after.threaded_comments.iter().any(|thread| thread.cell == compatibility.cell && thread.id.eq_ignore_ascii_case(id)))
                    && !after_comments.iter().any(|comment| comment.cell == compatibility.cell)
                {
                    after_comments.push(compatibility.clone());
                }
            }
            after_comments.sort_by_key(|comment| comment.cell);
            if !legacy_comments_equivalent(&before.comments, &after_comments) {
                self.legacy_comments.insert(index, after_comments);
            }
            let held_ids: std::collections::HashSet<&str> = before.threaded_comments.iter()
                .map(|comment| comment.id.as_str()).collect();
            let after_ids: std::collections::HashSet<&str> = after.threaded_comments.iter()
                .map(|comment| comment.id.as_str()).collect();
            for comment in &before.threaded_comments {
                if !after_ids.contains(comment.id.as_str()) {
                    self.threaded_comment_deletions.entry(index).or_default().insert(comment.id.clone());
                }
            }
            for comment in &after.threaded_comments {
                if held_ids.contains(comment.id.as_str()) {
                    let original = before.threaded_comments.iter().find(|held| held.id == comment.id).unwrap();
                    if original.text != comment.text || original.utf16_text != comment.utf16_text {
                        self.threaded_comment_text.insert((index, comment.id.clone()), (comment.text.clone(), comment.utf16_text.clone()));
                    }
                    if original.resolved != comment.resolved {
                        self.threaded_comment_resolved.insert((index, comment.id.clone()), comment.resolved);
                    }
                } else {
                    let mut addition = comment.clone();
                    let mut ids: std::collections::HashSet<String> = self.workbook.sheets.iter()
                        .flat_map(|sheet| sheet.threaded_comments.iter())
                        .map(|held| held.id.clone()).collect();
                    if addition.id.is_empty() || ids.contains(&addition.id) {
                        addition.id = unique_threaded_guid(
                            &format!("{}:{}:{}:{}:{}", index, addition.cell.0, addition.cell.1, addition.text, self.original_data.len()),
                            &ids,
                        );
                    }
                    ids.insert(addition.id.clone());
                    if addition.person_id.is_empty() {
                        let people: std::collections::HashSet<String> = self.workbook.sheets.iter()
                            .flat_map(|sheet| sheet.threaded_comments.iter())
                            .map(|held| held.person_id.clone()).collect();
                        addition.person_id = unique_threaded_guid("Oxi threaded-comment author User", &people);
                    }
                    self.threaded_comment_additions.entry(index).or_default().push(addition);
                }
            }
            if before.tab_color != after.tab_color {
                self.tab_colors.insert(index, after.tab_color.clone());
            }
            let mut changes: Vec<((usize, u32, u32), CellEditValue)> = Vec::new();
            let held = cells_of(before);
            let now = cells_of(after);
            let memo = std::cell::RefCell::new(BTreeMap::new());
            let multiple = |name: &str| cached_named_array_shape(&edited.defined_names, &after.name, name, &memo);
            for (&(row, col), cell) in &now {
                let same = held
                    .get(&(row, col))
                    .is_some_and(|before| same_content(before, cell));
                if !same {
                    if let Some(value) = edit_for(cell, row, &multiple) {
                        changes.push(((index, row, col), value));
                    }
                }
            }
            // A cell the run emptied has to be written as empty, not left alone.
            for &(row, col) in held.keys() {
                if !now.contains_key(&(row, col)) {
                    changes.push(((index, row, col), CellEditValue::Empty));
                }
            }
            for (key, value) in changes {
                self.edits.insert(key, value);
            }

            let mut rows: Vec<(u32, bool)> = Vec::new();
            for row in &after.rows {
                let was = before
                    .rows
                    .iter()
                    .find(|held| held.index == row.index)
                    .is_some_and(|held| held.hidden);
                if was != row.hidden {
                    rows.push((row.index, row.hidden));
                }
            }
            for row in &before.rows {
                let gone = !after.rows.iter().any(|held| held.index == row.index);
                if gone && row.hidden {
                    rows.push((row.index, false));
                }
            }
            for (row, hidden) in rows {
                self.row_hidden.insert((index, row), hidden);
            }

            for row in &after.rows {
                let was = before.rows.iter().find(|held| held.index == row.index);
                let had = was.and_then(|held| held.height);
                let pinned = was.is_some_and(|held| held.custom_height);
                if had != row.height || pinned != row.custom_height {
                    self.row_height
                        .insert((index, row.index), (row.height, row.custom_height));
                }
            }

            let mut cols: Vec<(u32, bool)> = Vec::new();
            for col in &after.hidden_cols {
                if !before.hidden_cols.contains(col) {
                    cols.push((*col, true));
                }
            }
            for col in &before.hidden_cols {
                if !after.hidden_cols.contains(col) {
                    cols.push((*col, false));
                }
            }
            for (col, hidden) in cols {
                self.col_hidden.insert((index, col), hidden);
            }

            for (col, width) in after.col_widths.iter().enumerate() {
                let was = before.col_widths.get(col).copied().unwrap_or(0.0);
                let custom = after.column_has_custom_width(col as u32);
                if *width != was || custom != before.column_has_custom_width(col as u32) {
                    self.col_width.insert((index, col as u32), *width);
                    self.col_custom_width.insert((index, col as u32), custom);
                }
            }
            if before.default_col_width != after.default_col_width || before.default_col_width_explicit != after.default_col_width_explicit {
                self.default_col_width.insert(index, after.default_col_width);
            }

            if (before.frozen_rows, before.frozen_cols)
                != (after.frozen_rows, after.frozen_cols)
            {
                self.panes
                    .insert(index, (after.frozen_rows, after.frozen_cols));
            }

            if !same_merges(&before.merge_cells, &after.merge_cells) {
                self.merges.insert(index, after.merge_cells.clone());
            }

            if before.auto_filter != after.auto_filter {
                self.filters.insert(index, after.auto_filter.clone());
            }
            if before.validations != after.validations {
                self.validations.insert(index, after.validations.clone());
            }
            if before.conditional_rules != after.conditional_rules {
                self.conditional_rules.insert(index, after.conditional_rules.clone());
            }
            for row in before.band_styles.rows.keys().chain(after.band_styles.rows.keys()).copied().collect::<std::collections::BTreeSet<_>>() {
                let was = before.band_styles.rows.get(&row);
                let now = after.band_styles.rows.get(&row);
                if was != now { self.row_band_styles.insert((index, row), now.cloned()); }
            }
            let columns: std::collections::BTreeSet<u32> = if before.band_styles.sheet != after.band_styles.sheet {
                (0..16_384).collect()
            } else {
                before.band_styles.columns.keys().chain(after.band_styles.columns.keys()).copied().collect()
            };
            for column in columns {
                let was = before.band_styles.columns.get(&column).or(before.band_styles.sheet.as_ref());
                let now = after.band_styles.columns.get(&column).or(after.band_styles.sheet.as_ref());
                if was != now { self.column_band_styles.insert((index, column), now.cloned()); }
            }

            for (&(row, col), cell) in &now {
                let was = held.get(&(row, col)).map(|before| &before.style);
                if was != Some(&cell.style) {
                    self.styles
                        .insert((index, row, col), cell.style.clone());
                }
            }
        }
        Ok(())
    }

    /// Hide or reveal a whole row. `row` is one-based, as OOXML counts them.
    pub fn set_row_hidden(&mut self, sheet_index: usize, row: u32, hidden: bool) {
        self.row_hidden.insert((sheet_index, row), hidden);
    }

    /// Sets how tall a row is, in points.
    ///
    /// `None` puts the row back on the height Excel works out from what is in
    /// it. `pinned` says the height was chosen rather than computed, which is
    /// what stops Excel working it out again the next time the file is opened.
    pub fn set_row_height(
        &mut self,
        sheet_index: usize,
        row: u32,
        height: Option<f32>,
        pinned: bool,
    ) {
        self.row_height.insert((sheet_index, row), (height, pinned));
    }

    /// Hide or reveal a whole column. `col` is zero-based, as the IR counts it.
    pub fn set_col_hidden(&mut self, sheet_index: usize, col: u32, hidden: bool) {
        self.col_hidden.insert((sheet_index, col), hidden);
    }

    /// Sets how wide a column is.
    ///
    /// The width is in characters of the workbook's default font with the
    /// gutter either side counted in, which is what the file states and not
    /// what Excel's own box shows: typing 10 there stores 10.625.
    pub fn set_col_width(&mut self, sheet_index: usize, col: u32, width: f32) {
        self.col_width.insert((sheet_index, col), width);
    }

    /// Replaces every merged cell on a sheet.
    pub fn set_merges(&mut self, sheet_index: usize, merges: Vec<MergeCell>) {
        self.merges.insert(sheet_index, merges);
    }

    /// Gives a cell a style of its own.
    pub fn set_cell_style(
        &mut self,
        sheet_index: usize,
        row: u32,
        col: u32,
        style: CellStyle,
    ) {
        self.styles.insert((sheet_index, row, col), style);
    }

    /// Holds `rows` at the top of the sheet and `cols` at the left while the
    /// rest scrolls. Both zero takes the freeze away.
    pub fn set_frozen_panes(&mut self, sheet_index: usize, rows: u32, cols: u32) {
        self.panes.insert(sheet_index, (rows, cols));
    }

    /// Puts a sheet under a filter, or takes its filter away with `None`.
    pub fn set_auto_filter(&mut self, sheet_index: usize, filter: Option<AutoFilter>) {
        self.filters.insert(sheet_index, filter);
    }

    pub fn has_edits(&self) -> bool {
        self.has_non_chart_edits() || !self.chart_edits.is_empty()
    }

    fn has_non_chart_edits(&self) -> bool {
        !self.edits.is_empty()
            || !self.row_hidden.is_empty()
            || !self.row_height.is_empty()
            || !self.col_hidden.is_empty()
            || !self.col_width.is_empty()
            || !self.default_col_width.is_empty()
            || !self.panes.is_empty()
            || !self.merges.is_empty()
            || !self.styles.is_empty()
            || !self.row_band_styles.is_empty()
            || !self.column_band_styles.is_empty()
            || !self.filters.is_empty()
            || !self.validations.is_empty()
            || !self.conditional_rules.is_empty()
            || !self.hyperlinks.is_empty()
            || !self.tab_colors.is_empty()
            || self.sheet_plan.is_some()
            || self.name_plan.is_some()
            || self.calculation_plan.is_some()
            || self.indexed_colors.is_some()
            || !self.threaded_comment_text.is_empty()
            || !self.threaded_comment_additions.is_empty()
            || !self.threaded_comment_deletions.is_empty()
            || !self.threaded_comment_resolved.is_empty()
            || !self.legacy_comments.is_empty()
            || self.table_plan.is_some()
            || self.dynamic_plan.as_ref().is_some_and(|plan| *plan != dynamic_array_plan(&self.workbook))
            || !self.drawing_line_edits.is_empty()
            || !self.drawing_fill_edits.is_empty()
            || !self.drawing_text_edits.is_empty()
            || !self.drawing_picture_edits.is_empty()
            || !self.drawing_geometry_edits.is_empty()
            || !self.drawing_order_edits.is_empty()
            || !self.drawing_deletions.is_empty()
            || !self.drawing_additions.is_empty()
            || !self.drawing_group_additions.is_empty()
    }

    /// Save edited xlsx.
    pub fn save(&self) -> Result<Vec<u8>, XlsxError> {
        if !self.has_edits() {
            return Ok(self.original_data.clone());
        }

        // Determine which sheet files need patching
        let sheet_paths = self.resolve_sheet_paths()?;
        let mut source=OoxmlArchive::new(&self.original_data)?;
        let chart_parts=crate::chart_edit::chart_part_names(&mut source)?;
        let drawing_parts = resolve_drawing_parts(&mut source, &sheet_paths)?;
        let drawing_package = drawing_addition_package(
            &mut source,
            &sheet_paths,
            &drawing_parts,
            &self.drawing_additions,
            &self.drawing_group_additions,
            &self.drawing_line_edits,
            &self.drawing_fill_edits,
            &self.drawing_text_edits,
            &self.drawing_picture_edits,
            &self.drawing_geometry_edits,
            &self.drawing_order_edits,
            &self.drawing_states,
            &self.drawing_deletions,
        )?;
        let threaded_edits = threaded_comment_package_edits(
            &mut source,
            &sheet_paths,
            &drawing_package,
            &self.threaded_comment_text,
            &self.threaded_comment_additions,
            &self.threaded_comment_deletions,
            &self.threaded_comment_resolved,
        )?;
        let mut legacy_comments = self.legacy_comments.clone();
        for (&sheet, additions) in &self.threaded_comment_additions {
            let Some(original) = self.workbook.sheets.get(sheet) else { continue };
            for thread in additions.iter().filter(|thread| thread.parent_id.is_none()) {
                let comments = legacy_comments.entry(sheet).or_insert_with(|| original.comments.clone());
                if let Some(existing) = comments.iter().find(|comment| comment.cell == thread.cell) {
                    if existing.author.eq_ignore_ascii_case(&format!("tc={}", thread.id)) { continue; }
                    return Err(XlsxError::InvalidData("threaded comment conflicts with an existing note".into()));
                }
                comments.push(threaded_compatibility_note(thread));
            }
        }
        for (&sheet, deleted) in &self.threaded_comment_deletions {
            let Some(original) = self.workbook.sheets.get(sheet) else { continue };
            let matching = |comment: &crate::ir::Comment| comment.author.strip_prefix("tc=")
                .is_some_and(|id| deleted.iter().any(|deleted| deleted.eq_ignore_ascii_case(id)));
            if legacy_comments.get(&sheet).unwrap_or(&original.comments).iter().any(matching) {
                legacy_comments.entry(sheet).or_insert_with(|| original.comments.clone()).retain(|comment| !matching(comment));
            }
        }
        // Excel also updates the legacy compatibility note when a thread or
        // one of its replies changes, so older readers see the current text.
        let mut touched = HashMap::<usize, std::collections::HashSet<(u32, u32)>>::new();
        for (&sheet, additions) in &self.threaded_comment_additions {
            touched.entry(sheet).or_default().extend(additions.iter().map(|comment| comment.cell));
        }
        for ((sheet, id), _) in &self.threaded_comment_text {
            if let Some(comment) = self.workbook.sheets.get(*sheet).and_then(|s| s.threaded_comments.iter().find(|comment| comment.id == *id)) {
                touched.entry(*sheet).or_default().insert(comment.cell);
            }
        }
        for (&sheet, deleted) in &self.threaded_comment_deletions {
            if let Some(original) = self.workbook.sheets.get(sheet) {
                touched.entry(sheet).or_default().extend(original.threaded_comments.iter().filter(|comment| deleted.contains(&comment.id)).map(|comment| comment.cell));
            }
        }
        for (sheet, cells) in touched {
            let Some(original) = self.workbook.sheets.get(sheet) else { continue };
            let deleted = self.threaded_comment_deletions.get(&sheet);
            let mut current: Vec<_> = original.threaded_comments.iter()
                .filter(|comment| !deleted.is_some_and(|ids| ids.contains(&comment.id))).cloned().collect();
            for comment in &mut current {
                if let Some((text, units)) = self.threaded_comment_text.get(&(sheet, comment.id.clone())) { comment.text = text.clone(); comment.utf16_text = units.clone(); }
            }
            if let Some(additions) = self.threaded_comment_additions.get(&sheet) { current.extend(additions.iter().cloned()); }
            for root in current.iter().filter(|comment| comment.parent_id.is_none() && cells.contains(&comment.cell)) {
                let comments = legacy_comments.entry(sheet).or_insert_with(|| original.comments.clone());
                let author = format!("tc={}", root.id);
                if !comments.iter().any(|comment| comment.cell == root.cell && comment.author.eq_ignore_ascii_case(&author)) {
                    if comments.iter().any(|comment| comment.cell == root.cell) { return Err(XlsxError::InvalidData("threaded comment conflicts with an existing note".into())); }
                    comments.push(threaded_compatibility_note(root));
                }
                let compatibility = comments.iter_mut().find(|comment| comment.cell == root.cell && comment.author.eq_ignore_ascii_case(&author)).unwrap();
                let mut replies: Vec<_> = current.iter().filter(|comment| comment.cell == root.cell && comment.parent_id.is_some()).collect();
                replies.sort_by(|a, b| a.date.cmp(&b.date));
                refresh_threaded_compatibility_note(compatibility, root, &replies);
            }
        }
        let legacy_edits = legacy_comment_package_edits(
            &mut source,
            &sheet_paths,
            &legacy_comments,
            &self.workbook.sheets,
            &threaded_edits,
            &drawing_package,
        )?;

        let cursor = Cursor::new(&self.original_data);
        let mut archive = ZipArchive::new(cursor)
            .map_err(|e| XlsxError::InvalidData(e.to_string()))?;

        // Group edits by sheet index
        let mut edits_by_sheet: HashMap<usize, HashMap<(u32, u32), &CellEditValue>> =
            HashMap::new();
        for ((si, row, col), val) in &self.edits {
            edits_by_sheet
                .entry(*si)
                .or_default()
                .insert((*row, *col), val);
        }

        // One xf per distinct style, so a run that bolds a hundred cells adds
        // one entry rather than a hundred.
        let mut distinct: Vec<CellStyle> = Vec::new();
        let mut style_slot: HashMap<(usize, u32, u32), usize> = HashMap::new();
        let has_styles_part = archive.file_names().any(|name| name == "xl/styles.xml");
        for (key, style) in &self.styles {
            if !has_styles_part && *style == CellStyle::default() {
                continue;
            }
            let slot = match distinct.iter().position(|held| held == style) {
                Some(slot) => slot,
                None => {
                    distinct.push(style.clone());
                    distinct.len() - 1
                }
            };
            style_slot.insert(*key, slot);
        }
        let mut row_band_slots = HashMap::new();
        let mut column_band_slots = HashMap::new();
        for (source, slots) in [(&self.row_band_styles, &mut row_band_slots), (&self.column_band_styles, &mut column_band_slots)] {
            for (key, style) in source {
                let slot = style.as_ref().map(|style| {
                    match distinct.iter().position(|held| held == style) {
                        Some(slot) => slot,
                        None => { distinct.push(style.clone()); distinct.len() - 1 }
                    }
                });
                slots.insert(*key, slot);
            }
        }
        let mut added_style_slot = HashMap::new();
        let mut added_row_band_slots = HashMap::new();
        let mut added_column_band_slots = HashMap::new();
        if let Some(plan) = &self.sheet_plan {
            for (index, planned) in plan.iter().enumerate() {
                let PlannedSheet::Added(sheet) = planned else { continue };
                for row in &sheet.rows {
                    for cell in &row.cells {
                        if !has_styles_part && cell.style == CellStyle::default() {
                            continue;
                        }
                        let slot = match distinct.iter().position(|held| held == &cell.style) {
                            Some(slot) => slot,
                            None => {
                                distinct.push(cell.style.clone());
                                distinct.len() - 1
                            }
                        };
                        added_style_slot.insert((index, row.index, cell.col), slot);
                    }
                }
                for (&row, style) in &sheet.band_styles.rows {
                    let slot = match distinct.iter().position(|held| held == style) {
                        Some(slot) => slot, None => { distinct.push(style.clone()); distinct.len() - 1 }
                    };
                    added_row_band_slots.insert((index, row), slot);
                }
                let columns: std::collections::BTreeSet<u32> = if sheet.band_styles.sheet.is_some() {
                    (0..16_384).collect()
                } else { sheet.band_styles.columns.keys().copied().collect() };
                for column in columns {
                    let Some(style) = sheet.band_styles.columns.get(&column).or(sheet.band_styles.sheet.as_ref()) else { continue; };
                    let slot = match distinct.iter().position(|held| held == style) {
                        Some(slot) => slot, None => { distinct.push(style.clone()); distinct.len() - 1 }
                    };
                    added_column_band_slots.insert((index, column), slot);
                }
            }
        }
        let mut conditional_styles: Vec<(Option<DiffStyle>, DiffStyle)> = Vec::new();
        let mut conditional_style_slots = HashMap::new();
        let mut added_conditional_style_slots = HashMap::new();
        for (&sheet, rules) in &self.conditional_rules {
            let before = &self.workbook.sheets[sheet].conditional_rules;
            for rule in rules {
                let old = rule.origin_priority.and_then(|origin| before.iter().find(|old| old.origin_priority == Some(origin)));
                let needed = old.map_or(rule.style.is_some(), |old| old.style != rule.style);
                if needed {
                    let source = old.and_then(|old| old.style.clone()).or_else(|| rule.style.as_ref().and_then(|style| style.source_xml.as_ref()).and_then(|raw|
                        self.workbook.sheets.iter().flat_map(|sheet| &sheet.conditional_rules).filter_map(|rule| rule.style.as_ref())
                            .find(|style| style.source_xml.as_ref() == Some(raw)).cloned()));
                    if old.is_none() && rule.style.as_ref().and_then(|style| style.source_xml.as_ref()).is_some() && source.is_none() {
                        return Err(XlsxError::InvalidData("missing copied differential style".into()));
                    }
                    let pair = (source, rule.style.clone().unwrap_or_default());
                    let slot = conditional_styles.iter().position(|held| held == &pair).unwrap_or_else(|| { conditional_styles.push(pair); conditional_styles.len() - 1 });
                    conditional_style_slots.insert((sheet, old.is_some(), rule.origin_priority.unwrap_or(rule.priority)), slot);
                }
            }
        }
        if let Some(plan) = &self.sheet_plan {
            for (sheet, planned) in plan.iter().enumerate() {
                let PlannedSheet::Added(sheet_data) = planned else { continue; };
                for rule in &sheet_data.conditional_rules {
                    if let Some(style) = &rule.style {
                        let original = style.source_xml.as_ref().and_then(|raw| self.workbook.sheets.iter()
                            .flat_map(|sheet| &sheet.conditional_rules).filter_map(|rule| rule.style.as_ref())
                            .find(|held| held.source_xml.as_ref() == Some(raw)).cloned());
                        if style.source_xml.is_some() && original.is_none() {
                            return Err(XlsxError::InvalidData("missing original differential style".into()));
                        }
                        let pair = (original, style.clone());
                        let slot = conditional_styles.iter().position(|held| held == &pair).unwrap_or_else(|| { conditional_styles.push(pair); conditional_styles.len() - 1 });
                        added_conditional_style_slots.insert((sheet, rule.priority), slot);
                    }
                }
            }
        }
        let create_styles_part = !has_styles_part
            && (self.indexed_colors.is_some() || !distinct.is_empty() || !conditional_styles.is_empty());

        // The style sheet is read ahead of the walk below, since a worksheet
        // needs the indices and may come before it in the archive.
        let mut patched_styles: Option<String> = None;
        // What each planned sheet is called in the package, and which of the
        // original parts it comes from.
        let sheet_paths_by_index = self.resolve_sheet_paths()?;
        let mut hyperlink_relationship_xml = HashMap::new();
        let mut hyperlink_ids: HashMap<usize, Vec<Option<String>>> = HashMap::new();
        for (&sheet, links) in &self.hyperlinks {
            let Some(sheet_path) = sheet_paths_by_index.get(sheet) else { continue };
            let rels_path = sheet_rels_path(sheet_path);
            let original_rels = source.try_read_part(&rels_path)?;
            let mut rels = original_rels.clone().unwrap_or_else(|| {
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>".to_string()
            });
            for additions in [
                threaded_edits.relationship_xml.get(&rels_path),
                legacy_edits.relationship_xml.get(&rels_path),
                drawing_package.relationship_xml.get(&rels_path),
            ].into_iter().flatten() {
                rels = merge_relationship_additions(&rels, additions)?;
            }
            if let Some((_, additions)) = threaded_edits.new_parts.iter().find(|(path, _)| path == &rels_path) {
                let additions = String::from_utf8_lossy(additions);
                rels = merge_relationship_additions(&rels, &additions)?;
            }
            if let Some((_, additions)) = drawing_package.new_parts.iter().find(|(path, _)| path == &rels_path) {
                rels = merge_relationship_additions(&rels, additions)?;
            }
            if let Some((_, additions)) = legacy_edits.new_parts.iter().find(|(path, _)| path == &rels_path) {
                rels = merge_relationship_additions(&rels, additions)?;
            }
            let mut ids = Vec::with_capacity(links.len());
            for link in links {
                if link.internal {
                    ids.push(None);
                } else {
                    let (updated, id) = append_hyperlink_relationship(
                        Some(&rels),
                        &link.target,
                    )?;
                    rels = updated;
                    ids.push(Some(id));
                }
            }
            if !links.iter().all(|link| link.internal)
                && (original_rels.as_deref() != Some(rels.as_str())
                    || threaded_edits.relationship_xml.contains_key(&rels_path)
                    || legacy_edits.relationship_xml.contains_key(&rels_path)
                    || drawing_package.relationship_xml.contains_key(&rels_path))
            {
                hyperlink_relationship_xml.insert(rels_path, rels);
            }
            hyperlink_ids.insert(sheet, ids);
        }
        let mut planned_parts: Vec<(String, Option<usize>)> = Vec::new();
        if let Some(plan) = self.sheet_plan.as_ref() {
            let mut next_part = sheet_paths_by_index.len() + 1;
            for planned in plan {
                match planned {
                    PlannedSheet::Held { origin, .. } => {
                        let path = sheet_paths_by_index.get(*origin).cloned().ok_or_else(|| {
                            XlsxError::InvalidData(
                                "a kept sheet is not in the original workbook".to_string(),
                            )
                        })?;
                        planned_parts.push((path, Some(*origin)));
                    }
                    PlannedSheet::Added(_) => {
                        // Take a part name nothing is using.
                        let mut path = format!("xl/worksheets/sheet{next_part}.xml");
                        while sheet_paths_by_index.contains(&path) {
                            next_part += 1;
                            path = format!("xl/worksheets/sheet{next_part}.xml");
                        }
                        next_part += 1;
                        planned_parts.push((path, None));
                    }
                }
            }
        }
        let dropped: Vec<String> = match self.sheet_plan.as_ref() {
            None => Vec::new(),
            Some(_) => sheet_paths_by_index
                .iter()
                .filter(|path| !planned_parts.iter().any(|(held, _)| held == *path))
                .cloned()
                .collect(),
        };

        let mut style_indices: Vec<u32> = Vec::new();
        let mut conditional_style_indices: Vec<u32> = Vec::new();
        if !distinct.is_empty() || self.indexed_colors.is_some() || !conditional_styles.is_empty() {
            let mut xml = String::new();
            if create_styles_part {
                xml = default_styles_xml();
            } else {
                archive
                    .by_name("xl/styles.xml")
                    .map_err(|_| {
                        XlsxError::InvalidData(
                            "the workbook has no style sheet to add styles to".to_string(),
                        )
                    })?
                    .read_to_string(&mut xml)
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            let (mut patched, indices) = if distinct.is_empty() {
                (xml, Vec::new())
            } else {
                patch_styles_xml(&xml, &distinct)?
            };
            if let Some(indexed_colors) = &self.indexed_colors {
                patched = patch_indexed_colors_xml(&patched, indexed_colors)?;
            }
            if !conditional_styles.is_empty() {
                (patched, conditional_style_indices) = append_differential_styles(&patched, &conditional_styles)?;
            }
            patched_styles = Some(patched);
            style_indices = indices;
        }

        // Where each cell's style ended up, ready for the sheet writer.
        let mut cell_styles: HashMap<usize, BTreeMap<(u32, u32), u32>> = HashMap::new();
        for ((sheet, row, col), slot) in &style_slot {
            if let Some(index) = style_indices.get(*slot) {
                cell_styles
                    .entry(*sheet)
                    .or_default()
                    .insert((*row, *col), *index);
            }
        }
        let mut added_cell_styles: HashMap<usize, BTreeMap<(u32, u32), u32>> = HashMap::new();
        for ((sheet, row, col), slot) in &added_style_slot {
            added_cell_styles.entry(*sheet).or_default().insert((*row, *col), style_indices[*slot]);
        }
        let mut added_row_band_styles: HashMap<usize, BTreeMap<u32, u32>> = HashMap::new();
        for ((sheet, row), slot) in &added_row_band_slots {
            added_row_band_styles.entry(*sheet).or_default().insert(*row, style_indices[*slot]);
        }
        let mut added_column_band_styles: HashMap<usize, BTreeMap<u32, u32>> = HashMap::new();
        for ((sheet, column), slot) in &added_column_band_slots {
            added_column_band_styles.entry(*sheet).or_default().insert(*column, style_indices[*slot]);
        }

        let mut rows_by_sheet: HashMap<usize, BTreeMap<u32, RowWant>> = HashMap::new();
        for ((sheet, row), hidden) in &self.row_hidden {
            rows_by_sheet
                .entry(*sheet)
                .or_default()
                .entry(*row)
                .or_default()
                .hidden = Some(*hidden);
        }
        for ((sheet, row), (height, pinned)) in &self.row_height {
            let want = rows_by_sheet.entry(*sheet).or_default().entry(*row).or_default();
            want.height = Some(*height);
            want.pinned = *pinned;
        }
        for ((sheet, row), slot) in &row_band_slots {
            rows_by_sheet.entry(*sheet).or_default().entry(*row).or_default().style = Some(slot.map(|slot| style_indices[slot]));
        }
        let mut cols_by_sheet: HashMap<usize, BTreeMap<u32, ColumnWant>> = HashMap::new();
        for ((sheet, col), hidden) in &self.col_hidden {
            cols_by_sheet
                .entry(*sheet)
                .or_default()
                .entry(*col)
                .or_default()
                .hidden = Some(*hidden);
        }
        for ((sheet, col), width) in &self.col_width {
            cols_by_sheet
                .entry(*sheet)
                .or_default()
                .entry(*col)
                .or_default()
                .width = Some(*width);
        }
        for ((sheet, col), custom) in &self.col_custom_width {
            cols_by_sheet.entry(*sheet).or_default().entry(*col).or_default().custom = Some(*custom);
        }
        for ((sheet, col), slot) in &column_band_slots {
            cols_by_sheet.entry(*sheet).or_default().entry(*col).or_default().style = Some(slot.map(|slot| style_indices[slot]));
        }


        // Map sheet path -> everything to change in that sheet
        let empty_cells: HashMap<(u32, u32), &CellEditValue> = HashMap::new();
        let empty_lines: BTreeMap<u32, RowWant> = BTreeMap::new();
        let empty_cols: BTreeMap<u32, ColumnWant> = BTreeMap::new();
        let empty_styles: BTreeMap<(u32, u32), u32> = BTreeMap::new();
        let mut path_edits: HashMap<String, SheetEdits<'_>> = HashMap::new();
        for sheet in edits_by_sheet
            .keys()
            .chain(rows_by_sheet.keys())
            .chain(cols_by_sheet.keys())
            .chain(self.merges.keys())
            .chain(cell_styles.keys())
            .chain(self.filters.keys())
            .chain(self.validations.keys())
            .chain(self.conditional_rules.keys())
            .chain(self.hyperlinks.keys())
            .chain(self.panes.keys())
            .copied()
            .collect::<std::collections::BTreeSet<_>>()
        {
            if let Some(path) = sheet_paths.get(sheet) {
                path_edits.insert(
                    path.clone(),
                    SheetEdits {
                        cells: edits_by_sheet.get(&sheet).unwrap_or(&empty_cells),
                        rows: rows_by_sheet.get(&sheet).unwrap_or(&empty_lines),
                        cols: cols_by_sheet.get(&sheet).unwrap_or(&empty_cols),
                        merges: self.merges.get(&sheet).map(Vec::as_slice),
                        styles: cell_styles.get(&sheet).unwrap_or(&empty_styles),
                        filter: self.filters.get(&sheet),
                        panes: self.panes.get(&sheet).copied(),
                        hyperlinks: self.hyperlinks.get(&sheet).map(Vec::as_slice),
                        hyperlink_ids: hyperlink_ids.get(&sheet).map(Vec::as_slice).unwrap_or(&[]),
                    },
                );
            }
        }

        let mut output = Vec::new();
        {
            let mut writer = ZipWriter::new(Cursor::new(&mut output));

            for i in 0..archive.len() {
                let mut entry = archive.by_index(i)
                    .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                let name = entry.name().to_string();
                let options = SimpleFileOptions::default()
                    .compression_method(entry.compression());

                writer.start_file(&name, options)
                    .map_err(|e| XlsxError::InvalidData(e.to_string()))?;

                if let Ok(origin)=chart_parts.binary_search(&name) {
                    if let Some((before,after))=self.chart_edits.get(&(origin as u32)) {
                        let mut xml=String::new();
                        entry.read_to_string(&mut xml).map_err(|error|XlsxError::InvalidData(error.to_string()))?;
                        let patched=crate::chart_edit::patch_chart(&xml,before,after)?;
                        writer.write_all(patched.as_bytes()).map_err(|error|XlsxError::InvalidData(error.to_string()))?;
                        continue;
                    }
                }

                if let Some(patched) = drawing_package.parts.get(&name) {
                    writer.write_all(patched.as_bytes()).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    continue;
                }

                if let Some(xml) = threaded_edits.person_xml.get(&name) {
                    writer.write_all(xml.as_bytes())
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    continue;
                }

                if let Some(edit) = threaded_edits.comments.get(&name) {
                    let mut xml = String::new();
                    entry.read_to_string(&mut xml)
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    let patched = patch_threaded_comment_xml(&xml, &edit.text, &edit.deletions, &edit.resolved)?;
                    let patched = append_threaded_comments(&patched, &edit.additions)?;
                    writer.write_all(patched.as_bytes())
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    continue;
                }

                if let Some(xml) = legacy_edits.parts.get(&name) {
                    writer.write_all(xml.as_bytes())
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    continue;
                }

                if dropped.contains(&name) {
                    // The sheet is gone; leave its part out of the package.
                    writer
                        .abort_file()
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    continue;
                }
                // The calculation chain names the cells that hold formulas and
                // the order Excel works them out in. Writing a value over a
                // formula takes the formula away, and a chain still pointing at
                // that cell makes Excel refuse the WHOLE workbook  E`fies_t2`
                // has one entry, `C1`, and replacing C1's `<f>C9</f>` with a
                // value lost the file. The chain is only a cache: Excel builds
                // it again on open, so the safe answer is to leave it out
                // whenever a cell has been touched.
                if name == "xl/calcChain.xml" && self.has_non_chart_edits() {
                    writer
                        .abort_file()
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    continue;
                }
                if name == "xl/workbook.xml"
                    && (self.sheet_plan.is_some() || self.name_plan.is_some() || self.calculation_plan.is_some())
                {
                    let mut xml = String::new();
                    entry
                        .read_to_string(&mut xml)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    if let Some(plan) = self.sheet_plan.as_ref() {
                        let names: Vec<(String, Option<&'static str>)> = plan
                            .iter()
                            .map(|planned| match planned {
                                PlannedSheet::Held { name, visibility, .. } => {
                                    (name.clone(), visibility.stated())
                                }
                                PlannedSheet::Added(sheet) => {
                                    (sheet.name.clone(), sheet.visibility.stated())
                                }
                            })
                            .collect();
                        xml = patch_workbook_sheets(&xml, &names)?;
                    }
                    if let Some(plan) = self.name_plan.as_ref() {
                        xml = patch_workbook_names(&xml, plan, self.name_attributes_plan.as_ref().unwrap_or(&self.workbook.defined_name_attributes))?;
                    }
                    if let Some(plan) = self.calculation_plan {
                        xml = patch_workbook_calculation(&xml, plan)?;
                    }
                    writer
                        .write_all(xml.as_bytes())
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    continue;
                }
                if name != "xl/_rels/workbook.xml.rels"
                    && name != "[Content_Types].xml"
                    && (threaded_edits.relationship_xml.contains_key(&name)
                        || legacy_edits.relationship_xml.contains_key(&name)
                        || drawing_package.relationship_xml.contains_key(&name)
                        || hyperlink_relationship_xml.contains_key(&name))
                {
                    let mut xml = String::new();
                    entry.read_to_string(&mut xml)
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    let mut patched = xml;
                    if let Some(additions) = threaded_edits.relationship_xml.get(&name) {
                        patched = merge_relationship_additions(&patched, additions)?;
                    }
                    if let Some(additions) = legacy_edits.relationship_xml.get(&name) {
                        patched = merge_relationship_additions(&patched, additions)?;
                    }
                    if let Some(additions) = drawing_package.relationship_xml.get(&name) {
                        patched = merge_relationship_additions(&patched, additions)?;
                    }
                    if let Some(additions) = hyperlink_relationship_xml.get(&name) {
                        patched = merge_relationship_additions(additions, &patched)?;
                    }
                    writer.write_all(patched.as_bytes())
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    continue;
                }
                if self.sheet_plan.is_some() && name == "xl/_rels/workbook.xml.rels" {
                    let mut xml = String::new();
                    entry
                        .read_to_string(&mut xml)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    let targets: Vec<String> = planned_parts
                        .iter()
                        .map(|(path, _)| {
                            path.strip_prefix("xl/").unwrap_or(path).to_string()
                        })
                        .collect();
                    let mut patched = patch_workbook_rels(&xml, &targets)?;
                    if let Some(additions) = threaded_edits.relationship_xml.get(&name) {
                        patched = merge_relationship_additions(&patched, additions)?;
                    }
                    if create_styles_part {
                        patched = append_workbook_styles_relationship(&patched)?;
                    }
                    writer
                        .write_all(patched.as_bytes())
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    continue;
                }
                if self.sheet_plan.is_some() && name == "[Content_Types].xml" {
                    let mut xml = String::new();
                    entry
                        .read_to_string(&mut xml)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    let parts: Vec<String> =
                        planned_parts.iter().map(|(path, _)| path.clone()).collect();
                    let mut patched = patch_content_types(&xml, &parts)?;
                    for (part, kind) in &threaded_edits.content_types {
                        patched = append_content_type(&patched, part, kind)?;
                    }
                    for (part, kind) in &legacy_edits.content_types {
                        patched = append_content_type(&patched, part, kind)?;
                    }
                    for (part, kind) in &drawing_package.content_types {
                        patched = append_content_type(&patched, part, kind)?;
                    }
                    if create_styles_part {
                        patched = append_content_type(
                            &patched,
                            "xl/styles.xml",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                        )?;
                    }
                    if legacy_edits.needs_vml_default {
                        patched = append_default_content_type(&patched, "vml", "application/vnd.openxmlformats-officedocument.vmlDrawing")?;
                    }
                    writer
                        .write_all(patched.as_bytes())
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    continue;
                }
                // The package must not name a part that is no longer in it:
                // the chain's own relationship and content type go with it.
                // Excel opens the workbook either way, but a relationship
                // pointing at nothing is not a package, and "invalid but
                // tolerated" is what this whole exercise is meant to catch.
                if self.has_non_chart_edits() && (name == "xl/_rels/workbook.xml.rels" || name == "[Content_Types].xml") {
                    let mut xml = String::new();
                    entry
                        .read_to_string(&mut xml)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    let mut patched = without_calc_chain(&xml);
                    if name == "[Content_Types].xml" {
                        for (part, kind) in &threaded_edits.content_types {
                            patched = append_content_type(&patched, part, kind)?;
                        }
                        for (part, kind) in &legacy_edits.content_types {
                            patched = append_content_type(&patched, part, kind)?;
                        }
                    for (part, kind) in &drawing_package.content_types {
                        patched = append_content_type(&patched, part, kind)?;
                    }
                        if legacy_edits.needs_vml_default {
                            patched = append_default_content_type(&patched, "vml", "application/vnd.openxmlformats-officedocument.vmlDrawing")?;
                        }
                    }
                    if create_styles_part && name == "xl/_rels/workbook.xml.rels" {
                        patched = append_workbook_styles_relationship(&patched)?;
                    }
                    if create_styles_part && name == "[Content_Types].xml" {
                        patched = append_content_type(
                            &patched,
                            "xl/styles.xml",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                        )?;
                    }
                    if let Some(additions) = threaded_edits.relationship_xml.get(&name) {
                        patched = merge_relationship_additions(&patched, additions)?;
                    }
                    if let Some(additions) = legacy_edits.relationship_xml.get(&name) {
                        patched = merge_relationship_additions(&patched, additions)?;
                    }
                    writer
                        .write_all(patched.as_bytes())
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    continue;
                }
                if name == "xl/styles.xml" {
                    if let Some(patched) = patched_styles.as_deref() {
                        writer
                            .write_all(patched.as_bytes())
                            .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        continue;
                    }
                }
                // A worksheet part whose sheet had its tab colour changed
                // gets that written in, whether or not its cells also changed.
                let colour_change = sheet_paths_by_index
                    .iter()
                    .position(|part| *part == name)
                    .and_then(|origin| self.tab_colors.get(&origin));
                let default_width = sheet_paths_by_index.iter().position(|part| *part == name)
                    .and_then(|origin| self.default_col_width.get(&origin)).copied();
                let validation_change = sheet_paths_by_index.iter().position(|part| *part == name)
                    .and_then(|origin| self.validations.get(&origin).map(|rules| (origin, rules)));
                let conditional_change = sheet_paths_by_index.iter().position(|part| *part == name)
                    .and_then(|origin| self.conditional_rules.get(&origin).map(|rules| (origin, rules)));
                if let Some(sheet_edits) = path_edits.get(&name) {
                    let mut xml = String::new();
                    entry.read_to_string(&mut xml)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    let mut patched = patch_worksheet_xml(&xml, sheet_edits)?;
                    if let Some((origin, rules)) = validation_change {
                        patched = patch_data_validations_xml(&patched, &self.workbook.sheets[origin].validations, rules)?;
                    }
                    if let Some((origin, rules)) = conditional_change {
                        let ids = conditional_style_slots.iter().filter_map(|(&(sheet, is_existing, priority), &slot)| (sheet == origin).then_some(((is_existing, priority), conditional_style_indices[slot]))).collect();
                        patched = patch_conditional_rules_xml(&patched, &self.workbook.sheets[origin].conditional_rules, rules, &ids)?;
                    }
                    if let Some(links) = sheet_edits.hyperlinks {
                        patched = patch_hyperlinks_xml(&patched, links, sheet_edits.hyperlink_ids)?;
                    }
                    if let Some(width) = default_width { patched = patch_default_column_width(&patched, width)?; }
                    if let Some(colour) = colour_change {
                        patched = set_tab_color_in_worksheet(&patched, colour);
                    }
                    if let Some(rid) = legacy_edits.worksheet_drawing.get(&name) {
                        patched = set_legacy_drawing(&patched, rid)?;
                    }
                    if let Some(rid) = drawing_package.worksheet_legacy_drawing.get(&name) {
                        patched = set_legacy_drawing(&patched, rid)?;
                    }
                    if let Some(rid) = drawing_package.worksheet_drawing.get(&name) {
                        patched = set_drawing(&patched, rid)?;
                    }
                    if let Some(additions) = drawing_package.worksheet_ole_objects.get(&name) {
                        patched = append_ole_objects_xml(&patched, additions)?;
                    }
                    if let Some(anchors) = drawing_package.ole_anchors.get(&name) {
                        patched = patch_ole_object_anchors(&patched, anchors)?;
                    }
                    writer.write_all(patched.as_bytes())
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                } else if colour_change.is_some() || default_width.is_some() || legacy_edits.worksheet_drawing.contains_key(&name) || drawing_package.worksheet_drawing.contains_key(&name) || drawing_package.worksheet_legacy_drawing.contains_key(&name) || drawing_package.ole_anchors.contains_key(&name) || drawing_package.worksheet_ole_objects.contains_key(&name) {
                    let mut xml = String::new();
                    entry.read_to_string(&mut xml)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    let mut recoloured = xml;
                    if let Some(colour) = colour_change { recoloured = set_tab_color_in_worksheet(&recoloured, colour); }
                    if let Some(width) = default_width { recoloured = patch_default_column_width(&recoloured, width)?; }
                    if let Some(rid) = legacy_edits.worksheet_drawing.get(&name) {
                        recoloured = set_legacy_drawing(&recoloured, rid)?;
                    }
                    if let Some(rid) = drawing_package.worksheet_legacy_drawing.get(&name) {
                        recoloured = set_legacy_drawing(&recoloured, rid)?;
                    }
                    if let Some(rid) = drawing_package.worksheet_drawing.get(&name) {
                        recoloured = set_drawing(&recoloured, rid)?;
                    }
                    if let Some(additions) = drawing_package.worksheet_ole_objects.get(&name) {
                        recoloured = append_ole_objects_xml(&recoloured, additions)?;
                    }
                    if let Some(anchors) = drawing_package.ole_anchors.get(&name) {
                        recoloured = patch_ole_object_anchors(&recoloured, anchors)?;
                    }
                    writer.write_all(recoloured.as_bytes())
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                } else {
                    let mut buf = Vec::new();
                    entry.read_to_end(&mut buf)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    writer.write_all(&buf)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                }
            }

            // Sheets the file never held are written after everything it did.
            if let Some(plan) = self.sheet_plan.as_ref() {
                for (index, (planned, (path, origin))) in plan.iter().zip(&planned_parts).enumerate() {
                    if origin.is_some() {
                        continue;
                    }
                    let PlannedSheet::Added(sheet) = planned else {
                        continue;
                    };
                    writer
                        .start_file(path, SimpleFileOptions::default())
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    let empty_styles = BTreeMap::new();
                    let empty_band_styles = BTreeMap::new();
                    let mut relationships = None;
                    let mut relationship_ids = Vec::with_capacity(sheet.hyperlinks.len());
                    for hyperlink in &sheet.hyperlinks {
                        if hyperlink.internal {
                            relationship_ids.push(None);
                        } else {
                            let (updated, id) = append_hyperlink_relationship(
                                relationships.as_deref(),
                                &hyperlink.target,
                            )?;
                            relationships = Some(updated);
                            relationship_ids.push(Some(id));
                        }
                    }
                    let memo = std::cell::RefCell::new(BTreeMap::new());
                    let names = self.name_plan.as_deref().unwrap_or(&self.workbook.defined_names);
                    let multiple = |name: &str| cached_named_array_shape(names, &sheet.name, name, &memo);
                    let mut sheet_xml = write_new_sheet(
                        sheet,
                        added_cell_styles.get(&index).unwrap_or(&empty_styles),
                        added_row_band_styles.get(&index).unwrap_or(&empty_band_styles),
                        added_column_band_styles.get(&index).unwrap_or(&empty_band_styles),
                        &multiple,
                    )?;
                    if !sheet.conditional_rules.is_empty() {
                        let ids = added_conditional_style_slots.iter().filter_map(|(&(sheet_index, priority), &slot)| (sheet_index == index).then_some((priority, conditional_style_indices[slot]))).collect();
                        sheet_xml = append_conditional_rules_xml(std::str::from_utf8(&sheet_xml).map_err(|error| XlsxError::InvalidData(error.to_string()))?, &sheet.conditional_rules, &ids)?.into_bytes();
                    }
                    if !sheet.hyperlinks.is_empty() {
                        sheet_xml = patch_hyperlinks_xml(
                            std::str::from_utf8(&sheet_xml)
                                .map_err(|error| XlsxError::InvalidData(error.to_string()))?,
                            &sheet.hyperlinks,
                            &relationship_ids,
                        )?.into_bytes();
                    }
                    writer.write_all(&sheet_xml)
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    if let Some(relationships) = relationships {
                        let rels_path = sheet_rels_path(path);
                        writer.start_file(rels_path, SimpleFileOptions::default())
                            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                        writer.write_all(relationships.as_bytes())
                            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    }
                }
            }

            let mut emitted_new_parts = std::collections::HashSet::new();
            for (path, bytes) in &threaded_edits.new_parts {
                writer.start_file(path, SimpleFileOptions::default())
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                let xml = String::from_utf8_lossy(bytes);
                let mut patched = if let Some(edit) = threaded_edits.comments.get(path) {
                    append_threaded_comments(&xml, &edit.additions)?
                } else {
                    xml.into_owned()
                };
                if path.ends_with(".rels") {
                    if let Some((_, additions)) = legacy_edits.new_parts.iter().find(|(held, _)| held == path) {
                        patched = merge_relationship_additions(&patched, additions)?;
                        emitted_new_parts.insert(path.clone());
                    }
                    if let Some((_, additions)) = drawing_package.new_parts.iter().find(|(held, _)| held == path) {
                        patched = merge_relationship_additions(&patched, additions)?;
                        emitted_new_parts.insert(path.clone());
                    }
                    if let Some(additions) = drawing_package.relationship_xml.get(path) {
                        patched = merge_relationship_additions(&patched, additions)?;
                        emitted_new_parts.insert(path.clone());
                    }
                    if let Some(additions) = hyperlink_relationship_xml.get(path) {
                        patched = merge_relationship_additions(&patched, additions)?;
                        emitted_new_parts.insert(path.clone());
                    }
                }
                writer.write_all(patched.as_bytes())
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }

            for (path, xml) in &legacy_edits.new_parts {
                if emitted_new_parts.contains(path) { continue; }
                writer.start_file(path, SimpleFileOptions::default())
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                let mut patched = xml.clone();
                if path.ends_with(".rels") {
                    if let Some((_, additions)) = drawing_package.new_parts.iter().find(|(held, _)| held == path) {
                        patched = merge_relationship_additions(&patched, additions)?;
                        emitted_new_parts.insert(path.clone());
                    }
                    if let Some(additions) = hyperlink_relationship_xml.get(path) {
                        patched = merge_relationship_additions(&patched, additions)?;
                        emitted_new_parts.insert(path.clone());
                    }
                }
                writer.write_all(patched.as_bytes())
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }

            for (path, xml) in &drawing_package.new_parts {
                if emitted_new_parts.contains(path) { continue; }
                writer.start_file(path, SimpleFileOptions::default())
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                let mut patched = xml.clone();
                if path.ends_with(".rels") {
                    if let Some(additions) = threaded_edits.relationship_xml.get(path) {
                        patched = merge_relationship_additions(&patched, additions)?;
                    }
                    if let Some(additions) = legacy_edits.relationship_xml.get(path) {
                        patched = merge_relationship_additions(&patched, additions)?;
                    }
                    if let Some((_, additions)) = threaded_edits.new_parts.iter().find(|(held, _)| held == path) {
                        let additions = String::from_utf8_lossy(additions);
                        patched = merge_relationship_additions(&patched, &additions)?;
                    }
                    if let Some((_, additions)) = legacy_edits.new_parts.iter().find(|(held, _)| held == path) {
                        patched = merge_relationship_additions(&patched, additions)?;
                    }
                    if let Some(additions) = hyperlink_relationship_xml.get(path) {
                        patched = merge_relationship_additions(&patched, additions)?;
                    }
                }
                writer.write_all(patched.as_bytes())
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }

            for (path, bytes) in &drawing_package.new_binary_parts {
                writer.start_file(path, SimpleFileOptions::default())
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                writer.write_all(bytes)
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }

            for (path, xml) in &hyperlink_relationship_xml {
                if emitted_new_parts.contains(path) || archive.file_names().any(|name| name == path) { continue; }
                writer.start_file(path, SimpleFileOptions::default())
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                let mut patched = xml.clone();
                if let Some(additions) = threaded_edits.relationship_xml.get(path) {
                    patched = merge_relationship_additions(&patched, additions)?;
                }
                if let Some(additions) = legacy_edits.relationship_xml.get(path) {
                    patched = merge_relationship_additions(&patched, additions)?;
                }
                if let Some(additions) = drawing_package.relationship_xml.get(path) {
                    patched = merge_relationship_additions(&patched, additions)?;
                }
                writer.write_all(patched.as_bytes())
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                emitted_new_parts.insert(path.clone());
            }

            if create_styles_part {
                let styles = patched_styles.as_deref().ok_or_else(|| {
                    XlsxError::InvalidData("palette edit did not create styles.xml".into())
                })?;
                writer.start_file("xl/styles.xml", SimpleFileOptions::default())
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                writer.write_all(styles.as_bytes())
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }

            writer.finish().map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        }

        let output = if let Some(tables) = &self.table_plan { patch_workbook_tables(&output, tables)? } else { output };
        let mut arrays = self.dynamic_plan.clone().unwrap_or_else(|| dynamic_array_plan(&self.workbook));
        if self.dynamic_plan.is_none() {
            for (sheet, entries) in arrays.iter_mut().enumerate() {
                entries.retain(|((col, row), _)| !self.edits.contains_key(&(sheet, row + 1, *col)));
            }
        }
        let output = patch_dynamic_array_package(&output, &arrays)?;
        let contains_rich_error = |sheet: &crate::ir::Sheet| sheet.rows.iter().flat_map(|row| &row.cells)
            .any(|cell| matches!(&cell.value, crate::ir::CellValue::Error(error) if crate::parser::rich_error_type(error).is_some()));
        if !self.edits.values().any(|value| value.error_text().and_then(crate::parser::rich_error_type).is_some())
            && !self.workbook.sheets.iter().any(contains_rich_error)
            && !self.sheet_plan.as_ref().is_some_and(|plan| plan.iter().any(|sheet| matches!(sheet, PlannedSheet::Added(sheet) if contains_rich_error(sheet)))) {
            return Ok(output);
        }
        let origins: Vec<_> = match &self.sheet_plan {
            Some(plan) => plan.iter().map(|sheet| match sheet { PlannedSheet::Held { origin, .. } => Some(*origin), PlannedSheet::Added(_) => None }).collect(),
            None => (0..self.workbook.sheets.len()).map(Some).collect(),
        };
        let clear_errors: Vec<std::collections::BTreeSet<_>> = origins.iter().map(|origin| self.edits.iter()
            .filter(|((sheet, _, _), value)| Some(*sheet) == *origin && value.error_text().and_then(crate::parser::rich_error_type).is_none())
            .map(|((_, row, col), _)| (*row, *col)).collect()).collect();
        patch_rich_error_package(&output, &clear_errors)
    }

    /// Resolve sheet index -> ZIP path for each sheet.
    fn resolve_sheet_paths(&self) -> Result<Vec<String>, XlsxError> {
        Self::sheet_paths_in(&self.original_data)
    }

    fn sheet_paths_in(data: &[u8]) -> Result<Vec<String>, XlsxError> {
        let mut archive = OoxmlArchive::new(data)?;
        let workbook_xml = archive.read_part("xl/workbook.xml")?;
        let rels_xml = archive.read_part("xl/_rels/workbook.xml.rels")?;

        // Parse sheet rIds
        let mut reader = Reader::from_str(&workbook_xml);
        let mut r_ids = Vec::new();
        loop {
            match reader.read_event().map_err(XlsxError::Xml)? {
                Event::Start(e) | Event::Empty(e) => {
                    if local_name(e.name().as_ref()) == "sheet" {
                        let r_id = get_attr(&e, "id")
                            .or_else(|| {
                                for attr in e.attributes().flatten() {
                                    let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                                    if key == "r:id" {
                                        return Some(String::from_utf8_lossy(&attr.value).to_string());
                                    }
                                }
                                None
                            })
                            .unwrap_or_default();
                        r_ids.push(r_id);
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }

        let rels = parse_relationships(&rels_xml)?;
        let rid_to_path: HashMap<String, String> = rels
            .into_iter()
            .map(|(id, rel)| (id, rel.target))
            .collect();

        let mut paths = Vec::new();
        for r_id in &r_ids {
            if let Some(target) = rid_to_path.get(r_id) {
                let path = oxidocs_common::security::sanitize_rel_target("xl", target)
                    .unwrap_or_default();
                paths.push(path);
            } else {
                paths.push(String::new());
            }
        }

        Ok(paths)
    }
}

/// Retain unknown table XML while changing the attributes modelled by the IR.
fn table_root_attributes(xml: &str, attributes: &[(&str, String)]) -> Result<String, XlsxError> {
    table_root_attributes_removed(xml, attributes, &[])
}

fn table_root_attributes_removed(xml: &str, attributes: &[(&str, String)], removed: &[&str]) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(xml);
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        let end = reader.buffer_position() as usize;
        if let Event::Start(element) | Event::Empty(element) = &event {
            let mut patched = element.to_owned();
            patched.clear_attributes();
            for attribute in element.attributes().flatten() {
                if !attributes.iter().any(|(key, _)| attribute.key.as_ref() == key.as_bytes())
                    && !removed.iter().any(|key| attribute.key.as_ref() == key.as_bytes()) {
                    patched.push_attribute(attribute);
                }
            }
            for (key, value) in attributes { patched.push_attribute((*key, value.as_str())); }
            let mut writer = Writer::new(Cursor::new(Vec::new()));
            writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(patched) } else { Event::Start(patched) })
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            let head = String::from_utf8(writer.into_inner().into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            return Ok(format!("{}{head}{}", &xml[..start], &xml[end..]));
        }
        if matches!(event, Event::Eof) { return Err(XlsxError::InvalidData("table XML has no root".into())); }
    }
}

fn table_root_attribute(xml: &str, key: &str) -> Option<String> {
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event().ok()? {
            Event::Start(element) | Event::Empty(element) => return get_attr(&element, key),
            Event::Eof => return None,
            _ => {}
        }
    }
}

fn table_column_elements(xml: &str) -> Result<Vec<String>, XlsxError> {
    table_child_elements(xml, "tableColumn")
}

fn table_child_elements(xml: &str, wanted: &str) -> Result<Vec<String>, XlsxError> {
    Ok(table_child_spans(xml, wanted)?.into_iter().map(|(start, end)| xml[start..end].to_string()).collect())
}

fn table_child_spans(xml: &str, wanted: &str) -> Result<Vec<(usize, usize)>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut result = Vec::new();
    let mut depth = 0usize;
    let mut start = None;
    loop {
        let before = reader.buffer_position() as usize;
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(element) => {
                if depth == 1 && local_name(element.name().as_ref()) == wanted { start = Some(before); }
                depth += 1;
            }
            Event::Empty(element) if depth == 1 && local_name(element.name().as_ref()) == wanted =>
                result.push((before, reader.buffer_position() as usize)),
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 1 { if let Some(start) = start.take() { result.push((start, reader.buffer_position() as usize)); } }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(result)
}

fn table_filter_buttons(xml: &str, table: &crate::ir::Table) -> Result<String, XlsxError> {
    let source = table_child_elements(xml, "filterColumn")?;
    let mut columns = String::new();
    for column in 0..table.columns.len() as u32 {
        let hidden = table.hidden_filter_buttons.contains(&column);
        let original = source.iter().find(|xml| table_root_attribute(xml, "colId").and_then(|id| id.parse::<u32>().ok()) == Some(column));
        if let Some(original) = original {
            columns.push_str(&table_root_attributes(original, &[
                ("hiddenButton", u8::from(hidden).to_string()), ("showButton", "1".into()),
            ])?);
        } else if hidden {
            columns.push_str(&format!("<filterColumn xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" colId=\"{column}\" hiddenButton=\"1\"/>"));
        }
    }
    let mut output = xml.to_string();
    for (start, end) in table_child_spans(xml, "filterColumn")?.into_iter().rev() {
        output.replace_range(start..end, "");
    }
    if !columns.is_empty() {
        output = set_worksheet_drawing_child(&output, "filterColumn", &columns, &["sortState", "extLst"])?;
    }
    Ok(output)
}

fn table_xml(table: &crate::ir::Table, id: u32, original: Option<&str>) -> Result<String, XlsxError> {
    if table.start_row == 0 || table.end_row < table.start_row || table.end_col < table.start_col
        || table.columns.len() != (table.end_col - table.start_col + 1) as usize {
        return Err(XlsxError::InvalidData("invalid table range or column count".into()));
    }
    let reference = format!("{}{}:{}{}", col_to_letter(table.start_col), table.start_row, col_to_letter(table.end_col), table.end_row);
    let empty = "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></table>";
    let mut output = table_root_attributes(original.unwrap_or(empty), &[
        ("id", id.to_string()), ("name", table.name.clone()), ("displayName", table.name.clone()),
        ("ref", reference.clone()), ("headerRowCount", table.header_rows.to_string()),
        ("totalsRowCount", table.totals_rows.to_string()), ("totalsRowShown", u8::from(table.totals_rows > 0).to_string()),
        ("insertRow", u8::from(table.insert_row).to_string()),
    ])?;
    let original_columns = direct_xml_child(&output, "tableColumns")?.0
        .map(|(start, end)| table_column_elements(&output[start..end])).transpose()?.unwrap_or_default();
    let mut used = std::collections::HashSet::new();
    let mut next_id = original_columns.iter().filter_map(|column| table_root_attribute(column, "id")?.parse::<u32>().ok()).max().unwrap_or(0) + 1;
    let mut columns = String::new();
    for (index, name) in table.columns.iter().enumerate() {
        let source = original_columns.iter().enumerate().find(|(at, column)| !used.contains(at) && table_root_attribute(column, "name").as_deref() == Some(name.as_str()))
            .map(|(at, _)| at).or_else(|| (index < original_columns.len() && !used.contains(&index)).then_some(index));
        let (xml, column_id) = if let Some(at) = source {
            used.insert(at);
            (original_columns[at].as_str(), table_root_attribute(&original_columns[at], "id").unwrap_or_else(|| (at + 1).to_string()))
        } else {
            let id = next_id; next_id += 1; ("<tableColumn/>", id.to_string())
        };
        let mut column = table_root_attributes(xml, &[("id", column_id), ("name", name.clone())])?;
        if let Some(totals) = table.totals.get(index) {
            let mut attributes = Vec::new();
            if let Some(function) = &totals.function { attributes.push(("totalsRowFunction", function.clone())); }
            if let Some(label) = &totals.label { attributes.push(("totalsRowLabel", crate::string_value::encode(label))); }
            column = table_root_attributes_removed(&column, &attributes, &["totalsRowFunction", "totalsRowLabel"])?;
            if let Some(formula) = &totals.formula {
                let original = direct_xml_child(&column, "totalsRowFormula")?.0.map(|(start, end)| &column[start..end]);
                let mut reader = Reader::from_str(original.unwrap_or("<totalsRowFormula xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>"));
                let root = loop { match reader.read_event().map_err(XlsxError::Xml)? {
                    Event::Start(element) | Event::Empty(element) => break element.into_owned(),
                    Event::Eof => return Err(XlsxError::InvalidData("totals formula has no root".into())),
                    _ => {}
                } };
                let end = BytesEnd::new(String::from_utf8_lossy(root.name().as_ref()).into_owned());
                let formula_text = crate::string_value::encode(formula);
                let mut writer = Writer::new(Cursor::new(Vec::new()));
                for event in [Event::Start(root), Event::Text(BytesText::new(&formula_text)), Event::End(end)] {
                    writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                }
                let formula = String::from_utf8(writer.into_inner().into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                column = set_worksheet_drawing_child(&column, "totalsRowFormula", &formula, &["extLst"])?;
            } else { column = replace_direct_xml_child(&column, "totalsRowFormula", None)?; }
        }
        columns.push_str(&column);
    }
    let columns = format!("<tableColumns xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{}\">{columns}</tableColumns>", table.columns.len());
    output = set_worksheet_drawing_child(&output, "tableColumns", &columns, &["tableStyleInfo", "extLst"])?;
    if table.header_rows > 0 && table.show_auto_filter {
        let filter_ref = format!("{}{}:{}{}", col_to_letter(table.start_col), table.start_row, col_to_letter(table.end_col), table.end_row.saturating_sub(table.totals_rows).max(table.start_row));
        let filter = table.auto_filter_xml.clone().unwrap_or_else(|| "<autoFilter xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>".into());
        let filter = table_root_attributes(&filter, &[("ref", filter_ref)])?;
        let filter = table_filter_buttons(&filter, table)?;
        output = set_worksheet_drawing_child(&output, "autoFilter", &filter, &["sortState", "tableColumns", "tableStyleInfo", "extLst"])?;
    } else { output = replace_direct_xml_child(&output, "autoFilter", None)?; }
    let style = direct_xml_child(&output, "tableStyleInfo")?.0.map(|(start, end)| output[start..end].to_string())
        .unwrap_or_else(|| "<tableStyleInfo xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" showFirstColumn=\"0\" showLastColumn=\"0\" showColumnStripes=\"0\"/>".into());
    let style = table_root_attributes(&style, &[
        ("name", table.style.clone().unwrap_or_default()), ("showRowStripes", u8::from(table.banded_rows).to_string()),
        ("showColumnStripes", u8::from(table.banded_columns).to_string()),
        ("showFirstColumn", u8::from(table.emphasize_first_column).to_string()),
        ("showLastColumn", u8::from(table.emphasize_last_column).to_string()),
    ])?;
    set_worksheet_drawing_child(&output, "tableStyleInfo", &style, &["extLst"])
}

struct TableFormulaRename {
    old: String,
    new: String,
    columns: Vec<(String, String)>,
}

fn rewrite_table_formula_xml(xml: &str, renames: &[TableFormulaRename]) -> Result<String, XlsxError> {
    let own = table_root_attribute(xml, "displayName").or_else(|| table_root_attribute(xml, "name")).unwrap_or_default();
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut active = false;
    let mut text = String::new();
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        match &event {
            Event::Start(element) if matches!(local_name(element.name().as_ref()).as_str(), "calculatedColumnFormula" | "totalsRowFormula") => {
                active = true; text.clear();
            }
            Event::Text(value) if active => { text.push_str(&value.unescape().map_err(XlsxError::Xml)?); continue; }
            Event::CData(value) if active => { text.push_str(&String::from_utf8_lossy(value)); continue; }
            Event::End(element) if active && matches!(local_name(element.name().as_ref()).as_str(), "calculatedColumnFormula" | "totalsRowFormula") => {
                let units = crate::string_value::decode(&text);
                let units = oxicells_calc::parser::transform_utf16_formula(&units, |formula| {
                    let mut formula = formula.to_string();
                    for rename in renames {
                        formula = oxicells_calc::map_table_references(&formula, &|name, asked| {
                            (oxicells_calc::case_fold::table_key(name) == oxicells_calc::case_fold::table_key(&rename.old)).then(|| match asked {
                                Some(asked) => format!("{}[{asked}]", rename.new), None => rename.new.clone(),
                            })
                        });
                        let inside = oxicells_calc::case_fold::table_key(&own) == oxicells_calc::case_fold::table_key(&rename.old)
                            || oxicells_calc::case_fold::table_key(&own) == oxicells_calc::case_fold::table_key(&rename.new);
                        for (old, new) in &rename.columns {
                            formula = oxicells_calc::rename_table_column(&formula, &rename.new, inside, old, new);
                        }
                    }
                    Ok(formula)
                }).map_err(XlsxError::InvalidData)?;
                writer.write_event(Event::Text(BytesText::new(&crate::string_value::encode(&units))))
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                active = false;
            }
            Event::Eof => break,
            _ => {}
        }
        writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }
    String::from_utf8(writer.into_inner().into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn patch_workbook_tables(data: &[u8], planned: &[Option<Vec<crate::ir::Table>>]) -> Result<Vec<u8>, XlsxError> {
    let mut archive = OoxmlArchive::new(data)?;
    let workbook = archive.read_part("xl/workbook.xml")?;
    let relationships = parse_relationships(&archive.read_part("xl/_rels/workbook.xml.rels")?)
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    let mut reader = Reader::from_str(&workbook);
    let mut paths = Vec::new();
    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(element) | Event::Empty(element) if local_name(element.name().as_ref()) == "sheet" => {
                let relation = get_attr(&element, "id").and_then(|id| relationships.get(&id))
                    .ok_or_else(|| XlsxError::InvalidData("worksheet relationship is missing".into()))?;
                paths.push(resolve_part_target("xl/workbook.xml", &relation.target));
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if paths.len() != planned.len() { return Err(XlsxError::InvalidData("table plan does not match worksheet order".into())); }
    let mut names: std::collections::HashSet<String> = archive.file_names().into_iter().collect();
    let mut next_id = 1u32;
    for path in names.iter().filter(|name| name.starts_with("xl/tables/") && name.ends_with(".xml")) {
        if let Some(id) = table_root_attribute(&archive.read_part(path)?, "id").and_then(|id| id.parse::<u32>().ok()) { next_id = next_id.max(id.saturating_add(1)); }
    }
    let mut replacements = HashMap::<String, Vec<u8>>::new();
    let mut removed = std::collections::HashSet::<String>::new();
    let mut types = archive.read_part("[Content_Types].xml")?;
    let mut formula_renames = Vec::new();
    for (sheet_path, plan) in paths.iter().zip(planned) {
        let Some(tables) = plan else { continue };
        let sheet = archive.read_part(sheet_path)?;
        let rels_path = sheet_rels_path(sheet_path);
        let rels_xml = archive.try_read_part(&rels_path)?.unwrap_or_else(|| "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>".into());
        let rels = parse_relationships(&rels_xml).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let mut existing = rels.iter().filter(|(_, relation)| relation.rel_type.ends_with("/table"))
            .map(|(rid, relation)| (rid.clone(), resolve_part_target(sheet_path, &relation.target))).collect::<Vec<_>>();
        existing.sort_by(|left, right| left.1.cmp(&right.1));
        if existing.is_empty() && tables.is_empty() { continue; }
        let mut parts = Vec::new();
        for (rid, path) in &existing { parts.push((rid.clone(), path.clone(), archive.read_part(path)?)); }
        let mut used = std::collections::HashSet::new();
        let mut table_parts = String::new();
        let mut new_relationships = String::new();
        let mut used_rids: std::collections::HashSet<String> = rels.keys().cloned().collect();
        for (index, table) in tables.iter().enumerate() {
            let reference = format!("{}{}:{}{}", col_to_letter(table.start_col), table.start_row, col_to_letter(table.end_col), table.end_row);
            let source = parts.iter().enumerate().find(|(at, (_, path, _))| !used.contains(at) && table.origin_part.as_deref() == Some(path.as_str())).map(|(at, _)| at)
                .or_else(|| parts.iter().enumerate().find(|(at, (_, _, xml))| !used.contains(at) && table_root_attribute(xml, "displayName").as_deref() == Some(table.name.as_str())).map(|(at, _)| at))
                .or_else(|| parts.iter().enumerate().find(|(at, (_, _, xml))| !used.contains(at) && table_root_attribute(xml, "ref").as_deref() == Some(reference.as_str())).map(|(at, _)| at))
                .or_else(|| (parts.len() == tables.len() && !used.contains(&index)).then_some(index));
            let (rid, path, original, id) = if let Some(at) = source {
                used.insert(at);
                let (rid, path, xml) = &parts[at];
                let id = table_root_attribute(xml, "id").and_then(|id| id.parse().ok()).unwrap_or_else(|| { let id = next_id; next_id += 1; id });
                (rid.clone(), path.clone(), Some(xml.as_str()), id)
            } else {
                let mut number = 1;
                let path = loop { let path = format!("xl/tables/table{number}.xml"); if names.insert(path.clone()) { break path; } number += 1; };
                let mut number = 1;
                let rid = loop { let rid = format!("rId{number}"); if used_rids.insert(rid.clone()) { break rid; } number += 1; };
                let id = next_id; next_id += 1;
                (rid, path, None, id)
            };
            let xml = table_xml(table, id, original)?;
            if let Some(original) = original {
                let old = table_root_attribute(original, "displayName").or_else(|| table_root_attribute(original, "name")).unwrap_or_default();
                let old_columns = direct_xml_child(original, "tableColumns")?.0
                    .map(|(start, end)| table_column_elements(&original[start..end])).transpose()?.unwrap_or_default();
                let mut columns = Vec::new();
                for (at, column) in old_columns.iter().enumerate() {
                    if let (Some(old), Some(new)) = (table_root_attribute(column, "name"), table.columns.get(at)) {
                        if old != *new && !table.columns.contains(&old) { columns.push((old, new.clone())); }
                    }
                }
                if old != table.name || !columns.is_empty() { formula_renames.push(TableFormulaRename { old, new: table.name.clone(), columns }); }
            }
            replacements.insert(path.clone(), xml.into_bytes());
            types = append_content_type(&types, &path, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")?;
            table_parts.push_str(&format!("<tablePart r:id=\"{}\"/>", escape(&rid)));
            let target = format!("../tables/{}", path.rsplit('/').next().unwrap_or_default());
            new_relationships.push_str(&format!("<Relationship xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\" Id=\"{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"{}\"/>", escape(&rid), escape(&target)));
        }
        for (at, (_, path, _)) in parts.iter().enumerate() { if !used.contains(&at) { removed.insert(path.clone()); } }
        let mut reader = Reader::from_str(&rels_xml);
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        loop {
            let event = reader.read_event().map_err(XlsxError::Xml)?;
            match &event {
                Event::Start(element) | Event::Empty(element) if local_name(element.name().as_ref()) == "Relationship" && get_attr(element, "Type").is_some_and(|kind| kind.ends_with("/table")) => {
                    if matches!(event, Event::Start(_)) { reader.read_to_end(element.name()).map_err(XlsxError::Xml)?; }
                    continue;
                }
                Event::Empty(element) if local_name(element.name().as_ref()) == "Relationships" => {
                    writer.write_event(Event::Start(element.clone())).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    writer.get_mut().write_all(new_relationships.as_bytes()).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    writer.write_event(Event::End(BytesEnd::new(String::from_utf8_lossy(element.name().as_ref()).to_string()))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    continue;
                }
                Event::End(element) if local_name(element.name().as_ref()) == "Relationships" =>
                    writer.get_mut().write_all(new_relationships.as_bytes()).map_err(|error| XlsxError::InvalidData(error.to_string()))?,
                Event::Eof => break,
                _ => {}
            }
            writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        }
        replacements.insert(rels_path, writer.into_inner().into_inner());
        let sheet = if tables.is_empty() { replace_direct_xml_child(&sheet, "tableParts", None)? } else {
            set_worksheet_drawing_child(&sheet, "tableParts", &format!("<tableParts xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" count=\"{}\">{table_parts}</tableParts>", tables.len()), &["extLst"])?
        };
        replacements.insert(sheet_path.clone(), sheet.into_bytes());
    }
    // A calculated column can reference a table on another sheet. Update
    // every table definition, including ones whose own properties did not change.
    if !formula_renames.is_empty() {
        for path in names.iter().filter(|path| path.starts_with("xl/tables/") && path.ends_with(".xml") && !removed.contains(*path)) {
            let xml = if let Some(bytes) = replacements.get(path) {
                String::from_utf8(bytes.clone()).map_err(|error| XlsxError::InvalidData(error.to_string()))?
            } else { archive.read_part(path)? };
            let rewritten = rewrite_table_formula_xml(&xml, &formula_renames)?;
            if rewritten != xml { replacements.insert(path.clone(), rewritten.into_bytes()); }
        }
    }
    // Unlisted/deleted tables have no remaining worksheet relationship.
    for path in &removed {
        let mut reader = Reader::from_str(&types);
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        loop {
            let event = reader.read_event().map_err(XlsxError::Xml)?;
            if matches!(&event, Event::Empty(element) if local_name(element.name().as_ref()) == "Override" && get_attr(element, "PartName").as_deref() == Some(&format!("/{path}"))) { continue; }
            if matches!(event, Event::Eof) { break; }
            writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        }
        types = String::from_utf8(writer.into_inner().into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }
    replacements.insert("[Content_Types].xml".into(), types.into_bytes());
    let mut input = ZipArchive::new(Cursor::new(data)).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    let mut output = ZipWriter::new(Cursor::new(Vec::new()));
    for at in 0..input.len() {
        let entry = input.by_index(at).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let path = entry.name().to_string();
        if removed.contains(&path) { continue; }
        if let Some(bytes) = replacements.remove(&path) {
            output.start_file(&path, SimpleFileOptions::default()).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            output.write_all(&bytes).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        } else { output.raw_copy_file(entry).map_err(|error| XlsxError::InvalidData(error.to_string()))?; }
    }
    let mut additions: Vec<_> = replacements.into_iter().collect();
    additions.sort_by(|left, right| left.0.cmp(&right.0));
    for (path, bytes) in additions {
        output.start_file(path, SimpleFileOptions::default()).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        output.write_all(&bytes).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }
    Ok(output.finish().map_err(|error| XlsxError::InvalidData(error.to_string()))?.into_inner())
}

impl CellEditValue {
    fn error_text(&self) -> Option<&str> {
        match self {
            Self::Error(error) | Self::CachedFormula { value: crate::ir::CellValue::Error(error), .. }
                | Self::CachedUtf16Formula { value: crate::ir::CellValue::Error(error), .. } => Some(error),
            _ => None,
        }
    }

    fn cell_type(&self) -> Option<&'static str> {
        match self {
            Self::String(_) | Self::Utf16(_) | Self::RichText { .. } => Some("str"),
            Self::Boolean(_) => Some("b"),
            Self::Error(_) => Some("e"),
            Self::CachedFormula { value, .. } | Self::CachedUtf16Formula { value, .. } => match value {
                crate::ir::CellValue::String(_) | crate::ir::CellValue::Utf16(_) => Some("str"),
                crate::ir::CellValue::Boolean(_) => Some("b"),
                crate::ir::CellValue::Error(_) => Some("e"),
                _ => None,
            },
            Self::Number(_) | Self::Formula(_) | Self::ArrayFormula { .. } | Self::Empty => None,
        }
    }

    fn value_text(&self) -> Result<Option<String>, XlsxError> {
        match self {
            Self::String(value) => Ok(Some(crate::string_value::encode(&value.encode_utf16().collect::<Vec<_>>()))),
            Self::Utf16(units) => Ok(Some(crate::string_value::encode(units))),
            Self::RichText { value, .. } => value.string_units()
                .map(|units|Some(crate::string_value::encode(&units)))
                .ok_or_else(||XlsxError::InvalidData("rich text needs a text cell value".into())),
            Self::Number(value) if value.is_finite() => Ok(Some(value.to_string())),
            Self::Number(_) => Err(XlsxError::InvalidData(
                "spreadsheet numbers must be finite".to_string(),
            )),
            Self::Boolean(value) => Ok(Some(if *value { "1" } else { "0" }.to_string())),
            Self::Error(value) => Ok(Some(value.clone())),
            Self::CachedFormula { value, .. } | Self::CachedUtf16Formula { value, .. } => match value {
                crate::ir::CellValue::Empty => Ok(None),
                crate::ir::CellValue::Number(value) => Self::Number(*value).value_text(),
                crate::ir::CellValue::String(value) => Self::String(value.clone()).value_text(),
                crate::ir::CellValue::Utf16(units) => Self::Utf16(units.clone()).value_text(),
                crate::ir::CellValue::Error(value) => Ok(Some(value.clone())),
                crate::ir::CellValue::Boolean(value) => Self::Boolean(*value).value_text(),
            },
            Self::Formula(_) | Self::ArrayFormula { .. } | Self::Empty => Ok(None),
        }
    }

    fn formula_units(&self) -> Option<std::borrow::Cow<'_, [u16]>> {
        match self {
            Self::Formula(value) | Self::ArrayFormula { formula: value, .. }
                | Self::CachedFormula { formula: value, .. } => {
                Some(value.strip_prefix('=').unwrap_or(value).encode_utf16().collect::<Vec<_>>().into())
            }
            Self::CachedUtf16Formula { formula, .. } => Some(formula.strip_prefix(&[b'=' as u16]).unwrap_or(formula).into()),
            _ => None,
        }
    }

    /// The block an array formula is dealt across, for the `<f>` element.
    fn array_reach(&self) -> Option<&str> {
        match self {
            Self::ArrayFormula { reach, .. } => Some(reach),
            Self::CachedFormula { reach, .. } | Self::CachedUtf16Formula { reach, .. } => reach.as_deref(),
            _ => None,
        }
    }
}

fn resolve_drawing_parts(
    archive: &mut OoxmlArchive,
    sheet_paths: &[String],
) -> Result<HashMap<usize, String>, XlsxError> {
    let mut parts = HashMap::new();
    for (index, sheet_path) in sheet_paths.iter().enumerate() {
        if sheet_path.is_empty() { continue; }
        let Some((directory, filename)) = sheet_path.rsplit_once('/') else { continue; };
        let rels_path = format!("{directory}/_rels/{filename}.rels");
        let Ok(xml) = archive.read_part(&rels_path) else { continue; };
        let relationships = parse_relationships(&xml).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        if let Some(relation) = relationships.values().find(|relation| relation.rel_type.ends_with("/drawing")) {
            parts.insert(index, resolve_part_target(sheet_path, &relation.target));
        }
    }
    Ok(parts)
}

fn drawing_position_changed(before: &crate::ir::Drawing, after: &crate::ir::Drawing) -> bool {
    if before.ole_placement != after.ole_placement
        || before.ole_auto_load != after.ole_auto_load
        || before.ole_enabled != after.ole_enabled
        || before.ole_locked != after.ole_locked
        || before.ole_print_object != after.ole_print_object
        || before.ole_shadow != after.ole_shadow
        || before.ole_on_action != after.ole_on_action
        || before.ole_border_color != after.ole_border_color
        || before.ole_border_weight != after.ole_border_weight
        || before.ole_interior_color != after.ole_interior_color
    { return true; }
    let anchor_changed = |left: crate::ir::Anchor, right: crate::ir::Anchor| {
        left.col != right.col || left.row != right.row
            || left.col_off.abs_diff(right.col_off) > 2
            || left.row_off.abs_diff(right.row_off) > 2
    };
    if anchor_changed(before.from, after.from) { return true; }
    match before.to {
        Some(before_to) => after.to.is_none_or(|after_to| anchor_changed(before_to, after_to)),
        None => match (before.extent, after.extent) {
            (Some(before_extent), Some(after_extent)) => before_extent.0.abs_diff(after_extent.0) > 2 || before_extent.1.abs_diff(after_extent.1) > 2,
            (None, None) => false,
            _ => true,
        },
    }
}

fn shape_text_formatting_changed(before: &crate::ir::ShapeText, after: &crate::ir::ShapeText) -> bool {
    let paragraph_runs = |paragraph: &crate::ir::ShapeParagraph| {
        if paragraph.runs.is_empty() {
            vec![(paragraph.text.clone(), paragraph.bold, paragraph.italic, false,
                paragraph.face.clone(), paragraph.size, paragraph.color.as_deref().unwrap_or("000000").to_ascii_uppercase(), None, None, None)]
        } else {
            paragraph.runs.iter().map(|run| (
                run.text.clone(),
                run.bold,
                run.italic.unwrap_or(paragraph.italic),
                run.underline,
                run.face.clone().or_else(|| paragraph.face.clone()),
                run.size.unwrap_or(paragraph.size),
                run.color.as_deref().or(paragraph.color.as_deref()).unwrap_or("000000").to_ascii_uppercase(),
                run.color_theme,
                run.color_brightness,
                run.legacy_color_index,
            )).collect()
        }
    };
    before.anchor != after.anchor || before.insets != after.insets || before.wrap != after.wrap || before.clip != after.clip
        || before.paragraphs.len() != after.paragraphs.len()
        || before.paragraphs.iter().zip(&after.paragraphs).any(|(left, right)| {
            left.text != right.text || left.align != right.align || left.line_pitch != right.line_pitch
                || left.line_scale != right.line_scale || paragraph_runs(left) != paragraph_runs(right)
        })
}

#[derive(Default)]
struct DrawingAdditionPackage {
    parts: HashMap<String, String>,
    ole_anchors: HashMap<String, HashMap<u32, crate::ir::Drawing>>,
    worksheet_ole_objects: HashMap<String, Vec<String>>,
    relationship_xml: HashMap<String, String>,
    new_parts: Vec<(String, String)>,
    new_binary_parts: Vec<(String, Vec<u8>)>,
    worksheet_drawing: HashMap<String, String>,
    worksheet_legacy_drawing: HashMap<String, String>,
    content_types: Vec<(String, String)>,
}

#[derive(Debug, Clone)]
struct ShapeTextEdit {
    before: crate::ir::ShapeText,
    after: crate::ir::ShapeText,
}

#[derive(Debug, Clone)]
struct PictureStyleEdit {
    before: crate::ir::PictureStyle,
    after: crate::ir::PictureStyle,
}

#[derive(Debug, Clone)]
struct DrawingGroupAddition {
    name: String,
    from: crate::ir::Anchor,
    to: crate::ir::Anchor,
    transform: crate::ir::GroupTransform,
    members: Vec<(String, (i64, i64, i64, i64))>,
}

fn drawing_addition_package(
    archive: &mut OoxmlArchive,
    sheet_paths: &[String],
    existing_parts: &HashMap<usize, String>,
    additions: &HashMap<usize, Vec<crate::ir::Drawing>>,
    group_additions: &HashMap<usize, Vec<DrawingGroupAddition>>,
    line_edits: &HashMap<(usize, String), Option<crate::ir::ShapeLine>>,
    fill_edits: &HashMap<(usize, String), (Option<String>, Option<usize>, Option<u8>, Option<f32>)>,
    text_edits: &HashMap<(usize, String), ShapeTextEdit>,
    picture_edits: &HashMap<(usize, String), PictureStyleEdit>,
    geometry_edits: &HashMap<(usize, String), crate::ir::Drawing>,
    drawing_order_edits: &HashMap<usize, Vec<String>>,
    drawing_states: &HashMap<(usize, String), crate::ir::Drawing>,
    deletions: &HashMap<usize, std::collections::HashSet<String>>,
) -> Result<DrawingAdditionPackage, XlsxError> {
    let mut plan = DrawingAdditionPackage::default();
    let mut names: std::collections::HashSet<String> = archive.file_names().into_iter().collect();
    let mut part_additions: HashMap<String, Vec<crate::ir::Drawing>> = HashMap::new();
    let mut ole_additions: HashMap<usize, Vec<crate::ir::Drawing>> = HashMap::new();
    let mut fresh_ole_additions: HashMap<usize, Vec<crate::ir::Drawing>> = HashMap::new();
    let mut fresh_ole_parts: HashMap<usize, String> = HashMap::new();
    for (sheet, drawings) in additions {
        let ole: Vec<_> = drawings.iter().filter(|drawing| drawing.ole_source_name.is_some()).cloned().collect();
        let fresh_ole: Vec<_> = drawings.iter().filter(|drawing| drawing.ole_source_name.is_none()
            && drawing.ole_prog_id.is_some() && drawing.ole_embedded_package.is_some()).cloned().collect();
        let regular: Vec<_> = drawings.iter().filter(|drawing| drawing.ole_source_name.is_none()
            && !(drawing.ole_prog_id.is_some() && drawing.ole_embedded_package.is_some())).cloned().collect();
        if !ole.is_empty() {
            if !existing_parts.contains_key(sheet) {
                return Err(XlsxError::InvalidData("an OLE duplicate has no source drawing part".into()));
            }
            ole_additions.entry(*sheet).or_default().extend(ole);
        }
        if !fresh_ole.is_empty() { fresh_ole_additions.insert(*sheet, fresh_ole.clone()); }
        let part = if let Some(part) = existing_parts.get(sheet) {
            part.clone()
        } else {
            if regular.is_empty() && fresh_ole.is_empty() { continue; }
            let sheet_path = sheet_paths.get(*sheet).filter(|path| !path.is_empty()).ok_or_else(|| XlsxError::InvalidData(format!("sheet {sheet} is not in the source workbook while adding a shape")))?;
            let drawing_path = unique_part_path(&names, "xl/drawings", "drawing", ".xml");
            names.insert(drawing_path.clone());
            let mut drawing_xml = String::from("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"></xdr:wsDr>");
            drawing_xml = append_drawing_shapes_xml(&drawing_xml, &regular)?;
            plan.new_parts.push((drawing_path.clone(), drawing_xml));
            plan.content_types.push((drawing_path.clone(), "application/vnd.openxmlformats-officedocument.drawing+xml".into()));

            let rels_path = sheet_rels_path(sheet_path);
            let original_rels = archive.try_read_part(&rels_path)?;
            let original_ids = original_rels.as_deref().map(parse_relationships).transpose()
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?.unwrap_or_default();
            let target = relative_sheet_target(sheet_path, &drawing_path);
            let updated_rels = append_relationship(
                original_rels.as_deref(),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                &target,
            )?;
            let parsed_updated = parse_relationships(&updated_rels)
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            let relationship_id = parsed_updated.keys().find(|id| !original_ids.contains_key(*id)).cloned()
                .ok_or_else(|| XlsxError::InvalidData("new drawing relationship id was not allocated".into()))?;
            if original_rels.is_some() {
                plan.relationship_xml.insert(rels_path, updated_rels);
            } else {
                plan.new_parts.push((rels_path, updated_rels));
            }
            plan.worksheet_drawing.insert(sheet_path.clone(), relationship_id);
            fresh_ole_parts.insert(*sheet, drawing_path.clone());
            drawing_path
        };
        if existing_parts.contains_key(sheet) {
            part_additions.entry(part.clone()).or_default().extend(regular);
            if !fresh_ole.is_empty() { fresh_ole_parts.insert(*sheet, part); }
        }
    }

    for (sheet, drawings) in fresh_ole_additions {
        let part = fresh_ole_parts.get(&sheet).ok_or_else(|| XlsxError::InvalidData("new OLE object has no drawing part".into()))?.clone();
        let sheet_path = sheet_paths.get(sheet).filter(|path| !path.is_empty())
            .ok_or_else(|| XlsxError::InvalidData(format!("sheet {sheet} is not in the source workbook while adding an OLE object")))?;
        let mut xml = plan.new_parts.iter().find(|(path, _)| path == &part).map(|(_, xml)| xml.clone())
            .or_else(|| plan.parts.get(&part).cloned());
        if xml.is_none() { xml = Some(archive.read_part(&part)?); }
        let mut xml = xml.expect("drawing XML loaded");
        let rels_path = sheet_rels_path(sheet_path);
        let original_rels = archive.try_read_part(&rels_path)?;
        let mut worksheet_rels = plan.relationship_xml.get(&rels_path).cloned()
            .or_else(|| plan.new_parts.iter().find(|(path, _)| path == &rels_path).map(|(_, xml)| xml.clone()))
            .or(original_rels.clone()).unwrap_or_else(|| r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.into());
        let mut worksheet_objects = Vec::new();
        for drawing in drawings {
            let (shape_id, icon_relationship_id) = append_new_ole_vml(
                archive, sheet_path, &drawing, &xml, &mut worksheet_rels, &mut names, &mut plan,
            )?;
            let word_document = drawing.ole_prog_id.as_deref().is_some_and(|prog_id| {
                prog_id.eq_ignore_ascii_case("Word.Document.12")
            });
            let embedding = unique_embedding_path(
                &names,
                if word_document {
                    "xl/embeddings/Microsoft_Word_Document.docx"
                } else {
                    "xl/embeddings/oleObject.bin"
                },
            );
            names.insert(embedding.clone());
            let package = drawing.ole_embedded_package.as_ref()
                .ok_or_else(|| XlsxError::InvalidData("new OLE object has no embedded package".into()))?;
            plan.new_binary_parts.push((embedding.clone(), package.clone()));
            plan.content_types.push((embedding.clone(), if word_document {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document".into()
            } else {
                "application/vnd.openxmlformats-officedocument.oleObject".into()
            }));
            let target = relative_sheet_target(sheet_path, &embedding);
            let before: std::collections::HashSet<String> = parse_relationships(&worksheet_rels)
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?.into_keys().collect();
            worksheet_rels = append_relationship(Some(&worksheet_rels),
                if word_document {
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
                } else {
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
                }, &target)?;
            let after = parse_relationships(&worksheet_rels).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            let relationship_id = after.keys().find(|id| !before.contains(*id)).cloned()
                .ok_or_else(|| XlsxError::InvalidData("OLE relationship id was not allocated".into()))?;
            xml = append_xml_child(&xml, "wsDr", &write_ole_drawing_shape(&drawing, shape_id)?)?;
            worksheet_objects.push(write_worksheet_ole_object(&drawing, shape_id, &relationship_id, &icon_relationship_id)?);
        }
        if let Some((_, contents)) = plan.new_parts.iter_mut().find(|(path, _)| path == &part) {
            *contents = xml;
        } else {
            plan.parts.insert(part, xml);
        }
        plan.worksheet_ole_objects.entry(sheet_path.clone()).or_default().extend(worksheet_objects);
        if original_rels.is_some() {
            plan.relationship_xml.insert(rels_path, worksheet_rels);
        } else if let Some((_, contents)) = plan.new_parts.iter_mut().find(|(path, _)| path == &rels_path) {
            *contents = worksheet_rels;
        } else {
            plan.new_parts.push((rels_path, worksheet_rels));
        }
    }

    let mut edits_by_part: HashMap<String, HashMap<String, Option<crate::ir::ShapeLine>>> = HashMap::new();
    for ((sheet, shape_name), after) in line_edits {
        let part = existing_parts.get(sheet).ok_or_else(|| XlsxError::InvalidData(format!("drawing part for sheet {sheet} was not found while saving shape {shape_name:?}")))?;
        edits_by_part.entry(part.clone()).or_default().insert(shape_name.clone(), after.clone());
    }
    let mut fills_by_part: HashMap<String, HashMap<String, (Option<String>, Option<usize>, Option<u8>, Option<f32>)>> = HashMap::new();
    for ((sheet, shape_name), fill) in fill_edits {
        let part = existing_parts.get(sheet).ok_or_else(|| XlsxError::InvalidData(format!("drawing part for sheet {sheet} was not found while saving shape {shape_name:?} fill")))?;
        fills_by_part.entry(part.clone()).or_default().insert(shape_name.clone(), fill.clone());
    }
    let mut text_by_part: HashMap<String, HashMap<String, ShapeTextEdit>> = HashMap::new();
    for ((sheet, shape_name), text) in text_edits {
        let part = existing_parts.get(sheet).ok_or_else(|| XlsxError::InvalidData(format!("drawing part for sheet {sheet} was not found while saving shape {shape_name:?} text")))?;
        text_by_part.entry(part.clone()).or_default().insert(shape_name.clone(), text.clone());
    }
    let mut pictures_by_part: HashMap<String, HashMap<String, PictureStyleEdit>> = HashMap::new();
    for ((sheet, shape_name), format) in picture_edits {
        let part = existing_parts.get(sheet).ok_or_else(|| XlsxError::InvalidData(format!("drawing part for sheet {sheet} was not found while saving picture {shape_name:?}")))?;
        pictures_by_part.entry(part.clone()).or_default().insert(shape_name.clone(), format.clone());
    }
    let mut geometry_by_part: HashMap<String, HashMap<String, crate::ir::Drawing>> = HashMap::new();
    for ((sheet, shape_name), after) in geometry_edits {
        let part = existing_parts.get(sheet).ok_or_else(|| XlsxError::InvalidData(format!("drawing part for sheet {sheet} was not found while moving shape {shape_name:?}")))?;
        geometry_by_part.entry(part.clone()).or_default().insert(shape_name.clone(), after.clone());
    }
    let mut drawings_by_part: HashMap<String, HashMap<String, crate::ir::Drawing>> = HashMap::new();
    for ((sheet, shape_name), drawing) in drawing_states {
        let part = existing_parts.get(sheet).ok_or_else(|| XlsxError::InvalidData(format!("drawing part for sheet {sheet} was not found while saving shape {shape_name:?}")))?;
        drawings_by_part.entry(part.clone()).or_default().insert(shape_name.clone(), drawing.clone());
    }
    let mut deletions_by_part: HashMap<String, std::collections::HashSet<String>> = HashMap::new();
    for (sheet, shape_names) in deletions {
        let part = existing_parts.get(sheet).ok_or_else(|| XlsxError::InvalidData(format!("drawing part for sheet {sheet} was not found while deleting a shape")))?;
        deletions_by_part.entry(part.clone()).or_default().extend(shape_names.iter().cloned());
    }
    let ole_parts: HashMap<String, usize> = ole_additions.iter().filter_map(|(sheet, _)| {
        existing_parts.get(sheet).map(|part| (part.clone(), *sheet))
    }).collect();
    let groups_by_part: HashMap<String, Vec<DrawingGroupAddition>> = group_additions.iter().filter_map(|(sheet, groups)| existing_parts.get(sheet).map(|part| (part.clone(), groups.clone()))).collect();
    for part in part_additions.keys().chain(edits_by_part.keys()).chain(fills_by_part.keys()).chain(text_by_part.keys()).chain(pictures_by_part.keys()).chain(geometry_by_part.keys()).chain(deletions_by_part.keys()).chain(ole_parts.keys()).chain(groups_by_part.keys()).chain(fresh_ole_parts.values()).cloned().collect::<std::collections::HashSet<_>>() {
        let mut xml = plan.parts.remove(&part)
            .or_else(|| plan.new_parts.iter().find(|(path, _)| path == &part).map(|(_, xml)| xml.clone()));
        if xml.is_none() { xml = Some(archive.read_part(&part)?); }
        let mut xml = xml.expect("drawing XML loaded");
        if let Some(edits) = edits_by_part.get(&part) { xml = patch_drawing_line_xml(&xml, edits)?; }
        if let Some(edits) = fills_by_part.get(&part) { xml = patch_drawing_fill_xml(&xml, edits)?; }
        if let Some(edits) = text_by_part.get(&part) { xml = patch_drawing_text_xml(&xml, edits)?; }
        if let Some(edits) = pictures_by_part.get(&part) {
            xml = patch_drawing_picture_xml(&xml, edits)?;
            xml = patch_drawing_picture_background_xml(&xml, edits)?;
        }
        if let Some(names) = deletions_by_part.get(&part) {
            xml = patch_drawing_deletions_xml(
                &xml,
                names,
                geometry_by_part.get(&part),
                drawings_by_part.get(&part),
            )?;
        }
        if let Some(edits) = geometry_by_part.get(&part) { xml = patch_drawing_geometry_xml(&xml, edits)?; }
        if let Some(drawings) = part_additions.get(&part) { xml = append_drawing_shapes_xml(&xml, drawings)?; }
        if let Some(groups) = groups_by_part.get(&part) {
            for group in groups { xml = patch_drawing_group_addition(&xml, group)?; }
        }

        if let Some(sheet) = ole_parts.get(&part).copied() {
            let sheet_path = sheet_paths.get(sheet).filter(|path| !path.is_empty())
                .ok_or_else(|| XlsxError::InvalidData("OLE duplicate worksheet part is missing".into()))?;
            let source_drawing_xml = archive.read_part(&part)?;
            let source_sheet_xml = archive.read_part(sheet_path)?;
            let rels_path = sheet_rels_path(sheet_path);
            let source_rels = archive.try_read_part(&rels_path)?;
            let mut relationship_xml = source_rels.clone().ok_or_else(|| XlsxError::InvalidData("OLE duplicate worksheet relationships are missing".into()))?;
            let source_relationships = parse_relationships(&relationship_xml)
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            let vml_path = source_relationships.values().find(|relation| relation.rel_type.ends_with("/vmlDrawing"))
                .map(|relation| resolve_part_target(sheet_path, &relation.target));
            let mut vml_xml = vml_path.as_deref().map(|path| archive.read_part(path)).transpose()?;
            let content_types = archive.read_part("[Content_Types].xml")?;
            let ole_copies = ole_additions.get(&sheet).expect("sheet collected from OLE additions");
            let mut appended_objects = Vec::new();
            for drawing in ole_copies {
                let source_name = drawing.ole_source_name.as_deref()
                    .ok_or_else(|| XlsxError::InvalidData("OLE duplicate has no source name".into()))?;
                let source_shape_id = drawing_shape_id(&source_drawing_xml, source_name)?
                    .ok_or_else(|| XlsxError::InvalidData(format!("OLE source drawing {source_name:?} has no shape id")))?;
                let source_ole = find_ole_object_xml(&source_sheet_xml, source_shape_id)?
                    .ok_or_else(|| XlsxError::InvalidData(format!("OLE source {source_name:?} has no worksheet object")))?;
                let source_rel_id = xml_attribute_value(&source_ole, "id")
                    .ok_or_else(|| XlsxError::InvalidData("OLE worksheet object has no relationship id".into()))?;
                let source_relation = source_relationships.get(&source_rel_id)
                    .ok_or_else(|| XlsxError::InvalidData(format!("OLE relationship {source_rel_id:?} is missing")))?;
                let source_target = resolve_part_target(sheet_path, &source_relation.target);
                let source_bytes = if let Some(package) = drawing.ole_embedded_package.as_ref() {
                    package.clone()
                } else {
                    archive.read_binary_part(&source_target)?
                };
                let new_target = unique_embedding_path(&names, &source_target);
                names.insert(new_target.clone());
                plan.new_binary_parts.push((new_target.clone(), source_bytes));
                let relationship_target = relative_sheet_target(sheet_path, &new_target);
                let before_ids: std::collections::HashSet<String> = parse_relationships(&relationship_xml)
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?.into_keys().collect();
                relationship_xml = append_relationship(Some(&relationship_xml), &source_relation.rel_type, &relationship_target)?;
                let after_relationships = parse_relationships(&relationship_xml)
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                let new_rel_id = after_relationships.keys().find(|id| !before_ids.contains(*id)).cloned()
                    .ok_or_else(|| XlsxError::InvalidData("OLE duplicate relationship id was not allocated".into()))?;

                let new_shape_id = max_drawing_shape_id(&xml)?.checked_add(1)
                    .ok_or_else(|| XlsxError::InvalidData("drawing shape id overflow".into()))?;
                let mut anchor = clone_drawing_anchor_xml(&source_drawing_xml, source_name, drawing.name.as_deref().unwrap_or("Object"), source_shape_id, new_shape_id)?;
                xml = append_xml_child(&xml, "wsDr", &anchor)?;
                let name = drawing.name.as_deref().unwrap_or("Object");
                xml = patch_drawing_geometry_xml(&xml, &HashMap::from([(name.to_string(), drawing.clone())]))?;
                anchor.clear();

                let object_xml = clone_ole_object_xml(&source_ole, source_shape_id, new_shape_id, &new_rel_id)?;
                appended_objects.push(object_xml);
                plan.ole_anchors.entry(sheet_path.clone()).or_default().insert(new_shape_id, drawing.clone());
                if let Some(vml_xml) = vml_xml.as_mut() {
                    let z_index = max_vml_z_index(vml_xml)?.checked_add(1)
                        .ok_or_else(|| XlsxError::InvalidData("VML z-index overflow".into()))?;
                    let vml_shape = clone_vml_ole_shape_xml(vml_xml, source_shape_id, new_shape_id, z_index, drawing)?;
                    *vml_xml = append_xml_child(vml_xml, "xml", &vml_shape)?;
                }

                let content_type = content_type_for_part(&content_types, &source_target)?;
                if source_target.rsplit_once('.').is_some_and(|(_, extension)| extension.eq_ignore_ascii_case("bin")) {
                    plan.content_types.push((new_target, content_type));
                }
            }
            plan.relationship_xml.insert(rels_path, relationship_xml);
            plan.worksheet_ole_objects.entry(sheet_path.clone()).or_default().extend(appended_objects);
            if let (Some(path), Some(xml)) = (vml_path, vml_xml) {
                plan.parts.insert(path, xml);
            }
        }
        if let Some((_, contents)) = plan.new_parts.iter_mut().find(|(path, _)| path == &part) {
            *contents = xml;
        } else {
            plan.parts.insert(part, xml);
        }
    }
    // Excel stores OLE positions twice: in DrawingML and in worksheet
    // oleObject/objectPr anchors. Keep the latter in sync when a shape moves.
    let mut ole_edits_by_sheet: HashMap<String, HashMap<u32, crate::ir::Drawing>> = HashMap::new();
    for ((sheet, shape_name), drawing) in geometry_edits {
        let drawing_part = existing_parts.get(sheet).ok_or_else(|| XlsxError::InvalidData(format!("drawing part for sheet {sheet} was not found while saving OLE position")))?;
        let drawing_xml = archive.read_part(drawing_part)?;
        let shape_id = drawing_shape_id(&drawing_xml, shape_name)?.ok_or_else(|| XlsxError::InvalidData(format!("drawing shape {shape_name:?} has no id while saving OLE position")))?;
        let sheet_path = sheet_paths.get(*sheet).filter(|path| !path.is_empty()).ok_or_else(|| XlsxError::InvalidData(format!("sheet {sheet} is not in the source workbook while saving OLE position")))?;
        ole_edits_by_sheet.entry(sheet_path.clone()).or_default().insert(shape_id, drawing.clone());
    }
    for (sheet_path, edits) in ole_edits_by_sheet {
        plan.ole_anchors.entry(sheet_path).or_default().extend(edits);
    }
    let mut vml_shadow_edits: HashMap<String, HashMap<u32, bool>> = HashMap::new();
    for ((sheet, shape_name), drawing) in geometry_edits {
        let Some(shadow) = drawing.ole_shadow else { continue };
        let drawing_part = existing_parts.get(sheet).ok_or_else(|| XlsxError::InvalidData(format!("drawing part for sheet {sheet} was not found while saving OLE shadow")))?;
        let drawing_xml = archive.read_part(drawing_part)?;
        let Some(shape_id) = drawing_shape_id(&drawing_xml, shape_name)? else { continue };
        let sheet_path = sheet_paths.get(*sheet).filter(|path| !path.is_empty()).ok_or_else(|| XlsxError::InvalidData(format!("sheet {sheet} is not in the source workbook while saving OLE shadow")))?;
        let rels_path = sheet_rels_path(sheet_path);
        let Some(rels_xml) = archive.try_read_part(&rels_path)? else { continue };
        let relationships = parse_relationships(&rels_xml).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let Some(vml_relation) = relationships.values().find(|relation| relation.rel_type.ends_with("/vmlDrawing")) else { continue };
        let vml_path = resolve_part_target(sheet_path, &vml_relation.target);
        vml_shadow_edits.entry(vml_path).or_default().insert(shape_id, shadow);
    }
    for (vml_path, edits) in vml_shadow_edits {
        let xml = plan.parts.get(&vml_path).cloned().unwrap_or(archive.read_part(&vml_path)?);
        plan.parts.insert(vml_path, patch_vml_ole_shadows(&xml, &edits)?);
    }
    let mut vml_style_edits: HashMap<String, HashMap<u32, crate::ir::Drawing>> = HashMap::new();
    for ((sheet, shape_name), drawing) in geometry_edits {
        if drawing.ole_border_color.is_none() && drawing.ole_border_weight.is_none() && drawing.ole_interior_color.is_none() { continue; }
        let drawing_part = existing_parts.get(sheet).ok_or_else(|| XlsxError::InvalidData(format!("drawing part for sheet {sheet} was not found while saving OLE style")))?;
        let drawing_xml = archive.read_part(drawing_part)?;
        let Some(shape_id) = drawing_shape_id(&drawing_xml, shape_name)? else { continue };
        let sheet_path = sheet_paths.get(*sheet).filter(|path| !path.is_empty()).ok_or_else(|| XlsxError::InvalidData(format!("sheet {sheet} is not in the source workbook while saving OLE style")))?;
        let rels_path = sheet_rels_path(sheet_path);
        let Some(rels_xml) = archive.try_read_part(&rels_path)? else { continue };
        let relationships = parse_relationships(&rels_xml).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let Some(vml_relation) = relationships.values().find(|relation| relation.rel_type.ends_with("/vmlDrawing")) else { continue };
        let vml_path = resolve_part_target(sheet_path, &vml_relation.target);
        vml_style_edits.entry(vml_path).or_default().insert(shape_id, drawing.clone());
    }
    for (vml_path, edits) in vml_style_edits {
        let xml = plan.parts.get(&vml_path).cloned().unwrap_or(archive.read_part(&vml_path)?);
        plan.parts.insert(vml_path, patch_vml_ole_styles(&xml, &edits)?);
    }
    for (sheet, order) in drawing_order_edits {
        let Some(part) = existing_parts.get(sheet) else { continue };
        let xml = plan.parts.remove(part).map_or_else(|| archive.read_part(part), Ok)?;
        plan.parts.insert(part.clone(), patch_drawing_order(&xml, order)?);
    }
    Ok(plan)
}

fn patch_drawing_order(xml: &str, order: &[String]) -> Result<String, XlsxError> {
    let rank: HashMap<&str, usize> = order.iter().enumerate().map(|(index, name)| (name.as_str(), index)).collect();
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut in_root = false;
    let mut children: Vec<(usize, usize, Vec<Event<'static>>)> = Vec::new();
    let mut sequence = 0usize;
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?.into_owned();
        match event {
            Event::Start(start) if !in_root => {
                in_root = true;
                writer.write_event(Event::Start(start)).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            Event::Start(start) => {
                let mut node = vec![Event::Start(start)];
                let mut depth = 1usize;
                while depth > 0 {
                    let next = reader.read_event().map_err(XlsxError::Xml)?.into_owned();
                    match &next {
                        Event::Start(_) => depth += 1,
                        Event::End(_) => depth -= 1,
                        Event::Eof => return Err(XlsxError::InvalidData("drawing anchor ended unexpectedly".into())),
                        _ => {}
                    }
                    node.push(next);
                }
                let name = node.iter().find_map(|event| match event {
                    Event::Start(element) | Event::Empty(element) if local_name(element.name().as_ref()) == "cNvPr" => get_attr(element, "name"),
                    _ => None,
                });
                let is_anchor = node.iter().any(|event| match event {
                    Event::Start(element) | Event::Empty(element) => matches!(local_name(element.name().as_ref()).as_str(), "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" | "relSizeAnchor" | "absSizeAnchor"),
                    _ => false,
                });
                let priority = if is_anchor {
                    name.as_deref().and_then(|name| rank.get(name).copied()).unwrap_or(order.len() + sequence)
                } else {
                    usize::MAX / 2 + sequence
                };
                children.push((priority, sequence, node));
                sequence += 1;
            }
            Event::Empty(start) => {
                let local = local_name(start.name().as_ref());
                let is_anchor = matches!(local.as_str(), "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" | "relSizeAnchor" | "absSizeAnchor");
                let name = if is_anchor { get_attr(&start, "name") } else { None };
                let priority = if is_anchor {
                    name.as_deref().and_then(|name| rank.get(name).copied()).unwrap_or(order.len() + sequence)
                } else {
                    usize::MAX / 2 + sequence
                };
                children.push((priority, sequence, vec![Event::Empty(start)]));
                sequence += 1;
            }
            Event::End(end) if in_root && local_name(end.name().as_ref()) == "wsDr" => {
                children.sort_by_key(|(priority, sequence, _)| (*priority, *sequence));
                for (_, _, events) in std::mem::take(&mut children) {
                    for child in events {
                        writer.write_event(child).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    }
                }
                writer.write_event(Event::End(end)).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                in_root = false;
            }
            Event::Eof => break,
            other => writer.write_event(other).map_err(|error| XlsxError::InvalidData(error.to_string()))?,
        }
    }
    String::from_utf8(writer.into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn top_level_drawing_ranges(xml: &str) -> Result<Vec<(usize, usize)>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut root_seen = false;
    let mut depth = 0usize;
    let mut current = None;
    let mut ranges = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(_) if !root_seen => { root_seen = true; depth = 1; }
            Event::Start(_) => {
                if depth == 1 { current = Some(start); }
                depth += 1;
            }
            Event::Empty(_) if root_seen && depth == 1 => ranges.push((start, end)),
            Event::End(_) if depth > 0 => {
                depth -= 1;
                if depth == 1 {
                    if let Some(start) = current.take() { ranges.push((start, end)); }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(ranges)
}

fn first_xml_element(xml: &str, target: &str) -> Result<Option<(usize, usize)>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut start_at = None;
    let mut depth = 0usize;
    loop {
        let event_start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        let event_end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) => {
                if start_at.is_none() && local_name(element.name().as_ref()) == target { start_at = Some(event_start); }
                if start_at.is_some() { depth += 1; }
            }
            Event::Empty(element) if start_at.is_none() && local_name(element.name().as_ref()) == target => return Ok(Some((event_start, event_end))),
            Event::End(_) if start_at.is_some() => {
                depth = depth.saturating_sub(1);
                if depth == 0 { return Ok(Some((start_at.expect("set above"), event_end))); }
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
    }
}

fn group_anchor(anchor: crate::ir::Anchor) -> String {
	format!("<xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff>",
		anchor.col, anchor.col_off, anchor.row, anchor.row_off)
}

fn patch_drawing_group_addition(xml: &str, group: &DrawingGroupAddition) -> Result<String, XlsxError> {
    let mut selected = Vec::new();
    let mut child_shapes = Vec::new();
    for (name, (x, y, cx, cy)) in &group.members {
        let mut found = None;
        for (start, end) in top_level_drawing_ranges(xml)? {
            let wrapper = &xml[start..end];
            if drawing_shape_id(wrapper, name)?.is_some() { found = Some((start, end, wrapper.to_string())); break; }
        }
        let (start, end, wrapper) = found.ok_or_else(|| XlsxError::InvalidData(format!("group member drawing {name:?} is missing")))?;
        let (shape_start, shape_end) = first_xml_element(&wrapper, "sp")?.ok_or_else(|| XlsxError::InvalidData(format!("group member drawing {name:?} is not a shape")))?;
        let mut shape = wrapper[shape_start..shape_end].to_string();
        let (transform_start, transform_end) = first_xml_element(&shape, "xfrm")?.ok_or_else(|| XlsxError::InvalidData(format!("group member drawing {name:?} has no transform")))?;
        let transform = format!("<a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>");
        shape.replace_range(transform_start..transform_end, &transform);
        child_shapes.push(shape);
        selected.push((start, end));
    }
    let id = max_drawing_shape_id(xml)?.checked_add(1).ok_or_else(|| XlsxError::InvalidData("group shape id overflow".into()))?;
    let transform = group.transform;
    let from = group_anchor(group.from);
    let to = format!("<xdr:to><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:to>", group.to.col, group.to.col_off, group.to.row, group.to.row_off);
    let name = escape(&group.name);
    let group_xml = format!("<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" Requires=\"a14\"><xdr:twoCellAnchor><xdr:from>{}</xdr:from>{}<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"{}\" name=\"{}\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/><a:chOff x=\"{}\" y=\"{}\"/><a:chExt cx=\"{}\" cy=\"{}\"/></a:xfrm></xdr:grpSpPr>{}</xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent>",
        from, to, id, name,
        transform.offset.0, transform.offset.1, transform.extent.0, transform.extent.1,
        transform.child_offset.0, transform.child_offset.1, transform.child_extent.0, transform.child_extent.1,
        child_shapes.join(""));
    selected.sort_by_key(|(start, _)| *start);
    let insertion = selected[0].0;
    let mut output = xml.to_string();
    for (start, end) in selected.into_iter().rev() { output.replace_range(start..end, ""); }
    output.insert_str(insertion, &group_xml);
    Ok(output)
}

fn vml_shape_id(element: &BytesStart<'_>) -> Option<u32> {
    get_attr(element, "id")?.rsplit_once("_s")?.1.parse().ok()
}

fn patch_vml_ole_styles(xml: &str, edits: &HashMap<u32, crate::ir::Drawing>) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len() + edits.len() * 96));
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?.into_owned();
        let empty_element = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(start) | Event::Empty(start) if local_name(start.name().as_ref()) == "shape" => {
                if let Some(drawing) = vml_shape_id(&start).and_then(|id| edits.get(&id)) {
                    let name = start.name().as_ref().to_vec();
                    let mut changed = BytesStart::new(String::from_utf8_lossy(&name).into_owned());
                    for attribute in start.attributes().flatten() {
                        let local = local_name(attribute.key.as_ref());
                        if (local == "strokecolor" && drawing.ole_border_color.is_some())
                            || (local == "fillcolor" && drawing.ole_interior_color.is_some())
                            || (local == "strokeweight" && drawing.ole_border_weight.is_some()) { continue; }
                        changed.push_attribute(attribute);
                    }
                    if let Some(color) = &drawing.ole_border_color { changed.push_attribute(("strokecolor", format!("#{color}").as_str())); }
                    if let Some(color) = &drawing.ole_interior_color { changed.push_attribute(("fillcolor", format!("#{color}").as_str())); }
                    if let Some(weight) = drawing.ole_border_weight {
                        let points = match weight { 1 => "0.25pt", 2 => "0.75pt", -4138 => "1.5pt", 4 => "1.95pt", _ => continue };
                        changed.push_attribute(("strokeweight", points));
                    }
                    writer.write_event(if empty_element { Event::Empty(changed) } else { Event::Start(changed) }).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                } else {
                    writer.write_event(if empty_element { Event::Empty(start) } else { Event::Start(start) }).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                }
            }
            Event::Empty(start) => writer.write_event(Event::Empty(start)).map_err(|error| XlsxError::InvalidData(error.to_string()))?,
            Event::Eof => break,
            other => writer.write_event(other).map_err(|error| XlsxError::InvalidData(error.to_string()))?,
        }
    }
    String::from_utf8(writer.into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn patch_vml_ole_shadows(xml: &str, edits: &HashMap<u32, bool>) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len() + edits.len() * 64));
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?.into_owned();
        match &event {
            Event::Start(start) if local_name(start.name().as_ref()) == "shape" => {
                let mut shape = vec![event];
                let mut depth = 1usize;
                while depth > 0 {
                    let next = reader.read_event().map_err(XlsxError::Xml)?.into_owned();
                    match &next {
                        Event::Start(_) => depth += 1,
                        Event::End(_) => depth -= 1,
                        Event::Eof => return Err(XlsxError::InvalidData("VML shape ended unexpectedly".into())),
                        _ => {}
                    }
                    shape.push(next);
                }
                let shape_id = match &shape[0] { Event::Start(start) => vml_shape_id(start), _ => None };
                if let Some((_, shadow)) = shape_id.and_then(|id| edits.get(&id).map(|shadow| (id, *shadow))) {
                    let patched = rewrite_vml_shape_shadow(shape, shadow)?;
                    for event in patched { writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?; }
                } else {
                    for event in shape { writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?; }
                }
            }
            Event::Eof => break,
            _ => writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?,
        }
    }
    String::from_utf8(writer.into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn rewrite_vml_shape_shadow(events: Vec<Event<'static>>, enabled: bool) -> Result<Vec<Event<'static>>, XlsxError> {
    let prefix = match &events[0] {
        Event::Start(start) => std::str::from_utf8(start.name().as_ref()).ok().and_then(|name| name.rsplit_once(':').map(|(prefix, _)| prefix.to_string())).unwrap_or_else(|| "v".into()),
        _ => return Err(XlsxError::InvalidData("VML shape has no start event".into())),
    };
    let mut output = Vec::with_capacity(events.len() + usize::from(enabled));
    let mut found_shadow = false;
    let mut inserted = false;
    let mut index = 0usize;
    while index < events.len() {
        if let Event::Start(start) | Event::Empty(start) = &events[index] {
            let local = local_name(start.name().as_ref());
            if enabled && local == "ClientData" && !inserted && !found_shadow {
                let mut shadow = BytesStart::new(format!("{prefix}:shadow"));
                shadow.push_attribute(("on", "t"));
                shadow.push_attribute(("color", "black"));
                shadow.push_attribute(("obscured", "t"));
                output.push(Event::Empty(shadow.into_owned()));
                inserted = true;
            }
            if local == "shadow" {
                found_shadow = true;
                if enabled {
                    let mut changed = start.to_owned();
                    changed.clear_attributes();
                    for attribute in start.attributes().flatten() {
                        if local_name(attribute.key.as_ref()) != "on" { changed.push_attribute(attribute); }
                    }
                    changed.push_attribute(("on", "t"));
                    output.push(if matches!(events[index], Event::Empty(_)) { Event::Empty(changed) } else { Event::Start(changed) });
                }
                if matches!(events[index], Event::Empty(_)) { index += 1; continue; }
                let mut depth = 0usize;
                while index < events.len() {
                    match &events[index] {
                        Event::Start(_) => depth += 1,
                        Event::End(_) => { depth -= 1; index += 1; if depth == 0 { break; } continue; }
                        _ => {}
                    }
                    index += 1;
                }
                continue;
            }
        }
        if enabled && matches!(&events[index], Event::End(end) if local_name(end.name().as_ref()) == "shape") && !inserted && !found_shadow {
            let mut shadow = BytesStart::new(format!("{prefix}:shadow"));
            shadow.push_attribute(("on", "t"));
            shadow.push_attribute(("color", "black"));
            shadow.push_attribute(("obscured", "t"));
            output.push(Event::Empty(shadow.into_owned()));
            inserted = true;
        }
        output.push(events[index].clone());
        index += 1;
    }
    Ok(output)
}

fn drawing_shape_id(xml: &str, shape_name: &str) -> Result<Option<u32>, XlsxError> {
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(element) | Event::Empty(element) if local_name(element.name().as_ref()) == "cNvPr" && get_attr(&element, "name").as_deref() == Some(shape_name) => {
                return Ok(get_attr(&element, "id").and_then(|id| id.parse().ok()));
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
    }
}

fn xml_root_close(xml: &str, root: &str) -> Result<(usize, String), XlsxError> {
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::End(end) if local_name(end.name().as_ref()) == root => {
                let qname = std::str::from_utf8(end.name().as_ref())
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?.to_string();
                let position = reader.buffer_position() as usize;
                let close_at = position.checked_sub(qname.len() + 3)
                    .ok_or_else(|| XlsxError::InvalidData("XML closing element is malformed".into()))?;
                return Ok((close_at, qname));
            }
            Event::Eof => return Err(XlsxError::InvalidData(format!("XML part has no {root} root"))),
            _ => {}
        }
    }
}

fn append_xml_child(xml: &str, root: &str, child: &str) -> Result<String, XlsxError> {
    let (close_at, _) = xml_root_close(xml, root)?;
    Ok(format!("{}{}{}", &xml[..close_at], child, &xml[close_at..]))
}

fn max_drawing_shape_id(xml: &str) -> Result<u32, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut max_id = 0u32;
    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(start) | Event::Empty(start) if local_name(start.name().as_ref()) == "cNvPr" => {
                max_id = max_id.max(get_attr(&start, "id").and_then(|id| id.parse().ok()).unwrap_or(0));
            }
            Event::Eof => return Ok(max_id),
            _ => {}
        }
    }
}

fn clone_drawing_anchor_xml(
    xml: &str,
    source_name: &str,
    new_name: &str,
    source_id: u32,
    new_id: u32,
) -> Result<String, XlsxError> {
    let bytes = xml.as_bytes();
    let mut reader = Reader::from_str(xml);
    let mut anchor_start = None;
    let mut depth = 0usize;
    let mut matched = false;
    loop {
        let event_start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        let event_end = reader.buffer_position() as usize;
        if anchor_start.is_none() {
            if matches!(&event, Event::Start(start) if local_name(start.name().as_ref()) == "twoCellAnchor" || local_name(start.name().as_ref()) == "oneCellAnchor") {
                anchor_start = Some(event_start);
                depth = 1;
            }
            continue;
        }
        match &event {
            Event::Start(start) => {
                if local_name(start.name().as_ref()) == "cNvPr" && get_attr(start, "name").as_deref() == Some(source_name) {
                    matched = true;
                }
                depth += 1;
            }
            Event::Empty(start) => {
                if local_name(start.name().as_ref()) == "cNvPr" && get_attr(start, "name").as_deref() == Some(source_name) {
                    matched = true;
                }
            }
            Event::End(end) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    if matched {
                        let start = anchor_start.expect("anchor start set");
                        let mut anchor = String::from_utf8_lossy(&bytes[start..event_end]).into_owned();
                        anchor = anchor.replacen(&format!("name=\"{source_name}\""), &format!("name=\"{new_name}\""), 1);
                        anchor = anchor.replacen(&format!("id=\"{source_id}\""), &format!("id=\"{new_id}\""), 1);
                        anchor = anchor.replace(&format!("_x0000_s{source_id}"), &format!("_x0000_s{new_id}"));
                        let anchor_tag = if anchor.contains("<xdr:twoCellAnchor") { "<xdr:twoCellAnchor" } else { "<xdr:oneCellAnchor" };
                        anchor = anchor.replacen(anchor_tag, &format!("{anchor_tag} xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\""), 1);
                        return Ok(anchor);
                    }
                    anchor_start = None;
                }
                let _ = end;
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Err(XlsxError::InvalidData(format!("drawing anchor {source_name:?} is missing")))
}

fn find_ole_object_xml(xml: &str, shape_id: u32) -> Result<Option<String>, XlsxError> {
    let bytes = xml.as_bytes();
    let mut reader = Reader::from_str(xml);
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) if local_name(element.name().as_ref()) == "oleObject"
                && get_attr(&element, "shapeId").and_then(|id| id.parse::<u32>().ok()) == Some(shape_id) =>
            {
                let mut depth = 1usize;
                loop {
                    let nested = reader.read_event().map_err(XlsxError::Xml)?;
                    match nested {
                        Event::Start(_) => depth += 1,
                        Event::End(element) if local_name(element.name().as_ref()) == "oleObject" => {
                            depth = depth.saturating_sub(1);
                            if depth == 0 {
                                let finish = reader.buffer_position() as usize;
                                return Ok(Some(String::from_utf8_lossy(&bytes[start..finish]).into_owned()));
                            }
                        }
                        Event::End(_) => depth = depth.saturating_sub(1),
                        Event::Eof => return Err(XlsxError::InvalidData("OLE object element is unterminated".into())),
                        _ => {}
                    }
                }
            }
            Event::Empty(element) if local_name(element.name().as_ref()) == "oleObject"
                && get_attr(&element, "shapeId").and_then(|id| id.parse::<u32>().ok()) == Some(shape_id) =>
            {
                return Ok(Some(String::from_utf8_lossy(&bytes[start..end]).into_owned()));
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
    }
}

fn find_vml_ole_shape_xml(xml: &str, shape_id: u32) -> Result<Option<String>, XlsxError> {
    let bytes = xml.as_bytes();
    let expected_id = format!("_x0000_s{shape_id}");
    let mut reader = Reader::from_str(xml);
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(element) if local_name(element.name().as_ref()) == "shape"
                && get_attr(&element, "id").as_deref() == Some(expected_id.as_str()) =>
            {
                let mut depth = 1usize;
                loop {
                    let nested = reader.read_event().map_err(XlsxError::Xml)?;
                    match nested {
                        Event::Start(_) => depth += 1,
                        Event::End(element) if local_name(element.name().as_ref()) == "shape" => {
                            depth = depth.saturating_sub(1);
                            if depth == 0 {
                                let finish = reader.buffer_position() as usize;
                                return Ok(Some(String::from_utf8_lossy(&bytes[start..finish]).into_owned()));
                            }
                        }
                        Event::End(_) => depth = depth.saturating_sub(1),
                        Event::Eof => return Err(XlsxError::InvalidData("VML shape is unterminated".into())),
                        _ => {}
                    }
                }
            }
            Event::Empty(element) if local_name(element.name().as_ref()) == "shape"
                && get_attr(&element, "id").as_deref() == Some(expected_id.as_str()) =>
            {
                return Ok(Some(String::from_utf8_lossy(&bytes[start..end]).into_owned()));
            }
            Event::Eof => return Ok(None),
            _ => {}
        }
    }
}

fn max_vml_z_index(xml: &str) -> Result<u32, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut max_index = 0u32;
    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(element) | Event::Empty(element) if local_name(element.name().as_ref()) == "shape" => {
                if let Some(style) = get_attr(&element, "style") {
                    for property in style.split(';') {
                        if let Some(value) = property.trim().strip_prefix("z-index:") {
                            max_index = max_index.max(value.trim().parse().unwrap_or(0));
                        }
                    }
                }
            }
            Event::Eof => return Ok(max_index),
            _ => {}
        }
    }
}

fn max_vml_shape_id(xml: &str) -> Result<u32, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut max_id = 1024u32;
    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(element) | Event::Empty(element) if local_name(element.name().as_ref()) == "shape" => {
                if let Some(id) = get_attr(&element, "id").and_then(|id| id.strip_prefix("_x0000_s").and_then(|id| id.parse::<u32>().ok())) {
                    max_id = max_id.max(id);
                }
            }
            Event::Eof => return Ok(max_id),
            _ => {}
        }
    }
}

fn write_vml_ole_shape(drawing: &crate::ir::Drawing, shape_id: u32, z_index: u32, image_relationship: &str) -> String {
    let to = drawing.to.unwrap_or(drawing.from);
    let pixel_offset = |emu: i64| (emu as f64 / 9_525.0).round() as i64;
    let anchor = format!(
        "{}, {}, {}, {}, {}, {}, {}, {}",
        drawing.from.col, pixel_offset(drawing.from.col_off), drawing.from.row, pixel_offset(drawing.from.row_off),
        to.col, pixel_offset(to.col_off), to.row, pixel_offset(to.row_off),
    );
    let (left, top) = drawing.ole_clone_position.unwrap_or((0.0, 0.0));
    let (width, height) = drawing.extent
        .map(|(width, height)| (width as f64 / 12_700.0, height as f64 / 12_700.0))
        .unwrap_or((51.0, 40.5));
    format!(
        r##"<v:shape id="_x0000_s{shape_id}" type="#_x0000_t75" style='position:absolute;margin-left:{}pt;margin-top:{}pt;width:{}pt;height:{}pt;z-index:{z_index}' filled="t" fillcolor="window [65]" stroked="t" strokecolor="windowText [64]" o:insetmode="auto"><v:fill color2="window [65]"/><v:imagedata o:relid="{}" o:title=""/><x:ClientData ObjectType="Pict"><x:SizeWithCells/><x:Anchor>{anchor}</x:Anchor><x:CF>Pict</x:CF><x:AutoPict/></x:ClientData></v:shape>"##,
        format_vml_point(left), format_vml_point(top),
        format_vml_point(width), format_vml_point(height), escape(image_relationship),
    )
}

fn vml_picture_shapetype() -> &'static str {
    r#"<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>"#
}

fn format_vml_point(point: f64) -> String {
    let formatted = format!("{point:.2}");
    formatted.trim_end_matches('0').trim_end_matches('.').to_string()
}

fn clone_vml_ole_shape_xml(
    xml: &str,
    source_shape_id: u32,
    new_shape_id: u32,
    z_index: u32,
    drawing: &crate::ir::Drawing,
) -> Result<String, XlsxError> {
    let source = find_vml_ole_shape_xml(xml, source_shape_id)?
        .ok_or_else(|| XlsxError::InvalidData(format!("VML OLE fallback for shape {source_shape_id} is missing")))?;
    let position = drawing.ole_clone_position
        .ok_or_else(|| XlsxError::InvalidData("OLE duplicate has no VML point position".into()))?;
    let from = drawing.from;
    let to = drawing.to.ok_or_else(|| XlsxError::InvalidData("OLE duplicate has no far anchor".into()))?;
    let pixel_offset = |emu: i64| (emu as f64 / 9_525.0).round() as i64;
    let anchor = format!(
        "{}, {}, {}, {}, {}, {}, {}, {}",
        from.col, pixel_offset(from.col_off), from.row, pixel_offset(from.row_off),
        to.col, pixel_offset(to.col_off), to.row, pixel_offset(to.row_off),
    );
    let mut cloned = source.replacen(
        &format!("id=\"_x0000_s{source_shape_id}\""),
        &format!("id=\"_x0000_s{new_shape_id}\""),
        1,
    );
    let style = xml_attribute_value(&source, "style")
        .ok_or_else(|| XlsxError::InvalidData("VML OLE shape has no style".into()))?;
    let mut properties = style.split(';').map(str::trim).filter(|property| !property.is_empty())
        .filter(|property| !property.starts_with("margin-left:") && !property.starts_with("margin-top:") && !property.starts_with("z-index:") && !property.starts_with("visibility:"))
        .map(str::to_string).collect::<Vec<_>>();
    properties.push(format!("margin-left:{}pt", format_vml_point(position.0)));
    properties.push(format!("margin-top:{}pt", format_vml_point(position.1)));
    properties.push(format!("z-index:{z_index}"));
    if drawing.ole_clone_visible == Some(false) {
        properties.push("visibility:hidden".into());
    }
    let quote = if source.contains(&format!("style='")) { '\'' } else { '"' };
    let style_prefix = format!("style={quote}{style}{quote}");
    let style_replacement = format!("style={quote}{}{quote}", properties.join(";"));
    cloned = cloned.replacen(&style_prefix, &style_replacement, 1);
    let anchor_open = "<x:Anchor>";
    let anchor_close = "</x:Anchor>";
    if let (Some(start), Some(relative_end)) = (cloned.find(anchor_open), cloned.find(anchor_close)) {
        let content_start = start + anchor_open.len();
        if relative_end >= content_start {
            cloned.replace_range(content_start..relative_end, &anchor);
        }
    }
    Ok(cloned)
}

fn xml_attribute_value(xml: &str, name: &str) -> Option<String> {
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event().ok()? {
            Event::Start(start) | Event::Empty(start) => return get_attr(&start, name),
            Event::Eof => return None,
            _ => {}
        }
    }
}

fn clone_ole_object_xml(xml: &str, source_id: u32, new_id: u32, new_relationship: &str) -> Result<String, XlsxError> {
    let old_relationship = xml_attribute_value(xml, "id")
        .ok_or_else(|| XlsxError::InvalidData("OLE object has no relationship id".into()))?;
    let mut clone = xml.replacen(&format!("shapeId=\"{source_id}\""), &format!("shapeId=\"{new_id}\""), 1);
    clone = clone.replacen(&format!("r:id=\"{old_relationship}\""), &format!("r:id=\"{new_relationship}\""), 1);
    Ok(clone)
}

fn write_ole_drawing_shape(drawing: &crate::ir::Drawing, id: u32) -> Result<String, XlsxError> {
    let name = escape(drawing.name.as_deref().unwrap_or("Object"));
    let to = drawing.to.ok_or_else(|| XlsxError::InvalidData("new OLE object has no far anchor".into()))?;
    let marker = |tag: &str, anchor: crate::ir::Anchor| format!(
        "<xdr:{tag}><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:{tag}>",
        anchor.col, anchor.col_off, anchor.row, anchor.row_off,
    );
    Ok(format!(
        r#"<xdr:twoCellAnchor editAs="oneCell">{}{}<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{id}" name="{name}" hidden="1"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>"#,
        marker("from", drawing.from), marker("to", to),
    ))
}

fn write_worksheet_ole_object(drawing: &crate::ir::Drawing, shape_id: u32, relationship_id: &str, icon_relationship_id: &str) -> Result<String, XlsxError> {
    let to = drawing.to.ok_or_else(|| XlsxError::InvalidData("new OLE object has no far anchor".into()))?;
    let marker = |tag: &str, anchor: crate::ir::Anchor| format!(
        "<{tag}><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></{tag}>",
        anchor.col, anchor.col_off, anchor.row, anchor.row_off,
    );
    let prog_id = escape(drawing.ole_prog_id.as_deref().unwrap_or("Package"));
    let aspect = if drawing.ole_display_as_icon.unwrap_or(false) { " dvAspect=\"DVASPECT_ICON\"" } else { "" };
    let word_document = drawing.ole_prog_id.as_deref().is_some_and(|id| id.eq_ignore_ascii_case("Word.Document.12"));
    let auto_load = if word_document { "" } else { " autoLoad=\"0\"" };
    let object_properties = if word_document {
        format!(
            r#"<objectPr defaultSize="0" r:id="{}"><anchor moveWithCells="1">{}{}</anchor></objectPr>"#,
            escape(icon_relationship_id), marker("from", drawing.from), marker("to", to),
        )
    } else {
        format!(
            r#"<objectPr defaultSize="0" disabled="0" locked="1" print="1" r:id="{}"><anchor moveWithCells="1" sizeWithCells="0">{}{}</anchor></objectPr>"#,
            escape(icon_relationship_id), marker("from", drawing.from), marker("to", to),
        )
    };
    Ok(format!(
        r#"<oleObject xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" progId="{prog_id}"{aspect} shapeId="{shape_id}" r:id="{}"{auto_load}>{object_properties}</oleObject>"#,
        escape(relationship_id),
    ))
}

fn append_new_ole_vml(
    archive: &mut OoxmlArchive,
    sheet_path: &str,
    drawing: &crate::ir::Drawing,
    drawing_xml: &str,
    worksheet_rels: &mut String,
    names: &mut std::collections::HashSet<String>,
    plan: &mut DrawingAdditionPackage,
) -> Result<(u32, String), XlsxError> {
    let sheet_relationships = parse_relationships(worksheet_rels)
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    let vml_relation = sheet_relationships.iter().find(|(_, relation)| relation.rel_type.ends_with("/vmlDrawing"));
    let (vml_path, vml_id) = if let Some((id, relation)) = vml_relation {
        (resolve_part_target(sheet_path, &relation.target), id.clone())
    } else {
        let path = unique_part_path(names, "xl/drawings", "vmlDrawing", ".vml");
        names.insert(path.clone());
        let before = sheet_relationships.keys().cloned().collect::<std::collections::HashSet<_>>();
        *worksheet_rels = append_relationship(
            Some(worksheet_rels),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
            &relative_sheet_target(sheet_path, &path),
        )?;
        let after = parse_relationships(worksheet_rels)
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let id = after.keys().find(|id| !before.contains(*id)).cloned()
            .ok_or_else(|| XlsxError::InvalidData("VML relationship id was not allocated".into()))?;
        plan.worksheet_legacy_drawing.insert(sheet_path.to_string(), id.clone());
        plan.content_types.push((path.clone(), "application/vnd.openxmlformats-officedocument.vmlDrawing".into()));
        plan.new_parts.push((path.clone(), empty_vml_drawing()));
        (path, id)
    };
    plan.worksheet_legacy_drawing.insert(sheet_path.to_string(), vml_id);

    let archived_vml = archive.try_read_part(&vml_path)?;
    let mut vml_xml = plan.parts.get(&vml_path).cloned()
        .or_else(|| plan.new_parts.iter().find(|(path, _)| path == &vml_path).map(|(_, xml)| xml.clone()))
        .or(archived_vml)
        .unwrap_or_else(empty_vml_drawing);
    if !vml_xml.contains("id=\"_x0000_t75\"") {
        vml_xml = append_xml_child(&vml_xml, "xml", vml_picture_shapetype())?;
    }

    // Excel stores one EMF preview relationship on the worksheet and another
    // from its VML fallback. Both point at the same package icon image.
    let image_path = unique_part_path(names, "xl/media", "image", ".emf");
    names.insert(image_path.clone());
    let word_document = drawing.ole_prog_id.as_deref().is_some_and(|prog_id| {
        prog_id.eq_ignore_ascii_case("Word.Document.12")
    });
    let icon = if word_document && drawing.ole_display_as_icon.unwrap_or(false) {
        include_bytes!("data/ole_word_icon.emf").as_slice()
    } else if word_document {
        include_bytes!("data/ole_word_content.emf").as_slice()
    } else if drawing.ole_package_is_empty == Some(true) {
        include_bytes!("data/ole_package_empty_icon.emf").as_slice()
    } else {
        include_bytes!("data/ole_package_icon.emf").as_slice()
    };
    plan.new_binary_parts.push((image_path.clone(), icon.to_vec()));
    plan.content_types.push((image_path.clone(), "image/x-emf".into()));

    let before_sheet: std::collections::HashSet<String> = parse_relationships(worksheet_rels)
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?.into_keys().collect();
    *worksheet_rels = append_relationship(
        Some(worksheet_rels),
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        &relative_sheet_target(sheet_path, &image_path),
    )?;
    let after_sheet = parse_relationships(worksheet_rels)
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    let icon_relationship_id = after_sheet.keys().find(|id| !before_sheet.contains(*id)).cloned()
        .ok_or_else(|| XlsxError::InvalidData("OLE icon relationship id was not allocated".into()))?;

    let vml_rels_path = sheet_rels_path(&vml_path);
    let archived_vml_rels = archive.try_read_part(&vml_rels_path)?;
    let original_vml_rels = plan.relationship_xml.get(&vml_rels_path).cloned()
        .or_else(|| plan.new_parts.iter().find(|(path, _)| path == &vml_rels_path).map(|(_, xml)| xml.clone()))
        .or(archived_vml_rels);
    let mut vml_rels = original_vml_rels.clone().unwrap_or_else(|| r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.into());
    let before_vml: std::collections::HashSet<String> = parse_relationships(&vml_rels)
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?.into_keys().collect();
    vml_rels = append_relationship(Some(&vml_rels),
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        &relative_sheet_target(&vml_path, &image_path))?;
    let after_vml = parse_relationships(&vml_rels)
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    let vml_image_id = after_vml.keys().find(|id| !before_vml.contains(*id)).cloned()
        .ok_or_else(|| XlsxError::InvalidData("VML icon relationship id was not allocated".into()))?;

    let shape_id = max_drawing_shape_id(drawing_xml)?.max(max_vml_shape_id(&vml_xml)?).max(1024).checked_add(1)
        .ok_or_else(|| XlsxError::InvalidData("drawing shape id overflow".into()))?;
    let z_index = max_vml_z_index(&vml_xml)?.checked_add(1)
        .ok_or_else(|| XlsxError::InvalidData("VML z-index overflow".into()))?;
    vml_xml = append_xml_child(&vml_xml, "xml", &write_vml_ole_shape(drawing, shape_id, z_index, &vml_image_id))?;

    if archive.try_read_part(&vml_path)?.is_some() {
        plan.parts.insert(vml_path, vml_xml);
    } else if let Some((_, xml)) = plan.new_parts.iter_mut().find(|(path, _)| path == &vml_path) {
        *xml = vml_xml;
    } else {
        plan.new_parts.push((vml_path, vml_xml));
    }
    if original_vml_rels.is_some() && archive.try_read_part(&vml_rels_path)?.is_some() {
        plan.relationship_xml.insert(vml_rels_path, vml_rels);
    } else if let Some((_, xml)) = plan.new_parts.iter_mut().find(|(path, _)| path == &vml_rels_path) {
        *xml = vml_rels;
    } else {
        plan.new_parts.push((vml_rels_path, vml_rels));
    }
    Ok((shape_id, icon_relationship_id))
}

fn append_ole_objects_xml(xml: &str, additions: &[String]) -> Result<String, XlsxError> {
    if additions.is_empty() { return Ok(xml.to_string()); }
    if xml.contains("</oleObjects>") {
        return Ok(xml.replacen("</oleObjects>", &format!("{}</oleObjects>", additions.join("")), 1));
    }
    let (close_at, _) = xml_root_close(xml, "worksheet")?;
    Ok(format!("{}<oleObjects>{}</oleObjects>{}", &xml[..close_at], additions.join(""), &xml[close_at..]))
}

fn unique_embedding_path(names: &std::collections::HashSet<String>, source: &str) -> String {
    let (directory, filename) = source.rsplit_once('/').unwrap_or(("", source));
    let (stem, extension) = filename.rsplit_once('.').unwrap_or((filename, "bin"));
    let digit_start = stem.char_indices().rev().take_while(|(_, ch)| ch.is_ascii_digit()).last().map(|(index, _)| index);
    let (prefix, first_suffix) = digit_start.map_or((stem, 1), |index| {
        let suffix = stem[index..].parse::<u32>().unwrap_or(0).saturating_add(1).max(1);
        (&stem[..index], suffix)
    });
    for suffix in first_suffix.. {
        let candidate_name = format!("{prefix}{suffix}.{extension}");
        let candidate = if directory.is_empty() { candidate_name } else { format!("{directory}/{candidate_name}") };
        if !names.contains(&candidate) { return candidate; }
    }
    unreachable!()
}

fn content_type_for_part(content_types: &str, part: &str) -> Result<String, XlsxError> {
    let extension = part.rsplit_once('.').map(|(_, extension)| extension).unwrap_or_default();
    let part_name = format!("/{part}");
    let mut reader = Reader::from_str(content_types);
    let mut default_type = None;
    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(element) | Event::Empty(element) => match local_name(element.name().as_ref()).as_str() {
                "Override" if get_attr(&element, "PartName").as_deref() == Some(part_name.as_str()) => {
                    return get_attr(&element, "ContentType").ok_or_else(|| XlsxError::InvalidData("content type override has no ContentType".into()));
                }
                "Default" if get_attr(&element, "Extension").is_some_and(|value| value.eq_ignore_ascii_case(extension)) => {
                    default_type = get_attr(&element, "ContentType");
                }
                _ => {}
            },
            Event::Eof => return default_type.ok_or_else(|| XlsxError::InvalidData(format!("content type for {part:?} is missing"))),
            _ => {}
        }
    }
}

fn patch_ole_object_anchors(xml: &str, edits: &HashMap<u32, crate::ir::Drawing>) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len() + edits.len() * 128));
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?.into_owned();
        match &event {
            Event::Empty(start) if local_name(start.name().as_ref()) == "oleObject" => {
                let shape_id = get_attr(start, "shapeId").and_then(|id| id.parse::<u32>().ok());
                if let Some(auto_load) = shape_id.and_then(|id| edits.get(&id)).and_then(|drawing| drawing.ole_auto_load) {
                    let mut changed = start.to_owned();
                    changed.clear_attributes();
                    for attribute in start.attributes().flatten() {
                        if local_name(attribute.key.as_ref()) != "autoLoad" { changed.push_attribute(attribute); }
                    }
                    changed.push_attribute(("autoLoad", if auto_load { "1" } else { "0" }));
                    writer.write_event(Event::Empty(changed)).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                } else {
                    writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                }
            }
            Event::Start(start) if local_name(start.name().as_ref()) == "oleObject" => {
                let shape_id = get_attr(start, "shapeId").and_then(|id| id.parse::<u32>().ok());
                let mut object = vec![event];
                let mut depth = 1usize;
                while depth > 0 {
                    let next = reader.read_event().map_err(XlsxError::Xml)?.into_owned();
                    match &next {
                        Event::Start(_) => depth += 1,
                        Event::End(_) => depth -= 1,
                        Event::Eof => return Err(XlsxError::InvalidData("OLE object ended unexpectedly".into())),
                        _ => {}
                    }
                    object.push(next);
                }
                if let Some(drawing) = shape_id.and_then(|id| edits.get(&id)) {
                    let patched = rewrite_ole_object_anchor(object, drawing)?;
                    for event in patched { writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?; }
                } else {
                    for event in object { writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?; }
                }
            }
            Event::Eof => break,
            _ => writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?,
        }
    }
    String::from_utf8(writer.into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn rewrite_ole_object_anchor(events: Vec<Event<'static>>, drawing: &crate::ir::Drawing) -> Result<Vec<Event<'static>>, XlsxError> {
    let mut output = Vec::with_capacity(events.len());
    let mut index = 0usize;
    let mut from_count = 0usize;
    let mut to_count = 0usize;
    while index < events.len() {
        if let Event::Start(start) | Event::Empty(start) = &events[index] {
            let local = local_name(start.name().as_ref());
            let mut updates = Vec::<(&str, &str)>::new();
            let bool_text = |value: bool| if value { "1" } else { "0" };
            match local.as_str() {
                "oleObject" => if let Some(value) = drawing.ole_auto_load { updates.push(("autoLoad", bool_text(value))); },
                "objectPr" => {
                    if let Some(value) = drawing.ole_locked { updates.push(("locked", bool_text(value))); }
                    if let Some(value) = drawing.ole_print_object { updates.push(("print", bool_text(value))); }
                    if let Some(value) = drawing.ole_enabled { updates.push(("disabled", bool_text(!value))); }
                    if let Some(value) = drawing.ole_on_action.as_deref() { updates.push(("macro", value)); }
                }
                _ => {}
            }
            if !updates.is_empty() {
                let mut changed = start.to_owned();
                changed.clear_attributes();
                for attribute in start.attributes().flatten() {
                    let name = local_name(attribute.key.as_ref());
                    if !updates.iter().any(|(key, _)| *key == name) { changed.push_attribute(attribute); }
                }
                for (name, value) in updates { changed.push_attribute((name, value)); }
                output.push(if matches!(events[index], Event::Empty(_)) { Event::Empty(changed) } else { Event::Start(changed) });
                index += 1;
                continue;
            }
        }
        if let (Some(placement), Event::Start(start) | Event::Empty(start)) = (drawing.ole_placement, &events[index]) {
            if local_name(start.name().as_ref()) == "anchor" {
                let (move_with_cells, size_with_cells) = match placement {
                    1 => (true, true),
                    2 => (true, false),
                    _ => (false, false),
                };
                let mut changed = start.to_owned();
                changed.clear_attributes();
                for attribute in start.attributes().flatten() {
                    if !matches!(attribute.key.as_ref(), b"moveWithCells" | b"sizeWithCells") {
                        changed.push_attribute(attribute);
                    }
                }
                if move_with_cells { changed.push_attribute(("moveWithCells", "1")); }
                if size_with_cells { changed.push_attribute(("sizeWithCells", "1")); }
                output.push(if matches!(events[index], Event::Empty(_)) { Event::Empty(changed) } else { Event::Start(changed) });
                index += 1;
                continue;
            }
        }
        if let Event::Start(start) = &events[index] {
            let local = local_name(start.name().as_ref());
            if local == "from" || (local == "to" && drawing.to.is_some()) {
                let anchor = if local == "from" { drawing.from } else { drawing.to.expect("checked") };
                let tag = std::str::from_utf8(start.name().as_ref()).unwrap_or("xdr:from").to_string();
                let marker = tag.clone();
                output.push(Event::Start(BytesStart::new(marker.clone()).into_owned()));
                let child_prefix = events[index + 1..].iter().find_map(|event| match event {
                    Event::Start(element) | Event::Empty(element) if local_name(element.name().as_ref()) == "col" => {
                        std::str::from_utf8(element.name().as_ref()).ok().and_then(|name| name.rsplit_once(':').map(|(prefix, _)| prefix.to_string()))
                    }
                    _ => None,
                }).unwrap_or_default();
                for (field, value) in [("col", i64::from(anchor.col)), ("colOff", anchor.col_off), ("row", i64::from(anchor.row)), ("rowOff", anchor.row_off)] {
                    let child = if child_prefix.is_empty() { field.to_string() } else { format!("{child_prefix}:{field}") };
                    output.push(Event::Start(BytesStart::new(child.clone()).into_owned()));
                    output.push(Event::Text(BytesText::new(&value.to_string()).into_owned()));
                    output.push(Event::End(BytesEnd::new(child).into_owned()));
                }
                output.push(Event::End(BytesEnd::new(marker).into_owned()));
                if local == "from" { from_count += 1; } else { to_count += 1; }
                let mut depth = 0usize;
                while index < events.len() {
                    match &events[index] {
                        Event::Start(_) => depth += 1,
                        Event::End(_) => { depth -= 1; index += 1; if depth == 0 { break; } continue; }
                        _ => {}
                    }
                    index += 1;
                }
                continue;
            }
        }
        output.push(events[index].clone());
        index += 1;
    }
    if from_count != 1 || (drawing.to.is_some() && to_count != 1) {
        return Err(XlsxError::InvalidData("OLE object anchor is missing or ambiguous while saving its position".into()));
    }
    Ok(output)
}

fn append_drawing_shapes_xml(
    xml: &str,
    drawings: &[crate::ir::Drawing],
) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut max_id = 0u32;
    let mut root_end: Option<(usize, String)> = None;
    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(start) | Event::Empty(start) if local_name(start.name().as_ref()) == "cNvPr" => {
                max_id = max_id.max(get_attr(&start, "id").and_then(|id| id.parse().ok()).unwrap_or(0));
            }
            Event::End(end) if local_name(end.name().as_ref()) == "wsDr" => {
                let qname = std::str::from_utf8(end.name().as_ref()).map_err(|error| XlsxError::InvalidData(error.to_string()))?.to_string();
                root_end = Some((reader.buffer_position() as usize, qname));
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let (end, qname) = root_end.ok_or_else(|| XlsxError::InvalidData("drawing part has no worksheet drawing root".into()))?;
    let closing_len = qname.len() + 3;
    let close_at = end.checked_sub(closing_len).ok_or_else(|| XlsxError::InvalidData("drawing root end is malformed".into()))?;
    let mut additions = String::new();
    for drawing in drawings {
        max_id = max_id.checked_add(1).ok_or_else(|| XlsxError::InvalidData("drawing shape id overflow".into()))?;
        additions.push_str(&write_drawing_shape(drawing, max_id)?);
    }
    Ok(format!("{}{}{}", &xml[..close_at], additions, &xml[close_at..]))
}

fn write_drawing_shape(drawing: &crate::ir::Drawing, id: u32) -> Result<String, XlsxError> {
    let name = escape(drawing.name.as_deref().unwrap_or("Shape"));
    let crate::ir::DrawingKind::Shape(shape) = &drawing.kind else {
        return Err(XlsxError::InvalidData("only shapes can be added to a drawing part by this writer".into()));
    };
    let point = |value: i64| value.to_string();
    let marker = |tag: &str, anchor: crate::ir::Anchor| format!(
        "<xdr:{tag}><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:{tag}>",
        anchor.col, anchor.col_off, anchor.row, anchor.row_off,
    );
    let (anchor_open, anchor_close, to_xml) = match drawing.to {
        Some(to) => ("twoCellAnchor", "</xdr:twoCellAnchor>", marker("to", to)),
        None => ("oneCellAnchor", "</xdr:oneCellAnchor>", format!("<xdr:ext cx=\"{}\" cy=\"{}\"", drawing.extent.map_or(0, |size| size.0), drawing.extent.map_or(0, |size| size.1)) + "/>"),
    };
    let (width, height) = drawing.extent.unwrap_or_else(|| {
        let to = drawing.to.unwrap_or(drawing.from);
        ((to.col_off - drawing.from.col_off).max(1), (to.row_off - drawing.from.row_off).max(1))
    });
    let rotation = (shape.rotation != 0).then(|| format!(" rot=\"{}\"", shape.rotation)).unwrap_or_default();
    let flip_h = shape.flip_h.then_some(" flipH=\"1\"").unwrap_or_default();
    let flip_v = shape.flip_v.then_some(" flipV=\"1\"").unwrap_or_default();
    let mut paint = String::new();
    if let Some(fill) = shape.fill.as_deref() {
        let color = write_drawing_color(fill, shape.fill_theme, shape.fill_legacy_color_index, shape.fill_brightness, None);
        paint.push_str(&format!("<a:solidFill>{color}</a:solidFill>"));
    } else {
        paint.push_str("<a:noFill/>");
    }
    if let Some(line) = shape.line.as_ref() {
        paint.push_str(&format!("<a:ln w=\"{}\">", line.width.max(0)));
        if let Some(pattern) = line.pattern.as_deref() {
            let foreground = write_drawing_color(&line.color, line.color_theme, line.legacy_color_index, line.color_brightness, line.color_shade);
            let background = write_drawing_color(line.back_color.as_deref().unwrap_or("FFFFFF"), line.back_color_theme, line.back_legacy_color_index, line.back_color_brightness, line.back_color_shade);
            paint.push_str(&format!("<a:pattFill prst=\"{}\"><a:fgClr>{foreground}</a:fgClr><a:bgClr>{background}</a:bgClr></a:pattFill>", escape(pattern)));
        } else {
            let color = write_drawing_color(&line.color, line.color_theme, line.legacy_color_index, line.color_brightness, line.color_shade);
            paint.push_str(&format!("<a:solidFill>{color}</a:solidFill>"));
        }
        if let Some(dash) = line.dash.as_deref() { paint.push_str(&format!("<a:prstDash val=\"{}\"/>", escape(dash))); }
        for (tag, kind, width, length) in [("headEnd", line.head_end.as_deref(), line.head_end_width.as_deref(), line.head_end_length.as_deref()), ("tailEnd", line.tail_end.as_deref(), line.tail_end_width.as_deref(), line.tail_end_length.as_deref())] {
            let kind = kind.unwrap_or("none");
            let width = width.map(|value| format!(" w=\"{}\"", escape(value))).unwrap_or_default();
            let length = length.map(|value| format!(" len=\"{}\"", escape(value))).unwrap_or_default();
            paint.push_str(&format!("<a:{tag} type=\"{}\"{width}{length}/>", escape(kind)));
        }
        paint.push_str("</a:ln>");
    } else {
        paint.push_str("<a:ln><a:noFill/></a:ln>");
    }
    let mut adjusts = String::new();
    for (name, value) in &shape.adjusts {
        adjusts.push_str(&format!("<a:gd name=\"{}\" fmla=\"val {}\"/>", escape(name), value));
    }
    let connector = shape.geometry == "line" || shape.geometry == "straightConnector1";
    let (object, nonvisual, properties) = if connector {
        ("cxnSp", "nvCxnSpPr", "cNvCxnSpPr")
    } else {
        ("sp", "nvSpPr", "cNvSpPr")
    };
    let text_box = (!connector && shape.text_box).then_some(" txBox=\"1\"").unwrap_or_default();
    let text = if connector { String::new() } else { write_shape_text(shape.text.as_ref()) };
    let geometry = escape(&shape.geometry);
    let av_list = format!("<a:avLst>{adjusts}</a:avLst>");
    let xfrm = format!("<a:xfrm{rotation}{flip_h}{flip_v}><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm>", point(drawing.from.col_off), point(drawing.from.row_off), width, height);
    let style = "<xdr:style><a:lnRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"dk1\"/></a:fontRef></xdr:style>";
    let anchor = format!("<xdr:{anchor_open}>{}{}<xdr:{object} macro=\"\"><xdr:{nonvisual}><xdr:cNvPr id=\"{id}\" name=\"{name}\"/><xdr:{properties}{text_box}/></xdr:{nonvisual}><xdr:spPr>{xfrm}<a:prstGeom prst=\"{geometry}\">{av_list}</a:prstGeom>{paint}</xdr:spPr>{style}{text}</xdr:{object}><xdr:clientData/>{anchor_close}", marker("from", drawing.from), to_xml);
    Ok(anchor)
}

fn drawing_color_modifiers(brightness: Option<f32>) -> String {
    let Some(brightness) = brightness.filter(|value| value.abs() > f32::EPSILON) else { return String::new(); };
    if brightness > 0.0 {
        format!("<a:lumMod val=\"{}\"/><a:lumOff val=\"{}\"/>", ((1.0 - brightness) * 100_000.0).round() as i32, (brightness * 100_000.0).round() as i32)
    } else {
        format!("<a:lumMod val=\"{}\"/>", ((1.0 + brightness) * 100_000.0).round() as i32)
    }
}

fn write_drawing_color(color: &str, theme: Option<usize>, legacy: Option<u8>, brightness: Option<f32>, shade: Option<f32>) -> String {
    let mods = shade.filter(|value| value.is_finite()).map(|value| format!("<a:shade val=\"{}\"/>", (value * 100_000.0).round() as i32)).unwrap_or_else(|| drawing_color_modifiers(brightness));
    if let Some(theme) = theme {
        let name = match theme {
            1 => "dk1", 2 => "lt1", 3 => "dk2", 4 => "lt2",
            5 => "accent1", 6 => "accent2", 7 => "accent3", 8 => "accent4",
            9 => "accent5", 10 => "accent6", 11 => "hlink", 12 => "folHlink",
            _ => return format!("<a:srgbClr val=\"{}\">{mods}</a:srgbClr>", escape(color)),
        };
        return format!("<a:schemeClr val=\"{name}\">{mods}</a:schemeClr>");
    }
    if let Some(index) = legacy {
        return format!("<a:srgbClr val=\"{}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"a14\" xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" a14:legacySpreadsheetColorIndex=\"{index}\">{mods}</a:srgbClr>", escape(color));
    }
    format!("<a:srgbClr val=\"{}\">{mods}</a:srgbClr>", escape(color))
}

fn drawing_theme_token(theme: usize) -> Option<&'static str> {
    Some(match theme {
        1 => "dk1", 2 => "lt1", 3 => "dk2", 4 => "lt2",
        5 => "accent1", 6 => "accent2", 7 => "accent3", 8 => "accent4",
        9 => "accent5", 10 => "accent6", 11 => "hlink", 12 => "folHlink",
        _ => return None,
    })
}

fn write_shape_text(text: Option<&crate::ir::ShapeText>) -> String {
    let Some(text) = text else {
        return "<xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\" rtlCol=\"0\" anchor=\"t\"/><a:lstStyle/><a:p><a:pPr algn=\"l\"/><a:endParaRPr lang=\"en-US\" sz=\"1100\"/></a:p></xdr:txBody>".into();
    };
    let (left, top, right, bottom) = text.insets;
    let anchor = text.anchor.as_deref().unwrap_or("t");
    let wrap = if text.wrap { "square" } else { "none" };
    let overflow = if text.clip { "clip" } else { "overflow" };
    let mut paragraphs = String::new();
    for paragraph in &text.paragraphs {
        let align = paragraph.align.as_deref().unwrap_or("l");
        let line_spacing = if let Some(points) = paragraph.line_pitch {
            format!("<a:lnSpc><a:spcPts val=\"{}\"/></a:lnSpc>", (points * 100.0).round() as i32)
        } else if let Some(scale) = paragraph.line_scale {
            format!("<a:lnSpc><a:spcPct val=\"{}\"/></a:lnSpc>", (scale * 100_000.0).round() as i32)
        } else { String::new() };
        if line_spacing.is_empty() {
            paragraphs.push_str(&format!("<a:p><a:pPr algn=\"{}\"/>", escape(align)));
        } else {
            paragraphs.push_str(&format!("<a:p><a:pPr algn=\"{}\">{line_spacing}</a:pPr>", escape(align)));
        }
        let runs = if paragraph.runs.is_empty() { vec![crate::ir::ShapeRun { text: paragraph.text.clone(), bold: paragraph.bold, italic: None, underline: false, face: None, size: None, color: paragraph.color.clone(), color_theme: None, color_brightness: None, legacy_color_index: None }] } else { paragraph.runs.clone() };
        for run in runs {
            let bold = run.bold.then_some(" b=\"1\"").unwrap_or_default();
            let italic = run.italic.unwrap_or(paragraph.italic).then_some(" i=\"1\"").unwrap_or_default();
            let underline = run.underline.then_some(" u=\"sng\"").unwrap_or_default();
            let color = if let Some(index) = run.legacy_color_index {
                run.color.or_else(|| paragraph.color.clone()).map(|color| format!(
                    "<a:solidFill><a:srgbClr val=\"{}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"a14\" xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" a14:legacySpreadsheetColorIndex=\"{index}\"/></a:solidFill>",
                    escape(&color)
                )).unwrap_or_default()
            } else if let Some(theme) = run.color_theme {
                let scheme = match theme {
                    1 => "dk1", 2 => "lt1", 3 => "dk2", 4 => "lt2", 5 => "accent1", 6 => "accent2",
                    7 => "accent3", 8 => "accent4", 9 => "accent5", 10 => "accent6", 11 => "hlink", 12 => "folHlink",
                    _ => "dk1",
                };
                let modifier = run.color_brightness.map(|brightness| {
                    if brightness > 0.0 {
                        format!("<a:lumMod val=\"{}\"/><a:lumOff val=\"{}\"/>", ((1.0 - brightness) * 100_000.0).round() as i32, (brightness * 100_000.0).round() as i32)
                    } else if brightness < 0.0 {
                        format!("<a:lumMod val=\"{}\"/>", ((1.0 + brightness) * 100_000.0).round() as i32)
                    } else { String::new() }
                }).unwrap_or_default();
                let scheme_color = if modifier.is_empty() {
                    format!("<a:schemeClr val=\"{scheme}\"/>")
                } else {
                    format!("<a:schemeClr val=\"{scheme}\">{modifier}</a:schemeClr>")
                };
                format!("<a:solidFill>{scheme_color}</a:solidFill>")
            } else {
                let modifier = run.color_brightness.map(|brightness| {
                    if brightness > 0.0 {
                        format!("<a:lumMod val=\"{}\"/><a:lumOff val=\"{}\"/>", ((1.0 - brightness) * 100_000.0).round() as i32, (brightness * 100_000.0).round() as i32)
                    } else if brightness < 0.0 {
                        format!("<a:lumMod val=\"{}\"/>", ((1.0 + brightness) * 100_000.0).round() as i32)
                    } else { String::new() }
                }).unwrap_or_default();
                run.color.or_else(|| paragraph.color.clone()).map(|color| {
                    if modifier.is_empty() {
                        format!("<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>", escape(&color))
                    } else {
                        format!("<a:solidFill><a:srgbClr val=\"{}\">{modifier}</a:srgbClr></a:solidFill>", escape(&color))
                    }
                }).unwrap_or_default()
            };
            let latin = run.face.as_deref().or(paragraph.face.as_deref()).map(|face| format!("<a:latin typeface=\"{}\" charset=\"{}\" pitchFamily=\"{}\"/>", escape(face), paragraph.charset.unwrap_or(0), paragraph.pitch_family.unwrap_or(0))).unwrap_or_default();
            let size = run.size.unwrap_or(paragraph.size);
            let rpr = format!("<a:rPr lang=\"en-US\" sz=\"{}\"{bold}{italic}{underline}>{color}{latin}</a:rPr>", (size * 100.0).round() as i32);
            let segments = run.text.split('\n').collect::<Vec<_>>();
            for (index, segment) in segments.iter().enumerate() {
                if !segment.is_empty() {
                    paragraphs.push_str(&format!("<a:r>{rpr}<a:t xml:space=\"preserve\">{}</a:t></a:r>", escape(segment)));
                }
                if index + 1 < segments.len() {
                    paragraphs.push_str(&format!("<a:br>{rpr}</a:br>"));
                }
            }
        }
        paragraphs.push_str(&format!("<a:endParaRPr lang=\"en-US\" sz=\"{}\"/></a:p>", (paragraph.size * 100.0).round() as i32));
    }
    format!("<xdr:txBody><a:bodyPr wrap=\"{wrap}\" vertOverflow=\"{overflow}\" lIns=\"{left}\" tIns=\"{top}\" rIns=\"{right}\" bIns=\"{bottom}\" vert=\"horz\" rtlCol=\"0\" anchor=\"{}\"/><a:lstStyle/>{paragraphs}</xdr:txBody>", escape(anchor))
}

fn write_line_fill_option<W: Write>(writer: &mut Writer<W>, line: Option<&crate::ir::ShapeLine>, prefix: &str) -> Result<(), XlsxError> {
    if let Some(line) = line {
        return write_line_fill(writer, line, prefix);
    }
    let tag = if prefix.is_empty() { "noFill".to_owned() } else { format!("{prefix}:noFill") };
    writer.write_event(Event::Start(BytesStart::new(tag.as_str())))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    writer.write_event(Event::End(BytesEnd::new(tag)))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn patch_drawing_picture_xml(
    xml: &str,
    edits: &HashMap<String, PictureStyleEdit>,
) -> Result<String, XlsxError> {
    fn write_adjustments<W: Write>(writer: &mut Writer<W>, style: &crate::ir::PictureStyle) -> Result<(), XlsxError> {
        let mut write_empty = |tag: &str, attrs: &[(&str, String)]| -> Result<(), XlsxError> {
            let mut element = BytesStart::new(tag);
            for (key, value) in attrs { element.push_attribute((*key, value.as_str())); }
            writer.write_event(Event::Empty(element)).map_err(|error| XlsxError::InvalidData(error.to_string()))
        };
        match style.color_type {
            2 => write_empty("a:grayscl", &[])?,
            3 => write_empty("a:biLevel", &[("thresh", "50000".to_owned())])?,
            _ => {}
        }
        if style.brightness != 0 || style.contrast != 0 {
            let bright = style.brightness.to_string();
            let contrast = style.contrast.to_string();
            write_empty("a:lum", &[("bright", bright), ("contrast", contrast)])?;
        }
        Ok(())
    }
    fn write_crop<W: Write>(writer: &mut Writer<W>, style: &crate::ir::PictureStyle) -> Result<(), XlsxError> {
        let mut element = BytesStart::new("a:srcRect");
        let attributes = [
            ("l", style.crop.0, style.crop.0.to_string()),
            ("t", style.crop.1, style.crop.1.to_string()),
            ("r", style.crop.2, style.crop.2.to_string()),
            ("b", style.crop.3, style.crop.3.to_string()),
        ];
        for (name, value, encoded) in &attributes {
            if *value != 0 { element.push_attribute((*name, encoded.as_str())); }
        }
        writer.write_event(Event::Empty(element)).map_err(|error| XlsxError::InvalidData(error.to_string()))
    }
    fn write_transparency_color<W: Write>(writer: &mut Writer<W>, color: Option<&str>) -> Result<(), XlsxError> {
        let Some(color) = color else { return Ok(()); };
        writer.write_event(Event::Start(BytesStart::new("a:clrChange"))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        writer.write_event(Event::Start(BytesStart::new("a:clrFrom"))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let mut from = BytesStart::new("a:srgbClr");
        from.push_attribute(("val", color));
        writer.write_event(Event::Empty(from)).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        writer.write_event(Event::End(BytesEnd::new("a:clrFrom"))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        writer.write_event(Event::Start(BytesStart::new("a:clrTo"))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let mut to = BytesStart::new("a:srgbClr");
        to.push_attribute(("val", color));
        writer.write_event(Event::Start(to)).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let mut alpha = BytesStart::new("a:alpha");
        alpha.push_attribute(("val", "0"));
        writer.write_event(Event::Empty(alpha)).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        writer.write_event(Event::End(BytesEnd::new("a:srgbClr"))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        writer.write_event(Event::End(BytesEnd::new("a:clrTo"))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        writer.write_event(Event::End(BytesEnd::new("a:clrChange"))).map_err(|error| XlsxError::InvalidData(error.to_string()))
    }

    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len() + edits.len() * 96));
    let mut current_shape: Option<String> = None;
    let mut active_picture: Option<String> = None;
    let mut active_fill: Option<String> = None;
    let mut crop_hits: HashMap<String, usize> = HashMap::new();
    let mut adjustments_written = std::collections::HashSet::new();
    let mut transparency_written = std::collections::HashSet::new();
    let mut skip_other_depth = 0usize;
    let mut skip_adjustment_depth = 0usize;
    let mut hits: HashMap<String, usize> = HashMap::new();
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        match &event {
            Event::Start(element) => {
                let tag = local_name(element.name().as_ref());
                if tag == "cNvPr" { current_shape = get_attr(element, "name"); }
                if skip_adjustment_depth > 0 {
                    skip_adjustment_depth += 1;
                    continue;
                }
                if skip_other_depth > 0 { skip_other_depth += 1; continue; }
                if tag == "blipFill" && active_fill.is_none() {
                    active_fill = current_shape.as_ref().filter(|name| edits.contains_key(*name)).cloned();
                }
                if tag == "srcRect" {
                    if let Some(name) = active_fill.as_ref() {
                        *crop_hits.entry(name.clone()).or_default() += 1;
                        write_crop(&mut writer, &edits[name].after)?;
                        skip_adjustment_depth = 1;
                        continue;
                    }
                }
                if tag == "stretch" {
                    if let Some(name) = active_fill.as_ref() {
                        if crop_hits.get(name).copied().unwrap_or(0) == 0 {
                            write_crop(&mut writer, &edits[name].after)?;
                            *crop_hits.entry(name.clone()).or_default() += 1;
                        }
                    }
                }
                if tag == "blip" && active_picture.is_none() {
                    active_picture = current_shape.as_ref().filter(|name| edits.contains_key(*name)).cloned();
                    if let Some(name) = active_picture.as_ref() {
                        *hits.entry(name.clone()).or_default() += 1;
                    }
                } else if active_picture.is_some() && matches!(tag.as_str(), "lum" | "grayscl" | "biLevel") {
                    skip_adjustment_depth = 1;
                    continue;
                } else if active_picture.is_some() && tag == "clrChange" {
                    if let Some(name) = active_picture.as_ref() {
                        write_transparency_color(&mut writer, edits[name].after.transparency_color.as_deref())?;
                        transparency_written.insert(name.clone());
                    }
                    skip_other_depth = 1;
                    continue;
                } else if active_picture.is_some() && tag == "extLst" {
                    let name = active_picture.as_ref().expect("checked picture");
                    if !adjustments_written.contains(name) {
                        write_adjustments(&mut writer, &edits[name].after)?;
                        adjustments_written.insert(name.clone());
                    }
                    if !transparency_written.contains(name) {
                        write_transparency_color(&mut writer, edits[name].after.transparency_color.as_deref())?;
                        transparency_written.insert(name.clone());
                    }
                }
                writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            Event::Empty(element) => {
                let tag = local_name(element.name().as_ref());
                if tag == "cNvPr" { current_shape = get_attr(element, "name"); }
                if skip_adjustment_depth > 0 { continue; }
                if skip_other_depth > 0 { continue; }
                if tag == "blipFill" && active_fill.is_none() {
                    active_fill = current_shape.as_ref().filter(|name| edits.contains_key(*name)).cloned();
                }
                if tag == "srcRect" {
                    if let Some(name) = active_fill.as_ref() {
                        *crop_hits.entry(name.clone()).or_default() += 1;
                        write_crop(&mut writer, &edits[name].after)?;
                        continue;
                    }
                }
                if tag == "stretch" {
                    if let Some(name) = active_fill.as_ref() {
                        if crop_hits.get(name).copied().unwrap_or(0) == 0 {
                            write_crop(&mut writer, &edits[name].after)?;
                            *crop_hits.entry(name.clone()).or_default() += 1;
                        }
                    }
                }
                if tag == "blip" {
                    if let Some(name) = current_shape.as_ref().filter(|name| edits.contains_key(*name)) {
                        *hits.entry(name.clone()).or_default() += 1;
                        let qualified_name = element.name();
                        let qualified = std::str::from_utf8(qualified_name.as_ref()).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                        writer.write_event(Event::Start(element.to_owned()))
                            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                        write_adjustments(&mut writer, &edits[name].after)?;
                        adjustments_written.insert(name.clone());
                        write_transparency_color(&mut writer, edits[name].after.transparency_color.as_deref())?;
                        transparency_written.insert(name.clone());
                        writer.write_event(Event::End(BytesEnd::new(qualified)))
                            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    } else {
                        writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    }
                    continue;
                }
                if active_picture.is_some() && matches!(tag.as_str(), "lum" | "grayscl" | "biLevel") { continue; }
                if tag == "clrChange" {
                    if let Some(name) = active_picture.as_ref() {
                        write_transparency_color(&mut writer, edits[name].after.transparency_color.as_deref())?;
                        transparency_written.insert(name.clone());
                        continue;
                    }
                }
                writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            Event::End(element) => {
                let tag = local_name(element.name().as_ref());
                if skip_adjustment_depth > 0 {
                    skip_adjustment_depth -= 1;
                    continue;
                }
                if skip_other_depth > 0 { skip_other_depth -= 1; continue; }
                if tag == "blip" {
                    if let Some(name) = active_picture.take() {
                        if !adjustments_written.contains(&name) {
                            write_adjustments(&mut writer, &edits[&name].after)?;
                            adjustments_written.insert(name.clone());
                        }
                        if !transparency_written.contains(&name) {
                            write_transparency_color(&mut writer, edits[&name].after.transparency_color.as_deref())?;
                            transparency_written.insert(name.clone());
                        }
                    }
                }
                if tag == "blipFill" {
                    if let Some(name) = active_fill.take() {
                        if crop_hits.get(&name).copied().unwrap_or(0) == 0 {
                            write_crop(&mut writer, &edits[&name].after)?;
                            *crop_hits.entry(name).or_default() += 1;
                        }
                    }
                }
                writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                if matches!(tag.as_str(), "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" | "relSizeAnchor" | "absSizeAnchor") { current_shape = None; }
            }
            Event::Eof => break,
            _ => writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?,
        }
    }
    for name in edits.keys() {
        if hits.get(name).copied().unwrap_or(0) != 1 {
            return Err(XlsxError::InvalidData(format!("picture {name:?} has {} blip elements while saving PictureFormat", hits.get(name).copied().unwrap_or(0))));
        }
    }
    String::from_utf8(writer.into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn patch_drawing_picture_background_xml(
    xml: &str,
    edits: &HashMap<String, PictureStyleEdit>,
) -> Result<String, XlsxError> {
    fn write_alpha_zero<W: Write>(writer: &mut Writer<W>) -> Result<(), XlsxError> {
        let mut alpha = BytesStart::new("a:alpha");
        alpha.push_attribute(("val", "0"));
        writer.write_event(Event::Empty(alpha)).map_err(|error| XlsxError::InvalidData(error.to_string()))
    }
    fn write_transparent_fill<W: Write>(writer: &mut Writer<W>) -> Result<(), XlsxError> {
        writer.write_event(Event::Start(BytesStart::new("a:solidFill"))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let mut color = BytesStart::new("a:srgbClr");
        color.push_attribute(("val", "000000"));
        writer.write_event(Event::Start(color)).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        write_alpha_zero(writer)?;
        writer.write_event(Event::End(BytesEnd::new("a:srgbClr"))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        writer.write_event(Event::End(BytesEnd::new("a:solidFill"))).map_err(|error| XlsxError::InvalidData(error.to_string()))
    }

    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len() + edits.len() * 32));
    let mut current_shape: Option<String> = None;
    let mut active_shape: Option<String> = None;
    let mut sppr_depth = 0usize;
    let mut color_depth = 0usize;
    let mut color_found = false;
    let mut alpha_zero_found = false;
    let mut skip_depth = 0usize;
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        match &event {
            Event::Start(element) => {
                let tag = local_name(element.name().as_ref());
                if tag == "cNvPr" { current_shape = get_attr(element, "name"); }
                if skip_depth > 0 { skip_depth += 1; continue; }
                if tag == "spPr" && active_shape.is_none() {
                    active_shape = current_shape.as_ref().filter(|name| edits.get(*name).is_some_and(|edit| edit.before.transparent_background != edit.after.transparent_background)).cloned();
                    if active_shape.is_some() { sppr_depth = 1; color_found = false; alpha_zero_found = false; }
                } else if active_shape.is_some() {
                    sppr_depth += 1;
                    if tag == "alpha" && get_attr(element, "val").as_deref() == Some("0") {
                        if !edits[active_shape.as_ref().unwrap()].after.transparent_background {
                            skip_depth = 1;
                            continue;
                        }
                        alpha_zero_found = true;
                    }
                    if color_depth == 0 && matches!(tag.as_str(), "srgbClr" | "schemeClr" | "sysClr" | "prstClr") {
                        color_depth = 1;
                        color_found = true;
                    } else if color_depth > 0 {
                        color_depth += 1;
                    }
                }
                writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            Event::Empty(element) => {
                let tag = local_name(element.name().as_ref());
                if tag == "cNvPr" { current_shape = get_attr(element, "name"); }
                if skip_depth > 0 { continue; }
                if active_shape.is_some() {
                    if tag == "alpha" && get_attr(element, "val").as_deref() == Some("0") {
                        if edits[active_shape.as_ref().unwrap()].after.transparent_background { alpha_zero_found = true; }
                        else { continue; }
                    }
                    if color_depth == 0 && matches!(tag.as_str(), "srgbClr" | "schemeClr" | "sysClr" | "prstClr") {
                        color_found = true;
                        if edits[active_shape.as_ref().unwrap()].after.transparent_background && !alpha_zero_found {
                            let qualified = element.name();
                            let name = std::str::from_utf8(qualified.as_ref()).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                            writer.write_event(Event::Start(element.to_owned())).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                            write_alpha_zero(&mut writer)?;
                            writer.write_event(Event::End(BytesEnd::new(name))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                            alpha_zero_found = true;
                            continue;
                        }
                    }
                }
                writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            Event::End(element) => {
                let tag = local_name(element.name().as_ref());
                if skip_depth > 0 { skip_depth -= 1; continue; }
                if active_shape.is_some() {
                    if color_depth > 0 {
                        color_depth -= 1;
                        if color_depth == 0 && edits[active_shape.as_ref().unwrap()].after.transparent_background && !alpha_zero_found {
                            write_alpha_zero(&mut writer)?;
                            alpha_zero_found = true;
                        }
                    }
                    if tag == "spPr" && sppr_depth == 1 {
                        if edits[active_shape.as_ref().unwrap()].after.transparent_background && !color_found {
                            write_transparent_fill(&mut writer)?;
                        }
                        active_shape = None;
                        sppr_depth = 0;
                    } else {
                        sppr_depth = sppr_depth.saturating_sub(1);
                    }
                }
                writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                if matches!(tag.as_str(), "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" | "relSizeAnchor" | "absSizeAnchor") { current_shape = None; }
            }
            Event::Eof => break,
            _ => writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?,
        }
    }
    String::from_utf8(writer.into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn patch_drawing_line_xml(
    xml: &str,
    edits: &HashMap<String, Option<crate::ir::ShapeLine>>,
) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len() + edits.len() * 128));
    let mut current_shape: Option<String> = None;
    let mut in_line = false;
    let mut line_depth = 0usize;
    let mut skip_fill_depth = 0usize;
    let mut hit: HashMap<String, usize> = HashMap::new();

    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        match &event {
            Event::Start(start) => {
                let tag = local_name(start.name().as_ref());
                if tag == "cNvPr" {
                    current_shape = get_attr(start, "name");
                }
                let selected = current_shape.as_ref().and_then(|name| edits.get(name));
                if tag == "ln" && !in_line {
                    if let (Some(shape), Some(_after)) = (current_shape.as_ref(), selected) {
                        *hit.entry(shape.clone()).or_default() += 1;
                        in_line = true;
                        line_depth = 0;
                        skip_fill_depth = 0;
                        writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        continue;
                    }
                }
                if in_line {
                    if skip_fill_depth > 0 {
                        skip_fill_depth += 1;
                        continue;
                    }
                    if line_depth == 0 && matches!(tag.as_str(), "solidFill" | "pattFill") {
                        skip_fill_depth = 1;
                        continue;
                    }
                    line_depth += 1;
                }
                writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            }
            Event::Empty(empty) => {
                let tag = local_name(empty.name().as_ref());
                if tag == "cNvPr" { current_shape = get_attr(empty, "name"); }
                if tag == "ln" && !in_line {
                    if let (Some(shape), Some(after)) = (current_shape.as_ref(), current_shape.as_ref().and_then(|name| edits.get(name))) {
                        *hit.entry(shape.clone()).or_default() += 1;
                        let qualified_name = empty.name();
                        let line_tag = std::str::from_utf8(qualified_name.as_ref()).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        let prefix = line_tag.rsplit_once(':').map_or("", |(prefix, _)| prefix);
                        writer.write_event(Event::Start(BytesStart::new(line_tag))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        write_line_fill_option(&mut writer, after.as_ref(), prefix)?;
                        writer.write_event(Event::End(BytesEnd::new(line_tag))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        continue;
                    }
                }
                if in_line && (skip_fill_depth > 0 || (line_depth == 0 && matches!(tag.as_str(), "solidFill" | "pattFill"))) {
                    continue;
                }
                writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            }
            Event::End(end) => {
                let tag = local_name(end.name().as_ref());
                if in_line && tag == "ln" {
                    if let Some(after) = current_shape.as_ref().and_then(|name| edits.get(name)) {
                        let qualified_name = end.name();
                        let line_tag = std::str::from_utf8(qualified_name.as_ref()).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        let prefix = line_tag.rsplit_once(':').map_or("", |(prefix, _)| prefix);
                        write_line_fill_option(&mut writer, after.as_ref(), prefix)?;
                    }
                    writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    in_line = false;
                    line_depth = 0;
                    skip_fill_depth = 0;
                    continue;
                }
                if skip_fill_depth > 0 {
                    skip_fill_depth -= 1;
                    continue;
                }
                if in_line { line_depth = line_depth.saturating_sub(1); }
                writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                if matches!(tag.as_str(), "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" | "relSizeAnchor" | "absSizeAnchor") {
                    current_shape = None;
                }
            }
            Event::Eof => break,
            _ => {
                if skip_fill_depth == 0 {
                    writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                }
            }
        }
    }
    for name in edits.keys() {
        match hit.get(name).copied().unwrap_or(0) {
            1 => {}
            0 => return Err(XlsxError::InvalidData(format!("drawing shape {name:?} was not found while saving its line"))),
            _ => return Err(XlsxError::InvalidData(format!("drawing shape {name:?} is ambiguous while saving its line"))),
        }
    }
    String::from_utf8(writer.into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn patch_drawing_fill_xml(
    xml: &str,
    edits: &HashMap<String, (Option<String>, Option<usize>, Option<u8>, Option<f32>)>,
) -> Result<String, XlsxError> {
    fn write_fill<W: Write>(writer: &mut Writer<W>, fill: Option<&(Option<String>, Option<usize>, Option<u8>, Option<f32>)>, prefix: &str) -> Result<(), XlsxError> {
        let qualified = |local: &str| if prefix.is_empty() { local.to_owned() } else { format!("{prefix}:{local}") };
        let (outer, inner) = match fill {
            Some((Some(_), _, _, _)) => ("solidFill", "srgbClr"),
            Some((None, _, _, _)) | None => ("noFill", ""),
        };
        let outer = qualified(outer);
        writer.write_event(Event::Start(BytesStart::new(outer.as_str())))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        if let Some((Some(color), theme, legacy, brightness)) = fill {
            let (tag, scheme) = theme.and_then(drawing_theme_token).map_or((inner, None), |scheme| ("schemeClr", Some(scheme)));
            let mut element = BytesStart::new(qualified(tag));
            element.push_attribute(("val", scheme.unwrap_or(color.as_str())));
            if let Some(index) = legacy {
                element.push_attribute(("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"));
                element.push_attribute(("mc:Ignorable", "a14"));
                element.push_attribute(("xmlns:a14", "http://schemas.microsoft.com/office/drawing/2010/main"));
                let index = index.to_string();
                element.push_attribute(("a14:legacySpreadsheetColorIndex", index.as_str()));
            }
            if let Some(brightness) = brightness.filter(|value| value.abs() > f32::EPSILON) {
                writer.write_event(Event::Start(element))
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                let values = if brightness > 0.0 {
                    vec![("lumMod", ((1.0 - brightness) * 100_000.0).round() as i32), ("lumOff", (brightness * 100_000.0).round() as i32)]
                } else { vec![("lumMod", ((1.0 + brightness) * 100_000.0).round() as i32)] };
                for (name, value) in values {
                    let mut modifier = BytesStart::new(qualified(name));
                    let value = value.to_string();
                    modifier.push_attribute(("val", value.as_str()));
                    writer.write_event(Event::Empty(modifier)).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                }
                let tag = qualified(tag);
                writer.write_event(Event::End(BytesEnd::new(tag))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            } else {
                writer.write_event(Event::Empty(element)).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
        }
        writer.write_event(Event::End(BytesEnd::new(outer)))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))
    }

    let prefix = {
        let mut reader = Reader::from_str(xml);
        let mut prefix = "a".to_owned();
        loop {
            match reader.read_event().map_err(XlsxError::Xml)? {
                Event::Start(element) | Event::Empty(element) if local_name(element.name().as_ref()) == "prstGeom" => {
                    let qualified = element.name();
                    let name = std::str::from_utf8(qualified.as_ref()).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    prefix = name.rsplit_once(':').map_or("", |(prefix, _)| prefix).to_owned();
                    break;
                }
                Event::Eof => break,
                _ => {}
            }
        }
        prefix
    };
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len() + edits.len() * 64));
    let mut current_shape: Option<String> = None;
    let mut active_shape: Option<String> = None;
    let mut sppr_depth = 0usize;
    let mut skip_depth = 0usize;
    let mut hits: HashMap<String, usize> = HashMap::new();
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        match &event {
            Event::Start(element) => {
                let tag = local_name(element.name().as_ref());
                if tag == "cNvPr" { current_shape = get_attr(element, "name"); }
                if skip_depth > 0 {
                    skip_depth += 1;
                    continue;
                }
                if tag == "spPr" && sppr_depth == 0 {
                    active_shape = current_shape.as_ref().filter(|name| edits.contains_key(*name)).cloned();
                    sppr_depth = 1;
                    writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    continue;
                }
                if sppr_depth > 0 {
                    if sppr_depth == 1 && matches!(tag.as_str(), "solidFill" | "noFill" | "pattFill" | "gradFill" | "blipFill") {
                        if let Some(shape) = active_shape.as_ref() {
                            *hits.entry(shape.clone()).or_default() += 1;
                            write_fill(&mut writer, edits.get(shape), &prefix)?;
                            skip_depth = 1;
                            continue;
                        }
                    }
                    sppr_depth += 1;
                }
                writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            Event::Empty(element) => {
                let tag = local_name(element.name().as_ref());
                if tag == "cNvPr" { current_shape = get_attr(element, "name"); }
                if skip_depth > 0 { continue; }
                if sppr_depth == 0 && tag == "spPr" {
                    active_shape = current_shape.as_ref().filter(|name| edits.contains_key(*name)).cloned();
                    if let Some(shape) = active_shape.take() {
                        let qualified = element.name();
                        let tag = std::str::from_utf8(qualified.as_ref()).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                        writer.write_event(Event::Start(BytesStart::new(tag))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                        write_fill(&mut writer, edits.get(&shape), &prefix)?;
                        writer.write_event(Event::End(BytesEnd::new(tag))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                        *hits.entry(shape).or_default() += 1;
                    } else {
                        writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    }
                    continue;
                }
                if sppr_depth == 1 && matches!(tag.as_str(), "solidFill" | "noFill" | "pattFill" | "gradFill" | "blipFill") {
                    if let Some(shape) = active_shape.as_ref() {
                        *hits.entry(shape.clone()).or_default() += 1;
                        write_fill(&mut writer, edits.get(shape), &prefix)?;
                        continue;
                    }
                }
                writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            Event::End(element) => {
                let tag = local_name(element.name().as_ref());
                if skip_depth > 0 {
                    skip_depth -= 1;
                    continue;
                }
                if tag == "spPr" && sppr_depth == 1 {
                    if let Some(shape) = active_shape.take() {
                        let count = hits.get(&shape).copied().unwrap_or(0);
                        if count == 0 {
                            write_fill(&mut writer, edits.get(&shape), &prefix)?;
                            *hits.entry(shape).or_default() += 1;
                        } else if count > 1 {
                            return Err(XlsxError::InvalidData(format!("shape {shape:?} has multiple direct fill elements")));
                        }
                    }
                    sppr_depth = 0;
                    writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    continue;
                }
                if sppr_depth > 0 { sppr_depth = sppr_depth.saturating_sub(1); }
                writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                if matches!(tag.as_str(), "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" | "relSizeAnchor" | "absSizeAnchor") { current_shape = None; }
            }
            Event::Eof => break,
            _ => writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?,
        }
    }
    for name in edits.keys() {
        match hits.get(name).copied().unwrap_or(0) {
            1 => {}
            0 => return Err(XlsxError::InvalidData(format!("drawing shape {name:?} was not found while saving its fill"))),
            _ => return Err(XlsxError::InvalidData(format!("drawing shape {name:?} is ambiguous while saving its fill"))),
        }
    }
    String::from_utf8(writer.into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn patch_drawing_text_xml(
    xml: &str,
    edits: &HashMap<String, ShapeTextEdit>,
) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len() + edits.len() * 256));
    let mut current_shape: Option<String> = None;
    let mut hits: HashMap<String, usize> = HashMap::new();
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        match &event {
            Event::Start(element) if local_name(element.name().as_ref()) == "cNvPr" => {
                current_shape = get_attr(element, "name");
                writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            Event::Empty(element) if local_name(element.name().as_ref()) == "cNvPr" => {
                current_shape = get_attr(element, "name");
                writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            Event::Start(element) if local_name(element.name().as_ref()) == "txBody" => {
                let Some(name) = current_shape.as_ref().filter(|name| edits.contains_key(*name)).cloned() else {
                    writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    continue;
                };
                let mut body = vec![event.into_owned()];
                let mut depth = 1usize;
                while depth > 0 {
                    let next = reader.read_event().map_err(XlsxError::Xml)?.into_owned();
                    match &next {
                        Event::Start(_) => depth += 1,
                        Event::End(_) => depth -= 1,
                        Event::Eof => return Err(XlsxError::InvalidData("shape text body ended unexpectedly".into())),
                        _ => {}
                    }
                    body.push(next);
                }
                *hits.entry(name.clone()).or_default() += 1;
                let edit = edits.get(&name).expect("shape text edit was selected by name");
                let replacement = write_shape_text(Some(&edit.after))
                    .replacen("<xdr:txBody>", "<xdr:txBody xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">", 1);
                let mut replacement_reader = Reader::from_str(&replacement);
                let mut replacement_events = Vec::new();
                loop {
                    let replacement_event = replacement_reader.read_event().map_err(XlsxError::Xml)?;
                    if matches!(replacement_event, Event::Eof) { break; }
                    replacement_events.push(replacement_event.into_owned());
                }
                for replacement_event in preserve_shape_text_metadata(&body, &replacement_events, &edit.before, &edit.after)? {
                    writer.write_event(replacement_event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                }
            }
            Event::End(element) => {
                let tag = local_name(element.name().as_ref());
                writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                if matches!(tag.as_str(), "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" | "relSizeAnchor" | "absSizeAnchor") {
                    current_shape = None;
                }
            }
            Event::Eof => break,
            _ => writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?,
        }
    }
    for name in edits.keys() {
        match hits.get(name).copied().unwrap_or(0) {
            1 => {}
            0 => return Err(XlsxError::InvalidData(format!("drawing shape {name:?} text was not found while saving its formatting"))),
            _ => return Err(XlsxError::InvalidData(format!("drawing shape {name:?} text is ambiguous while saving its formatting"))),
        }
    }
    String::from_utf8(writer.into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn drawing_text_subtree(events: &[Event<'static>], wanted: &str, occurrence: usize) -> Option<Vec<Event<'static>>> {
    let mut found = 0usize;
    for (index, event) in events.iter().enumerate() {
        let start = match event {
            Event::Start(start) | Event::Empty(start) if local_name(start.name().as_ref()) == wanted => start,
            _ => continue,
        };
        if found != occurrence {
            found += 1;
            continue;
        }
        if matches!(event, Event::Empty(_)) {
            return Some(vec![event.clone()]);
        }
        let _ = start;
        let mut depth = 1usize;
        for end in index + 1..events.len() {
            match &events[end] {
                Event::Start(_) => depth += 1,
                Event::End(_) => depth -= 1,
                _ => {}
            }
            if depth == 0 {
                return Some(events[index..=end].to_vec());
            }
        }
        return None;
    }
    None
}

fn merge_drawing_text_attributes(
    original: &[Event<'static>],
    replacement: &[Event<'static>],
    overridden: &[&str],
) -> Result<Vec<Event<'static>>, XlsxError> {
    merge_drawing_text_attributes_mode(original, replacement, overridden, false)
}

fn merge_drawing_text_attributes_mode(
    original: &[Event<'static>],
    replacement: &[Event<'static>],
    overridden: &[&str],
    remove_missing: bool,
) -> Result<Vec<Event<'static>>, XlsxError> {
    let (Some(original_start), Some(replacement_start)) = (original.first(), replacement.first()) else {
        return Ok(replacement.to_vec());
    };
    let original_start = match original_start {
        Event::Start(start) | Event::Empty(start) => start,
        _ => return Ok(replacement.to_vec()),
    };
    let replacement_start = match replacement_start {
        Event::Start(start) | Event::Empty(start) => start,
        _ => return Ok(replacement.to_vec()),
    };
    let mut desired = HashMap::new();
    for key in overridden {
        if let Some(value) = get_attr(replacement_start, key) {
            desired.insert((*key).to_string(), value);
        }
    }
    let name = std::str::from_utf8(original_start.name().as_ref())
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?.to_string();
    let mut merged = BytesStart::new(name.as_str());
    let mut copied = std::collections::HashSet::new();
    for attribute in original_start.attributes().flatten() {
        let key = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?
            .to_string();
        let local = local_name(key.as_bytes());
        if overridden.contains(&local.as_str()) {
            if let Some(value) = desired.get(&local) {
                merged.push_attribute((key.as_str(), value.as_str()));
                copied.insert(local);
                continue;
            } else if remove_missing {
                continue;
            }
        }
        let raw = String::from_utf8_lossy(attribute.value.as_ref());
        let value = quick_xml::escape::unescape(&raw)
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        merged.push_attribute((key.as_str(), value.as_ref()));
    }
    for (key, value) in desired {
        if copied.contains(&key) { continue; }
        merged.push_attribute((key.as_str(), value.as_str()));
    }
    let first = if matches!(original.first(), Some(Event::Empty(_))) {
        Event::Empty(merged.into_owned())
    } else {
        Event::Start(merged.into_owned())
    };
    let mut result = original.to_vec();
    result[0] = first;
    Ok(result)
}

fn drawing_event_subtree_end(events: &[Event<'static>], index: usize) -> Result<usize, XlsxError> {
    if !matches!(events.get(index), Some(Event::Start(_))) { return Ok(index); }
    let mut depth = 1usize;
    let mut end = index;
    while depth > 0 {
        end += 1;
        match events.get(end) {
            Some(Event::Start(_)) => depth += 1,
            Some(Event::End(_)) => depth -= 1,
            Some(_) => {}
            None => return Err(XlsxError::InvalidData("shape text property ended unexpectedly".into())),
        }
    }
    Ok(end)
}

fn drawing_run_text(events: &[Event<'static>]) -> Result<String, XlsxError> {
    let mut text = String::new();
    let mut in_text = false;
    for event in events {
        match event {
            Event::Start(start) if local_name(start.name().as_ref()) == "t" => in_text = true,
            Event::End(end) if local_name(end.name().as_ref()) == "t" => in_text = false,
            Event::Text(value) if in_text => text.push_str(&value.unescape().map_err(XlsxError::Xml)?),
            _ => {}
        }
    }
    Ok(text)
}

fn source_run_properties(events: &[Event<'static>], paragraph: usize) -> Result<Vec<(usize, usize, Vec<Event<'static>>)>, XlsxError> {
    let Some(source_paragraph) = drawing_text_subtree(events, "p", paragraph) else { return Ok(Vec::new()); };
    let mut runs = Vec::new();
    let mut position = 0usize;
    let mut index = 1usize;
    while index + 1 < source_paragraph.len() {
        let event = &source_paragraph[index];
        let tag = match event {
            Event::Start(start) | Event::Empty(start) => local_name(start.name().as_ref()),
            _ => String::new(),
        };
        if !matches!(tag.as_str(), "r" | "br" | "fld") {
            index += 1;
            continue;
        }
        let end = drawing_event_subtree_end(&source_paragraph, index)?;
        let subtree = &source_paragraph[index..=end];
        let length = if tag == "br" { 1 } else { drawing_run_text(subtree)?.chars().count() };
        if length > 0 {
            if let Some(properties) = drawing_text_subtree(subtree, "rPr", 0) {
                runs.push((position, position + length, properties));
            }
        }
        position += length;
        index = end + 1;
    }
    Ok(runs)
}

#[derive(Debug, Clone, PartialEq)]
struct ShapeRunStyle {
    bold: bool,
    italic: bool,
    underline: bool,
    face: Option<String>,
    size: f32,
    color: Option<String>,
    color_theme: Option<usize>,
    color_brightness: Option<f32>,
    legacy_color_index: Option<u8>,
}

fn shape_run_style_at(text: &crate::ir::ShapeText, paragraph_index: usize, position: usize) -> Option<ShapeRunStyle> {
    let paragraph = text.paragraphs.get(paragraph_index)?;
    let style = |bold, italic, underline, face: Option<String>, size, color: Option<&str>, color_theme, color_brightness, legacy_color_index| ShapeRunStyle {
        bold, italic, underline, face, size,
        color: color.map(str::to_ascii_uppercase),
        color_theme,
        color_brightness,
        legacy_color_index,
    };
    if paragraph.runs.is_empty() {
        return (position < paragraph.text.chars().count()).then(|| style(
            paragraph.bold, paragraph.italic, false, paragraph.face.clone(), paragraph.size, paragraph.color.as_deref(), None, None, None,
        ));
    }
    let mut start = 0usize;
    for run in &paragraph.runs {
        let end = start + run.text.chars().count();
        if start <= position && position < end {
            return Some(style(
                run.bold, run.italic.unwrap_or(paragraph.italic), run.underline,
                run.face.clone().or_else(|| paragraph.face.clone()), run.size.unwrap_or(paragraph.size),
                run.color.as_deref().or(paragraph.color.as_deref()),
                run.color_theme,
                run.color_brightness,
                run.legacy_color_index,
            ));
        }
        start = end;
    }
    None
}

fn merge_drawing_run_properties(
    original: &[Event<'static>],
    replacement: &[Event<'static>],
    changed_attributes: &[&str],
    changed_children: &[&str],
) -> Result<Vec<Event<'static>>, XlsxError> {
    const RUN_CHILDREN: &[&str] = &["solidFill", "latin", "ea"];
    if changed_attributes.is_empty() && changed_children.is_empty() {
        return Ok(original.to_vec());
    }
    let start = merge_drawing_text_attributes_mode(original, replacement, changed_attributes, true)?;
    let mut result = vec![match &start[0] {
        Event::Start(element) | Event::Empty(element) => Event::Start(element.clone().into_owned()),
        event => event.clone(),
    }];
    let generated_children = ["solidFill", "latin", "ea"].map(|tag| drawing_text_subtree(replacement, tag, 0));
    let mut generated_used = [false; 3];
    let mut index = 1usize;
    while index + 1 < original.len() {
        let event = &original[index];
        let tag = match event {
            Event::Start(child) | Event::Empty(child) => local_name(child.name().as_ref()),
            _ => String::new(),
        };
        let end = drawing_event_subtree_end(original, index)?;
        if let Some(child_index) = RUN_CHILDREN.iter().position(|known| *known == tag)
            .filter(|_| changed_children.contains(&tag.as_str()))
        {
            if !generated_used[child_index] {
                if let Some(generated) = &generated_children[child_index] {
                    result.extend(generated.clone());
                    generated_used[child_index] = true;
                }
            }
        } else {
            result.extend_from_slice(&original[index..=end]);
        }
        index = end + 1;
    }
    for (child_index, generated) in generated_children.iter().enumerate() {
        if !generated_used[child_index] && changed_children.contains(&RUN_CHILDREN[child_index]) {
            if let Some(generated) = generated { result.extend(generated.clone()); }
        }
    }
    if let Some(end) = original.last().filter(|event| matches!(event, Event::End(_))) {
        result.push(end.clone());
    } else if let Some(Event::End(end)) = replacement.last() {
        result.push(Event::End(end.clone()));
    }
    Ok(result)
}

fn preserve_drawing_run_properties(
    original: &[Event<'static>],
    replacement: &[Event<'static>],
    paragraph: usize,
    position: usize,
    before_text: &crate::ir::ShapeText,
    after_text: &crate::ir::ShapeText,
) -> Result<Vec<Event<'static>>, XlsxError> {
    let Some(source) = source_run_properties(original, paragraph)?
        .into_iter().find(|(start, end, _)| *start <= position && position < *end)
        .map(|(_, _, properties)| properties) else {
        return Ok(replacement.to_vec());
    };
    let Some(generated) = drawing_text_subtree(replacement, "rPr", 0) else { return Ok(replacement.to_vec()); };
    let before = shape_run_style_at(before_text, paragraph, position);
    let after = shape_run_style_at(after_text, paragraph, position);
    let mut attributes = Vec::new();
    let mut children = Vec::new();
    if before.as_ref().map(|style| style.bold) != after.as_ref().map(|style| style.bold) { attributes.push("b"); }
    if before.as_ref().map(|style| style.italic) != after.as_ref().map(|style| style.italic) { attributes.push("i"); }
    if before.as_ref().map(|style| style.underline) != after.as_ref().map(|style| style.underline) { attributes.push("u"); }
    if before.as_ref().map(|style| style.size) != after.as_ref().map(|style| style.size) { attributes.push("sz"); }
    if before.as_ref().and_then(|style| style.color.as_ref()) != after.as_ref().and_then(|style| style.color.as_ref()) { children.push("solidFill"); }
    if before.as_ref().and_then(|style| style.color_theme) != after.as_ref().and_then(|style| style.color_theme)
        || before.as_ref().and_then(|style| style.color_brightness) != after.as_ref().and_then(|style| style.color_brightness)
        || before.as_ref().and_then(|style| style.legacy_color_index) != after.as_ref().and_then(|style| style.legacy_color_index)
    { children.push("solidFill"); }
    if before.as_ref().and_then(|style| style.face.as_ref()) != after.as_ref().and_then(|style| style.face.as_ref()) {
        children.extend(["latin", "ea"]);
    }
    let merged = merge_drawing_run_properties(&source, &generated, &attributes, &children)?;
    let mut result = Vec::with_capacity(replacement.len() + merged.len());
    let mut index = 0usize;
    while index < replacement.len() {
        let tag = match &replacement[index] {
            Event::Start(start) | Event::Empty(start) => local_name(start.name().as_ref()),
            _ => String::new(),
        };
        if tag == "rPr" {
            let end = drawing_event_subtree_end(replacement, index)?;
            result.extend_from_slice(&merged);
            index = end + 1;
        } else {
            result.push(replacement[index].clone());
            index += 1;
        }
    }
    Ok(result)
}

fn preserve_shape_text_metadata(
    original: &[Event<'static>],
    replacement: &[Event<'static>],
    before_text: &crate::ir::ShapeText,
    after_text: &crate::ir::ShapeText,
) -> Result<Vec<Event<'static>>, XlsxError> {
    let mut result = Vec::with_capacity(replacement.len());
    let mut paragraph = 0usize;
    let mut run_position = 0usize;
    let mut index = 0usize;
    while index < replacement.len() {
        let event = &replacement[index];
        let tag = match event {
            Event::Start(start) | Event::Empty(start) => local_name(start.name().as_ref()),
            _ => String::new(),
        };
        if tag == "p" { paragraph += 1; run_position = 0; }
        let is_metadata = matches!(tag.as_str(), "bodyPr" | "lstStyle" | "pPr" | "endParaRPr" | "r" | "br");
        let end = match event {
            Event::Start(_) if is_metadata => {
                let mut depth = 1usize;
                let mut end = index;
                while depth > 0 {
                    end += 1;
                    match replacement.get(end) {
                        Some(Event::Start(_)) => depth += 1,
                        Some(Event::End(_)) => depth -= 1,
                        Some(_) => {}
                        None => return Err(XlsxError::InvalidData("replacement shape text body ended unexpectedly".into())),
                    }
                }
                end
            }
            _ => index,
        };
        let generated = &replacement[index..=end];
        if tag == "r" || tag == "br" {
            let length = if tag == "br" { 1 } else { drawing_run_text(generated)?.chars().count() };
            result.extend(preserve_drawing_run_properties(
                original, generated, paragraph.saturating_sub(1), run_position, before_text, after_text,
            )?);
            run_position += length;
            index = end + 1;
            continue;
        } else if tag == "bodyPr" {
            if let Some(source) = drawing_text_subtree(original, "bodyPr", 0) {
                result.extend(merge_drawing_text_attributes(
                    &source, generated,
                    &["wrap", "vertOverflow", "lIns", "tIns", "rIns", "bIns", "anchor"],
                )?);
                index = end + 1;
                continue;
            }
        } else if tag == "lstStyle" {
            if let Some(source) = drawing_text_subtree(original, "lstStyle", 0) {
                result.extend(source);
                index = end + 1;
                continue;
            }
        } else if tag == "pPr" {
            if let Some(source) = drawing_text_subtree(original, "pPr", paragraph.saturating_sub(1)) {
                result.extend(merge_drawing_text_attributes(&source, generated, &["algn"])?);
                index = end + 1;
                continue;
            }
        } else if tag == "endParaRPr" {
            if let Some(source) = drawing_text_subtree(original, "endParaRPr", paragraph.saturating_sub(1)) {
                result.extend(source);
                index = end + 1;
                continue;
            }
        }
        result.extend_from_slice(generated);
        index = end + 1;
    }
    Ok(result)
}

fn patch_drawing_deletions_xml(
    xml: &str,
    deletions: &std::collections::HashSet<String>,
    ungrouped: Option<&HashMap<String, crate::ir::Drawing>>,
    drawings: Option<&HashMap<String, crate::ir::Drawing>>,
) -> Result<String, XlsxError> {
    fn span(events: &[Event<'static>], start: usize) -> usize {
        match events.get(start) {
            Some(Event::Empty(_)) => start + 1,
            Some(Event::Start(_)) => {
                let mut depth = 0usize;
                for (index, event) in events.iter().enumerate().skip(start) {
                    match event {
                        Event::Start(_) => depth += 1,
                        Event::End(_) => {
                            depth -= 1;
                            if depth == 0 { return index + 1; }
                        }
                        _ => {}
                    }
                }
                events.len()
            }
            _ => start + 1,
        }
    }
    fn named_object(events: &[Event<'static>], start: usize, end: usize) -> Option<String> {
        events[start..end].iter().find_map(|event| match event {
            Event::Start(element) | Event::Empty(element) if local_name(element.name().as_ref()) == "cNvPr" => get_attr(element, "name"),
            _ => None,
        })
    }
    fn group_children(
        events: &[Event<'static>],
        group_start: usize,
        group_end: usize,
    ) -> Vec<(usize, usize, Option<String>)> {
        let mut children = Vec::new();
        let mut index = group_start + 1;
        while index + 1 < group_end {
            let object = match &events[index] {
                Event::Start(element) | Event::Empty(element) => matches!(
                    local_name(element.name().as_ref()).as_str(),
                    "sp" | "pic" | "graphicFrame" | "cxnSp" | "grpSp"
                ),
                _ => false,
            };
            if object {
                let end = span(events, index);
                children.push((index, end, named_object(events, index, end)));
                index = end;
            } else {
                index += 1;
            }
        }
        children
    }
    fn grouped_child_frame(events: &[Event<'static>]) -> Option<((i64, i64), (i64, i64))> {
        let mut in_group_props = false;
        let mut in_xfrm = false;
        let mut ch_off = None;
        let mut ch_ext = None;
        let mut group_depth = 0usize;
        for event in events {
            let element = match event {
                Event::Start(element) | Event::Empty(element) => element,
                Event::End(end) => {
                    if local_name(end.name().as_ref()) == "grpSp" { group_depth = group_depth.saturating_sub(1); }
                    if local_name(end.name().as_ref()) == "xfrm" { in_xfrm = false; }
                    if local_name(end.name().as_ref()) == "grpSpPr" { in_group_props = false; }
                    continue;
                }
                _ => continue,
            };
            match local_name(element.name().as_ref()).as_str() {
                "grpSp" => group_depth += 1,
                "grpSpPr" if group_depth == 1 => in_group_props = true,
                "xfrm" if in_group_props => in_xfrm = true,
                "chOff" if group_depth == 1 && in_group_props && in_xfrm => {
                    ch_off = Some((
                        get_attr(element, "x")?.parse().ok()?,
                        get_attr(element, "y")?.parse().ok()?,
                    ));
                }
                "chExt" if group_depth == 1 && in_group_props && in_xfrm => {
                    ch_ext = Some((
                        get_attr(element, "cx")?.parse().ok()?,
                        get_attr(element, "cy")?.parse().ok()?,
                    ));
                }
                _ => {}
            }
        }
        Some((ch_off?, ch_ext?))
    }
    fn normalize_grouped_child(
        events: &mut [Event<'static>],
        drawing: &crate::ir::Drawing,
        child_origin: (i64, i64),
        child_extent: (i64, i64),
    ) -> Result<(), XlsxError> {
        fn replace_attributes(event: &mut Event<'static>, updates: &[(&str, &str)]) {
            let (source, empty) = match event {
                Event::Start(element) => (element, false),
                Event::Empty(element) => (element, true),
                _ => return,
            };
            let mut changed = source.to_owned();
            changed.clear_attributes();
            for attribute in source.attributes().flatten() {
                if !updates.iter().any(|(name, _)| attribute.key.as_ref() == name.as_bytes()) {
                    changed.push_attribute(attribute);
                }
            }
            for (name, value) in updates { changed.push_attribute((*name, *value)); }
            *event = if empty { Event::Empty(changed) } else { Event::Start(changed) };
        }
        let frame = drawing.frame.ok_or_else(|| XlsxError::InvalidData(format!("nested shape {:?} has no root-group frame", drawing.name)))?;
        let rotation = match &drawing.kind {
            crate::ir::DrawingKind::Shape(shape) => shape.rotation,
            _ => 0,
        };
        let off_x = child_origin.0 + (frame.x * child_extent.0 as f64).round() as i64;
        let off_y = child_origin.1 + (frame.y * child_extent.1 as f64).round() as i64;
        let ext_x = (frame.w * child_extent.0 as f64).round() as i64;
        let ext_y = (frame.h * child_extent.1 as f64).round() as i64;
        let rotation = rotation.to_string();
        let mut in_sp_pr = false;
        let mut in_xfrm = false;
        let mut off_done = false;
        let mut ext_done = false;
        for event in events {
            if matches!(event, Event::End(end) if local_name(end.name().as_ref()) == "xfrm") {
                in_xfrm = false;
                continue;
            }
            if matches!(event, Event::End(end) if local_name(end.name().as_ref()) == "spPr") {
                in_sp_pr = false;
                continue;
            }
            let element = match event {
                Event::Start(element) | Event::Empty(element) => element,
                _ => continue,
            };
            let tag = local_name(element.name().as_ref());
            if tag == "spPr" { in_sp_pr = true; }
            if tag == "xfrm" && in_sp_pr {
                replace_attributes(event, &[("rot", rotation.as_str())]);
                in_xfrm = true;
                continue;
            }
            if !in_xfrm { continue; }
            if tag == "off" && !off_done {
                replace_attributes(event, &[("x", &off_x.to_string()), ("y", &off_y.to_string())]);
                off_done = true;
            } else if tag == "ext" && !ext_done {
                replace_attributes(event, &[("cx", &ext_x.to_string()), ("cy", &ext_y.to_string())]);
                ext_done = true;
            }
        }
        if !off_done || !ext_done {
            return Err(XlsxError::InvalidData(format!("nested shape {:?} has no complete DrawingML transform", drawing.name)));
        }
        Ok(())
    }
    fn replace_attributes(event: &mut Event<'static>, updates: &[(&str, &str)]) {
        let (source, empty) = match event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            _ => return,
        };
        let mut changed = source.to_owned();
        changed.clear_attributes();
        for attribute in source.attributes().flatten() {
            if !updates.iter().any(|(name, _)| attribute.key.as_ref() == name.as_bytes()) {
                changed.push_attribute(attribute);
            }
        }
        for (name, value) in updates { changed.push_attribute((*name, *value)); }
        *event = if empty { Event::Empty(changed) } else { Event::Start(changed) };
    }
    fn update_group_transform(events: &mut [Event<'static>], drawing: &crate::ir::Drawing) {
        let Some(transform) = drawing.group_transform else { return };
        let mut group_depth = 0usize;
        let mut in_group_props = false;
        let mut in_xfrm = false;
        for event in events {
            if matches!(event, Event::End(end) if local_name(end.name().as_ref()) == "xfrm") {
                in_xfrm = false;
                continue;
            }
            if matches!(event, Event::End(end) if local_name(end.name().as_ref()) == "grpSpPr") {
                in_group_props = false;
                continue;
            }
            if matches!(event, Event::End(end) if local_name(end.name().as_ref()) == "grpSp") {
                group_depth = group_depth.saturating_sub(1);
                continue;
            }
            let element = match event {
                Event::Start(element) | Event::Empty(element) => element,
                _ => continue,
            };
            let tag = local_name(element.name().as_ref());
            if tag == "grpSp" { group_depth += 1; }
            if tag == "grpSpPr" && group_depth == 1 { in_group_props = true; }
            if tag == "xfrm" && in_group_props { in_xfrm = true; continue; }
            if !in_xfrm { continue; }
            match tag.as_str() {
                "off" => replace_attributes(event, &[("x", &transform.offset.0.to_string()), ("y", &transform.offset.1.to_string())]),
                "ext" => replace_attributes(event, &[("cx", &transform.extent.0.to_string()), ("cy", &transform.extent.1.to_string())]),
                "chOff" => replace_attributes(event, &[("x", &transform.child_offset.0.to_string()), ("y", &transform.child_offset.1.to_string())]),
                "chExt" => replace_attributes(event, &[("cx", &transform.child_extent.0.to_string()), ("cy", &transform.child_extent.1.to_string())]),
                _ => {}
            }
        }
    }
    fn update_anchor_values(events: &mut [Event<'static>], drawing: &crate::ir::Drawing) {
        let mut anchor: Option<crate::ir::Anchor> = None;
        let mut field: Option<String> = None;
        for event in events {
            match event {
                Event::Start(element) => match local_name(element.name().as_ref()).as_str() {
                    "from" => { anchor = Some(drawing.from); field = None; }
                    "to" => { anchor = drawing.to; field = None; }
                    "col" | "colOff" | "row" | "rowOff" if anchor.is_some() => field = Some(local_name(element.name().as_ref())),
                    _ => {}
                },
                Event::End(end) if matches!(local_name(end.name().as_ref()).as_str(), "from" | "to") => {
                    anchor = None;
                    field = None;
                }
                Event::Text(_) if anchor.is_some() => {
                    let held = anchor.as_mut().unwrap();
                    let value = match field.take() {
                        Some(field) if field == "col" => held.col.to_string(),
                        Some(field) if field == "colOff" => held.col_off.to_string(),
                        Some(field) if field == "row" => held.row.to_string(),
                        Some(field) if field == "rowOff" => held.row_off.to_string(),
                        _ => continue,
                    };
                    *event = Event::Text(quick_xml::events::BytesText::new(&value).into_owned());
                }
                _ => {}
            }
        }
    }

    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len()));
    let mut hit: HashMap<String, usize> = HashMap::new();
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?.into_owned();
        match &event {
            Event::Start(start) if matches!(local_name(start.name().as_ref()).as_str(), "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" | "relSizeAnchor" | "absSizeAnchor") => {
                let mut anchor = vec![event];
                let mut depth = 1usize;
                while depth > 0 {
                    let next = reader.read_event().map_err(XlsxError::Xml)?.into_owned();
                    match &next {
                        Event::Start(_) => depth += 1,
                        Event::End(_) => depth -= 1,
                        Event::Eof => return Err(XlsxError::InvalidData("drawing anchor ended unexpectedly".into())),
                        _ => {}
                    }
                    anchor.push(next);
                }
                let shape_name = anchor.iter().find_map(|event| match event {
                    Event::Start(element) | Event::Empty(element) if local_name(element.name().as_ref()) == "cNvPr" => get_attr(element, "name"),
                    _ => None,
                });
                if let Some(shape_name) = shape_name.filter(|name| deletions.contains(name)) {
                    *hit.entry(shape_name).or_default() += 1;
                } else {
                    let group_start = anchor.iter().position(|event| matches!(event,
                        Event::Start(element) if local_name(element.name().as_ref()) == "grpSp"));
                    let mut patched = anchor.clone();
                    let mut remove_anchor = false;
                    if let Some(group_start) = group_start {
                        let original_group_end = span(&patched, group_start);
                        let root_name = named_object(&patched, group_start, original_group_end);
                        let root_drawing = root_name.as_ref().and_then(|root| {
                            drawings.and_then(|states| states.values().find(|drawing| drawing.group_name.as_deref() == Some(root)))
                        });
                        let root_frame = root_drawing.and_then(|drawing| drawing.group_transform.map(|transform| {
                            (transform.child_offset, transform.child_extent)
                        })).or_else(|| grouped_child_frame(&patched[group_start..original_group_end]));
                        let mut nested_changed = false;
                        let original_children = group_children(&patched, group_start, original_group_end);
                        for (child_start, child_end, child_name) in original_children.iter().rev() {
                            let is_nested_group = matches!(patched.get(*child_start), Some(Event::Start(element)) if local_name(element.name().as_ref()) == "grpSp");
                            if !is_nested_group { continue; }
                            let nested = patched[*child_start..*child_end].to_vec();
                            let nested_children = group_children(&nested, 0, nested.len());
                            let removed = nested_children
                                .iter()
                                .filter(|(_, _, name)| name.as_ref().is_some_and(|name| deletions.contains(name)))
                                .collect::<Vec<_>>();
                            if removed.is_empty() { continue; }
                            nested_changed = true;
                            for (_, _, name) in &removed {
                                if let Some(name) = name { *hit.entry(name.clone()).or_default() += 1; }
                            }
                            let survivors = nested_children
                                .iter()
                                .filter(|(_, _, name)| !name.as_ref().is_some_and(|name| deletions.contains(name)))
                                .collect::<Vec<_>>();
                            let replacement = if survivors.is_empty() {
                                Vec::new()
                            } else if survivors.len() == 1 {
                                let (_, _, name) = survivors[0];
                                let (start, end, _) = nested_children
                                    .iter()
                                    .find(|(_, _, held)| held == name)
                                    .cloned()
                                    .ok_or_else(|| XlsxError::InvalidData("the nested surviving shape disappeared".into()))?;
                                let mut shape = nested[start..end].to_vec();
                                let drawing = name.as_ref().and_then(|name| drawings.and_then(|states| states.get(name)))
                                    .ok_or_else(|| XlsxError::InvalidData("the nested surviving shape has no root-group geometry".into()))?;
                                let (root_off, root_ext) = root_frame
                                    .ok_or_else(|| XlsxError::InvalidData("the parent group has no child-coordinate transform".into()))?;
                                normalize_grouped_child(&mut shape, drawing, root_off, root_ext)?;
                                shape
                            } else {
                                let mut result = nested;
                                for (start, end, _) in removed.iter().rev() {
                                    result.drain(*start..*end);
                                }
                                result
                            };
                            patched.splice(*child_start..*child_end, replacement);
                            let _ = child_name;
                        }
                        let group_end = span(&patched, group_start);
                        let children = group_children(&patched, group_start, group_end);
                        let removed = children.iter().filter(|(_, _, name)| name.as_ref().is_some_and(|name| deletions.contains(name))).collect::<Vec<_>>();
                        if !removed.is_empty() || nested_changed {
                            for (_, _, name) in &removed {
                                if let Some(name) = name { *hit.entry(name.clone()).or_default() += 1; }
                            }
                            let survivors = children.iter().filter(|(_, _, name)| !name.as_ref().is_some_and(|name| deletions.contains(name))).collect::<Vec<_>>();
                            if survivors.is_empty() {
                                remove_anchor = true;
                            } else {
                                if nested_changed && survivors.len() > 1 {
                                    if let Some(drawing) = root_drawing {
                                        let transform = drawing.group_transform.ok_or_else(|| XlsxError::InvalidData("the surviving outer group has no transform".into()))?;
                                        for (child_start, child_end, name) in children.iter().rev() {
                                            let Some(name) = name else { continue };
                                            let Some(child_drawing) = drawings.and_then(|states| states.get(name)) else { continue };
                                            let mut shape = patched[*child_start..*child_end].to_vec();
                                            normalize_grouped_child(
                                                &mut shape,
                                                child_drawing,
                                                transform.child_offset,
                                                transform.child_extent,
                                            )?;
                                            patched.splice(*child_start..*child_end, shape);
                                        }
                                        let group_end = span(&patched, group_start);
                                        update_group_transform(&mut patched[group_start..group_end], drawing);
                                        update_anchor_values(&mut patched, drawing);
                                    }
                                }
                                let mut group = Vec::new();
                                let mut cursor = group_start;
                                for (child_start, child_end, name) in &children {
                                    if name.as_ref().is_some_and(|name| deletions.contains(name)) {
                                        group.extend_from_slice(&patched[cursor..*child_start]);
                                        cursor = *child_end;
                                    }
                                }
                                group.extend_from_slice(&patched[cursor..group_end]);
                                if survivors.len() == 1 {
                                    let (child_start, child_end, child_name) = survivors[0];
                                    let mut child = patched[*child_start..*child_end].to_vec();
                                    if let Some(drawing) = child_name.as_ref().and_then(|name| ungrouped.and_then(|edits| edits.get(name))) {
                                        normalize_ungrouped_transform(&mut child, drawing)?;
                                    } else {
                                        return Err(XlsxError::InvalidData("the surviving grouped shape has no ungrouped geometry".into()));
                                    }
                                    patched.splice(group_start..group_end, child);
                                } else {
                                    patched.splice(group_start..group_end, group);
                                }
                            }
                        }
                    }
                    if !remove_anchor {
                        for event in patched { writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?; }
                    }
                }
            }
            Event::Eof => break,
            _ => writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?,
        }
    }
    for name in deletions {
        match hit.get(name).copied().unwrap_or(0) {
            1 => {}
            0 => return Err(XlsxError::InvalidData(format!("drawing shape {name:?} was not found while deleting it"))),
            _ => return Err(XlsxError::InvalidData(format!("drawing shape {name:?} is ambiguous while deleting it"))),
        }
    }
    String::from_utf8(writer.into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn normalize_ungrouped_transform(
    events: &mut [Event<'static>],
    drawing: &crate::ir::Drawing,
) -> Result<(), XlsxError> {
    fn replace_attributes(event: &mut Event<'static>, updates: &[(&str, &str)]) {
        let (source, empty) = match event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            _ => return,
        };
        let mut changed = source.to_owned();
        changed.clear_attributes();
        for attribute in source.attributes().flatten() {
            if !updates.iter().any(|(name, _)| attribute.key.as_ref() == name.as_bytes()) {
                changed.push_attribute(attribute);
            }
        }
        for (name, value) in updates { changed.push_attribute((*name, *value)); }
        *event = if empty { Event::Empty(changed) } else { Event::Start(changed) };
    }

    let (cx, cy) = drawing.extent.ok_or_else(|| XlsxError::InvalidData("the surviving grouped shape has no extent".into()))?;
    let rotation = match &drawing.kind {
        crate::ir::DrawingKind::Shape(shape) => shape.rotation,
        _ => 0,
    };
    let rotation = rotation.to_string();
    let mut in_transform = false;
    let mut off_done = false;
    let mut ext_done = false;
    for event in events {
        if matches!(event, Event::End(end) if local_name(end.name().as_ref()) == "xfrm") {
            in_transform = false;
            continue;
        }
        let element = match event {
            Event::Start(element) | Event::Empty(element) => element,
            _ => continue,
        };
        let tag = local_name(element.name().as_ref());
        if tag == "xfrm" {
            replace_attributes(event, &[("rot", rotation.as_str())]);
            in_transform = true;
            continue;
        }
        if !in_transform { continue; }
        if tag == "off" && !off_done {
            replace_attributes(event, &[("x", "0"), ("y", "0")]);
            off_done = true;
        } else if tag == "ext" && !ext_done {
            replace_attributes(event, &[("cx", &cx.to_string()), ("cy", &cy.to_string())]);
            ext_done = true;
        }
    }
    if !off_done || !ext_done {
        return Err(XlsxError::InvalidData(format!("shape {:?} has no complete DrawingML transform to ungroup", drawing.name)));
    }
    Ok(())
}

fn patch_drawing_geometry_xml(
    xml: &str,
    edits: &HashMap<String, crate::ir::Drawing>,
) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::with_capacity(xml.len() + edits.len() * 64));
    let mut hit: HashMap<String, usize> = HashMap::new();
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?.into_owned();
        match &event {
            Event::Start(start) if matches!(local_name(start.name().as_ref()).as_str(), "twoCellAnchor" | "oneCellAnchor" | "absoluteAnchor" | "relSizeAnchor" | "absSizeAnchor") => {
                let mut anchor = vec![event];
                let mut depth = 1usize;
                while depth > 0 {
                    let next = reader.read_event().map_err(XlsxError::Xml)?.into_owned();
                    match &next {
                        Event::Start(_) => depth += 1,
                        Event::End(_) => depth -= 1,
                        Event::Eof => return Err(XlsxError::InvalidData("drawing anchor ended unexpectedly".into())),
                        _ => {}
                    }
                    anchor.push(next);
                }
                let shape_name = anchor.iter().find_map(|event| match event {
                    Event::Start(element) | Event::Empty(element) if local_name(element.name().as_ref()) == "cNvPr" => get_attr(element, "name").filter(|name| edits.contains_key(name)),
                    _ => None,
                });
                if let Some((shape_name, drawing)) = shape_name.as_ref().and_then(|name| edits.get(name).map(|drawing| (name.clone(), drawing))) {
                    *hit.entry(shape_name).or_default() += 1;
                    let patched = rewrite_drawing_anchor_geometry(anchor, drawing)?;
                    for event in patched { writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?; }
                } else {
                    for event in anchor { writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?; }
                }
            }
            Event::Eof => break,
            _ => writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?,
        }
    }
    for name in edits.keys() {
        match hit.get(name).copied().unwrap_or(0) {
            1 => {}
            0 => return Err(XlsxError::InvalidData(format!("drawing shape {name:?} was not found while saving its position"))),
            _ => return Err(XlsxError::InvalidData(format!("drawing shape {name:?} is ambiguous while saving its position"))),
        }
    }
    String::from_utf8(writer.into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn rewrite_drawing_anchor_geometry(
    events: Vec<Event<'static>>,
    drawing: &crate::ir::Drawing,
) -> Result<Vec<Event<'static>>, XlsxError> {
    fn qualified(tag: &str, local: &str) -> String {
        tag.rsplit_once(':').map_or_else(|| local.to_string(), |(prefix, _)| format!("{prefix}:{local}"))
    }
    fn write_marker(tag: &str, anchor: crate::ir::Anchor, output: &mut Vec<Event<'static>>) {
        let marker = tag.to_string();
        output.push(Event::Start(BytesStart::new(marker.clone()).into_owned()));
        for (name, value) in [("col", anchor.col.to_string()), ("colOff", anchor.col_off.to_string()), ("row", anchor.row.to_string()), ("rowOff", anchor.row_off.to_string())] {
            let child = qualified(tag, name);
            output.push(Event::Start(BytesStart::new(child.clone()).into_owned()));
            output.push(Event::Text(BytesText::new(&value).into_owned()));
            output.push(Event::End(BytesEnd::new(child).into_owned()));
        }
        output.push(Event::End(BytesEnd::new(marker).into_owned()));
    }
    fn skip_element(events: &[Event<'static>], start: usize) -> usize {
        let mut depth = 0usize;
        for (index, event) in events.iter().enumerate().skip(start) {
            match event {
                Event::Start(_) => depth += 1,
                Event::End(_) => {
                    depth -= 1;
                    if depth == 0 { return index + 1; }
                }
                _ => {}
            }
        }
        events.len()
    }

    let mut output = Vec::with_capacity(events.len());
    let mut depth = 0usize;
    let mut from_count = 0usize;
    let mut to_count = 0usize;
    let mut extent_count = 0usize;
    let mut index = 0usize;
    while index < events.len() {
        let event = &events[index];
        if let Event::Start(start) = event {
            let tag = local_name(start.name().as_ref());
            if depth == 1 && tag == "from" {
                from_count += 1;
                write_marker(std::str::from_utf8(start.name().as_ref()).unwrap_or("xdr:from"), drawing.from, &mut output);
                index = skip_element(&events, index);
                continue;
            }
            if depth == 1 && tag == "to" {
                if let Some(to) = drawing.to {
                    to_count += 1;
                    write_marker(std::str::from_utf8(start.name().as_ref()).unwrap_or("xdr:to"), to, &mut output);
                    index = skip_element(&events, index);
                    continue;
                }
            }
            if depth == 1 && tag == "ext" && drawing.to.is_none() {
                let (cx, cy) = drawing.extent.ok_or_else(|| XlsxError::InvalidData("one-cell drawing edit has no extent".into()))?;
                let tag_name = std::str::from_utf8(start.name().as_ref()).map_err(|e| XlsxError::InvalidData(e.to_string()))?.to_string();
                let mut extent = BytesStart::new(tag_name);
                let cx = cx.to_string();
                let cy = cy.to_string();
                extent.push_attribute(("cx", cx.as_str()));
                extent.push_attribute(("cy", cy.as_str()));
                output.push(Event::Empty(extent.into_owned()));
                extent_count += 1;
                index = skip_element(&events, index);
                continue;
            }
            depth += 1;
        } else if let Event::Empty(empty) = event {
            if depth == 1 && local_name(empty.name().as_ref()) == "ext" && drawing.to.is_none() {
                let (cx, cy) = drawing.extent.ok_or_else(|| XlsxError::InvalidData("one-cell drawing edit has no extent".into()))?;
                let tag_name = std::str::from_utf8(empty.name().as_ref()).map_err(|e| XlsxError::InvalidData(e.to_string()))?.to_string();
                let mut extent = BytesStart::new(tag_name);
                let cx = cx.to_string();
                let cy = cy.to_string();
                extent.push_attribute(("cx", cx.as_str()));
                extent.push_attribute(("cy", cy.as_str()));
                output.push(Event::Empty(extent.into_owned()));
                extent_count += 1;
                index += 1;
                continue;
            }
        } else if matches!(event, Event::End(_)) {
            depth = depth.saturating_sub(1);
        }
        output.push(event.clone());
        index += 1;
    }
    let expected_to = if drawing.to.is_some() { 1 } else { 0 };
    let expected_extent = if drawing.to.is_none() { 1 } else { 0 };
    if from_count != 1 || to_count != expected_to || extent_count != expected_extent {
        return Err(XlsxError::InvalidData(format!("drawing anchor geometry is incomplete (from={from_count}, to={to_count}, ext={extent_count})")));
    }
    Ok(output)
}

fn write_line_fill<W: Write>(writer: &mut Writer<W>, line: &crate::ir::ShapeLine, prefix: &str) -> Result<(), XlsxError> {
    fn qualified(prefix: &str, local: &str) -> String {
        if prefix.is_empty() { local.into() } else { format!("{prefix}:{local}") }
    }
    fn write_color<W: Write>(writer: &mut Writer<W>, prefix: &str, element: &str, color: &str, theme: Option<usize>, legacy: Option<u8>, brightness: Option<f32>) -> Result<(), XlsxError> {
        let wrapper = BytesStart::new(qualified(prefix, element));
        writer.write_event(Event::Start(wrapper.to_owned())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        let (tag, scheme) = theme.and_then(drawing_theme_token).map_or(("srgbClr", None), |scheme| ("schemeClr", Some(scheme)));
        let mut value = BytesStart::new(qualified(prefix, tag));
        value.push_attribute(("val", scheme.unwrap_or(color)));
        if let Some(index) = legacy {
            value.push_attribute(("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"));
            value.push_attribute(("mc:Ignorable", "a14"));
            value.push_attribute(("xmlns:a14", "http://schemas.microsoft.com/office/drawing/2010/main"));
            let index = index.to_string();
            value.push_attribute(("a14:legacySpreadsheetColorIndex", index.as_str()));
        }
        if let Some(brightness) = brightness.filter(|value| value.abs() > f32::EPSILON) {
            writer.write_event(Event::Start(value)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            let values = if brightness > 0.0 {
                vec![("lumMod", ((1.0 - brightness) * 100_000.0).round() as i32), ("lumOff", (brightness * 100_000.0).round() as i32)]
            } else { vec![("lumMod", ((1.0 + brightness) * 100_000.0).round() as i32)] };
            for (name, value) in values {
                let mut modifier = BytesStart::new(qualified(prefix, name));
                let value = value.to_string();
                modifier.push_attribute(("val", value.as_str()));
                writer.write_event(Event::Empty(modifier)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            }
            writer.write_event(Event::End(BytesEnd::new(qualified(prefix, tag)))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        } else {
            writer.write_event(Event::Empty(value)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        }
        writer.write_event(Event::End(BytesEnd::new(qualified(prefix, element)))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        Ok(())
    }

    if let Some(pattern) = line.pattern.as_deref() {
        let fill_tag = qualified(prefix, "pattFill");
        let mut fill = BytesStart::new(fill_tag.as_str());
        fill.push_attribute(("prst", pattern));
        writer.write_event(Event::Start(fill)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        write_color(writer, prefix, "fgClr", &line.color, line.color_theme, line.legacy_color_index, line.color_brightness)?;
        write_color(writer, prefix, "bgClr", line.back_color.as_deref().unwrap_or("FFFFFF"), line.back_color_theme, line.back_legacy_color_index, line.back_color_brightness)?;
        writer.write_event(Event::End(BytesEnd::new(fill_tag))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
    } else {
        let solid_fill_tag = qualified(prefix, "solidFill");
        writer.write_event(Event::Start(BytesStart::new(solid_fill_tag.as_str()))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        let (tag, scheme) = line.color_theme.and_then(drawing_theme_token).map_or(("srgbClr", None), |scheme| ("schemeClr", Some(scheme)));
        let mut color = BytesStart::new(qualified(prefix, tag));
        color.push_attribute(("val", scheme.unwrap_or(line.color.as_str())));
        if let Some(index) = line.legacy_color_index {
            color.push_attribute(("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"));
            color.push_attribute(("mc:Ignorable", "a14"));
            color.push_attribute(("xmlns:a14", "http://schemas.microsoft.com/office/drawing/2010/main"));
            let index = index.to_string();
            color.push_attribute(("a14:legacySpreadsheetColorIndex", index.as_str()));
        }
        if let Some(brightness) = line.color_brightness.filter(|value| value.abs() > f32::EPSILON) {
            writer.write_event(Event::Start(color)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            let values = if brightness > 0.0 {
                vec![("lumMod", ((1.0 - brightness) * 100_000.0).round() as i32), ("lumOff", (brightness * 100_000.0).round() as i32)]
            } else { vec![("lumMod", ((1.0 + brightness) * 100_000.0).round() as i32)] };
            for (name, value) in values {
                let mut modifier = BytesStart::new(qualified(prefix, name));
                let value = value.to_string();
                modifier.push_attribute(("val", value.as_str()));
                writer.write_event(Event::Empty(modifier)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            }
            writer.write_event(Event::End(BytesEnd::new(qualified(prefix, tag)))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        } else {
            writer.write_event(Event::Empty(color)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        }
        writer.write_event(Event::End(BytesEnd::new(solid_fill_tag))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
    }
    Ok(())
}

fn write_formula(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    formula: &[u16],
    reach: Option<&str>,
) -> Result<(), XlsxError> {
    let formula = oxicells_calc::parser::transform_utf16_formula(formula, oxicells_calc::lexer::name_formula_to_ooxml)
        .unwrap_or_else(|_| formula.to_vec());
    let mut element = BytesStart::new("f");
    if let Some(reach) = reach {
        element.push_attribute(("t", "array"));
        element.push_attribute(("ref", reach));
    }
    writer
        .write_event(Event::Start(element))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    writer
        .write_event(Event::Text(BytesText::new(&crate::string_value::encode(&formula))))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    writer
        .write_event(Event::End(BytesEnd::new("f")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// Write one cell.
///
/// `styles` says which cellXfs index each cell should carry. A cell being
/// written for the first time has to be given its style here: nothing later
/// puts one on it, and without it the cell comes out under General however it
/// was styled  Ewhich is how a date lands in the file as a bare serial number.
fn write_cell_value(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    row: u32,
    col: u32,
    value: &CellEditValue,
    styles: &BTreeMap<(u32, u32), u32>,
) -> Result<(), XlsxError> {
    let reference = format!("{}{row}", col_to_letter(col));
    let style_text = styles.get(&(row, col)).map(|index| index.to_string());
    let mut cell = BytesStart::new("c");
    cell.push_attribute(("r", reference.as_str()));
    if let Some(style) = style_text.as_deref() {
        cell.push_attribute(("s", style));
    }
    if let Some(cell_type) = value.cell_type() {
        cell.push_attribute(("t", cell_type));
    }
    writer
        .write_event(Event::Start(cell))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    if let Some(formula) = value.formula_units() {
        write_formula(writer, &formula, value.array_reach())?;
    }
    if let Some(text) = value.value_text()? {
        writer
            .write_event(Event::Start(BytesStart::new("v")))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        writer
            .write_event(Event::Text(BytesText::new(&text)))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        writer
            .write_event(Event::End(BytesEnd::new("v")))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }
    writer
        .write_event(Event::End(BytesEnd::new("c")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// Every cell a sheet holds, keyed by its one-based row and zero-based column.
fn cells_of(sheet: &crate::ir::Sheet) -> BTreeMap<(u32, u32), &crate::ir::Cell> {
    let mut cells = BTreeMap::new();
    for row in &sheet.rows {
        for cell in &row.cells {
            cells.insert((row.index, cell.col), cell);
        }
    }
    cells
}

/// Merges compare as a set: the order a sheet lists them in means nothing.
fn same_merges(before: &[MergeCell], after: &[MergeCell]) -> bool {
    if before.len() != after.len() {
        return false;
    }
    let key = |merge: &MergeCell| {
        (merge.start_row, merge.start_col, merge.end_row, merge.end_col)
    };
    let mut before: Vec<_> = before.iter().map(key).collect();
    let mut after: Vec<_> = after.iter().map(key).collect();
    before.sort_unstable();
    after.sort_unstable();
    before == after
}

/// The evaluation class of a name in a worksheet's scope. Undefined names
/// can acquire an array definition later, so they also need intersection.
pub fn named_formula_needs_intersection(names: &[(String, String)], sheet: &str, requested: &str) -> bool {
    let (owner, leaf) = crate::ir::defined_name_scope(requested).unwrap_or_else(|| (sheet.to_string(), requested));
    let key = oxicells_calc::case_fold::name_key;
    let definition = names.iter().find(|(held, _)| crate::ir::defined_name_scope(held)
        .is_some_and(|(scope, name)| key(&scope) == key(&owner) && key(name) == key(leaf)))
        .or_else(|| names.iter().find(|(held, _)| !held.contains('!') && key(held) == key(leaf)));
    definition.is_none_or(|(_, formula)| oxicells_calc::lexer::named_formula_needs_intersection(formula))
}

fn cached_named_array_shape(names: &[(String, String)], sheet: &str, requested: &str,
    memo: &std::cell::RefCell<BTreeMap<String, bool>>) -> bool {
    let key = oxicells_calc::case_fold::name_key(requested);
    if let Some(value) = memo.borrow().get(&key) { return *value; }
    let value = named_formula_needs_intersection(names, sheet, requested);
    memo.borrow_mut().insert(key, value);
    value
}

fn dynamic_array_plan(book: &Workbook) -> Vec<Vec<crate::ir::CachedSpill>> {
    book.sheets.iter().map(|sheet| {
        let mut arrays = Vec::new();
        for row in &sheet.rows {
            for cell in row.cells.iter().filter(|cell| cell.spill == crate::ir::Spill::Dynamic) {
                if cell.formula.is_none() { continue; }
                let anchor = (cell.col, row.index.saturating_sub(1));
                let cached = sheet.cached_spills.as_ref().and_then(|blocks| blocks.iter().find(|(at, _)| *at == anchor)).map(|(_, size)| *size);
                let size = cached.unwrap_or_else(|| {
                    if sheet.cached_spills.is_some() { return (1, 1); }
                    let member = |col, index| sheet.rows.iter().find(|r| r.index == index)
                        .and_then(|r| r.cells.iter().find(|c| c.col == col)).is_some_and(|c| c.spill == crate::ir::Spill::Member);
                    let (mut width, mut height) = (1, 1);
                    while cell.col.checked_add(width).is_some_and(|col| col <= oxicells_calc::reference::MAX_COL && member(col, row.index)) { width += 1; }
                    while row.index.checked_add(height).is_some_and(|row| row <= oxicells_calc::reference::MAX_ROW + 1 && member(cell.col, row)) { height += 1; }
                    (width, height)
                });
                arrays.push((anchor, size));
            }
        }
        arrays
    }).collect()
}

const SPREADSHEET_METADATA_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn metadata_tag(e: &BytesStart<'_>, changes: &[(&str, Option<String>)]) -> BytesStart<'static> {
    let mut tag = BytesStart::new(String::from_utf8_lossy(e.name().as_ref()).into_owned());
    for attr in e.attributes().flatten() {
        if !changes.iter().any(|(key, _)| key.as_bytes() == attr.key.as_ref()) { tag.push_attribute(attr); }
    }
    for (key, value) in changes {
        if let Some(value) = value { tag.push_attribute((*key, value.as_str())); }
    }
    tag
}

fn xml_start_text(tag: BytesStart<'_>) -> Result<String, XlsxError> {
    let mut writer = Writer::new(Vec::new());
    writer.write_event(Event::Start(tag)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
    String::from_utf8(writer.into_inner()).map_err(|e| XlsxError::InvalidData(e.to_string()))
}

/// Append one record without rebuilding existing metadata payloads or namespaces.
fn append_metadata_record(xml: &str, section: &str, name: Option<&str>, record: &str) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut target = None::<(usize, usize, BytesStart<'static>)>;
    let mut count = 0usize;
    let record_name = if section == "metadataTypes" { "metadataType" } else { "bk" };
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(e) | Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                if depth == 1 && local == section && name.is_none_or(|wanted| get_attr(&e, "name").as_deref() == Some(wanted)) {
                    if empty {
                        let header = xml_start_text(metadata_tag(&e, &[("count", Some("1".into()))]))?;
                        let closing = format!("</{}>", String::from_utf8_lossy(e.name().as_ref()));
                        return Ok(format!("{}{}{}{}{}", &xml[..start], header, record, closing, &xml[reader.buffer_position() as usize..]));
                    }
                    target = Some((start, reader.buffer_position() as usize, e.clone().into_owned()));
                } else if target.is_some() && depth == 2 && local == record_name { count += 1; }
                if !empty { depth += 1; }
            }
            Event::End(e) => {
                if depth == 2 && local_name(e.name().as_ref()) == section {
                    if let Some((first, content, tag)) = target.take() {
                        let header = xml_start_text(metadata_tag(&tag, &[("count", Some((count + 1).to_string()))]))?;
                        return Ok(format!("{}{}{}{}{}", &xml[..first], header, &xml[content..start], record, &xml[start..]));
                    }
                }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let attribute = name.map(|name| format!(" name=\"{}\"", escape(name))).unwrap_or_default();
    let addition = format!("<{section} xmlns=\"{SPREADSHEET_METADATA_NS}\"{attribute} count=\"1\">{record}</{section}>");
    let rank = |name: &str| match name { "metadataTypes" => 0, "mdxMetadata" => 1, "futureMetadata" => 2, "cellMetadata" => 3, "valueMetadata" => 4, "extLst" => 5, _ => 0 };
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(e) => {
                if depth == 1 && rank(&local_name(e.name().as_ref())) > rank(section) {
                    return Ok(format!("{}{}{}", &xml[..start], addition, &xml[start..]));
                }
                depth += 1;
            }
            Event::Empty(e) => {
                if depth == 0 && local_name(e.name().as_ref()) == "metadata" {
                    let closing = format!("</{}>", String::from_utf8_lossy(e.name().as_ref()));
                    return Ok(format!("{}{}{}{}{}", &xml[..start], xml_start_text(e.clone().into_owned())?, addition,
                        closing, &xml[reader.buffer_position() as usize..]));
                }
                if depth == 1 && rank(&local_name(e.name().as_ref())) > rank(section) {
                    return Ok(format!("{}{}{}", &xml[..start], addition, &xml[start..]));
                }
            }
            Event::End(_) => {
                if depth == 1 { return Ok(format!("{}{}{}", &xml[..start], addition, &xml[start..])); }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => return Err(XlsxError::InvalidData("metadata part has no root element".into())),
            _ => {}
        }
    }
}

/// Remove only this metadata type from a cell block; keep all other records
/// and extension payloads so mixed metadata survives formula replacement.
fn without_metadata_type(block: &str, type_index: usize) -> Result<(String, bool), XlsxError> {
    let mut reader = Reader::from_str(block);
    let mut writer = Writer::new(Vec::new());
    let mut depth = 0usize;
    let mut skip = 0usize;
    let mut retained = false;
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        if skip > 0 {
            match event { Event::Start(_) => skip += 1, Event::End(_) => skip -= 1, Event::Eof => break, _ => {} }
            continue;
        }
        match &event {
            Event::Start(e) | Event::Empty(e) => {
                let empty = matches!(&event, Event::Empty(_));
                if skip > 0 { if !empty { skip += 1; } continue; }
                if depth == 1 && local_name(e.name().as_ref()) == "rc"
                    && get_attr(e, "t").and_then(|s| s.parse::<usize>().ok()) == Some(type_index) {
                    if !empty { skip = 1; }
                    continue;
                }
                if depth == 1 { retained = true; }
                if depth == 0 && e.attributes().flatten().any(|a| !a.key.as_ref().starts_with(b"xmlns")) { retained = true; }
                if !empty { depth += 1; }
            }
            Event::End(_) => {
                if skip > 0 { skip -= 1; continue; }
                depth = depth.saturating_sub(1);
            }
            Event::Text(text) if !text.as_ref().iter().all(u8::is_ascii_whitespace) => retained = true,
            Event::Eof => break,
            _ if skip > 0 => continue,
            _ => {}
        }
        writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
    }
    Ok((String::from_utf8(writer.into_inner()).map_err(|e| XlsxError::InvalidData(e.to_string()))?, retained))
}

fn append_cell_metadata_record(block: &str, record: &str) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(block);
    let mut depth = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(e) => {
                if depth == 1 && local_name(e.name().as_ref()) == "extLst" {
                    return Ok(format!("{}{}{}", &block[..start], record, &block[start..]));
                }
                depth += 1;
            }
            Event::Empty(e) => {
                if depth == 0 {
                    let closing = format!("</{}>", String::from_utf8_lossy(e.name().as_ref()));
                    return Ok(format!("{}{}{}{}{}", &block[..start], xml_start_text(e.into_owned())?, record, closing, &block[reader.buffer_position() as usize..]));
                }
                if depth == 1 && local_name(e.name().as_ref()) == "extLst" {
                    return Ok(format!("{}{}{}", &block[..start], record, &block[start..]));
                }
            }
            Event::End(_) => {
                if depth == 1 { return Ok(format!("{}{}{}", &block[..start], record, &block[start..])); }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => return Err(XlsxError::InvalidData("cell metadata block has no root".into())),
            _ => {}
        }
    }
}

type DynamicMetadataChanges = BTreeMap<(bool, Option<usize>), Option<usize>>;

fn dynamic_metadata_changes(xml: &str, requested: &std::collections::BTreeSet<(bool, Option<usize>)>)
    -> Result<(String, DynamicMetadataChanges), XlsxError>
{
    let info = crate::parser::dynamic_metadata(xml)?;
    let mut output = xml.to_string();
    let mut changes = BTreeMap::new();
    let adding = requested.iter().any(|(dynamic, id)| *dynamic && !id.is_some_and(|id| info.cells.contains(&id)));
    let type_index = info.type_index.unwrap_or(info.type_count + 1);
    if adding {
        if info.type_index.is_none() {
            let record = format!("<metadataType xmlns=\"{SPREADSHEET_METADATA_NS}\" name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/>");
            output = append_metadata_record(&output, "metadataTypes", None, &record)?;
        }
        let record = format!("<bk xmlns=\"{SPREADSHEET_METADATA_NS}\"><extLst><ext uri=\"{{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}}\"><xda:dynamicArrayProperties xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\" fDynamic=\"1\" fCollapsed=\"0\"/></ext></extLst></bk>");
        output = append_metadata_record(&output, "futureMetadata", Some("XLDAPR"), &record)?;
    }
    let mut count = info.cell_count;
    for &(dynamic, id) in requested {
        if dynamic && id.is_some_and(|id| info.cells.contains(&id)) { changes.insert((dynamic, id), id); continue; }
        let base = match id {
            Some(index) => info.cell_blocks.get(index.saturating_sub(1)).ok_or_else(|| XlsxError::InvalidData("cell metadata index is out of range".into()))?.clone(),
            None => format!("<bk xmlns=\"{SPREADSHEET_METADATA_NS}\"></bk>"),
        };
        let (mut block, retained) = without_metadata_type(&base, type_index)?;
        if !dynamic && !retained { changes.insert((dynamic, id), None); continue; }
        if dynamic {
            let record = format!("<rc xmlns=\"{SPREADSHEET_METADATA_NS}\" t=\"{type_index}\" v=\"{}\"/>", info.future_count);
            block = append_cell_metadata_record(&block, &record)?;
        }
        output = append_metadata_record(&output, "cellMetadata", None, &block)?;
        count += 1;
        changes.insert((dynamic, id), Some(count));
    }
    Ok((output, changes))
}

fn cell_metadata_index(tag: &BytesStart<'_>) -> Option<usize> {
    get_attr(tag, "cm").and_then(|s| s.parse().ok()).filter(|id| *id > 0)
}

fn dynamic_range(anchor: (u32, u32), size: (u32, u32)) -> Result<String, XlsxError> {
    let (col, row) = anchor;
    let (width, height) = size;
    let right = width.checked_sub(1).and_then(|n| col.checked_add(n)).filter(|c| *c <= oxicells_calc::reference::MAX_COL);
    let bottom = height.checked_sub(1).and_then(|n| row.checked_add(n)).filter(|r| *r <= oxicells_calc::reference::MAX_ROW);
    let (Some(right), Some(bottom)) = (right, bottom) else { return Err(XlsxError::InvalidData("dynamic array extent is outside the worksheet".into())); };
    let first = format!("{}{}", col_to_letter(col), row + 1);
    Ok(if (col, row) == (right, bottom) { first } else { format!("{first}:{}{}", col_to_letter(right), bottom + 1) })
}

fn worksheet_metadata_requests(xml: &str, arrays: &BTreeMap<(u32, u32), (u32, u32)>, dynamic_ids: &std::collections::BTreeSet<usize>)
    -> Result<std::collections::BTreeSet<(bool, Option<usize>)>, XlsxError>
{
    let mut reader = Reader::from_str(xml);
    let mut path = Vec::<String>::new();
    let mut requested = std::collections::BTreeSet::new();
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        match &event {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                if name == "c" && path.last().map(String::as_str) == Some("row")
                    && path.iter().rev().nth(1).map(String::as_str) == Some("sheetData") {
                    let dynamic = get_attr(e, "r").and_then(|s| oxicells_calc::reference::parse_a1(&s))
                        .is_some_and(|r| arrays.contains_key(&(r.col, r.row)));
                    let id = cell_metadata_index(e);
                    if dynamic || id.is_some_and(|id| dynamic_ids.contains(&id)) { requested.insert((dynamic, id)); }
                }
                if matches!(&event, Event::Start(_)) { path.push(name); }
            }
            Event::End(_) => { path.pop(); }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(requested)
}

fn patch_dynamic_worksheet(xml: &str, arrays: &BTreeMap<(u32, u32), (u32, u32)>, dynamic_ids: &std::collections::BTreeSet<usize>, changes: &DynamicMetadataChanges)
    -> Result<String, XlsxError>
{
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Vec::new());
    let mut path = Vec::<String>::new();
    let mut active = None;
    let mut seen = std::collections::BTreeSet::new();
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref());
                let mut tag = e.clone().into_owned();
                if name == "c" && path.last().map(String::as_str) == Some("row")
                    && path.iter().rev().nth(1).map(String::as_str) == Some("sheetData") {
                    active = get_attr(&e, "r").and_then(|s| oxicells_calc::reference::parse_a1(&s)).map(|r| (r.col, r.row));
                    let dynamic = active.is_some_and(|at| arrays.contains_key(&at));
                    let id = cell_metadata_index(&e);
                    if dynamic || id.is_some_and(|id| dynamic_ids.contains(&id)) {
                        let replacement = changes.get(&(dynamic, id)).ok_or_else(|| XlsxError::InvalidData("missing dynamic metadata mapping".into()))?;
                        tag = metadata_tag(&e, &[("cm", replacement.map(|n| n.to_string()))]);
                    }
                } else if name == "f" && path.last().map(String::as_str) == Some("c") {
                    if let Some(at) = active {
                        if let Some(size) = arrays.get(&at) {
                            if !seen.insert(at) { return Err(XlsxError::InvalidData("multiple formulas for a dynamic anchor".into())); }
                            tag = metadata_tag(&e, &[("t", Some("array".into())), ("ref", Some(dynamic_range(at, *size)?)), ("si", None)]);
                        }
                    }
                }
                writer.write_event(if empty { Event::Empty(tag) } else { Event::Start(tag) })
                    .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                if !empty { path.push(name); } else if name == "c" { active = None; }
            }
            Event::End(e) => {
                if local_name(e.name().as_ref()) == "c" && path.iter().rev().nth(1).map(String::as_str) == Some("row") { active = None; }
                path.pop();
                writer.write_event(Event::End(e)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            }
            Event::Eof => break,
            other => writer.write_event(other).map_err(|e| XlsxError::InvalidData(e.to_string()))?,
        }
    }
    if seen.len() != arrays.len() { return Err(XlsxError::InvalidData("a dynamic array anchor has no saved formula".into())); }
    String::from_utf8(writer.into_inner()).map_err(|e| XlsxError::InvalidData(e.to_string()))
}

fn patch_dynamic_array_package(data: &[u8], plan: &[Vec<crate::ir::CachedSpill>]) -> Result<Vec<u8>, XlsxError> {
    let mut source = OoxmlArchive::new(data)?;
    let mut rels = source.read_part("xl/_rels/workbook.xml.rels")?;
    let relationships = parse_relationships(&rels)?;
    let metadata_rel = relationships.values().find(|r| r.rel_type.ends_with("/sheetMetadata") && r.target_mode.as_deref() != Some("External"));
    let metadata_path = match metadata_rel {
        Some(rel) => oxidocs_common::security::sanitize_rel_target("xl", &rel.target)
            .map_err(|e| XlsxError::InvalidData(e.to_string()))?,
        None => {
            let mut candidate = "xl/metadata.xml".to_string();
            let mut index = 1;
            while source.try_read_part(&candidate)?.is_some() { candidate = format!("xl/metadata{index}.xml"); index += 1; }
            candidate
        }
    };
    let original = source.try_read_part(&metadata_path)?.unwrap_or_else(|| format!("<metadata xmlns=\"{SPREADSHEET_METADATA_NS}\"></metadata>"));
    let info = crate::parser::dynamic_metadata(&original)?;
    if plan.iter().all(Vec::is_empty) && info.cells.is_empty() { return Ok(data.to_vec()); }
    let paths = XlsxEditor::sheet_paths_in(data)?;
    if paths.len() != plan.len() { return Err(XlsxError::InvalidData("dynamic array plan does not match the sheet order".into())); }
    let mut sheets = Vec::new();
    let mut requested = std::collections::BTreeSet::new();
    for (path, arrays) in paths.iter().zip(plan) {
        let arrays: BTreeMap<_, _> = arrays.iter().copied().collect();
        for (at, size) in &arrays { dynamic_range(*at, *size)?; }
        let xml = source.read_part(path)?;
        let requests = worksheet_metadata_requests(&xml, &arrays, &info.cells)?;
        if requests.is_empty() && arrays.is_empty() { continue; }
        requested.extend(requests);
        sheets.push((path.clone(), xml, arrays));
    }
    let (metadata, changes) = dynamic_metadata_changes(&original, &requested)?;
    let mut replacements = BTreeMap::<String, Vec<u8>>::new();
    for (path, xml, arrays) in sheets {
        replacements.insert(path, patch_dynamic_worksheet(&xml, &arrays, &info.cells, &changes)?.into_bytes());
    }
    if metadata != original || metadata_rel.is_none() && plan.iter().any(|v| !v.is_empty()) {
        replacements.insert(metadata_path.clone(), metadata.into_bytes());
    }
    if metadata_rel.is_none() && plan.iter().any(|v| !v.is_empty()) {
        let target = metadata_path.strip_prefix("xl/").ok_or_else(|| XlsxError::InvalidData("metadata part is outside the workbook directory".into()))?;
        rels = append_relationship(Some(&rels), "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata", target)?;
        replacements.insert("xl/_rels/workbook.xml.rels".into(), rels.into_bytes());
        let types = append_content_type(&source.read_part("[Content_Types].xml")?, &metadata_path,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml")?;
        replacements.insert("[Content_Types].xml".into(), types.into_bytes());
    }
    replace_package_parts(data, replacements)
}

fn replace_package_parts(data: &[u8], mut replacements: BTreeMap<String, Vec<u8>>) -> Result<Vec<u8>, XlsxError> {
    let mut archive = ZipArchive::new(Cursor::new(data)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    for i in 0..archive.len() {
        let entry = archive.by_index(i).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        let path = entry.name().to_string();
        if let Some(bytes) = replacements.remove(&path) {
            writer.start_file(path, SimpleFileOptions::default().compression_method(entry.compression())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            writer.write_all(&bytes).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        } else {
            writer.raw_copy_file(entry).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        }
    }
    for (path, bytes) in replacements {
        writer.start_file(path, SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        writer.write_all(&bytes).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
    }
    Ok(writer.finish().map_err(|e| XlsxError::InvalidData(e.to_string()))?.into_inner())
}

const RICH_DATA_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";

fn xml_root_attributes(xml: &str, changes: &[(&str, Option<String>)]) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(xml);
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                let mut writer = Writer::new(Vec::new());
                let tag = metadata_tag(&tag, changes);
                writer.write_event(if empty { Event::Empty(tag) } else { Event::Start(tag) }).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                let header = String::from_utf8(writer.into_inner()).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                return Ok(format!("{}{}{}", &xml[..start], header, &xml[reader.buffer_position() as usize..]));
            }
            Event::Eof => return Err(XlsxError::InvalidData("XML part has no root element".into())),
            _ => {}
        }
    }
}

/// Rich-data tables count only their own records. Preserve all existing
/// records, their order, namespaces and extension payloads while appending.
fn append_rich_record(xml: &str, child: &str, record: &str) -> Result<String, XlsxError> {
    let count = crate::parser::metadata_elements(xml)?.iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == child).count();
    let xml = xml_root_attributes(xml, &[("count", Some((count + 1).to_string()))])?;
    let (extension, end, empty, root) = direct_xml_child(&xml, "extLst")?;
    let at = extension.map_or(end, |(start, _)| start);
    if !empty { return Ok(format!("{}{}{}", &xml[..at], record, &xml[at..])); }
    let mut reader = Reader::from_str(&xml);
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Empty(tag) => return Ok(format!("{}{}{}</{}>{}", &xml[..start], xml_start_text(tag.into_owned())?, record, root, &xml[end..])),
            Event::Eof => return Err(XlsxError::InvalidData("rich-data table has no root element".into())),
            _ => {}
        }
    }
}

struct RichErrorCell<'a> {
    xml: &'a str,
    metadata: Option<usize>,
    kind: Option<u32>,
}

fn rich_error_cells<'a>(xml: &'a str, clears: &std::collections::BTreeSet<(u32, u32)>, known: &BTreeMap<usize, &'static str>) -> Result<Vec<RichErrorCell<'a>>, XlsxError> {
    if !["#SPILL!", "#CONNECT!", "#BLOCKED!", "#UNKNOWN!", "#FIELD!", "#CALC!"].iter().any(|error| xml.contains(error))
        && (clears.is_empty() || known.is_empty() || !xml.contains("vm=")) { return Ok(Vec::new()); }
    let mut result = Vec::new();
    for (_, data) in crate::parser::metadata_elements(xml)?.into_iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == "sheetData") {
        for (_, row) in crate::parser::metadata_elements(data)?.into_iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == "row") {
            for (tag, cell) in crate::parser::metadata_elements(row)?.into_iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == "c") {
                let metadata = get_attr(&tag, "vm").and_then(|s| s.parse().ok()).filter(|index| *index > 0);
                let clearing = metadata.is_some_and(|index| known.contains_key(&index)) && get_attr(&tag, "r")
                    .and_then(|s| oxicells_calc::reference::parse_a1(&s)).is_some_and(|at| clears.contains(&(at.row + 1, at.col)));
                let mut kind = None;
                if get_attr(&tag, "t").as_deref() == Some("e") {
                    if let Some((_, value)) = crate::parser::metadata_elements(cell)?.into_iter().find(|(tag, _)| local_name(tag.name().as_ref()) == "v") {
                        let mut reader = Reader::from_str(value);
                        if let Event::Start(start) = reader.read_event().map_err(XlsxError::Xml)? {
                            let text = reader.read_text(start.name()).map_err(XlsxError::Xml)?;
                            kind = crate::parser::rich_error_type(text.trim());
                        }
                    }
                }
                if kind.is_some() || clearing { result.push(RichErrorCell { xml: cell, metadata, kind }); }
            }
        }
    }
    Ok(result)
}

fn generic_error_structure(xml: &mut String, subtype: bool) -> Result<usize, XlsxError> {
    let structures = crate::parser::metadata_elements(xml)?;
    let mut count = 0;
    for (tag, record) in structures.iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == "s") {
        let fields = crate::parser::metadata_elements(record)?;
        let keys: Vec<_> = fields.iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == "k").map(|(tag, _)| (get_attr(tag, "n"), get_attr(tag, "t"))).collect();
        let expected = if subtype { vec![(Some("errorType".into()), Some("i".into())), (Some("subType".into()), Some("i".into()))] }
            else { vec![(Some("errorType".into()), Some("i".into()))] };
        if get_attr(tag, "t").as_deref() == Some("_error") && keys == expected { return Ok(count); }
        count += 1;
    }
    let second = if subtype { "<k n=\"subType\" t=\"i\"/>" } else { "" };
    let record = format!("<s xmlns=\"{RICH_DATA_NS}\" t=\"_error\"><k n=\"errorType\" t=\"i\"/>{second}</s>");
    *xml = append_rich_record(xml, "s", &record)?;
    Ok(count)
}

/// New errors use Excel's generic error payload; existing compatible value
/// metadata is reused verbatim, including subtype and spill-offset fields.
fn rich_error_records(metadata: &str, structures: &mut String, values: &mut String,
    requests: &std::collections::BTreeSet<(u32, Option<usize>)>) -> Result<(String, BTreeMap<(u32, Option<usize>), usize>), XlsxError> {
    let info = crate::parser::rich_error_metadata(metadata, structures, values)?;
    let type_index = info.type_index.unwrap_or(info.type_count + 1);
    let mut output = metadata.to_string();
    let mut future_count = info.future_count;
    let mut count = info.blocks.len();
    let mut added = BTreeMap::new();
    let mut ids = BTreeMap::new();
    for &(kind, old) in requests {
        if old.and_then(|index| info.errors.get(&index)).is_some_and(|error| crate::parser::rich_error_type(error) == Some(kind)) {
            ids.insert((kind, old), old.unwrap());
            continue;
        }
        if info.type_index.is_none() && added.is_empty() {
            let record = format!("<metadataType xmlns=\"{SPREADSHEET_METADATA_NS}\" name=\"XLRICHVALUE\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\"/>");
            output = append_metadata_record(&output, "metadataTypes", None, &record)?;
        }
        let future = match added.get(&kind) {
            Some(index) => *index,
            None => {
                let structure = generic_error_structure(structures, kind != 11)?;
                let index = crate::parser::metadata_elements(values)?.iter().filter(|(tag, _)| local_name(tag.name().as_ref()) == "rv").count();
                let subtype = if kind == 11 { String::new() } else { format!("<v>{}</v>", if kind == 10 { -1 } else { 0 }) };
                *values = append_rich_record(values, "rv", &format!("<rv xmlns=\"{RICH_DATA_NS}\" s=\"{structure}\"><v>{kind}</v>{subtype}</rv>"))?;
                let record = format!("<bk xmlns=\"{SPREADSHEET_METADATA_NS}\"><extLst><ext uri=\"{{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}}\"><xlrd:rvb xmlns:xlrd=\"{RICH_DATA_NS}\" i=\"{index}\"/></ext></extLst></bk>");
                output = append_metadata_record(&output, "futureMetadata", Some("XLRICHVALUE"), &record)?;
                added.insert(kind, future_count);
                future_count += 1;
                future_count - 1
            }
        };
        let base = match old {
            Some(index) => info.blocks.get(index - 1).ok_or_else(|| XlsxError::InvalidData("value metadata index is out of range".into()))?.clone(),
            None => format!("<bk xmlns=\"{SPREADSHEET_METADATA_NS}\"/>"),
        };
        let (base, _) = without_metadata_type(&base, type_index)?;
        let block = append_cell_metadata_record(&base, &format!("<rc xmlns=\"{SPREADSHEET_METADATA_NS}\" t=\"{type_index}\" v=\"{future}\"/>"))?;
        output = append_metadata_record(&output, "valueMetadata", None, &block)?;
        count += 1;
        ids.insert((kind, old), count);
    }
    Ok((output, ids))
}

fn clear_rich_error_records(metadata: &str, structures: &str, values: &str, requests: &std::collections::BTreeSet<usize>)
    -> Result<(String, BTreeMap<usize, Option<usize>>), XlsxError> {
    let info = crate::parser::rich_error_metadata(metadata, structures, values)?;
    let mut output = metadata.to_string();
    let mut ids = BTreeMap::new();
    let mut count = info.blocks.len();
    for index in requests {
        let block = info.blocks.get(index - 1).ok_or_else(|| XlsxError::InvalidData("value metadata index is out of range".into()))?;
        let (block, retained) = without_metadata_type(block, info.type_index.ok_or_else(|| XlsxError::InvalidData("rich error has no metadata type".into()))?)?;
        if retained {
            output = append_metadata_record(&output, "valueMetadata", None, &block)?;
            count += 1;
            ids.insert(*index, Some(count));
        } else { ids.insert(*index, None); }
    }
    Ok((output, ids))
}

fn patch_rich_error_workbook(xml: &str) -> Result<String, XlsxError> {
    const FEATURES: &str = "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures";
    const EXTENSION: &str = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}";
    // Rich error caches need the calculation capability and version markers,
    // including when a reader leaves a workbook in manual calculation mode.
    fn root(xml: &str) -> Result<(BytesStart<'static>, bool, usize, usize), XlsxError> {
        let mut reader = Reader::from_str(xml);
        loop {
            let start = reader.buffer_position() as usize;
            let event = reader.read_event().map_err(XlsxError::Xml)?;
            let empty = matches!(event, Event::Empty(_));
            match event {
                Event::Start(tag) | Event::Empty(tag) => return Ok((tag.into_owned(), empty, start, reader.buffer_position() as usize)),
                Event::Eof => return Err(XlsxError::InvalidData("workbook element is missing".into())),
                _ => {}
            }
        }
    }
    fn append(xml: &str, child: &str) -> Result<String, XlsxError> {
        let (tag, empty, start, end) = root(xml)?;
        if empty {
            let name = String::from_utf8_lossy(tag.name().as_ref()).into_owned();
            return Ok(format!("{}{}{child}</{name}>{}", &xml[..start], xml_start_text(tag)?, &xml[end..]));
        }
        append_xml_child(xml, &local_name(tag.name().as_ref()), child)
    }
    let version_span = direct_xml_child(xml, "fileVersion")?.0;
    let version = version_span.map(|(start, end)| &xml[start..end]).unwrap_or(
        r#"<fileVersion xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#);
    let (tag, _, _, _) = root(version)?;
    let mut changes = Vec::new();
    if get_attr(&tag, "appName").as_deref() != Some("xl") { changes.push(("appName", Some("xl".to_string()))); }
    for name in ["lastEdited", "lowestEdited"] {
        if get_attr(&tag, name).is_none() { changes.push((name, Some("7".to_string()))); }
    }
    let version = if changes.is_empty() { version.to_string() } else { xml_root_attributes(version, &changes)? };
    let mut output = if version_span.is_some() {
        replace_direct_xml_child(xml, "fileVersion", Some(&version))?
    } else {
        let (_, empty, _, end) = root(xml)?;
        if empty { append(xml, &version)? } else { format!("{}{}{}", &xml[..end], version, &xml[end..]) }
    };
    let calc_span = direct_xml_child(&output, "calcPr")?.0;
    let calc = calc_span.map(|(start, end)| &output[start..end]).unwrap_or(
        r#"<calcPr xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#);
    let (tag, _, _, _) = root(calc)?;
    let calc = if get_attr(&tag, "calcId").and_then(|id| id.parse::<u32>().ok()).is_some_and(|id| id >= 191029) {
        calc.to_string()
    } else { xml_root_attributes(calc, &[("calcId", Some("191029".into()))])? };
    output = set_worksheet_drawing_child(&output, "calcPr", &calc,
        &["oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"])?;

    let mut reader = quick_xml::NsReader::from_str(&output);
    let mut depth = 0usize;
    let mut extension_start = None;
    let mut feature_start = None;
    let mut extension_span = None;
    let mut feature_span = None;
    loop {
        let start = reader.buffer_position() as usize;
        let (namespace, event) = reader.read_resolved_event().map_err(XlsxError::Xml)?;
        let spreadsheet = matches!(&namespace, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == SPREADSHEET_METADATA_NS.as_bytes());
        let feature_namespace = matches!(&namespace, quick_xml::name::ResolveResult::Bound(ns) if ns.as_ref() == FEATURES.as_bytes());
        let end = reader.buffer_position() as usize;
        let empty = matches!(event, Event::Empty(_));
        match event {
            Event::Start(tag) | Event::Empty(tag) => {
                let local = local_name(tag.name().as_ref());
                if depth == 2 && spreadsheet && local == "ext" && get_attr(&tag, "uri").as_deref() == Some(EXTENSION) {
                    if empty { extension_span = Some((start, end)); } else { extension_start = Some(start); }
                }
                if depth == 3 && extension_start.is_some() && feature_namespace && local == "calcFeatures" {
                    if empty { feature_span = Some((start, end)); } else { feature_start = Some(start); }
                }
                if depth == 4 && feature_start.is_some() && feature_namespace && local == "feature"
                    && get_attr(&tag, "name").as_deref() == Some("microsoft.com:RD") { return Ok(output); }
                if !empty { depth += 1; }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 3 { if let Some(start) = feature_start.take() { feature_span = Some((start, end)); } }
                if depth == 2 { if let Some(start) = extension_start.take() { extension_span = Some((start, end)); } }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let feature = format!(r#"<xcalcf:feature xmlns:xcalcf="{FEATURES}" name="microsoft.com:RD"/>"#);
    if let Some((start, end)) = feature_span {
        let changed = append(&output[start..end], &feature)?;
        return Ok(format!("{}{}{}", &output[..start], changed, &output[end..]));
    }
    let features = format!(r#"<xcalcf:calcFeatures xmlns:xcalcf="{FEATURES}">{feature}</xcalcf:calcFeatures>"#);
    if let Some((start, end)) = extension_span {
        let changed = append(&output[start..end], &features)?;
        return Ok(format!("{}{}{}", &output[..start], changed, &output[end..]));
    }
    let extension = format!(r#"<ext xmlns="{SPREADSHEET_METADATA_NS}" uri="{EXTENSION}">{features}</ext>"#);
    let extensions = match direct_xml_child(&output, "extLst")?.0 {
        Some((start, end)) => append(&output[start..end], &extension)?,
        None => format!(r#"<extLst xmlns="{SPREADSHEET_METADATA_NS}">{extension}</extLst>"#),
    };
    replace_direct_xml_child(&output, "extLst", Some(&extensions))
}

fn patch_rich_error_package(data: &[u8], clear_plan: &[std::collections::BTreeSet<(u32, u32)>]) -> Result<Vec<u8>, XlsxError> {
    let mut source = OoxmlArchive::new(data)?;
    let mut rels = source.read_part("xl/_rels/workbook.xml.rels")?;
    let relationships = parse_relationships(&rels)?;
    let mut types = source.read_part("[Content_Types].xml")?;
    let mut type_info = String::from("<rvTypesInfo xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2\"><global><keyFlags>");
    for key in ["_Self", "_DisplayString", "_Flags", "_Format", "_SubLabel", "_Attribution", "_Icon", "_Display", "_CanonicalPropertyNames", "_ClassificationId"] {
        type_info.push_str(&format!("<key name=\"{key}\">{}<flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>",
            if key == "_Self" { "<flag name=\"ExcludeFromFile\" value=\"1\"/>" } else { "" }));
    }
    type_info.push_str("</keyFlags></global></rvTypesInfo>");
    let specifications = [
        ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata", "xl/metadata.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml", format!("<metadata xmlns=\"{SPREADSHEET_METADATA_NS}\"/>")),
        ("http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure", "xl/richData/rdrichvaluestructure.xml", "application/vnd.ms-excel.rdrichvaluestructure+xml", format!("<rvStructures xmlns=\"{RICH_DATA_NS}\" count=\"0\"/>")),
        ("http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue", "xl/richData/rdrichvalue.xml", "application/vnd.ms-excel.rdrichvalue+xml", format!("<rvData xmlns=\"{RICH_DATA_NS}\" count=\"0\"/>")),
        ("http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes", "xl/richData/rdRichValueTypes.xml", "application/vnd.ms-excel.rdrichvaluetypes+xml", type_info),
    ];
    let mut parts = Vec::new();
    for (kind, default, content_type, empty) in specifications {
        let relation = relationships.values().find(|rel| rel.rel_type == kind && rel.target_mode.as_deref() != Some("External"));
        let path = match relation {
            Some(rel) => oxidocs_common::security::sanitize_rel_target("xl", &rel.target).map_err(|e| XlsxError::InvalidData(e.to_string()))?,
            None => {
                let mut path = default.to_string();
                let mut index = 1;
                while source.try_read_part(&path)?.is_some() { path = format!("{}{index}.xml", default.trim_end_matches(".xml")); index += 1; }
                path
            }
        };
        let xml = source.try_read_part(&path)?.unwrap_or(empty);
        if relation.is_none() {
            let target = path.strip_prefix("xl/").ok_or_else(|| XlsxError::InvalidData("rich-data part is outside the workbook directory".into()))?;
            rels = append_relationship(Some(&rels), kind, target)?;
        }
        types = append_content_type(&types, &path, content_type)?;
        parts.push((path, xml));
    }
    let original = crate::parser::rich_error_metadata(&parts[0].1, &parts[1].1, &parts[2].1)?;
    let paths = XlsxEditor::sheet_paths_in(data)?;
    if paths.len() != clear_plan.len() { return Err(XlsxError::InvalidData("value metadata plan does not match the sheet order".into())); }
    let mut sheets = Vec::new();
    let mut requests = std::collections::BTreeSet::new();
    let mut clearing = std::collections::BTreeSet::new();
    for (path, clears) in paths.into_iter().zip(clear_plan) {
        let xml = source.read_part(&path)?;
        let cells = rich_error_cells(&xml, clears, &original.errors)?;
        if cells.is_empty() { continue; }
        for cell in cells {
            if let Some(kind) = cell.kind { requests.insert((kind, cell.metadata)); }
            else if let Some(index) = cell.metadata { clearing.insert(index); }
        }
        sheets.push((path, xml, clears));
    }
    if sheets.is_empty() { return Ok(data.to_vec()); }
    let mut structures = parts[1].1.clone();
    let mut values = parts[2].1.clone();
    let (metadata, ids) = rich_error_records(&parts[0].1, &mut structures, &mut values, &requests)?;
    let (metadata, clear_ids) = clear_rich_error_records(&metadata, &structures, &values, &clearing)?;
    parts[0].1 = metadata;
    parts[1].1 = structures;
    parts[2].1 = values;
    let mut replacements = BTreeMap::new();
    for (path, xml, clears) in sheets {
        let mut edits = Vec::new();
        for cell in rich_error_cells(&xml, clears, &original.errors)? {
            let (body, id) = if let Some(kind) = cell.kind {
                (replace_direct_xml_child(cell.xml, "v", Some(&format!("<v xmlns=\"{SPREADSHEET_METADATA_NS}\">#VALUE!</v>")))?, Some(ids[&(kind, cell.metadata)]))
            } else {
                (cell.xml.to_string(), clear_ids[&cell.metadata.unwrap()])
            };
            let patched = xml_root_attributes(&body, &[("vm", id.map(|id| id.to_string()))])?;
            let start = cell.xml.as_ptr() as usize - xml.as_ptr() as usize;
            edits.push((start, start + cell.xml.len(), patched));
        }
        let mut result = String::new();
        let mut from = 0;
        for (start, end, body) in edits { result.push_str(&xml[from..start]); result.push_str(&body); from = end; }
        result.push_str(&xml[from..]);
        replacements.insert(path, result.into_bytes());
    }
    for (path, xml) in parts { replacements.insert(path, xml.into_bytes()); }
    if !requests.is_empty() {
        let workbook = source.read_part("xl/workbook.xml")?;
        replacements.insert("xl/workbook.xml".into(), patch_rich_error_workbook(&workbook)?.into_bytes());
    }
    replacements.insert("xl/_rels/workbook.xml.rels".into(), rels.into_bytes());
    replacements.insert("[Content_Types].xml".into(), types.into_bytes());
    replace_package_parts(data, replacements)
}

fn same_content(before: &crate::ir::Cell, after: &crate::ir::Cell) -> bool {
    before.formula == after.formula
        && before.formula_utf16 == after.formula_utf16
        && before.array_block == after.array_block
        && before.spill == after.spill
        && same_value(&before.value, &after.value)
        && before.runs == after.runs
}

fn same_value(before: &crate::ir::CellValue, after: &crate::ir::CellValue) -> bool {
    use crate::ir::CellValue::*;
    match (before, after) {
        (Empty, Empty) => true,
        (String(before), String(after)) => before == after,
        (Utf16(before), Utf16(after)) => before == after,
        (Number(before), Number(after)) => before == after,
        (Boolean(before), Boolean(after)) => before == after,
        (Error(before), Error(after)) => before == after,
        _ => false,
    }
}

/// How a cell should be written back, or `None` for one this editor cannot
/// express  Ean error value has no edit of its own.
///
/// A member of an array formula is written the way the file keeps it: the
/// formula on the block's top-left cell, with the block beside it, and only
/// a value on every other cell.
fn edit_for(cell: &crate::ir::Cell, row: u32, multiple: &dyn Fn(&str) -> bool) -> Option<CellEditValue> {
    let legacy = cell.array_block.is_none();
    let saved_formula = |formula: &str| {
        if !legacy || !formula.contains('@') { return formula.to_string(); }
        // Ordinary OOXML formulas already imply scalar intersections. Drop
        // only operators that legacy evaluation can reproduce; an explicit
        // intersection inside SUM, for example, changes the result.
        oxicells_calc::lexer::legacy_formula(formula, multiple)
    };
    if let Some(units) = cell.formula_units() {
        let reach = cell.array_block.and_then(|(top, left, bottom, right)|
            (row == top && cell.col == left).then(|| format!("{}{top}:{}{bottom}", col_to_letter(left), col_to_letter(right))));
        if cell.array_block.is_none() || reach.is_some() {
            let formula = if legacy && units.contains(&(b'@' as u16)) {
                oxicells_calc::parser::transform_utf16_formula(units, |formula| Ok(saved_formula(formula)))
                    .unwrap_or_else(|_| units.to_vec())
            } else { units.to_vec() };
            return Some(CellEditValue::CachedUtf16Formula { formula, reach, value: cell.value.clone() });
        }
    }
    if let Some((top, left, bottom, right)) = cell.array_block {
        if row == top && cell.col == left {
            let formula = cell.formula.clone().unwrap_or_default();
            let reach = format!(
                "{}{top}:{}{bottom}",
                col_to_letter(left),
                col_to_letter(right)
            );
            return Some(CellEditValue::CachedFormula { formula, reach: Some(reach), value: cell.value.clone() });
        }
    } else if let Some(formula) = cell.formula.as_ref() {
        return Some(CellEditValue::CachedFormula { formula: saved_formula(formula), reach: None, value: cell.value.clone() });
    }
    if cell.value.is_text() && !cell.runs.is_empty() {
        return Some(CellEditValue::RichText { value: cell.value.clone(), runs: cell.runs.clone() });
    }
    match &cell.value {
        crate::ir::CellValue::Empty => Some(CellEditValue::Empty),
        crate::ir::CellValue::String(value) => Some(CellEditValue::String(value.clone())),
        crate::ir::CellValue::Utf16(value) => Some(CellEditValue::Utf16(value.clone())),
        crate::ir::CellValue::Number(value) => Some(CellEditValue::Number(*value)),
        crate::ir::CellValue::Boolean(value) => Some(CellEditValue::Boolean(*value)),
        crate::ir::CellValue::Error(value) => Some(CellEditValue::Error(value.clone())),
    }
}

/// Adds the styles a save needs to `styles.xml`, and says which `cellXfs`
/// index each one ended up at.
///
/// Styles are appended rather than matched against what is already there. A
/// style read back out of a file is only as complete as the parser that read
/// it, so reusing an existing entry risks carrying over something the parser
/// never saw. Appending costs a few bytes and keeps every original entry as it
/// was for the cells still using it.
fn patch_styles_xml(
    xml: &str,
    wanted: &[CellStyle],
) -> Result<(String, Vec<u32>), XlsxError> {
    if wanted.is_empty() {
        return Ok((xml.to_string(), Vec::new()));
    }

    let fonts_at = count_of(xml, "fonts");
    let fills_at = count_of(xml, "fills");
    let borders_at = count_of(xml, "borders");
    let xfs_at = count_of(xml, "cellXfs");
    let mut next_number_format = 164_u32.max(
        // Custom formats are numbered from 164; step past any the file already has.
        highest_custom_number_format(xml).map_or(164, |highest| highest + 1),
    );

    let mut fonts = String::new();
    let mut fills = String::new();
    let mut borders = String::new();
    let mut xfs = String::new();
    let mut number_formats = String::new();
    let mut indices = Vec::with_capacity(wanted.len());

    for (offset, style) in wanted.iter().enumerate() {
        let offset = offset as u32;
        fonts.push_str(&font_xml(style));
        fills.push_str(&fill_xml(style));
        borders.push_str(&border_xml(style));

        let number_format_id = match style.number_format.as_deref() {
            None => 0,
            Some(format) => match builtin_number_format_id(format) {
                Some(id) => id,
                None => {
                    number_formats.push_str(&format!(
                        "<numFmt numFmtId=\"{next_number_format}\" formatCode=\"{}\"/>",
                        escape(format)
                    ));
                    let id = next_number_format;
                    next_number_format += 1;
                    id
                }
            },
        };

        xfs.push_str(&xf_xml(
            style,
            number_format_id,
            fonts_at + offset,
            fills_at + offset,
            borders_at + offset,
        ));
        indices.push(xfs_at + offset);
    }

    let added = wanted.len() as u32;
    let mut patched = xml.to_string();
    patched = append_section(&patched, "numFmts", &number_formats, true)?;
    patched = append_section(&patched, "fonts", &fonts, false)?;
    patched = append_section(&patched, "fills", &fills, false)?;
    patched = append_section(&patched, "borders", &borders, false)?;
    patched = append_section(&patched, "cellXfs", &xfs, false)?;
    let _ = added;
    Ok((patched, indices))
}

fn default_styles_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>"#.to_string()
}

fn append_workbook_styles_relationship(xml: &str) -> Result<String, XlsxError> {
    let held = parse_relationships(xml)
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    const STYLE_RELATIONSHIP: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    if held.values().any(|relation| relation.rel_type == STYLE_RELATIONSHIP) {
        return Ok(xml.to_string());
    }
    let mut suffix = 0u32;
    let id = loop {
        let candidate = if suffix == 0 { "rIdStyles".to_string() } else { format!("rIdStyles{suffix}") };
        if !held.contains_key(&candidate) { break candidate; }
        suffix += 1;
    };
    let at = xml.rfind("</Relationships>")
        .ok_or_else(|| XlsxError::InvalidData("workbook relationships have no closing element".into()))?;
    let mut patched = String::with_capacity(xml.len() + 128);
    patched.push_str(&xml[..at]);
    patched.push_str(&format!(
        "<Relationship Id=\"{}\" Type=\"{}\" Target=\"styles.xml\"/>",
        escape(&id), STYLE_RELATIONSHIP
    ));
    patched.push_str(&xml[at..]);
    Ok(patched)
}

fn patch_indexed_colors_xml(xml: &str, colors: &[String]) -> Result<String, XlsxError> {
    let mut section = String::from("<colors><indexedColors>");
    for color in colors {
        let color = color.strip_prefix('#').unwrap_or(color);
        if color.len() != 6 || !color.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(XlsxError::InvalidData(format!(
                "invalid indexed color {color:?}"
            )));
        }
        section.push_str(&format!("<rgbColor rgb=\"00{}\"/>", color.to_ascii_uppercase()));
    }
    section.push_str("</indexedColors></colors>");

    if let Some(start) = xml.find("<colors") {
        let Some(relative_end) = xml[start..].find('>') else {
            return Err(XlsxError::InvalidData("invalid colors section in styles.xml".into()));
        };
        let open_end = start + relative_end + 1;
        if xml[start..open_end].ends_with("/>") {
            let mut patched = String::with_capacity(xml.len() + section.len());
            patched.push_str(&xml[..start]);
            patched.push_str(&section);
            patched.push_str(&xml[open_end..]);
            return Ok(patched);
        }
        let Some(close_relative) = xml[open_end..].find("</colors>") else {
            return Err(XlsxError::InvalidData("unterminated colors section in styles.xml".into()));
        };
        let end = open_end + close_relative + "</colors>".len();
        let mut patched = String::with_capacity(xml.len() + section.len());
        patched.push_str(&xml[..start]);
        patched.push_str(&section);
        patched.push_str(&xml[end..]);
        return Ok(patched);
    }

    let insertion = xml.find("<extLst").or_else(|| xml.rfind("</"))
        .ok_or_else(|| XlsxError::InvalidData("styles.xml has no styleSheet end tag".into()))?;
    let mut patched = String::with_capacity(xml.len() + section.len());
    patched.push_str(&xml[..insertion]);
    patched.push_str(&section);
    patched.push_str(&xml[insertion..]);
    Ok(patched)
}

/// How many entries a section already holds, read from its `count` attribute.
fn count_of(xml: &str, section: &str) -> u32 {
    let open = format!("<{section} ");
    let Some(start) = xml.find(&open).or_else(|| xml.find(&format!("<{section}>"))) else {
        return 0;
    };
    let Some(end) = xml[start..].find('>') else {
        return 0;
    };
    let tag = &xml[start..start + end];
    tag.find("count=\"")
        .and_then(|at| {
            let rest = &tag[at + 7..];
            rest.find('"').and_then(|close| rest[..close].parse().ok())
        })
        .unwrap_or(0)
}

fn highest_custom_number_format(xml: &str) -> Option<u32> {
    let mut highest = None;
    let mut rest = xml;
    while let Some(at) = rest.find("numFmtId=\"") {
        rest = &rest[at + 10..];
        let Some(close) = rest.find('"') else { break };
        if let Ok(id) = rest[..close].parse::<u32>() {
            if id >= 164 {
                highest = Some(highest.map_or(id, |held: u32| held.max(id)));
            }
        }
    }
    highest
}

fn append_differential_styles(xml: &str, styles: &[(Option<DiffStyle>, DiffStyle)]) -> Result<(String, Vec<u32>), XlsxError> {
    let fail = || XlsxError::InvalidData("cannot safely append differential styles".into());
    let mut reader = Reader::from_str(xml);
    let mut in_dxfs = false;
    let mut count = 0u32;
    let mut has_dxfs = false;
    loop {
        match reader.read_event().map_err(|_| fail())? {
            Event::Start(ref e) if local_name(e.name().as_ref()) == "dxfs" => { in_dxfs = true; has_dxfs = true; }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == "dxfs" => { has_dxfs = true; }
            Event::Start(ref e) | Event::Empty(ref e) if in_dxfs && local_name(e.name().as_ref()) == "dxf" => count += 1,
            Event::End(ref e) if local_name(e.name().as_ref()) == "dxfs" => in_dxfs = false,
            Event::Eof => break,
            _ => {}
        }
    }
    if has_dxfs && count_of(xml, "dxfs") != count { return Err(fail()); }
    let mut entries = String::new();
    for (before, style) in styles {
        if let Some(before) = before {
            let raw = before.source_xml.as_deref().ok_or_else(fail)?;
            entries.push_str(&patch_differential_style(raw, before, style)?);
            continue;
        }
        if style.number_format.is_some() { return Err(fail()); }
        entries.push_str("<dxf>");
        if style.font_color.is_some() || style.bold.is_some() || style.italic.is_some() || style.underline.is_some() || style.strikethrough.is_some() {
            entries.push_str("<font>");
            for (tag, value) in [("b", style.bold), ("i", style.italic), ("u", style.underline), ("strike", style.strikethrough)] {
                if let Some(value) = value { entries.push_str(&format!("<{tag} val=\"{}\"/>", u8::from(value))); }
            }
            if let Some(color) = &style.font_color { entries.push_str(&format!("<color rgb=\"FF{}\"/>", differential_color(color)?)); }
            entries.push_str("</font>");
        }
        if let Some(color) = &style.bg_color { entries.push_str(&format!("<fill><patternFill><bgColor rgb=\"FF{}\"/></patternFill></fill>", differential_color(color)?)); }
        entries.push_str("</dxf>");
    }
    let indices = (0..styles.len()).map(|index| count + index as u32).collect();
    if has_dxfs { return Ok((append_section(xml, "dxfs", &entries, false)?, indices)); }
    let insertion = ["<tableStyles", "<colors", "<extLst", "</styleSheet>"]
        .iter().filter_map(|tag| xml.find(tag)).min().ok_or_else(fail)?;
    Ok((format!("{}<dxfs count=\"{}\">{entries}</dxfs>{}", &xml[..insertion], styles.len(), &xml[insertion..]), indices))
}

fn differential_color(color: &str) -> Result<String, XlsxError> {
    let color = color.strip_prefix('#').unwrap_or(color);
    if color.len() != 6 || !color.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(XlsxError::InvalidData("invalid differential color".into()));
    }
    Ok(color.to_ascii_uppercase())
}

/// Find a direct child and the insertion point of an XML element, preserving
/// the original bytes of all unrelated children and attributes.
fn direct_xml_child(xml: &str, wanted: &str) -> Result<(Option<(usize, usize)>, usize, bool, String), XlsxError> {
    let fail = || XlsxError::InvalidData("cannot safely edit differential style".into());
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut root = String::new();
    let mut root_head = 0usize;
    let mut child_start = None;
    let mut found = None;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|_| fail())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref e) => {
                if depth == 0 { root = String::from_utf8(e.name().as_ref().to_vec()).map_err(|_| fail())?; root_head = end; }
                else if depth == 1 && local_name(e.name().as_ref()) == wanted {
                    if found.is_some() || child_start.is_some() { return Err(fail()); }
                    child_start = Some(start);
                }
                depth += 1;
            }
            Event::Empty(ref e) if depth == 0 => {
                root = String::from_utf8(e.name().as_ref().to_vec()).map_err(|_| fail())?;
                return Ok((None, end, true, root));
            }
            Event::Empty(ref e) if depth == 1 && local_name(e.name().as_ref()) == wanted => {
                if found.is_some() || child_start.is_some() { return Err(fail()); }
                found = Some((start, end));
            }
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(fail)?;
                if depth == 1 { if let Some(child) = child_start.take() { found = Some((child, end)); } }
                if depth == 0 { return Ok((found, start, false, root)); }
            }
            Event::Eof => return Err(fail()),
            _ => {}
        }
        if root_head > xml.len() { return Err(fail()); }
    }
}

fn replace_direct_xml_child(xml: &str, wanted: &str, replacement: Option<&str>) -> Result<String, XlsxError> {
    let (child, insert, empty, root) = direct_xml_child(xml, wanted)?;
    if let Some((start, end)) = child { return Ok(format!("{}{}{}", &xml[..start], replacement.unwrap_or(""), &xml[end..])); }
    let Some(replacement) = replacement else { return Ok(xml.to_string()); };
    if empty {
        if insert < 2 || &xml[insert - 2..insert] != "/>" { return Err(XlsxError::InvalidData("invalid empty differential element".into())); }
        return Ok(format!("{}>{replacement}</{root}>{}", &xml[..insert - 2], &xml[insert..]));
    }
    Ok(format!("{}{replacement}{}", &xml[..insert], &xml[insert..]))
}

fn patch_differential_style(xml: &str, before: &DiffStyle, after: &DiffStyle) -> Result<String, XlsxError> {
    if before.number_format != after.number_format || before.source_xml != after.source_xml {
        return Err(XlsxError::InvalidData("unsupported differential style change".into()));
    }
    let mut output = xml.to_string();
    let font_changed = before.bold != after.bold || before.italic != after.italic || before.underline != after.underline
        || before.strikethrough != after.strikethrough || before.font_color != after.font_color;
    if font_changed {
        let (span, _, _, _) = direct_xml_child(&output, "font")?;
        let mut font = span.map(|(start, end)| output[start..end].to_string()).unwrap_or_else(|| "<font/>".into());
        for (tag, old, new) in [("b", before.bold, after.bold), ("i", before.italic, after.italic),
            ("u", before.underline, after.underline), ("strike", before.strikethrough, after.strikethrough)] {
            if old != new {
                let element = new.map(|value| format!("<{tag} val=\"{}\"/>", u8::from(value)));
                font = replace_direct_xml_child(&font, tag, element.as_deref())?;
            }
        }
        if before.font_color != after.font_color {
            let element = after.font_color.as_ref().map(|color| differential_color(color).map(|color| format!("<color rgb=\"FF{color}\"/>"))).transpose()?;
            font = replace_direct_xml_child(&font, "color", element.as_deref())?;
        }
        output = replace_direct_xml_child(&output, "font", Some(&font))?;
    }
    if before.bg_color != after.bg_color {
        let (span, _, _, _) = direct_xml_child(&output, "fill")?;
        let mut fill = span.map(|(start, end)| output[start..end].to_string()).unwrap_or_else(|| "<fill/>".into());
        if direct_xml_child(&fill, "gradientFill")?.0.is_some() { return Err(XlsxError::InvalidData("cannot safely change a gradient differential fill".into())); }
        let (span, _, _, _) = direct_xml_child(&fill, "patternFill")?;
        let mut pattern = span.map(|(start, end)| fill[start..end].to_string()).unwrap_or_else(|| "<patternFill/>".into());
        let element = after.bg_color.as_ref().map(|color| differential_color(color).map(|color| format!("<bgColor rgb=\"FF{color}\"/>"))).transpose()?;
        pattern = replace_direct_xml_child(&pattern, "bgColor", element.as_deref())?;
        if after.bg_color.is_none() { pattern = replace_direct_xml_child(&pattern, "fgColor", None)?; }
        fill = replace_direct_xml_child(&fill, "patternFill", Some(&pattern))?;
        output = replace_direct_xml_child(&output, "fill", Some(&fill))?;
    }
    Ok(output)
}

/// Puts new entries at the end of a section, growing its `count`. A section
/// that is not there is created when `create` is set, which only numFmts needs.
fn append_section(
    xml: &str,
    section: &str,
    entries: &str,
    create: bool,
) -> Result<String, XlsxError> {
    if entries.is_empty() {
        return Ok(xml.to_string());
    }
    let added = entries.matches("<numFmt ").count().max(
        entries
            .matches(&format!("<{}", singular(section)))
            .count(),
    ) as u32;

    let closing = format!("</{section}>");
    if let Some(at) = xml.find(&closing) {
        let mut patched = String::with_capacity(xml.len() + entries.len());
        patched.push_str(&xml[..at]);
        patched.push_str(entries);
        patched.push_str(&xml[at..]);
        return Ok(bump_count(&patched, section, added));
    }

    // A self-closing section, such as <numFmts count="0"/>.
    let empty = format!("<{section} ");
    if let Some(at) = xml.find(&empty) {
        if let Some(end) = xml[at..].find("/>") {
            let head = &xml[..at + end];
            let mut patched = String::with_capacity(xml.len() + entries.len());
            patched.push_str(head);
            patched.push('>');
            patched.push_str(entries);
            patched.push_str(&closing);
            patched.push_str(&xml[at + end + 2..]);
            return Ok(bump_count(&patched, section, added));
        }
    }

    if !create {
        return Err(XlsxError::InvalidData(format!(
            "the workbook's styles have no {section} to add to"
        )));
    }
    // numFmts belongs at the head of the style sheet, ahead of the fonts.
    let at = xml.find("<fonts").ok_or_else(|| {
        XlsxError::InvalidData("the workbook's styles have no fonts".to_string())
    })?;
    let mut patched = String::with_capacity(xml.len() + entries.len() + 40);
    patched.push_str(&xml[..at]);
    patched.push_str(&format!("<{section} count=\"{added}\">"));
    patched.push_str(entries);
    patched.push_str(&closing);
    patched.push_str(&xml[at..]);
    Ok(patched)
}

fn singular(section: &str) -> &str {
    match section {
        "fonts" => "font",
        "fills" => "fill",
        "borders" => "border",
        "cellXfs" => "xf",
        "dxfs" => "dxf",
        other => other,
    }
}

fn bump_count(xml: &str, section: &str, added: u32) -> String {
    let open = format!("<{section} ");
    let Some(start) = xml.find(&open) else {
        return xml.to_string();
    };
    let Some(end) = xml[start..].find('>') else {
        return xml.to_string();
    };
    let tag = &xml[start..start + end];
    let Some(at) = tag.find("count=\"") else {
        return xml.to_string();
    };
    let rest = &tag[at + 7..];
    let Some(close) = rest.find('"') else {
        return xml.to_string();
    };
    let held: u32 = rest[..close].parse().unwrap_or(0);
    let replaced = format!("{}count=\"{}\"{}", &tag[..at], held + added, &rest[close + 1..]);
    format!("{}{}{}", &xml[..start], replaced, &xml[start + end..])
}

fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn resolve_part_target(from: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.trim_start_matches('/').to_string();
    }
    let mut parts: Vec<&str> = from.rsplit_once('/').map_or("", |(dir, _)| dir).split('/').collect();
    parts.retain(|part| !part.is_empty());
    for step in target.split('/') {
        match step {
            "" | "." => {}
            ".." => { parts.pop(); }
            step => parts.push(step),
        }
    }
    parts.join("/")
}

fn patch_threaded_comment_xml(
    xml: &str,
    replacements: &HashMap<String, String>,
    deletions: &std::collections::HashSet<String>,
    resolved: &HashMap<String, bool>,
) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut current_id: Option<String> = None;
    let mut in_text = false;
    let mut wrote_text = false;
    let mut skipping_comment = false;
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        match &event {
            Event::Start(element) if local_name(element.name().as_ref()) == "threadedComment" => {
                current_id = get_attr(element, "id");
                wrote_text = false;
                if current_id.as_ref().is_some_and(|id| deletions.contains(id)) {
                    skipping_comment = true;
                    continue;
                }
                if let Some(done) = current_id.as_ref().and_then(|id| resolved.get(id)) {
                    let mut patched = element.to_owned();
                    patched.clear_attributes();
                    for attribute in element.attributes().flatten() {
                        if attribute.key.as_ref() != b"done" { patched.push_attribute(attribute); }
                    }
                    patched.push_attribute(("done", if *done { "1" } else { "0" }));
                    writer.write_event(Event::Start(patched))
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    continue;
                }
            }
            Event::Empty(element) if local_name(element.name().as_ref()) == "threadedComment"
                && get_attr(element, "id").is_some_and(|id| deletions.contains(&id)) => continue,
            Event::Start(element) if local_name(element.name().as_ref()) == "text" => in_text = true,
            Event::Empty(element) if local_name(element.name().as_ref()) == "text" => {
                if let Some(text) = current_id.as_ref().and_then(|id| replacements.get(id)) {
                    let name = String::from_utf8(element.name().as_ref().to_vec())
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    writer.write_event(Event::Start(BytesStart::new(&name)))
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    writer.write_event(Event::Text(BytesText::new(text)))
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    writer.write_event(Event::End(BytesEnd::new(&name)))
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    wrote_text = true;
                    continue;
                }
            }
            Event::Text(_) if in_text && !wrote_text => {
                if let Some(text) = current_id.as_ref().and_then(|id| replacements.get(id)) {
                    writer.write_event(Event::Text(BytesText::new(text)))
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    wrote_text = true;
                    continue;
                }
            }
            Event::End(element) if local_name(element.name().as_ref()) == "text" => {
                if in_text && !wrote_text {
                    if let Some(text) = current_id.as_ref().and_then(|id| replacements.get(id)) {
                        writer.write_event(Event::Text(BytesText::new(text)))
                            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                        wrote_text = true;
                    }
                }
                in_text = false;
            }
            Event::End(element) if local_name(element.name().as_ref()) == "threadedComment" => {
                current_id = None;
                if skipping_comment {
                    skipping_comment = false;
                    continue;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        if skipping_comment { continue; }
        writer.write_event(event.into_owned())
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }
    let bytes = writer.into_inner().into_inner();
    String::from_utf8(bytes).map_err(|error| XlsxError::InvalidData(error.to_string()))
}

#[derive(Default)]
struct ThreadedCommentPackageEdits {
    comments: HashMap<String, ThreadedCommentPartEdit>,
    relationship_xml: HashMap<String, String>,
    person_xml: HashMap<String, String>,
    new_parts: Vec<(String, Vec<u8>)>,
    content_types: Vec<(String, String)>,
}

#[derive(Default)]
struct ThreadedCommentPartEdit {
    text: HashMap<String, String>,
    resolved: HashMap<String, bool>,
    additions: Vec<crate::ir::ThreadedComment>,
    deletions: std::collections::HashSet<String>,
}

#[derive(Default)]
struct LegacyCommentPackageEdits {
    parts: HashMap<String, String>,
    new_parts: Vec<(String, String)>,
    relationship_xml: HashMap<String, String>,
    worksheet_drawing: HashMap<String, String>,
    content_types: Vec<(String, String)>,
    needs_vml_default: bool,
}

fn legacy_comment_package_edits(
    archive: &mut OoxmlArchive,
    sheet_paths: &[String],
    changed: &HashMap<usize, Vec<crate::ir::Comment>>,
    original_sheets: &[crate::ir::Sheet],
    threaded: &ThreadedCommentPackageEdits,
    drawing: &DrawingAdditionPackage,
) -> Result<LegacyCommentPackageEdits, XlsxError> {
    let mut plan = LegacyCommentPackageEdits::default();
    let mut names: std::collections::HashSet<String> = archive.file_names().into_iter().collect();
    let archive_names = names.clone();
    for (sheet_index, comments) in changed {
        let Some(sheet_path) = sheet_paths.get(*sheet_index) else { continue };
        let rels_path = sheet_rels_path(sheet_path);
        let mut rels_xml = archive.try_read_part(&rels_path)?;
        let drawing_rels = drawing.relationship_xml.get(&rels_path).cloned().or_else(|| {
            drawing.new_parts.iter().find(|(path, _)| path == &rels_path).map(|(_, xml)| xml.clone())
        });
        if let Some(additions) = drawing_rels {
            rels_xml = Some(if let Some(original) = rels_xml.as_deref() {
                merge_relationship_additions(original, &additions)?
            } else { additions });
        }
        let threaded_rels = threaded.relationship_xml.get(&rels_path).cloned().or_else(|| {
            threaded.new_parts.iter().find(|(path, _)| path == &rels_path)
                .map(|(_, bytes)| String::from_utf8_lossy(bytes).into_owned())
        });
        if let Some(additions) = threaded_rels {
            rels_xml = Some(if let Some(original) = rels_xml.as_deref() {
                merge_relationship_additions(original, &additions)?
            } else {
                additions
            });
        }
        let rels = rels_xml.as_deref().map(parse_relationships).transpose()
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?.unwrap_or_default();
        let comment_path = rels.values().find(|rel| rel.rel_type.ends_with("/comments"))
            .map(|rel| resolve_part_target(sheet_path, &rel.target));
        let vml_relation = rels.iter().find(|(_, rel)| rel.rel_type.ends_with("/vmlDrawing"));
        let existing_vml_rid = vml_relation.map(|(id, _)| id.clone());
        let vml_path = vml_relation.map(|(_, rel)| resolve_part_target(sheet_path, &rel.target));

        let comment_path = match comment_path {
            Some(path) => path,
            None => {
                let path = unique_part_path(&names, "xl", "comments", ".xml");
                names.insert(path.clone());
                let target = relative_sheet_target(sheet_path, &path);
                rels_xml = Some(append_relationship(
                    rels_xml.as_deref(),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                    &target,
                )?);
                plan.content_types.push((path.clone(), "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml".into()));
                plan.new_parts.push((path.clone(), String::new()));
                path
            }
        };
        let mut legacy_drawing_rid = existing_vml_rid;
        let vml_path = match vml_path {
            Some(path) => path,
            None => {
                let path = unique_part_path(&names, "xl/drawings", "vmlDrawing", ".vml");
                names.insert(path.clone());
                let target = relative_sheet_target(sheet_path, &path);
                let before = rels_xml.clone();
                let updated = append_relationship(
                    rels_xml.as_deref(),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
                    &target,
                )?;
                let parsed_before = before.as_deref().map(parse_relationships).transpose()
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?.unwrap_or_default();
                let parsed_after = parse_relationships(&updated)
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                legacy_drawing_rid = parsed_after.keys().find(|id| !parsed_before.contains_key(*id)).cloned();
                rels_xml = Some(updated);
                plan.new_parts.push((path.clone(), empty_vml_drawing()));
                plan.needs_vml_default = true;
                path
            }
        };

        let original_comment_part_exists = archive_names.contains(&comment_path);
        let comments_xml = if original_comment_part_exists {
            let original_xml = archive.try_read_part(&comment_path)?
                .ok_or_else(|| XlsxError::InvalidData("missing comments part".into()))?;
            patch_legacy_comments(&original_xml, &original_sheets[*sheet_index].comments, comments)?
        } else {
            write_legacy_comments(comments)
        };
        if original_comment_part_exists {
            plan.parts.insert(comment_path.clone(), comments_xml);
        } else if let Some((_, xml)) = plan.new_parts.iter_mut().find(|(path, _)| path == &comment_path) {
            *xml = comments_xml;
        } else {
            plan.new_parts.push((comment_path.clone(), comments_xml));
        }

        let original_vml = archive.try_read_part(&vml_path)?.unwrap_or_else(empty_vml_drawing);
        let vml = patch_vml_notes(&original_vml, comments)?;
        if archive_names.contains(&vml_path) {
            plan.parts.insert(vml_path, vml);
        } else if let Some((_, xml)) = plan.new_parts.iter_mut().find(|(path, _)| path == &vml_path) {
            *xml = vml;
        } else {
            plan.new_parts.push((vml_path, vml));
        }
        if let Some(updated) = rels_xml {
            if archive_names.contains(&rels_path) {
                plan.relationship_xml.insert(rels_path, updated);
            } else {
                plan.new_parts.push((rels_path, updated));
            }
        }
        if let Some(rid) = legacy_drawing_rid {
            plan.worksheet_drawing.insert(sheet_path.clone(), rid);
        }
    }
    Ok(plan)
}

fn unique_part_path(names: &std::collections::HashSet<String>, directory: &str, stem: &str, extension: &str) -> String {
    let mut index = 1;
    loop {
        let path = format!("{directory}/{stem}{index}{extension}");
        if !names.contains(&path) { return path; }
        index += 1;
    }
}

fn relative_sheet_target(sheet_path: &str, part_path: &str) -> String {
    let (directory, _) = sheet_path.rsplit_once('/').unwrap_or(("", sheet_path));
    let mut from: Vec<&str> = directory.split('/').filter(|piece| !piece.is_empty()).collect();
    let mut to: Vec<&str> = part_path.split('/').filter(|piece| !piece.is_empty()).collect();
    while from.first() == to.first() && !from.is_empty() {
        from.remove(0);
        to.remove(0);
    }
    let mut result = vec![".."; from.len()];
    result.extend(to);
    result.join("/")
}

fn legacy_comments_equivalent(before: &[crate::ir::Comment], after: &[crate::ir::Comment]) -> bool {
    fn normalized(comments: &[crate::ir::Comment]) -> Vec<crate::ir::Comment> {
        let mut comments = comments.to_vec();
        for comment in &mut comments {
            for paragraph in &mut comment.text.paragraphs {
                if paragraph.align.as_deref() == Some("l") { paragraph.align = None; }
                if let [run] = paragraph.runs.as_slice() {
                    if run.text == paragraph.text && run.bold == paragraph.bold
                        && run.italic.unwrap_or(paragraph.italic) == paragraph.italic && !run.underline
                        && run.face.as_ref().or(paragraph.face.as_ref()) == paragraph.face.as_ref()
                        && run.size.unwrap_or(paragraph.size) == paragraph.size
                        && run.color.as_ref().or(paragraph.color.as_ref()) == paragraph.color.as_ref()
                        && run.color_theme.is_none() && run.color_brightness.is_none() && run.legacy_color_index.is_none()
                    { paragraph.runs.clear(); }
                }
            }
        }
        comments
    }
    normalized(before) == normalized(after)
}

fn threaded_compatibility_note(thread: &crate::ir::ThreadedComment) -> crate::ir::Comment {
    let text = format!("[スレッド化されたコメント]\n\n使用している Excel のバージョンでは、このスレッド化されたコメントを表示できますが、新しいバージョンの Excel でファイルを開いた場合はコメントに対する編集がすべて削除されます。詳細: https://go.microsoft.com/fwlink/?linkid=870924\n\nコメント:\n    {}", thread.text);
    let mut units: Vec<u16> = text[..text.len() - thread.text.len()].encode_utf16().collect();
    units.extend(thread.utf16_text.clone().unwrap_or_else(|| thread.text.encode_utf16().collect()));
    let utf16_text = String::from_utf16(&units).err().map(|_| units);
    let (row, col) = thread.cell;
    crate::ir::Comment {
        author: format!("tc={}", thread.id), utf16_text,
        from: crate::ir::Anchor { col: col.saturating_add(1), col_off: 0, row, row_off: 0 },
        to: None, size: (96.0, 59.25),
        text: crate::ir::ShapeText { paragraphs: text.split('\n').map(|line| crate::ir::ShapeParagraph {
            text: line.into(), runs: Vec::new(), align: None, size: 9.0, bold: false, italic: false,
            face: None, charset: None, pitch_family: None, color: None, line_pitch: None, line_scale: None,
        }).collect(), anchor: Some("t".into()), insets: (91_440, 45_720, 91_440, 45_720), wrap: true, clip: true },
        fill: Some("FFFFE1".into()), fill_visible: true, fill_transparency: 0.0,
        line_color: None, line_visible: true, line_weight: 0.75, line_transparency: 0.0,
        cell: thread.cell, visible: false,
    }
}

fn refresh_threaded_compatibility_note(note: &mut crate::ir::Comment, root: &crate::ir::ThreadedComment, replies: &[&crate::ir::ThreadedComment]) {
    let existing = note.text.paragraphs.iter().map(|paragraph| paragraph.text.as_str()).collect::<Vec<_>>().join("\n");
    let (header, reply_label) = if let Some((header, _)) = existing.split_once("\n\nコメント:\n") {
        (format!("{header}\n\nコメント:\n    "), "返信")
    } else if let Some((header, _)) = existing.split_once("\n\nComment:\n") {
        (format!("{header}\n\nComment:\n    "), "Reply")
    } else {
        let fallback = threaded_compatibility_note(root);
        let text = fallback.text.paragraphs.iter().map(|paragraph| paragraph.text.as_str()).collect::<Vec<_>>().join("\n");
        (text.split_once("\n\nコメント:\n").map(|(header, _)| format!("{header}\n\nコメント:\n    ")).unwrap_or_default(), "返信")
    };
    let mut units: Vec<u16> = header.encode_utf16().collect();
    units.extend(root.utf16_text.clone().unwrap_or_else(|| root.text.encode_utf16().collect()));
    for reply in replies {
        units.extend(format!("\n{reply_label}:\n").encode_utf16());
        units.extend(reply.utf16_text.clone().unwrap_or_else(|| reply.text.encode_utf16().collect()));
    }
    let text = String::from_utf16_lossy(&units);
    let utf16_text = String::from_utf16(&units).err().map(|_| units);
    if text == existing && utf16_text == note.utf16_text { return; }
    if let Some(style) = note.text.paragraphs.first().cloned() {
        note.text.paragraphs = text.split('\n').map(|line| { let mut paragraph = style.clone(); paragraph.text = line.into(); paragraph.runs.clear(); paragraph }).collect();
        note.utf16_text = utf16_text;
    }
}

fn write_legacy_comments(comments: &[crate::ir::Comment]) -> String {
    let mut comments = comments.to_vec();
    comments.sort_by_key(|comment| comment.cell);
    let mut authors = Vec::<String>::new();
    for comment in &comments {
        if !authors.contains(&comment.author) { authors.push(comment.author.clone()); }
    }
    let author_xml = authors.iter().map(|author| format!("<author>{}</author>", escape(author))).collect::<String>();
    let entries = comments.iter().map(|comment| {
        let author_id = authors.iter().position(|author| author == &comment.author).unwrap_or(0);
        write_legacy_comment(comment, author_id)
    }).collect::<String>();
    format!("<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors>{author_xml}</authors><commentList>{entries}</commentList></comments>")
}

fn write_legacy_comment(comment: &crate::ir::Comment, author_id: usize) -> String {
        let text = comment.text.paragraphs.iter().map(|paragraph| paragraph.text.as_str()).collect::<Vec<_>>().join("\n");
        let mut runs = String::new();
        let paragraphs = &comment.text.paragraphs;
        let emit = |text: &str, p: Option<&crate::ir::ShapeParagraph>, r: Option<&crate::ir::ShapeRun>| {
            let mut props = String::new();
            if r.map(|r| r.bold).or_else(|| p.map(|p| p.bold)).unwrap_or(false) { props.push_str("<b/>"); }
            if r.and_then(|r| r.italic).or_else(|| p.map(|p| p.italic)).unwrap_or(false) { props.push_str("<i/>"); }
            if r.is_some_and(|r| r.underline) { props.push_str("<u/>"); }
            if let Some(size) = r.and_then(|r| r.size).or_else(|| p.map(|p| p.size)) { props.push_str(&format!("<sz val=\"{size}\"/>")); }
            if let Some(color) = r.and_then(|r| r.color.as_deref()).or_else(|| p.and_then(|p| p.color.as_deref())) { props.push_str(&format!("<color rgb=\"FF{}\"/>", escape(color.trim_start_matches('#')))); }
            if let Some(face) = r.and_then(|r| r.face.as_deref()).or_else(|| p.and_then(|p| p.face.as_deref())) { props.push_str(&format!("<rFont val=\"{}\"/>", escape(face))); }
            if let Some(charset) = p.and_then(|p| p.charset) { props.push_str(&format!("<charset val=\"{charset}\"/>")); }
            let props = if props.is_empty() { String::new() } else { format!("<rPr>{props}</rPr>") };
            format!("<r>{props}<t xml:space=\"preserve\">{}</t></r>", escape(text))
        };
        if let Some(units) = &comment.utf16_text {
            runs.push_str(&emit(&crate::string_value::encode(units), paragraphs.first(), None));
        } else if paragraphs.iter().any(|p| !p.runs.is_empty()) {
            for (index, paragraph) in paragraphs.iter().enumerate() {
                if paragraph.runs.is_empty() {
                    let mut line = paragraph.text.clone();
                    if index + 1 < paragraphs.len() { line.push('\n'); }
                    runs.push_str(&emit(&crate::string_value::encode(&line.encode_utf16().collect::<Vec<_>>()), Some(paragraph), None));
                } else {
                    for (position, run) in paragraph.runs.iter().enumerate() {
                        let mut line = run.text.clone();
                        if position + 1 == paragraph.runs.len() && index + 1 < paragraphs.len() { line.push('\n'); }
                        runs.push_str(&emit(&crate::string_value::encode(&line.encode_utf16().collect::<Vec<_>>()), Some(paragraph), Some(run)));
                    }
                }
            }
        } else {
            runs.push_str(&emit(&crate::string_value::encode(&text.encode_utf16().collect::<Vec<_>>()), paragraphs.first(), None));
        }
        let (row, col) = comment.cell;
        let uid = comment.author.strip_prefix("tc=").map(|id| format!(" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xr:uid=\"{}\"", escape(id))).unwrap_or_default();
        format!("<comment ref=\"{}{}\" authorId=\"{author_id}\" shapeId=\"0\"{uid}><text>{runs}</text></comment>", col_to_letter(col), row + 1)
}

// Preserve the original comment elements whenever their content has not
// changed. Excel stores rich text and threaded-comment identifiers there.
fn patch_legacy_comments(xml: &str, before: &[crate::ir::Comment], after: &[crate::ir::Comment]) -> Result<String, XlsxError> {
    let fail = || XlsxError::InvalidData("invalid legacy comments part".into());
    let (authors_span, _, _, _) = direct_xml_child(xml, "authors")?;
    let (list_span, _, _, _) = direct_xml_child(xml, "commentList")?;
    let (authors_start, authors_end) = authors_span.ok_or_else(fail)?;
    let (list_start, list_end) = list_span.ok_or_else(fail)?;
    let authors_xml = &xml[authors_start..authors_end];
    let list_xml = &xml[list_start..list_end];
    let mut reader = Reader::from_str(authors_xml);
    let mut authors = Vec::<String>::new();
    let mut current = None::<String>;
    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(e) if local_name(e.name().as_ref()) == "author" => current = Some(String::new()),
            Event::Empty(e) if local_name(e.name().as_ref()) == "author" => authors.push(String::new()),
            Event::Text(t) if current.is_some() => current.as_mut().unwrap().push_str(&t.unescape().map_err(XlsxError::Xml)?),
            Event::End(e) if local_name(e.name().as_ref()) == "author" => authors.push(current.take().ok_or_else(fail)?),
            Event::Eof => break,
            _ => {}
        }
    }
    let mut raw = HashMap::<String, &str>::new();
    let mut reader = Reader::from_str(list_xml);
    let mut depth = 0usize;
    let mut held = None::<(usize, String)>;
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Start(e) => {
                if depth == 1 && local_name(e.name().as_ref()) == "comment" {
                    held = Some((start, get_attr(&e, "ref").ok_or_else(fail)?));
                }
                depth += 1;
            }
            Event::Empty(e) if depth == 1 && local_name(e.name().as_ref()) == "comment" => {
                let key = get_attr(&e, "ref").ok_or_else(fail)?;
                if raw.insert(key, &list_xml[start..reader.buffer_position() as usize]).is_some() { return Err(fail()); }
            }
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(fail)?;
                if depth == 1 {
                    if let Some((begin, key)) = held.take() {
                        if raw.insert(key, &list_xml[begin..reader.buffer_position() as usize]).is_some() { return Err(fail()); }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let mut added_authors = String::new();
    let mut sorted = after.to_vec();
    sorted.sort_by_key(|comment| comment.cell);
    let mut entries = String::new();
    for comment in &sorted {
        let address = format!("{}{}", col_to_letter(comment.cell.1), comment.cell.0 + 1);
        if before.iter().any(|old| old.cell == comment.cell && old.author == comment.author && old.utf16_text == comment.utf16_text && old.text.paragraphs == comment.text.paragraphs) {
            if let Some(original) = raw.get(&address) { entries.push_str(original); continue; }
        }
        let author_id = match authors.iter().position(|author| author == &comment.author) {
            Some(index) => index,
            None => {
                let index = authors.len();
                authors.push(comment.author.clone());
                added_authors.push_str(&format!("<author>{}</author>", escape(&comment.author)));
                index
            }
        };
        let replacement = write_legacy_comment(comment, author_id);
        if before.iter().any(|old| old.cell == comment.cell && old.author == comment.author) {
            if let Some(original) = raw.get(&address) {
                let (text_span, _, _, _) = direct_xml_child(&replacement, "text")?;
                let (start, end) = text_span.ok_or_else(fail)?;
                entries.push_str(&replace_direct_xml_child(original, "text", Some(&replacement[start..end]))?);
                continue;
            }
        }
        entries.push_str(&replacement);
    }
    let new_authors = if added_authors.is_empty() { authors_xml.to_string() } else { append_xml_child(authors_xml, "authors", &added_authors)? };
    let (_, list_close, list_empty, list_root) = direct_xml_child(list_xml, "__no_such_child__")?;
    let new_list = if list_empty {
        format!("{}>{entries}</{list_root}>", &list_xml[..list_close - 2])
    } else {
        let mut reader = Reader::from_str(list_xml);
        loop {
            match reader.read_event().map_err(XlsxError::Xml)? {
                Event::Start(_) => break,
                Event::Eof => return Err(fail()),
                _ => {}
            }
        }
        let head = reader.buffer_position() as usize;
        format!("{}{entries}{}", &list_xml[..head], &list_xml[list_close..])
    };
    let mut replacements = vec![(authors_start, authors_end, new_authors), (list_start, list_end, new_list)];
    replacements.sort_by_key(|(start, _, _)| std::cmp::Reverse(*start));
    let mut output = xml.to_string();
    for (start, end, replacement) in replacements { output.replace_range(start..end, &replacement); }
    Ok(output)
}

fn empty_vml_drawing() -> String {
    "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype></xml>".to_string()
}

fn patch_vml_notes(xml: &str, comments: &[crate::ir::Comment]) -> Result<String, XlsxError> {
    let mut base = xml.to_string();
    if base.is_empty() { base = empty_vml_drawing(); }
    let mut kept = String::with_capacity(base.len());
    let mut cursor = 0;
    let mut max_id = 1024u32;
    while let Some(relative) = base[cursor..].find("<v:shape ") {
        let start = cursor + relative;
        kept.push_str(&base[cursor..start]);
        let Some(end_rel) = base[start..].find("</v:shape>") else { return Err(XlsxError::InvalidData("unterminated VML shape".into())); };
        let end = start + end_rel + "</v:shape>".len();
        let shape = &base[start..end];
        for capture in shape.split("_x0000_s").skip(1) {
            let digits: String = capture.chars().take_while(char::is_ascii_digit).collect();
            if let Ok(id) = digits.parse::<u32>() { max_id = max_id.max(id); }
        }
        if !shape.contains("ObjectType=\"Note\"") && !shape.contains("ObjectType='Note'") { kept.push_str(shape); }
        cursor = end;
    }
    kept.push_str(&base[cursor..]);
    let at = kept.rfind("</xml>").ok_or_else(|| XlsxError::InvalidData("VML drawing has no closing xml element".into()))?;
    let mut generated = String::new();
    let mut sorted = comments.to_vec();
    sorted.sort_by_key(|comment| comment.cell);
    for (ordinal, comment) in sorted.iter().enumerate() {
        generated.push_str(&write_vml_note(comment, max_id + ordinal as u32 + 1, ordinal as u32 + 1));
    }
    kept.insert_str(at, &generated);
    Ok(kept)
}

fn write_vml_note(comment: &crate::ir::Comment, shape_id: u32, z_index: u32) -> String {
    let (row, col) = comment.cell;
    let from = comment.from;
    let to = comment.to.unwrap_or(crate::ir::Anchor {
        col: from.col,
        col_off: from.col_off + (f64::from(comment.size.0) / 0.75 * 9525.0) as i64,
        row: from.row,
        row_off: from.row_off + (f64::from(comment.size.1) / 0.75 * 9525.0) as i64,
    });
    let px = |value: i64| (value as f64 / 9525.0).round() as i64;
    let fill = vml_fill_color(comment.fill.as_deref().unwrap_or("FFFFE1"));
    let fill_opacity = vml_opacity_attribute(comment.fill_transparency);
    let line = comment.line_color.as_deref().map(vml_line_color)
        .unwrap_or_else(|| "none [81]".to_string());
    let line_opacity = vml_opacity_attribute(comment.line_transparency);
    let stroke = if line_opacity.is_empty() { String::new() } else { format!("<v:stroke{line_opacity}/>") };
    let filled = if comment.fill_visible { "" } else { " filled=\"f\"" };
    let stroked = if comment.line_visible { "" } else { " stroked=\"f\"" };
    let visibility = if comment.visible { "visible" } else { "hidden" };
    let visible = if comment.visible { "<x:Visible/>" } else { "" };
    let anchor = format!("{}, {}, {}, {}, {}, {}, {}, {}", from.col, px(from.col_off), from.row, px(from.row_off), to.col, px(to.col_off), to.row, px(to.row_off));
    format!("<v:shape id=\"_x0000_s{shape_id}\" type=\"#_x0000_t202\" style='position:absolute;margin-left:0pt;margin-top:0pt;width:{}pt;height:{}pt;z-index:{};visibility:{}'{filled} fillcolor=\"{}\" strokecolor=\"{}\" strokeweight=\"{}pt\"{stroked} o:insetmode=\"auto\"><v:fill{fill_opacity} color2=\"{}\"/>{stroke}<v:shadow color=\"none [81]\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><v:textbox style='mso-direction-alt:auto' inset=\"{}pt,{}pt,{}pt,{}pt\"><div style='text-align:left'></div></v:textbox><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{row}</x:Row><x:Column>{col}</x:Column>{visible}</x:ClientData></v:shape>", comment.size.0, comment.size.1, z_index, visibility, escape(&fill), escape(&line), comment.line_weight, escape(&fill), comment.text.insets.0 as f64 / 12700.0, comment.text.insets.1 as f64 / 12700.0, comment.text.insets.2 as f64 / 12700.0, comment.text.insets.3 as f64 / 12700.0)
}

fn vml_opacity_attribute(transparency: f32) -> String {
    if transparency <= 0.0 {
        return String::new();
    }
    let opacity = ((1.0 - f64::from(transparency.clamp(0.0, 1.0))) * 65_536.0).round() as u32;
    format!(" opacity=\"{opacity}f\"")
}

fn vml_fill_color(color: &str) -> String {
    let normalized = color.trim_start_matches('#').trim();
    if normalized.eq_ignore_ascii_case("FFFFE1")
        || normalized.eq_ignore_ascii_case("infoBackground [80]")
    {
        "infoBackground [80]".into()
    } else {
        format!("#{normalized}")
    }
}

fn vml_line_color(color: &str) -> String {
    let hex = color.trim_start_matches('#').to_ascii_uppercase();
    match hex.as_str() {
        "000000" => "black".into(),
        "FFFFFF" => "white".into(),
        "FF0000" => "red".into(),
        "00FF00" => "lime".into(),
        "0000FF" => "blue".into(),
        "FFFF00" => "yellow".into(),
        "00FFFF" => "cyan".into(),
        "FF00FF" => "magenta".into(),
        "808080" => "gray".into(),
        "008000" => "green".into(),
        "800000" => "maroon".into(),
        "000080" => "navy".into(),
        "808000" => "olive".into(),
        "800080" => "purple".into(),
        "008080" => "teal".into(),
        _ => format!("#{hex}"),
    }
}

fn threaded_comment_package_edits(
    archive: &mut OoxmlArchive,
    sheet_paths: &[String],
    drawing: &DrawingAdditionPackage,
    text_edits: &HashMap<(usize, String), (String, Option<Vec<u16>>)>,
    additions: &HashMap<usize, Vec<crate::ir::ThreadedComment>>,
    deletions: &HashMap<usize, std::collections::HashSet<String>>,
    resolved: &HashMap<(usize, String), bool>,
) -> Result<ThreadedCommentPackageEdits, XlsxError> {
    let mut plan = ThreadedCommentPackageEdits::default();
    let names: std::collections::HashSet<String> = archive.file_names().into_iter().collect();
    let mut rels_cache = HashMap::<String, String>::new();
    for sheet_path in sheet_paths {
        let rels_path = sheet_rels_path(sheet_path);
        let drawing_rels = drawing.relationship_xml.get(&rels_path).cloned().or_else(|| {
            drawing.new_parts.iter().find(|(path, _)| path == &rels_path).map(|(_, xml)| xml.clone())
        });
        if let Some(additions) = drawing_rels {
            let original = archive.try_read_part(&rels_path)?;
            let combined = if let Some(original) = original { merge_relationship_additions(&original, &additions)? } else { additions };
            rels_cache.insert(rels_path, combined);
        }
    }

    for ((sheet, id), (text, units)) in text_edits {
        let Some(sheet_path) = sheet_paths.get(*sheet) else { continue };
        let rels_path = sheet_rels_path(sheet_path);
        let rels_xml = if let Some(xml) = rels_cache.get(&rels_path) {
            Some(xml.clone())
        } else {
            let xml = archive.try_read_part(&rels_path)?;
            if let Some(xml) = &xml { rels_cache.insert(rels_path.clone(), xml.clone()); }
            xml
        };
        let Some(rels_xml) = rels_xml else { continue };
        let rels = parse_relationships(&rels_xml)
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let Some(relation) = rels.values().find(|relation| relation.rel_type.ends_with("/threadedComment")) else { continue };
        let part = resolve_part_target(sheet_path, &relation.target);
        plan.comments.entry(part).or_default().text.insert(id.clone(), crate::string_value::encode(&units.clone().unwrap_or_else(|| text.encode_utf16().collect())));
    }

    for ((sheet, id), done) in resolved {
        let Some(sheet_path) = sheet_paths.get(*sheet) else { continue };
        let rels_path = sheet_rels_path(sheet_path);
        let rels_xml = if let Some(xml) = rels_cache.get(&rels_path) {
            Some(xml.clone())
        } else {
            let xml = archive.try_read_part(&rels_path)?;
            if let Some(xml) = &xml { rels_cache.insert(rels_path.clone(), xml.clone()); }
            xml
        };
        let Some(rels_xml) = rels_xml else { continue };
        let rels = parse_relationships(&rels_xml)
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let Some(relation) = rels.values().find(|relation| relation.rel_type.ends_with("/threadedComment")) else { continue };
        let part = resolve_part_target(sheet_path, &relation.target);
        plan.comments.entry(part).or_default().resolved.insert(id.clone(), *done);
    }

    for (sheet, ids) in deletions {
        let Some(sheet_path) = sheet_paths.get(*sheet) else { continue };
        let rels_path = sheet_rels_path(sheet_path);
        let rels_xml = if let Some(xml) = rels_cache.get(&rels_path) {
            Some(xml.clone())
        } else {
            let xml = archive.try_read_part(&rels_path)?;
            if let Some(xml) = &xml { rels_cache.insert(rels_path.clone(), xml.clone()); }
            xml
        };
        let Some(rels_xml) = rels_xml else { continue };
        let rels = parse_relationships(&rels_xml)
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let Some(relation) = rels.values().find(|relation| relation.rel_type.ends_with("/threadedComment")) else { continue };
        let part = resolve_part_target(sheet_path, &relation.target);
        plan.comments.entry(part).or_default().deletions.extend(ids.iter().cloned());
    }

    for (sheet, comments) in additions {
        let Some(sheet_path) = sheet_paths.get(*sheet) else { continue };
        let rels_path = sheet_rels_path(sheet_path);
        let rels_xml = if let Some(xml) = rels_cache.get(&rels_path) {
            Some(xml.clone())
        } else {
            let xml = archive.try_read_part(&rels_path)?;
            if let Some(xml) = &xml { rels_cache.insert(rels_path.clone(), xml.clone()); }
            xml
        };
        let existing = if let Some(rels_xml) = &rels_xml {
            let rels = parse_relationships(rels_xml)
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            rels.values().find(|relation| relation.rel_type.ends_with("/threadedComment"))
                .map(|relation| resolve_part_target(sheet_path, &relation.target))
        } else { None };
        let part = if let Some(part) = existing {
            part
        } else {
            let mut number = 1;
            let part = loop {
                let candidate = format!("xl/threadedComments/threadedComment{number}.xml");
                if !names.contains(&candidate) { break candidate; }
                number += 1;
            };
            let target = format!("../threadedComments/threadedComment{number}.xml");
            let updated = append_relationship(rels_xml.as_deref(), "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment", &target)?;
            rels_cache.insert(rels_path.clone(), updated.clone());
            if rels_xml.is_some() {
                plan.relationship_xml.insert(rels_path, updated);
            } else {
                plan.new_parts.push((rels_path, updated.into_bytes()));
            }
            plan.content_types.push((part.clone(), "application/vnd.ms-excel.threadedcomments+xml".into()));
            plan.new_parts.push((part.clone(), empty_threaded_comment_xml().as_bytes().to_vec()));
            part
        };
        plan.comments.entry(part).or_default().additions.extend(comments.iter().cloned());
    }

    if !additions.is_empty() {
        ensure_person_part(archive, &names, &mut plan)?;
    }
    Ok(plan)
}

fn ensure_person_part(
    archive: &mut OoxmlArchive,
    names: &std::collections::HashSet<String>,
    plan: &mut ThreadedCommentPackageEdits,
) -> Result<(), XlsxError> {
    let workbook_rels_path = "xl/_rels/workbook.xml.rels";
    let mut workbook_rels = archive.read_part(workbook_rels_path)?;
    let relationships = parse_relationships(&workbook_rels)
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    let existing = relationships.values()
        .find(|relation| relation.rel_type.ends_with("/person"))
        .map(|relation| resolve_part_target("xl/workbook.xml", &relation.target));
    let person_path = if let Some(path) = existing {
        path
    } else {
        let path = if names.contains("xl/persons/person.xml") {
            let mut index = 1;
            loop {
                let candidate = format!("xl/persons/person{index}.xml");
                if !names.contains(&candidate) { break candidate; }
                index += 1;
            }
        } else { "xl/persons/person.xml".to_string() };
        let target = path.strip_prefix("xl/").unwrap_or(&path);
        workbook_rels = append_relationship(Some(&workbook_rels), "http://schemas.microsoft.com/office/2017/10/relationships/person", target)?;
        plan.relationship_xml.insert(workbook_rels_path.into(), workbook_rels);
        plan.new_parts.push((path.clone(), empty_person_list_xml().as_bytes().to_vec()));
        plan.content_types.push((path.clone(), "application/vnd.ms-excel.person+xml".into()));
        path
    };
    let authors: std::collections::BTreeMap<String, (String, String, String)> = plan.comments.values()
        .flat_map(|edit| edit.additions.iter())
        .filter(|comment| !comment.person_id.is_empty())
        .map(|comment| (comment.person_id.clone(), (
            if comment.person_name.is_empty() { "User".to_string() } else { comment.person_name.clone() },
            if comment.person_user_id.is_empty() { "User".to_string() } else { comment.person_user_id.clone() },
            if comment.person_provider_id.is_empty() { "None".to_string() } else { comment.person_provider_id.clone() },
        )))
        .collect();
    if !authors.is_empty() {
        if let Some(xml) = archive.try_read_part(&person_path)? {
            let updated = authors.iter().try_fold(xml, |xml, (id, (name, user_id, provider_id))| append_person(&xml, id, name, user_id, provider_id))?;
            plan.person_xml.insert(person_path, updated);
        } else if let Some((_, bytes)) = plan.new_parts.iter_mut().find(|(path, _)| path == &person_path) {
            let mut xml = String::from_utf8_lossy(bytes).into_owned();
            for (id, (name, user_id, provider_id)) in &authors { xml = append_person(&xml, id, name, user_id, provider_id)?; }
            *bytes = xml.into_bytes();
        }
    }
    Ok(())
}

fn sheet_rels_path(sheet_path: &str) -> String {
    let (directory, file) = sheet_path.rsplit_once('/').unwrap_or(("", sheet_path));
    format!("{directory}/_rels/{file}.rels")
}

fn append_hyperlink_relationship(
    xml: Option<&str>,
    target: &str,
) -> Result<(String, String), XlsxError> {
    let mut current = xml.map(str::to_owned).unwrap_or_else(|| {
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>".to_string()
    });
    let parsed = parse_relationships(&current)
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    let mut index = 1;
    let id = loop {
        let candidate = format!("rId{index}");
        if !parsed.contains_key(&candidate) { break candidate; }
        index += 1;
    };
    let at = current.rfind("</Relationships>")
        .ok_or_else(|| XlsxError::InvalidData("relationships part has no closing element".into()))?;
    current.insert_str(at, &format!(
        "<Relationship Id=\"{id}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{}\" TargetMode=\"External\"/>",
        escape(target),
    ));
    Ok((current, id))
}

fn patch_hyperlinks_xml(
    xml: &str,
    links: &[Hyperlink],
    relationship_ids: &[Option<String>],
) -> Result<String, XlsxError> {
    if links.len() != relationship_ids.len() {
        return Err(XlsxError::InvalidData("hyperlink relationship ids do not match the link list".into()));
    }
    let mut output = xml.to_string();
    if let Some(start) = output.find("<hyperlinks") {
        let tag_end = output[start..].find('>').map(|offset| start + offset + 1)
            .ok_or_else(|| XlsxError::InvalidData("unterminated hyperlinks element".into()))?;
        let end = if output[start..tag_end].trim_end().ends_with("/>") {
            tag_end
        } else {
            output[tag_end..].find("</hyperlinks>").map(|offset| tag_end + offset + "</hyperlinks>".len())
                .ok_or_else(|| XlsxError::InvalidData("unterminated hyperlinks block".into()))?
        };
        output.replace_range(start..end, "");
    }
    if links.is_empty() { return Ok(output); }

    let mut entries = String::new();
    for (link, relationship_id) in links.iter().zip(relationship_ids) {
        let first = format!("{}{}", col_to_letter(link.cell.1), link.cell.0 + 1);
        let reference = link.end_cell.map_or(first.clone(), |(row, col)| {
            format!("{first}:{}{}", col_to_letter(col), row + 1)
        });
        let mut attrs = format!(" ref=\"{}\"", escape(&reference));
        if let Some(id) = relationship_id {
            attrs.push_str(&format!(" r:id=\"{}\"", escape(id)));
        }
        if link.internal {
            attrs.push_str(&format!(" location=\"{}\"", escape(&link.target)));
        } else if let Some(location) = link.sub_address.as_deref() {
            attrs.push_str(&format!(" location=\"{}\"", escape(location)));
        }
        if let Some(tooltip) = link.tooltip.as_deref() {
            attrs.push_str(&format!(" tooltip=\"{}\"", escape(tooltip)));
        }
        entries.push_str(&format!("<hyperlink{attrs}/>") );
    }
    let block = format!("<hyperlinks>{entries}</hyperlinks>");
    if relationship_ids.iter().any(Option::is_some) && !output.contains("xmlns:r=") {
        let root = output.find("<worksheet")
            .ok_or_else(|| XlsxError::InvalidData("worksheet root element is missing".into()))?;
        let end = output[root..].find('>').map(|offset| root + offset)
            .ok_or_else(|| XlsxError::InvalidData("worksheet root element is unterminated".into()))?;
        output.insert_str(end, " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
    }
    let before = ["<printOptions", "<pageMargins", "<pageSetup", "<headerFooter", "<drawing", "<legacyDrawing", "<picture", "<oleObjects", "<controls", "<tableParts", "<extLst", "</worksheet>"];
    let insertion = before.iter().filter_map(|tag| output.find(tag)).min()
        .ok_or_else(|| XlsxError::InvalidData("worksheet closing element is missing".into()))?;
    output.insert_str(insertion, &block);
    Ok(output)
}

fn append_relationship(xml: Option<&str>, rel_type: &str, target: &str) -> Result<String, XlsxError> {
    let mut current = xml.map(str::to_owned).unwrap_or_else(|| {
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>".to_string()
    });
    let parsed = parse_relationships(&current)
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    let mut index = 1;
    let id = loop {
        let candidate = format!("rId{index}");
        if !parsed.contains_key(&candidate) { break candidate; }
        index += 1;
    };
    let at = current.rfind("</Relationships>")
        .ok_or_else(|| XlsxError::InvalidData("relationships part has no closing element".into()))?;
    current.insert_str(at, &format!("<Relationship Id=\"{id}\" Type=\"{}\" Target=\"{}\"/>", escape(rel_type), escape(target)));
    Ok(current)
}

fn empty_threaded_comment_xml() -> &'static str {
    "<ThreadedComments xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></ThreadedComments>"
}

fn empty_person_list_xml() -> &'static str {
    "<personList xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></personList>"
}

fn append_person(xml: &str, id: &str, display_name: &str, user_id: &str, provider_id: &str) -> Result<String, XlsxError> {
    if xml.contains(&format!("id=\"{}\"", escape(id))) { return Ok(xml.to_string()); }
    let at = xml.rfind("</personList>")
        .ok_or_else(|| XlsxError::InvalidData("person list has no closing element".into()))?;
    let mut output = xml.to_string();
    output.insert_str(at, &format!("<person displayName=\"{}\" id=\"{}\" userId=\"{}\" providerId=\"{}\"/>", escape(display_name), escape(id), escape(user_id), escape(provider_id)));
    Ok(output)
}

fn append_threaded_comments(xml: &str, comments: &[crate::ir::ThreadedComment]) -> Result<String, XlsxError> {
    let at = xml.rfind("</ThreadedComments>")
        .ok_or_else(|| XlsxError::InvalidData("threaded comments part has no closing element".into()))?;
    let mut entries = String::new();
    for comment in comments {
        let (row, col) = comment.cell;
        let mut attributes = format!("ref=\"{}{}\" personId=\"{}\" id=\"{}\"", col_to_letter(col), row + 1, escape(&comment.person_id), escape(&comment.id));
        if !comment.date.is_empty() { attributes.push_str(&format!(" dT=\"{}\"", escape(&comment.date))); }
        if let Some(parent) = comment.parent_id.as_deref() { attributes.push_str(&format!(" parentId=\"{}\"", escape(parent))); }
        if comment.resolved { attributes.push_str(" done=\"1\""); }
        entries.push_str(&format!("<threadedComment {attributes}><text>{}</text></threadedComment>", escape(&crate::string_value::encode(&comment.utf16_text.clone().unwrap_or_else(|| comment.text.encode_utf16().collect())))));
    }
    let mut output = xml.to_string();
    output.insert_str(at, &entries);
    Ok(output)
}

/// Reconcile removals and range changes without rewriting differential styles.
/// Other edits are refused until they can preserve the rule's unmodelled XML.
fn patch_conditional_rules_xml(xml: &str, before: &[ConditionalRule], after: &[ConditionalRule], style_ids: &HashMap<(bool, i32), u32>) -> Result<String, XlsxError> {
    let fail = || XlsxError::InvalidData("cannot safely reconcile worksheet conditional formatting".into());
    if before.iter().any(|rule| rule.visual_extension_only) || after.iter().any(|rule| rule.visual_extension_only) {
        // An x14-only rule has no base cfRule: update or remove its extension.
        let mut extension_updates = BTreeMap::new();
        for original in before.iter().filter(|rule| rule.visual_extension_only) {
            let id = original.visual_extension_id.clone().ok_or_else(fail)?;
            let matches: Vec<_> = after.iter().filter(|rule| rule.visual_extension_only && rule.origin_priority == original.origin_priority).collect();
            if matches.is_empty() {
                if extension_updates.insert(id, None).is_some() { return Err(fail()); }
                continue;
            }
            let [updated] = matches.as_slice() else { return Err(fail()); };
            let mut expected = original.clone();
            expected.ranges = updated.ranges.clone();
            expected.priority = updated.priority;
            expected.visual_xml = updated.visual_xml.clone();
            expected.visual_extension_xml = updated.visual_extension_xml.clone();
            if expected != **updated { return Err(fail()); }
            if original.visual_xml != updated.visual_xml || original.visual_extension_xml != updated.visual_extension_xml {
                let source = original.visual_extension_xml.as_deref().ok_or_else(fail)?;
                let visual = updated.visual_xml.as_deref().ok_or_else(fail)?;
                if update_extended_icon_visual(source, visual).map_err(|_| fail())? != *updated.visual_extension_xml.as_ref().ok_or_else(fail)? { return Err(fail()); }
            }
            if original.ranges != updated.ranges || original.priority != updated.priority || original.visual_extension_xml != updated.visual_extension_xml {
                let mut payload = if original.visual_extension_xml != updated.visual_extension_xml { Some(updated.visual_extension_xml.clone().ok_or_else(fail)?) } else { None };
                if original.priority != updated.priority {
                    let raw = payload.as_deref().or(updated.visual_extension_xml.as_deref()).ok_or_else(fail)?;
                    payload = Some(set_extended_rule_priority(raw, &id, updated.priority).map_err(|_| fail())?);
                }
                let sqref = updated.ranges.iter().map(|&(r0,c0,r1,c1)| {
                    let first = format!("{}{}", col_to_letter(c0), r0 + 1);
                    if (r0,c0) == (r1,c1) { first } else { format!("{}:{}{}", first, col_to_letter(c1), r1 + 1) }
                }).collect::<Vec<_>>().join(" ");
                if extension_updates.insert(id, Some((sqref, payload))).is_some() { return Err(fail()); }
            }
        }
        if after.iter().filter(|rule| rule.visual_extension_only && rule.origin_priority.is_some()).any(|rule| !before.iter().any(|original| original.visual_extension_only && original.origin_priority == rule.origin_priority)) { return Err(fail()); }
        let mut priorities = std::collections::BTreeSet::new();
        if after.iter().any(|rule| rule.priority <= 0 || !priorities.insert(rule.priority)) { return Err(fail()); }
        let base_before: Vec<_> = before.iter().filter(|rule| !rule.visual_extension_only).cloned().collect();
        let base_after: Vec<_> = after.iter().filter(|rule| !rule.visual_extension_only).cloned().collect();
        let output = patch_conditional_rules_xml(xml, &base_before, &base_after, style_ids)?;
        let output = patch_extended_conditional_ranges(&output, &extension_updates)?;
        let added: Vec<_> = after.iter().filter(|rule| rule.visual_extension_only && rule.origin_priority.is_none()).cloned().collect();
        return append_conditional_rules_xml(&output, &added, &HashMap::new());
    }
    let mut priorities = BTreeMap::new();
    for rule in before {
        if rule.priority <= 0 || priorities.insert(rule.priority, rule).is_some() { return Err(fail()); }
    }
    let mut retained = BTreeMap::new();
    let mut added = Vec::new();
    let mut all_priorities = std::collections::BTreeSet::new();
    for rule in after {
        if rule.priority <= 0 || !all_priorities.insert(rule.priority) { return Err(fail()); }
        let Some(origin) = rule.origin_priority else { added.push(rule.clone()); continue; };
        let old = priorities.get(&origin).ok_or_else(fail)?;
        let mut expected = (*old).clone();
        let corner = |ranges: &[(u32,u32,u32,u32)]| Some((ranges.iter().map(|range| range.0).min()?, ranges.iter().map(|range| range.1).min()?));
        let old_corner = corner(&old.ranges).ok_or_else(fail)?;
        let new_corner = corner(&rule.ranges).ok_or_else(fail)?;
        let row_shift = i64::from(new_corner.0) - i64::from(old_corner.0);
        let col_shift = i64::from(new_corner.1) - i64::from(old_corner.1);
        expected.ranges = rule.ranges.clone();
        expected.priority = rule.priority;
        expected.stop_if_true = rule.stop_if_true;
        if old.style != rule.style {
            let old_style = old.style.as_ref().cloned().unwrap_or_default();
            let new_style = rule.style.as_ref().cloned().unwrap_or_default();
            if old_style.number_format != new_style.number_format || old_style.source_xml != new_style.source_xml { return Err(fail()); }
            if !style_ids.contains_key(&(true, origin)) { return Err(fail()); }
            expected.style = rule.style.clone();
        }
        if old.kind == "top10" && rule.kind == "top10" {
            expected.rank = rule.rank; expected.bottom = rule.bottom; expected.percent = rule.percent;
        }
        if matches!(old.kind.as_str(), "uniqueValues" | "duplicateValues") && matches!(rule.kind.as_str(), "uniqueValues" | "duplicateValues") {
            expected.kind = rule.kind.clone();
        }
        if old.kind == "aboveAverage" && rule.kind == "aboveAverage" {
            expected.above_average = rule.above_average; expected.equal_average = rule.equal_average; expected.standard_deviations = rule.standard_deviations;
        }
        if old.kind == rule.kind && matches!(old.kind.as_str(), "colorScale" | "dataBar" | "iconSet") {
            if let Some(raw) = &rule.visual_xml { if !valid_conditional_visual(raw, &rule.kind) { return Err(fail()); } }
            expected.visual_xml = rule.visual_xml.clone();
            if old.visual_extension_id == rule.visual_extension_id {
                if old.visual_extension_xml.is_some() && rule.visual_extension_xml.is_none() { return Err(fail()); }
                expected.visual_extension_xml = rule.visual_extension_xml.clone();
            } else if old.kind == "dataBar" && old.visual_extension_id.is_none() && old.visual_extension_xml.is_none() && rule.visual_extension_id.is_some() && rule.visual_extension_xml.is_some() {
                expected.visual_extension_id = rule.visual_extension_id.clone();
                expected.visual_extension_xml = rule.visual_extension_xml.clone();
            }
        }
        if matches!(old.kind.as_str(), "cellIs" | "expression") && matches!(rule.kind.as_str(), "cellIs" | "expression") {
            expected.kind = rule.kind.clone();
            expected.operator = rule.operator.clone();
            expected.formulas = rule.formulas.clone();
        } else if old.kind == rule.kind && matches!(old.kind.as_str(),
            "containsText" | "notContainsText" | "beginsWith" | "endsWith"
            | "containsBlanks" | "notContainsBlanks" | "containsErrors" | "notContainsErrors") {
            // These rules carry ordinary top-level formulas as well. Structural
            // edits can shift references without moving the rule's anchor.
            expected.formulas = rule.formulas.clone();
        } else {
            for formula in &mut expected.formulas {
                *formula = oxicells_calc::lexer::translate_validation_references(formula, row_shift, col_shift).map_err(|_| fail())?;
            }
        }
        if expected != *rule || retained.insert(origin, rule).is_some() { return Err(fail()); }
    }
    struct Block { start: usize, head_end: usize, other: bool, rules: Vec<(usize, usize, i32)> }
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut block: Option<Block> = None;
    let mut open_rule: Option<(usize, i32)> = None;
    let mut seen = std::collections::BTreeSet::new();
    let mut replacements: Vec<(usize, usize, String)> = Vec::new();
    let mut extension_ranges = BTreeMap::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|_| fail())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) if depth == 1 && local_name(element.name().as_ref()) == "conditionalFormatting" => {
                if block.is_some() { return Err(fail()); }
                block = Some(Block { start, head_end: end, other: false, rules: Vec::new() });
                depth += 1;
            }
            Event::Start(ref element) | Event::Empty(ref element) if depth == 2 && block.is_some() && local_name(element.name().as_ref()) == "cfRule" => {
                let priority = get_attr(element, "priority").and_then(|value| value.parse::<i32>().ok()).ok_or_else(fail)?;
                if !priorities.contains_key(&priority) || !seen.insert(priority) { return Err(fail()); }
                if matches!(event, Event::Empty(_)) {
                    block.as_mut().unwrap().rules.push((start, end, priority));
                } else {
                    open_rule = Some((start, priority));
                    depth += 1;
                }
            }
            Event::Start(_) if depth == 2 && block.is_some() => { block.as_mut().unwrap().other = true; depth += 1; }
            Event::Empty(_) if depth == 2 && block.is_some() => { block.as_mut().unwrap().other = true; }
            Event::Start(_) => depth += 1,
            Event::End(ref element) => {
                depth = depth.checked_sub(1).ok_or_else(fail)?;
                if depth == 2 && local_name(element.name().as_ref()) == "cfRule" {
                    let (rule_start, priority) = open_rule.take().ok_or_else(fail)?;
                    block.as_mut().ok_or_else(fail)?.rules.push((rule_start, end, priority));
                } else if depth == 1 && local_name(element.name().as_ref()) == "conditionalFormatting" {
                    let block = block.take().ok_or_else(fail)?;
                    let changed_ranges = block.rules.iter().any(|(_, _, priority)| retained.get(priority).is_some_and(|rule| rule.ranges != priorities[priority].ranges));
                    // Extended data bars store their range again in the sheet
                    // extension, linked to this base rule by an x14 id.
                    for (rule_start, rule_end, priority) in &block.rules {
                        let raw = &xml[*rule_start..*rule_end];
                        let updated = retained.get(priority);
                        if priorities[priority].visual_extension_id.is_none() && updated.is_some_and(|rule| rule.visual_extension_id.is_some()) { continue; }
                        let changed = updated.map_or(true, |rule| rule.ranges != priorities[priority].ranges || rule.visual_extension_xml != priorities[priority].visual_extension_xml);
                        if !changed || !raw.contains("<extLst") { continue; }
                        let id = raw.split_once("<x14:id>").and_then(|(_, rest)| rest.split_once("</x14:id>").map(|(id, _)| id)).ok_or_else(fail)?;
                        if !id.starts_with('{') || !id.ends_with('}') || extension_ranges.contains_key(id) { return Err(fail()); }
                        let change = if let Some(rule) = updated {
                            let sqref = rule.ranges.iter().map(|&(r0,c0,r1,c1)| {
                                let first = format!("{}{}", col_to_letter(c0), r0 + 1);
                                if (r0,c0) == (r1,c1) { first } else { format!("{}:{}{}", first, col_to_letter(c1), r1 + 1) }
                            }).collect::<Vec<_>>().join(" ");
                            let payload = if rule.visual_extension_xml != priorities[priority].visual_extension_xml {
                                rule.visual_extension_xml.as_deref().map(omit_hidden_databar_axis_color).transpose()?
                            } else { None };
                            Some((sqref, payload))
                        } else { None };
                        extension_ranges.insert(id.to_string(), change);
                    }
                    if changed_ranges {
                        if block.other { return Err(fail()); }
                        let mut replacement = String::new();
                        let header = &xml[block.start..block.head_end];
                        let key = header.find("sqref=").ok_or_else(fail)? + "sqref=".len();
                        let quote = header.as_bytes().get(key).copied().filter(|c| *c == b'\"' || *c == b'\'').ok_or_else(fail)?;
                        let quote_end = header[key + 1..].find(quote as char).map(|n| key + 1 + n).ok_or_else(fail)?;
                        for (rule_start, rule_end, priority) in block.rules {
                            let Some(rule) = retained.get(&priority) else { continue; };
                            let sqref = rule.ranges.iter().map(|&(r0,c0,r1,c1)| {
                                let first = format!("{}{}", col_to_letter(c0), r0 + 1);
                                if (r0,c0) == (r1,c1) { first } else { format!("{}:{}{}", first, col_to_letter(c1), r1 + 1) }
                            }).collect::<Vec<_>>().join(" ");
                            replacement.push_str(&header[..key + 1]); replacement.push_str(&escape(&sqref)); replacement.push_str(&header[quote_end..]);
                            replacement.push_str(&patch_conditional_rule(&xml[rule_start..rule_end], priorities[&priority], rule, style_ids.get(&(true, priority)).copied())?);
                            replacement.push_str(&xml[start..end]);
                        }
                        replacements.push((block.start, end, replacement));
                    } else {
                        let deleted: Vec<_> = block.rules.iter().filter(|(_, _, priority)| !retained.contains_key(priority)).collect();
                        if !block.rules.is_empty() && deleted.len() == block.rules.len() {
                            if block.other { return Err(fail()); }
                            replacements.push((block.start, end, String::new()));
                        } else {
                            for (rule_start, rule_end, _) in deleted { replacements.push((*rule_start, *rule_end, String::new())); }
                            for (rule_start, rule_end, priority) in &block.rules {
                                if let Some(rule) = retained.get(priority) {
                                    if *rule != priorities[priority] {
                                        replacements.push((*rule_start, *rule_end, patch_conditional_rule(&xml[*rule_start..*rule_end], priorities[priority], rule, style_ids.get(&(true, *priority)).copied())?));
                                    }
                                }
                            }
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if block.is_some() || open_rule.is_some() || seen.len() != before.len() { return Err(fail()); }
    let mut output = xml.to_string();
    replacements.sort_unstable_by(|a, b| b.0.cmp(&a.0));
    for (start, end, replacement) in replacements { output.replace_range(start..end, &replacement); }
    output = patch_extended_conditional_ranges(&output, &extension_ranges)?;
    let mut new_extensions = String::new();
    for (origin, rule) in &retained {
        let before = priorities[origin];
        if before.kind != "dataBar" || before.visual_extension_id.is_some() { continue; }
        let Some(id) = rule.visual_extension_id.as_deref() else { continue; };
        let raw = rule.visual_extension_xml.as_deref().ok_or_else(fail)?;
        if !id.starts_with('{') || !id.ends_with('}') || raw.matches(id).count() != 1 || xml.contains(id) || new_extensions.contains(id) { return Err(fail()); }
        let sqref = rule.ranges.iter().map(|&(r0,c0,r1,c1)| {
            let first = format!("{}{}", col_to_letter(c0), r0 + 1);
            if (r0,c0) == (r1,c1) { first } else { format!("{}:{}{}", first, col_to_letter(c1), r1 + 1) }
        }).collect::<Vec<_>>().join(" ");
        let mut ranges = BTreeMap::new(); ranges.insert(id.to_string(), Some((sqref, None)));
        new_extensions.push_str(&patch_extended_conditional_ranges(&omit_hidden_databar_axis_color(raw)?, &ranges)?);
    }
    output = append_extended_conditional_rules(&output, &new_extensions)?;
    let added_styles = style_ids.iter().filter_map(|(&(existing, priority), &id)| (!existing).then_some((priority, id))).collect();
    append_conditional_rules_xml(&output, &added, &added_styles)
}

/// Change the priority of one x14 rule without rewriting its payload.
fn set_extended_rule_priority(xml: &str, id: &str, priority: i32) -> Result<String, String> {
    if priority <= 0 { return Err("invalid extended conditional priority".into()); }
    let mut reader = Reader::from_str(xml);
    let mut replacement = None;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) | Event::Empty(ref element)
                if element.name().as_ref() == b"x14:cfRule" && get_attr(element, "id").as_deref() == Some(id) => {
                    if replacement.is_some() { return Err("duplicate extended conditional id".into()); }
                    let mut tag = element.to_owned(); tag.clear_attributes();
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| error.to_string())?;
                        if local_name(attribute.key.as_ref()) != "priority" { tag.push_attribute(attribute); }
                    }
                    let value = priority.to_string(); tag.push_attribute(("priority", value.as_str()));
                    let mut writer = Writer::new(Vec::new());
                    writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(tag) } else { Event::Start(tag) }).map_err(|error| error.to_string())?;
                    replacement = Some((start, end, String::from_utf8(writer.into_inner()).map_err(|error| error.to_string())?));
                }
            Event::Eof => break,
            _ => {}
        }
    }
    let (start, end, value) = replacement.ok_or_else(|| "missing extended conditional id".to_string())?;
    Ok(format!("{}{}{}", &xml[..start], value, &xml[end..]))
}

/// Patch a linked x14 conditional formatting block without rewriting its
/// unmodelled settings. Ambiguous or missing links must fail the save.
fn patch_extended_conditional_ranges(xml: &str, updates: &BTreeMap<String, Option<(String, Option<String>)>>) -> Result<String, XlsxError> {
    if updates.is_empty() { return Ok(xml.to_string()); }
    let fail = || XlsxError::InvalidData("cannot safely reconcile extended conditional formatting".into());
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut block_depth = None;
    let mut block_start = 0usize;
    let mut ids = Vec::new();
    let mut range_start = None;
    let mut range = None;
    let mut replacements = Vec::new();
    let mut found = std::collections::BTreeSet::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|_| fail())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref e) => {
                let name = e.name();
                if name.as_ref() == b"x14:conditionalFormatting" {
                    if block_depth.is_some() { return Err(fail()); }
                    block_depth = Some(depth); block_start = start; ids.clear(); range = None;
                } else if block_depth.is_some() && name.as_ref() == b"x14:cfRule" {
                    ids.push(get_attr(e, "id").ok_or_else(fail)?);
                } else if block_depth.is_some() && name.as_ref() == b"xm:sqref" {
                    if range_start.replace(end).is_some() || range.is_some() { return Err(fail()); }
                } else if range_start.is_some() { return Err(fail()); }
                depth += 1;
            }
            Event::Empty(ref e) => {
                if range_start.is_some() { return Err(fail()); }
                if block_depth.is_some() && e.name().as_ref() == b"x14:cfRule" { ids.push(get_attr(e, "id").ok_or_else(fail)?); }
            }
            Event::End(ref e) => {
                depth = depth.checked_sub(1).ok_or_else(fail)?;
                if e.name().as_ref() == b"xm:sqref" && block_depth.is_some() {
                    range = Some((range_start.take().ok_or_else(fail)?, start));
                } else if e.name().as_ref() == b"x14:conditionalFormatting" && block_depth == Some(depth) {
                    let matching: Vec<_> = ids.iter().filter(|id| updates.contains_key(*id)).collect();
                    if !matching.is_empty() {
                        if ids.len() != 1 || matching.len() != 1 || !found.insert(matching[0].clone()) { return Err(fail()); }
                        if let Some((sqref, payload)) = &updates[matching[0]] {
                            if let Some(payload) = payload {
                                let mut nested = BTreeMap::new();
                                nested.insert(matching[0].clone(), Some((sqref.clone(), None)));
                                let rewritten = patch_extended_conditional_ranges(payload, &nested)?;
                                replacements.push((block_start, end, rewritten));
                            } else {
                                let (first, last) = range.ok_or_else(fail)?;
                                replacements.push((first, last, escape(sqref)));
                            }
                        } else {
                            replacements.push((block_start, end, String::new()));
                        }
                    }
                    block_depth = None;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if block_depth.is_some() || range_start.is_some() || found.len() != updates.len() { return Err(fail()); }
    let mut output = xml.to_string();
    replacements.sort_unstable_by(|a, b| b.0.cmp(&a.0));
    for (start, end, replacement) in replacements { output.replace_range(start..end, &replacement); }
    if updates.values().any(Option::is_none) {
        output = prune_empty_conditional_extension(&output, "x14:conditionalFormattings", None, None)?;
        output = prune_empty_conditional_extension(&output, "ext", Some("{78C0D931-6437-407D-A8EE-F0AAD7539E65}"), None)?;
        output = prune_empty_conditional_extension(&output, "extLst", None, Some(1))?;
    }
    Ok(output)
}

fn prune_empty_conditional_extension(xml: &str, tag: &str, uri: Option<&str>, required_depth: Option<usize>) -> Result<String, XlsxError> {
    let fail = || XlsxError::InvalidData("cannot safely prune conditional formatting extension".into());
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut candidates = Vec::new();
    let mut replacements = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|_| fail())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref e) => {
                if e.name().as_ref() == tag.as_bytes() && required_depth.map_or(true, |wanted| depth == wanted)
                    && uri.map_or(true, |wanted| get_attr(e, "uri").is_some_and(|actual| actual.eq_ignore_ascii_case(wanted))) {
                    candidates.push((depth, start, end));
                }
                depth += 1;
            }
            Event::End(ref e) => {
                depth = depth.checked_sub(1).ok_or_else(fail)?;
                if e.name().as_ref() == tag.as_bytes() && candidates.last().is_some_and(|candidate| candidate.0 == depth) {
                    let (_, first, content) = candidates.pop().ok_or_else(fail)?;
                    if xml[content..start].trim().is_empty() { replacements.push((first, end)); }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if !candidates.is_empty() { return Err(fail()); }
    let mut output = xml.to_string();
    replacements.sort_unstable_by(|a, b| b.0.cmp(&a.0));
    for (start, end) in replacements { output.replace_range(start..end, ""); }
    Ok(output)
}

fn append_conditional_rules_xml(xml: &str, rules: &[ConditionalRule], style_ids: &HashMap<i32, u32>) -> Result<String, XlsxError> {
    if rules.is_empty() { return Ok(xml.to_string()); }
    let fail = || XlsxError::InvalidData("cannot safely write conditional rule".into());
    let mut entries = String::new();
    let mut extended_entries = String::new();
    let mut seen = std::collections::BTreeSet::new();
    for rule in rules {
        if !matches!(rule.kind.as_str(), "cellIs" | "expression" | "colorScale" | "dataBar" | "iconSet" | "top10" | "uniqueValues" | "duplicateValues" | "aboveAverage" | "containsText" | "notContainsText" | "beginsWith" | "endsWith" | "containsBlanks" | "notContainsBlanks" | "containsErrors" | "notContainsErrors")
            || rule.priority <= 0 || !seen.insert(rule.priority) || rule.ranges.is_empty() { return Err(fail()); }
        let sqref = rule.ranges.iter().map(|&(r0,c0,r1,c1)| {
            let first = format!("{}{}", col_to_letter(c0), r0 + 1);
            if (r0,c0) == (r1,c1) { first } else { format!("{}:{}{}", first, col_to_letter(c1), r1 + 1) }
        }).collect::<Vec<_>>().join(" ");
        if rule.visual_extension_only {
            let old_id = rule.visual_extension_id.as_deref().filter(|id| id.starts_with('{') && id.ends_with('}')).ok_or_else(fail)?;
            let raw = rule.visual_extension_xml.as_deref().ok_or_else(fail)?;
            if rule.kind != "iconSet" || raw.matches(old_id).count() != 1 { return Err(fail()); }
            use std::hash::{Hash, Hasher};
            let mut hasher = std::collections::hash_map::DefaultHasher::new();
            xml.hash(&mut hasher); old_id.hash(&mut hasher); rule.priority.hash(&mut hasher); sqref.hash(&mut hasher);
            let mut candidate = hasher.finish();
            let new_id = loop {
                let id = format!("{{{:08X}-{:04X}-4{:03X}-8{:03X}-{:012X}}}", (candidate >> 32) as u32, (candidate >> 16) as u16, (candidate >> 4) & 0xfff, candidate & 0xfff, candidate & 0xffffffffffff);
                if !xml.contains(&id) && !extended_entries.contains(&id) { break id; }
                candidate = candidate.wrapping_add(1);
            };
            let changed = raw.replace(old_id, &new_id);
            let changed = set_extended_rule_priority(&changed, &new_id, rule.priority).map_err(|_| fail())?;
            let mut ranges = BTreeMap::new(); ranges.insert(new_id, Some((sqref, None)));
            extended_entries.push_str(&patch_extended_conditional_ranges(&changed, &ranges)?);
            continue;
        }
        let extension_id = match (&rule.visual_extension_id, &rule.visual_extension_xml) {
            (None, None) => None,
            (Some(old_id), Some(raw)) if rule.kind == "dataBar" && old_id.starts_with('{') && old_id.ends_with('}') => {
                use std::hash::{Hash, Hasher};
                let mut hasher = std::collections::hash_map::DefaultHasher::new();
                xml.hash(&mut hasher); old_id.hash(&mut hasher); rule.priority.hash(&mut hasher); sqref.hash(&mut hasher);
                let mut candidate = hasher.finish();
                let new_id = loop {
                    let id = format!("{{{:08X}-{:04X}-4{:03X}-8{:03X}-{:012X}}}", (candidate >> 32) as u32, (candidate >> 16) as u16, (candidate >> 4) & 0xfff, candidate & 0xfff, candidate & 0xffffffffffff);
                    if !xml.contains(&id) && !extended_entries.contains(&id) { break id; }
                    candidate = candidate.wrapping_add(1);
                };
                if raw.matches(old_id).count() != 1 { return Err(fail()); }
                let changed = omit_hidden_databar_axis_color(raw)?.replace(old_id, &new_id);
                let mut ranges = BTreeMap::new(); ranges.insert(new_id.clone(), Some((sqref.clone(), None)));
                let changed = patch_extended_conditional_ranges(&changed, &ranges)?;
                extended_entries.push_str(&changed);
                Some(new_id)
            }
            _ => return Err(fail()),
        };
        entries.push_str(&format!("<conditionalFormatting sqref=\"{}\"><cfRule type=\"{}\" priority=\"{}\" stopIfTrue=\"{}\"", escape(&sqref), escape(&rule.kind), rule.priority, u8::from(rule.stop_if_true)));
        if let Some(style) = &rule.style {
            let id = style_ids.get(&rule.priority).ok_or_else(fail)?;
            let _ = style;
            entries.push_str(&format!(" dxfId=\"{id}\""));
        }
        if let Some(operator) = &rule.operator { entries.push_str(&format!(" operator=\"{}\"", escape(operator))); }
        if let Some(text) = &rule.text { entries.push_str(&format!(" text=\"{}\"", escape(text))); }
        if rule.kind == "top10" {
            entries.push_str(&format!(" rank=\"{}\" bottom=\"{}\" percent=\"{}\"", rule.rank.unwrap_or(10), u8::from(rule.bottom), u8::from(rule.percent)));
        }
        if rule.kind == "aboveAverage" {
            entries.push_str(&format!(" aboveAverage=\"{}\" equalAverage=\"{}\"", u8::from(rule.above_average), u8::from(rule.equal_average)));
            if let Some(count) = rule.standard_deviations { entries.push_str(&format!(" stdDev=\"{count}\"")); }
        }
        entries.push('>');
        if let Some(raw) = &rule.visual_xml {
            if !matches!(rule.kind.as_str(), "colorScale" | "dataBar" | "iconSet") || !valid_conditional_visual(raw, &rule.kind) { return Err(fail()); }
            entries.push_str(raw);
        } else if matches!(rule.kind.as_str(), "colorScale" | "dataBar" | "iconSet") {
            entries.push_str(default_conditional_visual(&rule.kind, rule.color_scale_points).ok_or_else(fail)?);
        }
        for formula in &rule.formulas { entries.push_str(&format!("<formula>{}</formula>", escape(formula))); }
        if let Some(id) = extension_id {
            entries.push_str(&format!("<extLst><ext uri=\"{{B025F937-C7B1-47D3-B67F-A62EFF666E3E}}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:id>{}</x14:id></ext></extLst>", escape(&id)));
        }
        entries.push_str("</cfRule></conditionalFormatting>");
    }
    // Only direct worksheet children are legal insertion points. An extLst
    // inside a data bar is part of its rule, not the worksheet's extLst.
    let later = ["dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "drawing", "legacyDrawing", "tableParts", "extLst"];
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut insertion = None;
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event().map_err(|_| fail())? {
            Event::Start(ref e) => {
                if depth == 1 && later.contains(&local_name(e.name().as_ref()).as_str()) { insertion = Some(start); break; }
                depth += 1;
            }
            Event::Empty(ref e) => {
                if depth == 1 && later.contains(&local_name(e.name().as_ref()).as_str()) { insertion = Some(start); break; }
            }
            Event::End(ref e) => {
                depth = depth.checked_sub(1).ok_or_else(fail)?;
                if depth == 0 && local_name(e.name().as_ref()) == "worksheet" { insertion = Some(start); break; }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let insertion = insertion.ok_or_else(fail)?;
    let output = format!("{}{entries}{}", &xml[..insertion], &xml[insertion..]);
    append_extended_conditional_rules(&output, &extended_entries)
}

fn append_extended_conditional_rules(xml: &str, entries: &str) -> Result<String, XlsxError> {
    if entries.is_empty() { return Ok(xml.to_string()); }
    let fail = || XlsxError::InvalidData("cannot safely append extended conditional formatting".into());
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut container = None;
    let mut close = None;
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event().map_err(|_| fail())? {
            Event::Start(ref e) => {
                if e.name().as_ref() == b"x14:conditionalFormattings" {
                    if container.is_some() || close.is_some() { return Err(fail()); }
                    container = Some(depth);
                }
                depth += 1;
            }
            Event::End(ref e) => {
                depth = depth.checked_sub(1).ok_or_else(fail)?;
                if e.name().as_ref() == b"x14:conditionalFormattings" && container == Some(depth) { close = Some(start); container = None; }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if let Some(at) = close { return Ok(format!("{}{entries}{}", &xml[..at], &xml[at..])); }
    if container.is_some() { return Err(fail()); }
    let extension = format!("<ext uri=\"{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:conditionalFormattings>{entries}</x14:conditionalFormattings></ext>");
    let (child, insert, empty, _) = direct_xml_child(xml, "extLst")?;
    if empty { return Err(fail()); }
    if let Some((start, end)) = child {
        let raw = &xml[start..end];
        let updated = if raw.ends_with("/>") { format!("{}>{extension}</extLst>", &raw[..raw.len() - 2]) }
            else { append_xml_child(raw, "extLst", &extension)? };
        Ok(format!("{}{updated}{}", &xml[..start], &xml[end..]))
    } else { Ok(format!("{}<extLst>{extension}</extLst>{}", &xml[..insert], &xml[insert..])) }
}

/// Rewrite formula-valued thresholds without disturbing visual settings.
pub fn rewrite_conditional_visual_formulas(xml: &str, rewrite: impl Fn(&str) -> Result<String, String>) -> Result<String, String> {
    let mut reader = Reader::from_str(xml);
    let mut replacements = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) | Event::Empty(ref element)
                if local_name(element.name().as_ref()) == "cfvo" && matches!(get_attr(element, "type").as_deref(), Some("num" | "percent" | "percentile" | "formula")) => {
                // Excel writes a threshold without `val` after assigning VBA Null.
                // There is no reference to rewrite in that case.
                let Some(formula) = get_attr(element, "val") else { continue; };
                let updated = rewrite(&formula)?;
                if updated == formula { continue; }
                let mut tag = element.to_owned(); tag.clear_attributes();
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| error.to_string())?;
                    if attribute.key.as_ref() == b"val" { tag.push_attribute(("val", updated.as_str())); }
                    else { tag.push_attribute(attribute); }
                }
                let mut writer = Writer::new(Vec::new());
                writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(tag) } else { Event::Start(tag) }).map_err(|error| error.to_string())?;
                replacements.push((start, end, String::from_utf8(writer.into_inner()).map_err(|error| error.to_string())?));
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let mut output = xml.to_string();
    for (start, end, value) in replacements.into_iter().rev() { output.replace_range(start..end, &value); }
    Ok(output)
}

/// Rewrite formula thresholds in the linked x14 data-bar payload.
pub fn rewrite_extended_conditional_formulas(xml: &str, rewrite: impl Fn(&str) -> Result<String, String>) -> Result<String, String> {
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut formula_cfvo = None;
    let mut formula_text = None;
    let mut replacements = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref e) => {
                if e.name().as_ref() == b"x14:cfvo" && matches!(get_attr(e, "type").as_deref(), Some("num" | "percent" | "percentile" | "formula")) {
                    if formula_cfvo.replace(depth).is_some() { return Err("nested conditional thresholds".into()); }
                } else if formula_cfvo.is_some() && e.name().as_ref() == b"xm:f" {
                    if formula_text.replace((depth, end)).is_some() { return Err("nested threshold formula".into()); }
                } else if formula_text.is_some() { return Err("unsupported threshold formula XML".into()); }
                depth += 1;
            }
            Event::Empty(_) if formula_text.is_some() => return Err("missing conditional threshold formula".into()),
            Event::End(ref e) => {
                depth = depth.checked_sub(1).ok_or_else(|| "invalid threshold XML".to_string())?;
                if e.name().as_ref() == b"xm:f" && formula_text.is_some_and(|(held, _)| held == depth) {
                    let (_, first) = formula_text.take().unwrap();
                    let formula = quick_xml::escape::unescape(&xml[first..start]).map_err(|error| error.to_string())?;
                    let updated = rewrite(&formula)?;
                    if updated != formula { replacements.push((first, start, escape(&updated))); }
                } else if e.name().as_ref() == b"x14:cfvo" && formula_cfvo == Some(depth) {
                    formula_cfvo = None;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if depth != 0 || formula_cfvo.is_some() || formula_text.is_some() { return Err("incomplete threshold XML".into()); }
    let mut output = xml.to_string();
    for (start, end, value) in replacements.into_iter().rev() { output.replace_range(start..end, &value); }
    Ok(output)
}

pub fn default_conditional_visual(kind: &str, points: Option<u8>) -> Option<&'static str> {
    match (kind, points) {
        ("colorScale", Some(2)) => Some(r#"<colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF7128"/><color rgb="FFFFEF9C"/></colorScale>"#),
        ("colorScale", Some(3)) => Some(r#"<colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale>"#),
        ("dataBar", _) => Some(r#"<dataBar minLength="0" maxLength="100"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>"#),
        ("iconSet", _) => Some(r#"<iconSet><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet>"#),
        _ => None,
    }
}

/// Set a length on a base data bar or its linked x14 data bar.
pub fn set_databar_length(xml: &str, extended: bool, name: &str, value: u8) -> Result<String, String> {
    if !matches!(name, "minLength" | "maxLength") || value > 100 || (!extended && !valid_conditional_visual(xml, "dataBar")) { return Err("invalid data bar length".into()); }
    let mut reader = Reader::from_str(xml);
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) if if extended { element.name().as_ref() == b"x14:dataBar" } else { local_name(element.name().as_ref()) == "dataBar" } => {
                let mut tag = element.to_owned(); tag.clear_attributes();
                for attr in element.attributes() { let attr = attr.map_err(|error| error.to_string())?; if local_name(attr.key.as_ref()) != name { tag.push_attribute(attr); } }
                let value = value.to_string(); tag.push_attribute((name, value.as_str()));
                let mut writer = Writer::new(Vec::new()); writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(tag) } else { Event::Start(tag) }).map_err(|error| error.to_string())?;
                return Ok(format!("{}{}{}", &xml[..start], String::from_utf8(writer.into_inner()).map_err(|error| error.to_string())?, &xml[end..]));
            }
            Event::Eof => return Err("missing data bar settings".into()),
            _ => {}
        }
    }
}

/// Set an x14 data-bar display attribute, preserving thresholds and colors.
pub fn set_extended_databar_display(xml: &str, name: &str, value: Option<&str>, ensure_axis_color: bool) -> Result<String, String> {
    if !matches!(name, "direction" | "gradient" | "axisPosition" | "border" | "negativeBarColorSameAsPositive" | "negativeBarBorderColorSameAsPositive") { return Err("invalid data bar display setting".into()); }
    let mut reader = Reader::from_str(xml); let mut depth = 0usize; let mut bar = None;
    let mut head = None; let mut close = None; let mut has_axis = false;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) if element.name().as_ref() == b"x14:dataBar" => {
                if bar.replace(depth).is_some() || head.is_some() { return Err("multiple extended data bars".into()); }
                let mut tag = element.to_owned(); tag.clear_attributes();
                for attr in element.attributes() { let attr = attr.map_err(|error| error.to_string())?; if local_name(attr.key.as_ref()) != name { tag.push_attribute(attr); } }
                if let Some(value) = value { tag.push_attribute((name, value)); }
                let mut writer = Writer::new(Vec::new()); writer.write_event(Event::Start(tag)).map_err(|error| error.to_string())?;
                head = Some((start, end, String::from_utf8(writer.into_inner()).map_err(|error| error.to_string())?)); depth += 1;
            }
            Event::Start(ref element) if bar.is_some_and(|n| depth == n + 1) && element.name().as_ref() == b"x14:axisColor" => { has_axis = true; depth += 1; }
            Event::Empty(ref element) if bar.is_some_and(|n| depth == n + 1) && element.name().as_ref() == b"x14:axisColor" => has_axis = true,
            Event::Start(_) => depth += 1,
            Event::End(ref element) => {
                depth = depth.checked_sub(1).ok_or_else(|| "invalid extended data bar".to_string())?;
                if element.name().as_ref() == b"x14:dataBar" && bar == Some(depth) { close = Some(start); bar = None; }
            }
            Event::Eof => break, _ => {}
        }
    }
    let (start, end, tag) = head.ok_or_else(|| "missing extended data bar".to_string())?;
    let close = close.ok_or_else(|| "invalid extended data bar".to_string())?;
    let mut output = xml.to_string();
    if ensure_axis_color && !has_axis { output.insert_str(close, "<x14:axisColor rgb=\"FF000000\"/>"); }
    output.replace_range(start..end, &tag);
    Ok(output)
}

/// Set the x14 axis color, adding it after the other data-bar children if absent.
pub fn set_extended_databar_axis_color(xml: &str, name: &str, value: &str) -> Result<String, String> {
    set_extended_databar_child_color(xml, "axisColor", name, value)
}

/// Set an x14 data-bar color without changing the remaining extension data.
pub fn set_extended_databar_child_color(xml: &str, child: &str, name: &str, value: &str) -> Result<String, String> {
    if !matches!(child, "axisColor" | "borderColor" | "negativeFillColor" | "negativeBorderColor") { return Err("invalid data bar color".into()); }
    if !matches!(name, "rgb" | "theme" | "tint" | "auto") { return Err("invalid data bar color".into()); }
    let mut reader = Reader::from_str(xml); let mut depth = 0usize; let mut bar = None;
    let mut found = None; let mut close = None; let mut following = None;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref e) if e.name().as_ref() == b"x14:dataBar" => { if bar.replace(depth).is_some() { return Err("multiple data bars".into()); } depth += 1; }
            Event::Empty(ref e) if bar.is_some_and(|n| depth == n + 1) && local_name(e.name().as_ref()) == child => {
                if found.is_some() { return Err("multiple axis colors".into()); } found = Some((start, end, e.to_owned()));
            }
            Event::Empty(ref e) if bar.is_some_and(|n| depth == n + 1) && match child { "borderColor" => ["negativeFillColor", "negativeBorderColor", "axisColor"].contains(&local_name(e.name().as_ref()).as_str()), "negativeFillColor" => ["negativeBorderColor", "axisColor"].contains(&local_name(e.name().as_ref()).as_str()), "negativeBorderColor" => local_name(e.name().as_ref()) == "axisColor", _ => false } => { if following.is_none() { following = Some(start); } }
            Event::Start(_) => depth += 1,
            Event::End(ref e) => { depth = depth.checked_sub(1).ok_or_else(|| "invalid data bar".to_string())?; if e.name().as_ref() == b"x14:dataBar" && bar == Some(depth) { close = Some(start); bar = None; } }
            Event::Eof => break, _ => {}
        }
    }
    let (start, end, element) = if let Some(found) = found { found } else { let at = following.or(close).ok_or_else(|| "missing data bar".to_string())?; (at, at, quick_xml::events::BytesStart::new(format!("x14:{child}")).to_owned()) };
    let mut tag = element.to_owned(); tag.clear_attributes();
    for attr in element.attributes() {
        let attr = attr.map_err(|error| error.to_string())?; let key = local_name(attr.key.as_ref());
        if key != name && !(name != "tint" && ["rgb", "theme", "indexed", "auto", "tint"].contains(&key.as_str())) { tag.push_attribute(attr); }
    }
    if name == "tint" && start == end { tag.push_attribute(("rgb", "FF000000")); }
    tag.push_attribute((name, value));
    let mut writer = Writer::new(Vec::new()); writer.write_event(Event::Empty(tag)).map_err(|error| error.to_string())?;
    Ok(format!("{}{}{}", &xml[..start], String::from_utf8(writer.into_inner()).map_err(|error| error.to_string())?, &xml[end..]))
}

/// Drop an explicitly independent x14 color when the matching color is linked.
pub fn remove_extended_databar_child_color(xml: &str, child: &str) -> Result<String, String> {
    if !matches!(child, "negativeFillColor" | "negativeBorderColor") { return Err("invalid data bar color".into()); }
    let mut reader = Reader::from_str(xml); let mut depth = 0usize; let mut bar = None;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref e) if e.name().as_ref() == b"x14:dataBar" => { bar = Some(depth); depth += 1; }
            Event::Empty(ref e) if bar.is_some_and(|n| depth == n + 1) && local_name(e.name().as_ref()) == child => return Ok(format!("{}{}", &xml[..start], &xml[end..])),
            Event::Start(_) => depth += 1,
            Event::End(ref e) => { depth = depth.checked_sub(1).ok_or_else(|| "invalid data bar".to_string())?; if e.name().as_ref() == b"x14:dataBar" && bar == Some(depth) { bar = None; } }
            Event::Eof => return Ok(xml.to_string()), _ => {}
        }
    }
}

fn omit_hidden_databar_axis_color(xml: &str) -> Result<String, XlsxError> {
    let fail = || XlsxError::InvalidData("invalid extended data bar axis".into());
    let mut reader = Reader::from_str(xml); let mut depth = 0usize; let mut bar = None; let mut hidden = false; let mut no_border = false; let mut same_fill = false; let mut same_border = false; let mut remove = Vec::new(); let mut head = None;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|_| fail())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref e) if e.name().as_ref() == b"x14:dataBar" => {
                bar = Some(depth); hidden = get_attr(e, "axisPosition").as_deref() == Some("none"); no_border = !matches!(get_attr(e, "border").as_deref(), Some("1" | "true")); same_fill = matches!(get_attr(e, "negativeBarColorSameAsPositive").as_deref(), Some("1" | "true")); same_border = matches!(get_attr(e, "negativeBarBorderColorSameAsPositive").as_deref(), Some("1" | "true"));
                if no_border && get_attr(e, "negativeBarBorderColorSameAsPositive").is_some() {
                    let mut tag=e.to_owned(); tag.clear_attributes();
                    for attr in e.attributes() { let attr=attr.map_err(|_| fail())?; if local_name(attr.key.as_ref()) != "negativeBarBorderColorSameAsPositive" { tag.push_attribute(attr); } }
                    let mut writer=Writer::new(Vec::new()); writer.write_event(Event::Start(tag)).map_err(|_| fail())?;
                    head=Some((start,end,String::from_utf8(writer.into_inner()).map_err(|_| fail())?));
                }
                depth += 1;
            }
            Event::Empty(ref e) if bar.is_some_and(|n| depth == n + 1) && ((hidden && e.name().as_ref() == b"x14:axisColor") || (no_border && matches!(e.name().as_ref(), b"x14:borderColor" | b"x14:negativeBorderColor")) || (same_fill && e.name().as_ref() == b"x14:negativeFillColor") || (same_border && e.name().as_ref() == b"x14:negativeBorderColor")) => remove.push((start,end)),
            Event::Start(_) => depth += 1,
            Event::End(ref e) => { depth = depth.checked_sub(1).ok_or_else(fail)?; if e.name().as_ref() == b"x14:dataBar" && bar == Some(depth) { bar = None; } }
            Event::Eof => break, _ => {}
        }
    }
    let mut output = xml.to_string(); for (start,end) in remove.into_iter().rev() { output.replace_range(start..end, ""); } if let Some((start,end,tag))=head { output.replace_range(start..end, &tag); } Ok(output)
}

/// Change a visual flag while preserving the other attributes and children.
pub fn set_conditional_visual_boolean(xml: &str, kind: &str, name: &str, value: bool) -> Result<String, String> {
    let default = match (kind, name) { ("dataBar" | "iconSet", "showValue") => true, ("iconSet", "reverse") => false, _ => return Err("invalid conditional visual flag".into()) };
    if !valid_conditional_visual(xml, kind) { return Err("invalid conditional visual".into()); }
    let mut reader = Reader::from_str(xml);
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) if local_name(element.name().as_ref()) == kind => {
                let mut tag = element.to_owned();
                tag.clear_attributes();
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| error.to_string())?;
                    if local_name(attribute.key.as_ref()) != name { tag.push_attribute(attribute); }
                }
                if value != default { tag.push_attribute((name, if value { "1" } else { "0" })); }
                let mut writer = Writer::new(Vec::new());
                writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(tag) } else { Event::Start(tag) }).map_err(|error| error.to_string())?;
                let changed = String::from_utf8(writer.into_inner()).map_err(|error| error.to_string())?;
                return Ok(format!("{}{}{}", &xml[..start], changed, &xml[end..]));
            }
            Event::Eof => return Err("missing conditional visual".into()),
            _ => {}
        }
    }
}

/// Change one icon threshold without disturbing its other attributes.
pub fn set_icon_criterion(xml: &str, index: usize, kind: Option<&str>, value: Option<Option<&str>>, greater_or_equal: Option<bool>) -> Result<String, String> {
    set_conditional_threshold(xml, "iconSet", index, kind, value, greater_or_equal)
}

/// Change a threshold in a base conditional visual, preserving other fields.
pub fn set_conditional_threshold(xml: &str, visual: &str, index: usize, kind: Option<&str>, value: Option<Option<&str>>, greater_or_equal: Option<bool>) -> Result<String, String> {
    if !valid_conditional_visual(xml, visual) { return Err("invalid conditional visual".into()); }
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut seen = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) if depth == 1 && local_name(element.name().as_ref()) == "cfvo" => {
                if seen == index {
                    let mut tag = element.to_owned(); tag.clear_attributes();
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| error.to_string())?;
                        let key = local_name(attribute.key.as_ref());
                        if (key == "type" && kind.is_some()) || (key == "val" && value.is_some()) || (key == "gte" && greater_or_equal.is_some()) { continue; }
                        tag.push_attribute(attribute);
                    }
                    if let Some(kind) = kind { tag.push_attribute(("type", kind)); }
                    if let Some(Some(value)) = value { tag.push_attribute(("val", value)); }
                    if let Some(false) = greater_or_equal { tag.push_attribute(("gte", "0")); }
                    let mut writer = Writer::new(Vec::new());
                    writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(tag) } else { Event::Start(tag) }).map_err(|error| error.to_string())?;
                    let changed = String::from_utf8(writer.into_inner()).map_err(|error| error.to_string())?;
                    return Ok(format!("{}{}{}", &xml[..start], changed, &xml[end..]));
                }
                seen += 1;
                if matches!(event, Event::Start(_)) { depth += 1; }
            }
            Event::Start(_) => depth += 1,
            Event::End(_) => depth = depth.checked_sub(1).ok_or_else(|| "invalid icon set".to_string())?,
            Event::Eof => return Err("icon criterion is out of range".into()),
            _ => {}
        }
    }
}

/// Synchronize one linked x14 data-bar threshold with its base rule.
pub fn set_extended_databar_threshold(xml: &str, index: usize, kind: &str, value: Option<&str>) -> Result<String, String> {
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut bar = None;
    let mut current: Option<(usize, usize, quick_xml::events::BytesStart<'static>)> = None;
    let mut seen = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let end = reader.buffer_position() as usize;
        let mut replacement = None;
        match event {
            Event::Start(ref e) if e.name().as_ref() == b"x14:dataBar" => { if bar.replace(depth).is_some() { return Err("multiple extended data bars".into()); } depth += 1; }
            Event::Start(ref e) if bar.is_some_and(|n| depth == n + 1) && e.name().as_ref() == b"x14:cfvo" => {
                if seen == index { current = Some((depth, start, e.to_owned())); } seen += 1; depth += 1;
            }
            Event::Empty(ref e) if bar.is_some_and(|n| depth == n + 1) && e.name().as_ref() == b"x14:cfvo" => {
                if seen == index { replacement = Some((start, end, e.to_owned())); } seen += 1;
            }
            Event::Start(ref e) if current.is_some() => {
                if e.name().as_ref() != b"xm:f" { return Err("unsupported data bar threshold".into()); }
                depth += 1;
            }
            Event::Start(_) => depth += 1,
            Event::End(ref e) => {
                depth = depth.checked_sub(1).ok_or_else(|| "invalid extended data bar".to_string())?;
                if e.name().as_ref() == b"x14:cfvo" && current.as_ref().is_some_and(|item| item.0 == depth) {
                    let (_, first, tag) = current.take().unwrap(); replacement = Some((first, end, tag));
                } else if e.name().as_ref() == b"x14:dataBar" && bar == Some(depth) { bar = None; }
            }
            Event::Eof => return Err("data bar threshold is missing".into()),
            _ => {}
        }
        if let Some((first, last, original)) = replacement {
            let mut tag = original.to_owned(); tag.clear_attributes();
            for attr in original.attributes() { let attr = attr.map_err(|error| error.to_string())?; if local_name(attr.key.as_ref()) != "type" { tag.push_attribute(attr); } }
            tag.push_attribute(("type", kind));
            let mut writer = Writer::new(Vec::new());
            if let Some(value) = value {
                writer.write_event(Event::Start(tag)).map_err(|error| error.to_string())?;
                writer.write_event(Event::Start(quick_xml::events::BytesStart::new("xm:f"))).map_err(|error| error.to_string())?;
                writer.write_event(Event::Text(quick_xml::events::BytesText::new(value))).map_err(|error| error.to_string())?;
                writer.write_event(Event::End(quick_xml::events::BytesEnd::new("xm:f"))).map_err(|error| error.to_string())?;
                writer.write_event(Event::End(quick_xml::events::BytesEnd::new("x14:cfvo"))).map_err(|error| error.to_string())?;
            } else { writer.write_event(Event::Empty(tag)).map_err(|error| error.to_string())?; }
            return Ok(format!("{}{}{}", &xml[..first], String::from_utf8(writer.into_inner()).map_err(|error| error.to_string())?, &xml[last..]));
        }
    }
}

/// Apply the base-shaped icon view to its x14-only payload, retaining the
/// icon set's other attributes and any custom icon children.
pub fn update_extended_icon_visual(xml: &str, visual: &str) -> Result<String, String> {
    if !valid_conditional_visual(visual, "iconSet") { return Err("invalid icon set".into()); }
    let mut base = Reader::from_str(visual);
    let mut flags = (None, None);
    let mut icon_set = None;
    let mut reset_custom = false;
    let mut replace_custom = false;
    let mut custom_icons = Vec::new();
    let mut criteria = Vec::new();
    let mut depth = 0usize;
    loop {
        let event = base.read_event().map_err(|error| error.to_string())?;
        match event {
            Event::Start(ref e) if depth == 0 && local_name(e.name().as_ref()) == "iconSet" => {
                flags = (get_attr(e, "showValue"), get_attr(e, "reverse")); icon_set = get_attr(e, "iconSet");
                reset_custom = matches!(get_attr(e, "custom").as_deref(), Some("0" | "false")); depth += 1;
                replace_custom = matches!(get_attr(e, "custom").as_deref(), Some("1" | "true"));
            }
            Event::Start(ref e) | Event::Empty(ref e) if depth == 1 && local_name(e.name().as_ref()) == "cfvo" => {
                criteria.push((get_attr(e, "type").ok_or_else(|| "missing icon threshold type".to_string())?, get_attr(e, "val"), get_attr(e, "gte")));
                if matches!(event, Event::Start(_)) { depth += 1; }
            }
            Event::Empty(ref e) if depth == 1 && local_name(e.name().as_ref()) == "cfIcon" && replace_custom => {
                let set = get_attr(e, "iconSet").ok_or_else(|| "missing custom icon set".to_string())?;
                let index = get_attr(e, "iconId").ok_or_else(|| "missing custom icon index".to_string())?;
                custom_icons.push(format!("<x14:cfIcon iconSet=\"{}\" iconId=\"{}\"/>", escape(&set), escape(&index)));
            }
            Event::Start(_) => depth += 1,
            Event::End(_) => depth = depth.checked_sub(1).ok_or_else(|| "invalid icon set".to_string())?,
            Event::Eof => break,
            _ => {}
        }
    }
    let mut reader = Reader::from_str(xml);
    let mut icon_depth = None;
    let mut current: Option<(usize, usize, quick_xml::events::BytesStart<'static>)> = None;
    let mut formula_depth = None;
    let mut replacements = Vec::new();
    let mut depth = 0usize;
    let mut seen = 0usize;
    let mut last_threshold = None;
    let mut custom_icon_start = None;
    let render_threshold = |element: &quick_xml::events::BytesStart<'_>, index: usize| -> Result<String, String> {
        let (kind, value, gte) = criteria.get(index).ok_or_else(|| "icon threshold count mismatch".to_string())?;
        let mut tag = element.to_owned(); tag.clear_attributes();
        for attr in element.attributes() {
            let attr = attr.map_err(|error| error.to_string())?;
            if !["type", "gte"].contains(&local_name(attr.key.as_ref()).as_str()) { tag.push_attribute(attr); }
        }
        tag.push_attribute(("type", kind.as_str()));
        if let Some(gte) = gte { tag.push_attribute(("gte", gte.as_str())); }
        let mut writer = Writer::new(Vec::new());
        if let Some(value) = value {
            writer.write_event(Event::Start(tag)).map_err(|error| error.to_string())?;
            writer.write_event(Event::Start(quick_xml::events::BytesStart::new("xm:f"))).map_err(|error| error.to_string())?;
            writer.write_event(Event::Text(quick_xml::events::BytesText::new(value))).map_err(|error| error.to_string())?;
            writer.write_event(Event::End(quick_xml::events::BytesEnd::new("xm:f"))).map_err(|error| error.to_string())?;
            writer.write_event(Event::End(quick_xml::events::BytesEnd::new("x14:cfvo"))).map_err(|error| error.to_string())?;
        } else { writer.write_event(Event::Empty(tag)).map_err(|error| error.to_string())?; }
        String::from_utf8(writer.into_inner()).map_err(|error| error.to_string())
    };
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref e) if e.name().as_ref() == b"x14:iconSet" => {
                if icon_depth.replace(depth).is_some() { return Err("multiple extended icon sets".into()); }
                let mut tag = e.to_owned(); tag.clear_attributes();
                let mut has_icon_set = false;
                for attr in e.attributes() {
                    let attr = attr.map_err(|error| error.to_string())?;
                    let name = local_name(attr.key.as_ref());
                    if name == "iconSet" {
                        has_icon_set = true;
                        if let Some(value) = &icon_set { tag.push_attribute(("iconSet", value.as_str())); } else { tag.push_attribute(attr); }
                    } else if !["showValue", "reverse"].contains(&name.as_str()) && !((reset_custom || replace_custom) && name == "custom") { tag.push_attribute(attr); }
                }
                if !has_icon_set { if let Some(value) = &icon_set { tag.push_attribute(("iconSet", value.as_str())); } }
                if replace_custom { tag.push_attribute(("custom", "1")); }
                if let Some(value) = &flags.0 { tag.push_attribute(("showValue", value.as_str())); }
                if let Some(value) = &flags.1 { tag.push_attribute(("reverse", value.as_str())); }
                let mut writer = Writer::new(Vec::new()); writer.write_event(Event::Start(tag)).map_err(|error| error.to_string())?;
                replacements.push((start, end, String::from_utf8(writer.into_inner()).map_err(|error| error.to_string())?));
                depth += 1;
            }
            Event::Start(ref e) if icon_depth.is_some_and(|held| depth == held + 1) && e.name().as_ref() == b"x14:cfvo" => {
                if current.replace((depth, start, e.to_owned())).is_some() { return Err("nested icon thresholds".into()); }
                depth += 1;
            }
            Event::Empty(ref e) if icon_depth.is_some_and(|held| depth == held + 1) && e.name().as_ref() == b"x14:cfvo" => {
                replacements.push((start, end, if seen < criteria.len() { render_threshold(e, seen)? } else { String::new() }));
                last_threshold = Some((end, e.to_owned())); seen += 1;
            }
            Event::Empty(ref e) if (reset_custom || replace_custom) && icon_depth.is_some_and(|held| depth == held + 1) && e.name().as_ref() == b"x14:cfIcon" => {
                replacements.push((start, end, String::new()));
            }
            Event::Start(ref e) if (reset_custom || replace_custom) && icon_depth.is_some_and(|held| depth == held + 1) && e.name().as_ref() == b"x14:cfIcon" => {
                custom_icon_start = Some((depth, start)); depth += 1;
            }
            Event::Start(ref e) if current.is_some() => {
                if e.name().as_ref() != b"xm:f" || formula_depth.replace(depth).is_some() { return Err("unsupported extended icon threshold".into()); }
                depth += 1;
            }
            Event::Empty(_) if current.is_some() => return Err("unsupported extended icon threshold".into()),
            Event::Start(_) => depth += 1,
            Event::End(ref e) => {
                depth = depth.checked_sub(1).ok_or_else(|| "invalid extended icon set".to_string())?;
                if e.name().as_ref() == b"xm:f" && formula_depth == Some(depth) { formula_depth = None; }
                else if e.name().as_ref() == b"x14:cfvo" && current.as_ref().is_some_and(|held| held.0 == depth) {
                    let (_, first, tag) = current.take().unwrap();
                    replacements.push((first, end, if seen < criteria.len() { render_threshold(&tag, seen)? } else { String::new() }));
                    last_threshold = Some((end, tag)); seen += 1;
                } else if e.name().as_ref() == b"x14:cfIcon" && custom_icon_start.is_some_and(|held| held.0 == depth) {
                    let (_, first) = custom_icon_start.take().unwrap(); replacements.push((first, end, String::new()));
                } else if e.name().as_ref() == b"x14:iconSet" && icon_depth == Some(depth) {
                    if replace_custom { replacements.push((start, start, custom_icons.join(""))); }
                    icon_depth = None;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if depth != 0 || icon_depth.is_some() || current.is_some() || custom_icon_start.is_some() || seen == 0 || !(3..=5).contains(&criteria.len()) { return Err("invalid extended icon set".into()); }
    if replace_custom && custom_icons.len() != criteria.len() { return Err("custom icon count mismatch".into()); }
    if !reset_custom && !replace_custom && seen != criteria.len() && xml.contains(":cfIcon") { return Err("cannot resize custom icon set".into()); }
    if seen < criteria.len() {
        let (end, tag) = last_threshold.ok_or_else(|| "missing extended icon thresholds".to_string())?;
        let mut added = String::new();
        for index in seen..criteria.len() { added.push_str(&render_threshold(&tag, index)?); }
        replacements.push((end, end, added));
    }
    let mut output = xml.to_string();
    replacements.sort_by_key(|item| std::cmp::Reverse(item.0));
    for (start, end, replacement) in replacements { output.replace_range(start..end, &replacement); }
    Ok(output)
}

/// Replace one colour-scale colour, keeping the rest of the rule intact.
pub fn set_color_scale_color(xml: &str, index: usize, rgb: &str) -> Result<String, String> {
    set_conditional_visual_color(xml, "colorScale", index, rgb)
}

/// Replace a conditional visual's RGB color without changing its other fields.
pub fn set_conditional_visual_color(xml: &str, visual: &str, index: usize, rgb: &str) -> Result<String, String> {
    set_conditional_visual_color_inner(xml, visual, index, Some(rgb))
}

/// Set an automatic/no-color conditional visual color.
pub fn clear_conditional_visual_color(xml: &str, visual: &str, index: usize) -> Result<String, String> {
    set_conditional_visual_color_inner(xml, visual, index, None)
}

/// Set a theme color or change only its tint on a conditional visual.
pub fn set_conditional_visual_color_component(xml: &str, visual: &str, name: &str, value: &str) -> Result<String, String> {
    if !valid_conditional_visual(xml, visual) || !matches!(name, "theme" | "tint") { return Err("invalid conditional color component".into()); }
    let mut reader = Reader::from_str(xml); let mut depth = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) if depth == 1 && local_name(element.name().as_ref()) == "color" => {
                let mut tag = element.to_owned(); tag.clear_attributes();
                for attr in element.attributes() {
                    let attr = attr.map_err(|error| error.to_string())?;
                    let key = local_name(attr.key.as_ref());
                    if key != name && !(name == "theme" && ["rgb", "indexed", "auto", "tint"].contains(&key.as_str())) { tag.push_attribute(attr); }
                }
                tag.push_attribute((name, value));
                let mut writer = Writer::new(Vec::new()); writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(tag) } else { Event::Start(tag) }).map_err(|error| error.to_string())?;
                return Ok(format!("{}{}{}", &xml[..start], String::from_utf8(writer.into_inner()).map_err(|error| error.to_string())?, &xml[end..]));
            }
            Event::Start(_) => depth += 1,
            Event::End(_) => depth = depth.checked_sub(1).ok_or_else(|| "invalid conditional visual".to_string())?,
            Event::Eof => return Err("missing conditional color".into()),
            _ => {}
        }
    }
}

fn set_conditional_visual_color_inner(xml: &str, visual: &str, index: usize, rgb: Option<&str>) -> Result<String, String> {
    if !valid_conditional_visual(xml, visual) || rgb.is_some_and(|rgb| rgb.len() != 6 || !rgb.bytes().all(|b| b.is_ascii_hexdigit())) { return Err("invalid conditional color".into()); }
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut seen = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) if depth == 1 && local_name(element.name().as_ref()) == "color" => {
                if seen == index {
                    let mut tag = element.to_owned(); tag.clear_attributes();
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| error.to_string())?;
                        if !["rgb", "theme", "indexed", "auto", "tint"].contains(&local_name(attribute.key.as_ref()).as_str()) { tag.push_attribute(attribute); }
                    }
                    let color = rgb.map(|rgb| format!("FF{}", rgb.to_ascii_uppercase()));
                    if let Some(color) = &color { tag.push_attribute(("rgb", color.as_str())); } else { tag.push_attribute(("auto", "1")); }
                    let mut writer = Writer::new(Vec::new());
                    writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(tag) } else { Event::Start(tag) }).map_err(|error| error.to_string())?;
                    let changed = String::from_utf8(writer.into_inner()).map_err(|error| error.to_string())?;
                    return Ok(format!("{}{}{}", &xml[..start], changed, &xml[end..]));
                }
                seen += 1;
                if matches!(event, Event::Start(_)) { depth += 1; }
            }
            Event::Start(_) => depth += 1,
            Event::End(_) => depth = depth.checked_sub(1).ok_or_else(|| "invalid color scale".to_string())?,
            Event::Eof => break,
            _ => {}
        }
    }
    Err("color scale criterion is out of range".into())
}

/// Replace a colour-scale threshold while preserving the other criteria.
pub fn set_color_scale_threshold(xml: &str, index: usize, kind: &str, value: Option<&str>) -> Result<String, String> {
    if !valid_conditional_visual(xml, "colorScale") || !["num", "min", "max", "percent", "formula", "percentile"].contains(&kind) { return Err("invalid color scale threshold".into()); }
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut seen = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| error.to_string())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) if depth == 1 && local_name(element.name().as_ref()) == "cfvo" => {
                if seen == index {
                    let mut tag = element.to_owned(); tag.clear_attributes();
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| error.to_string())?;
                        if !["type", "val"].contains(&local_name(attribute.key.as_ref()).as_str()) { tag.push_attribute(attribute); }
                    }
                    tag.push_attribute(("type", kind));
                    if let Some(value) = value { tag.push_attribute(("val", value)); }
                    let mut writer = Writer::new(Vec::new());
                    writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(tag) } else { Event::Start(tag) }).map_err(|error| error.to_string())?;
                    let changed = String::from_utf8(writer.into_inner()).map_err(|error| error.to_string())?;
                    return Ok(format!("{}{}{}", &xml[..start], changed, &xml[end..]));
                }
                seen += 1;
                if matches!(event, Event::Start(_)) { depth += 1; }
            }
            Event::Start(_) => depth += 1,
            Event::End(_) => depth = depth.checked_sub(1).ok_or_else(|| "invalid color scale".to_string())?,
            Event::Eof => break,
            _ => {}
        }
    }
    Err("color scale criterion is out of range".into())
}

fn valid_conditional_visual(xml: &str, kind: &str) -> bool {
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    let mut roots = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => {
                if depth == 0 { roots += 1; if local_name(e.name().as_ref()) != kind { return false; } }
                depth += 1;
            }
            Ok(Event::Empty(e)) => {
                if depth == 0 { roots += 1; if local_name(e.name().as_ref()) != kind { return false; } }
            }
            Ok(Event::End(_)) => { let Some(next) = depth.checked_sub(1) else { return false; }; depth = next; }
            Ok(Event::Text(e)) if depth == 0 && !e.as_ref().iter().all(u8::is_ascii_whitespace) => return false,
            Ok(Event::Eof) => return depth == 0 && roots == 1,
            Ok(_) => {}
            Err(_) => return false,
        }
    }
}

fn patch_conditional_rule(xml: &str, before: &ConditionalRule, after: &ConditionalRule, style_id: Option<u32>) -> Result<String, XlsxError> {
    let mut updates: Vec<(&str, Option<String>)> = Vec::new();
    if before.priority != after.priority { updates.push(("priority", Some(after.priority.to_string()))); }
    if before.operator != after.operator { updates.push(("operator", after.operator.clone())); }
    if before.style != after.style {
        updates.push(("dxfId", Some(style_id.ok_or_else(|| XlsxError::InvalidData("missing differential style".into()))?.to_string())));
    }
    if before.kind != after.kind { updates.push(("type", Some(after.kind.clone()))); }
    if before.stop_if_true != after.stop_if_true { updates.push(("stopIfTrue", Some(u8::from(after.stop_if_true).to_string()))); }
    if before.rank != after.rank { updates.push(("rank", after.rank.map(|value| value.to_string()))); }
    if before.bottom != after.bottom { updates.push(("bottom", Some(u8::from(after.bottom).to_string()))); }
    if before.percent != after.percent { updates.push(("percent", Some(u8::from(after.percent).to_string()))); }
    if before.above_average != after.above_average { updates.push(("aboveAverage", Some(u8::from(after.above_average).to_string()))); }
    if before.equal_average != after.equal_average { updates.push(("equalAverage", Some(u8::from(after.equal_average).to_string()))); }
    if before.standard_deviations != after.standard_deviations { updates.push(("stdDev", after.standard_deviations.map(|value| value.to_string()))); }
    let mut output = if updates.is_empty() { xml.to_string() } else {
        let mut reader = Reader::from_str(xml);
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        let end = reader.buffer_position() as usize;
        let (start, empty) = match event { Event::Start(start) => (start, false), Event::Empty(start) => (start, true), _ => return Err(XlsxError::InvalidData("invalid conditional rule".into())) };
        let mut changed = start.to_owned();
        changed.clear_attributes();
        for attribute in start.attributes() {
            let attribute = attribute.map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            if !updates.iter().any(|(name, _)| attribute.key.as_ref() == name.as_bytes()) { changed.push_attribute(attribute); }
        }
        for (name, value) in &updates { if let Some(value) = value { changed.push_attribute((*name, value.as_str())); } }
        let mut writer = Writer::new(Vec::new());
        writer.write_event(if empty { Event::Empty(changed) } else { Event::Start(changed) }).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let mut output = String::from_utf8(writer.into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        output.push_str(&xml[end..]);
        output
    };
    if before.visual_xml != after.visual_xml {
        let raw = after.visual_xml.as_deref().ok_or_else(|| XlsxError::InvalidData("missing conditional visual settings".into()))?;
        if before.kind != after.kind || !valid_conditional_visual(raw, &after.kind) { return Err(XlsxError::InvalidData("invalid conditional visual settings".into())); }
        output = replace_direct_xml_child(&output, &after.kind, Some(raw))?;
    }
    if before.visual_extension_id.is_none() && after.kind == "dataBar" {
        if let Some(id) = after.visual_extension_id.as_deref() {
            let link = format!("<ext uri=\"{{B025F937-C7B1-47D3-B67F-A62EFF666E3E}}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:id>{}</x14:id></ext>", escape(id));
            let (child, _, empty, _) = direct_xml_child(&output, "extLst")?;
            if empty { return Err(XlsxError::InvalidData("cannot attach data bar extension".into())); }
            if let Some((start, end)) = child {
                let raw = &output[start..end];
                let updated = if raw.ends_with("/>") { format!("{}>{link}</extLst>", &raw[..raw.len()-2]) } else { append_xml_child(raw, "extLst", &link)? };
                output.replace_range(start..end, &updated);
            } else { output = append_xml_child(&output, "cfRule", &format!("<extLst>{link}</extLst>"))?; }
        }
    }
    output = patch_conditional_formulas(&output, &before.formulas, &after.formulas)?;
    Ok(output)
}

fn patch_conditional_formulas(xml: &str, before: &[String], after: &[String]) -> Result<String, XlsxError> {
    if before == after { return Ok(xml.to_string()); }
    let fail = || XlsxError::InvalidData("cannot safely rewrite conditional formulas".into());
    if before.len() != after.len() {
        let mut reader = Reader::from_str(xml);
        let mut depth = 0usize;
        let mut open = None;
        let mut spans = Vec::new();
        let mut insertion = None;
        let formulas = after.iter().map(|formula| format!("<formula>{}</formula>", escape(formula))).collect::<String>();
        loop {
            let start = reader.buffer_position() as usize;
            let event = reader.read_event().map_err(|_| fail())?;
            let end = reader.buffer_position() as usize;
            match event {
                Event::Start(ref element) => {
                    if depth == 1 && local_name(element.name().as_ref()) == "formula" {
                        if element.attributes().next().is_some() || open.replace(start).is_some() { return Err(fail()); }
                    } else if depth == 1 && local_name(element.name().as_ref()) == "extLst" && insertion.is_none() { insertion = Some(start); }
                    depth += 1;
                }
                Event::Empty(ref element) if depth == 0 => {
                    if !before.is_empty() || end < 2 || &xml[end - 2..end] != "/>" { return Err(fail()); }
                    let name = std::str::from_utf8(element.name().as_ref()).map_err(|_| fail())?.to_string();
                    return Ok(format!("{}>{formulas}</{name}>{}", &xml[..end - 2], &xml[end..]));
                }
                Event::Empty(ref element) if depth == 1 && local_name(element.name().as_ref()) == "formula" => {
                    if element.attributes().next().is_some() { return Err(fail()); }
                    spans.push((start, end));
                }
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(fail)?;
                    if depth == 1 { if let Some(begin) = open.take() { spans.push((begin, end)); } }
                    if depth == 0 { insertion.get_or_insert(start); break; }
                }
                Event::Eof => return Err(fail()),
                _ => {}
            }
        }
        if open.is_some() || spans.len() != before.len() { return Err(fail()); }
        let mut output = xml.to_string();
        if spans.is_empty() { output.insert_str(insertion.ok_or_else(fail)?, &formulas); }
        else {
            for (index, (start, end)) in spans.into_iter().enumerate().rev() {
                output.replace_range(start..end, if index == 0 { &formulas } else { "" });
            }
        }
        return Ok(output);
    }
    let mut reader = Reader::from_str(xml);
    let mut open = None;
    let mut spans = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|_| fail())?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) if local_name(element.name().as_ref()) == "formula" => {
                if open.replace(end).is_some() { return Err(fail()); }
            }
            Event::End(ref element) if local_name(element.name().as_ref()) == "formula" => spans.push((open.take().ok_or_else(fail)?, start)),
            Event::Eof => break,
            _ => {}
        }
    }
    if open.is_some() || spans.len() != after.len() { return Err(fail()); }
    let mut output = xml.to_string();
    for ((start, end), formula) in spans.into_iter().zip(after).rev() { output.replace_range(start..end, &escape(formula)); }
    Ok(output)
}

/// Replace worksheet validation rules while preserving the exact XML of any
/// unchanged rule. When only its range moved, retain its unmodelled attributes
/// and children and rewrite just sqref.
fn patch_data_validations_xml(xml: &str, before: &[DataValidation], after: &[DataValidation]) -> Result<String, XlsxError> {
    let block_start = xml.find("<dataValidations");
    let mut old_entries = Vec::new();
    let mut old_block = None;
    if let Some(start) = block_start {
        let head_end = xml[start..].find('>').map(|n| start + n + 1)
            .ok_or_else(|| XlsxError::InvalidData("incomplete dataValidations element".into()))?;
        if xml[start..head_end].trim_end().ends_with("/>") {
            old_block = Some((start, head_end));
        } else {
            let tail = xml[head_end..].find("</dataValidations>").map(|n| head_end + n)
                .ok_or_else(|| XlsxError::InvalidData("unclosed dataValidations element".into()))?;
            let mut cursor = head_end;
            while let Some(offset) = xml[cursor..tail].find("<dataValidation") {
                let entry_start = cursor + offset;
                let entry_head = xml[entry_start..tail].find('>').map(|n| entry_start + n + 1)
                    .ok_or_else(|| XlsxError::InvalidData("incomplete dataValidation element".into()))?;
                let entry_end = if xml[entry_start..entry_head].trim_end().ends_with("/>") { entry_head }
                    else { xml[entry_head..tail].find("</dataValidation>").map(|n| entry_head + n + "</dataValidation>".len())
                        .ok_or_else(|| XlsxError::InvalidData("unclosed dataValidation element".into()))? };
                old_entries.push(&xml[entry_start..entry_end]);
                cursor = entry_end;
            }
            old_block = Some((start, tail + "</dataValidations>".len()));
        }
    }
    if old_entries.len() != before.len() {
        return Err(XlsxError::InvalidData("cannot safely reconcile worksheet validation rules".into()));
    }
    let sqref = |rule: &DataValidation| rule.ranges.iter().map(|&(r0, c0, r1, c1)| {
        let first = format!("{}{}", col_to_letter(c0), r0 + 1);
        if (r0, c0) == (r1, c1) { first } else { format!("{}:{}{}", first, col_to_letter(c1), r1 + 1) }
    }).collect::<Vec<_>>().join(" ");
    let mut used = vec![false; before.len()];
    let mut entries = String::new();
    for rule in after {
        let reference = sqref(rule);
        if reference.is_empty() { continue; }
        let same = before.iter().enumerate().position(|(i, old)| !used[i] && old == rule);
        if let Some(i) = same { used[i] = true; entries.push_str(old_entries[i]); continue; }
        let moved = before.iter().enumerate().position(|(i, old)| {
            if used[i] { return false; }
            let mut copy = old.clone(); copy.ranges = rule.ranges.clone(); &copy == rule
        });
        if let Some(i) = moved {
            let raw = old_entries[i];
            let head_end = raw.find('>').ok_or_else(|| XlsxError::InvalidData("incomplete validation rule".into()))?;
            let key = raw[..head_end].find("sqref=").ok_or_else(|| XlsxError::InvalidData("validation rule has no sqref".into()))?;
            let quote_at = key + "sqref=".len();
            let quote = raw.as_bytes().get(quote_at).copied().filter(|c| *c == b'\"' || *c == b'\'')
                .ok_or_else(|| XlsxError::InvalidData("invalid validation sqref".into()))?;
            let end = raw[quote_at + 1..head_end].find(quote as char).map(|n| quote_at + 1 + n)
                .ok_or_else(|| XlsxError::InvalidData("unclosed validation sqref".into()))?;
            entries.push_str(&raw[..quote_at + 1]); entries.push_str(&escape(&reference)); entries.push_str(&raw[end..]);
            used[i] = true; continue;
        }
        entries.push_str(&format!("<dataValidation type=\"{}\" sqref=\"{}\" allowBlank=\"{}\" showDropDown=\"{}\" showInputMessage=\"{}\" showErrorMessage=\"{}\"",
            escape(if rule.kind == "any" { "none" } else { &rule.kind }), escape(&reference), u8::from(rule.allow_blank),
            u8::from(!rule.in_cell_dropdown), u8::from(rule.show_input), u8::from(rule.show_error)));
        for (key, value) in [("operator", &rule.operator), ("errorStyle", &rule.error_style), ("promptTitle", &rule.prompt_title),
            ("errorTitle", &rule.error_title), ("prompt", &rule.prompt), ("error", &rule.error)] {
            if let Some(value) = value { entries.push_str(&format!(" {key}=\"{}\"", escape(value))); }
        }
        entries.push('>');
        if let Some(formula) = &rule.formula1 { entries.push_str(&format!("<formula1>{}</formula1>", escape(formula))); }
        if let Some(formula) = &rule.formula2 { entries.push_str(&format!("<formula2>{}</formula2>", escape(formula))); }
        entries.push_str("</dataValidation>");
    }
    let count = after.iter().filter(|rule| !rule.ranges.is_empty()).count();
    let replacement = if count == 0 { String::new() } else { format!("<dataValidations count=\"{count}\">{entries}</dataValidations>") };
    if let Some((start, end)) = old_block { return Ok(format!("{}{}{}", &xml[..start], replacement, &xml[end..])); }
    if replacement.is_empty() { return Ok(xml.to_string()); }
    let insert = ["<hyperlinks", "<printOptions", "<pageMargins", "<pageSetup", "<headerFooter", "<rowBreaks", "<colBreaks", "<drawing", "<legacyDrawing", "<tableParts", "<extLst", "</worksheet>"]
        .iter().filter_map(|tag| xml.find(tag)).min()
        .ok_or_else(|| XlsxError::InvalidData("worksheet has no validation insertion point".into()))?;
    Ok(format!("{}{}{}", &xml[..insert], replacement, &xml[insert..]))
}

fn append_content_type(xml: &str, part: &str, kind: &str) -> Result<String, XlsxError> {
    if xml.contains(&format!("PartName=\"/{}\"", part)) { return Ok(xml.to_string()); }
    let at = xml.rfind("</Types>")
        .ok_or_else(|| XlsxError::InvalidData("content types part has no closing element".into()))?;
    let mut output = xml.to_string();
    output.insert_str(at, &format!("<Override PartName=\"/{}\" ContentType=\"{}\"/>", escape(part), escape(kind)));
    Ok(output)
}

fn append_default_content_type(xml: &str, extension: &str, kind: &str) -> Result<String, XlsxError> {
    if xml.to_ascii_lowercase().contains(&format!("extension=\"{}\"", extension.to_ascii_lowercase())) {
        return Ok(xml.to_string());
    }
    let at = xml.rfind("</Types>")
        .ok_or_else(|| XlsxError::InvalidData("content types part has no closing element".into()))?;
    let mut output = xml.to_string();
    output.insert_str(at, &format!("<Default Extension=\"{}\" ContentType=\"{}\"/>", escape(extension), escape(kind)));
    Ok(output)
}

fn set_legacy_drawing(xml: &str, relationship_id: &str) -> Result<String, XlsxError> {
    // Minimal worksheets often bind only the default SpreadsheetML
    // namespace. Keep the relationship prefix in scope even on those files.
    let replacement = format!(
        "<legacyDrawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{}\"/>",
        escape(relationship_id),
    );
    set_worksheet_drawing_child(xml, "legacyDrawing", &replacement,
        &["legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"])
}

fn set_drawing(xml: &str, relationship_id: &str) -> Result<String, XlsxError> {
    let replacement = format!(
        "<drawing xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{}\"/>",
        escape(relationship_id),
    );
    set_worksheet_drawing_child(xml, "drawing", &replacement,
        &["legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"])
}

fn set_worksheet_drawing_child(xml: &str, name: &str, replacement: &str, later: &[&str]) -> Result<String, XlsxError> {
    let (existing, close, empty, _) = direct_xml_child(xml, name)?;
    if let Some((start, end)) = existing {
        return Ok(format!("{}{}{}", &xml[..start], replacement, &xml[end..]));
    }
    if empty { return replace_direct_xml_child(xml, name, Some(replacement)); }
    let mut insertion = close;
    for element in later {
        if let Some((start, _)) = direct_xml_child(xml, element)?.0 {
            insertion = insertion.min(start);
        }
    }
    Ok(format!("{}{}{}", &xml[..insertion], replacement, &xml[insertion..]))
}

fn merge_relationship_additions(base: &str, edited: &str) -> Result<String, XlsxError> {
    let held = parse_relationships(base)
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    let additions = parse_relationships(edited)
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    let mut output = base.to_string();
    for (id, relation) in additions {
        if let Some(existing) = held.get(&id) {
            if existing.rel_type != relation.rel_type || existing.target != relation.target || existing.target_mode != relation.target_mode {
                return Err(XlsxError::InvalidData(format!("conflicting relationship ID {id}")));
            }
            continue;
        }
        let at = output.rfind("</Relationships>")
            .ok_or_else(|| XlsxError::InvalidData("relationships part has no closing element".into()))?;
        let mode = relation.target_mode.as_ref().map(|mode| format!(" TargetMode=\"{}\"", escape(mode))).unwrap_or_default();
        output.insert_str(at, &format!("<Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"{mode}/>", escape(&id), escape(&relation.rel_type), escape(&relation.target)));
    }
    Ok(output)
}

fn unique_threaded_guid(seed: &str, existing: &std::collections::HashSet<String>) -> String {
    use std::hash::{Hash, Hasher};
    let mut attempt = 0u64;
    loop {
        let mut first = std::collections::hash_map::DefaultHasher::new();
        seed.hash(&mut first);
        attempt.hash(&mut first);
        let mut second = std::collections::hash_map::DefaultHasher::new();
        "threaded-comment-guid".hash(&mut second);
        seed.hash(&mut second);
        attempt.wrapping_add(1).hash(&mut second);
        let mut bytes = [0u8; 16];
        bytes[..8].copy_from_slice(&first.finish().to_be_bytes());
        bytes[8..].copy_from_slice(&second.finish().to_be_bytes());
        bytes[6] = (bytes[6] & 0x0f) | 0x40;
        bytes[8] = (bytes[8] & 0x3f) | 0x80;
        let id = format!(
            "{{{:02X}{:02X}{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}}}",
            bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
            bytes[8], bytes[9], bytes[10], bytes[11], bytes[12], bytes[13], bytes[14], bytes[15]
        );
        if !existing.contains(&id) { return id; }
        attempt = attempt.wrapping_add(1);
    }
}

fn font_xml(style: &CellStyle) -> String {
    let mut font = String::from("<font>");
    if style.bold {
        font.push_str("<b/>");
    }
    if style.italic {
        font.push_str("<i/>");
    }
    if style.underline {
        font.push_str("<u/>");
    }
    if style.strikethrough {
        font.push_str("<strike/>");
    }
    if let Some(size) = style.font_size {
        font.push_str(&format!("<sz val=\"{size}\"/>"));
    }
    if let Some(color) = style.font_color.as_deref() {
        font.push_str(&format!("<color rgb=\"FF{}\"/>", escape(color)));
    }
    if let Some(name) = style.font_name.as_deref() {
        font.push_str(&format!("<name val=\"{}\"/>", escape(name)));
    }
    font.push_str("</font>");
    font
}

fn fill_xml(style: &CellStyle) -> String {
    match style.bg_color.as_deref() {
        Some(color) => format!(
            "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF{}\"/><bgColor indexed=\"64\"/></patternFill></fill>",
            escape(color)
        ),
        None => "<fill><patternFill patternType=\"none\"/></fill>".to_string(),
    }
}

fn border_xml(style: &CellStyle) -> String {
    let edge = |name: &str, line: &Option<BorderLine>| match line {
        Some(line) => {
            let colour = match line.color.as_deref() {
                Some(colour) => format!("<color rgb=\"FF{}\"/>", escape(colour)),
                None => "<color indexed=\"64\"/>".to_string(),
            };
            format!(
                "<{name} style=\"{}\">{colour}</{name}>",
                escape(&line.style)
            )
        }
        None => format!("<{name}/>"),
    };
    format!(
        "<border>{}{}{}{}<diagonal/></border>",
        edge("left", &style.border_left),
        edge("right", &style.border_right),
        edge("top", &style.border_top),
        edge("bottom", &style.border_bottom)
    )
}

fn xf_xml(
    style: &CellStyle,
    number_format_id: u32,
    font_id: u32,
    fill_id: u32,
    border_id: u32,
) -> String {
    let mut xf = format!(
        "<xf numFmtId=\"{number_format_id}\" fontId=\"{font_id}\" fillId=\"{fill_id}\" borderId=\"{border_id}\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\""
    );
    if number_format_id != 0 {
        xf.push_str(" applyNumberFormat=\"1\"");
    }
    if style.quote_prefix {
        xf.push_str(" quotePrefix=\"1\"");
    }
    // Everything an `<alignment>` can say goes in one element, so it is built
    // up rather than written from the first part that happens to be set. Wrap,
    // indent and the vertical part were read all along and never written: a
    // cell told to wrap looked right on screen and came back unwrapped, which
    // is worse than not being able to tell it to.
    let mut alignment = String::new();
    if let Some(horizontal) = style.horizontal_align.as_deref() {
        alignment.push_str(&format!(" horizontal=\"{}\"", escape(horizontal)));
    }
    if let Some(vertical) = style.vertical_align.as_deref() {
        alignment.push_str(&format!(" vertical=\"{}\"", escape(vertical)));
    }
    if style.wrap_text {
        alignment.push_str(" wrapText=\"1\"");
    }
    if style.shrink_to_fit {
        alignment.push_str(" shrinkToFit=\"1\"");
    }
    if style.indent > 0 {
        alignment.push_str(&format!(" indent=\"{}\"", style.indent));
    }
    if alignment.is_empty() {
        xf.push_str("/>");
    } else {
        xf.push_str(" applyAlignment=\"1\">");
        xf.push_str(&format!("<alignment{alignment}/>"));
        xf.push_str("</xf>");
    }
    xf
}

/// The well-known number formats, so a common one does not need a new entry.
fn builtin_number_format_id(format: &str) -> Option<u32> {
    match format {
        "General" => Some(0),
        "0" => Some(1),
        "0.00" => Some(2),
        "#,##0" => Some(3),
        "#,##0.00" => Some(4),
        "$#,##0.00_);($#,##0.00)" => Some(7),
        "0%" => Some(9),
        "0.00%" => Some(10),
        "0.00E+00" => Some(11),
        "mm-dd-yy" => Some(14),
        "m/d/yy h:mm" => Some(22),
        _ => None,
    }
}

/// Rewrites names with their current sheet index and retained attributes.
fn patch_workbook_names(xml: &str, names: &[(String, String)], attributes: &crate::ir::DefinedNameAttributes) -> Result<String, XlsxError> {
    let (start, end) = match xml.find("<definedNames") {
        Some(start) => {
            let head_end = xml[start..].find('>').map(|at| start + at + 1)
                .ok_or_else(|| XlsxError::InvalidData("the workbook's name list has no end".to_string()))?;
            let end = if xml[start..head_end].ends_with("/>") { head_end } else {
                xml[head_end..].find("</definedNames>").map(|at| head_end + at + "</definedNames>".len())
                    .ok_or_else(|| XlsxError::InvalidData("the workbook's name list has no end".to_string()))?
            };
            (start, end)
        }
        None => {
            let after = xml.find("</sheets>").map(|at| at + "</sheets>".len())
                .or_else(|| xml.find("<sheets").and_then(|at| xml[at..].find("/>").map(|n| at + n + 2)))
                .ok_or_else(|| XlsxError::InvalidData("the workbook does not list its sheets".to_string()))?;
            (after, after)
        }
    };
    if names.is_empty() { return Ok(format!("{}{}", &xml[..start], &xml[end..])); }
    let mut sheets = Vec::new();
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(e) | Event::Empty(e)) if local_name(e.name().as_ref()) == "sheet" => {
                sheets.push(get_attr(&e, "name").unwrap_or_default());
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    let mut listed = String::from("<definedNames>");
    for (name, refers_to) in names {
        let scope = crate::ir::defined_name_scope(name);
        let leaf = scope.as_ref().map(|(_, leaf)| *leaf).unwrap_or(name);
        let held = attributes.get(name);
        let imported = held.and_then(|a| a.get("name"));
        let builtin = imported.is_some_and(|old| old.strip_prefix("_xlnm.").is_some_and(|old| old == leaf))
            || ["_FilterDatabase", "Print_Area", "Print_Titles", "Criteria", "Extract"].contains(&leaf);
        let xml_name = if builtin { format!("_xlnm.{leaf}") } else { leaf.to_string() };
        let mut element = BytesStart::new("definedName");
        element.push_attribute(("name", xml_name.as_str()));
        let local_index;
        if let Some((owner, _)) = scope {
            let index = sheets.iter().position(|sheet| oxicells_calc::case_fold::name_key(sheet) == oxicells_calc::case_fold::name_key(&owner))
                .ok_or_else(|| XlsxError::InvalidData(format!("defined name {name:?} belongs to a missing sheet")))?;
            local_index = index.to_string();
            element.push_attribute(("localSheetId", local_index.as_str()));
        }
        if let Some(held) = held {
            for (key, value) in held {
                if key != "name" && key != "localSheetId" { element.push_attribute((key.as_str(), value.as_str())); }
            }
        }
        if leaf == "_FilterDatabase" && !held.is_some_and(|a| a.contains_key("hidden")) { element.push_attribute(("hidden", "1")); }
        let mut writer = Writer::new(Vec::new());
        writer.write_event(Event::Start(element)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        let formula = oxicells_calc::lexer::name_formula_to_ooxml(refers_to).unwrap_or_else(|_| refers_to.clone());
        writer.write_event(Event::Text(BytesText::new(&formula))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        writer.write_event(Event::End(BytesEnd::new("definedName"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
        listed.push_str(&String::from_utf8(writer.into_inner()).map_err(|e| XlsxError::InvalidData(e.to_string()))?);
    }
    listed.push_str("</definedNames>");
    Ok(format!("{}{}{}", &xml[..start], listed, &xml[end..]))
}

/// Rewrites the `<sheets>` list of `xl/workbook.xml`.
///
/// Each sheet is named in order and given a relationship id matching its
/// position, which is what `patch_workbook_rels` writes on the other side.
fn patch_workbook_sheets(
    xml: &str,
    names: &[(String, Option<&'static str>)],
) -> Result<String, XlsxError> {
    let start = xml.find("<sheets").ok_or_else(|| {
        XlsxError::InvalidData("the workbook does not list its sheets".to_string())
    })?;
    let end = xml[start..]
        .find("</sheets>")
        .map(|at| start + at + "</sheets>".len())
        .or_else(|| {
            // A self-closing <sheets/> holds nothing, but is still the place.
            xml[start..].find("/>").map(|at| start + at + 2)
        })
        .ok_or_else(|| {
            XlsxError::InvalidData("the workbook's sheet list has no end".to_string())
        })?;

    let mut listed = String::from("<sheets>");
    for (position, (name, state)) in names.iter().enumerate() {
        // A hidden sheet says so here and nowhere else, so writing the list
        // without its state would quietly bring it back into view.
        let state = state
            .map(|state| format!(" state=\"{state}\""))
            .unwrap_or_default();
        listed.push_str(&format!(
            "<sheet name=\"{}\"{state} sheetId=\"{}\" r:id=\"rIdSheet{}\"/>",
            escape(name),
            position + 1,
            position + 1
        ));
    }
    listed.push_str("</sheets>");
    Ok(format!("{}{}{}", &xml[..start], listed, &xml[end..]))
}

/// The same XML with any element naming the calculation chain taken out.
///
/// Used on `[Content_Types].xml` and on the workbook's relationships, where
/// the chain is declared. Both state it as a single self-closing element, so
/// cutting from its `<` to the `>` that closes it is enough and leaves the
/// rest of the part byte for byte as it was.
fn without_calc_chain(xml: &str) -> String {
    let mut out = String::with_capacity(xml.len());
    let mut rest = xml;
    while let Some(at) = rest.find("calcChain.xml") {
        let Some(open) = rest[..at].rfind('<') else {
            break;
        };
        let Some(close) = rest[at..].find('>') else {
            break;
        };
        out.push_str(&rest[..open]);
        rest = &rest[at + close + 1..];
    }
    out.push_str(rest);
    out
}

/// Rewrites the worksheet relationships of `xl/_rels/workbook.xml.rels`,
/// leaving every other relationship  Estyles, theme, shared strings  Eas it is.
fn patch_workbook_rels(xml: &str, targets: &[String]) -> Result<String, XlsxError> {
    const WORKSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    let original_types:HashMap<String,String>=parse_relationships(xml)?.into_values()
        .map(|rel|(rel.target.trim_start_matches('/').strip_prefix("xl/").unwrap_or(rel.target.trim_start_matches('/')).to_string(),rel.rel_type)).collect();

    let mut kept = String::new();
    let mut rest = xml;
    let end = rest.find("</Relationships>").ok_or_else(|| {
        XlsxError::InvalidData("the workbook has no relationships".to_string())
    })?;
    let tail = &rest[end..];
    rest = &rest[..end];

    // Keep everything that is not a worksheet, in the order it was written.
    let head_end = rest.find('>').map(|at| at + 1).unwrap_or(0);
    let mut head = rest[..head_end].to_string();
    if !head.contains("<Relationships") {
        // The declaration came first; find the real opening tag.
        let at = rest.find("<Relationships").ok_or_else(|| {
            XlsxError::InvalidData("the workbook has no relationships".to_string())
        })?;
        let close = rest[at..].find('>').map(|offset| at + offset + 1).unwrap();
        head = rest[..close].to_string();
        rest = &rest[close..];
    } else {
        rest = &rest[head_end..];
    }

    for piece in rest.split_inclusive("/>") {
        if piece.contains("/worksheet\"") || piece.contains("/chartsheet\"") {
            continue;
        }
        kept.push_str(piece);
    }

    let mut added = String::new();
    for (position, target) in targets.iter().enumerate() {
        let relationship_type=original_types.get(target).map(String::as_str).unwrap_or(WORKSHEET);
        added.push_str(&format!(
            "<Relationship Id=\"rIdSheet{}\" Type=\"{}\" Target=\"{}\"/>",
            position + 1,
            escape(relationship_type),
            escape(target)
        ));
    }
    Ok(format!("{head}{added}{kept}{tail}"))
}

/// Rewrites the worksheet overrides of `[Content_Types].xml`.
fn patch_content_types(xml: &str, parts: &[String]) -> Result<String, XlsxError> {
    const WORKSHEET: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    let mut original_types=HashMap::new();
    let mut reader=Reader::from_str(xml);
    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Empty(element)|Event::Start(element) if local_name(element.name().as_ref())=="Override"=> {
                if let (Some(part),Some(kind))=(get_attr(&element,"PartName"),get_attr(&element,"ContentType")) {
                    original_types.insert(part.trim_start_matches('/').to_string(),kind);
                }
            }
            Event::Eof=>break,
            _=>{}
        }
    }

    let end = xml.find("</Types>").ok_or_else(|| {
        XlsxError::InvalidData("the package has no content types".to_string())
    })?;
    let (body, tail) = xml.split_at(end);

    let mut kept = String::new();
    for piece in body.split_inclusive("/>") {
        if piece.contains("spreadsheetml.worksheet+xml") || piece.contains("spreadsheetml.chartsheet+xml") {
            continue;
        }
        kept.push_str(piece);
    }

    let mut added = String::new();
    for part in parts {
        let content_type=original_types.get(part).map(String::as_str).unwrap_or(WORKSHEET);
        added.push_str(&format!(
            "<Override PartName=\"/{}\" ContentType=\"{}\"/>",
            escape(part),escape(content_type)
        ));
    }
    Ok(format!("{kept}{added}{tail}"))
}

/// Writes a worksheet part from scratch, for a sheet the file never held.
fn patch_workbook_calculation(xml: &str, settings: CalculationSettings) -> Result<String, XlsxError> {
    fn element(original: Option<&BytesStart<'_>>, settings: CalculationSettings, fallback: &str) -> BytesStart<'static> {
        let name = original.map(|e| String::from_utf8_lossy(e.name().as_ref()).into_owned()).unwrap_or_else(|| fallback.to_string());
        let mut result = BytesStart::new(name);
        if let Some(original) = original {
            for attr in original.attributes().flatten() {
                if !matches!(attr.key.as_ref(), b"calcMode" | b"calcOnSave" | b"calcCompleted") {
                    result.push_attribute(attr);
                }
            }
        }
        match settings.mode {
            CalculationMode::Automatic => {}
            CalculationMode::Manual => result.push_attribute(("calcMode", "manual")),
            CalculationMode::AutomaticExceptTables => result.push_attribute(("calcMode", "autoNoTable")),
        }
        if !settings.on_save { result.push_attribute(("calcOnSave", "0")); }
        if let Some(completed) = settings.completed {
            result.push_attribute(("calcCompleted", if completed { "1" } else { "0" }));
        }
        result.into_owned()
    }
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut written = false;
    let mut skipped = 0;
    let mut calc_name = "calcPr".to_string();
    loop {
        let event = reader.read_event().map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        if skipped > 0 {
            match event { Event::Start(_) => skipped += 1, Event::End(_) => skipped -= 1, Event::Eof => break, _ => {} }
            continue;
        }
        match &event {
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == "workbook" => {
                if let Some((prefix, _)) = std::str::from_utf8(e.name().as_ref()).ok().and_then(|name| name.split_once(':')) {
                    calc_name = format!("{prefix}:calcPr");
                }
                if matches!(event, Event::Empty(_)) {
                    let name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.write_event(Event::Start(e.clone())).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    writer.write_event(Event::Empty(element(None, settings, &calc_name))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    writer.write_event(Event::End(BytesEnd::new(name))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    written = true;
                    continue;
                }
            }
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == "calcPr" => {
                writer.write_event(Event::Empty(element(Some(e), settings, &calc_name))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                written = true;
                if matches!(event, Event::Start(_)) { skipped = 1; }
                continue;
            }
            Event::Start(e) | Event::Empty(e) if !written && matches!(local_name(e.name().as_ref()).as_str(),
                "oleSize" | "customWorkbookViews" | "pivotCaches" | "smartTagPr" | "smartTagTypes" | "webPublishing" | "fileRecoveryPr" | "webPublishObjects" | "extLst") => {
                writer.write_event(Event::Empty(element(None, settings, &calc_name))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                written = true;
            }
            Event::End(e) if !written && local_name(e.name().as_ref()) == "workbook" => {
                writer.write_event(Event::Empty(element(None, settings, &calc_name))).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            Event::Eof => break,
            _ => {}
        }
        writer.write_event(event).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }
    let patched = String::from_utf8(writer.into_inner().into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    crate::compatibility::patch_version(&patched, settings.compatibility_version)
}

fn patch_default_column_width(xml: &str, width: f32) -> Result<String, XlsxError> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut written = false;
    loop {
        let event = reader.read_event().map_err(XlsxError::Xml)?;
        match &event {
            Event::Start(start) | Event::Empty(start) if start.local_name().as_ref() == b"sheetFormatPr" => {
                let mut changed = start.to_owned();
                changed.clear_attributes();
                for attribute in start.attributes().flatten() {
                    if attribute.key.as_ref() != b"defaultColWidth" { changed.push_attribute(attribute); }
                }
                changed.push_attribute(("defaultColWidth", width_text(width).as_str()));
                writer.write_event(if matches!(event, Event::Empty(_)) { Event::Empty(changed) } else { Event::Start(changed) })
                    .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                written = true;
                continue;
            }
            Event::Start(start) | Event::Empty(start) if !written && matches!(start.local_name().as_ref(), b"cols" | b"sheetData") => {
                let mut format = BytesStart::new("sheetFormatPr");
                format.push_attribute(("defaultRowHeight", "15"));
                format.push_attribute(("defaultColWidth", width_text(width).as_str()));
                writer.write_event(Event::Empty(format)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                written = true;
            }
            Event::Eof => break,
            _ => {}
        }
        writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
    }
    String::from_utf8(writer.into_inner().into_inner()).map_err(|e| XlsxError::InvalidData(e.to_string()))
}

fn write_new_sheet(
    sheet: &crate::ir::Sheet,
    styles: &BTreeMap<(u32, u32), u32>,
    row_styles: &BTreeMap<u32, u32>,
    column_styles: &BTreeMap<u32, u32>,
    multiple: &dyn Fn(&str) -> bool,
) -> Result<Vec<u8>, XlsxError> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut root = BytesStart::new("worksheet");
    root.push_attribute((
        "xmlns",
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    ));
    writer
        .write_event(Event::Start(root))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;

    if let Some(hex) = &sheet.tab_color {
        writer
            .write_event(Event::Start(BytesStart::new("sheetPr")))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let mut colour = BytesStart::new("tabColor");
        colour.push_attribute(("rgb", format!("FF{hex}").as_str()));
        writer
            .write_event(Event::Empty(colour))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        writer
            .write_event(Event::End(BytesEnd::new("sheetPr")))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }

    if sheet.has_default_column_width() {
        let mut format = BytesStart::new("sheetFormatPr");
        format.push_attribute(("defaultRowHeight", width_text(sheet.default_row_height).as_str()));
        format.push_attribute(("defaultColWidth", width_text(sheet.default_col_width).as_str()));
        writer.write_event(Event::Empty(format)).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }
    let columns: std::collections::BTreeSet<u32> = (0..sheet.col_widths.len() as u32)
        .chain(sheet.hidden_cols.iter().copied()).chain(sheet.col_custom_widths.keys().copied()).collect();
    let mut pending = BTreeMap::new();
    for column in columns {
        let width = sheet.col_widths.get(column as usize).copied().unwrap_or(0.0);
        let hidden = sheet.hidden_cols.contains(&column);
        let custom = sheet.column_has_custom_width(column);
        if width > 0.0 || hidden || custom { pending.insert(column, ColumnWant { hidden: Some(hidden), width: Some(width), custom: Some(custom), style: None, }); }
    }
    for (&column, &style) in column_styles {
        pending.entry(column).or_default().style = Some(Some(style));
    }
    if !pending.is_empty() {
        open_cols(&mut writer)?;
        write_new_cols(&mut writer, &mut pending)?;
        close_cols(&mut writer)?;
    }

    if sheet.rows.iter().all(|row| row.cells.is_empty() && !row.hidden && row.height.is_none()) && row_styles.is_empty() {
        writer
            .write_event(Event::Empty(BytesStart::new("sheetData")))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    } else {
        writer
            .write_event(Event::Start(BytesStart::new("sheetData")))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        let all_rows: std::collections::BTreeSet<u32> = sheet.rows.iter().map(|row| row.index).chain(row_styles.keys().copied()).collect();
        for index in all_rows {
            let row = sheet.rows.iter().find(|row| row.index == index);
            let mut cells = BTreeMap::new();
            for cell in row.into_iter().flat_map(|row| &row.cells) {
                if let Some(value) = edit_for(cell, index, multiple) {
                    cells.insert(cell.col, value);
                }
            }
            let hidden = row.is_some_and(|row| row.hidden);
            let height = row.and_then(|row| row.height);
            if cells.is_empty() && !hidden && height.is_none() && !row_styles.contains_key(&index) {
                continue;
            }
            write_inserted_row(
                &mut writer,
                index,
                cells,
                Some(RowWant {
                    hidden: Some(hidden),
                    height: Some(height),
                    pinned: row.is_some_and(|row| row.custom_height),
                    style: row_styles.get(&index).map(|style| Some(*style)),
                }),
                styles,
            )?;
        }
        writer
            .write_event(Event::End(BytesEnd::new("sheetData")))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }

    write_merges(&mut writer, &sheet.merge_cells)?;
    writer
        .write_event(Event::End(BytesEnd::new("worksheet")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;

    let mut part = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.to_vec();
    part.extend(writer.into_inner().into_inner());
    let edits: HashMap<_,_> = sheet.rows.iter().flat_map(|row|row.cells.iter().filter_map(move |cell|edit_for(cell,row.index,multiple).map(|value|((row.index,cell.col),value)))).collect();
    let refs=edits.iter().map(|(at,value)|(*at,value)).collect();
    let rich = crate::rich_text::patch(std::str::from_utf8(&part).map_err(|error|XlsxError::InvalidData(error.to_string()))?, &refs)?;
    patch_data_validations_xml(&rich, &[], &sheet.validations).map(String::into_bytes)
}

/// Serialize filter criteria for a table part or another spreadsheet container.
pub fn auto_filter_xml(filter: &AutoFilter) -> Result<String, XlsxError> {
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    write_auto_filter(&mut writer, Some(filter))?;
    let xml = String::from_utf8(writer.into_inner().into_inner()).map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    table_root_attributes(&xml, &[("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main".into())])
}

/// Writes an `<autoFilter>` block, or nothing when the filter has been taken
/// away.
///
/// A criterion is stated the way VBA states it  E`"apple"`, `">15"`  Eand is
/// written as a plain value list when it names values, or as a custom filter
/// when it compares.
fn write_auto_filter(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    filter: Option<&AutoFilter>,
) -> Result<(), XlsxError> {
    let Some(filter) = filter else {
        return Ok(());
    };
    let reference = format!(
        "{}{}:{}{}",
        col_to_letter(filter.start_col),
        filter.start_row,
        col_to_letter(filter.end_col),
        filter.end_row
    );
    let mut block = BytesStart::new("autoFilter");
    block.push_attribute(("ref", reference.as_str()));
    if filter.columns.is_empty() {
        return writer
            .write_event(Event::Empty(block))
            .map_err(|error| XlsxError::InvalidData(error.to_string()));
    }
    writer
        .write_event(Event::Start(block))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;

    for column in &filter.columns {
        let mut tag = BytesStart::new("filterColumn");
        tag.push_attribute(("colId", (column.field.saturating_sub(1)).to_string().as_str()));
        writer
            .write_event(Event::Start(tag))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;

        let comparing = column
            .criteria
            .iter()
            .any(|criterion| filter_comparison(criterion).is_some());
        if comparing {
            let mut customs = BytesStart::new("customFilters");
            if !column.either {
                customs.push_attribute(("and", "1"));
            }
            writer
                .write_event(Event::Start(customs))
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            for criterion in &column.criteria {
                let (operator, value) = match filter_comparison(criterion) {
                    Some(split) => split,
                    None => ("equal", criterion.as_str()),
                };
                let mut custom = BytesStart::new("customFilter");
                custom.push_attribute(("operator", operator));
                custom.push_attribute(("val", value));
                writer
                    .write_event(Event::Empty(custom))
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("customFilters")))
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        } else {
            writer
                .write_event(Event::Start(BytesStart::new("filters")))
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            for criterion in &column.criteria {
                let mut value = BytesStart::new("filter");
                value.push_attribute(("val", criterion.as_str()));
                writer
                    .write_event(Event::Empty(value))
                    .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("filters")))
                .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("filterColumn")))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }

    writer
        .write_event(Event::End(BytesEnd::new("autoFilter")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// Splits a comparing criterion into the operator OOXML names it by and the
/// value it compares against. A criterion that only names a value answers None.
fn filter_comparison(criterion: &str) -> Option<(&'static str, &str)> {
    for (prefix, operator) in [
        ("<>", "notEqual"),
        (">=", "greaterThanOrEqual"),
        ("<=", "lessThanOrEqual"),
        (">", "greaterThan"),
        ("<", "lessThan"),
    ] {
        if let Some(rest) = criterion.strip_prefix(prefix) {
            return Some((operator, rest));
        }
    }
    None
}

/// Writes a `<mergeCells>` block, or nothing at all when there is none left.
fn write_merges(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    merges: &[MergeCell],
) -> Result<(), XlsxError> {
    if merges.is_empty() {
        return Ok(());
    }
    let mut block = BytesStart::new("mergeCells");
    block.push_attribute(("count", merges.len().to_string().as_str()));
    writer
        .write_event(Event::Start(block))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    for merge in merges {
        let reference = format!(
            "{}{}:{}{}",
            col_to_letter(merge.start_col),
            merge.start_row,
            col_to_letter(merge.end_col),
            merge.end_row
        );
        let mut span = BytesStart::new("mergeCell");
        span.push_attribute(("ref", reference.as_str()));
        writer
            .write_event(Event::Empty(span))
            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    }
    writer
        .write_event(Event::End(BytesEnd::new("mergeCells")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// What a row should end up as. A field left `None` keeps whatever the sheet
/// already says, so hiding a row does not disturb its height.
#[derive(Clone, Copy, Default, PartialEq)]
struct RowWant {
    hidden: Option<bool>,
    /// `Some(None)` puts the row back on the sheet's automatic height.
    height: Option<Option<f32>>,
    /// Whether that height was chosen by hand. Excel writes `customHeight="1"`
    /// beside a height someone dragged to and leaves it off one it worked out
    /// itself, and the difference is whether it recomputes the row on open.
    /// Pinning a height Excel computed would fix a number that was a cache.
    pinned: bool,
    /// Some(None) removes the row format; None leaves it unchanged.
    style: Option<Option<u32>>,
}

/// What a column should end up as. A field left `None` keeps whatever the
/// sheet already says about it, so hiding a column does not disturb its width
/// and widening one does not un-hide it.
#[derive(Clone, Copy, Default, PartialEq)]
struct ColumnWant {
    hidden: Option<bool>,
    /// In characters of the workbook's default font, gutter included.
    width: Option<f32>,
    custom: Option<bool>,
    /// Some(None) removes the column format; None leaves it unchanged.
    style: Option<Option<u32>>,
}

/// Writes one `<col>` span, splitting it where a change covers only part of it.
///
/// A span reads `<col min="1" max="3" .../>` and covers three columns at once,
/// so hiding the middle one means writing three spans in its place.
fn write_col_span(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    start: &BytesStart<'_>,
    pending: &mut BTreeMap<u32, ColumnWant>,
) -> Result<(), XlsxError> {
    let min = get_attr(start, "min")
        .and_then(|value| value.parse::<u32>().ok())
        .unwrap_or(1);
    let max = get_attr(start, "max")
        .and_then(|value| value.parse::<u32>().ok())
        .unwrap_or(min);
    let was_hidden = matches!(get_attr(start, "hidden").as_deref(), Some("1") | Some("true"));
    let was_width = get_attr(start, "width");
    let was_custom = matches!(
        get_attr(start, "customWidth").as_deref(),
        Some("1") | Some("true")
    );

    // Group the columns of this span into runs that share an answer.
    let mut runs: Vec<(u32, u32, ColumnFace)> = Vec::new();
    for column in min..=max {
        let want = pending.remove(&(column - 1)).unwrap_or_default();
        let face = ColumnFace {
            hidden: want.hidden.unwrap_or(was_hidden),
            width: match want.width {
                Some(width) if width == 0.0 && want.custom == Some(false) => None,
                Some(width) => Some(width_text(width)),
                None => was_width.clone(),
            },
            resized: want.width.is_some(),
            custom: want.custom.unwrap_or(want.width.is_some() || was_custom),
            style: match want.style { Some(value) => value.map(|style| style.to_string()), None => get_attr(start, "style") },
            style_changed: want.style.is_some(),
        };
        match runs.last_mut() {
            Some((_, last, held)) if *held == face && *last + 1 == column => *last = column,
            _ => runs.push((column, column, face)),
        }
    }

    for (first, last, face) in runs {
        write_col(writer, Some(start), first, last, &face)?;
    }
    Ok(())
}

/// What one column should look like once everything has been applied.
#[derive(Clone, PartialEq)]
struct ColumnFace {
    hidden: bool,
    width: Option<String>,
    /// True when this run's width was set here rather than read off the file.
    resized: bool,
    custom: bool,
    style: Option<String>,
    style_changed: bool,
}

/// A width as the file states it.
///
/// The shortest form that still reads back as the same number, which is what
/// Excel writes: 10.625 for a column someone typed 10 into, and a bare `9` for
/// one left at the default. Rounding to a fixed number of places would write
/// 10.63, a different width.
fn width_text(width: f32) -> String {
    format!("{width}")
}

/// Writes one `<col>` span, keeping whatever attributes `from` carried that
/// this span does not decide for itself  Ea style, an outline level, a best
/// fit that still stands.
fn write_col(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    from: Option<&BytesStart<'_>>,
    first: u32,
    last: u32,
    face: &ColumnFace,
) -> Result<(), XlsxError> {
    let mut span = BytesStart::new("col");
    span.push_attribute(("min", first.to_string().as_str()));
    span.push_attribute(("max", last.to_string().as_str()));
    if let Some(width) = &face.width {
        span.push_attribute(("width", width.as_str()));
    }
    if let Some(start) = from {
        for attribute in start.attributes().flatten() {
            let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
            // Everything this span states for itself, and  Ewhen a width has
            // been set by hand  Ethe claim that the old width was a fit to
            // something. Excel drops bestFit on a resize, measured.
            if matches!(
                key.as_str(),
                "min" | "max" | "width" | "hidden" | "customWidth"
            ) || (face.resized && key == "bestFit") || (face.style_changed && key == "style")
            {
                continue;
            }
            let value = String::from_utf8_lossy(&attribute.value).into_owned();
            span.push_attribute((key.as_str(), value.as_str()));
        }
    }
    if face.style_changed {
        if let Some(style) = &face.style { span.push_attribute(("style", style.as_str())); }
    }
    if face.hidden {
        span.push_attribute(("hidden", "1"));
    }
    if face.custom {
        span.push_attribute(("customWidth", "1"));
    }
    writer
        .write_event(Event::Empty(span))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// Writes `<col>` spans for columns the sheet never described.
fn write_new_cols(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    pending: &mut BTreeMap<u32, ColumnWant>,
) -> Result<(), XlsxError> {
    let mut runs: Vec<(u32, u32, ColumnFace)> = Vec::new();
    for (column, want) in std::mem::take(pending) {
        let hidden = want.hidden.unwrap_or(false);
        if !hidden && want.width.is_none() && !matches!(want.style, Some(Some(_))) {
            continue;
        }
        let face = ColumnFace {
            hidden,
            width: want.width.filter(|width| *width != 0.0 || want.custom != Some(false)).map(width_text),
            resized: want.width.is_some(),
            custom: want.custom.unwrap_or(want.width.is_some()),
            style: want.style.flatten().map(|style| style.to_string()),
            style_changed: want.style.is_some(),
        };
        match runs.last_mut() {
            Some((_, last, held)) if *held == face && *last + 1 == column + 1 => *last = column + 1,
            _ => runs.push((column + 1, column + 1, face)),
        }
    }
    for (first, last, face) in runs { write_col(writer, None, first, last, &face)?; }
    Ok(())
}

/// Writes the `<pane>` that holds `rows` at the top and `cols` at the left.
///
/// Nothing is written when both are zero: an unfrozen sheet has no pane at
/// all rather than one saying it holds nothing.
fn write_pane(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    rows: u32,
    cols: u32,
) -> Result<(), XlsxError> {
    if rows == 0 && cols == 0 {
        return Ok(());
    }
    let mut pane = BytesStart::new("pane");
    if cols > 0 {
        pane.push_attribute(("xSplit", cols.to_string().as_str()));
    }
    if rows > 0 {
        pane.push_attribute(("ySplit", rows.to_string().as_str()));
    }
    // The first cell past the freeze, which is where the scrolling part starts.
    let corner = format!("{}{}", col_to_letter(cols), rows + 1);
    pane.push_attribute(("topLeftCell", corner.as_str()));
    // Which of the four corners the cursor lands in. With only rows held there
    // is no right-hand pane to be in, and with only columns no bottom one.
    pane.push_attribute((
        "activePane",
        match (rows > 0, cols > 0) {
            (true, true) => "bottomRight",
            (true, false) => "bottomLeft",
            (false, true) => "topRight",
            (false, false) => "topLeft",
        },
    ));
    pane.push_attribute(("state", "frozen"));
    writer
        .write_event(Event::Empty(pane))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// Writes a whole `<sheetViews>` block, for a sheet that had none.
fn write_sheet_views(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    rows: u32,
    cols: u32,
) -> Result<(), XlsxError> {
    writer
        .write_event(Event::Start(BytesStart::new("sheetViews")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    let mut view = BytesStart::new("sheetView");
    view.push_attribute(("workbookViewId", "0"));
    writer
        .write_event(Event::Start(view))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    write_pane(writer, rows, cols)?;
    writer
        .write_event(Event::End(BytesEnd::new("sheetView")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    writer
        .write_event(Event::End(BytesEnd::new("sheetViews")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn open_cols(writer: &mut Writer<Cursor<Vec<u8>>>) -> Result<(), XlsxError> {
    writer
        .write_event(Event::Start(BytesStart::new("cols")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

fn close_cols(writer: &mut Writer<Cursor<Vec<u8>>>) -> Result<(), XlsxError> {
    writer
        .write_event(Event::End(BytesEnd::new("cols")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// Copies a start tag, pointing it at a different entry in the style sheet.
/// Copies a row's start tag with whatever `want` decides applied to it.
///
/// Everything the row was carrying that this is not deciding  Eits span, its
/// own format, the descent Excel notes on it  Ecomes through untouched.
///
/// The flag beside the height is load-bearing. Excel writes `customHeight="1"`
/// on a row someone dragged to a size and leaves it off one it worked out
/// itself, and that is what decides whether it works the height out again the
/// next time the file is opened. Pinning a height Excel computed would fix a
/// number that was only ever a cache of what the content needed.
fn row_as_wanted(start: &BytesStart<'_>, want: &RowWant) -> BytesStart<'static> {
    let hidden = want.hidden.unwrap_or_else(|| {
        matches!(get_attr(start, "hidden").as_deref(), Some("1") | Some("true"))
    });
    let height = match want.height {
        Some(height) => height.map(|points| (format!("{points}"), want.pinned)),
        None => get_attr(start, "ht").map(|text| {
            let pinned = matches!(
                get_attr(start, "customHeight").as_deref(),
                Some("1") | Some("true")
            );
            (text, pinned)
        }),
    };
    let mut rebuilt = BytesStart::new("row");
    for attribute in start.attributes().flatten() {
        let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        if matches!(key.as_str(), "hidden" | "ht" | "customHeight")
            || (want.style.is_some() && matches!(key.as_str(), "s" | "customFormat")) {
            continue;
        }
        let value = String::from_utf8_lossy(&attribute.value).into_owned();
        rebuilt.push_attribute((key.as_str(), value.as_str()));
    }
    if let Some((points, pinned)) = &height {
        rebuilt.push_attribute(("ht", points.as_str()));
        if *pinned {
            rebuilt.push_attribute(("customHeight", "1"));
        }
    }
    if hidden {
        rebuilt.push_attribute(("hidden", "1"));
    }
    if let Some(Some(style)) = want.style {
        rebuilt.push_attribute(("s", style.to_string().as_str()));
        rebuilt.push_attribute(("customFormat", "1"));
    }
    rebuilt
}

fn with_style(start: &BytesStart<'_>, style: u32) -> BytesStart<'static> {
    let name = String::from_utf8_lossy(start.name().as_ref()).into_owned();
    let mut rewritten = BytesStart::new(name);
    for attribute in start.attributes().flatten() {
        let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
        if key == "s" {
            continue;
        }
        let value = String::from_utf8_lossy(&attribute.value).into_owned();
        rewritten.push_attribute((key.as_str(), value.as_str()));
    }
    rewritten.push_attribute(("s", style.to_string().as_str()));
    rewritten
}

/// Copies a start tag, replacing whatever it said about being hidden.
/// Writes a row the file does not have yet.
///
/// There is nothing here to preserve  Ethe row is new  Eso it wears whatever
/// `want` says and nothing else. Copying an EXISTING row's start tag, which
/// has to keep the height and span and hidden flag it was already carrying, is
/// `row_as_wanted`'s job.
fn write_inserted_row(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    row: u32,
    cells: BTreeMap<u32, CellEditValue>,
    want: Option<RowWant>,
    styles: &BTreeMap<(u32, u32), u32>,
) -> Result<(), XlsxError> {
    let row_text = row.to_string();
    let mut bare = BytesStart::new("row");
    bare.push_attribute(("r", row_text.as_str()));
    let row_start = match want {
        Some(want) => row_as_wanted(&bare, &want),
        None => bare,
    };
    writer
        .write_event(Event::Start(row_start))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
    for (col, value) in cells {
        write_cell_value(writer, row, col, &value, styles)?;
    }
    writer
        .write_event(Event::End(BytesEnd::new("row")))
        .map_err(|error| XlsxError::InvalidData(error.to_string()))
}

/// Patch worksheet XML, replacing or inserting cells at specified positions.
/// Each edit is written with its corresponding OOXML scalar type.
/// Everything one worksheet has to change.
struct SheetEdits<'a> {
    cells: &'a HashMap<(u32, u32), &'a CellEditValue>,
    /// One-based row -> what it should become.
    rows: &'a BTreeMap<u32, RowWant>,
    /// Zero-based column -> what it should become.
    cols: &'a BTreeMap<u32, ColumnWant>,
    /// Every merge the sheet should end up with, when they are being replaced.
    merges: Option<&'a [MergeCell]>,
    /// (one-based row, zero-based column) -> the cellXfs index it should carry.
    styles: &'a BTreeMap<(u32, u32), u32>,
    /// The filter the sheet should end up under, when it is being replaced.
    /// `Some(None)` takes an existing filter away.
    filter: Option<&'a Option<AutoFilter>>,
    /// The complete cell-link set when this sheet's hyperlinks changed.
    hyperlinks: Option<&'a [Hyperlink]>,
    /// Relationship ids parallel to `hyperlinks`; internal destinations
    /// have no relationship id.
    hyperlink_ids: &'a [Option<String>],
    /// The rows and columns to hold in view, when they are being set. Both
    /// zero takes the freeze away.
    panes: Option<(u32, u32)>,
}

/// Emit each missing row once, in order, merging cell and row-format edits.
fn flush_inserted_rows(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    before: Option<u32>,
    pending: &mut BTreeMap<u32, BTreeMap<u32, CellEditValue>>,
    pending_rows: &mut std::collections::BTreeSet<u32>,
    edits: &SheetEdits<'_>,
) -> Result<(), XlsxError> {
    let rows: std::collections::BTreeSet<u32> = pending.keys().chain(pending_rows.iter())
        .copied().filter(|row| before.is_none_or(|limit| *row < limit)).collect();
    for row in rows {
        let cells = pending.remove(&row).unwrap_or_default();
        pending_rows.remove(&row);
        let want = edits.rows.get(&row).copied();
        if !cells.is_empty() || want.is_some_and(|want| want.hidden == Some(true) || want.height.is_some() || matches!(want.style, Some(Some(_)))) {
            write_inserted_row(writer, row, cells, want, edits.styles)?;
        }
    }
    Ok(())
}

/// Set, change or remove a worksheet part's `<sheetPr><tabColor>`. `colour` is
/// the RRGGBB the tab should wear (written with Excel's leading FF alpha), or
/// `None` to take any colour away. Done as string surgery so the streaming
/// cell patcher stays untouched; a worksheet part never holds a `<sheetPr` or a
/// raw `>` anywhere but in this one element's start tag.
fn set_tab_color_in_worksheet(xml: &str, colour: &Option<String>) -> String {
    let element = colour.as_ref().map(|hex| format!("<tabColor rgb=\"FF{hex}\"/>"));
    if let Some(pr_at) = xml.find("<sheetPr") {
        let Some(rel_end) = xml[pr_at..].find('>') else {
            return xml.to_string();
        };
        let tag_end = pr_at + rel_end;
        let self_closing = xml.as_bytes()[tag_end - 1] == b'/';
        if self_closing {
            match &element {
                Some(el) => format!("{}>{el}</sheetPr>{}", &xml[..tag_end - 1], &xml[tag_end + 1..]),
                None => xml.to_string(),
            }
        } else {
            let body_start = tag_end + 1;
            let Some(close_rel) = xml[body_start..].find("</sheetPr>") else {
                return xml.to_string();
            };
            let close_at = body_start + close_rel;
            let body = remove_tab_color(&xml[body_start..close_at]);
            let inner = match &element {
                Some(el) => format!("{el}{body}"),
                None => body,
            };
            format!("{}{}{}", &xml[..body_start], inner, &xml[close_at..])
        }
    } else {
        match &element {
            Some(el) => match xml.find("<worksheet").and_then(|ws| {
                xml[ws..].find('>').map(|rel| ws + rel + 1)
            }) {
                Some(at) => format!("{}<sheetPr>{el}</sheetPr>{}", &xml[..at], &xml[at..]),
                None => xml.to_string(),
            },
            None => xml.to_string(),
        }
    }
}

/// A `<sheetPr>` body with any `<tabColor>` taken out.
fn remove_tab_color(body: &str) -> String {
    let Some(at) = body.find("<tabColor") else {
        return body.to_string();
    };
    if let Some(rel) = body[at..].find("/>") {
        return format!("{}{}", &body[..at], &body[at + rel + 2..]);
    }
    if let Some(rel) = body[at..].find("</tabColor>") {
        return format!("{}{}", &body[..at], &body[at + rel + "</tabColor>".len()..]);
    }
    body.to_string()
}

fn patch_worksheet_xml(xml: &str, sheet_edits: &SheetEdits<'_>) -> Result<String, XlsxError> {
    let edits = sheet_edits.cells;
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut pending: BTreeMap<u32, BTreeMap<u32, CellEditValue>> = BTreeMap::new();
    for (&(row, col), value) in edits {
        pending
            .entry(row)
            .or_default()
            .insert(col, (*value).clone());
    }

    let mut pending_rows: std::collections::BTreeSet<u32> =
        sheet_edits.rows.keys().copied().collect();
    let mut pending_cols: BTreeMap<u32, ColumnWant> = sheet_edits.cols.clone();
    let mut seen_cols = false;
    let mut skip_merges_depth = 0_u32;
    // Whether the sheet already describes its merges decides where the new
    // block goes: in place of the old one, or straight after the data.
    let has_merge_block = xml.contains("<mergeCells");
    let mut wrote_merges = false;
    let has_filter_block = xml.contains("<autoFilter");
    let mut wrote_filter = false;
    let mut skip_filter_depth = 0_u32;
    // Whether we are inside a `<sheetView>` whose panes are being replaced.
    // Its own attributes come through untouched  Ea sheet with the gridlines
    // turned off must not have them turned back on  Eand only its `<pane>` and
    // `<selection>` children are dropped, since a selection names a corner
    // that may no longer exist.
    let mut in_view = false;
    let mut wrote_views = false;

    let mut current_row: u32 = 0;
    let mut in_row = false;
    let mut in_cell = false;
    let mut cell_col: u32;
    let mut cell_row: u32;
    let mut in_value = false;
    let mut value_written = false;
    let mut current_edit: Option<CellEditValue> = None;
    let mut skip_value_text = false;
    let mut skip_replaced_content_depth = 0_u32;

    loop {
        match reader.read_event().map_err(XlsxError::Xml)? {
            Event::Eof => break,
            Event::Start(ref e) => {
                if skip_filter_depth > 0 {
                    skip_filter_depth += 1;
                    continue;
                }
                if skip_merges_depth > 0 {
                    skip_merges_depth += 1;
                    continue;
                }
                if skip_replaced_content_depth > 0 {
                    skip_replaced_content_depth += 1;
                    continue;
                }
                let name = local_name(e.name().as_ref());
                if in_cell && current_edit.is_some() && (name == "f" || name == "is") {
                    skip_replaced_content_depth = 1;
                    continue;
                }
                if name == "cols" {
                    seen_cols = true;
                }
                if name == "sheetView" && sheet_edits.panes.is_some() {
                    writer
                        .write_event(Event::Start(e.clone()))
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    if let Some((rows, cols)) = sheet_edits.panes {
                        write_pane(&mut writer, rows, cols)?;
                    }
                    in_view = true;
                    wrote_views = true;
                    continue;
                }
                if in_view && matches!(name.as_str(), "pane" | "selection") {
                    skip_replaced_content_depth = 1;
                    continue;
                }
                if name == "autoFilter" && sheet_edits.filter.is_some() {
                    skip_filter_depth = 1;
                    continue;
                }
                if name == "mergeCells" && sheet_edits.merges.is_some() {
                    skip_merges_depth = 1;
                    continue;
                }
                if !wrote_views
                    && sheet_edits.panes.is_some()
                    && matches!(name.as_str(), "sheetFormatPr" | "cols" | "sheetData")
                {
                    if let Some((rows, cols)) = sheet_edits.panes {
                        write_sheet_views(&mut writer, rows, cols)?;
                    }
                    wrote_views = true;
                }
                if name == "sheetData" && !seen_cols && !pending_cols.is_empty() {
                    open_cols(&mut writer)?;
                    write_new_cols(&mut writer, &mut pending_cols)?;
                    close_cols(&mut writer)?;
                    seen_cols = true;
                }
                match name.as_str() {
                    "row" => {
                        let next_row = get_attr(e, "r")
                            .and_then(|value| value.parse().ok())
                            .unwrap_or(0);
                        flush_inserted_rows(&mut writer, Some(next_row), &mut pending, &mut pending_rows, sheet_edits)?;
                        in_row = true;
                        current_row = next_row;
                        pending_rows.remove(&next_row);
                        if let Some(hidden) = sheet_edits.rows.get(&next_row) {
                            writer
                                .write_event(Event::Start(row_as_wanted(e, hidden)))
                                .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            continue;
                        }
                    }
                    "c" if in_row => {
                        in_cell = true;
                        // Whether this cell has been GIVEN its new value yet,
                        // which is not the same question as whether we are
                        // inside its `<v>` right now. Asking `in_value` at
                        // `</c>`  Eby which point `</v>` has already cleared
                        // it  Eput a SECOND `<v>` after the one just written,
                        // and a cell holding two of them is a workbook Excel
                        // refuses to open. It is the same conflation that
                        // wrote a docx run's text once per `<w:t>`.
                        value_written = false;
                        let cell_ref = get_attr(e, "r").unwrap_or_default();
                        let (col, row) = crate::parser::parse_cell_ref(&cell_ref);
                        cell_col = col;
                        cell_row = if row > 0 { row + 1 } else { current_row };

                        if let Some(cells) = pending.get_mut(&cell_row) {
                            let missing_cols: Vec<u32> =
                                cells.range(..cell_col).map(|(&col, _)| col).collect();
                            for col in missing_cols {
                                if let Some(value) = cells.remove(&col) {
                                    write_cell_value(
                                        &mut writer,
                                        cell_row,
                                        col,
                                        &value,
                                        sheet_edits.styles,
                                    )?;
                                }
                            }
                            current_edit = cells.remove(&cell_col);
                        } else {
                            current_edit = None;
                        }

                        let restyle = sheet_edits.styles.get(&(cell_row, cell_col)).copied();
                        if current_edit.is_none() {
                            if let Some(style) = restyle {
                                writer
                                    .write_event(Event::Start(with_style(e, style)))
                                    .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                                continue;
                            }
                        }
                        if let Some(value) = &current_edit {
                            // Rewrite the cell type while preserving its reference and style.
                            let mut new_start = BytesStart::new("c");
                            for attr in e.attributes().flatten() {
                                let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                                if key == "t" || (key == "s" && restyle.is_some()) {
                                    continue;
                                }
                                new_start.push_attribute((key, std::str::from_utf8(&attr.value).unwrap_or("")));
                            }
                            if let Some(style) = restyle {
                                new_start.push_attribute(("s", style.to_string().as_str()));
                            }
                            if let Some(cell_type) = value.cell_type() {
                                new_start.push_attribute(("t", cell_type));
                            }
                            writer.write_event(Event::Start(new_start)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            if let Some(formula) = value.formula_units() {
                                write_formula(&mut writer, &formula, value.array_reach())?;
                            }
                        } else {
                            writer.write_event(Event::Start(e.clone())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        }
                        continue;
                    }
                    "v" if in_cell && current_edit.is_some() => {
                        in_value = true;
                        value_written = true;
                        skip_value_text = true;
                        // Write <v> start
                        writer.write_event(Event::Start(e.clone())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        // Write new value text
                        if let Some(text) = current_edit
                            .as_ref()
                            .map(CellEditValue::value_text)
                            .transpose()?
                            .flatten()
                        {
                            writer.write_event(Event::Text(BytesText::new(&text))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        }
                        continue;
                    }
                    _ => {}
                }
                writer.write_event(Event::Start(e.clone())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            }
            Event::End(ref e) => {
                if skip_filter_depth > 0 {
                    skip_filter_depth -= 1;
                    if skip_filter_depth == 0 {
                        if let Some(filter) = sheet_edits.filter {
                            write_auto_filter(&mut writer, filter.as_ref())?;
                            wrote_filter = true;
                        }
                    }
                    continue;
                }
                if skip_merges_depth > 0 {
                    skip_merges_depth -= 1;
                    if skip_merges_depth == 0 {
                        if let Some(merges) = sheet_edits.merges {
                            write_merges(&mut writer, merges)?;
                            wrote_merges = true;
                        }
                    }
                    continue;
                }
                if skip_replaced_content_depth > 0 {
                    skip_replaced_content_depth -= 1;
                    continue;
                }
                let name = local_name(e.name().as_ref());
                if name == "sheetView" {
                    in_view = false;
                }
                if name == "cols" {
                    write_new_cols(&mut writer, &mut pending_cols)?;
                }
                match name.as_str() {
                    "row" => {
                        if let Some(cells) = pending.remove(&current_row) {
                            for (col, value) in cells {
                                write_cell_value(
                                    &mut writer,
                                    current_row,
                                    col,
                                    &value,
                                        sheet_edits.styles,
                                )?;
                            }
                        }
                        in_row = false;
                    }
                    "c" => {
                        if in_cell && current_edit.is_some() {
                            // A cell that carried no `<v>` of its own still
                            // needs one to hold the new value.
                            if !value_written {
                                if let Some(text) = current_edit
                                    .as_ref()
                                    .map(CellEditValue::value_text)
                                    .transpose()?
                                    .flatten()
                                {
                                    writer.write_event(Event::Start(BytesStart::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                                    writer.write_event(Event::Text(BytesText::new(&text))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                                    writer.write_event(Event::End(BytesEnd::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                                }
                            }
                        }
                        in_cell = false;
                        in_value = false;
                        current_edit = None;
                        skip_value_text = false;
                    }
                    "v" => {
                        in_value = false;
                        skip_value_text = false;
                    }
                    "sheetData" => {
                        flush_inserted_rows(&mut writer, None, &mut pending, &mut pending_rows, sheet_edits)?;
                    }
                    _ => {}
                }
                writer.write_event(Event::End(e.clone())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                // A filter sits between the data and the merges.
                if name == "sheetData" && !wrote_filter && !has_filter_block {
                    if let Some(filter) = sheet_edits.filter {
                        write_auto_filter(&mut writer, filter.as_ref())?;
                        wrote_filter = true;
                    }
                }
                // A sheet that never had a mergeCells element needs one, and it
                // belongs directly after the data.
                if name == "sheetData" && !wrote_merges && !has_merge_block {
                    if let Some(merges) = sheet_edits.merges {
                        write_merges(&mut writer, merges)?;
                        wrote_merges = true;
                    }
                }
            }
            Event::Text(ref e) => {
                if skip_replaced_content_depth > 0 || (skip_value_text && in_value) {
                    // Already wrote the new value, skip the original
                    continue;
                }
                writer.write_event(Event::Text(e.clone())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            }
            Event::Empty(ref e) => {
                if skip_replaced_content_depth > 0 {
                    continue;
                }
                let name = local_name(e.name().as_ref());
                if in_cell && current_edit.is_some() && matches!(name.as_str(), "f" | "is") {
                    continue;
                }
                if name == "v" && in_cell && current_edit.is_some() {
                    value_written = true;
                    if let Some(text) = current_edit.as_ref().map(CellEditValue::value_text).transpose()?.flatten() {
                        writer.write_event(Event::Start(BytesStart::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        writer.write_event(Event::Text(BytesText::new(&text))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        writer.write_event(Event::End(BytesEnd::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    }
                    continue;
                }
                if name == "sheetView" && sheet_edits.panes.is_some() {
                    // An unfrozen sheet says `<sheetView workbookViewId="0"/>`
                    // with nothing inside it, so it has to be opened up before
                    // a pane can be put in.
                    writer
                        .write_event(Event::Start(e.clone()))
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    if let Some((rows, cols)) = sheet_edits.panes {
                        write_pane(&mut writer, rows, cols)?;
                    }
                    writer
                        .write_event(Event::End(BytesEnd::new("sheetView")))
                        .map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                    wrote_views = true;
                    continue;
                }
                if in_view && matches!(name.as_str(), "pane" | "selection") {
                    continue;
                }
                if name == "col" {
                    seen_cols = true;
                    write_col_span(&mut writer, e, &mut pending_cols)?;
                    continue;
                }
                if name == "autoFilter" && sheet_edits.filter.is_some() {
                    if let Some(filter) = sheet_edits.filter {
                        write_auto_filter(&mut writer, filter.as_ref())?;
                        wrote_filter = true;
                    }
                    continue;
                }
                if name == "mergeCells" && sheet_edits.merges.is_some() {
                    if let Some(merges) = sheet_edits.merges {
                        write_merges(&mut writer, merges)?;
                        wrote_merges = true;
                    }
                    continue;
                }
                if in_cell && current_edit.is_some() && (name == "f" || name == "is") {
                    continue;
                }
                if !wrote_views
                    && sheet_edits.panes.is_some()
                    && matches!(name.as_str(), "sheetFormatPr" | "cols" | "sheetData")
                {
                    if let Some((rows, cols)) = sheet_edits.panes {
                        write_sheet_views(&mut writer, rows, cols)?;
                    }
                    wrote_views = true;
                }
                if name == "sheetData" && !seen_cols && !pending_cols.is_empty() {
                    open_cols(&mut writer)?;
                    write_new_cols(&mut writer, &mut pending_cols)?;
                    close_cols(&mut writer)?;
                    seen_cols = true;
                }
                if name == "sheetData" && !pending.is_empty() {
                    writer
                        .write_event(Event::Start(e.clone()))
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    flush_inserted_rows(&mut writer, None, &mut pending, &mut pending_rows, sheet_edits)?;
                    writer
                        .write_event(Event::End(BytesEnd::new("sheetData")))
                        .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                    continue;
                }
                if name == "row" {
                    let row_num = get_attr(e, "r")
                        .and_then(|value| value.parse().ok())
                        .unwrap_or(0);
                    flush_inserted_rows(&mut writer, Some(row_num), &mut pending, &mut pending_rows, sheet_edits)?;
                    pending_rows.remove(&row_num);
                    if let Some(cells) = pending.remove(&row_num) {
                        let row_start = sheet_edits.rows.get(&row_num)
                            .map(|want| row_as_wanted(e, want)).unwrap_or_else(|| e.clone());
                        writer.write_event(Event::Start(row_start)).map_err(|error| {
                            XlsxError::InvalidData(error.to_string())
                        })?;
                        for (col, value) in cells {
                            write_cell_value(&mut writer, row_num, col, &value, sheet_edits.styles)?;
                        }
                        writer.write_event(Event::End(BytesEnd::new("row"))).map_err(
                            |error| XlsxError::InvalidData(error.to_string()),
                        )?;
                        continue;
                    }
                    if let Some(want) = sheet_edits.rows.get(&row_num) {
                        writer.write_event(Event::Empty(row_as_wanted(e, want)))
                            .map_err(|error| XlsxError::InvalidData(error.to_string()))?;
                        continue;
                    }
                }
                if name == "c" && in_row {
                    let cell_ref = get_attr(e, "r").unwrap_or_default();
                    let (col, row) = crate::parser::parse_cell_ref(&cell_ref);
                    let row_num = if row > 0 { row + 1 } else { current_row };

                    let mut edit = None;
                    if let Some(cells) = pending.get_mut(&row_num) {
                        let missing_cols: Vec<u32> =
                            cells.range(..col).map(|(&column, _)| column).collect();
                        for missing_col in missing_cols {
                            if let Some(value) = cells.remove(&missing_col) {
                                write_cell_value(
                                    &mut writer,
                                    row_num,
                                    missing_col,
                                    &value,
                                    sheet_edits.styles,
                                )?;
                            }
                        }
                        edit = cells.remove(&col);
                    }

                    if let Some(value) = edit {
                        // Convert the empty cell while preserving its reference and style.
                        let mut new_start = BytesStart::new("c");
                        let restyle = sheet_edits.styles.get(&(row_num, col));
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "t" || (key == "s" && restyle.is_some()) { continue; }
                            new_start.push_attribute((key, std::str::from_utf8(&attr.value).unwrap_or("")));
                        }
                        if let Some(style) = restyle {
                            new_start.push_attribute(("s", style.to_string().as_str()));
                        }
                        if let Some(cell_type) = value.cell_type() {
                            new_start.push_attribute(("t", cell_type));
                        }
                        if let Some(formula) = value.formula_units() {
                            writer.write_event(Event::Start(new_start)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            write_formula(&mut writer, &formula, value.array_reach())?;
                            if let Some(text) = value.value_text()? {
                                writer.write_event(Event::Start(BytesStart::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                                writer.write_event(Event::Text(BytesText::new(&text))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                                writer.write_event(Event::End(BytesEnd::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            }
                            writer.write_event(Event::End(BytesEnd::new("c"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        } else if let Some(text) = value.value_text()? {
                            writer.write_event(Event::Start(new_start)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            writer.write_event(Event::Start(BytesStart::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            writer.write_event(Event::Text(BytesText::new(&text))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            writer.write_event(Event::End(BytesEnd::new("v"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                            writer.write_event(Event::End(BytesEnd::new("c"))).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        } else {
                            writer.write_event(Event::Empty(new_start)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        }
                        continue;
                    }
                    if let Some(style) = sheet_edits.styles.get(&(row_num, col)) {
                        let mut restyled = BytesStart::new("c");
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() != b"s" {
                                restyled.push_attribute(attr);
                            }
                        }
                        restyled.push_attribute(("s", style.to_string().as_str()));
                        writer.write_event(Event::Empty(restyled)).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
                        continue;
                    }
                }
                writer.write_event(Event::Empty(e.clone())).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            }
            event => {
                if skip_replaced_content_depth > 0 {
                    continue;
                }
                writer.write_event(event).map_err(|e| XlsxError::InvalidData(e.to_string()))?;
            }
        }
    }

    let result = writer.into_inner().into_inner();
    let result = String::from_utf8(result).map_err(|_| XlsxError::InvalidData("UTF-8 error".to_string()))?;
    crate::rich_text::patch(&result, sheet_edits.cells)
}

#[cfg(test)]
mod tests {
    #[test]
    fn worksheet_drawings_are_direct_children_even_with_nested_extensions() {
        let xml = "<worksheet><conditionalFormatting><cfRule><extLst/></cfRule></conditionalFormatting><pageMargins/><extLst/></worksheet>";
        let with_note = super::set_legacy_drawing(xml, "rId2").unwrap();
        assert!(with_note.contains("</conditionalFormatting><pageMargins/><legacyDrawing "));
        assert!(with_note.contains("r:id=\"rId2\"/><extLst/></worksheet>"));
        let with_drawing = super::set_drawing(&with_note, "rId3").unwrap();
        assert!(with_drawing.contains("r:id=\"rId3\"/><legacyDrawing "));
        let replaced = super::set_legacy_drawing(&with_drawing, "rId4").unwrap();
        assert!(replaced.contains("r:id=\"rId4\"/><extLst/></worksheet>"));
        assert_eq!(replaced.matches("<legacyDrawing ").count(), 1);
    }
    #[test]
    fn extended_icon_edits_preserve_custom_icons_and_escape_formulas() {
        let raw = r#"<x14:conditionalFormatting><x14:cfRule type="iconSet" id="{A}"><x14:iconSet iconSet="3Stars" custom="1"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo><x14:cfIcon iconSet="3Flags" iconId="1"/></x14:iconSet></x14:cfRule><xm:sqref>A2:A6</xm:sqref></x14:conditionalFormatting>"#;
        let visual = r#"<iconSet iconSet="3Stars" reverse="1"><cfvo type="percent" val="0"/><cfvo type="formula" val="$B$1&amp;$C$1" gte="0"/><cfvo type="percent" val="67"/></iconSet>"#;
        let result = super::update_extended_icon_visual(raw, visual).unwrap();
        assert!(result.contains(r#"<x14:iconSet iconSet="3Stars" custom="1" reverse="1">"#));
        assert!(result.contains(r#"<x14:cfvo type="formula" gte="0"><xm:f>$B$1&amp;$C$1</xm:f></x14:cfvo>"#));
        assert!(result.contains(r#"<x14:cfIcon iconSet="3Flags" iconId="1"/>"#));
        assert!(super::update_extended_icon_visual(raw, r#"<iconSet><cfvo type="num" val="1"/></iconSet>"#).is_err());
    }

    #[test]
    fn extended_conditional_range_follows_the_base_rule() {
        let xml = r#"<worksheet><conditionalFormatting sqref="B2:B6"><cfRule type="dataBar" priority="2"><dataBar><color rgb="FF638EC6"/></dataBar><extLst><ext><x14:id>{ABC}</x14:id></ext></extLst></cfRule></conditionalFormatting><extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="{ABC}"><x14:dataBar axisPosition="none"/></x14:cfRule><xm:sqref>B2:B6</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext><ext uri="other"><payload/></ext></extLst></worksheet>"#;
        let bytes = include_bytes!("../../../tests/fixtures/conditional/cell_is.xlsx");
        let mut before = crate::parse_xlsx(bytes).unwrap().sheets[0].conditional_rules[0].clone();
        before.kind = "dataBar".into(); before.operator = None; before.formulas.clear(); before.style = None;
        before.priority = 2; before.origin_priority = Some(2); before.ranges = vec![(1, 1, 5, 1)];
        let mut after = before.clone(); after.ranges = vec![(1, 1, 6, 1)];
        let output = super::patch_conditional_rules_xml(xml, &[before.clone()], &[after.clone()], &std::collections::HashMap::new()).unwrap();
        assert!(output.contains(r#"sqref="B2:B7""#));
        assert!(output.contains("<xm:sqref>B2:B7</xm:sqref>"));
        assert!(output.contains(r#"<x14:dataBar axisPosition="none"/>"#));
        assert!(output.contains(r#"<color rgb="FF638EC6"/>"#));
        let missing = xml.replace("id=\"{ABC}\"", "id=\"{OTHER}\"");
        assert!(super::patch_conditional_rules_xml(&missing, &[before.clone()], &[after], &std::collections::HashMap::new()).is_err());
        let deleted = super::patch_conditional_rules_xml(xml, &[before], &[], &std::collections::HashMap::new()).unwrap();
        assert!(!deleted.contains("conditionalFormatting"));
        assert!(deleted.contains(r#"<ext uri="other"><payload/></ext>"#));
    }

    #[test]
    fn conditional_rule_removal_preserves_surviving_rules_and_styles() {
        let bytes = include_bytes!("../../../tests/fixtures/conditional/cell_is.xlsx");
        let mut workbook = crate::parse_xlsx(bytes).unwrap();
        assert_eq!(workbook.sheets[0].conditional_rules.len(), 2);
        let surviving = workbook.sheets[0].conditional_rules[1].clone();
        workbook.sheets[0].conditional_rules.remove(0);
        let mut editor = super::XlsxEditor::new(bytes).unwrap();
        editor.apply_workbook(&workbook).unwrap();
        let saved = editor.save().unwrap();
        let reopened = crate::parse_xlsx(&saved).unwrap();
        assert_eq!(reopened.sheets[0].conditional_rules, vec![surviving]);
        let original_styles = oxidocs_common::archive::OoxmlArchive::new(bytes).unwrap().read_part("xl/styles.xml").unwrap();
        let saved_styles = oxidocs_common::archive::OoxmlArchive::new(&saved).unwrap().read_part("xl/styles.xml").unwrap();
        assert_eq!(original_styles, saved_styles);
    }

    #[test]
    fn conditional_rule_edits_fail_instead_of_being_silently_lost() {
        let bytes = include_bytes!("../../../tests/fixtures/conditional/cell_is.xlsx");
        let mut workbook = crate::parse_xlsx(bytes).unwrap();
        workbook.sheets[0].conditional_rules[0].text = Some("unmodelled change".into());
        let mut editor = super::XlsxEditor::new(bytes).unwrap();
        editor.apply_workbook(&workbook).unwrap();
        assert!(editor.save().is_err());
    }

    #[test]
    fn conditional_ranges_split_without_losing_other_rules() {
        let bytes = include_bytes!("../../../tests/fixtures/conditional/cell_is.xlsx");
        let mut workbook = crate::parse_xlsx(bytes).unwrap();
        let unchanged = workbook.sheets[0].conditional_rules[1].clone();
        workbook.sheets[0].conditional_rules[0].ranges = vec![(1, 0, 2, 0), (4, 0, 7, 0)];
        let mut editor = super::XlsxEditor::new(bytes).unwrap();
        editor.apply_workbook(&workbook).unwrap();
        let reopened = crate::parse_xlsx(&editor.save().unwrap()).unwrap();
        assert_eq!(reopened.sheets[0].conditional_rules, workbook.sheets[0].conditional_rules);
        assert_eq!(reopened.sheets[0].conditional_rules[1], unchanged);
    }

    #[test]
    fn conditional_settings_keep_unmodelled_xml_attributes() {
        let bytes = include_bytes!("../../../tests/fixtures/conditional/cell_is.xlsx");
        let workbook = crate::parse_xlsx(bytes).unwrap();
        let mut before = workbook.sheets[0].conditional_rules[0].clone();
        before.kind = "aboveAverage".into(); before.operator = None; before.formulas.clear(); before.style = None;
        let mut after = before.clone();
        after.above_average = false; after.standard_deviations = Some(2); after.stop_if_true = true;
        let xml = r#"<worksheet><conditionalFormatting sqref="A2:A8"><cfRule type="aboveAverage" priority="1" custom="a&amp;b"/></conditionalFormatting></worksheet>"#;
        let patched = super::patch_conditional_rules_xml(xml, &[before], &[after], &std::collections::HashMap::new()).unwrap();
        assert!(patched.contains(r#"custom="a&amp;b""#), "{patched}");
        assert!(patched.contains(r#"aboveAverage="0""#), "{patched}");
        assert!(patched.contains(r#"stdDev="2""#), "{patched}");
        assert!(patched.contains(r#"stopIfTrue="1""#), "{patched}");
    }

    #[test]
    fn differential_style_edit_keeps_unmodelled_children() {
        let raw = r#"<dxf><font><name val="Example"/><b/><extLst><ext uri="x"/></extLst></font><fill><patternFill patternType="solid"><fgColor rgb="FFABCDEF"/><bgColor rgb="FF112233"/></patternFill></fill><border><left style="thin"/></border></dxf>"#;
        let mut before = crate::ir::DiffStyle::default();
        before.source_xml = Some(raw.into()); before.bold = Some(true); before.bg_color = Some("112233".into());
        let mut after = before.clone(); after.bold = Some(false); after.bg_color = Some("FF0000".into());
        let patched = super::patch_differential_style(raw, &before, &after).unwrap();
        for expected in [r#"<name val="Example"/>"#, r#"<ext uri="x"/>"#, r#"<fgColor rgb="FFABCDEF"/>"#, r#"<left style="thin"/>"#, r#"<b val="0"/>"#, r#"<bgColor rgb="FFFF0000"/>"#] {
            assert!(patched.contains(expected), "missing {expected}: {patched}");
        }
    }
    use super::*;

    #[test]
    fn moving_a_validation_preserves_unmodelled_rule_metadata() {
        let bytes = include_bytes!("../tests/fixtures/unsupported_bits.xlsx");
        let workbook = crate::parse_xlsx(bytes).unwrap();
        let before = &workbook.sheets[0].validations;
        let mut after = before.clone();
        after[0].ranges = vec![(1, 1, 1, 1)];
        let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut xml = String::new();
        archive.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut xml).unwrap();
        let patched = patch_data_validations_xml(&xml, before, &after).unwrap();
        assert!(patched.contains("sqref=\"B2\""));
        assert!(patched.contains("xr:uid=\"{1294A447-6A7E-4403-89DB-A76DEB859ED2}\""));
        assert!(patched.contains("<formula1>\"a,b,c\"</formula1>"));
    }

    #[test]
    fn drawing_order_edit_reorders_existing_worksheet_anchors() {
        let bytes = include_bytes!("../../../tests/fixtures/vba/scale_picture_original_ole.xlsx");
        let original = crate::parse_xlsx(bytes).unwrap();
        let mut edited = original.clone();
        edited.sheets[0].drawings.swap(0, 1);
        let mut editor = XlsxEditor::new(bytes).unwrap();
        editor.apply_workbook(&edited).unwrap();
        assert_eq!(editor.drawing_order_edits.get(&0), Some(&vec!["Object 2".to_string(), "Object 1".to_string()]));
        let saved = editor.save().unwrap();
        let drawing_xml = OoxmlArchive::new(&saved).unwrap().read_part("xl/drawings/drawing1.xml").unwrap();
        assert!(drawing_xml.find("name=\"Object 2\"").unwrap() < drawing_xml.find("name=\"Object 1\"").unwrap());
        let reopened = crate::parse_xlsx(&saved).unwrap();
        assert_eq!(reopened.sheets[0].drawings[0].name.as_deref(), Some("Object 2"));
        assert_eq!(reopened.sheets[0].drawings[1].name.as_deref(), Some("Object 1"));
    }

    fn workbook_with_line_shape() -> Vec<u8> {
        let parts = [
            ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#),
            ("xl/workbook.xml", r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#),
            ("xl/_rels/workbook.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#),
            ("xl/worksheets/sheet1.xml", r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rId1"/></worksheet>"#),
            ("xl/worksheets/_rels/sheet1.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#),
            ("xl/drawings/drawing1.xml", r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Patterned"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln><a:solidFill><a:srgbClr val="111111"/></a:solidFill><a:prstDash val="dash"/></a:ln></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#),
        ];
        let mut bytes = Vec::new();
        let mut writer = ZipWriter::new(Cursor::new(&mut bytes));
        for (name, xml) in parts {
            writer.start_file(name, SimpleFileOptions::default()).unwrap();
            writer.write_all(xml.as_bytes()).unwrap();
        }
        writer.finish().unwrap();
        bytes
    }

    fn simple_added_shape(name: &str) -> crate::ir::Drawing {
        crate::ir::Drawing {
            name: Some(name.into()),
            ole_prog_id: None,
            ole_source_name: None,
            ole_clone_position: None,
            ole_embedded_package: None,
            ole_package_is_empty: None,
            ole_clone_visible: None,
            ole_placement: None,
            ole_type: None,
            ole_display_as_icon: None,
            ole_auto_load: None,
            ole_enabled: None,
            ole_locked: None,
            ole_print_object: None,
            ole_shadow: None,
            ole_on_action: None,
            ole_border_color: None,
            ole_border_weight: None,
            ole_interior_color: None,
            from: crate::ir::Anchor { col: 1, col_off: 12_700, row: 2, row_off: 25_400 },
            to: Some(crate::ir::Anchor { col: 3, col_off: 38_100, row: 5, row_off: 50_800 }),
            extent: Some((1_000_000, 600_000)),
            kind: crate::ir::DrawingKind::Shape(crate::ir::Shape {
                geometry: "rect".into(), text_box: false, fill: Some("112233".into()), fill_theme: None, fill_legacy_color_index: None, fill_brightness: None,
                line: Some(crate::ir::ShapeLine {
                    color: "FF0000".into(), pattern: None, back_color: None, back_color_theme: None, back_legacy_color_index: None, back_color_brightness: None, back_color_shade: None, width: 19_050,
                    color_theme: None, legacy_color_index: None, color_brightness: None, color_shade: None,
                    dash: Some("dash".into()), cap: None, head_end: None, tail_end: None,
                    head_end_length: None, head_end_width: None, tail_end_length: None, tail_end_width: None,
                }),
                adjusts: Vec::new(), path: None, flip_h: false, flip_v: false, rotation: 0, text: None,
            }),
            frame: None,
            grouped: false,
            group_name: None,
            group_rotation: 0,
            group_transform: None,
            group_child_transform: None,
            group_nested: false,
        }
    }

    fn workbook_without_drawings() -> Vec<u8> {
        let parts = [
            ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#),
            ("xl/workbook.xml", r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#),
            ("xl/_rels/workbook.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#),
            ("xl/worksheets/sheet1.xml", r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/></worksheet>"#),
        ];
        let mut bytes = Vec::new();
        let mut writer = ZipWriter::new(Cursor::new(&mut bytes));
        for (name, xml) in parts {
            writer.start_file(name, SimpleFileOptions::default()).unwrap();
            writer.write_all(xml.as_bytes()).unwrap();
        }
        writer.finish().unwrap();
        bytes
    }

    #[test]
    fn drawing_line_patch_replaces_only_the_line_paint() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Patterned"/></xdr:nvSpPr><xdr:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="111111"/></a:solidFill><a:prstDash val="dash"/><a:headEnd type="triangle"/></a:ln></xdr:spPr></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>"#;
        let line = crate::ir::ShapeLine {
            color: "FF0000".into(),
            color_theme: None,
            legacy_color_index: None,
            color_brightness: None,
            color_shade: None,
            pattern: Some("pct10".into()),
            back_color: Some("00FF00".into()),
            back_color_theme: None,
            back_legacy_color_index: None,
            back_color_brightness: None,
            back_color_shade: None,
            width: 19050,
            dash: Some("dash".into()),
            cap: None,
            head_end: Some("triangle".into()),
            tail_end: None,
            head_end_length: None,
            head_end_width: None,
            tail_end_length: None,
            tail_end_width: None,
        };
        let saved = patch_drawing_line_xml(xml, &HashMap::from([("Patterned".into(), Some(line))])).unwrap();
        assert!(saved.contains("<a:pattFill prst=\"pct10\">"));
        assert!(saved.contains("<a:fgClr><a:srgbClr val=\"FF0000\"/></a:fgClr>"));
        assert!(saved.contains("<a:bgClr><a:srgbClr val=\"00FF00\"/></a:bgClr>"));
        assert!(saved.contains("<a:prstDash val=\"dash\"/>"));
        assert!(saved.contains("<a:headEnd type=\"triangle\"/>"));
        assert!(!saved.contains("val=\"111111\""));
    }

    #[test]
    fn drawing_deletion_removes_only_the_named_anchor() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:twoCellAnchor><xdr:from/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="Keep"/></xdr:nvSpPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor><xdr:oneCellAnchor><xdr:from/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Delete &amp; Me"/></xdr:nvSpPr></xdr:sp><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>"#;
        let saved = patch_drawing_deletions_xml(xml, &std::collections::HashSet::from(["Delete & Me".into()]), None, None).unwrap();
        assert!(saved.contains("name=\"Keep\""));
        assert!(!saved.contains("Delete"));
        let mut reader = Reader::from_str(&saved);
        loop {
            match reader.read_event().unwrap() {
                Event::Eof => break,
                _ => {}
            }
        }
    }

    #[test]
    fn drawing_position_patch_updates_two_cell_anchors_only() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="Move"/></xdr:nvSpPr><xdr:spPr><a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:off x="0" y="0"/><a:ext cx="100" cy="200"/></a:xfrm></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;
        let after = crate::ir::Drawing {
            name: Some("Move".into()),
            ole_prog_id: None,
            ole_source_name: None,
            ole_clone_position: None,
            ole_embedded_package: None,
            ole_package_is_empty: None,
            ole_clone_visible: None,
            ole_placement: None,
            ole_type: None,
            ole_display_as_icon: None,
            ole_auto_load: None,
            ole_enabled: None,
            ole_locked: None,
            ole_print_object: None,
            ole_shadow: None,
            ole_on_action: None,
            ole_border_color: None,
            ole_border_weight: None,
            ole_interior_color: None,
            from: crate::ir::Anchor { col: 1, col_off: 127, row: 3, row_off: 254 },
            to: Some(crate::ir::Anchor { col: 4, col_off: 381, row: 6, row_off: 508 }),
            extent: Some((300, 400)),
            kind: crate::ir::DrawingKind::Shape(crate::ir::Shape {
                geometry: "rect".into(), text_box: false, fill: None, fill_theme: None, fill_legacy_color_index: None, fill_brightness: None, line: None,
                adjusts: Vec::new(), path: None, flip_h: false, flip_v: false, rotation: 0, text: None,
            }),
            frame: None,
            grouped: false,
            group_name: None,
            group_rotation: 0,
            group_transform: None,
            group_child_transform: None,
            group_nested: false,
        };
        let saved = patch_drawing_geometry_xml(xml, &HashMap::from([("Move".into(), after)])).unwrap();
        assert!(saved.contains("<xdr:col>1</xdr:col><xdr:colOff>127</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>254</xdr:rowOff>"));
        assert!(saved.contains("<xdr:col>4</xdr:col><xdr:colOff>381</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>508</xdr:rowOff>"));
        assert!(saved.contains("<a:off x=\"0\" y=\"0\"/><a:ext cx=\"100\" cy=\"200\"/>"));
    }

    #[test]
    fn drawing_fill_patch_changes_shape_fill_without_changing_line_paint() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="Paint"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;
        let saved = patch_drawing_fill_xml(xml, &HashMap::from([("Paint".into(), (Some("AABBCC".into()), None, None, None))])).unwrap();
        assert!(saved.contains(r#"<a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill>"#));
        assert!(saved.contains(r#"<a:srgbClr val="112233"/>"#));
    }

    #[test]
    fn drawing_fill_patch_adds_or_removes_explicit_fill() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="Paint"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;
        let colored = patch_drawing_fill_xml(xml, &HashMap::from([("Paint".into(), (Some("AABBCC".into()), None, None, None))])).unwrap();
        assert!(colored.contains(r#"<a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill>"#));
        let cleared = patch_drawing_fill_xml(&colored, &HashMap::from([("Paint".into(), (None, None, None, None))])).unwrap();
        assert!(cleared.contains("<a:noFill></a:noFill>"));
        assert!(!cleared.contains("AABBCC"));
    }

    #[test]
    fn drawing_text_color_patch_changes_run_and_default_run_colors_only() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="Text"/></xdr:nvSpPr><xdr:txBody><a:bodyPr wrap="none" vertOverflow="clip" anchor="ctr" vert="vert" rtlCol="1" numCol="2" customAttr="keep"><a:noAutofit/></a:bodyPr><a:lstStyle><a:lvl1pPr marL="1200"/></a:lstStyle><a:p><a:pPr algn="r" marL="2400"><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buNone/></a:pPr><a:r><a:rPr b="1" baseline="1500" cap="small"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:highlight><a:srgbClr val="FFFF00"/></a:highlight></a:rPr><a:t>first</a:t></a:r><a:endParaRPr sz="1100" baseline="1500" kern="900"/></a:p><a:p><a:r><a:rPr b="0"><a:solidFill><a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/></a:rPr><a:t>second</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;
        let text = crate::ir::ShapeText {
            paragraphs: vec![
                crate::ir::ShapeParagraph {
                    text: "first".into(), runs: vec![crate::ir::ShapeRun { text: "first".into(), bold: true, italic: None, underline: false, face: None, size: None, color: Some("00FF00".into()), color_theme: None, color_brightness: None, legacy_color_index: None }],
                    align: Some("l".into()), size: 11.0, bold: true, italic: false, face: Some("Arial".into()), charset: None,
                    pitch_family: None, color: Some("00FF00".into()), line_pitch: None, line_scale: None,
                },
                crate::ir::ShapeParagraph {
                    text: "second".into(), runs: vec![crate::ir::ShapeRun { text: "second".into(), bold: false, italic: None, underline: false, face: None, size: None, color: Some("445566".into()), color_theme: None, color_brightness: None, legacy_color_index: None }],
                    align: Some("l".into()), size: 11.0, bold: false, italic: false, face: Some("Arial".into()), charset: None,
                    pitch_family: None, color: Some("445566".into()), line_pitch: None, line_scale: None,
                },
            ],
            anchor: Some("t".into()), insets: (91440, 45720, 91440, 45720), wrap: true, clip: false,
        };
        let mut before = text.clone();
        before.paragraphs[0].runs[0].color = Some("112233".into());
        before.paragraphs[0].color = Some("112233".into());
        let saved = patch_drawing_text_xml(xml, &HashMap::from([("Text".into(), ShapeTextEdit { before, after: text })])).unwrap();
        assert!(saved.contains(r#"<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>"#));
        assert!(saved.contains("b=\"1\""));
        assert!(saved.contains("sz=\"1100\""));
        assert!(saved.contains("numCol=\"2\""));
        assert!(saved.contains("customAttr=\"keep\""));
        assert!(saved.contains("vert=\"vert\""));
        assert!(saved.contains("rtlCol=\"1\""));
        assert!(saved.contains("<a:noAutofit"));
        assert!(saved.contains("<a:lvl1pPr marL=\"1200\""));
        assert!(saved.contains("algn=\"l\" marL=\"2400\""));
        assert!(saved.contains("<a:spcBef><a:spcPts val=\"500\"/>"));
        assert!(saved.contains("<a:buNone"));
        assert!(saved.contains("baseline=\"1500\" kern=\"900\""));
        assert!(saved.contains("baseline=\"1500\" cap=\"small\""));
        assert!(saved.contains("<a:highlight><a:srgbClr val=\"FFFF00\"/></a:highlight>"));
        assert!(saved.contains("<a:schemeClr val=\"accent1\"><a:tint val=\"50000\"/></a:schemeClr>"));
        assert!(saved.contains("<a:latin typeface=\"+mn-lt\"/>"));
        assert!(!saved.contains("112233"));
    }

    #[test]
    fn shape_fill_change_is_saved_in_the_workbook() {
        let source = workbook_with_line_shape();
        let mut editor = XlsxEditor::new(&source).unwrap();
        let mut edited = editor.workbook().clone();
        let crate::ir::DrawingKind::Shape(shape) = &mut edited.sheets[0].drawings[0].kind else { panic!("shape expected") };
        shape.fill = Some("AABBCC".into());
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let reparsed = crate::parser::parse_xlsx(&saved).unwrap();
        let crate::ir::DrawingKind::Shape(shape) = &reparsed.sheets[0].drawings[0].kind else { panic!("shape expected") };
        assert_eq!(shape.fill.as_deref(), Some("AABBCC"));
        assert_eq!(shape.line.as_ref().unwrap().color, "111111");
    }

    #[test]
    fn drawing_position_patch_updates_one_cell_extent_without_changing_anchor_kind() {
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="100" cy="200"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="Resize"/></xdr:nvSpPr></xdr:sp><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>"#;
        let after = crate::ir::Drawing {
            name: Some("Resize".into()),
            ole_prog_id: None,
            ole_source_name: None,
            ole_clone_position: None,
            ole_embedded_package: None,
            ole_package_is_empty: None,
            ole_clone_visible: None,
            ole_placement: None,
            ole_type: None,
            ole_display_as_icon: None,
            ole_auto_load: None,
            ole_enabled: None,
            ole_locked: None,
            ole_print_object: None,
            ole_shadow: None,
            ole_on_action: None,
            ole_border_color: None,
            ole_border_weight: None,
            ole_interior_color: None,
            from: crate::ir::Anchor { col: 1, col_off: 127, row: 3, row_off: 254 },
            to: None,
            extent: Some((300, 400)),
            kind: crate::ir::DrawingKind::Shape(crate::ir::Shape {
                geometry: "rect".into(), text_box: false, fill: None, fill_theme: None, fill_legacy_color_index: None, fill_brightness: None, line: None,
                adjusts: Vec::new(), path: None, flip_h: false, flip_v: false, rotation: 0, text: None,
            }),
            frame: None,
            grouped: false,
            group_name: None,
            group_rotation: 0,
            group_transform: None,
            group_child_transform: None,
            group_nested: false,
        };
        let saved = patch_drawing_geometry_xml(xml, &HashMap::from([("Resize".into(), after)])).unwrap();
        assert!(saved.contains("<xdr:col>1</xdr:col><xdr:colOff>127</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>254</xdr:rowOff>"));
        assert!(saved.contains("<xdr:ext cx=\"300\" cy=\"400\"/>"));
        assert!(!saved.contains("<xdr:to>"));
    }

    #[test]
    fn deleted_shape_is_removed_from_the_saved_workbook() {
        let source = workbook_with_line_shape();
        let mut editor = XlsxEditor::new(&source).unwrap();
        let mut edited = editor.workbook().clone();
        edited.sheets[0].drawings.clear();
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let reparsed = crate::parser::parse_xlsx(&saved).unwrap();
        assert!(reparsed.sheets[0].drawings.is_empty());
        let mut package = ZipArchive::new(Cursor::new(&saved)).unwrap();
        let mut drawing_xml = String::new();
        package.by_name("xl/drawings/drawing1.xml").unwrap().read_to_string(&mut drawing_xml).unwrap();
        assert!(!drawing_xml.contains("Patterned"));
    }

    #[test]
    fn moved_shape_is_saved_with_its_new_cell_anchors() {
        let source = workbook_with_line_shape();
        let mut editor = XlsxEditor::new(&source).unwrap();
        let mut edited = editor.workbook().clone();
        edited.sheets[0].drawings[0].from = crate::ir::Anchor { col: 1, col_off: 12_700, row: 3, row_off: 25_400 };
        edited.sheets[0].drawings[0].to = Some(crate::ir::Anchor { col: 4, col_off: 38_100, row: 6, row_off: 50_800 });
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let reparsed = crate::parser::parse_xlsx(&saved).unwrap();
        assert_eq!(reparsed.sheets[0].drawings[0].from, edited.sheets[0].drawings[0].from);
        assert_eq!(reparsed.sheets[0].drawings[0].to, edited.sheets[0].drawings[0].to);
    }

    #[test]
    fn moved_ole_shape_updates_worksheet_object_pr_anchor() {
        let xml = r#"<worksheet xmlns:xdr="urn:xdr"><oleObjects><oleObject shapeId="1025"><objectPr defaultSize="0"><anchor moveWithCells="1"><from><xdr:col>4</xdr:col><xdr:colOff>10</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>20</xdr:rowOff></from><to><xdr:col>5</xdr:col><xdr:colOff>30</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>40</xdr:rowOff></to></anchor></objectPr></oleObject><oleObject shapeId="1026"/></oleObjects></worksheet>"#;
        let source = workbook_with_line_shape();
        let editor = XlsxEditor::new(&source).unwrap();
        let mut moved = editor.workbook().sheets[0].drawings[0].clone();
        moved.from = crate::ir::Anchor { col: 4, col_off: 50, row: 6, row_off: 20 };
        moved.to = Some(crate::ir::Anchor { col: 5, col_off: 70, row: 8, row_off: 40 });
        moved.ole_prog_id = Some("Package".into());
        moved.ole_placement = Some(1);
        let saved = patch_ole_object_anchors(xml, &HashMap::from([(1025, moved)])).unwrap();
        assert!(saved.contains("<xdr:colOff>50</xdr:colOff>"));
        assert!(saved.contains("<xdr:colOff>70</xdr:colOff>"));
        assert!(saved.contains("<anchor moveWithCells=\"1\" sizeWithCells=\"1\">"));
        assert!(saved.contains("<oleObject shapeId=\"1026\"/>"));
        assert!(saved.contains("<objectPr defaultSize=\"0\">") && saved.contains("moveWithCells=\"1\""));
    }

    #[test]
    fn ole_shadow_updates_the_matching_vml_shape() {
        let xml = r#"<xml xmlns:v="urn:vml" xmlns:x="urn:excel"><v:shape id="_x0000_s1025"><v:imagedata/><x:ClientData/></v:shape><v:shape id="_x0000_s1026"><v:shadow on="t" color="black"/><x:ClientData/></v:shape></xml>"#;
        let saved = patch_vml_ole_shadows(xml, &HashMap::from([(1025, true), (1026, false)])).unwrap();
        assert!(saved.contains("<v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>"));
        assert!(!saved.contains(r#"_x0000_s1026"><v:shadow"#));
        assert!(saved.contains("<v:imagedata/><v:shadow") && saved.contains("<x:ClientData/>"));
    }

    #[test]
    fn patterned_shape_line_survives_workbook_save_and_reparse() {
        let source = workbook_with_line_shape();
        let mut editor = XlsxEditor::new(&source).unwrap();
        let mut edited = editor.workbook().clone();
        let crate::ir::DrawingKind::Shape(shape) = &mut edited.sheets[0].drawings[0].kind else { panic!("shape expected") };
        let line = shape.line.as_mut().unwrap();
        line.pattern = Some("pct10".into());
        line.color = "FF0000".into();
        line.back_color = Some("00FF00".into());
        editor.apply_workbook(&edited).unwrap();
        assert!(editor.has_edits());
        let saved = editor.save().unwrap();
        let mut package = ZipArchive::new(Cursor::new(&saved)).unwrap();
        let mut drawing_xml = String::new();
        package.by_name("xl/drawings/drawing1.xml").unwrap().read_to_string(&mut drawing_xml).unwrap();
        assert!(drawing_xml.contains("<a:pattFill prst=\"pct10\">"));
        let reparsed = crate::parser::parse_xlsx(&saved).unwrap();
        let crate::ir::DrawingKind::Shape(shape) = &reparsed.sheets[0].drawings[0].kind else { panic!("shape expected after reparse") };
        let line = shape.line.as_ref().unwrap();
        assert_eq!(line.pattern.as_deref(), Some("pct10"));
        assert_eq!(line.color, "FF0000");
        assert_eq!(line.back_color.as_deref(), Some("00FF00"));
        assert_eq!(line.dash.as_deref(), Some("dash"));
    }

    #[test]
    fn added_shape_is_written_to_an_existing_drawing_part() {
        let source = workbook_with_line_shape();
        let mut editor = XlsxEditor::new(&source).unwrap();
        let mut edited = editor.workbook().clone();
        let mut added = simple_added_shape("New & Added");
        let crate::ir::DrawingKind::Shape(shape) = &mut added.kind else { unreachable!() };
        let line = shape.line.as_mut().unwrap();
        line.color_theme = Some(5);
        line.color_shade = Some(0.15);
        edited.sheets[0].drawings.push(added);
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let reparsed = crate::parser::parse_xlsx(&saved).unwrap();
        assert_eq!(reparsed.sheets[0].drawings.len(), 2);
        let added = &reparsed.sheets[0].drawings[1];
        assert_eq!(added.name.as_deref(), Some("New & Added"));
        let crate::ir::DrawingKind::Shape(shape) = &added.kind else { panic!("added shape expected") };
        assert_eq!(shape.geometry, "rect");
        assert_eq!(shape.fill.as_deref(), Some("112233"));
        assert_eq!(shape.line.as_ref().and_then(|line| line.dash.as_deref()), Some("dash"));
        assert_eq!(shape.line.as_ref().and_then(|line| line.color_shade), Some(0.15));
        let mut package = ZipArchive::new(Cursor::new(&saved)).unwrap();
        let mut drawing_xml = String::new();
        package.by_name("xl/drawings/drawing1.xml").unwrap().read_to_string(&mut drawing_xml).unwrap();
        assert!(drawing_xml.contains("<a:schemeClr val=\"accent1\"><a:shade val=\"15000\"/></a:schemeClr>"));
    }

    #[test]
    fn added_shape_creates_drawing_relationship_and_content_type() {
        let source = workbook_without_drawings();
        let mut editor = XlsxEditor::new(&source).unwrap();
        let mut edited = editor.workbook().clone();
        edited.sheets[0].drawings.push(simple_added_shape("Rectangle 1"));
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let reparsed = crate::parser::parse_xlsx(&saved).unwrap();
        assert_eq!(reparsed.sheets[0].drawings.len(), 1);
        assert_eq!(reparsed.sheets[0].drawings[0].name.as_deref(), Some("Rectangle 1"));
        let mut package = ZipArchive::new(Cursor::new(&saved)).unwrap();
        let mut types = String::new();
        package.by_name("[Content_Types].xml").unwrap().read_to_string(&mut types).unwrap();
        assert!(types.contains("application/vnd.openxmlformats-officedocument.drawing+xml"));
        assert!(package.file_names().any(|name| name == "xl/drawings/drawing1.xml"));
    }

    #[test]
    fn added_line_is_saved_as_a_connector_shape() {
        let source = workbook_without_drawings();
        let mut editor = XlsxEditor::new(&source).unwrap();
        let mut edited = editor.workbook().clone();
        let mut line = simple_added_shape("Line 1");
        let crate::ir::DrawingKind::Shape(shape) = &mut line.kind else { unreachable!() };
        shape.geometry = "line".into();
        shape.fill = None;
        edited.sheets[0].drawings.push(line);
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let mut package = ZipArchive::new(Cursor::new(&saved)).unwrap();
        let mut drawing_xml = String::new();
        package.by_name("xl/drawings/drawing1.xml").unwrap().read_to_string(&mut drawing_xml).unwrap();
        assert!(drawing_xml.contains("<xdr:cxnSp"));
        assert!(drawing_xml.contains("<a:prstGeom prst=\"line\">"));
        let reparsed = crate::parser::parse_xlsx(&saved).unwrap();
        let crate::ir::DrawingKind::Shape(shape) = &reparsed.sheets[0].drawings[0].kind else { panic!("connector expected") };
        assert_eq!(shape.geometry, "line");
    }

    #[test]
    fn added_text_box_keeps_its_text_typeface_and_colour() {
        let source = workbook_without_drawings();
        let mut editor = XlsxEditor::new(&source).unwrap();
        let mut edited = editor.workbook().clone();
        let mut drawing = simple_added_shape("TextBox 1");
        let crate::ir::DrawingKind::Shape(shape) = &mut drawing.kind else { unreachable!() };
        shape.text_box = true;
        shape.text = Some(crate::ir::ShapeText {
            paragraphs: vec![crate::ir::ShapeParagraph {
                text: "A & <B>".into(), runs: Vec::new(), align: Some("ctr".into()), size: 12.0,
                bold: true, italic: true, face: Some("Arial".into()), charset: Some(0), pitch_family: Some(34),
                color: Some("112233".into()), line_pitch: None, line_scale: None,
            }],
            anchor: Some("ctr".into()), insets: (9_144, 9_144, 9_144, 9_144), wrap: true, clip: false,
        });
        edited.sheets[0].drawings.push(drawing);
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let reparsed = crate::parser::parse_xlsx(&saved).unwrap();
        let crate::ir::DrawingKind::Shape(shape) = &reparsed.sheets[0].drawings[0].kind else { panic!("text box expected") };
        assert!(shape.text_box);
        let text = shape.text.as_ref().expect("text box content");
        assert_eq!(text.paragraphs[0].text, "A & <B>");
        assert_eq!(text.paragraphs[0].face.as_deref(), Some("Arial"));
        assert_eq!(text.paragraphs[0].color.as_deref(), Some("112233"));
        assert!(text.paragraphs[0].bold);
        assert!(text.paragraphs[0].italic);
    }

    #[test]
    fn legacy_comment_creation_update_and_deletion_roundtrip_through_save() {
        let base = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(base).unwrap();
        let mut edited = editor.workbook().clone();
        edited.sheets[0].comments.push(crate::ir::Comment {
            author: "VBA Author".into(),
            utf16_text: None,
            from: crate::ir::Anchor { col: 2, col_off: 15 * 9525, row: 4, row_off: 10 * 9525 },
            size: (120.0, 59.25),
            to: Some(crate::ir::Anchor { col: 5, col_off: 47 * 9525, row: 7, row_off: 9 * 9525 }),
            text: crate::ir::ShapeText {
                paragraphs: vec![crate::ir::ShapeParagraph {
                    text: "A & <B>".into(), runs: Vec::new(), align: None, size: 9.0, bold: false, italic: false,
                    face: Some("Arial".into()), charset: None, pitch_family: None, color: None, line_pitch: None, line_scale: None,
                }],
                anchor: Some("t".into()), insets: (91_440, 45_720, 91_440, 45_720), wrap: true, clip: true,
            },
            fill: Some("FFFFE1".into()),
            fill_visible: false,
            fill_transparency: 0.4,
            line_color: Some("FF0000".into()),
            line_visible: false,
            line_weight: 2.25,
            line_transparency: 0.4,
            cell: (4, 2),
            visible: true,
        });
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let mut package = ZipArchive::new(Cursor::new(saved.clone())).unwrap();
        let mut sheet_xml = String::new();
        package
            .by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut sheet_xml)
            .unwrap();
        assert!(sheet_xml.contains("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""));
        let mut reader = Reader::from_str(&sheet_xml);
        while reader.read_event().unwrap() != Event::Eof {}
        let mut vml = String::new();
        package
            .by_name("xl/drawings/vmlDrawing1.vml")
            .unwrap()
            .read_to_string(&mut vml)
            .unwrap();
        assert!(vml.contains("fillcolor=\"infoBackground [80]\""));
        assert!(vml.contains("filled=\"f\""));
        assert!(vml.contains("<v:fill opacity=\"39322f\""));
        assert!(vml.contains("<v:stroke opacity=\"39322f\"/>"));
        assert!(vml.contains("strokecolor=\"red\""));
        let reopened = crate::parse_xlsx(&saved).unwrap();
        let note = &reopened.sheets[0].comments[0];
        let original_xml = "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:xr=\"urn:test\"><authors><author>VBA Author</author></authors><commentList><comment ref=\"C5\" authorId=\"0\" xr:uid=\"keep-me\"><text><r><rPr><b/></rPr><t>A &amp; </t></r><r><t>&lt;B&gt;</t></r></text></comment></commentList></comments>";
        let mut added = note.clone();
        added.cell = (5, 3);
        added.author = "New & Author".into();
        let patched = patch_legacy_comments(original_xml, std::slice::from_ref(note), &[note.clone(), added]).unwrap();
        assert!(patched.contains("xr:uid=\"keep-me\"><text><r><rPr><b/></rPr><t>A &amp; </t></r><r><t>&lt;B&gt;</t></r></text></comment>"));
        assert!(patched.contains("<author>New &amp; Author</author>"));
        assert!(patched.contains("ref=\"D6\" authorId=\"1\""));
        let mut changed_text = note.clone();
        changed_text.text.paragraphs[0].text = "replacement".into();
        let changed_xml = patch_legacy_comments(original_xml, std::slice::from_ref(note), &[changed_text]).unwrap();
        assert!(changed_xml.contains("ref=\"C5\" authorId=\"0\" xr:uid=\"keep-me\""));
        assert!(changed_xml.contains(">replacement</t>"));
        let mut rich = note.clone();
        rich.text.paragraphs[0].runs = vec![
            crate::ir::ShapeRun { text: "A & ".into(), bold: true, italic: Some(false), underline: false, face: None, size: Some(9.0), color: None, color_theme: None, color_brightness: None, legacy_color_index: None },
            crate::ir::ShapeRun { text: "<B>".into(), bold: false, italic: Some(true), underline: false, face: None, size: Some(9.0), color: Some("FF0000".into()), color_theme: None, color_brightness: None, legacy_color_index: None },
        ];
        let rich_xml = patch_legacy_comments(original_xml, std::slice::from_ref(note), &[rich]).unwrap();
        assert!(rich_xml.contains("<b/><sz val=\"9\""));
        assert!(rich_xml.contains("<i/><sz val=\"9\"/><color rgb=\"FFFF0000\"/>"));
        assert!(rich_xml.contains("&lt;B&gt;</t>"));
        assert_eq!(note.author, "VBA Author");
        assert_eq!(note.cell, (4, 2));
        assert_eq!(note.text.paragraphs[0].text, "A & <B>");
        assert_eq!(note.size, (120.0, 59.25));
        assert_eq!(note.line_color.as_deref(), Some("FF0000"));
        assert!(!note.fill_visible);
        assert!((note.fill_transparency - 0.399_993_9).abs() < 0.000_001);
        assert!((note.line_transparency - 0.399_993_9).abs() < 0.000_001);
        assert!(!note.line_visible);
        assert_eq!(note.line_weight, 2.25);
        assert!(note.visible);

        let mut second = XlsxEditor::new(&saved).unwrap();
        let mut changed = second.workbook().clone();
        changed.sheets[0].comments[0].text.paragraphs[0].text = "edited text".into();
        changed.sheets[0].comments[0].line_visible = true;
        changed.sheets[0].comments[0].line_weight = 1.5;
        second.apply_workbook(&changed).unwrap();
        let saved = second.save().unwrap();
        let reopened = crate::parse_xlsx(&saved).unwrap();
        assert_eq!(reopened.sheets[0].comments[0].text.paragraphs[0].text, "edited text");
        assert!(reopened.sheets[0].comments[0].line_visible);
        assert_eq!(reopened.sheets[0].comments[0].line_weight, 1.5);

        let mut third = XlsxEditor::new(&saved).unwrap();
        let mut deleted = third.workbook().clone();
        deleted.sheets[0].comments.clear();
        third.apply_workbook(&deleted).unwrap();
        let reopened = crate::parse_xlsx(&third.save().unwrap()).unwrap();
        assert!(reopened.sheets[0].comments.is_empty());
    }

    #[test]
    fn edited_threaded_comment_text_roundtrips_through_save() {
        let base = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut input = ZipArchive::new(Cursor::new(base.to_vec())).unwrap();
        let mut package = Vec::new();
        {
            let mut output = ZipWriter::new(Cursor::new(&mut package));
            for index in 0..input.len() {
                let mut entry = input.by_index(index).unwrap();
                let name = entry.name().to_string();
                let mut bytes = Vec::new();
                entry.read_to_end(&mut bytes).unwrap();
                if name == "[Content_Types].xml" {
                    let xml = String::from_utf8(bytes).unwrap();
                    let xml = xml.replace(
                        "</Types>",
                        "<Override PartName=\"/xl/threadedComments/threadedComment1.xml\" ContentType=\"application/vnd.ms-excel.threadedcomments+xml\"/></Types>",
                    );
                    bytes = xml.into_bytes();
                }
                output.start_file(&name, SimpleFileOptions::default()).unwrap();
                output.write_all(&bytes).unwrap();
            }
            output.start_file("xl/worksheets/_rels/sheet1.xml.rels", SimpleFileOptions::default()).unwrap();
            output.write_all(br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdThread" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>"#).unwrap();
            output.start_file("xl/threadedComments/threadedComment1.xml", SimpleFileOptions::default()).unwrap();
            output.write_all(br#"<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="A1" dT="2026-01-01T00:00:00Z" personId="{PERSON}" id="{COMMENT}"><text>old</text></threadedComment></ThreadedComments>"#).unwrap();
            output.finish().unwrap();
        }

        let mut editor = XlsxEditor::new(&package).unwrap();
        let mut edited = editor.workbook().clone();
        edited.sheets[0].threaded_comments.push(crate::ir::ThreadedComment {
            utf16_text: None,
            cell: (0, 0), id: "{COMMENT}".into(), parent_id: None,
            person_id: "{PERSON}".into(), date: "2026-01-01T00:00:00Z".into(), text: "new & <review>".into(),
            person_name: "User".into(),
            person_provider_id: "None".into(), person_user_id: "User".into(),
            resolved: false,
        });
        edited.sheets[0].threaded_comments.push(crate::ir::ThreadedComment {
            utf16_text: None,
            cell: (1, 1), id: String::new(), parent_id: None,
            person_id: String::new(), date: String::new(), text: "added beside existing".into(),
            person_name: "User".into(),
            person_provider_id: "None".into(), person_user_id: "User".into(),
            resolved: false,
        });
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let reopened = parse_xlsx(&saved).unwrap();
        assert_eq!(reopened.sheets[0].threaded_comments[0].text, "new & <review>");
        assert_eq!(reopened.sheets[0].threaded_comments[1].text, "added beside existing");
        assert_eq!(reopened.sheets[0].threaded_comments[1].person_name, "User");
        assert_eq!(reopened.sheets[0].threaded_comments[1].person_user_id, "User");
        assert_eq!(reopened.sheets[0].threaded_comments[1].person_provider_id, "None");
    }

    #[test]
    fn threaded_comment_unpaired_utf16_and_literal_escape_roundtrip() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(data).unwrap();
        let mut edited = editor.workbook().clone();
        let units = vec![65, 88, 0xDE00, 66];
        edited.sheets[0].threaded_comments.push(crate::ir::ThreadedComment {
            cell: (0, 0), id: String::new(), parent_id: None,
            person_id: String::new(), person_name: "User".into(),
            person_provider_id: "None".into(), person_user_id: "User".into(),
            date: String::new(), text: String::from_utf16_lossy(&units),
            utf16_text: Some(units.clone()), resolved: false,
        });
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let reopened = parse_xlsx(&saved).unwrap();
        assert_eq!(reopened.sheets[0].threaded_comments[0].utf16_text.as_ref(), Some(&units));
        assert!(reopened.sheets[0].comments[0].utf16_text.as_ref().unwrap().ends_with(&units));

        let mut editor = XlsxEditor::new(&saved).unwrap();
        let mut edited = editor.workbook().clone();
        edited.sheets[0].threaded_comments[0].text = "literal _xDE00_".into();
        edited.sheets[0].threaded_comments[0].utf16_text = None;
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let reopened = parse_xlsx(&saved).unwrap();
        assert_eq!(reopened.sheets[0].threaded_comments[0].text, "literal _xDE00_");
        assert_eq!(reopened.sheets[0].threaded_comments[0].utf16_text, None);
    }

    #[test]
    fn new_threaded_comment_writes_all_required_package_parts() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(data).unwrap();
        let mut edited = editor.workbook().clone();
        edited.sheets[0].threaded_comments.push(crate::ir::ThreadedComment {
            utf16_text: None,
            cell: (3, 2), id: String::new(), parent_id: None,
            person_id: String::new(), date: String::new(), text: "new thread & reply".into(),
            person_name: "User".into(),
            person_provider_id: "None".into(), person_user_id: "User".into(),
            resolved: true,
        });
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();

        let reopened = parse_xlsx(&saved).unwrap();
        let comment = &reopened.sheets[0].threaded_comments[0];
        assert_eq!(comment.cell, (3, 2));
        assert_eq!(comment.text, "new thread & reply");
        assert!(comment.id.starts_with('{') && comment.id.ends_with('}'));
        assert!(comment.person_id.starts_with('{') && comment.person_id.ends_with('}'));
        assert!(comment.resolved);
        assert_eq!(comment.person_name, "User");
        assert_eq!(comment.person_user_id, "User");
        assert_eq!(comment.person_provider_id, "None");

        let mut archive = ZipArchive::new(Cursor::new(saved.clone())).unwrap();
        for part in [
            "xl/threadedComments/threadedComment1.xml",
            "xl/persons/person.xml",
            "xl/worksheets/_rels/sheet1.xml.rels",
        ] {
            assert!(archive.by_name(part).is_ok(), "missing {part}");
        }
        let mut person_xml = String::new();
        archive.by_name("xl/persons/person.xml").unwrap().read_to_string(&mut person_xml).unwrap();
        assert!(person_xml.contains(&comment.person_id));
        assert!(person_xml.contains("providerId=\"None\""));

        let mut unresolved = reopened.clone();
        unresolved.sheets[0].threaded_comments[0].resolved = false;
        let mut editor = XlsxEditor::new(&saved).unwrap();
        editor.apply_workbook(&unresolved).unwrap();
        let unresolved = parse_xlsx(&editor.save().unwrap()).unwrap();
        assert!(!unresolved.sheets[0].threaded_comments[0].resolved);

        let mut removed = reopened;
        removed.sheets[0].threaded_comments.clear();
        let mut editor = XlsxEditor::new(&saved).unwrap();
        editor.apply_workbook(&removed).unwrap();
        let deleted = parse_xlsx(&editor.save().unwrap()).unwrap();
        assert!(deleted.sheets[0].threaded_comments.is_empty());
    }

    #[test]
    fn column_width_custom_state_roundtrips_even_when_width_is_unchanged() {
        let data = include_bytes!("../../../tests/fixtures/vba/cell_width.xlsm");
        let mut edited = parse_xlsx(data).unwrap();
        assert_eq!(edited.sheets[0].column_width_font.as_deref(), Some("Aptos Narrow"));
        edited.sheets[0].col_widths.resize(4, 0.0);
        edited.sheets[0].col_widths[3] = 9.140625;
        edited.sheets[0].col_custom_widths.insert(3, true);
        let mut editor = XlsxEditor::new(data).unwrap();
        editor.apply_workbook(&edited).unwrap();
        let sized = editor.save().unwrap();
        let mut reset = parse_xlsx(&sized).unwrap();
        assert!(reset.sheets[0].column_has_custom_width(3));
        reset.sheets[0].col_custom_widths.insert(3, false);
        let mut editor = XlsxEditor::new(&sized).unwrap();
        editor.apply_workbook(&reset).unwrap();
        let restored = parse_xlsx(&editor.save().unwrap()).unwrap();
        assert_eq!(restored.sheets[0].col_widths[3], 9.140625);
        assert!(!restored.sheets[0].column_has_custom_width(3));
    }

    #[test]
    fn standard_width_and_reset_column_survive_save() {
        let data = include_bytes!("../../../tests/fixtures/vba/cell_width.xlsm");
        let mut edited = parse_xlsx(data).unwrap();
        edited.sheets[0].col_widths.resize(4, 0.0);
        edited.sheets[0].col_widths[3] = 13.285156;
        edited.sheets[0].col_custom_widths.insert(3, true);
        let mut editor = XlsxEditor::new(data).unwrap();
        editor.apply_workbook(&edited).unwrap();
        let sized = editor.save().unwrap();
        let mut reset = parse_xlsx(&sized).unwrap();
        reset.sheets[0].col_widths[3] = 0.0;
        reset.sheets[0].col_custom_widths.insert(3, false);
        reset.sheets[0].default_col_width = 11.285156;
        reset.sheets[0].default_col_width_explicit = true;
        let mut editor = XlsxEditor::new(&sized).unwrap();
        editor.apply_workbook(&reset).unwrap();
        let restored = parse_xlsx(&editor.save().unwrap()).unwrap();
        assert_eq!(restored.sheets[0].default_col_width, 11.285156);
        assert!(!restored.sheets[0].column_has_custom_width(3));
        reset.sheets[0].default_col_width = 0.0;
        let mut editor = XlsxEditor::new(&sized).unwrap();
        editor.apply_workbook(&reset).unwrap();
        let zero = parse_xlsx(&editor.save().unwrap()).unwrap();
        assert!(zero.sheets[0].default_col_width_explicit);
        assert_eq!(zero.sheets[0].default_col_width, 0.0);
    }

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(1), "B");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
        assert_eq!(col_to_letter(27), "AB");
    }

    #[test]
    fn calculation_settings_roundtrip_preserves_unmodeled_attributes() {
        let settings = CalculationSettings { mode: CalculationMode::Manual, on_save: false, completed: Some(false), ..Default::default() };
        for xml in ["<workbook><sheets/><calcPr calcId=\"123\" iterate=\"1\"/></workbook>",
            "<workbook><sheets/><extLst/></workbook>", "<workbook/>",
            "<s:workbook xmlns:s=\"urn:test\"><s:sheets/></s:workbook>"] {
            let patched = patch_workbook_calculation(xml, settings).unwrap();
            let mut reader = Reader::from_str(&patched);
            let mut count = 0;
            loop {
                match reader.read_event().unwrap() {
                    Event::Empty(e) if local_name(e.name().as_ref()) == "calcPr" => {
                        count += 1;
                        assert_eq!(get_attr(&e, "calcMode").as_deref(), Some("manual"));
                        assert_eq!(get_attr(&e, "calcOnSave").as_deref(), Some("0"));
                        assert_eq!(get_attr(&e, "calcCompleted").as_deref(), Some("0"));
                    }
                    Event::Eof => break,
                    _ => {}
                }
            }
            assert_eq!(count, 1, "{patched}");
            if xml.contains("calcId") { assert!(patched.contains("calcId=\"123\" iterate=\"1\"")); }
            if xml.contains("s:workbook") { assert!(patched.contains("<s:calcPr")); }
            if xml.contains("extLst") { assert!(patched.find("calcPr").unwrap() < patched.find("extLst").unwrap()); }
        }
        let data = include_bytes!("../../../tests/fixtures/vba/save_calculation.xlsm");
        for mode in [CalculationMode::Automatic, CalculationMode::Manual, CalculationMode::AutomaticExceptTables] {
            for on_save in [true, false] {
                let mut workbook = parse_xlsx(data).unwrap();
                workbook.calculation = CalculationSettings { mode, on_save, completed: Some(false), ..Default::default() };
                let mut editor = XlsxEditor::new(data).unwrap();
                editor.apply_workbook(&workbook).unwrap();
                let bytes = editor.save().unwrap();
                assert_eq!(crate::parse_xlsx_preserving_values(&bytes).unwrap().calculation, workbook.calculation);
            }
        }
    }

    #[test]
    fn rich_error_cache_markers_preserve_manual_calculation_and_existing_extensions() {
        let ns = SPREADSHEET_METADATA_NS;
        let feature_ns = "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures";
        let extension = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}";
        let sources = [
            format!(r#"<workbook xmlns="{ns}"><sheets/></workbook>"#),
            format!(r#"<workbook xmlns="{ns}"/>"#),
            format!(r#"<workbook xmlns="{ns}"><fileVersion appName="other" lastEdited="4" lowestEdited="4" custom="keep"/><sheets/><calcPr calcId="123" calcMode="manual" calcOnSave="0" iterate="1" calcCompleted="0"/><extLst/></workbook>"#),
            format!(r#"<s:workbook xmlns:s="{ns}" xmlns:f="{feature_ns}" xmlns:v="urn:vendor"><s:sheets/><s:calcPr calcId="999999" calcMode="autoNoTable" calcOnSave="0"/><s:extLst custom="keep"><s:ext uri="urn:opaque"><v:calcFeatures><v:feature name="microsoft.com:RD"/></v:calcFeatures></s:ext><s:ext uri="{extension}"><f:calcFeatures><f:feature name="microsoft.com:FV"/></f:calcFeatures></s:ext></s:extLst></s:workbook>"#),
            format!(r#"<workbook xmlns="{ns}"><sheets/><extLst><ext uri="{extension}"><f:calcFeatures xmlns:f="{feature_ns}"/></ext></extLst></workbook>"#),
        ];
        for (index, source) in sources.iter().enumerate() {
            let changed = patch_rich_error_workbook(source).unwrap();
            assert_eq!(patch_rich_error_workbook(&changed).unwrap(), changed, "idempotence: {index}");
            let (version, _, _, _) = direct_xml_child(&changed, "fileVersion").unwrap();
            let (start, end) = version.unwrap();
            assert!(changed[start..end].contains("appName=\"xl\""));
            let mut reader = quick_xml::NsReader::from_str(&changed);
            let mut rich = 0;
            loop {
                let (resolved, event) = reader.read_resolved_event().unwrap();
                match event {
                    Event::Start(ref tag) | Event::Empty(ref tag)
                        if matches!(resolved, quick_xml::name::ResolveResult::Bound(ref bound) if bound.as_ref() == feature_ns.as_bytes())
                            && tag.local_name().as_ref() == b"feature"
                            && get_attr(tag, "name").as_deref() == Some("microsoft.com:RD") => rich += 1,
                    Event::Eof => break,
                    _ => {}
                }
            }
            assert_eq!(rich, 1, "{changed}");
            assert_eq!(changed.matches(extension).count(), 1);
            if let Some((sheets, _)) = direct_xml_child(&changed, "sheets").unwrap().0 { assert!(start < sheets); }
            let calc = direct_xml_child(&changed, "calcPr").unwrap().0.unwrap();
            let extensions = direct_xml_child(&changed, "extLst").unwrap().0.unwrap();
            assert!(calc.0 < extensions.0);
            if index == 2 {
                for held in [r#"lastEdited="4""#, r#"lowestEdited="4""#, r#"custom="keep""#,
                    r#"calcMode="manual""#, r#"calcOnSave="0""#, r#"iterate="1""#, r#"calcCompleted="0""#] {
                    assert!(changed.contains(held), "{held}: {changed}");
                }
                assert!(changed.contains(r#"calcId="191029""#));
            }
            if index == 3 {
                for held in [r#"calcId="999999""#, r#"calcMode="autoNoTable""#, r#"custom="keep""#,
                    "<v:calcFeatures><v:feature name=\"microsoft.com:RD\"/></v:calcFeatures>",
                    "<f:feature name=\"microsoft.com:FV\"/>"] { assert!(changed.contains(held), "{changed}"); }
            }
        }
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(data).unwrap();
        editor.set_cell_value(0, 2, 3, CellEditValue::Error("#CALC!".into()));
        let saved = editor.save().unwrap();
        let mut archive = OoxmlArchive::new(&saved).unwrap();
        let xml = archive.read_part("xl/workbook.xml").unwrap();
        assert!(xml.contains("microsoft.com:RD") && xml.contains(r#"appName="xl""#) && xml.contains(r#"calcId="191029""#));
        let reopened = crate::parse_xlsx_preserving_values(&saved).unwrap();
        assert!(matches!(&reopened.sheets[0].rows.iter().find(|row| row.index == 2).unwrap().cells.iter().find(|cell| cell.col == 3).unwrap().value,
            crate::ir::CellValue::Error(text) if text == "#CALC!"));
    }

    #[test]
    fn rich_error_metadata_resolves_fields_and_preserves_other_records() {
        let metadata = format!(r#"<m:metadata xmlns:m="{SPREADSHEET_METADATA_NS}" xmlns:q="urn:other" xmlns:r="{RICH_DATA_NS}"><m:metadataTypes count="2"><m:metadataType name="OTHER"/><m:metadataType name="XLRICHVALUE"/></m:metadataTypes><m:futureMetadata name="OTHER" count="1"><m:bk/></m:futureMetadata><m:futureMetadata name="XLRICHVALUE" count="3"><m:bk><m:extLst><m:ext><r:rvb i="0"/></m:ext></m:extLst></m:bk><m:bk><m:extLst><m:ext><r:rvb i="1"/></m:ext></m:extLst></m:bk><m:bk><m:extLst><m:ext><r:rvb i="2"/></m:ext></m:extLst></m:bk></m:futureMetadata><m:valueMetadata count="4"><m:bk><m:rc t="1" v="0"/></m:bk><m:bk><m:rc t="1" v="0"/><m:rc t="2" v="0"/><m:extLst><q:keep attr="a&amp;b"/></m:extLst></m:bk><m:bk><m:rc t="2" v="1"/></m:bk><m:bk><m:rc t="2" v="2"/></m:bk></m:valueMetadata></m:metadata>"#);
        let mut structures = format!(r#"<rvStructures xmlns="{RICH_DATA_NS}" count="3"><s t="_error"><k n="colOffset" t="i"/><k n="errorType" t="i"/><k n="rwOffset" t="i"/><k n="subType" t="i"/></s><s t="unsupported"><k n="errorType" t="i"/></s><s t="_error"><k n="errorType" t="i"/><k n="subType" t="i"/></s><extLst><q:keep xmlns:q="urn:other"/></extLst></rvStructures>"#);
        let mut values = format!(r#"<rvData xmlns="{RICH_DATA_NS}" count="3"><rv s="0"><v>0</v><v>8</v><v>2</v><v>1</v></rv><rv s="1"><v>13</v></rv><rv s="2"><v>13</v><v>3</v></rv></rvData>"#);
        let info = crate::parser::rich_error_metadata(&metadata, &structures, &values).unwrap();
        assert_eq!(info.errors, [(2, "#SPILL!"), (4, "#CALC!")].into_iter().collect());
        let old_structures = structures.clone();let old_values = values.clone();
        let (changed, ids) = rich_error_records(&metadata, &mut structures, &mut values, &[(9, Some(2)), (13, Some(4))].into_iter().collect()).unwrap();
        assert_eq!(ids[&(13, Some(4))], 4);
        assert_eq!(structures, old_structures);
        assert_eq!(crate::parser::metadata_elements(&old_values).unwrap().iter().map(|(_, xml)| *xml).collect::<Vec<_>>(), crate::parser::metadata_elements(&values).unwrap().iter().take(3).map(|(_, xml)| *xml).collect::<Vec<_>>());
        let after = crate::parser::rich_error_metadata(&changed, &structures, &values).unwrap();
        let added = ids[&(9, Some(2))];
        assert_eq!(after.errors[&added], "#CONNECT!");
        assert!(after.blocks[added - 1].contains(r#"t="1" v="0""#));
        assert!(after.blocks[added - 1].contains(r#"<q:keep attr="a&amp;b"/>"#));
        assert_eq!(after.errors[&2], "#SPILL!");
        let (cleared, ids) = clear_rich_error_records(&changed, &structures, &values, &[2, 4].into_iter().collect()).unwrap();
        assert_eq!(ids[&4], None);
        let after = crate::parser::rich_error_metadata(&cleared, &structures, &values).unwrap();
        let retained = ids[&2].unwrap();
        assert!(!after.errors.contains_key(&retained));
        assert!(after.blocks[retained - 1].contains(r#"<q:keep attr="a&amp;b"/>"#));
        assert!(after.blocks[retained - 1].contains(r#"t="1" v="0""#));
    }

    #[test]
    fn rich_error_values_and_dynamic_caches_survive_save_and_replacement() {
        use crate::ir::CellValue;
        let errors = ["#SPILL!", "#CONNECT!", "#BLOCKED!", "#UNKNOWN!", "#FIELD!", "#CALC!"];
        let original = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut base = XlsxEditor::new(original).unwrap();
        for (index, error) in errors.iter().enumerate() { base.set_cell_value(0, index as u32 + 1, 25, CellEditValue::Error((*error).into())); }
        for (row, value) in [(1, 10.), (2, 20.), (3, 30.)] { base.set_cell_value(0, row, 0, CellEditValue::Number(value)); }
        base.set_cell_value(0, 2, 3, CellEditValue::Formula("A1:A3".into()));
        base.set_cell_value(0, 3, 3, CellEditValue::Number(99.));
        base.set_cell_value(0, 6, 3, CellEditValue::Formula("SEQUENCE(0)".into()));
        let seed = base.save().unwrap();
        let mut book = crate::parse_xlsx_preserving_values(&seed).unwrap();
        for index in [2, 6] {
            let cell = book.sheets[0].rows.iter_mut().find(|row| row.index == index).unwrap().cells.iter_mut().find(|cell| cell.col == 3).unwrap();
            cell.spill = crate::ir::Spill::Dynamic;
        }
        crate::formula::evaluate_workbook_formulas(&mut book);
        let mut editor = XlsxEditor::new(&seed).unwrap();editor.apply_workbook(&book).unwrap();
        let saved = editor.save().unwrap();
        let reopened = crate::parse_xlsx_preserving_values(&saved).unwrap();
        let cell = |book: &Workbook, row, col| book.sheets[0].rows.iter().find(|held| held.index == row).unwrap().cells.iter().find(|held| held.col == col).unwrap().clone();
        for (index, error) in errors.iter().enumerate() { assert!(same_value(&cell(&reopened, index as u32 + 1, 25).value, &CellValue::Error((*error).into()))); }
        for (row, error) in [(2, "#SPILL!"), (6, "#CALC!")] {
            let anchor = cell(&reopened, row, 3);
            assert!(same_value(&anchor.value, &CellValue::Error(error.into())), "{row}: {:?}", anchor.value);
            assert_eq!(anchor.spill, crate::ir::Spill::Dynamic);
            assert!(anchor.array_block.is_none());
        }
        let mut archive = OoxmlArchive::new(&saved).unwrap();
        let xml = archive.read_part("xl/worksheets/sheet1.xml").unwrap();
        assert!(!xml.contains("<v>#SPILL!</v>") && !xml.contains("<v>#CALC!</v>"));
        let metadata = archive.read_part("xl/metadata.xml").unwrap();
        assert!(metadata.contains("XLDAPR") && metadata.contains("XLRICHVALUE"));
        if std::env::var_os("OXI_EXPORT_RICH_XLSX").is_some() {
            println!("OXI_RICH_ERROR_XLSX_HEX:{}", saved.iter().map(|byte| format!("{byte:02x}")).collect::<String>());
        }
        if let Some(path) = std::env::var_os("OXI_IMPORT_RICH_XLSX") {
            let native = crate::parse_xlsx_preserving_values(&std::fs::read(path).unwrap()).unwrap();
            for (index, error) in errors.iter().enumerate() { assert!(same_value(&cell(&native, index as u32 + 2, 1).value, &CellValue::Error((*error).into()))); }
        }
        let mut editor = XlsxEditor::new(&saved).unwrap();
        for (row, value) in [(1, CellEditValue::Number(7.)), (2, CellEditValue::Error("#VALUE!".into())),
            (3, CellEditValue::String("#BLOCKED!".into())), (4, CellEditValue::Error("#FIELD!".into())),
            (5, CellEditValue::Empty), (6, CellEditValue::Error("#CALC!".into()))] { editor.set_cell_value(0, row, 25, value); }
        let replaced = editor.save().unwrap();
        let changed = crate::parse_xlsx_preserving_values(&replaced).unwrap();
        for (row, expected) in [(1, CellValue::Number(7.)), (2, CellValue::Error("#VALUE!".into())),
            (3, CellValue::String("#BLOCKED!".into())), (4, CellValue::Error("#FIELD!".into())),
            (5, CellValue::Empty), (6, CellValue::Error("#CALC!".into()))] { assert!(same_value(&cell(&changed, row, 25).value, &expected), "{row}: {:?}", cell(&changed, row, 25).value); }
        let mut unchanged = XlsxEditor::new(&saved).unwrap();unchanged.set_cell_value(0, 1, 25, CellEditValue::Error("#SPILL!".into()));
        let retained = unchanged.save().unwrap();
        let mut after = OoxmlArchive::new(&retained).unwrap();
        assert_eq!(after.read_part("xl/metadata.xml").unwrap(), metadata);
        for part in ["xl/richData/rdrichvalue.xml", "xl/richData/rdrichvaluestructure.xml", "xl/richData/rdRichValueTypes.xml"] { assert_eq!(after.read_part(part).unwrap(), archive.read_part(part).unwrap()); }
        // The metadata edit follows the original sheet when a copied sheet
        // is inserted before it. Added sheets get independent cell records.
        let mut moved = reopened.clone();
        let mut copy = moved.sheets[0].clone();copy.name = "Copied".into();copy.origin_id = None;
        moved.sheets[0].rows.iter_mut().find(|row| row.index == 1).unwrap().cells.iter_mut().find(|cell| cell.col == 25).unwrap().value = CellValue::Number(7.);
        moved.sheets.insert(0, copy);
        let mut editor = XlsxEditor::new(&saved).unwrap();editor.apply_workbook(&moved).unwrap();
        let moved = crate::parse_xlsx_preserving_values(&editor.save().unwrap()).unwrap();
        assert_eq!(moved.sheets[0].name, "Copied");
        assert!(same_value(&cell(&moved, 1, 25).value, &CellValue::Error("#SPILL!".into())));
        let original = moved.sheets[1].rows.iter().find(|row| row.index == 1).unwrap().cells.iter().find(|cell| cell.col == 25).unwrap();
        assert!(same_value(&original.value, &CellValue::Number(7.)));
    }

    #[test]
    fn literal_error_cells_replace_text_formulas_and_missing_cells() {
        use crate::ir::CellValue;
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut original = XlsxEditor::new(data).unwrap();
        original.set_cell_value(0, 2, 1, CellEditValue::CachedFormula {
            formula: "1+1".to_string(), reach: None, value: CellValue::Number(2.0),
        });
        let data = original.save().unwrap();
        let mut edited = crate::parse_xlsx_preserving_values(&data).unwrap();
        for (row, column, text) in [(1, 0, "#N/A"), (2, 1, "#DIV/0!"), (1, 25, "#VALUE!")] {
            let index = edited.sheets[0].rows.iter().position(|held| held.index==row).unwrap();
            let template = edited.sheets[0].rows[index].cells[0].clone();
            let cells = &mut edited.sheets[0].rows[index].cells;
            if !cells.iter().any(|cell| cell.col==column) {
                let mut cell = template;
                cell.col = column;
                cells.push(cell);
            }
            let cell = cells.iter_mut().find(|cell| cell.col==column).unwrap();
            cell.value = CellValue::Error(text.to_string());
            cell.formula = None;
            cell.array_block = None;
            cell.runs.clear();
        }
        let mut editor = XlsxEditor::new(&data).unwrap();
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let reopened = crate::parse_xlsx_preserving_values(&saved).unwrap();
        for (row, column, text) in [(1, 0, "#N/A"), (2, 1, "#DIV/0!"), (1, 25, "#VALUE!")] {
            let cell = reopened.sheets[0].rows.iter().find(|held| held.index==row).unwrap()
                .cells.iter().find(|cell| cell.col==column).unwrap();
            assert!(matches!(&cell.value, CellValue::Error(value) if value==text));
            assert!(cell.formula.is_none());
        }
    }

    #[test]
    fn workbook_formula_caches_survive_save_without_recalculation() {
        use crate::ir::CellValue;
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        for cached in [CellValue::Number(2.0), CellValue::String("cached & text".to_string()),
            CellValue::Boolean(true), CellValue::Error("#N/A".to_string())] {
            for array in [false, true] {
                let mut edited = parse_xlsx(data).unwrap();
                let cell = &mut edited.sheets[0].rows[1].cells[1];
                cell.formula = Some("A1*2".to_string());
                cell.value = cached.clone();
                cell.array_block = array.then_some((2, 1, 3, 1));
                let mut editor = XlsxEditor::new(data).unwrap();
                editor.apply_workbook(&edited).unwrap();
                let saved = editor.save().unwrap();
                let reopened = crate::parse_xlsx_preserving_values(&saved).unwrap();
                let cell = &reopened.sheets[0].rows[1].cells[1];
                assert!(same_value(&cell.value, &cached), "{cached:?}: {:?}", cell.value);
                assert_eq!(cell.formula.as_deref(), Some("A1*2"));
                assert_eq!(cell.array_block.is_some(), array);
            }
        }
    }

    #[test]
    fn dynamic_metadata_keeps_other_types_records_and_namespaces() {
        let xml = r#"<m:metadata xmlns:m="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><m:metadataTypes count="2"><m:metadataType name="OTHER"/><m:metadataType name="XLDAPR"/></m:metadataTypes><m:futureMetadata name="OTHER" count="1"><m:bk><q:payload xmlns:q="urn:custom" keep="a&amp;b"/></m:bk></m:futureMetadata><m:futureMetadata name="XLDAPR" count="1"><m:bk><m:extLst><m:ext uri="native"><xda:dynamicArrayProperties fDynamic="1"/></m:ext></m:extLst></m:bk></m:futureMetadata><m:cellMetadata count="2"><m:bk><m:rc t="1" v="0"/></m:bk><m:bk><m:rc t="1" v="0"/><m:rc t="2" v="0"/></m:bk></m:cellMetadata></m:metadata>"#;
        let requests = [(false, Some(2)), (true, None), (true, Some(1))].into_iter().collect();
        let (changed, ids) = dynamic_metadata_changes(xml, &requests).unwrap();
        let info = crate::parser::dynamic_metadata(&changed).unwrap();
        assert_eq!(info.type_index, Some(2));
        assert_eq!(info.type_count, 2);
        assert_eq!(info.future_count, 2);
        assert_eq!(info.cell_count, 5);
        assert!(changed.contains(r#"<q:payload xmlns:q="urn:custom" keep="a&amp;b"/>"#));
        let removed = ids[&(false, Some(2))].unwrap();
        assert!(!info.cells.contains(&removed));
        assert!(info.cell_blocks[removed - 1].contains(r#"t="1" v="0""#));
        let merged = ids[&(true, Some(1))].unwrap();
        assert!(info.cells.contains(&merged));
        assert!(info.cell_blocks[merged - 1].contains(r#"t="1" v="0""#));
        assert!(info.cells.contains(&ids[&(true, None)].unwrap()));
        let (empty, _) = dynamic_metadata_changes(r#"<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><metadataTypes count="0"/><cellMetadata count="0"/><extLst/></metadata>"#,
            &[(true, None)].into_iter().collect()).unwrap();
        assert_eq!(crate::parser::dynamic_metadata(&empty).unwrap().cells.into_iter().collect::<Vec<_>>(), vec![1]);
        assert!(empty.find("futureMetadata").unwrap() < empty.find("cellMetadata").unwrap());
    }

    #[test]
    fn dynamic_array_mode_and_exact_owners_survive_save_and_recalculation() {
        let original = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut base = XlsxEditor::new(original).unwrap();
        for (row, value) in [(1, 10.), (2, 20.), (3, 30.)] { base.set_cell_value(0, row, 0, CellEditValue::Number(value)); }
        for row in 3..=8 { base.set_cell_value(0, row, 3, CellEditValue::Empty); }
        let data = base.save().unwrap();
        let mut edited = parse_xlsx(&data).unwrap();
        let anchor = edited.sheets[0].rows.iter_mut().find(|r| r.index == 2).unwrap().cells.iter_mut().find(|c| c.col == 3).unwrap();
        anchor.formula = Some("A1:A3".into());anchor.formula_utf16 = None;anchor.array_block = None;anchor.spill = crate::ir::Spill::Dynamic;
        crate::formula::evaluate_workbook_formulas(&mut edited);
        assert_eq!(edited.sheets[0].cached_spills, Some(vec![((3, 1), (1, 3))]));
        let mut editor = XlsxEditor::new(&data).unwrap();editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let mut reopened = crate::parse_xlsx_preserving_values(&saved).unwrap();
        assert_eq!(reopened.sheets[0].cached_spills, Some(vec![((3, 1), (1, 3))]));
        for row in reopened.sheets[0].rows.iter().filter(|r| (2..=4).contains(&r.index)) {
            let cell = row.cells.iter().find(|c| c.col == 3).unwrap();
            assert!(cell.array_block.is_none());
            assert_eq!(cell.spill, if row.index == 2 { crate::ir::Spill::Dynamic } else { crate::ir::Spill::Member });
            if row.index != 2 { assert!(cell.formula.is_none()); }
        }
        crate::formula::evaluate_workbook_formulas(&mut reopened);
        assert!(matches!(reopened.sheets[0].rows.iter().find(|r| r.index == 2).unwrap().cells.iter().find(|c| c.col == 3).unwrap().value, crate::ir::CellValue::Number(10.)));
        let mut archive = OoxmlArchive::new(&saved).unwrap();
        assert!(archive.read_part("xl/metadata.xml").unwrap().contains("XLDAPR"));
        if std::env::var_os("OXI_EXPORT_DYNAMIC_XLSX").is_some() {
            let hex: String = saved.iter().map(|byte| format!("{byte:02x}")).collect();
            println!("OXI_DYNAMIC_XLSX_HEX:{hex}");
        }
        // Excel distinguishes Formula from Formula2 even when both getters
        // display SUM(SEQUENCE(3)+1). Array evaluation must survive saving
        // the scalar result, not just a visibly spilling range.
        for (formula, modern, saved_array) in [
            ("SUM(SEQUENCE(3)+1)", true, true),
            ("SUM(SEQUENCE(3)+1)", false, false),
            ("SUM({1;2;3}+1)", true, true),
            ("SUM({1;2;3}+1)", false, false),
            ("SUMPRODUCT(A1:A3*B1:B3)", true, true),
            ("SUMPRODUCT(A1:A3*B1:B3)", false, false),
            ("SUM(SEQUENCE(3))", true, false),
            ("SUM(@SEQUENCE(3)+1)", true, false),
            ("HSTACK(A1)", true, true),
            ("@HSTACK(A1)", false, false),
            ("@(SEQUENCE(3)+1)", true, true),
        ] {
            let mut book = parse_xlsx(&data).unwrap();
            let sheet = &mut book.sheets[0];
            sheet.cached_spills = Some(vec![]);
            let cell = sheet.rows.iter_mut().find(|row| row.index == 2).unwrap().cells.iter_mut().find(|cell| cell.col == 3).unwrap();
            cell.formula = Some(formula.into());cell.formula_utf16 = None;cell.array_block = None;
            cell.spill = if modern && oxicells_calc::lexer::requires_array_storage(formula, &|_| false) { crate::ir::Spill::Dynamic }
                else if modern { crate::ir::Spill::Scalar } else { crate::ir::Spill::None };
            cell.value = crate::ir::CellValue::Number(9.);
            let mut editor = XlsxEditor::new(&data).unwrap();editor.apply_workbook(&book).unwrap();
            let saved = editor.save().unwrap();
            let reopened = crate::parse_xlsx_preserving_values(&saved).unwrap();
            let cell = reopened.sheets[0].rows.iter().find(|row| row.index == 2).unwrap().cells.iter().find(|cell| cell.col == 3).unwrap();
            assert_eq!(cell.spill == crate::ir::Spill::Dynamic, saved_array, "{formula} modern={modern}");
            assert!(cell.array_block.is_none());
            assert!(matches!(cell.value, crate::ir::CellValue::Number(9.)));
        }
    }

    #[test]
    fn named_evaluation_modes_preserve_scope_and_saved_array_metadata() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let definitions = vec![("Rate".into(), "10".into()), ("Sheet1!Rate".into(), "SUM(1)".into()),
            ("Alias".into(), "Rate".into()), ("Cell".into(), "Sheet1!$A$1".into())];
        for (sheet, name, expected) in [("Sheet1", "Rate", true), ("Other", "Rate", false),
            ("Other", "'Sheet1'!Rate", true), ("Sheet1", "Alias", true), ("Sheet1", "Cell", false), ("Sheet1", "Missing", true)] {
            assert_eq!(named_formula_needs_intersection(&definitions, sheet, name), expected, "{sheet}!{name}");
        }
        for (body, needs_array) in [("10", false), ("-1.5", false), ("\"txt\"", false), ("TRUE", false),
            ("Sheet1!$A$1", false), ("Sheet1!$A$1:$A$1", true), ("SUM(1)", true), ("(10)", true), ("OtherName", true)] {
            for modern in [false, true] {
                let mut book = parse_xlsx(data).unwrap();book.sheets[0].name="Sheet1".into();book.defined_names.push(("Measured".into(), body.into()));
                let sheet = &mut book.sheets[0];sheet.cached_spills = Some(vec![]);
                let cell = &mut sheet.rows[1].cells[1];cell.formula = Some(if !modern && needs_array { "@Measured" } else { "Measured" }.into());
                cell.formula_utf16 = None;cell.array_block = None;cell.spill = if modern && oxicells_calc::lexer::named_formula_needs_intersection(body) { crate::ir::Spill::Dynamic }
                    else if modern { crate::ir::Spill::Scalar } else { crate::ir::Spill::None };
                cell.value = crate::ir::CellValue::Number(7.);
                let mut editor = XlsxEditor::new(data).unwrap();editor.apply_workbook(&book).unwrap();let saved = editor.save().unwrap();
                let reopened = crate::parse_xlsx_preserving_values(&saved).unwrap();let cell = &reopened.sheets[0].rows[1].cells[1];
                assert_eq!(cell.spill == crate::ir::Spill::Dynamic, modern && needs_array, "{body} modern={modern}");
                assert!(cell.array_block.is_none());
                assert_eq!(cell.formula.as_deref(), Some("Measured"), "{body} modern={modern}");
            }
        }
    }

    #[test]
    fn legacy_saved_formulas_omit_only_reproducible_intersections_and_keep_utf16() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        for (formula, modern, expected) in [
            ("@MYOWN(B1)", false, "MYOWN(B1)"),
            ("@B1:B3", false, "B1:B3"),
            ("SUM(@B1:B3)", false, "SUM(_xlfn.SINGLE(B1:B3))"),
            ("SUM(@MYOWN(B1))", false, "SUM(_xlfn.SINGLE(MYOWN(B1)))"),
            ("@MYOWN(B1)", true, "MYOWN(B1)"),
            ("@MYOWN(\"a@b!\")", false, "MYOWN(\"a@b!\")"),
            ("SUM(@B1:B3)&\"a@!\"", false, "SUM(_xlfn.SINGLE(B1:B3))&\"a@!\""),
            ("@Rate", false, "_xlfn.SINGLE(Rate)"),
            ("@Rate", true, "_xlfn.SINGLE(Rate)"),
        ] {
            for unpaired in [false, true] {
                let mut edited = parse_xlsx(data).unwrap();
                edited.defined_names.push(("Rate".into(), "10".into()));
                let cell = &mut edited.sheets[0].rows[1].cells[1];
                let mut units: Vec<u16> = formula.encode_utf16().collect();
                let mut expected: Vec<u16> = expected.encode_utf16().collect();
                if unpaired && formula.contains('!') {
                    for text in [&mut units, &mut expected] { *text.iter_mut().find(|u| **u == b'!' as u16).unwrap() = 0xd83d; }
                }
                cell.formula = Some(String::from_utf16_lossy(&units));
                cell.formula_utf16 = String::from_utf16(&units).err().map(|_| units);
                cell.spill = if modern && oxicells_calc::lexer::requires_array_storage(formula, &|_| false) { crate::ir::Spill::Dynamic }
                    else if modern { crate::ir::Spill::Scalar } else { crate::ir::Spill::None };
                cell.value = crate::ir::CellValue::Number(5.);
                let mut editor = XlsxEditor::new(data).unwrap();
                editor.apply_workbook(&edited).unwrap();
                let saved = editor.save().unwrap();
                let reopened = crate::parse_xlsx_preserving_values(&saved).unwrap();
                let cell = &reopened.sheets[0].rows[1].cells[1];
                let actual = cell.formula_units().map(|u| u.to_vec()).unwrap_or_else(|| cell.formula.as_ref().unwrap().encode_utf16().collect());
                assert_eq!(actual, expected, "{formula} modern={modern} unpaired={unpaired}");
                assert!(matches!(cell.value, crate::ir::CellValue::Number(5.)));
            }
        }
    }

    #[test]
    fn saved_lexical_functions_use_excel_namespaces_and_keep_exact_utf16() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        for (source, xml, cached) in [
            ("LET(Rate,100,Rate)", "_xlfn.LET(_xlpm.Rate,100,_xlpm.Rate)", crate::ir::CellValue::Number(100.)),
            ("LAMBDA(x,x+1)(2)", "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)", crate::ir::CellValue::Number(3.)),
            ("LET(value,\"a!@b\",value)", "_xlfn.LET(_xlpm.value,\"a_xD83D_@b\",_xlpm.value)", crate::ir::CellValue::Utf16(vec![97,0xd83d,64,98])),
        ] {
            let mut edited = parse_xlsx(data).unwrap();
            let cell = &mut edited.sheets[0].rows[1].cells[1];
            let mut units: Vec<u16> = source.encode_utf16().collect();
            if let Some(at) = units.iter().position(|u| *u == b'!' as u16) { units[at] = 0xd83d; }
            cell.formula = Some(String::from_utf16_lossy(&units));
            cell.formula_utf16 = String::from_utf16(&units).err().map(|_| units.clone());
            cell.value = cached.clone();
            let mut editor = XlsxEditor::new(data).unwrap();
            editor.apply_workbook(&edited).unwrap();
            let saved = editor.save().unwrap();
            let sheet = OoxmlArchive::new(&saved).unwrap().read_part("xl/worksheets/sheet1.xml").unwrap();
            assert!(quick_xml::escape::unescape(&sheet).unwrap().contains(xml), "{source}: {sheet}");
            let reopened = crate::parse_xlsx_preserving_values(&saved).unwrap();
            let cell = &reopened.sheets[0].rows[1].cells[1];
            let raw = cell.formula_units().map(|u| u.to_vec()).unwrap_or_else(|| cell.formula.as_ref().unwrap().encode_utf16().collect());
            let displayed = oxicells_calc::parser::transform_utf16_formula(&raw, oxicells_calc::lexer::formula_from_ooxml).unwrap();
            assert_eq!(displayed, units, "{source}");
            assert!(same_value(&cell.value, &cached));
        }
    }

    #[test]
    fn added_sheet_keeps_its_cell_styles_and_formula_cache() {
        let data = include_bytes!("../tests/fixtures/hidden_rows_cols.xlsx");
        let mut edited = parse_xlsx(data).unwrap();
        let mut added = edited.sheets[0].clone();
        added.name = "Added".to_string();
        added.origin_id = None;
        let cell = &mut added.rows[0].cells[0];
        cell.style.bold = true;
        cell.style.number_format = Some("0.000".to_string());
        cell.formula = Some("A1*2".to_string());
        cell.value = crate::ir::CellValue::Number(2.0);
        edited.sheets.push(added);
        let mut editor = XlsxEditor::new(data).unwrap();
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let reopened = crate::parse_xlsx_preserving_values(&saved).unwrap();
        let cell = &reopened.sheets[1].rows[0].cells[0];
        assert!(cell.style.bold);
        assert_eq!(cell.style.number_format.as_deref(), Some("0.000"));
        assert!(matches!(cell.value, crate::ir::CellValue::Number(2.0)));
        assert_eq!(cell.formula.as_deref(), Some("A1*2"));
    }

    /// A workbook that had no names gets a list, put where the schema wants
    /// it: straight after the sheets.
    #[test]
    fn a_workbook_with_no_names_is_given_a_list() {
        let xml = "<workbook><sheets><sheet name=\"S\"/></sheets><calcPr/></workbook>";
        let patched =
            patch_workbook_names(xml, &[("Sales".to_string(), "S!$A$1:$A$5".to_string())], &Default::default()).unwrap();
        assert_eq!(
            patched,
            "<workbook><sheets><sheet name=\"S\"/></sheets>\
             <definedNames><definedName name=\"Sales\">S!$A$1:$A$5</definedName></definedNames>\
             <calcPr/></workbook>"
        );
    }

    #[test]
    fn local_names_and_builtin_attributes_survive_rewriting() {
        let xml = "<workbook><sheets><sheet name=\"S\"/><sheet name=\"Other\"/></sheets><definedNames/></workbook>";
        let names = vec![("S!Print_Area".to_string(), "S!$A$1:$B$9".to_string()),
            ("Other!Local".to_string(), "Other!$C$1".to_string()),
            ("Sales".to_string(), "S!$A$1".to_string()), ("Criteria".to_string(), "S!$A$1".to_string())];
        let mut attributes = crate::ir::DefinedNameAttributes::new();
        attributes.insert("Other!Local".into(), BTreeMap::from([("hidden".into(), "1".into()), ("comment".into(), "a & b".into())]));
        let patched = patch_workbook_names(xml, &names, &attributes).unwrap();
        assert!(patched.contains("name=\"_xlnm.Print_Area\" localSheetId=\"0\""));
        assert!(patched.contains("name=\"Local\" localSheetId=\"1\" comment=\"a &amp; b\" hidden=\"1\""));
        assert!(patched.contains("<definedName name=\"Sales\">S!$A$1</definedName>"));
        assert!(patched.contains("<definedName name=\"_xlnm.Criteria\">S!$A$1</definedName>"));
        assert!(patch_workbook_names(xml, &[("Missing!Local".into(), "1".into())], &Default::default()).is_err());
    }

    #[test]
    fn a_list_with_nothing_left_in_it_goes() {
        let xml = "<workbook><sheets/><definedNames><definedName name=\"Sales\">S!$A$1</definedName></definedNames><calcPr/></workbook>";
        let patched = patch_workbook_names(xml, &[], &Default::default()).unwrap();
        assert_eq!(patched, "<workbook><sheets/><calcPr/></workbook>");
    }

    /// A hidden sheet stays hidden through a save.
    ///
    /// The state lives only in the workbook's sheet list, which is rewritten
    /// whole whenever a sheet is added, renamed or moved  Eso without carrying
    /// it, any of those would quietly bring a hidden sheet back into view.
    #[test]
    fn a_hidden_sheet_stays_hidden() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut edited = parse_xlsx(data).expect("should parse");
        let mut second = edited.sheets[0].clone();
        second.name = "Working".to_string();
        second.visibility = crate::ir::Visibility::Hidden;
        edited.sheets.push(second);

        let mut editor = XlsxEditor::new(data).expect("should open");
        editor.apply_workbook(&edited).expect("should apply");
        let saved = editor.save().expect("should save");

        let read_back = parse_xlsx(&saved).expect("should parse");
        assert_eq!(read_back.sheets.len(), 2);
        assert_eq!(read_back.sheets[0].visibility, crate::ir::Visibility::Visible);
        assert_eq!(read_back.sheets[1].visibility, crate::ir::Visibility::Hidden);
    }

    /// Deleting one sheet and adding another keeps the count the same, which
    /// used to fool the position match into handing the new sheet the deleted
    /// one's part -- its cells and its tab colour. The stable id must prevent
    /// that: the kept sheet keeps its own content, and the added one is empty.
    #[test]
    fn an_add_and_delete_that_keep_the_count_do_not_swap_parts() {
        let data = include_bytes!("../../../tests/fixtures/multi_sheet.xlsx");
        let original = parse_xlsx(data).expect("should parse");
        assert_eq!(
            original.sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(),
            ["Data", "Summary"]
        );

        let mut edited = original.clone();
        let mut fresh = edited.sheets[0].clone();
        fresh.name = "Report".to_string();
        fresh.rows = Vec::new();
        fresh.origin_id = None;
        fresh.tab_color = None;
        // Drop Summary, keep Data, append the fresh sheet: two sheets again.
        edited.sheets = vec![edited.sheets[0].clone(), fresh];

        let mut editor = XlsxEditor::new(data).expect("should open");
        editor.apply_workbook(&edited).expect("should apply");
        let saved = editor.save().expect("should save");
        let read_back = parse_xlsx(&saved).expect("should parse");

        assert_eq!(
            read_back.sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(),
            ["Data", "Report"]
        );
        let report_cells: usize = read_back.sheets[1].rows.iter().map(|r| r.cells.len()).sum();
        assert_eq!(report_cells, 0, "the added sheet inherited another sheet's cells");
        let data_cells: usize = read_back.sheets[0].rows.iter().map(|r| r.cells.len()).sum();
        assert!(data_cells > 0, "the kept sheet lost its cells");
    }

    /// Renaming a sheet rewrites the formulas that name it, so a reference the
    /// user never touched keeps pointing at the same data.
    #[test]
    fn renaming_a_sheet_rewrites_the_formulas_that_name_it() {
        let data = include_bytes!("../../../tests/fixtures/multi_sheet.xlsx");
        let original = parse_xlsx(data).expect("should parse");

        let mut edited = original.clone();
        // Summary's first cell now reads from Data; then Data is renamed.
        edited.sheets[1].rows[0].cells[0].formula = Some("=Data!B2".to_string());
        edited.sheets[0].name = "Facts".to_string();

        let mut editor = XlsxEditor::new(data).expect("should open");
        editor.apply_workbook(&edited).expect("should apply");
        let saved = editor.save().expect("should save");
        let read_back = parse_xlsx(&saved).expect("should parse");

        assert_eq!(read_back.sheets[0].name, "Facts");
        let formula = read_back.sheets[1]
            .rows
            .iter()
            .flat_map(|row| &row.cells)
            .find_map(|cell| cell.formula.clone())
            .expect("the summary cell kept a formula");
        assert_eq!(formula, "Facts!B2", "the reference was not rewritten");
    }

    /// A tab colour set in the editor is written on save, and taken away again
    /// on a later save -- including on a sheet whose part has no <sheetPr> yet.
    #[test]
    fn set_tab_color_helper_handles_the_sheet_pr_cases() {
        // No <sheetPr>: one is inserted right after <worksheet ...>.
        let none = r#"<worksheet xmlns="x"><sheetData/></worksheet>"#;
        let out = set_tab_color_in_worksheet(none, &Some("FF0000".to_string()));
        assert!(out.contains(r#"<sheetPr><tabColor rgb="FFFF0000"/></sheetPr>"#), "insert: {out}");
        // Self-closing <sheetPr/>: expanded.
        let sc = r#"<worksheet xmlns="x"><sheetPr codeName="S"/><sheetData/></worksheet>"#;
        let out = set_tab_color_in_worksheet(sc, &Some("00B050".to_string()));
        assert!(out.contains(r#"<sheetPr codeName="S"><tabColor rgb="FF00B050"/></sheetPr>"#), "self-closing: {out}");
        // Existing tabColor: replaced.
        let ex = r#"<worksheet xmlns="x"><sheetPr><tabColor rgb="FFAAAAAA"/></sheetPr><sheetData/></worksheet>"#;
        let out = set_tab_color_in_worksheet(ex, &Some("123456".to_string()));
        assert!(out.contains(r#"<tabColor rgb="FF123456"/>"#) && !out.contains("FFAAAAAA"), "replace: {out}");
        // Remove.
        let out = set_tab_color_in_worksheet(ex, &None);
        assert!(!out.contains("tabColor"), "remove: {out}");
    }

    #[test]
    fn a_tab_colour_set_and_removed_survives_a_save() {
        let data = include_bytes!("../../../tests/fixtures/multi_sheet.xlsx");
        let original = parse_xlsx(data).expect("should parse");
        assert_eq!(original.sheets[0].tab_color, None);

        let mut edited = original.clone();
        edited.sheets[0].tab_color = Some("FF0000".to_string());
        let mut editor = XlsxEditor::new(data).expect("open");
        editor.apply_workbook(&edited).expect("apply");
        let painted = editor.save().expect("save");
        let read = parse_xlsx(&painted).expect("parse");
        assert_eq!(read.sheets[0].tab_color.as_deref(), Some("FF0000"), "colour not written");

        let mut cleared = read.clone();
        cleared.sheets[0].tab_color = None;
        let mut editor = XlsxEditor::new(&painted).expect("open");
        editor.apply_workbook(&cleared).expect("apply");
        let bare = editor.save().expect("save");
        let read2 = parse_xlsx(&bare).expect("parse");
        assert_eq!(read2.sheets[0].tab_color, None, "colour not removed");
    }

    /// A name written and read back comes through the file whole.
    #[test]
    fn a_name_reaches_the_saved_file() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut edited = parse_xlsx(data).expect("should parse");
        let sheet = edited.sheets[0].name.clone();
        edited
            .defined_names
            .push(("Sales".to_string(), format!("{sheet}!$A$1:$A$5")));

        let mut editor = XlsxEditor::new(data).expect("should open");
        editor.apply_workbook(&edited).expect("should apply");
        let saved = editor.save().expect("should save");

        let read_back = parse_xlsx(&saved).expect("should parse");
        assert_eq!(
            read_back.defined_names,
            vec![("Sales".to_string(), format!("{sheet}!$A$1:$A$5"))]
        );
    }

    #[test]
    fn workbook_indexed_colors_are_patched_and_round_trip() {
        let data = include_bytes!("../tests/fixtures/wrapped.xlsx");
        let mut edited = parse_xlsx(data).expect("should parse");
        let mut colors = vec!["000000".to_string(); 64];
        colors[27] = "FF0000".into();
        colors[28] = "7B2D43".into();
        edited.indexed_colors = colors.clone();

        let mut editor = XlsxEditor::new(data).expect("should open");
        editor.apply_workbook(&edited).expect("should apply palette");
        assert_eq!(editor.indexed_colors.as_deref(), Some(colors.as_slice()));
        let saved = editor.save().expect("should save");
        let read_back = parse_xlsx(&saved).expect("should parse saved palette");
        assert_eq!(read_back.indexed_colors, colors);
    }

    #[test]
    fn workbook_indexed_colors_create_missing_styles_part() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut edited = parse_xlsx(data).expect("should parse");
        let mut colors = vec!["000000".to_string(); 64];
        colors[27] = "7B2D43".into();
        edited.indexed_colors = colors.clone();

        let mut editor = XlsxEditor::new(data).expect("should open");
        editor.apply_workbook(&edited).expect("should apply palette");
        let saved = editor.save().expect("should save");
        let read_back = parse_xlsx(&saved).expect("should parse saved palette");
        assert_eq!(read_back.indexed_colors, colors);

        let mut package = ZipArchive::new(Cursor::new(saved)).expect("should open package");
        assert!(package.by_name("xl/styles.xml").is_ok());
        let mut relationships = String::new();
        package
            .by_name("xl/_rels/workbook.xml.rels")
            .unwrap()
            .read_to_string(&mut relationships)
            .unwrap();
        assert!(relationships.contains("/styles"));
        let mut content_types = String::new();
        package
            .by_name("[Content_Types].xml")
            .unwrap()
            .read_to_string(&mut content_types)
            .unwrap();
        assert!(content_types.contains("/xl/styles.xml"));
    }

    #[test]
    fn patch_indexed_colors_xml_replaces_or_inserts_before_extensions() {
        let colors = vec!["123456".to_string(), "AABBCC".to_string()];
        let existing = r#"<styleSheet><colors><indexedColors><rgbColor rgb="00000000"/></indexedColors></colors><extLst/></styleSheet>"#;
        let patched = patch_indexed_colors_xml(existing, &colors).unwrap();
        assert!(patched.contains("<rgbColor rgb=\"00123456\"/><rgbColor rgb=\"00AABBCC\"/>"));
        assert!(patched.find("</colors>").unwrap() < patched.find("<extLst").unwrap());

        let missing = "<styleSheet><fonts></fonts><cellXfs></cellXfs><extLst/></styleSheet>";
        let patched = patch_indexed_colors_xml(missing, &colors).unwrap();
        assert!(patched.find("</cellXfs>").unwrap() < patched.find("<colors>").unwrap());
        assert!(patched.find("</colors>").unwrap() < patched.find("<extLst").unwrap());
    }

    #[test]
    fn test_editor_round_trip() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let editor = XlsxEditor::new(data).expect("should open");
        let saved = editor.save().expect("should save");
        let wb = parse_xlsx(&saved).expect("should parse");
        assert_eq!(wb.sheets.len(), 1);
        assert_eq!(wb.sheets[0].name, "Sales");
    }

    #[test]
    fn test_editor_change_cell() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(data).expect("should open");

        // Change cell A1 (row 1, col 0)  E"Product" header
        editor.set_cell(0, 1, 0, "Item".to_string());

        let saved = editor.save().expect("should save");
        let wb = parse_xlsx(&saved).expect("should parse");

        let row1 = &wb.sheets[0].rows[0];
        assert!(matches!(&row1.cells[0].value, crate::ir::CellValue::String(s) if s == "Item"));
    }

    #[test]
    fn test_editor_inserts_missing_cells_and_rows() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(data).expect("should open");

        editor.set_cell(0, 1, 25, "New column".to_string());
        editor.set_cell(0, 4, 1, "New row".to_string());

        let saved = editor.save().expect("should save");
        let wb = parse_xlsx(&saved).expect("should parse");
        let sheet = &wb.sheets[0];
        let inserted_cell = sheet.rows[0]
            .cells
            .iter()
            .find(|cell| cell.col == 25)
            .expect("Z1 should be inserted");
        assert!(matches!(
            &inserted_cell.value,
            crate::ir::CellValue::String(value) if value == "New column"
        ));
        let inserted_row = sheet
            .rows
            .iter()
            .find(|row| row.index == 4)
            .expect("row 4 should be inserted");
        assert!(matches!(
            &inserted_row.cells[0].value,
            crate::ir::CellValue::String(value) if value == "New row"
        ));
    }

    #[test]
    fn test_editor_preserves_typed_cell_edits() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(data).expect("should open");
        editor.set_cell_value(0, 1, 20, CellEditValue::Number(12.5));
        editor.set_cell_value(0, 1, 21, CellEditValue::Boolean(true));
        editor.set_cell_value(0, 1, 22, CellEditValue::Empty);

        let saved = editor.save().expect("should save");
        let workbook = parse_xlsx(&saved).expect("should parse");
        let row = &workbook.sheets[0].rows[0];
        assert!(matches!(row.cells.iter().find(|cell| cell.col == 20).unwrap().value, crate::ir::CellValue::Number(12.5)));
        assert!(matches!(row.cells.iter().find(|cell| cell.col == 21).unwrap().value, crate::ir::CellValue::Boolean(true)));
        assert!(matches!(row.cells.iter().find(|cell| cell.col == 22).unwrap().value, crate::ir::CellValue::Empty));
    }

    #[test]
    fn hyperlink_addition_roundtrips_external_internal_and_external_location() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(data).expect("should open");
        let mut edited = editor.workbook().clone();
        edited.sheets[0].hyperlinks = vec![
            Hyperlink {
                cell: (0, 0),
                end_cell: Some((2, 3)),
                target: "https://example.test/path?a=1&b=2".into(),
                internal: false,
                sub_address: Some("Section 1".into()),
                tooltip: Some("open & read".into()),
            },
            Hyperlink {
                cell: (1, 2),
                end_cell: None,
                target: "'Sheet 1'!$A$1".into(),
                internal: true,
                sub_address: None,
                tooltip: None,
            },
        ];
        editor.apply_workbook(&edited).expect("record links");
        let saved = editor.save().expect("save hyperlinks");
        let reopened = parse_xlsx(&saved).expect("reopen hyperlinks");
        assert_eq!(reopened.sheets[0].hyperlinks, edited.sheets[0].hyperlinks);

        let mut archive = ZipArchive::new(Cursor::new(saved)).expect("zip output");
        let mut rels = String::new();
        archive.by_name("xl/worksheets/_rels/sheet1.xml.rels").expect("sheet relationships")
            .read_to_string(&mut rels).expect("read relationships");
        assert!(rels.contains("TargetMode=\"External\""), "{rels}");
        assert!(rels.contains("https://example.test/path?a=1&amp;b=2"), "{rels}");
    }

    #[test]
    fn internal_hyperlink_saves_without_a_relationship_part() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(data).unwrap();
        let mut edited = editor.workbook().clone();
        edited.sheets[0].hyperlinks.push(Hyperlink {
            cell: (0, 0), end_cell: None, target: "Sheet1!A1".into(), internal: true,
            sub_address: None, tooltip: None,
        });
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let reopened = parse_xlsx(&saved).unwrap();
        assert_eq!(reopened.sheets[0].hyperlinks, edited.sheets[0].hyperlinks);

        let mut delete = XlsxEditor::new(&saved).unwrap();
        let mut no_links = delete.workbook().clone();
        no_links.sheets[0].hyperlinks.clear();
        delete.apply_workbook(&no_links).unwrap();
        let deleted = delete.save().unwrap();
        let reopened = parse_xlsx(&deleted).unwrap();
        assert!(reopened.sheets[0].hyperlinks.is_empty());
    }

    #[test]
    fn hyperlinks_on_a_new_sheet_are_written_with_the_new_sheet_part() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(data).unwrap();
        let mut edited = editor.workbook().clone();
        let mut added = edited.sheets[0].clone();
        added.name = "Added".into();
        added.origin_id = None;
        added.hyperlinks = vec![Hyperlink {
            cell: (2, 3),
            end_cell: Some((4, 5)),
            target: "https://new-sheet.example/path?q=1&v=2".into(),
            internal: false,
            sub_address: Some("Details".into()),
            tooltip: Some("open details".into()),
        }];
        let expected = added.hyperlinks.clone();
        edited.sheets.push(added);
        editor.apply_workbook(&edited).unwrap();
        let saved = editor.save().unwrap();
        let reopened = parse_xlsx(&saved).unwrap();
        assert_eq!(reopened.sheets[1].hyperlinks, expected);
    }

    #[test]
    fn test_editor_preserves_formula_edits() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut editor = XlsxEditor::new(data).expect("should open");
        editor.set_cell_value(
            0,
            1,
            23,
            CellEditValue::Formula("=SUM(A2:A3)".to_string()),
        );

        let saved = editor.save().expect("should save");
        let workbook = parse_xlsx(&saved).expect("should parse");
        let cell = workbook.sheets[0].rows[0]
            .cells
            .iter()
            .find(|cell| cell.col == 23)
            .unwrap();
        assert_eq!(cell.formula.as_deref(), Some("SUM(A2:A3)"));
    }

    const STYLES: &str = concat!(
        r#"<styleSheet><numFmts count="0"/><fonts count="1"><font><sz val="11"/></font></fonts>"#,
        r#"<fills count="1"><fill><patternFill patternType="none"/></fill></fills>"#,
        r#"<borders count="1"><border/></borders>"#,
        r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>"#,
        r#"</styleSheet>"#
    );

    #[test]
    fn styles_are_appended_and_their_indices_reported() {
        let bold = CellStyle {
            bold: true,
            ..CellStyle::default()
        };
        let (patched, indices) = patch_styles_xml(STYLES, &[bold]).expect("appends");

        assert_eq!(indices, vec![1]);
        assert!(patched.contains("<font><b/></font>"));
        assert!(patched.contains(r#"<fonts count="2">"#));
        assert!(patched.contains(r#"<cellXfs count="2">"#));
        assert!(patched.contains(r#"fontId="1" fillId="1" borderId="1""#));
        // The entry that was already there is left alone.
        assert!(patched.contains(r#"<font><sz val="11"/></font>"#));
    }

    #[test]
    fn a_custom_number_format_gets_an_id_of_its_own() {
        let styled = CellStyle {
            number_format: Some("0.000".to_string()),
            ..CellStyle::default()
        };
        let (patched, _) = patch_styles_xml(STYLES, &[styled]).expect("appends");
        assert!(patched.contains(r#"<numFmt numFmtId="164" formatCode="0.000"/>"#));
        assert!(patched.contains(r#"numFmtId="164""#));
        assert!(patched.contains(r#"<numFmts count="1">"#));
    }

    #[test]
    fn a_well_known_number_format_reuses_its_id() {
        let styled = CellStyle {
            number_format: Some("0.00".to_string()),
            ..CellStyle::default()
        };
        let (patched, _) = patch_styles_xml(STYLES, &[styled]).expect("appends");
        assert!(!patched.contains("<numFmt "));
        assert!(patched.contains(r#"numFmtId="2""#));
    }

    #[test]
    fn alignment_and_colour_reach_the_style_sheet() {
        let styled = CellStyle {
            horizontal_align: Some("center".to_string()),
            bg_color: Some("FFFF00".to_string()),
            font_color: Some("FF0000".to_string()),
            border_top: Some(BorderLine { style: "thin".to_string(), color: None }),
            ..CellStyle::default()
        };
        let (patched, _) = patch_styles_xml(STYLES, &[styled]).expect("appends");
        assert!(patched.contains(r#"<alignment horizontal="center"/>"#));
        assert!(patched.contains(r#"<fgColor rgb="FFFFFF00"/>"#));
        assert!(patched.contains(r#"<color rgb="FFFF0000"/>"#));
        assert!(patched.contains(r#"<top style="thin">"#));
    }

    /// Wrapping, the vertical part and an indent were read out of a file and
    /// never written back, so a cell told to wrap came back unwrapped. They
    /// share one element with the horizontal part rather than each getting one.
    #[test]
    fn everything_an_alignment_says_is_written_in_one_element() {
        let laid_out = CellStyle {
            horizontal_align: Some("left".to_string()),
            vertical_align: Some("top".to_string()),
            wrap_text: true,
            shrink_to_fit: true,
            indent: 2,
            ..CellStyle::default()
        };
        let (patched, _) = patch_styles_xml(STYLES, &[laid_out]).expect("appends");
        assert!(patched.contains(
            r#"<alignment horizontal="left" vertical="top" wrapText="1" shrinkToFit="1" indent="2"/>"#
        ));
    }

    /// A style that lines nothing up writes no `<alignment>` at all: an empty
    /// one would claim the cell states something it does not.
    #[test]
    fn a_style_that_lines_nothing_up_writes_no_alignment() {
        let plain = CellStyle { bold: true, ..CellStyle::default() };
        let (patched, _) = patch_styles_xml(STYLES, &[plain]).expect("appends");
        assert!(!patched.contains("<alignment"));
    }

    #[test]
    fn asking_for_no_styles_changes_nothing() {
        let (patched, indices) = patch_styles_xml(STYLES, &[]).expect("does nothing");
        assert_eq!(patched, STYLES);
        assert!(indices.is_empty());
    }

    /// The patcher takes everything a sheet changes; these tests only change cells.
    fn cells_only<'a>(
        cells: &'a HashMap<(u32, u32), &'a CellEditValue>,
    ) -> SheetEdits<'a> {
        static NO_ROWS: std::sync::OnceLock<BTreeMap<u32, RowWant>> =
            std::sync::OnceLock::new();
        static NO_STYLES: std::sync::OnceLock<BTreeMap<(u32, u32), u32>> =
            std::sync::OnceLock::new();
        SheetEdits {
            cells,
            rows: NO_ROWS.get_or_init(BTreeMap::new),
            cols: nothing_cols(),
            merges: None,
            styles: NO_STYLES.get_or_init(BTreeMap::new),
            filter: None,
            panes: None,
            hyperlinks: None,
            hyperlink_ids: &[],
        }
    }

    #[test]
    fn worksheet_patch_replaces_an_existing_formula() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><f>OLD()</f><v>1</v></c></row></sheetData></worksheet>"#;
        let formula = CellEditValue::Formula("=SUM(B1:B2)".to_string());
        let edits = HashMap::from([((1, 0), &formula)]);

        let patched = patch_worksheet_xml(xml, &cells_only(&edits)).expect("should replace formula");
        assert!(patched.contains("<f>SUM(B1:B2)</f>"));
        assert!(!patched.contains("OLD()"));
        assert_eq!(patched.matches("<f>").count(), 1);
    }

    #[test]
    fn renamed_table_formula_metadata_follows_references_without_changing_literals() {
        let original = r#"<table name="Other" displayName="Other"><tableColumns count="1"><tableColumn id="1" name="Value"><calculatedColumnFormula>SUM(OldTable[X])&amp;"OldTable[X]a_xD83D_b"</calculatedColumnFormula><totalsRowFormula>SUM(OldTable[X])</totalsRowFormula></tableColumn></tableColumns></table>"#;
        let rewritten = rewrite_table_formula_xml(original, &[TableFormulaRename {
            old: "OldTable".into(), new: "NewTable".into(), columns: vec![("X".into(), "Z".into())],
        }]).unwrap();
        assert!(rewritten.contains("SUM(NewTable[Z])"));
        assert!(rewritten.contains("OldTable[X]a_xD83D_b"));
        assert_eq!(rewritten.matches("SUM(NewTable[Z])").count(), 2);
    }

    #[test]
    fn table_edits_preserve_column_formulas_extensions_and_style_flags() {
        let table = crate::ir::Table {
            origin_part: None,
            name: "Renamed".into(), columns: vec!["X & Y".into()],
            start_row: 1, start_col: 0, end_row: 5, end_col: 0,
            style: Some("TableStyleMedium2".into()), header_rows: 1, totals_rows: 0, totals: Vec::new(), banded_rows: true,
            banded_columns: true, emphasize_first_column: true, emphasize_last_column: false,
            show_auto_filter: true, hidden_filter_buttons: Vec::new(), insert_row: false,
            auto_filter_xml: Some("<autoFilter ref=\"A1:A4\"><filterColumn colId=\"0\"><filters><filter val=\"1\"/></filters></filterColumn></autoFilter>".into()),
            accent: None, band: None, rule: None, outline: None,
        };
        let original = r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:custom="urn:keep" custom:flag="yes" id="7" name="Old" displayName="Old" ref="A1:A4"><autoFilter ref="A1:A4"><filterColumn colId="0"><filters><filter val="1"/></filters></filterColumn></autoFilter><tableColumns count="1"><tableColumn id="42" name="Old"><calculatedColumnFormula>A1*2</calculatedColumnFormula><extLst><ext uri="keep-column"/></extLst></tableColumn></tableColumns><tableStyleInfo name="TableStyleMedium1" showFirstColumn="1" showColumnStripes="1"/><extLst><ext uri="keep-table"/></extLst></table>"#;
        let edited = table_xml(&table, 7, Some(original)).unwrap();
        assert!(edited.contains("custom:flag=\"yes\""));
        assert!(edited.contains("<calculatedColumnFormula>A1*2</calculatedColumnFormula>"));
        assert!(edited.contains("uri=\"keep-column\""));
        assert!(edited.contains("uri=\"keep-table\""));
        assert!(edited.contains("showFirstColumn=\"1\""));
        assert!(edited.contains("showColumnStripes=\"1\""));
        assert!(edited.contains("id=\"42\" name=\"X &amp; Y\""));
        assert!(edited.contains("<filter val=\"1\"/>"));
        assert_eq!(table_root_attribute(&edited, "ref").as_deref(), Some("A1:A5"));
    }

    #[test]
    fn cached_formula_replaces_empty_xml_cells_and_values() {
        let cached = CellEditValue::CachedFormula { formula: "2+2".to_string(), reach: None, value: crate::ir::CellValue::Number(99.0) };
        let edits = HashMap::from([((1, 0), &cached)]);
        let styles = BTreeMap::from([((1, 0), 9)]);
        for cell in ["<c r=\"A1\" s=\"0\"/>", "<c r=\"A1\"><f>OLD()</f><v/></c>", "<c r=\"A1\"><f/><v>1</v></c>"] {
            let xml = format!("<worksheet><sheetData><row r=\"1\">{cell}</row></sheetData></worksheet>");
            let patched = patch_worksheet_xml(&xml, &SheetEdits { styles: &styles, ..cells_only(&edits) }).unwrap();
            assert!(patched.contains("s=\"9\""), "{patched}");
            assert!(patched.contains("<v>99</v>"), "{patched}");
            assert!(patched.contains("<f>2+2</f>"), "{patched}");
            assert_eq!(patched.matches("<f").count(), 1, "{patched}");
            assert_eq!(patched.matches("<v").count(), 1, "{patched}");
        }
        let xml = "<worksheet><sheetData><row r=\"1\"><c r=\"A1\" s=\"0\"/></row></sheetData></worksheet>";
        let edits = HashMap::new();
        let patched = patch_worksheet_xml(xml, &SheetEdits { styles: &styles, ..cells_only(&edits) }).unwrap();
        assert!(patched.contains("s=\"9\""));
        assert!(!patched.contains("s=\"0\""));
    }

    #[test]
    fn the_calculation_chain_is_not_left_pointing_at_nothing() {
        // Writing a value over a formula takes the formula away. A chain still
        // naming that cell makes Excel refuse the WHOLE workbook  E`fies_t2`
        // has exactly one entry, C1, and replacing C1's `<f>C9</f>` lost the
        // file. The part is dropped, and so is anything that declares it.
        let types = concat!(
            r#"<Types><Override PartName="/xl/workbook.xml" ContentType="a"/>"#,
            r#"<Override PartName="/xl/calcChain.xml" ContentType="b"/>"#,
            r#"<Override PartName="/xl/styles.xml" ContentType="c"/></Types>"#,
        );
        let out = without_calc_chain(types);
        assert!(!out.contains("calcChain"), "{out}");
        assert!(out.contains("workbook.xml") && out.contains("styles.xml"), "{out}");
        let rels = concat!(
            r#"<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/>"#,
            r#"<Relationship Id="rId8" Target="calcChain.xml"/></Relationships>"#,
        );
        let out = without_calc_chain(rels);
        assert!(!out.contains("calcChain"), "{out}");
        assert!(out.contains("rId1"), "{out}");
    }

    #[test]
    fn a_replaced_cell_keeps_one_value_element() {
        // The check for "this cell had no `<v>` of its own" was reading the
        // flag that says "we are inside a `<v>` right now", which `</v>` had
        // already cleared  Eso the new value went in twice, and Excel refuses
        // to open a workbook whose cell holds two of them. `001250007` was
        // that workbook: 83 sheets, and the whole file would not load.
        let xml = concat!(
            r#"<worksheet><sheetData><row r="2">"#,
            r#"<c r="B2" s="1" t="s"><v>877</v></c>"#,
            r#"</row></sheetData></worksheet>"#,
        );
        let value = CellEditValue::String("fresh".to_string());
        let edits = HashMap::from([((2, 1), &value)]);
        let patched = patch_worksheet_xml(xml, &cells_only(&edits)).expect("should patch");
        assert_eq!(patched.matches("<v>").count(), 1, "{patched}");
        assert!(patched.contains("<v>fresh</v>"), "{patched}");
        assert!(!patched.contains("877"), "{patched}");
    }

    fn nothing_cols() -> &'static BTreeMap<u32, ColumnWant> {
        static NO_COLS: std::sync::OnceLock<BTreeMap<u32, ColumnWant>> =
            std::sync::OnceLock::new();
        NO_COLS.get_or_init(BTreeMap::new)
    }

    fn nothing_rows() -> &'static BTreeMap<u32, RowWant> {
        static NO_ROWS: std::sync::OnceLock<BTreeMap<u32, RowWant>> =
            std::sync::OnceLock::new();
        NO_ROWS.get_or_init(BTreeMap::new)
    }

    /// The edits for one sheet's columns and nothing else.
    fn cols_only(cols: &BTreeMap<u32, ColumnWant>) -> SheetEdits<'_> {
        static NO_CELLS: std::sync::OnceLock<HashMap<(u32, u32), &CellEditValue>> =
            std::sync::OnceLock::new();
        static NO_STYLES: std::sync::OnceLock<BTreeMap<(u32, u32), u32>> =
            std::sync::OnceLock::new();
        SheetEdits {
            cells: NO_CELLS.get_or_init(HashMap::new),
            rows: nothing_rows(),
            cols,
            merges: None,
            styles: NO_STYLES.get_or_init(BTreeMap::new),
            filter: None,
            panes: None,
            hyperlinks: None,
            hyperlink_ids: &[],
        }
    }

    /// The `<cols>` element of a patched sheet, or what stands in its place.
    fn cols_of(xml: &str, cols: &BTreeMap<u32, ColumnWant>) -> String {
        let patched = patch_worksheet_xml(xml, &cols_only(cols)).expect("should patch");
        match (patched.find("<cols>"), patched.find("</cols>")) {
            (Some(from), Some(to)) => patched[from..to + "</cols>".len()].to_string(),
            _ => format!("(no cols in {patched})"),
        }
    }

    const BARE: &str =
        r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;

    /// The edits for one sheet's rows and nothing else.
    fn rows_only(rows: &BTreeMap<u32, RowWant>) -> SheetEdits<'_> {
        static NO_CELLS: std::sync::OnceLock<HashMap<(u32, u32), &CellEditValue>> =
            std::sync::OnceLock::new();
        static NO_STYLES: std::sync::OnceLock<BTreeMap<(u32, u32), u32>> =
            std::sync::OnceLock::new();
        SheetEdits {
            cells: NO_CELLS.get_or_init(HashMap::new),
            rows,
            cols: nothing_cols(),
            merges: None,
            styles: NO_STYLES.get_or_init(BTreeMap::new),
            filter: None,
            panes: None,
            hyperlinks: None,
            hyperlink_ids: &[],
        }
    }

    /// The `<row>` start tags of a patched sheet.
    fn row_tags(xml: &str, rows: &BTreeMap<u32, RowWant>) -> String {
        let patched = patch_worksheet_xml(xml, &rows_only(rows)).expect("should patch");
        patched
            .match_indices("<row ")
            .map(|(at, _)| {
                let end = patched[at..].find('>').expect("a row tag closes") + at;
                patched[at..=end].to_string()
            })
            .collect::<Vec<_>>()
            .join("")
    }

    const THREE_ROWS: &str = concat!(
        r#"<worksheet><sheetData>"#,
        r#"<row r="1" spans="1:1"><c r="A1"><v>1</v></c></row>"#,
        r#"<row r="2" spans="1:1"><c r="A2"><v>2</v></c></row>"#,
        r#"<row r="3" spans="1:1" ht="30" customHeight="1"><c r="A3"><v>3</v></c></row>"#,
        r#"</sheetData></worksheet>"#,
    );

    #[test]
    fn a_height_someone_chose_is_written_as_chosen() {
        // Excel writes `ht="30" customHeight="1"` for a row dragged to a size,
        // and the flag is what stops it working the height out again on open.
        let want = BTreeMap::from([(1, RowWant {
            hidden: None, height: Some(Some(30.0)), pinned: true,
         style: None, })]);
        let out = row_tags(THREE_ROWS, &want);
        assert!(out.contains(r#"<row r="1" spans="1:1" ht="30" customHeight="1">"#), "{out}");
    }

    #[test]
    fn a_height_worked_out_rather_than_chosen_goes_in_without_the_flag() {
        // An autofitted row carries `ht="44.25"` and no flag: the number is
        // what the content needed, and Excel will work it out afresh.
        let want = BTreeMap::from([(2, RowWant {
            hidden: None, height: Some(Some(44.25)), pinned: false,
         style: None, })]);
        let out = row_tags(THREE_ROWS, &want);
        assert!(out.contains(r#"<row r="2" spans="1:1" ht="44.25">"#), "{out}");
        // Row 3 has a chosen height of its own, so the flag is looked for on
        // row 2's tag rather than anywhere in the sheet.
        let two = out.split(r#"<row r="3""#).next().unwrap();
        assert!(!two.contains("customHeight"), "{out}");
    }

    #[test]
    fn putting_a_row_back_on_automatic_takes_its_height_away() {
        // Autofitting a row whose content fits the default drops `ht`
        // altogether rather than writing the default number down.
        let want = BTreeMap::from([(3, RowWant {
            hidden: None, height: Some(None), pinned: false,
         style: None, })]);
        let out = row_tags(THREE_ROWS, &want);
        assert!(out.contains(r#"<row r="3" spans="1:1">"#), "{out}");
        assert!(!out.contains("ht="), "{out}");
        assert!(!out.contains("customHeight"), "{out}");
    }

    #[test]
    fn hiding_a_row_leaves_its_height_alone_and_the_other_way_round() {
        let hide = BTreeMap::from([(3, RowWant {
            hidden: Some(true), height: None, pinned: false,
         style: None, })]);
        let out = row_tags(THREE_ROWS, &hide);
        assert!(out.contains(r#"ht="30""#), "{out}");
        assert!(out.contains(r#"customHeight="1""#), "{out}");
        assert!(out.contains(r#"hidden="1""#), "{out}");

        let taller = BTreeMap::from([(3, RowWant {
            hidden: None, height: Some(Some(45.0)), pinned: true,
         style: None, })]);
        let out = row_tags(
            &THREE_ROWS.replace(r#"ht="30""#, r#"ht="30" hidden="1""#),
            &taller,
        );
        assert!(out.contains(r#"ht="45""#), "{out}");
        assert!(out.contains(r#"hidden="1""#), "{out}");
    }

    #[test]
    fn a_row_keeps_whatever_else_its_tag_was_carrying() {
        // The span and the descent Excel notes are not this edit's business,
        // and a row that loses its span is one Excel has to work out again.
        let want = BTreeMap::from([(1, RowWant {
            hidden: None, height: Some(Some(22.5)), pinned: true,
         style: None, })]);
        let xml = THREE_ROWS.replace(
            r#"<row r="1" spans="1:1">"#,
            r#"<row r="1" spans="1:1" s="4" customFormat="1" x14ac:dyDescent="0.4">"#,
        );
        let out = row_tags(&xml, &want);
        assert!(out.contains(r#"s="4""#), "{out}");
        assert!(out.contains(r#"customFormat="1""#), "{out}");
        assert!(out.contains(r#"x14ac:dyDescent="0.4""#), "{out}");
        assert!(out.contains(r#"ht="22.5""#), "{out}");
    }

    #[test]
    fn a_row_the_sheet_never_described_can_still_be_given_a_height() {
        let want = BTreeMap::from([(7, RowWant {
            hidden: None, height: Some(Some(36.0)), pinned: true,
         style: None, })]);
        let out = row_tags(THREE_ROWS, &want);
        assert!(out.contains(r#"<row r="7" ht="36" customHeight="1">"#), "{out}");
    }

    #[test]
    fn a_width_is_written_where_the_sheet_had_no_columns_at_all() {
        // A sheet nobody has resized carries no `<cols>` element  Emeasured on
        // a freshly saved workbook  Eso one has to be put in, before
        // `<sheetData>` where the schema wants it.
        let want = BTreeMap::from([(1, ColumnWant { hidden: None, width: Some(10.625), custom: None, style: None, })]);
        assert_eq!(
            cols_of(BARE, &want),
            r#"<cols><col min="2" max="2" width="10.625" customWidth="1"/></cols>"#,
        );
    }

    #[test]
    fn a_width_always_brings_the_flag_that_says_it_was_chosen() {
        // Excel writes `customWidth="1"` beside every width it stores. Without
        // it the number is there and Excel ignores it.
        let want = BTreeMap::from([(0, ColumnWant { hidden: None, width: Some(30.625), custom: None, style: None, })]);
        let out = cols_of(BARE, &want);
        assert!(out.contains(r#"width="30.625""#), "{out}");
        assert!(out.contains(r#"customWidth="1""#), "{out}");
    }

    #[test]
    fn resizing_one_column_of_a_run_splits_the_run() {
        // `A:E` at one width, then C widened, comes back as three spans  E        // 1-2, 3-3, 4-5  Ewhich is what Excel itself writes.
        let xml = concat!(
            r#"<worksheet><cols><col min="1" max="5" width="12.625" customWidth="1"/></cols>"#,
            r#"<sheetData/></worksheet>"#,
        );
        let want = BTreeMap::from([(2, ColumnWant { hidden: None, width: Some(30.625), custom: None, style: None, })]);
        assert_eq!(
            cols_of(xml, &want),
            concat!(
                r#"<cols><col min="1" max="2" width="12.625" customWidth="1"/>"#,
                r#"<col min="3" max="3" width="30.625" customWidth="1"/>"#,
                r#"<col min="4" max="5" width="12.625" customWidth="1"/></cols>"#,
            ),
        );
    }

    #[test]
    fn resizing_a_column_drops_the_claim_that_it_was_a_best_fit() {
        // An autofitted column carries `bestFit="1"`. Excel takes that off
        // when a width is set by hand, which is the sort of thing that has to
        // be looked at rather than reasoned about  Emeasured.
        let xml = concat!(
            r#"<worksheet><cols>"#,
            r#"<col min="2" max="2" width="7.5" bestFit="1" customWidth="1"/>"#,
            r#"</cols><sheetData/></worksheet>"#,
        );
        let want = BTreeMap::from([(1, ColumnWant { hidden: None, width: Some(20.625), custom: None, style: None, })]);
        let out = cols_of(xml, &want);
        assert!(!out.contains("bestFit"), "{out}");
        assert!(out.contains(r#"width="20.625""#), "{out}");
    }

    #[test]
    fn hiding_a_column_leaves_its_width_alone_and_the_other_way_round() {
        let xml = concat!(
            r#"<worksheet><cols>"#,
            r#"<col min="2" max="2" width="18.625" customWidth="1"/>"#,
            r#"</cols><sheetData/></worksheet>"#,
        );
        let hide = BTreeMap::from([(1, ColumnWant { hidden: Some(true), width: None, custom: None, style: None, })]);
        let out = cols_of(xml, &hide);
        assert!(out.contains(r#"width="18.625""#), "{out}");
        assert!(out.contains(r#"hidden="1""#), "{out}");

        let widen = BTreeMap::from([(1, ColumnWant { hidden: None, width: Some(9.0), custom: None, style: None, })]);
        let out = cols_of(
            &xml.replace(r#"width="18.625""#, r#"width="18.625" hidden="1""#),
            &widen,
        );
        assert!(out.contains(r#"hidden="1""#), "{out}");
        assert!(out.contains(r#"width="9""#), "{out}");
    }

    #[test]
    fn a_column_keeps_whatever_else_its_span_was_carrying() {
        // A `<col>` can dress its column in a style and set an outline level.
        // Neither is this edit's business, and losing them would change how
        // every cell in the column is drawn.
        let xml = concat!(
            r#"<worksheet><cols>"#,
            r#"<col min="3" max="3" width="8" style="4" outlineLevel="2"/>"#,
            r#"</cols><sheetData/></worksheet>"#,
        );
        let want = BTreeMap::from([(2, ColumnWant { hidden: None, width: Some(15.5), custom: None, style: None, })]);
        let out = cols_of(xml, &want);
        assert!(out.contains(r#"style="4""#), "{out}");
        assert!(out.contains(r#"outlineLevel="2""#), "{out}");
        assert!(out.contains(r#"width="15.5""#), "{out}");
    }

    #[test]
    fn columns_asked_for_the_same_width_are_written_as_one_span() {
        let want = BTreeMap::from([
            (0, ColumnWant { hidden: None, width: Some(12.0), custom: None, style: None, }),
            (1, ColumnWant { hidden: None, width: Some(12.0), custom: None, style: None, }),
            (2, ColumnWant { hidden: None, width: Some(12.0), custom: None, style: None, }),
        ]);
        let out = cols_of(BARE, &want);
        assert_eq!(out, r#"<cols><col min="1" max="3" width="12" customWidth="1"/></cols>"#);
        // The same coalescing applies when an existing span covers them.
        let xml = r#"<worksheet><cols><col min="1" max="3" width="8"/></cols><sheetData/></worksheet>"#;
        let out = cols_of(xml, &want);
        assert_eq!(
            out,
            r#"<cols><col min="1" max="3" width="12" customWidth="1"/></cols>"#,
        );
    }

    #[test]
    fn a_cell_written_for_the_first_time_wears_the_style_it_was_given() {
        // A cell the sheet already had gets its style put on as it is copied
        // through. A cell that did not exist goes down the insert path, which
        // did not consult the style map at all  Eso it came out under General
        // however it had been styled, and a date filled into an empty cell
        // landed in the file as a bare serial number.
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;
        let fresh = CellEditValue::Number(46053.0);
        let over = CellEditValue::Number(7.0);
        let edits = HashMap::from([((1, 1), &fresh), ((1, 0), &over)]);
        let styles = BTreeMap::from([((1, 1), 4u32), ((1, 0), 5u32)]);
        let patched = patch_worksheet_xml(xml, &SheetEdits {
            cells: &edits,
            rows: nothing_rows(),
            cols: nothing_cols(),
            merges: None,
            styles: &styles,
            filter: None,
            panes: None,
            hyperlinks: None,
            hyperlink_ids: &[],
        }).expect("should patch");
        assert!(patched.contains(r#"<c r="B1" s="4""#), "{patched}");
        assert!(patched.contains(r#"<c r="A1" s="5""#), "{patched}");
    }

    #[test]
    fn a_row_written_for_the_first_time_carries_its_cells_styles() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;
        let fresh = CellEditValue::Number(46054.0);
        let edits = HashMap::from([((3, 0), &fresh)]);
        let styles = BTreeMap::from([((3, 0), 2u32)]);
        let patched = patch_worksheet_xml(xml, &SheetEdits {
            cells: &edits,
            rows: nothing_rows(),
            cols: nothing_cols(),
            merges: None,
            styles: &styles,
            filter: None,
            panes: None,
            hyperlinks: None,
            hyperlink_ids: &[],
        }).expect("should patch");
        assert!(patched.contains(r#"<c r="A3" s="2""#), "{patched}");
    }

    #[test]
    fn worksheet_patch_orders_insertions_and_replaces_formula_content() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><f>1+1</f><v>2</v></c><c r="C1" t="inlineStr"><is><t>old</t></is></c></row><row r="3"/></sheetData></worksheet>"#;
        let formula_value = CellEditValue::String("9".to_string());
        let middle_value = CellEditValue::String("middle".to_string());
        let inline_value = CellEditValue::String("new".to_string());
        let new_row_value = CellEditValue::String("row two".to_string());
        let edits = HashMap::from([
            ((1, 0), &formula_value),
            ((1, 1), &middle_value),
            ((1, 2), &inline_value),
            ((2, 1), &new_row_value),
        ]);

        let patched = patch_worksheet_xml(xml, &cells_only(&edits)).expect("should patch worksheet");
        assert!(!patched.contains("<f>"));
        assert!(!patched.contains("<is>"));
        let a1 = patched.find("r=\"A1\"").unwrap();
        let b1 = patched.find("r=\"B1\"").unwrap();
        let c1 = patched.find("r=\"C1\"").unwrap();
        assert!(a1 < b1 && b1 < c1);
        let row1 = patched.find("<row r=\"1\"").unwrap();
        let row2 = patched.find("<row r=\"2\"").unwrap();
        let row3 = patched.find("<row r=\"3\"").unwrap();
        assert!(row1 < row2 && row2 < row3);
        assert!(patched.contains("<v>9</v>"));
        assert!(patched.contains("<v>middle</v>"));
        assert!(patched.contains("<v>new</v>"));
        assert!(patched.contains("<v>row two</v>"));
    }

    #[test]
    fn worksheet_patch_populates_an_empty_sheet_data_element() {
        let xml = r#"<worksheet><sheetData/></worksheet>"#;
        let value = CellEditValue::String("first value".to_string());
        let edits = HashMap::from([((1, 0), &value)]);

        let patched = patch_worksheet_xml(xml, &cells_only(&edits)).expect("should patch empty worksheet");
        let expected = "<sheetData><row r=\"1\"><c r=\"A1\" t=\"str\">\
                        <v>first value</v></c></row></sheetData>";
        assert!(patched.contains(expected));
    }
}
