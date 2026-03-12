//! Round-trip docx editor with full structural editing support.
//!
//! Preserves the original ZIP archive and XML structure.
//! Supports text editing, paragraph/run insertion/deletion,
//! formatting changes, table operations, and image insertion.

use std::collections::HashMap;
use std::io::{Cursor, Read, Write};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::ir::Document;
use crate::parser::{parse_docx, ParseError};

// ---------------------------------------------------------------------------
// Public types for edit operations
// ---------------------------------------------------------------------------

/// Run-level formatting properties.
/// Only `Some` values are applied; `None` means "don't change".
#[derive(Debug, Clone, Default, serde::Deserialize)]
pub struct RunProps {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    /// Underline style (e.g. "single", "double", "wave")
    pub underline_style: Option<String>,
    pub strikethrough: Option<bool>,
    /// Font family name (e.g. "Arial", "游ゴシック")
    pub font_family: Option<String>,
    /// East-Asian font family
    pub font_family_east_asia: Option<String>,
    /// Font size in points (e.g. 11.0)
    pub font_size: Option<f32>,
    /// Color as hex without '#' (e.g. "FF0000")
    pub color: Option<String>,
    /// Highlight color name (e.g. "yellow")
    pub highlight: Option<String>,
    /// Character spacing in points (w:spacing w:val in twips/20).
    /// Positive = expanded, negative = condensed.
    /// Applied after kerning, added to each character's advance width.
    pub character_spacing: Option<f32>,
    /// Kerning threshold in points (w:kern w:val in half-points/2).
    /// Characters at or above this size get kerned via system kerning.
    /// 0 = disabled. Common: 1.0 (kern everything >= 1pt).
    pub kerning: Option<f32>,
    /// Language tag (w:lang w:val, BCP 47). e.g. "en-US", "ja-JP"
    pub lang: Option<String>,
    /// East-Asian language tag (w:lang w:eastAsia). e.g. "ja-JP"
    pub lang_east_asia: Option<String>,
    /// Bidi language tag (w:lang w:bidi). e.g. "ar-SA"
    pub lang_bidi: Option<String>,
    /// Run style reference (w:rStyle w:val). e.g. "Hyperlink"
    pub run_style: Option<String>,
    /// No proofing (w:noProof). Boolean toggle element.
    pub no_proof: Option<bool>,
    /// Superscript/subscript (w:vertAlign w:val). "superscript" or "subscript"
    pub vertical_align: Option<String>,
}

/// Paragraph-level formatting properties.
/// Only `Some` values are applied; `None` means "don't change".
#[derive(Debug, Clone, Default, serde::Deserialize)]
pub struct ParaProps {
    /// "left", "center", "right", "both" (justify), "distribute"
    pub alignment: Option<String>,
    /// Space before in points
    pub space_before: Option<f32>,
    /// Space after in points
    pub space_after: Option<f32>,
    /// Line spacing multiplier (e.g. 1.15). Generates lineRule="auto".
    pub line_spacing: Option<f32>,
    /// Left indent in points
    pub indent_left: Option<f32>,
    /// Right indent in points
    pub indent_right: Option<f32>,
    /// First line indent in points (negative = hanging)
    pub indent_first_line: Option<f32>,
    /// Style ID (e.g. "Heading1", "Normal")
    pub style_id: Option<String>,
    /// Keep with next paragraph (w:keepNext). Boolean toggle.
    pub keep_next: Option<bool>,
    /// Keep lines together (w:keepLines). Boolean toggle.
    pub keep_lines: Option<bool>,
    /// Widow/orphan control (w:widowControl).
    /// Word default: true. Japanese Normal style often sets false.
    pub widow_control: Option<bool>,
    /// Snap to document grid (w:snapToGrid).
    /// Word default: true. val="0" disables.
    pub snap_to_grid: Option<bool>,
    /// Word wrap at word boundaries (w:wordWrap).
    /// val="0" allows mid-word breaks (CJK text fitting).
    /// Often seen with autoSpaceDE="0", autoSpaceDN="0", adjustRightInd="0".
    pub word_wrap: Option<bool>,
    /// Adjust right indent for grid (w:adjustRightInd).
    /// val="0" disables. Part of CJK formatting control group.
    pub adjust_right_ind: Option<bool>,
    /// Auto-space between CJK and Latin (w:autoSpaceDE). val="0" disables.
    pub auto_space_de: Option<bool>,
    /// Auto-space between CJK and numbers (w:autoSpaceDN). val="0" disables.
    pub auto_space_dn: Option<bool>,
    /// Page break before this paragraph (w:pageBreakBefore). Boolean toggle.
    pub page_break_before: Option<bool>,
}

/// A single edit operation on a .docx document.
#[derive(Debug, Clone)]
pub enum DocxEdit {
    // --- Text operations ---
    /// Replace text content of a specific run
    SetRunText {
        paragraph_index: usize,
        run_index: usize,
        new_text: String,
    },
    /// Insert a new paragraph at body-level index.
    /// If `runs` is empty, creates a paragraph with one empty run.
    InsertParagraph {
        index: usize,
        runs: Vec<(String, Option<RunProps>)>,
        style: Option<ParaProps>,
    },
    /// Delete a body-level paragraph by index.
    DeleteParagraph {
        index: usize,
    },
    /// Insert a new run into an existing paragraph.
    InsertRun {
        paragraph_index: usize,
        run_index: usize,
        text: String,
        style: Option<RunProps>,
    },
    /// Delete a run from a paragraph.
    DeleteRun {
        paragraph_index: usize,
        run_index: usize,
    },

    // --- Formatting ---
    /// Set formatting on an existing run (merge with existing).
    SetRunFormat {
        paragraph_index: usize,
        run_index: usize,
        style: RunProps,
    },
    /// Set paragraph-level formatting.
    SetParagraphFormat {
        paragraph_index: usize,
        style: ParaProps,
    },

    // --- Tables ---
    /// Insert a table at body-level index.
    /// `content[row][col]` = cell text. If None, cells are empty.
    InsertTable {
        index: usize,
        rows: usize,
        cols: usize,
        content: Option<Vec<Vec<String>>>,
        col_widths_pt: Option<Vec<f32>>,
    },
    /// Insert a row into an existing table.
    InsertTableRow {
        /// Body-level index of the table (counting only tables, 0-based)
        table_index: usize,
        row_index: usize,
        cells: Vec<String>,
    },
    /// Delete a row from a table.
    DeleteTableRow {
        table_index: usize,
        row_index: usize,
    },
    /// Set text in a table cell.
    SetCellText {
        table_index: usize,
        row: usize,
        col: usize,
        text: String,
    },

    // --- Images ---
    /// Insert an inline image as a new paragraph at body-level index.
    InsertImage {
        index: usize,
        data: Vec<u8>,
        width_pt: f32,
        height_pt: f32,
        content_type: String,
    },
}

/// Legacy text edit (kept for backward compatibility).
#[derive(Debug, Clone)]
pub struct TextEdit {
    pub paragraph_index: usize,
    pub run_index: usize,
    pub new_text: String,
}

// ---------------------------------------------------------------------------
// Body segment representation
// ---------------------------------------------------------------------------

/// A top-level element within `<w:body>`.
#[derive(Debug, Clone)]
enum BodySegment {
    Paragraph(String),
    Table(String),
    SectPr(String),
    Other(String),
}

impl BodySegment {
    fn xml(&self) -> &str {
        match self {
            Self::Paragraph(s) | Self::Table(s) | Self::SectPr(s) | Self::Other(s) => s,
        }
    }
}

// ---------------------------------------------------------------------------
// DocxEditor
// ---------------------------------------------------------------------------

/// Round-trip docx editor that preserves the original archive.
pub struct DocxEditor {
    original_data: Vec<u8>,
    document: Document,
    edits: Vec<DocxEdit>,
    /// Images to add to the ZIP (rel_id, data, content_type, filename)
    pending_images: Vec<PendingImage>,
}

struct PendingImage {
    rel_id: String,
    data: Vec<u8>,
    content_type: String,
    filename: String,
}

impl DocxEditor {
    /// Create a new editor from raw .docx bytes.
    pub fn new(data: &[u8]) -> Result<Self, ParseError> {
        let document = parse_docx(data)?;
        Ok(Self {
            original_data: data.to_vec(),
            document,
            edits: Vec::new(),
            pending_images: Vec::new(),
        })
    }

    /// Get a reference to the parsed document IR (read-only).
    pub fn document(&self) -> &Document {
        &self.document
    }

    // --- Legacy API (backward compatible) ---

    /// Set the text of a specific run.
    pub fn set_run_text(&mut self, paragraph_index: usize, run_index: usize, new_text: String) {
        self.edits.push(DocxEdit::SetRunText {
            paragraph_index,
            run_index,
            new_text,
        });
    }

    /// Apply multiple legacy text edits at once.
    pub fn apply_edits(&mut self, edits: &[TextEdit]) {
        for edit in edits {
            self.set_run_text(edit.paragraph_index, edit.run_index, edit.new_text.clone());
        }
    }

    /// Check if there are pending edits.
    pub fn has_edits(&self) -> bool {
        !self.edits.is_empty()
    }

    // --- New API ---

    /// Add an edit operation.
    pub fn add_edit(&mut self, edit: DocxEdit) {
        self.edits.push(edit);
    }

    /// Add multiple edit operations.
    pub fn add_edits(&mut self, edits: Vec<DocxEdit>) {
        self.edits.extend(edits);
    }

    /// Insert a paragraph with text at the given body-level index.
    pub fn insert_paragraph(
        &mut self,
        index: usize,
        text: &str,
        run_style: Option<RunProps>,
        para_style: Option<ParaProps>,
    ) {
        self.edits.push(DocxEdit::InsertParagraph {
            index,
            runs: vec![(text.to_string(), run_style)],
            style: para_style,
        });
    }

    /// Delete a paragraph at the given body-level index.
    pub fn delete_paragraph(&mut self, index: usize) {
        self.edits.push(DocxEdit::DeleteParagraph { index });
    }

    /// Insert a run into an existing paragraph.
    pub fn insert_run(
        &mut self,
        paragraph_index: usize,
        run_index: usize,
        text: &str,
        style: Option<RunProps>,
    ) {
        self.edits.push(DocxEdit::InsertRun {
            paragraph_index,
            run_index,
            text: text.to_string(),
            style,
        });
    }

    /// Delete a run from a paragraph.
    pub fn delete_run(&mut self, paragraph_index: usize, run_index: usize) {
        self.edits.push(DocxEdit::DeleteRun {
            paragraph_index,
            run_index,
        });
    }

    /// Set formatting on an existing run.
    pub fn set_run_format(
        &mut self,
        paragraph_index: usize,
        run_index: usize,
        style: RunProps,
    ) {
        self.edits.push(DocxEdit::SetRunFormat {
            paragraph_index,
            run_index,
            style,
        });
    }

    /// Set paragraph-level formatting.
    pub fn set_paragraph_format(&mut self, paragraph_index: usize, style: ParaProps) {
        self.edits.push(DocxEdit::SetParagraphFormat {
            paragraph_index,
            style,
        });
    }

    /// Insert a table at body-level index.
    pub fn insert_table(
        &mut self,
        index: usize,
        rows: usize,
        cols: usize,
        content: Option<Vec<Vec<String>>>,
        col_widths_pt: Option<Vec<f32>>,
    ) {
        self.edits.push(DocxEdit::InsertTable {
            index,
            rows,
            cols,
            content,
            col_widths_pt,
        });
    }

    /// Insert a row into a table.
    pub fn insert_table_row(
        &mut self,
        table_index: usize,
        row_index: usize,
        cells: Vec<String>,
    ) {
        self.edits.push(DocxEdit::InsertTableRow {
            table_index,
            row_index,
            cells,
        });
    }

    /// Delete a row from a table.
    pub fn delete_table_row(&mut self, table_index: usize, row_index: usize) {
        self.edits.push(DocxEdit::DeleteTableRow {
            table_index,
            row_index,
        });
    }

    /// Set text in a table cell.
    pub fn set_cell_text(
        &mut self,
        table_index: usize,
        row: usize,
        col: usize,
        text: &str,
    ) {
        self.edits.push(DocxEdit::SetCellText {
            table_index,
            row,
            col,
            text: text.to_string(),
        });
    }

    /// Insert an inline image as a new paragraph.
    pub fn insert_image(
        &mut self,
        index: usize,
        data: Vec<u8>,
        width_pt: f32,
        height_pt: f32,
        content_type: &str,
    ) {
        // Determine next relationship ID
        let img_idx = self.pending_images.len() + 1;
        let rel_id = format!("rIdImg{}", img_idx);
        let ext = match content_type {
            "image/png" => "png",
            "image/jpeg" | "image/jpg" => "jpeg",
            "image/gif" => "gif",
            "image/bmp" => "bmp",
            "image/tiff" => "tiff",
            _ => "png",
        };
        let filename = format!("image_inserted_{}.{}", img_idx, ext);

        self.pending_images.push(PendingImage {
            rel_id: rel_id.clone(),
            data,
            content_type: content_type.to_string(),
            filename: filename.clone(),
        });

        self.edits.push(DocxEdit::InsertImage {
            index,
            data: Vec::new(), // data is in pending_images
            width_pt,
            height_pt,
            content_type: content_type.to_string(),
        });
    }

    // --- Save ---

    /// Save the edited document as new .docx bytes.
    pub fn save(&self) -> Result<Vec<u8>, ParseError> {
        if self.edits.is_empty() && self.pending_images.is_empty() {
            return self.copy_archive();
        }

        let cursor = Cursor::new(&self.original_data);
        let mut archive = ZipArchive::new(cursor)?;

        let mut output = Vec::new();
        {
            let mut writer = ZipWriter::new(Cursor::new(&mut output));
            let opts = SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);

            for i in 0..archive.len() {
                let mut entry = archive.by_index(i)?;
                let name = entry.name().to_string();

                if name == "word/document.xml" {
                    let mut xml = String::new();
                    entry.read_to_string(&mut xml)?;
                    let patched = self.apply_all_edits(&xml)?;
                    writer.start_file(&name, opts)?;
                    writer.write_all(patched.as_bytes())?;
                } else if name == "word/_rels/document.xml.rels" && !self.pending_images.is_empty() {
                    let mut xml = String::new();
                    entry.read_to_string(&mut xml)?;
                    let patched = self.patch_rels(&xml);
                    writer.start_file(&name, opts)?;
                    writer.write_all(patched.as_bytes())?;
                } else if name == "[Content_Types].xml" && !self.pending_images.is_empty() {
                    let mut xml = String::new();
                    entry.read_to_string(&mut xml)?;
                    let patched = self.patch_content_types(&xml);
                    writer.start_file(&name, opts)?;
                    writer.write_all(patched.as_bytes())?;
                } else {
                    let entry_opts = SimpleFileOptions::default()
                        .compression_method(entry.compression());
                    writer.start_file(&name, entry_opts)?;
                    let mut buf = Vec::new();
                    entry.read_to_end(&mut buf)?;
                    writer.write_all(&buf)?;
                }
            }

            // Add image files to ZIP
            for img in &self.pending_images {
                let path = format!("word/media/{}", img.filename);
                writer.start_file(&path, opts)?;
                writer.write_all(&img.data)?;
            }

            writer.finish()?;
        }

        Ok(output)
    }

    /// Copy archive unchanged (no edits).
    fn copy_archive(&self) -> Result<Vec<u8>, ParseError> {
        let cursor = Cursor::new(&self.original_data);
        let mut archive = ZipArchive::new(cursor)?;
        let mut output = Vec::new();
        {
            let mut writer = ZipWriter::new(Cursor::new(&mut output));
            for i in 0..archive.len() {
                let mut entry = archive.by_index(i)?;
                let name = entry.name().to_string();
                let opts = SimpleFileOptions::default()
                    .compression_method(entry.compression());
                writer.start_file(&name, opts)?;
                let mut buf = Vec::new();
                entry.read_to_end(&mut buf)?;
                writer.write_all(&buf)?;
            }
            writer.finish()?;
        }
        Ok(output)
    }

    // -----------------------------------------------------------------------
    // Core: apply all edits to document.xml
    // -----------------------------------------------------------------------

    fn apply_all_edits(&self, xml: &str) -> Result<String, ParseError> {
        // 1. Split XML into: preamble, body segments, postamble
        let (preamble, mut segments, postamble) = split_body(xml)?;

        // 2. Classify edits
        let mut text_edits: HashMap<(usize, usize), String> = HashMap::new();
        let mut run_format_edits: HashMap<(usize, usize), RunProps> = HashMap::new();
        let mut para_format_edits: HashMap<usize, ParaProps> = HashMap::new();
        let mut para_deletions: Vec<usize> = Vec::new();
        let mut para_insertions: Vec<(usize, String)> = Vec::new(); // (index, xml)
        let mut run_insertions: HashMap<(usize, usize), (String, Option<RunProps>)> = HashMap::new();
        let mut run_deletions: Vec<(usize, usize)> = Vec::new();
        let mut table_insertions: Vec<(usize, String)> = Vec::new();
        let mut table_row_insertions: Vec<(usize, usize, Vec<String>)> = Vec::new();
        let mut table_row_deletions: Vec<(usize, usize)> = Vec::new();
        let mut cell_text_edits: HashMap<(usize, usize, usize), String> = HashMap::new();
        let mut image_insertions: Vec<(usize, usize)> = Vec::new(); // (body_index, pending_image_index)

        let mut image_counter = 0usize;

        for edit in &self.edits {
            match edit {
                DocxEdit::SetRunText { paragraph_index, run_index, new_text } => {
                    text_edits.insert((*paragraph_index, *run_index), new_text.clone());
                }
                DocxEdit::InsertParagraph { index, runs, style } => {
                    let xml_frag = generate_paragraph_xml(runs, style.as_ref());
                    para_insertions.push((*index, xml_frag));
                }
                DocxEdit::DeleteParagraph { index } => {
                    para_deletions.push(*index);
                }
                DocxEdit::InsertRun { paragraph_index, run_index, text, style } => {
                    run_insertions.insert(
                        (*paragraph_index, *run_index),
                        (text.clone(), style.clone()),
                    );
                }
                DocxEdit::DeleteRun { paragraph_index, run_index } => {
                    run_deletions.push((*paragraph_index, *run_index));
                }
                DocxEdit::SetRunFormat { paragraph_index, run_index, style } => {
                    run_format_edits.insert((*paragraph_index, *run_index), style.clone());
                }
                DocxEdit::SetParagraphFormat { paragraph_index, style } => {
                    para_format_edits.insert(*paragraph_index, style.clone());
                }
                DocxEdit::InsertTable { index, rows, cols, content, col_widths_pt } => {
                    let xml_frag = generate_table_xml(*rows, *cols, content.as_ref(), col_widths_pt.as_ref());
                    table_insertions.push((*index, xml_frag));
                }
                DocxEdit::InsertTableRow { table_index, row_index, cells } => {
                    table_row_insertions.push((*table_index, *row_index, cells.clone()));
                }
                DocxEdit::DeleteTableRow { table_index, row_index } => {
                    table_row_deletions.push((*table_index, *row_index));
                }
                DocxEdit::SetCellText { table_index, row, col, text } => {
                    cell_text_edits.insert((*table_index, *row, *col), text.clone());
                }
                DocxEdit::InsertImage { index, width_pt, height_pt, content_type, .. } => {
                    if image_counter < self.pending_images.len() {
                        let rel_id = &self.pending_images[image_counter].rel_id;
                        let xml_frag = generate_image_paragraph_xml(
                            rel_id,
                            *width_pt,
                            *height_pt,
                            image_counter + 1,
                        );
                        para_insertions.push((*index, xml_frag));
                        image_counter += 1;
                    }
                }
            }
        }

        // 3. Apply structural edits to segments

        // Sort insertions by index (descending) so we can insert without shifting
        para_insertions.sort_by(|a, b| b.0.cmp(&a.0));
        table_insertions.sort_by(|a, b| b.0.cmp(&a.0));

        // Count paragraphs and tables for index mapping
        let mut para_indices: Vec<usize> = Vec::new(); // segment index for each paragraph
        let mut table_indices: Vec<usize> = Vec::new(); // segment index for each table
        for (i, seg) in segments.iter().enumerate() {
            match seg {
                BodySegment::Paragraph(_) => para_indices.push(i),
                BodySegment::Table(_) => table_indices.push(i),
                _ => {}
            }
        }

        // Delete paragraphs (from highest index first)
        let mut sorted_deletions = para_deletions.clone();
        sorted_deletions.sort_unstable();
        sorted_deletions.dedup();
        for &para_idx in sorted_deletions.iter().rev() {
            if para_idx < para_indices.len() {
                let seg_idx = para_indices[para_idx];
                segments.remove(seg_idx);
                // Rebuild indices
                para_indices.clear();
                table_indices.clear();
                for (i, seg) in segments.iter().enumerate() {
                    match seg {
                        BodySegment::Paragraph(_) => para_indices.push(i),
                        BodySegment::Table(_) => table_indices.push(i),
                        _ => {}
                    }
                }
            }
        }

        // Insert paragraphs and tables
        // Merge para_insertions and table_insertions, sorted by body-level index
        let mut all_insertions: Vec<(usize, String, bool)> = Vec::new(); // (body_idx, xml, is_table)
        for (idx, xml) in &para_insertions {
            all_insertions.push((*idx, xml.clone(), false));
        }
        for (idx, xml) in &table_insertions {
            all_insertions.push((*idx, xml.clone(), true));
        }
        all_insertions.sort_by(|a, b| b.0.cmp(&a.0)); // descending

        for (body_idx, xml_frag, is_table) in &all_insertions {
            // Find the segment index to insert before
            // body_idx refers to body-level element count (paragraphs + tables)
            let mut body_count = 0;
            let mut insert_seg_idx = segments.len();
            // Insert before sectPr if at end
            for (i, seg) in segments.iter().enumerate() {
                match seg {
                    BodySegment::SectPr(_) => {
                        if body_count <= *body_idx {
                            insert_seg_idx = i;
                            break;
                        }
                    }
                    BodySegment::Paragraph(_) | BodySegment::Table(_) => {
                        if body_count == *body_idx {
                            insert_seg_idx = i;
                            break;
                        }
                        body_count += 1;
                    }
                    _ => {}
                }
            }

            let new_seg = if *is_table {
                BodySegment::Table(xml_frag.clone())
            } else {
                BodySegment::Paragraph(xml_frag.clone())
            };
            segments.insert(insert_seg_idx, new_seg);
        }

        // Rebuild indices after structural changes
        para_indices.clear();
        table_indices.clear();
        for (i, seg) in segments.iter().enumerate() {
            match seg {
                BodySegment::Paragraph(_) => para_indices.push(i),
                BodySegment::Table(_) => table_indices.push(i),
                _ => {}
            }
        }

        // 4. Apply in-place edits to paragraph segments
        for (para_idx, seg_idx) in para_indices.iter().enumerate() {
            let seg_xml = match &segments[*seg_idx] {
                BodySegment::Paragraph(xml) => xml.clone(),
                _ => continue,
            };

            let mut modified = false;
            let mut current_xml = seg_xml;

            // Text edits for this paragraph
            let para_text_edits: HashMap<usize, &String> = text_edits
                .iter()
                .filter(|((pi, _), _)| *pi == para_idx)
                .map(|((_, ri), t)| (*ri, t))
                .collect();

            // Run format edits for this paragraph
            let para_run_formats: HashMap<usize, &RunProps> = run_format_edits
                .iter()
                .filter(|((pi, _), _)| *pi == para_idx)
                .map(|((_, ri), s)| (*ri, s))
                .collect();

            // Run insertions for this paragraph
            let para_run_insertions: HashMap<usize, &(String, Option<RunProps>)> = run_insertions
                .iter()
                .filter(|((pi, _), _)| *pi == para_idx)
                .map(|((_, ri), v)| (*ri, v))
                .collect();

            // Run deletions for this paragraph
            let para_run_deletions: Vec<usize> = run_deletions
                .iter()
                .filter(|(pi, _)| *pi == para_idx)
                .map(|(_, ri)| *ri)
                .collect();

            // Paragraph format edit
            let para_fmt = para_format_edits.get(&para_idx);

            if !para_text_edits.is_empty()
                || !para_run_formats.is_empty()
                || !para_run_insertions.is_empty()
                || !para_run_deletions.is_empty()
                || para_fmt.is_some()
            {
                current_xml = patch_paragraph_xml(
                    &current_xml,
                    &para_text_edits,
                    &para_run_formats,
                    &para_run_insertions,
                    &para_run_deletions,
                    para_fmt,
                )?;
                modified = true;
            }

            if modified {
                segments[*seg_idx] = BodySegment::Paragraph(current_xml);
            }
        }

        // 5. Apply in-place edits to table segments
        for (tbl_idx, seg_idx) in table_indices.iter().enumerate() {
            let seg_xml = match &segments[*seg_idx] {
                BodySegment::Table(xml) => xml.clone(),
                _ => continue,
            };

            // Cell text edits
            let tbl_cell_edits: HashMap<(usize, usize), &String> = cell_text_edits
                .iter()
                .filter(|((ti, _, _), _)| *ti == tbl_idx)
                .map(|((_, r, c), t)| ((*r, *c), t))
                .collect();

            // Row insertions
            let tbl_row_ins: Vec<(usize, &Vec<String>)> = table_row_insertions
                .iter()
                .filter(|(ti, _, _)| *ti == tbl_idx)
                .map(|(_, ri, cells)| (*ri, cells))
                .collect();

            // Row deletions
            let tbl_row_dels: Vec<usize> = table_row_deletions
                .iter()
                .filter(|(ti, _)| *ti == tbl_idx)
                .map(|(_, ri)| *ri)
                .collect();

            if !tbl_cell_edits.is_empty() || !tbl_row_ins.is_empty() || !tbl_row_dels.is_empty() {
                let patched = patch_table_xml(
                    &seg_xml,
                    &tbl_cell_edits,
                    &tbl_row_ins,
                    &tbl_row_dels,
                )?;
                segments[*seg_idx] = BodySegment::Table(patched);
            }
        }

        // 6. Reassemble
        let mut result = preamble;
        for seg in &segments {
            result.push_str(seg.xml());
        }
        result.push_str(&postamble);

        Ok(result)
    }

    // -----------------------------------------------------------------------
    // Relationship and content type patching for images
    // -----------------------------------------------------------------------

    fn patch_rels(&self, xml: &str) -> String {
        let insert_before = "</Relationships>";
        let mut additions = String::new();
        for img in &self.pending_images {
            additions.push_str(&format!(
                r#"<Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{}"/>"#,
                img.rel_id, img.filename
            ));
        }
        xml.replace(insert_before, &format!("{}{}", additions, insert_before))
    }

    fn patch_content_types(&self, xml: &str) -> String {
        // Add Default entries for image extensions if not already present
        let mut result = xml.to_string();
        let extensions: Vec<&str> = self
            .pending_images
            .iter()
            .map(|img| {
                let ext = img.filename.rsplit('.').next().unwrap_or("png");
                ext
            })
            .collect();

        for ext in extensions {
            let check = format!("Extension=\"{}\"", ext);
            if !result.contains(&check) {
                let ct = match ext {
                    "png" => "image/png",
                    "jpeg" | "jpg" => "image/jpeg",
                    "gif" => "image/gif",
                    "bmp" => "image/bmp",
                    "tiff" => "image/tiff",
                    _ => "image/png",
                };
                let entry = format!(
                    r#"<Default Extension="{}" ContentType="{}"/>"#,
                    ext, ct
                );
                result = result.replace(
                    "</Types>",
                    &format!("{}</Types>", entry),
                );
            }
        }
        result
    }
}

// ===========================================================================
// XML generation helpers
// ===========================================================================

/// Generate run properties XML (`<w:rPr>` content).
///
/// Element ordering follows ECMA-376 Part 1 §17.3.2.28 (CT_RPr):
///   rFonts → b → bCs → i → iCs → strike → color → sz → szCs → highlight → u
fn generate_rpr_xml(props: &RunProps) -> String {
    let mut parts = Vec::new();

    // 0. rStyle (run style reference) — always first per Word output
    if let Some(ref rs) = props.run_style {
        parts.push(format!(r#"<w:rStyle w:val="{}"/>"#, rs));
    }

    // 1. rFonts (font family)
    if let Some(ref ff) = props.font_family {
        let ea = props.font_family_east_asia.as_deref().unwrap_or(ff);
        parts.push(format!(
            r#"<w:rFonts w:ascii="{}" w:hAnsi="{}" w:eastAsia="{}"/>"#,
            ff, ff, ea
        ));
    } else if let Some(ref ea) = props.font_family_east_asia {
        parts.push(format!(r#"<w:rFonts w:eastAsia="{}"/>"#, ea));
    }

    // 2. b, bCs (bold)
    if props.bold == Some(true) {
        parts.push("<w:b/>".to_string());
        parts.push("<w:bCs/>".to_string());
    }

    // 3. i, iCs (italic)
    if props.italic == Some(true) {
        parts.push("<w:i/>".to_string());
        parts.push("<w:iCs/>".to_string());
    }

    // 4. strike (strikethrough)
    if props.strikethrough == Some(true) {
        parts.push("<w:strike/>".to_string());
    }

    // 4.5. noProof — boolean toggle, bare element = true
    if props.no_proof == Some(true) {
        parts.push("<w:noProof/>".to_string());
    }

    // 5. color
    if let Some(ref color) = props.color {
        parts.push(format!(r#"<w:color w:val="{}"/>"#, color));
    }

    // 5.5. spacing (character spacing in twips)
    // Value in twips (1/20pt), added to advance width after kerning.
    if let Some(cs) = props.character_spacing {
        let twips = (cs * 20.0).round() as i32;
        parts.push(format!(r#"<w:spacing w:val="{}"/>"#, twips));
    }

    // 5.6. kern (kerning threshold in half-points)
    // Characters >= threshold get kerned via system kerning.
    if let Some(kern) = props.kerning {
        let half_pt = (kern * 2.0).round() as u32;
        parts.push(format!(r#"<w:kern w:val="{}"/>"#, half_pt));
    }

    // 6. sz, szCs (font size — half-points)
    if let Some(size) = props.font_size {
        let half_pt = (size * 2.0).round() as u32;
        parts.push(format!(r#"<w:sz w:val="{}"/>"#, half_pt));
        parts.push(format!(r#"<w:szCs w:val="{}"/>"#, half_pt));
    }

    // 7. highlight
    if let Some(ref hl) = props.highlight {
        parts.push(format!(r#"<w:highlight w:val="{}"/>"#, hl));
    }

    // 8. highlight
    if let Some(ref hl) = props.highlight {
        parts.push(format!(r#"<w:highlight w:val="{}"/>"#, hl));
    }

    // 9. u (underline) — comes after highlight per Word output
    if props.underline == Some(true) {
        let style = props.underline_style.as_deref().unwrap_or("single");
        parts.push(format!(r#"<w:u w:val="{}"/>"#, style));
    }

    // 10. vertAlign (superscript/subscript)
    if let Some(ref va) = props.vertical_align {
        parts.push(format!(r#"<w:vertAlign w:val="{}"/>"#, va));
    }

    // 11. lang (language tag) — always near the end per Word output
    if props.lang.is_some() || props.lang_east_asia.is_some() || props.lang_bidi.is_some() {
        let mut lang_attrs = Vec::new();
        if let Some(ref v) = props.lang {
            lang_attrs.push(format!(r#"w:val="{}""#, v));
        }
        if let Some(ref ea) = props.lang_east_asia {
            lang_attrs.push(format!(r#"w:eastAsia="{}""#, ea));
        }
        if let Some(ref bi) = props.lang_bidi {
            lang_attrs.push(format!(r#"w:bidi="{}""#, bi));
        }
        parts.push(format!("<w:lang {}/>", lang_attrs.join(" ")));
    }

    if parts.is_empty() {
        String::new()
    } else {
        format!("<w:rPr>{}</w:rPr>", parts.join(""))
    }
}


/// Generate a single `<w:r>` element.
fn generate_run_xml(text: &str, props: Option<&RunProps>) -> String {
    let rpr = props.map(generate_rpr_xml).unwrap_or_default();
    // Use xml:space="preserve" for text with leading/trailing whitespace
    let space_attr = if text.starts_with(' ') || text.ends_with(' ') || text.contains('\t') {
        r#" xml:space="preserve""#
    } else {
        ""
    };
    format!(
        "<w:r>{}<w:t{}>{}</w:t></w:r>",
        rpr,
        space_attr,
        escape_xml(text)
    )
}

/// Generate paragraph properties XML (`<w:pPr>` content).
///
/// Element ordering follows ECMA-376 Part 1 §17.3.1.26 (CT_PPr):
///   pStyle → keepNext → keepLines → spacing → ind → jc
fn generate_ppr_xml(props: &ParaProps) -> String {
    let mut parts = Vec::new();

    // 1. pStyle (paragraph style reference) — must be first
    if let Some(ref style_id) = props.style_id {
        parts.push(format!(r#"<w:pStyle w:val="{}"/>"#, style_id));
    }

    // 1.5. keepNext, keepLines, pageBreakBefore — Word order: after pStyle
    if props.keep_next == Some(true) {
        parts.push("<w:keepNext/>".to_string());
    }
    if props.keep_lines == Some(true) {
        parts.push("<w:keepLines/>".to_string());
    }
    if props.page_break_before == Some(true) {
        parts.push("<w:pageBreakBefore/>".to_string());
    }

    // 1.6. widowControl — bare = true, val="0" = disable
    if let Some(wc) = props.widow_control {
        if wc {
            parts.push("<w:widowControl/>".to_string());
        } else {
            parts.push(r#"<w:widowControl w:val="0"/>"#.to_string());
        }
    }

    // 1.7. wordWrap, autoSpaceDE, autoSpaceDN, adjustRightInd — CJK formatting group
    if let Some(ww) = props.word_wrap {
        if ww {
            parts.push("<w:wordWrap/>".to_string());
        } else {
            parts.push(r#"<w:wordWrap w:val="0"/>"#.to_string());
        }
    }
    if let Some(de) = props.auto_space_de {
        if !de {
            parts.push(r#"<w:autoSpaceDE w:val="0"/>"#.to_string());
        }
    }
    if let Some(dn) = props.auto_space_dn {
        if !dn {
            parts.push(r#"<w:autoSpaceDN w:val="0"/>"#.to_string());
        }
    }
    if let Some(ari) = props.adjust_right_ind {
        if ari {
            parts.push("<w:adjustRightInd/>".to_string());
        } else {
            parts.push(r#"<w:adjustRightInd w:val="0"/>"#.to_string());
        }
    }

    // 1.8. snapToGrid
    if let Some(sg) = props.snap_to_grid {
        if sg {
            parts.push("<w:snapToGrid/>".to_string());
        } else {
            parts.push(r#"<w:snapToGrid w:val="0"/>"#.to_string());
        }
    }

    // 2. spacing
    let mut spacing_attrs = Vec::new();
    if let Some(before) = props.space_before {
        spacing_attrs.push(format!(r#"w:before="{}""#, (before * 20.0).round() as i32));
    }
    if let Some(after) = props.space_after {
        spacing_attrs.push(format!(r#"w:after="{}""#, (after * 20.0).round() as i32));
    }
    if let Some(ls) = props.line_spacing {
        let line_val = (ls * 240.0).round() as i32;
        spacing_attrs.push(format!(r#"w:line="{}" w:lineRule="auto""#, line_val));
    }
    if !spacing_attrs.is_empty() {
        parts.push(format!("<w:spacing {}/>", spacing_attrs.join(" ")));
    }

    // 3. ind (indentation)
    let mut ind_attrs = Vec::new();
    if let Some(left) = props.indent_left {
        ind_attrs.push(format!(r#"w:left="{}""#, (left * 20.0).round() as i32));
    }
    if let Some(right) = props.indent_right {
        ind_attrs.push(format!(r#"w:right="{}""#, (right * 20.0).round() as i32));
    }
    if let Some(first) = props.indent_first_line {
        if first >= 0.0 {
            ind_attrs.push(format!(r#"w:firstLine="{}""#, (first * 20.0).round() as i32));
        } else {
            ind_attrs.push(format!(r#"w:hanging="{}""#, (-first * 20.0).round() as i32));
        }
    }
    if !ind_attrs.is_empty() {
        parts.push(format!("<w:ind {}/>", ind_attrs.join(" ")));
    }

    // 4. jc (alignment) — comes after ind per ECMA-376
    if let Some(ref align) = props.alignment {
        parts.push(format!(r#"<w:jc w:val="{}"/>"#, align));
    }

    if parts.is_empty() {
        String::new()
    } else {
        format!("<w:pPr>{}</w:pPr>", parts.join(""))
    }
}

/// Generate a complete `<w:p>` element.
fn generate_paragraph_xml(
    runs: &[(String, Option<RunProps>)],
    style: Option<&ParaProps>,
) -> String {
    let ppr = style.map(generate_ppr_xml).unwrap_or_default();
    let runs_xml: String = if runs.is_empty() {
        generate_run_xml("", None)
    } else {
        runs.iter()
            .map(|(text, props)| generate_run_xml(text, props.as_ref()))
            .collect()
    };
    format!("<w:p>{}{}</w:p>", ppr, runs_xml)
}

/// Generate a complete `<w:tbl>` element.
///
/// Matches Word's actual output structure:
///   tblPr: tblStyle → tblW → tblBorders → tblLook
///   tblBorders: top/left/bottom/right/insideH/insideV with color="000000"
///   tblW: w:type before w:w (Word attribute order)
fn generate_table_xml(
    rows: usize,
    cols: usize,
    content: Option<&Vec<Vec<String>>>,
    col_widths_pt: Option<&Vec<f32>>,
) -> String {
    // Word default: Letter page body width = 8640 twips (6in at 1440 twips/in)
    // with 1in margins on 8.5in page. A4 is similar.
    let total_width_twips = 8640;
    let default_col_width = total_width_twips / cols as i32;

    let mut xml = String::new();
    xml.push_str("<w:tbl>");

    // Table properties — element order matches Word output exactly
    xml.push_str(concat!(
        "<w:tblPr>",
        r#"<w:tblW w:type="auto" w:w="0"/>"#,
        "<w:tblBorders>",
        r#"<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>"#,
        r#"<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>"#,
        r#"<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>"#,
        r#"<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>"#,
        r#"<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>"#,
        r#"<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>"#,
        "</w:tblBorders>",
        r#"<w:tblLook w:firstColumn="1" w:firstRow="1" w:lastColumn="0" w:lastRow="0" w:noHBand="0" w:noVBand="1" w:val="04A0"/>"#,
        "</w:tblPr>"
    ));

    // Grid columns
    xml.push_str("<w:tblGrid>");
    for c in 0..cols {
        let w = if let Some(widths) = col_widths_pt {
            if c < widths.len() {
                (widths[c] * 20.0).round() as i32
            } else {
                default_col_width
            }
        } else {
            default_col_width
        };
        xml.push_str(&format!(r#"<w:gridCol w:w="{}"/>"#, w));
    }
    xml.push_str("</w:tblGrid>");

    // Rows
    for r in 0..rows {
        xml.push_str("<w:tr>");
        for c in 0..cols {
            let w = if let Some(widths) = col_widths_pt {
                if c < widths.len() {
                    (widths[c] * 20.0).round() as i32
                } else {
                    default_col_width
                }
            } else {
                default_col_width
            };
            let cell_text = content
                .and_then(|rows| rows.get(r))
                .and_then(|cols| cols.get(c))
                .map(|s| s.as_str())
                .unwrap_or("");
            xml.push_str(&format!(
                r#"<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{}"/></w:tcPr><w:p><w:r><w:t>{}</w:t></w:r></w:p></w:tc>"#,
                w,
                escape_xml(cell_text)
            ));
        }
        xml.push_str("</w:tr>");
    }

    xml.push_str("</w:tbl>");
    xml
}

/// Generate a table row XML.
///
/// Cell properties use Word's attribute order: w:type before w:w.
fn generate_table_row_xml(cells: &[String], col_width_twips: i32) -> String {
    let mut xml = String::from("<w:tr>");
    for cell_text in cells {
        xml.push_str(&format!(
            r#"<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="{}"/></w:tcPr><w:p><w:r><w:t>{}</w:t></w:r></w:p></w:tc>"#,
            col_width_twips,
            escape_xml(cell_text)
        ));
    }
    xml.push_str("</w:tr>");
    xml
}

/// Generate a paragraph containing an inline image (DrawingML).
///
/// Structure matches Word's actual output from with_image.docx:
///   wp:inline → wp:extent → wp:docPr → a:graphic → a:graphicData → pic:pic
///   pic:pic → pic:blipFill (a:blip + a:stretch) → pic:spPr (a:xfrm + a:ext)
/// No wp:effectExtent (Word omits it for simple inline images).
/// Namespace declarations on the elements that use them (matching Word output).
fn generate_image_paragraph_xml(
    rel_id: &str,
    width_pt: f32,
    height_pt: f32,
    pic_id: usize,
) -> String {
    // Convert points to EMU (1 pt = 12700 EMU)
    let cx = (width_pt * 12700.0).round() as i64;
    let cy = (height_pt * 12700.0).round() as i64;

    format!(
        concat!(
            "<w:p><w:r><w:drawing>",
            "<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" ",
            "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">",
            "<wp:extent cx=\"{}\" cy=\"{}\"/>",
            "<wp:docPr id=\"{}\" name=\"Picture {}\"/>",
            "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
            "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">",
            "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">",
            "<pic:blipFill>",
            "<a:blip r:embed=\"{}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>",
            "</pic:blipFill>",
            "<pic:spPr>",
            "<a:xfrm>",
            "<a:ext cx=\"{}\" cy=\"{}\"/>",
            "</a:xfrm>",
            "</pic:spPr>",
            "</pic:pic>",
            "</a:graphicData>",
            "</a:graphic>",
            "</wp:inline>",
            "</w:drawing></w:r></w:p>",
        ),
        cx, cy,
        pic_id, pic_id,
        rel_id,
        cx, cy,
    )
}

// ===========================================================================
// XML body splitting
// ===========================================================================

/// Split document XML into preamble (before body content), body segments, and postamble.
fn split_body(xml: &str) -> Result<(String, Vec<BodySegment>, String), ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut segments = Vec::new();

    // We'll track positions using the XML event stream
    let mut preamble = String::new();
    let mut postamble = String::new();
    let mut in_body = false;
    let mut body_depth: i32 = 0;
    let mut current_element = String::new();
    let mut current_writer: Option<Writer<Cursor<Vec<u8>>>> = None;
    let mut current_kind: Option<&str> = None;
    let mut wrote_body_start = false;
    let mut pre_writer = Writer::new(Cursor::new(Vec::new()));
    let mut post_collecting = false;
    let mut post_writer = Writer::new(Cursor::new(Vec::new()));

    loop {
        match reader.read_event()? {
            Event::Eof => break,
            event => {
                let local = match &event {
                    Event::Start(e) | Event::Empty(e) => {
                        Some(local_name_from_bytes(e.name().as_ref()))
                    }
                    Event::End(e) => Some(local_name_from_bytes(e.name().as_ref())),
                    _ => None,
                };

                if !in_body {
                    match &event {
                        Event::Start(e) if local.as_deref() == Some("body") => {
                            in_body = true;
                            wrote_body_start = true;
                            pre_writer.write_event(Event::Start(e.clone()))?;
                            continue;
                        }
                        _ => {
                            pre_writer.write_event(event)?;
                            continue;
                        }
                    }
                }

                // Inside body
                if post_collecting {
                    post_writer.write_event(event)?;
                    continue;
                }

                if body_depth == 0 {
                    match &event {
                        Event::Start(e) => {
                            let ln = local.as_deref().unwrap_or("");
                            body_depth = 1;
                            let mut w = Writer::new(Cursor::new(Vec::new()));
                            w.write_event(Event::Start(e.clone()))?;
                            current_kind = Some(match ln {
                                "p" => "p",
                                "tbl" => "tbl",
                                "sectPr" => "sectPr",
                                _ => "other",
                            });
                            current_writer = Some(w);
                        }
                        Event::Empty(e) => {
                            // Self-closing element at body level
                            let ln = local.as_deref().unwrap_or("");
                            let mut w = Writer::new(Cursor::new(Vec::new()));
                            w.write_event(Event::Empty(e.clone()))?;
                            let xml_bytes = w.into_inner().into_inner();
                            let xml_str = String::from_utf8(xml_bytes).unwrap_or_default();
                            let seg = match ln {
                                "sectPr" => BodySegment::SectPr(xml_str),
                                _ => BodySegment::Other(xml_str),
                            };
                            segments.push(seg);
                        }
                        Event::End(e) if local.as_deref() == Some("body") => {
                            // End of body
                            post_collecting = true;
                            post_writer.write_event(Event::End(e.clone()))?;
                        }
                        Event::Text(e) => {
                            // Whitespace between body elements — ignore
                        }
                        _ => {
                            // Other events at body level (comments, PI, etc.)
                        }
                    }
                } else {
                    // Inside a body-level element
                    if let Some(ref mut w) = current_writer {
                        match &event {
                            Event::Start(_) => {
                                body_depth += 1;
                                w.write_event(event)?;
                            }
                            Event::End(_) => {
                                body_depth -= 1;
                                w.write_event(event)?;
                                if body_depth == 0 {
                                    let writer = current_writer.take().unwrap();
                                    let xml_bytes = writer.into_inner().into_inner();
                                    let xml_str = String::from_utf8(xml_bytes).unwrap_or_default();
                                    let kind = current_kind.take().unwrap_or("other");
                                    let seg = match kind {
                                        "p" => BodySegment::Paragraph(xml_str),
                                        "tbl" => BodySegment::Table(xml_str),
                                        "sectPr" => BodySegment::SectPr(xml_str),
                                        _ => BodySegment::Other(xml_str),
                                    };
                                    segments.push(seg);
                                }
                            }
                            _ => {
                                w.write_event(event)?;
                            }
                        }
                    }
                }
            }
        }
    }

    preamble = String::from_utf8(pre_writer.into_inner().into_inner()).unwrap_or_default();
    postamble = String::from_utf8(post_writer.into_inner().into_inner()).unwrap_or_default();

    Ok((preamble, segments, postamble))
}

// ===========================================================================
// Paragraph XML patching
// ===========================================================================

/// Collect all child elements of an rPr or pPr element as (local_name, full_xml) pairs.
/// The input is the full element XML including the opening/closing tags.
fn collect_pr_children(pr_xml: &str) -> Vec<(String, String)> {
    let mut reader = Reader::from_str(pr_xml);
    let mut elements: Vec<(String, String)> = Vec::new();
    let mut depth: i32 = 0;
    let mut current_name = String::new();
    let mut current_writer: Option<Writer<Cursor<Vec<u8>>>> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(Event::Start(ref e)) => {
                depth += 1;
                if depth == 1 {
                    // This is the root element (rPr/pPr itself), skip
                    continue;
                }
                if depth == 2 {
                    // Direct child element start
                    current_name = local_name_from_bytes(e.name().as_ref());
                    let mut w = Writer::new(Cursor::new(Vec::new()));
                    w.write_event(Event::Start(e.clone())).ok();
                    current_writer = Some(w);
                } else if let Some(ref mut w) = current_writer {
                    w.write_event(Event::Start(e.clone())).ok();
                }
            }
            Ok(Event::End(ref e)) => {
                if depth == 2 {
                    if let Some(mut w) = current_writer.take() {
                        w.write_event(Event::End(e.clone())).ok();
                        let xml = String::from_utf8(w.into_inner().into_inner()).unwrap_or_default();
                        elements.push((current_name.clone(), xml));
                    }
                } else if depth > 2 {
                    if let Some(ref mut w) = current_writer {
                        w.write_event(Event::End(e.clone())).ok();
                    }
                }
                depth -= 1;
            }
            Ok(Event::Empty(ref e)) => {
                if depth == 1 {
                    // Direct child self-closing element
                    let name = local_name_from_bytes(e.name().as_ref());
                    let mut w = Writer::new(Cursor::new(Vec::new()));
                    w.write_event(Event::Empty(e.clone())).ok();
                    let xml = String::from_utf8(w.into_inner().into_inner()).unwrap_or_default();
                    elements.push((name, xml));
                } else if let Some(ref mut w) = current_writer {
                    w.write_event(Event::Empty(e.clone())).ok();
                }
            }
            Ok(event) => {
                if let Some(ref mut w) = current_writer {
                    w.write_event(event).ok();
                }
            }
            Err(_) => break,
        }
    }
    elements
}

/// Merge RunProps into existing rPr XML. Returns merged `<w:rPr>...</w:rPr>`.
fn merge_rpr_xml(existing_rpr: &str, props: &RunProps) -> String {
    let mut elements = collect_pr_children(existing_rpr);

    // Helper to remove elements by local name
    fn remove_by_name(elements: &mut Vec<(String, String)>, name: &str) {
        elements.retain(|(n, _)| n != name);
    }

    // Apply overrides
    if let Some(bold) = props.bold {
        remove_by_name(&mut elements, "b");
        remove_by_name(&mut elements, "bCs");
        if bold {
            elements.push(("b".to_string(), "<w:b/>".to_string()));
            elements.push(("bCs".to_string(), "<w:bCs/>".to_string()));
        }
    }
    if let Some(italic) = props.italic {
        remove_by_name(&mut elements, "i");
        remove_by_name(&mut elements, "iCs");
        if italic {
            elements.push(("i".to_string(), "<w:i/>".to_string()));
            elements.push(("iCs".to_string(), "<w:iCs/>".to_string()));
        }
    }
    if let Some(underline) = props.underline {
        remove_by_name(&mut elements, "u");
        if underline {
            let style = props.underline_style.as_deref().unwrap_or("single");
            elements.push(("u".to_string(), format!(r#"<w:u w:val="{}"/>"#, style)));
        }
    }
    if let Some(strike) = props.strikethrough {
        remove_by_name(&mut elements, "strike");
        if strike {
            elements.push(("strike".to_string(), "<w:strike/>".to_string()));
        }
    }
    if let Some(ref ff) = props.font_family {
        remove_by_name(&mut elements, "rFonts");
        let ea = props.font_family_east_asia.as_deref().unwrap_or(ff);
        elements.push(("rFonts".to_string(), format!(
            r#"<w:rFonts w:ascii="{}" w:hAnsi="{}" w:eastAsia="{}"/>"#,
            ff, ff, ea
        )));
    } else if let Some(ref ea) = props.font_family_east_asia {
        // Only update eastAsia — we need to preserve existing ascii/hAnsi
        // For simplicity, replace the whole rFonts with eastAsia only
        // A more complete implementation would parse existing rFonts attrs
        remove_by_name(&mut elements, "rFonts");
        elements.push(("rFonts".to_string(), format!(r#"<w:rFonts w:eastAsia="{}"/>"#, ea)));
    }
    if let Some(size) = props.font_size {
        remove_by_name(&mut elements, "sz");
        remove_by_name(&mut elements, "szCs");
        let half_pt = (size * 2.0).round() as u32;
        elements.push(("sz".to_string(), format!(r#"<w:sz w:val="{}"/>"#, half_pt)));
        elements.push(("szCs".to_string(), format!(r#"<w:szCs w:val="{}"/>"#, half_pt)));
    }
    if let Some(ref color) = props.color {
        remove_by_name(&mut elements, "color");
        elements.push(("color".to_string(), format!(r#"<w:color w:val="{}"/>"#, color)));
    }
    if let Some(ref hl) = props.highlight {
        remove_by_name(&mut elements, "highlight");
        elements.push(("highlight".to_string(), format!(r#"<w:highlight w:val="{}"/>"#, hl)));
    }

    // Character spacing (twips)
    if let Some(cs) = props.character_spacing {
        remove_by_name(&mut elements, "spacing");
        let twips = (cs * 20.0).round() as i32;
        elements.push(("spacing".to_string(), format!(r#"<w:spacing w:val="{}"/>"#, twips)));
    }

    // Kerning threshold (half-points)
    if let Some(kern) = props.kerning {
        remove_by_name(&mut elements, "kern");
        let half_pt = (kern * 2.0).round() as u32;
        elements.push(("kern".to_string(), format!(r#"<w:kern w:val="{}"/>"#, half_pt)));
    }

    // No proofing
    if let Some(np) = props.no_proof {
        remove_by_name(&mut elements, "noProof");
        if np {
            elements.push(("noProof".to_string(), "<w:noProof/>".to_string()));
        }
    }

    // Run style reference
    if let Some(ref rs) = props.run_style {
        remove_by_name(&mut elements, "rStyle");
        elements.push(("rStyle".to_string(), format!(r#"<w:rStyle w:val="{}"/>"#, rs)));
    }

    // Vertical align
    if let Some(ref va) = props.vertical_align {
        remove_by_name(&mut elements, "vertAlign");
        elements.push(("vertAlign".to_string(), format!(r#"<w:vertAlign w:val="{}"/>"#, va)));
    }

    // Language tag
    if props.lang.is_some() || props.lang_east_asia.is_some() || props.lang_bidi.is_some() {
        remove_by_name(&mut elements, "lang");
        let mut lang_attrs = Vec::new();
        if let Some(ref v) = props.lang {
            lang_attrs.push(format!(r#"w:val="{}""#, v));
        }
        if let Some(ref ea) = props.lang_east_asia {
            lang_attrs.push(format!(r#"w:eastAsia="{}""#, ea));
        }
        if let Some(ref bi) = props.lang_bidi {
            lang_attrs.push(format!(r#"w:bidi="{}""#, bi));
        }
        elements.push(("lang".to_string(), format!("<w:lang {}/>", lang_attrs.join(" "))));
    }

    if elements.is_empty() {
        String::new()
    } else {
        // Sort elements to match Word's output order (from Word's actual document.xml):
        //   rFonts → b → bCs → i → iCs → strike → color → sz → szCs → highlight → u
        // Unknown elements go to the end in their original order.
        sort_rpr_elements(&mut elements);
        let content: String = elements.iter().map(|(_, xml)| xml.as_str()).collect();
        format!("<w:rPr>{}</w:rPr>", content)
    }
}

/// Sort rPr child elements to match Word's output order.
fn sort_rpr_elements(elements: &mut Vec<(String, String)>) {
    // Word's rPr element order (observed from actual Word output)
    const RPR_ORDER: &[&str] = &[
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
        "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
        "noProof", "vanish", "color", "spacing", "w", "kern", "position",
        "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
        "fitText", "vertAlign", "rtl", "cs", "lang",
    ];

    elements.sort_by(|(a_name, _), (b_name, _)| {
        let a_pos = RPR_ORDER.iter().position(|&n| n == a_name).unwrap_or(999);
        let b_pos = RPR_ORDER.iter().position(|&n| n == b_name).unwrap_or(999);
        a_pos.cmp(&b_pos)
    });
}

/// Merge ParaProps into existing pPr XML. Returns merged `<w:pPr>...</w:pPr>`.
fn merge_ppr_xml(existing_ppr: &str, props: &ParaProps) -> String {
    let mut elements = collect_pr_children(existing_ppr);

    fn remove_by_name(elements: &mut Vec<(String, String)>, name: &str) {
        elements.retain(|(n, _)| n != name);
    }

    if let Some(ref style_id) = props.style_id {
        remove_by_name(&mut elements, "pStyle");
        elements.insert(0, ("pStyle".to_string(), format!(r#"<w:pStyle w:val="{}"/>"#, style_id)));
    }

    // Boolean toggle elements — Word uses bare element for true, w:val="0" for false
    fn merge_bool_toggle(elements: &mut Vec<(String, String)>, name: &str, val: Option<bool>) {
        if let Some(v) = val {
            elements.retain(|(n, _)| n != name);
            if v {
                elements.push((name.to_string(), format!("<w:{}/>", name)));
            } else {
                elements.push((name.to_string(), format!(r#"<w:{} w:val="0"/>"#, name)));
            }
        }
    }

    merge_bool_toggle(&mut elements, "keepNext", props.keep_next);
    merge_bool_toggle(&mut elements, "keepLines", props.keep_lines);
    merge_bool_toggle(&mut elements, "pageBreakBefore", props.page_break_before);
    merge_bool_toggle(&mut elements, "widowControl", props.widow_control);
    merge_bool_toggle(&mut elements, "wordWrap", props.word_wrap);
    if let Some(de) = props.auto_space_de {
        if !de {
            remove_by_name(&mut elements, "autoSpaceDE");
            elements.push(("autoSpaceDE".to_string(), r#"<w:autoSpaceDE w:val="0"/>"#.to_string()));
        }
    }
    if let Some(dn) = props.auto_space_dn {
        if !dn {
            remove_by_name(&mut elements, "autoSpaceDN");
            elements.push(("autoSpaceDN".to_string(), r#"<w:autoSpaceDN w:val="0"/>"#.to_string()));
        }
    }
    merge_bool_toggle(&mut elements, "adjustRightInd", props.adjust_right_ind);
    merge_bool_toggle(&mut elements, "snapToGrid", props.snap_to_grid);

    if let Some(ref align) = props.alignment {
        remove_by_name(&mut elements, "jc");
        elements.push(("jc".to_string(), format!(r#"<w:jc w:val="{}"/>"#, align)));
    }

    // Spacing — merge individual attributes
    if props.space_before.is_some() || props.space_after.is_some() || props.line_spacing.is_some() {
        // Extract existing spacing attributes if any
        let existing_spacing = elements.iter()
            .find(|(n, _)| n == "spacing")
            .map(|(_, xml)| xml.clone());

        let mut attrs: HashMap<String, String> = HashMap::new();
        if let Some(ref sp_xml) = existing_spacing {
            // Parse existing spacing attributes
            for part in sp_xml.split_whitespace() {
                if let Some((key, val)) = part.split_once('=') {
                    if key.starts_with("w:") {
                        let val = val.trim_matches('"').trim_matches('/').trim_matches('>');
                        attrs.insert(key.to_string(), val.to_string());
                    }
                }
            }
        }

        if let Some(before) = props.space_before {
            attrs.insert("w:before".to_string(), format!("{}", (before * 20.0).round() as i32));
        }
        if let Some(after) = props.space_after {
            attrs.insert("w:after".to_string(), format!("{}", (after * 20.0).round() as i32));
        }
        if let Some(ls) = props.line_spacing {
            let line_val = (ls * 240.0).round() as i32;
            attrs.insert("w:line".to_string(), format!("{}", line_val));
            attrs.insert("w:lineRule".to_string(), "auto".to_string());
        }

        remove_by_name(&mut elements, "spacing");
        let attr_str: String = attrs.iter()
            .map(|(k, v)| format!(r#"{}="{}""#, k, v))
            .collect::<Vec<_>>()
            .join(" ");
        elements.push(("spacing".to_string(), format!("<w:spacing {}/>", attr_str)));
    }

    // Indentation — merge individual attributes
    if props.indent_left.is_some() || props.indent_right.is_some() || props.indent_first_line.is_some() {
        let mut attrs: HashMap<String, String> = HashMap::new();

        // Parse existing ind attributes if any
        let existing_ind = elements.iter()
            .find(|(n, _)| n == "ind")
            .map(|(_, xml)| xml.clone());
        if let Some(ref ind_xml) = existing_ind {
            for part in ind_xml.split_whitespace() {
                if let Some((key, val)) = part.split_once('=') {
                    if key.starts_with("w:") {
                        let val = val.trim_matches('"').trim_matches('/').trim_matches('>');
                        attrs.insert(key.to_string(), val.to_string());
                    }
                }
            }
        }

        if let Some(left) = props.indent_left {
            attrs.insert("w:left".to_string(), format!("{}", (left * 20.0).round() as i32));
        }
        if let Some(right) = props.indent_right {
            attrs.insert("w:right".to_string(), format!("{}", (right * 20.0).round() as i32));
        }
        if let Some(first) = props.indent_first_line {
            attrs.remove("w:firstLine");
            attrs.remove("w:hanging");
            if first >= 0.0 {
                attrs.insert("w:firstLine".to_string(), format!("{}", (first * 20.0).round() as i32));
            } else {
                attrs.insert("w:hanging".to_string(), format!("{}", (-first * 20.0).round() as i32));
            }
        }

        remove_by_name(&mut elements, "ind");
        let attr_str: String = attrs.iter()
            .map(|(k, v)| format!(r#"{}="{}""#, k, v))
            .collect::<Vec<_>>()
            .join(" ");
        elements.push(("ind".to_string(), format!("<w:ind {}/>", attr_str)));
    }

    // Sort elements to match Word's output order
    sort_ppr_elements(&mut elements);
    let content: String = elements.iter().map(|(_, xml)| xml.as_str()).collect();
    format!("<w:pPr>{}</w:pPr>", content)
}

/// Sort pPr child elements to match Word's output order.
fn sort_ppr_elements(elements: &mut Vec<(String, String)>) {
    // Word's pPr element order (observed from actual Word output)
    const PPR_ORDER: &[&str] = &[
        "pStyle", "keepNext", "keepLines", "pageBreakBefore",
        "widowControl", "numPr", "suppressLineNumbers",
        "pBdr", "shd", "tabs", "suppressAutoHyphens",
        "kinsoku", "wordWrap", "overflowPunct", "topLinePunct",
        "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
        "snapToGrid", "spacing", "ind", "contextualSpacing",
        "mirrorIndents", "jc", "textDirection", "textAlignment",
        "outlineLvl", "rPr",
    ];

    elements.sort_by(|(a_name, _), (b_name, _)| {
        let a_pos = PPR_ORDER.iter().position(|&n| n == a_name).unwrap_or(999);
        let b_pos = PPR_ORDER.iter().position(|&n| n == b_name).unwrap_or(999);
        a_pos.cmp(&b_pos)
    });
}

/// Patch a single `<w:p>` XML fragment with text edits, formatting, run insertion/deletion.
fn patch_paragraph_xml(
    xml: &str,
    text_edits: &HashMap<usize, &String>,
    run_format_edits: &HashMap<usize, &RunProps>,
    run_insertions: &HashMap<usize, &(String, Option<RunProps>)>,
    run_deletions: &[usize],
    para_format: Option<&ParaProps>,
) -> Result<String, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));

    let mut run_idx: usize = 0;
    let mut in_run = false;
    let mut in_text = false;
    let mut text_replaced = false;
    let mut skip_run = false;
    let mut skip_depth: i32 = 0;
    // pPr/rPr collection for merging
    let mut collecting_ppr = false;
    let mut ppr_collector: Option<Writer<Cursor<Vec<u8>>>> = None;
    let mut ppr_depth: i32 = 0;
    let mut ppr_seen = false;
    let mut collecting_rpr = false;
    let mut rpr_collector: Option<Writer<Cursor<Vec<u8>>>> = None;
    let mut rpr_depth: i32 = 0;

    loop {
        match reader.read_event()? {
            Event::Eof => break,
            Event::Start(ref e) => {
                let local = local_name_from_bytes(e.name().as_ref());

                if skip_run {
                    skip_depth += 1;
                    continue;
                }

                // Collecting pPr content for merge
                if collecting_ppr {
                    ppr_depth += 1;
                    if let Some(ref mut w) = ppr_collector {
                        w.write_event(Event::Start(e.clone()))?;
                    }
                    continue;
                }

                // Collecting rPr content for merge
                if collecting_rpr {
                    rpr_depth += 1;
                    if let Some(ref mut w) = rpr_collector {
                        w.write_event(Event::Start(e.clone()))?;
                    }
                    continue;
                }

                match local.as_str() {
                    "p" => {
                        writer.write_event(Event::Start(e.clone()))?;
                    }
                    "pPr" => {
                        ppr_seen = true;
                        if para_format.is_some() {
                            // Collect existing pPr for merging
                            collecting_ppr = true;
                            ppr_depth = 1;
                            let mut w = Writer::new(Cursor::new(Vec::new()));
                            w.write_event(Event::Start(e.clone()))?;
                            ppr_collector = Some(w);
                        } else {
                            writer.write_event(Event::Start(e.clone()))?;
                        }
                    }
                    "r" if !collecting_ppr => {
                        if run_deletions.contains(&run_idx) {
                            skip_run = true;
                            skip_depth = 1;
                            continue;
                        }

                        // Insert run before this one if needed
                        if let Some((text, style)) = run_insertions.get(&run_idx) {
                            let new_run = generate_run_xml(text, style.as_ref());
                            writer.write_event(Event::Text(BytesText::from_escaped(&new_run)))?;
                        }

                        in_run = true;
                        writer.write_event(Event::Start(e.clone()))?;
                    }
                    "rPr" if in_run => {
                        if run_format_edits.contains_key(&run_idx) {
                            // Collect existing rPr for merging
                            collecting_rpr = true;
                            rpr_depth = 1;
                            let mut w = Writer::new(Cursor::new(Vec::new()));
                            w.write_event(Event::Start(e.clone()))?;
                            rpr_collector = Some(w);
                        } else {
                            writer.write_event(Event::Start(e.clone()))?;
                        }
                    }
                    "t" if in_run => {
                        in_text = true;
                        text_replaced = false;
                        writer.write_event(Event::Start(e.clone()))?;
                    }
                    _ => {
                        writer.write_event(Event::Start(e.clone()))?;
                    }
                }
            }
            Event::End(ref e) => {
                let local = local_name_from_bytes(e.name().as_ref());

                if skip_run {
                    skip_depth -= 1;
                    if skip_depth == 0 {
                        skip_run = false;
                        run_idx += 1;
                    }
                    continue;
                }

                // Collecting pPr
                if collecting_ppr {
                    ppr_depth -= 1;
                    if let Some(ref mut w) = ppr_collector {
                        w.write_event(Event::End(e.clone()))?;
                    }
                    if ppr_depth == 0 {
                        collecting_ppr = false;
                        if let Some(w) = ppr_collector.take() {
                            let existing = String::from_utf8(w.into_inner().into_inner()).unwrap_or_default();
                            let merged = merge_ppr_xml(&existing, para_format.unwrap());
                            writer.write_event(Event::Text(BytesText::from_escaped(&merged)))?;
                        }
                    }
                    continue;
                }

                // Collecting rPr
                if collecting_rpr {
                    rpr_depth -= 1;
                    if let Some(ref mut w) = rpr_collector {
                        w.write_event(Event::End(e.clone()))?;
                    }
                    if rpr_depth == 0 {
                        collecting_rpr = false;
                        if let Some(w) = rpr_collector.take() {
                            let existing = String::from_utf8(w.into_inner().into_inner()).unwrap_or_default();
                            let fmt = run_format_edits.get(&run_idx).unwrap();
                            let merged = merge_rpr_xml(&existing, fmt);
                            if !merged.is_empty() {
                                writer.write_event(Event::Text(BytesText::from_escaped(&merged)))?;
                            }
                        }
                    }
                    continue;
                }

                match local.as_str() {
                    "p" => {
                        // Insert paragraph format if no pPr was seen
                        if !ppr_seen {
                            if let Some(fmt) = para_format {
                                // We need to inject pPr right after <w:p>.
                                // Since we already wrote events, we need to handle this differently.
                                // Actually, this case means we need to prepend pPr.
                                // For now, we'll handle this by rewriting the output.
                                // A simpler approach: generate pPr XML and inject it.
                                let ppr = generate_ppr_xml(fmt);
                                if !ppr.is_empty() {
                                    // This is tricky because we've already written content.
                                    // We'll handle this case in a post-processing step.
                                }
                            }
                        }

                        // Insert runs at the end if needed
                        if let Some((text, style)) = run_insertions.get(&run_idx) {
                            let new_run = generate_run_xml(text, style.as_ref());
                            writer.write_event(Event::Text(BytesText::from_escaped(&new_run)))?;
                        }

                        writer.write_event(Event::End(e.clone()))?;
                    }
                    "r" if in_run => {
                        // If this run needs formatting but had no rPr, add one
                        if let Some(fmt) = run_format_edits.get(&run_idx) {
                            // rPr wasn't seen for this run — we need to inject it
                            // But we've already written the run content...
                            // This is handled by checking if rPr was seen when run started.
                        }
                        in_run = false;
                        writer.write_event(Event::End(e.clone()))?;
                        run_idx += 1;
                    }
                    "t" if in_text => {
                        if !text_replaced {
                            if let Some(new_text) = text_edits.get(&run_idx) {
                                writer.write_event(Event::Text(BytesText::new(new_text)))?;
                            }
                        }
                        in_text = false;
                        writer.write_event(Event::End(e.clone()))?;
                    }
                    _ => {
                        writer.write_event(Event::End(e.clone()))?;
                    }
                }
            }
            Event::Text(ref e) => {
                if skip_run || collecting_ppr || collecting_rpr {
                    if collecting_ppr {
                        if let Some(ref mut w) = ppr_collector {
                            w.write_event(Event::Text(e.clone()))?;
                        }
                    }
                    if collecting_rpr {
                        if let Some(ref mut w) = rpr_collector {
                            w.write_event(Event::Text(e.clone()))?;
                        }
                    }
                    continue;
                }

                if in_text {
                    if let Some(new_text) = text_edits.get(&run_idx) {
                        writer.write_event(Event::Text(BytesText::new(new_text)))?;
                        text_replaced = true;
                    } else {
                        writer.write_event(Event::Text(e.clone()))?;
                    }
                } else {
                    writer.write_event(Event::Text(e.clone()))?;
                }
            }
            Event::Empty(ref e) => {
                if skip_run {
                    continue;
                }

                let local = local_name_from_bytes(e.name().as_ref());

                if collecting_ppr {
                    if let Some(ref mut w) = ppr_collector {
                        w.write_event(Event::Empty(e.clone()))?;
                    }
                    continue;
                }

                if collecting_rpr {
                    if let Some(ref mut w) = rpr_collector {
                        w.write_event(Event::Empty(e.clone()))?;
                    }
                    continue;
                }

                // Handle self-closing rPr element (empty rPr)
                if local == "rPr" && in_run {
                    if let Some(fmt) = run_format_edits.get(&run_idx) {
                        // Empty rPr — generate new one from scratch
                        let new_rpr = generate_rpr_xml(fmt);
                        if !new_rpr.is_empty() {
                            writer.write_event(Event::Text(BytesText::from_escaped(&new_rpr)))?;
                        }
                        continue;
                    }
                }

                writer.write_event(Event::Empty(e.clone()))?;
            }
            event => {
                if skip_run || collecting_ppr || collecting_rpr {
                    continue;
                }
                writer.write_event(event)?;
            }
        }
    }

    // Post-processing: if paragraph format was requested but no pPr existed,
    // inject it after the opening <w:p> tag
    let mut result_bytes = writer.into_inner().into_inner();
    if !ppr_seen && para_format.is_some() {
        let result_str = String::from_utf8(result_bytes).map_err(|_| ParseError::InvalidAttribute("UTF-8 error".to_string()))?;
        let ppr = generate_ppr_xml(para_format.unwrap());
        if !ppr.is_empty() {
            // Find the end of the opening <w:p...> tag and inject pPr after it
            if let Some(pos) = result_str.find('>') {
                let mut new_result = String::with_capacity(result_str.len() + ppr.len());
                new_result.push_str(&result_str[..pos + 1]);
                new_result.push_str(&ppr);
                new_result.push_str(&result_str[pos + 1..]);
                return Ok(new_result);
            }
        }
        return Ok(result_str);
    }

    String::from_utf8(result_bytes).map_err(|_| ParseError::InvalidAttribute("UTF-8 error".to_string()))
}

/// Handle the case where a run needs formatting but has no existing rPr.
/// This is done by post-processing: if SetRunFormat targets a run that has no rPr,
/// we inject one after the `<w:r>` or `<w:r ...>` opening tag.
fn inject_rpr_into_run(run_xml: &str, props: &RunProps) -> String {
    let rpr = generate_rpr_xml(props);
    if rpr.is_empty() {
        return run_xml.to_string();
    }
    // Find the end of the opening <w:r> tag
    if let Some(pos) = run_xml.find('>') {
        let mut result = String::with_capacity(run_xml.len() + rpr.len());
        result.push_str(&run_xml[..pos + 1]);
        result.push_str(&rpr);
        result.push_str(&run_xml[pos + 1..]);
        result
    } else {
        run_xml.to_string()
    }
}

// ===========================================================================
// Table XML patching
// ===========================================================================

/// Patch a `<w:tbl>` XML fragment with cell text edits, row insertions/deletions.
fn patch_table_xml(
    xml: &str,
    cell_text_edits: &HashMap<(usize, usize), &String>,
    row_insertions: &[(usize, &Vec<String>)],
    row_deletions: &[usize],
) -> Result<String, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));

    let mut row_idx: usize = 0;
    let mut col_idx: usize = 0;
    let mut in_row = false;
    let mut in_cell = false;
    let mut in_cell_p = false;
    let mut in_cell_r = false;
    let mut in_cell_t = false;
    let mut cell_text_replaced = false;
    let mut skip_row = false;
    let mut skip_depth: i32 = 0;
    let mut cell_depth: i32 = 0;

    // Determine default column width from existing grid (approximate)
    let default_col_width = 4675; // twips

    loop {
        match reader.read_event()? {
            Event::Eof => break,
            Event::Start(ref e) => {
                let local = local_name_from_bytes(e.name().as_ref());

                if skip_row {
                    skip_depth += 1;
                    continue;
                }

                match local.as_str() {
                    "tr" => {
                        // Check for row insertion before this row
                        for (ins_idx, cells) in row_insertions {
                            if *ins_idx == row_idx {
                                let row_xml = generate_table_row_xml(cells, default_col_width);
                                writer.write_event(Event::Text(BytesText::from_escaped(&row_xml)))?;
                            }
                        }

                        if row_deletions.contains(&row_idx) {
                            skip_row = true;
                            skip_depth = 1;
                            continue;
                        }

                        in_row = true;
                        col_idx = 0;
                        writer.write_event(Event::Start(e.clone()))?;
                    }
                    "tc" if in_row => {
                        in_cell = true;
                        cell_depth = 1;
                        writer.write_event(Event::Start(e.clone()))?;
                    }
                    "p" if in_cell && !in_cell_p => {
                        in_cell_p = true;
                        writer.write_event(Event::Start(e.clone()))?;
                    }
                    "r" if in_cell_p => {
                        in_cell_r = true;
                        writer.write_event(Event::Start(e.clone()))?;
                    }
                    "t" if in_cell_r => {
                        in_cell_t = true;
                        cell_text_replaced = false;
                        writer.write_event(Event::Start(e.clone()))?;
                    }
                    _ => {
                        if in_cell {
                            cell_depth += 1;
                        }
                        writer.write_event(Event::Start(e.clone()))?;
                    }
                }
            }
            Event::End(ref e) => {
                let local = local_name_from_bytes(e.name().as_ref());

                if skip_row {
                    skip_depth -= 1;
                    if skip_depth == 0 {
                        skip_row = false;
                        row_idx += 1;
                    }
                    continue;
                }

                match local.as_str() {
                    "tbl" => {
                        // Check for row insertion at end
                        for (ins_idx, cells) in row_insertions {
                            if *ins_idx == row_idx {
                                let row_xml = generate_table_row_xml(cells, default_col_width);
                                writer.write_event(Event::Text(BytesText::from_escaped(&row_xml)))?;
                            }
                        }
                        writer.write_event(Event::End(e.clone()))?;
                    }
                    "tr" if in_row => {
                        in_row = false;
                        writer.write_event(Event::End(e.clone()))?;
                        row_idx += 1;
                    }
                    "tc" if in_cell => {
                        cell_depth -= 1;
                        if cell_depth == 0 {
                            in_cell = false;
                            writer.write_event(Event::End(e.clone()))?;
                            col_idx += 1;
                        } else {
                            writer.write_event(Event::End(e.clone()))?;
                        }
                    }
                    "p" if in_cell_p => {
                        in_cell_p = false;
                        writer.write_event(Event::End(e.clone()))?;
                    }
                    "r" if in_cell_r => {
                        in_cell_r = false;
                        writer.write_event(Event::End(e.clone()))?;
                    }
                    "t" if in_cell_t => {
                        if !cell_text_replaced {
                            if let Some(new_text) = cell_text_edits.get(&(row_idx, col_idx)) {
                                writer.write_event(Event::Text(BytesText::new(new_text)))?;
                            }
                        }
                        in_cell_t = false;
                        writer.write_event(Event::End(e.clone()))?;
                    }
                    _ => {
                        if in_cell {
                            cell_depth -= 1;
                        }
                        writer.write_event(Event::End(e.clone()))?;
                    }
                }
            }
            Event::Text(ref e) => {
                if skip_row {
                    continue;
                }

                if in_cell_t {
                    if let Some(new_text) = cell_text_edits.get(&(row_idx, col_idx)) {
                        writer.write_event(Event::Text(BytesText::new(new_text)))?;
                        cell_text_replaced = true;
                    } else {
                        writer.write_event(Event::Text(e.clone()))?;
                    }
                } else {
                    writer.write_event(Event::Text(e.clone()))?;
                }
            }
            event => {
                if skip_row {
                    continue;
                }
                writer.write_event(event)?;
            }
        }
    }

    let result = writer.into_inner().into_inner();
    String::from_utf8(result).map_err(|_| ParseError::InvalidAttribute("UTF-8 error".to_string()))
}

// ===========================================================================
// Utilities
// ===========================================================================

/// Extract local name from a potentially namespaced XML tag.
fn local_name_from_bytes(name: &[u8]) -> String {
    let s = std::str::from_utf8(name).unwrap_or("");
    match s.rfind(':') {
        Some(pos) => s[pos + 1..].to_string(),
        None => s.to_string(),
    }
}

/// Escape XML special characters (public access for build_docx).
pub fn escape_xml_public(s: &str) -> String {
    escape_xml(s)
}

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

// Keep old function name for compatibility
fn local_name(name: &[u8]) -> String {
    local_name_from_bytes(name)
}

// ===========================================================================
// Tests
// ===========================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_editor_round_trip_no_edits() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let editor = DocxEditor::new(data).expect("should open");
        assert!(!editor.has_edits());

        let saved = editor.save().expect("should save");
        let doc = parse_docx(&saved).expect("saved docx should parse");
        assert_eq!(doc.pages[0].blocks.len(), 4);

        if let crate::ir::Block::Paragraph(p) = &doc.pages[0].blocks[0] {
            assert_eq!(p.runs[0].text, "Oxidocs Test Document");
        } else {
            panic!("expected paragraph");
        }
    }

    #[test]
    fn test_editor_change_heading_text() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let mut editor = DocxEditor::new(data).expect("should open");

        editor.set_run_text(0, 0, "New Heading".to_string());
        assert!(editor.has_edits());

        let saved = editor.save().expect("should save");
        let doc = parse_docx(&saved).expect("saved docx should parse");

        if let crate::ir::Block::Paragraph(p) = &doc.pages[0].blocks[0] {
            assert_eq!(p.runs[0].text, "New Heading");
            assert!(p.runs[0].style.bold);
            assert_eq!(p.style.heading_level, Some(1));
        } else {
            panic!("expected paragraph");
        }

        if let crate::ir::Block::Paragraph(p) = &doc.pages[0].blocks[2] {
            assert!(p.runs[0].text.contains("日本語"));
        } else {
            panic!("expected paragraph");
        }
    }

    #[test]
    fn test_editor_change_multiple_runs() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let mut editor = DocxEditor::new(data).expect("should open");

        editor.apply_edits(&[
            TextEdit {
                paragraph_index: 0,
                run_index: 0,
                new_text: "タイトル変更".to_string(),
            },
            TextEdit {
                paragraph_index: 2,
                run_index: 0,
                new_text: "日本語テキスト変更済み".to_string(),
            },
        ]);

        let saved = editor.save().expect("should save");
        let doc = parse_docx(&saved).expect("saved docx should parse");

        if let crate::ir::Block::Paragraph(p) = &doc.pages[0].blocks[0] {
            assert_eq!(p.runs[0].text, "タイトル変更");
        } else {
            panic!("expected paragraph");
        }

        if let crate::ir::Block::Paragraph(p) = &doc.pages[0].blocks[2] {
            assert_eq!(p.runs[0].text, "日本語テキスト変更済み");
        } else {
            panic!("expected paragraph");
        }

        if let crate::ir::Block::Table(t) = &doc.pages[0].blocks[3] {
            assert_eq!(t.rows.len(), 2);
            assert!(t.style.border);
        } else {
            panic!("expected table");
        }
    }

    #[test]
    fn test_editor_preserves_all_zip_entries() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let mut editor = DocxEditor::new(data).expect("should open");
        editor.set_run_text(0, 0, "Modified".to_string());

        let saved = editor.save().expect("should save");

        let mut orig = ZipArchive::new(Cursor::new(data.to_vec())).unwrap();
        let mut saved_zip = ZipArchive::new(Cursor::new(&saved)).unwrap();

        let mut orig_names: Vec<String> = (0..orig.len())
            .map(|i| orig.by_index(i).unwrap().name().to_string())
            .collect();
        let mut saved_names: Vec<String> = (0..saved_zip.len())
            .map(|i| saved_zip.by_index(i).unwrap().name().to_string())
            .collect();
        orig_names.sort();
        saved_names.sort();
        assert_eq!(orig_names, saved_names);
    }

    #[test]
    fn test_insert_paragraph() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let mut editor = DocxEditor::new(data).expect("should open");

        // Insert a new paragraph at the beginning
        editor.insert_paragraph(0, "Inserted First", None, None);

        let saved = editor.save().expect("should save");
        let doc = parse_docx(&saved).expect("saved docx should parse");

        // New paragraph should be first
        if let crate::ir::Block::Paragraph(p) = &doc.pages[0].blocks[0] {
            assert_eq!(p.runs[0].text, "Inserted First");
        } else {
            panic!("expected paragraph");
        }

        // Original heading should now be second
        if let crate::ir::Block::Paragraph(p) = &doc.pages[0].blocks[1] {
            assert_eq!(p.runs[0].text, "Oxidocs Test Document");
        } else {
            panic!("expected paragraph");
        }
    }

    #[test]
    fn test_delete_paragraph() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let mut editor = DocxEditor::new(data).expect("should open");

        // Delete the heading (paragraph 0)
        editor.delete_paragraph(0);

        let saved = editor.save().expect("should save");
        let doc = parse_docx(&saved).expect("saved docx should parse");

        // First block should now be the mixed formatting paragraph
        if let crate::ir::Block::Paragraph(p) = &doc.pages[0].blocks[0] {
            assert!(!p.runs[0].text.contains("Oxidocs Test Document"));
        } else {
            panic!("expected paragraph");
        }
    }

    #[test]
    fn test_insert_paragraph_with_formatting() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let mut editor = DocxEditor::new(data).expect("should open");

        let style = RunProps {
            bold: Some(true),
            font_size: Some(24.0),
            color: Some("FF0000".to_string()),
            ..Default::default()
        };
        let para = ParaProps {
            alignment: Some("center".to_string()),
            ..Default::default()
        };
        editor.insert_paragraph(0, "Red Bold Title", Some(style), Some(para));

        let saved = editor.save().expect("should save");
        let doc = parse_docx(&saved).expect("saved docx should parse");

        if let crate::ir::Block::Paragraph(p) = &doc.pages[0].blocks[0] {
            assert_eq!(p.runs[0].text, "Red Bold Title");
            assert!(p.runs[0].style.bold);
            assert_eq!(p.runs[0].style.font_size, Some(24.0));
            assert_eq!(p.runs[0].style.color.as_deref(), Some("FF0000"));
            assert_eq!(p.alignment, crate::ir::Alignment::Center);
        } else {
            panic!("expected paragraph");
        }
    }

    #[test]
    fn test_insert_table() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let mut editor = DocxEditor::new(data).expect("should open");

        let content = vec![
            vec!["A1".to_string(), "B1".to_string()],
            vec!["A2".to_string(), "B2".to_string()],
        ];
        editor.insert_table(0, 2, 2, Some(content), None);

        let saved = editor.save().expect("should save");
        let doc = parse_docx(&saved).expect("saved docx should parse");

        // First block should be a table
        if let crate::ir::Block::Table(t) = &doc.pages[0].blocks[0] {
            assert_eq!(t.rows.len(), 2);
            assert_eq!(t.rows[0].cells.len(), 2);
            // Check cell text
            if let crate::ir::Block::Paragraph(p) = &t.rows[0].cells[0].blocks[0] {
                assert_eq!(p.runs[0].text, "A1");
            }
            if let crate::ir::Block::Paragraph(p) = &t.rows[0].cells[1].blocks[0] {
                assert_eq!(p.runs[0].text, "B1");
            }
        } else {
            panic!("expected table, got {:?}", &doc.pages[0].blocks[0]);
        }
    }

    #[test]
    fn test_set_cell_text() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let mut editor = DocxEditor::new(data).expect("should open");

        // basic_test.docx has a table at block index 3, which is table index 0
        editor.set_cell_text(0, 0, 0, "Modified Cell");

        let saved = editor.save().expect("should save");
        let doc = parse_docx(&saved).expect("saved docx should parse");

        if let crate::ir::Block::Table(t) = &doc.pages[0].blocks[3] {
            if let crate::ir::Block::Paragraph(p) = &t.rows[0].cells[0].blocks[0] {
                assert_eq!(p.runs[0].text, "Modified Cell");
            } else {
                panic!("expected paragraph in cell");
            }
        } else {
            panic!("expected table at index 3");
        }
    }

    #[test]
    fn test_insert_image() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let mut editor = DocxEditor::new(data).expect("should open");

        // Create a minimal 1x1 PNG
        let png_data = create_minimal_png();
        editor.insert_image(0, png_data, 100.0, 100.0, "image/png");

        let saved = editor.save().expect("should save");

        // Verify the ZIP contains the image
        let mut saved_zip = ZipArchive::new(Cursor::new(&saved)).unwrap();
        let names: Vec<String> = (0..saved_zip.len())
            .map(|i| saved_zip.by_index(i).unwrap().name().to_string())
            .collect();
        assert!(names.iter().any(|n| n.starts_with("word/media/")));

        // Verify it parses (image may appear as block)
        let doc = parse_docx(&saved).expect("saved docx should parse");
        assert!(!doc.pages[0].blocks.is_empty());
    }

    #[test]
    fn test_create_blank_and_edit() {
        let bytes = crate::create_blank_docx();
        let mut editor = DocxEditor::new(&bytes).expect("should open blank");

        // Add content
        editor.set_run_text(0, 0, "Hello World".to_string());
        editor.insert_paragraph(
            1,
            "Second paragraph",
            Some(RunProps {
                italic: Some(true),
                ..Default::default()
            }),
            None,
        );

        let saved = editor.save().expect("should save");
        let doc = parse_docx(&saved).expect("should parse");

        if let crate::ir::Block::Paragraph(p) = &doc.pages[0].blocks[0] {
            assert_eq!(p.runs[0].text, "Hello World");
        } else {
            panic!("expected paragraph");
        }

        if let crate::ir::Block::Paragraph(p) = &doc.pages[0].blocks[1] {
            assert_eq!(p.runs[0].text, "Second paragraph");
            assert!(p.runs[0].style.italic);
        } else {
            panic!("expected paragraph");
        }
    }

    #[test]
    fn test_set_run_format_merges_with_existing() {
        // The heading in basic_test.docx is bold. Setting color should keep bold.
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let mut editor = DocxEditor::new(data).expect("should open");

        // Heading (para 0, run 0) is bold. Add color without losing bold.
        editor.set_run_format(0, 0, RunProps {
            color: Some("0000FF".to_string()),
            ..Default::default()
        });

        let saved = editor.save().expect("should save");
        let doc = parse_docx(&saved).expect("saved docx should parse");

        if let crate::ir::Block::Paragraph(p) = &doc.pages[0].blocks[0] {
            assert_eq!(p.runs[0].text, "Oxidocs Test Document");
            // Bold should be preserved (was in original)
            assert!(p.runs[0].style.bold, "bold should be preserved after format merge");
            // Color should be added
            assert_eq!(p.runs[0].style.color.as_deref(), Some("0000FF"));
        } else {
            panic!("expected paragraph");
        }
    }

    #[test]
    fn test_set_paragraph_format_merges_with_existing() {
        // Heading1 has space_before/space_after. Setting alignment should keep spacing.
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let mut editor = DocxEditor::new(data).expect("should open");

        editor.set_paragraph_format(0, ParaProps {
            alignment: Some("center".to_string()),
            ..Default::default()
        });

        let saved = editor.save().expect("should save");
        let doc = parse_docx(&saved).expect("saved docx should parse");

        if let crate::ir::Block::Paragraph(p) = &doc.pages[0].blocks[0] {
            assert_eq!(p.alignment, crate::ir::Alignment::Center);
            // Heading level should be preserved
            assert_eq!(p.style.heading_level, Some(1), "heading level should be preserved");
        } else {
            panic!("expected paragraph");
        }
    }

    /// Create a minimal valid PNG (1x1 red pixel).
    fn create_minimal_png() -> Vec<u8> {
        let mut png = Vec::new();
        // PNG signature
        png.extend_from_slice(&[137, 80, 78, 71, 13, 10, 26, 10]);
        // IHDR chunk
        let ihdr_data = [
            0, 0, 0, 1, // width
            0, 0, 0, 1, // height
            8,  // bit depth
            2,  // color type (RGB)
            0,  // compression
            0,  // filter
            0,  // interlace
        ];
        let ihdr_crc = crc32(&[b'I', b'H', b'D', b'R'], &ihdr_data);
        png.extend_from_slice(&(ihdr_data.len() as u32).to_be_bytes());
        png.extend_from_slice(b"IHDR");
        png.extend_from_slice(&ihdr_data);
        png.extend_from_slice(&ihdr_crc.to_be_bytes());
        // IDAT chunk (compressed 1x1 RGB pixel)
        let raw_data = [0, 255, 0, 0]; // filter byte + R, G, B
        let compressed = deflate_simple(&raw_data);
        let idat_crc = crc32(b"IDAT", &compressed);
        png.extend_from_slice(&(compressed.len() as u32).to_be_bytes());
        png.extend_from_slice(b"IDAT");
        png.extend_from_slice(&compressed);
        png.extend_from_slice(&idat_crc.to_be_bytes());
        // IEND chunk
        let iend_crc = crc32(b"IEND", &[]);
        png.extend_from_slice(&0u32.to_be_bytes());
        png.extend_from_slice(b"IEND");
        png.extend_from_slice(&iend_crc.to_be_bytes());
        png
    }

    fn crc32(chunk_type: &[u8], data: &[u8]) -> u32 {
        let mut crc: u32 = 0xFFFFFFFF;
        for &byte in chunk_type.iter().chain(data.iter()) {
            crc ^= byte as u32;
            for _ in 0..8 {
                if crc & 1 != 0 {
                    crc = (crc >> 1) ^ 0xEDB88320;
                } else {
                    crc >>= 1;
                }
            }
        }
        crc ^ 0xFFFFFFFF
    }

    fn deflate_simple(data: &[u8]) -> Vec<u8> {
        // Minimal zlib-wrapped deflate (stored block, no compression)
        let mut out = Vec::new();
        out.push(0x78); // CMF
        out.push(0x01); // FLG
        // BFINAL=1, BTYPE=00 (no compression)
        out.push(0x01);
        let len = data.len() as u16;
        out.extend_from_slice(&len.to_le_bytes());
        out.extend_from_slice(&(!len).to_le_bytes());
        out.extend_from_slice(data);
        // Adler-32
        let adler = adler32(data);
        out.extend_from_slice(&adler.to_be_bytes());
        out
    }

    fn adler32(data: &[u8]) -> u32 {
        let mut a: u32 = 1;
        let mut b: u32 = 0;
        for &byte in data {
            a = (a + byte as u32) % 65521;
            b = (b + a) % 65521;
        }
        (b << 16) | a
    }
}
