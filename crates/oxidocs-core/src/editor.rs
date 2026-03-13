//! Round-trip docx editor.
//!
//! Preserves the original ZIP archive and XML structure.
//! Only patches `<w:t>` text nodes in `word/document.xml` on save,
//! so all formatting, styles, metadata, and unknown elements survive untouched.

use std::collections::HashMap;
use std::io::{Cursor, Read, Write};

use quick_xml::events::{BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::ir::Document;
use crate::parser::{parse_docx, ParseError};

/// A text edit operation keyed by (paragraph_index, run_index) within `w:body`.
/// Only body-level paragraphs are counted (table cells have their own coordinate space).
#[derive(Debug, Clone)]
pub struct TextEdit {
    /// 0-based paragraph index within the document body
    pub paragraph_index: usize,
    /// 0-based run index within the paragraph
    pub run_index: usize,
    /// New text content for this run
    pub new_text: String,
}

/// Round-trip docx editor that preserves the original archive.
pub struct DocxEditor {
    /// Original docx bytes (ZIP archive)
    original_data: Vec<u8>,
    /// Parsed IR for reading/display
    document: Document,
    /// Pending text edits: (para_idx, run_idx) -> new_text
    edits: HashMap<(usize, usize), String>,
}

impl DocxEditor {
    /// Create a new editor from raw .docx bytes.
    pub fn new(data: &[u8]) -> Result<Self, ParseError> {
        let document = parse_docx(data)?;
        Ok(Self {
            original_data: data.to_vec(),
            document,
            edits: HashMap::new(),
        })
    }

    /// Get a reference to the parsed document IR (read-only).
    pub fn document(&self) -> &Document {
        &self.document
    }

    /// Set the text of a specific run.
    pub fn set_run_text(&mut self, paragraph_index: usize, run_index: usize, new_text: String) {
        self.edits
            .insert((paragraph_index, run_index), new_text);
    }

    /// Apply multiple edits at once.
    pub fn apply_edits(&mut self, edits: &[TextEdit]) {
        for edit in edits {
            self.set_run_text(edit.paragraph_index, edit.run_index, edit.new_text.clone());
        }
    }

    /// Check if there are pending edits.
    pub fn has_edits(&self) -> bool {
        !self.edits.is_empty()
    }

    /// Save the edited document as new .docx bytes.
    ///
    /// Walks the original ZIP, copies every entry unchanged except `word/document.xml`,
    /// which is patched with the pending edits.
    pub fn save(&self) -> Result<Vec<u8>, ParseError> {
        let cursor = Cursor::new(&self.original_data);
        let mut archive = ZipArchive::new(cursor)?;

        let mut output = Vec::new();
        {
            let mut writer = ZipWriter::new(Cursor::new(&mut output));

            for i in 0..archive.len() {
                let mut entry = archive.by_index(i)?;
                let name = entry.name().to_string();
                let options = SimpleFileOptions::default()
                    .compression_method(entry.compression());

                writer.start_file(&name, options)?;

                if name == "word/document.xml" {
                    // Read original XML and patch it
                    let mut xml = String::new();
                    entry.read_to_string(&mut xml)?;
                    let patched = self.patch_document_xml(&xml)?;
                    writer.write_all(patched.as_bytes())?;
                } else {
                    // Copy unchanged
                    let mut buf = Vec::new();
                    entry.read_to_end(&mut buf)?;
                    writer.write_all(&buf)?;
                }
            }

            writer.finish()?;
        }

        Ok(output)
    }

    /// Patch `word/document.xml` by walking the XML event stream and replacing
    /// `<w:t>` text nodes at the specified (paragraph, run) coordinates.
    ///
    /// This preserves ALL original XML structure — namespaces, attributes, unknown
    /// elements, comments, processing instructions, etc.
    fn patch_document_xml(&self, xml: &str) -> Result<String, ParseError> {
        if self.edits.is_empty() {
            return Ok(xml.to_string());
        }

        let mut reader = Reader::from_str(xml);
        let mut writer = Writer::new(Cursor::new(Vec::new()));

        // Tracking state
        let mut in_body = false;
        let mut body_depth: i32 = 0;
        // Current paragraph/run index within body (only top-level, not inside tables)
        let mut para_idx: usize = 0;
        let mut run_idx: usize = 0;
        let mut in_paragraph = false;
        let mut in_run = false;
        let mut in_text = false;
        // Whether we already wrote replacement text for the current <w:t> element
        let mut text_replaced = false;
        // Depth tracking for nested elements (tables etc.) — we only edit top-level paragraphs
        let mut table_depth: usize = 0;

        loop {
            match reader.read_event()? {
                Event::Eof => break,
                Event::Start(ref e) => {
                    let local = local_name(e.name().as_ref());
                    match local.as_str() {
                        "body" => {
                            in_body = true;
                            body_depth = 0;
                        }
                        "tbl" if in_body && table_depth == 0 && body_depth == 0 => {
                            table_depth += 1;
                        }
                        "tbl" if table_depth > 0 => {
                            table_depth += 1;
                        }
                        "p" if in_body && table_depth == 0 && body_depth == 0 => {
                            in_paragraph = true;
                            run_idx = 0;
                        }
                        "p" if in_body && body_depth == 0 => {
                            // paragraph inside table — skip tracking
                        }
                        "r" if in_paragraph && table_depth == 0 => {
                            in_run = true;
                        }
                        "t" if in_run && table_depth == 0 => {
                            in_text = true;
                            text_replaced = false;
                        }
                        _ if in_body && !in_paragraph && table_depth == 0 && body_depth == 0 => {
                            body_depth += 1;
                        }
                        _ => {}
                    }
                    writer.write_event(Event::Start(e.clone()))?;
                }
                Event::End(ref e) => {
                    let local = local_name(e.name().as_ref());
                    match local.as_str() {
                        "body" => {
                            in_body = false;
                        }
                        "tbl" if table_depth > 0 => {
                            table_depth -= 1;
                        }
                        "p" if in_paragraph && table_depth == 0 => {
                            in_paragraph = false;
                            para_idx += 1;
                        }
                        "r" if in_run && table_depth == 0 => {
                            in_run = false;
                            run_idx += 1;
                        }
                        "t" if in_text && table_depth == 0 => {
                            // If the original <w:t> was empty and we have an edit,
                            // inject the replacement text before closing the element.
                            if !text_replaced {
                                if let Some(new_text) = self.edits.get(&(para_idx, run_idx)) {
                                    writer.write_event(Event::Text(BytesText::new(new_text)))?;
                                }
                            }
                            in_text = false;
                        }
                        _ if body_depth > 0 && !in_paragraph && table_depth == 0 => {
                            body_depth -= 1;
                        }
                        _ => {}
                    }
                    writer.write_event(Event::End(e.clone()))?;
                }
                Event::Text(ref e) => {
                    if in_text && table_depth == 0 {
                        // Check if this (para_idx, run_idx) has an edit
                        if let Some(new_text) = self.edits.get(&(para_idx, run_idx)) {
                            writer.write_event(Event::Text(BytesText::new(new_text)))?;
                            text_replaced = true;
                        } else {
                            writer.write_event(Event::Text(e.clone()))?;
                        }
                    } else {
                        writer.write_event(Event::Text(e.clone()))?;
                    }
                }
                event => {
                    writer.write_event(event)?;
                }
            }
        }

        let result = writer.into_inner().into_inner();
        String::from_utf8(result).map_err(|_| {
            ParseError::InvalidAttribute("UTF-8 error in patched XML".to_string())
        })
    }
}

/// Extract local name from a potentially namespaced XML tag.
fn local_name(name: &[u8]) -> String {
    let s = std::str::from_utf8(name).unwrap_or("");
    match s.rfind(':') {
        Some(pos) => s[pos + 1..].to_string(),
        None => s.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_editor_round_trip_no_edits() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.docx");
        let editor = DocxEditor::new(data).expect("should open");
        assert!(!editor.has_edits());

        // Save without edits — should produce a valid docx
        let saved = editor.save().expect("should save");
        let doc = parse_docx(&saved).expect("saved docx should parse");
        assert_eq!(doc.pages[0].blocks.len(), 4);

        // Verify text is unchanged
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

        // Change heading text (paragraph 0, run 0)
        editor.set_run_text(0, 0, "New Heading".to_string());
        assert!(editor.has_edits());

        let saved = editor.save().expect("should save");
        let doc = parse_docx(&saved).expect("saved docx should parse");

        // Heading should be changed
        if let crate::ir::Block::Paragraph(p) = &doc.pages[0].blocks[0] {
            assert_eq!(p.runs[0].text, "New Heading");
            // Style should be preserved
            assert!(p.runs[0].style.bold);
            assert_eq!(p.style.heading_level, Some(1));
        } else {
            panic!("expected paragraph");
        }

        // Other paragraphs should be unchanged
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

        // Apply multiple edits
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

        // Table should survive intact
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

        // Both original and saved should have the same ZIP entries
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
    fn test_patch_document_xml_directly() {
        // Test the XML patching logic with a minimal XML snippet
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p w:rsidR="001A2B3C">
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r w:rsidRPr="00AABB11">
        <w:rPr><w:b/></w:rPr>
        <w:t>Original Title</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Hello </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>World</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p>
      <w:r><w:t>After table</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#;

        // Create a minimal editor just to test patching
        // We'll call patch_document_xml directly
        let editor = DocxEditor {
            original_data: Vec::new(),
            document: Document {
                pages: vec![],
                styles: Default::default(),
                metadata: Default::default(),
                comments: vec![],
            },
            edits: HashMap::from([
                ((0, 0), "New Title".to_string()),
                ((1, 1), "Rust".to_string()),
                ((2, 0), "After table changed".to_string()),
            ]),
        };

        let patched = editor.patch_document_xml(xml).unwrap();

        // Verify the patched XML
        assert!(patched.contains("New Title"));
        assert!(!patched.contains("Original Title"));
        assert!(patched.contains("Hello ")); // para 1, run 0 unchanged
        assert!(patched.contains("Rust")); // para 1, run 1 changed
        assert!(!patched.contains(">World<")); // "World" replaced
        assert!(patched.contains("Cell")); // table cell unchanged (not tracked)
        assert!(patched.contains("After table changed")); // para 2 (after table) changed

        // Verify attributes preserved
        assert!(patched.contains("w:rsidR=\"001A2B3C\"")); // paragraph rsid preserved
        assert!(patched.contains("w:rsidRPr=\"00AABB11\"")); // run rsid preserved
        assert!(patched.contains("w:val=\"Heading1\"")); // style preserved
    }
}
