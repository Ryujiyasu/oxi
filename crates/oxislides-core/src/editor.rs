//! Round-trip pptx editor.
//!
//! Preserves the original ZIP archive. Patches `<a:t>` text nodes in slide XML
//! at specified (slide, shape, paragraph, run) coordinates.

use std::collections::HashMap;
use std::io::{Cursor, Read, Write};

use quick_xml::events::{BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::ir::Presentation;
use crate::parser::{parse_pptx, PptxError};
use oxi_common::archive::OoxmlArchive;
use oxi_common::relationships::parse_relationships;
use oxi_common::xml_utils::local_name;

/// A slide text edit operation.
#[derive(Debug, Clone)]
pub struct SlideTextEdit {
    /// 0-based slide index
    pub slide_index: usize,
    /// 0-based shape index within the slide
    pub shape_index: usize,
    /// 0-based paragraph index within the shape
    pub paragraph_index: usize,
    /// 0-based run index within the paragraph
    pub run_index: usize,
    /// New text
    pub new_text: String,
}

/// Round-trip pptx editor.
pub struct PptxEditor {
    original_data: Vec<u8>,
    presentation: Presentation,
    /// (slide_idx) -> { (shape, para, run) -> text }
    edits: HashMap<usize, HashMap<(usize, usize, usize), String>>,
}

impl PptxEditor {
    pub fn new(data: &[u8]) -> Result<Self, PptxError> {
        let presentation = parse_pptx(data)?;
        Ok(Self {
            original_data: data.to_vec(),
            presentation,
            edits: HashMap::new(),
        })
    }

    pub fn presentation(&self) -> &Presentation {
        &self.presentation
    }

    pub fn set_run_text(
        &mut self,
        slide_index: usize,
        shape_index: usize,
        paragraph_index: usize,
        run_index: usize,
        new_text: String,
    ) {
        self.edits
            .entry(slide_index)
            .or_default()
            .insert((shape_index, paragraph_index, run_index), new_text);
    }

    pub fn apply_edits(&mut self, edits: &[SlideTextEdit]) {
        for e in edits {
            self.set_run_text(
                e.slide_index,
                e.shape_index,
                e.paragraph_index,
                e.run_index,
                e.new_text.clone(),
            );
        }
    }

    pub fn has_edits(&self) -> bool {
        !self.edits.is_empty()
    }

    pub fn save(&self) -> Result<Vec<u8>, PptxError> {
        if self.edits.is_empty() {
            return Ok(self.original_data.clone());
        }

        // Resolve slide index -> ZIP path
        let slide_paths = self.resolve_slide_paths()?;

        // Map path -> slide edits
        let mut path_edits: HashMap<String, &HashMap<(usize, usize, usize), String>> =
            HashMap::new();
        for (si, edits) in &self.edits {
            if let Some(path) = slide_paths.get(*si) {
                path_edits.insert(path.clone(), edits);
            }
        }

        let cursor = Cursor::new(&self.original_data);
        let mut archive =
            ZipArchive::new(cursor).map_err(|e| PptxError::InvalidData(e.to_string()))?;

        let mut output = Vec::new();
        {
            let mut writer = ZipWriter::new(Cursor::new(&mut output));

            for i in 0..archive.len() {
                let mut entry = archive
                    .by_index(i)
                    .map_err(|e| PptxError::InvalidData(e.to_string()))?;
                let name = entry.name().to_string();
                let options = SimpleFileOptions::default().compression_method(entry.compression());

                writer
                    .start_file(&name, options)
                    .map_err(|e| PptxError::InvalidData(e.to_string()))?;

                if let Some(slide_edits) = path_edits.get(&name) {
                    let mut xml = String::new();
                    entry
                        .read_to_string(&mut xml)
                        .map_err(|e| PptxError::InvalidData(e.to_string()))?;
                    let patched = patch_slide_xml(&xml, slide_edits)?;
                    writer
                        .write_all(patched.as_bytes())
                        .map_err(|e| PptxError::InvalidData(e.to_string()))?;
                } else {
                    let mut buf = Vec::new();
                    entry
                        .read_to_end(&mut buf)
                        .map_err(|e| PptxError::InvalidData(e.to_string()))?;
                    writer
                        .write_all(&buf)
                        .map_err(|e| PptxError::InvalidData(e.to_string()))?;
                }
            }

            writer
                .finish()
                .map_err(|e| PptxError::InvalidData(e.to_string()))?;
        }

        Ok(output)
    }

    fn resolve_slide_paths(&self) -> Result<Vec<String>, PptxError> {
        let mut archive = OoxmlArchive::new(&self.original_data)?;
        let pres_xml = archive.read_part("ppt/presentation.xml")?;
        let rels_xml = archive.read_part("ppt/_rels/presentation.xml.rels")?;

        let mut reader = Reader::from_str(&pres_xml);
        let mut r_ids = Vec::new();
        loop {
            match reader.read_event().map_err(PptxError::Xml)? {
                Event::Start(e) | Event::Empty(e) => {
                    if local_name(e.name().as_ref()) == "sldId" {
                        let r_id = {
                            let mut found = None;
                            for attr in e.attributes().flatten() {
                                let key =
                                    std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                                if key == "r:id" {
                                    found = Some(
                                        String::from_utf8_lossy(&attr.value).to_string(),
                                    );
                                    break;
                                }
                            }
                            found.unwrap_or_default()
                        };
                        if !r_id.is_empty() {
                            r_ids.push(r_id);
                        }
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
                let path = if target.starts_with('/') {
                    target.trim_start_matches('/').to_string()
                } else {
                    format!("ppt/{}", target)
                };
                paths.push(path);
            } else {
                paths.push(String::new());
            }
        }

        Ok(paths)
    }
}

/// Patch slide XML, replacing `<a:t>` text nodes at (shape, para, run) coordinates.
fn patch_slide_xml(
    xml: &str,
    edits: &HashMap<(usize, usize, usize), String>,
) -> Result<String, PptxError> {
    let mut reader = Reader::from_str(xml);
    let mut writer = Writer::new(Cursor::new(Vec::new()));

    let mut in_sp_tree = false;
    let mut shape_idx: usize = 0;
    let mut in_shape = false;
    let mut para_idx: usize = 0;
    let mut in_paragraph = false;
    let mut run_idx: usize = 0;
    let mut in_run = false;
    let mut in_text = false;

    loop {
        match reader.read_event().map_err(PptxError::Xml)? {
            Event::Eof => break,
            Event::Start(ref e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "spTree" => {
                        in_sp_tree = true;
                        shape_idx = 0;
                    }
                    "sp" | "pic" if in_sp_tree => {
                        in_shape = true;
                        para_idx = 0;
                    }
                    "p" if in_shape => {
                        in_paragraph = true;
                        run_idx = 0;
                    }
                    "r" if in_paragraph => {
                        in_run = true;
                    }
                    "t" if in_run => {
                        in_text = true;
                    }
                    _ => {}
                }
                writer
                    .write_event(Event::Start(e.clone()))
                    .map_err(|e| PptxError::InvalidData(e.to_string()))?;
            }
            Event::End(ref e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "spTree" => {
                        in_sp_tree = false;
                    }
                    "sp" | "pic" if in_shape => {
                        in_shape = false;
                        shape_idx += 1;
                    }
                    "p" if in_paragraph => {
                        in_paragraph = false;
                        para_idx += 1;
                    }
                    "r" if in_run => {
                        in_run = false;
                        run_idx += 1;
                    }
                    "t" if in_text => {
                        in_text = false;
                    }
                    _ => {}
                }
                writer
                    .write_event(Event::End(e.clone()))
                    .map_err(|e| PptxError::InvalidData(e.to_string()))?;
            }
            Event::Text(ref e) => {
                if in_text {
                    if let Some(new_text) = edits.get(&(shape_idx, para_idx, run_idx)) {
                        writer
                            .write_event(Event::Text(BytesText::new(new_text)))
                            .map_err(|e| PptxError::InvalidData(e.to_string()))?;
                    } else {
                        writer
                            .write_event(Event::Text(e.clone()))
                            .map_err(|e| PptxError::InvalidData(e.to_string()))?;
                    }
                } else {
                    writer
                        .write_event(Event::Text(e.clone()))
                        .map_err(|e| PptxError::InvalidData(e.to_string()))?;
                }
            }
            event => {
                writer
                    .write_event(event)
                    .map_err(|e| PptxError::InvalidData(e.to_string()))?;
            }
        }
    }

    let result = writer.into_inner().into_inner();
    String::from_utf8(result).map_err(|_| PptxError::InvalidData("UTF-8 error".to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_editor_round_trip() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.pptx");
        let editor = PptxEditor::new(data).expect("should open");
        let saved = editor.save().expect("should save");
        let pres = parse_pptx(&saved).expect("should parse");
        assert_eq!(pres.slides.len(), 1);
    }

    #[test]
    fn test_editor_change_title() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.pptx");
        let mut editor = PptxEditor::new(data).expect("should open");

        // Change the title text (slide 0, shape 0, para 0, run 0)
        editor.set_run_text(0, 0, 0, 0, "New Title".to_string());

        let saved = editor.save().expect("should save");
        let pres = parse_pptx(&saved).expect("should parse");

        let slide = &pres.slides[0];
        if let crate::ir::ShapeContent::TextBox { paragraphs } = &slide.shapes[0].content {
            assert_eq!(paragraphs[0].runs[0].text, "New Title");
        } else {
            panic!("Expected TextBox");
        }
    }
}
