// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Round-trip pptx editor.
//!
//! Preserves the original ZIP archive. Patches `<a:t>` text nodes in slide XML
//! at specified (slide, shape, paragraph, run) coordinates.

use std::collections::HashMap;
use std::io::{Cursor, Read, Write};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::ir::Presentation;
use crate::parser::{parse_pptx, PptxError};
use oxidocs_common::archive::OoxmlArchive;
use oxidocs_common::relationships::parse_relationships;
use oxidocs_common::xml_utils::local_name;

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

/// A paragraph break: what pressing Enter asks for.
///
/// The paragraph is cut at `at_char` -- counted in CHARACTERS over the
/// paragraph's runs, the same way the layout counts them -- and the tail
/// becomes a new paragraph carrying a copy of the original's `a:pPr`, so the
/// two halves keep the level, alignment and bullet the whole had.
#[derive(Debug, Clone)]
pub struct SlideParagraphSplit {
    pub slide_index: usize,
    pub shape_index: usize,
    pub paragraph_index: usize,
    pub at_char: usize,
}

/// What to change about a run's look. `None` leaves a property alone.
///
/// `bold`/`italic`/`underline`/`font_size` are `a:rPr` ATTRIBUTES, rewritten
/// in place. `color` is a `<a:solidFill><a:srgbClr>` CHILD, a different shape
/// of edit: the writer drops the run's existing fill and injects the new one
/// in schema position (after any `a:ln`).
#[derive(Debug, Clone, Default)]
pub struct RunFormat {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    /// Size in POINTS. The file stores hundredths, which this converts.
    pub font_size: Option<f32>,
    /// Solid fill colour as a 6-hex string ("RRGGBB"), no leading `#`.
    pub color: Option<String>,
}

/// Where a shape sits and how big it is, in POINTS. The file counts EMU
/// (1pt = 12700 EMU), which the writer converts. Only top-level shapes should
/// be moved this way -- a group child's coordinates are in its group's space.
#[derive(Debug, Clone, Copy)]
pub struct ShapeGeometry {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
}

/// Round-trip pptx editor.
pub struct PptxEditor {
    original_data: Vec<u8>,
    presentation: Presentation,
    /// (slide_idx) -> { (shape, para, run) -> text }
    edits: HashMap<usize, HashMap<(usize, usize, usize), String>>,
    /// (slide_idx) -> { (shape, para) -> character offset to break at }
    splits: HashMap<usize, HashMap<(usize, usize), usize>>,
    /// (slide_idx) -> { (shape, para) } -- paragraphs joined onto the one before
    merges: HashMap<usize, std::collections::HashSet<(usize, usize)>>,
    /// (slide_idx) -> { (shape, para, run) -> what to change about its look }
    formats: HashMap<usize, HashMap<(usize, usize, usize), RunFormat>>,
    /// (slide_idx) -> { shape -> where and how big it should be }
    geoms: HashMap<usize, HashMap<usize, ShapeGeometry>>,
}

impl PptxEditor {
    pub fn new(data: &[u8]) -> Result<Self, PptxError> {
        let presentation = parse_pptx(data)?;
        Ok(Self {
            original_data: data.to_vec(),
            presentation,
            edits: HashMap::new(),
            splits: HashMap::new(),
            merges: HashMap::new(),
            formats: HashMap::new(),
            geoms: HashMap::new(),
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

    /// The runs this editor can address, as IT counts them.
    ///
    /// `set_run_text` takes a slide, a shape, a paragraph and a run, and the
    /// shape number counts the `<p:sp>` and `<p:pic>` children of the slide's
    /// shape tree — not the shapes the IR lists, which is a different set. A
    /// caller working from a `Presentation` therefore cannot tell which
    /// numbers to pass, and guessing puts the edit on the wrong shape without
    /// any error to show for it.
    ///
    /// One entry per slide, holding its shapes, holding their paragraphs,
    /// holding each run's current text.
    /// Break `paragraph_index` in two at `at_char`.
    ///
    /// One break per paragraph per save: a second call replaces the first,
    /// because the offsets of a paragraph that has already been cut are no
    /// longer the ones the caller measured.
    pub fn split_paragraph(
        &mut self,
        slide_index: usize,
        shape_index: usize,
        paragraph_index: usize,
        at_char: usize,
    ) {
        self.splits
            .entry(slide_index)
            .or_default()
            .insert((shape_index, paragraph_index), at_char);
    }

    /// Join `paragraph_index` onto the paragraph before it -- what Backspace at
    /// the start of a paragraph asks for.
    ///
    /// The joined paragraph keeps the FIRST one's properties: a line pulled up
    /// into a bulleted paragraph joins that bullet, it does not bring its own.
    /// Nothing happens for paragraph 0, which has nothing to join.
    pub fn merge_paragraph(
        &mut self,
        slide_index: usize,
        shape_index: usize,
        paragraph_index: usize,
    ) {
        if paragraph_index == 0 {
            return;
        }
        self.merges
            .entry(slide_index)
            .or_default()
            .insert((shape_index, paragraph_index));
    }

    /// Change how one run looks. Properties left `None` keep what they had.
    pub fn set_run_format(
        &mut self,
        slide_index: usize,
        shape_index: usize,
        paragraph_index: usize,
        run_index: usize,
        format: RunFormat,
    ) {
        self.formats
            .entry(slide_index)
            .or_default()
            .insert((shape_index, paragraph_index, run_index), format);
    }

    /// Move and size one top-level shape. `shape_index` is the editor's own
    /// `sp`/`pic` count (the IR's `sp_index`).
    pub fn set_shape_geometry(&mut self, slide_index: usize, shape_index: usize, geom: ShapeGeometry) {
        self.geoms
            .entry(slide_index)
            .or_default()
            .insert(shape_index, geom);
    }

    pub fn addressable_runs(&self) -> Result<Vec<Vec<Vec<Vec<String>>>>, PptxError> {
        let paths = self.resolve_slide_paths()?;
        let mut archive = OoxmlArchive::new(&self.original_data)?;
        let mut out = Vec::new();
        for path in &paths {
            let xml = archive.read_part(path)?;
            out.push(runs_of_slide(&xml));
        }
        Ok(out)
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
        if !self.splits.is_empty()
            || !self.merges.is_empty()
            || !self.formats.is_empty()
            || !self.geoms.is_empty()
        {
            return true;
        }
        !self.edits.is_empty()
    }

    pub fn save(&self) -> Result<Vec<u8>, PptxError> {
        if self.edits.is_empty()
            && self.splits.is_empty()
            && self.merges.is_empty()
            && self.formats.is_empty()
            && self.geoms.is_empty()
        {
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
        let mut path_splits: HashMap<String, &HashMap<(usize, usize), usize>> = HashMap::new();
        for (si, splits) in &self.splits {
            if let Some(path) = slide_paths.get(*si) {
                path_splits.insert(path.clone(), splits);
            }
        }
        let mut path_merges: HashMap<String, &std::collections::HashSet<(usize, usize)>> =
            HashMap::new();
        for (si, merges) in &self.merges {
            if let Some(path) = slide_paths.get(*si) {
                path_merges.insert(path.clone(), merges);
            }
        }
        let mut path_formats: HashMap<String, &HashMap<(usize, usize, usize), RunFormat>> =
            HashMap::new();
        for (si, formats) in &self.formats {
            if let Some(path) = slide_paths.get(*si) {
                path_formats.insert(path.clone(), formats);
            }
        }
        let mut path_geoms: HashMap<String, &HashMap<usize, ShapeGeometry>> = HashMap::new();
        for (si, geoms) in &self.geoms {
            if let Some(path) = slide_paths.get(*si) {
                path_geoms.insert(path.clone(), geoms);
            }
        }
        let no_formats: HashMap<(usize, usize, usize), RunFormat> = HashMap::new();
        let no_edits: HashMap<(usize, usize, usize), String> = HashMap::new();
        let no_splits: HashMap<(usize, usize), usize> = HashMap::new();
        let no_merges: std::collections::HashSet<(usize, usize)> =
            std::collections::HashSet::new();
        let no_geoms: HashMap<usize, ShapeGeometry> = HashMap::new();

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

                if path_edits.contains_key(&name)
                    || path_splits.contains_key(&name)
                    || path_merges.contains_key(&name)
                    || path_formats.contains_key(&name)
                    || path_geoms.contains_key(&name)
                {
                    let slide_edits = path_edits.get(&name).copied().unwrap_or(&no_edits);
                    let slide_splits = path_splits.get(&name).copied().unwrap_or(&no_splits);
                    let slide_merges = path_merges.get(&name).copied().unwrap_or(&no_merges);
                    let slide_formats = path_formats.get(&name).copied().unwrap_or(&no_formats);
                    let slide_geoms = path_geoms.get(&name).copied().unwrap_or(&no_geoms);
                    let mut xml = String::new();
                    entry
                        .read_to_string(&mut xml)
                        .map_err(|e| PptxError::InvalidData(e.to_string()))?;
                    let patched = patch_slide_xml(
                        &xml, slide_edits, slide_splits, slide_merges, slide_formats, slide_geoms,
                    )?;
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
                let path = oxidocs_common::security::sanitize_rel_target("ppt", target)
                    .unwrap_or_default();
                paths.push(path);
            } else {
                paths.push(String::new());
            }
        }

        Ok(paths)
    }
}

/// Patch slide XML, replacing `<a:t>` text nodes at (shape, para, run) coordinates.
/// One slide's runs, counted exactly as `patch_slide_xml` counts them:
/// shapes are the `<p:sp>` and `<p:pic>` children of `<p:spTree>`, paragraphs
/// are the `<a:p>` inside a shape, runs the `<a:r>` inside a paragraph.
///
/// Kept beside the patcher so the two cannot drift apart: a reader that counts
/// differently from the writer is worse than no reader, because every edit it
/// aims lands one shape over.
fn runs_of_slide(xml: &str) -> Vec<Vec<Vec<String>>> {
    let mut reader = Reader::from_str(xml);
    let mut shapes: Vec<Vec<Vec<String>>> = Vec::new();
    let mut paragraphs: Vec<Vec<String>> = Vec::new();
    let mut runs: Vec<String> = Vec::new();
    let mut held = String::new();
    let (mut in_tree, mut in_shape, mut in_para, mut in_run, mut in_text) =
        (false, false, false, false, false);
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => match local_name(e.name().as_ref()).as_str() {
                "spTree" => in_tree = true,
                "sp" | "pic" if in_tree && !in_shape => {
                    in_shape = true;
                    paragraphs = Vec::new();
                }
                "p" if in_shape => {
                    in_para = true;
                    runs = Vec::new();
                }
                "r" if in_para => {
                    in_run = true;
                    held.clear();
                }
                "t" if in_run => in_text = true,
                _ => {}
            },
            Ok(Event::Text(e)) if in_text => {
                if let Ok(text) = e.unescape() {
                    held.push_str(&text);
                }
            }
            Ok(Event::End(e)) => match local_name(e.name().as_ref()).as_str() {
                "spTree" => in_tree = false,
                "sp" | "pic" if in_shape => {
                    in_shape = false;
                    shapes.push(std::mem::take(&mut paragraphs));
                }
                "p" if in_para => {
                    in_para = false;
                    paragraphs.push(std::mem::take(&mut runs));
                }
                "r" if in_run => {
                    in_run = false;
                    runs.push(std::mem::take(&mut held));
                }
                "t" => in_text = false,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    shapes
}

/// Write out one buffered paragraph as TWO, cut at `at_char`.
///
/// The buffer holds every event of the original `<a:p>`, text edits already
/// applied. The cut is counted in characters over the paragraph's `<a:t>`
/// contents, so it is the offset the layout and the caret use.
///
/// Both halves keep the paragraph's own `a:pPr`: a break in the middle of a
/// bulleted, indented paragraph must not leave the tail unbulleted and flush
/// left. The run that straddles the cut keeps its `a:rPr` on both sides for the
/// same reason -- its size and weight belong to the text, not to the paragraph.
fn write_split_paragraph<W: std::io::Write>(
    writer: &mut Writer<W>,
    buffered: &[Event<'static>],
    at_char: usize,
) -> Result<(), PptxError> {
    let mut seen = 0usize;
    let mut done = false;
    let mut p_start: Option<BytesStart<'static>> = None;
    let mut ppr: Vec<Event<'static>> = Vec::new();
    let mut in_ppr = false;
    let mut cur_r: Option<BytesStart<'static>> = None;
    let mut cur_t: Option<BytesStart<'static>> = None;
    // The CURRENT run's own properties, kept the same way the paragraph's are:
    // the half after the break re-opens the run, and a run re-opened without
    // its `a:rPr` loses the size and weight of the text it carries.
    let mut rpr: Vec<Event<'static>> = Vec::new();
    let mut in_rpr = false;
    let mut in_text = false;

    let put = |w: &mut Writer<W>, ev: Event| -> Result<(), PptxError> {
        w.write_event(ev)
            .map_err(|e| PptxError::InvalidData(e.to_string()))
    };

    for ev in buffered {
        match ev {
            Event::Start(e) => {
                let name = local_name(e.name().as_ref());
                match name.as_str() {
                    "p" if p_start.is_none() => p_start = Some(e.clone()),
                    "pPr" => in_ppr = true,
                    "r" => {
                        cur_r = Some(e.clone());
                        rpr.clear();
                    }
                    "rPr" => in_rpr = true,
                    "t" => {
                        cur_t = Some(e.clone());
                        in_text = true;
                    }
                    _ => {}
                }
                if in_ppr {
                    ppr.push(ev.clone());
                }
                if in_rpr {
                    rpr.push(ev.clone());
                }
            }
            Event::End(e) => {
                let name = local_name(e.name().as_ref());
                if in_ppr {
                    ppr.push(ev.clone());
                }
                if in_rpr {
                    rpr.push(ev.clone());
                }
                if name == "pPr" {
                    in_ppr = false;
                }
                if name == "rPr" {
                    in_rpr = false;
                }
                if name == "t" {
                    in_text = false;
                }
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == "rPr" => {
                // `<a:rPr .../>` is the common shape and never opens, so it is
                // caught here rather than by the Start arm.
                rpr.clear();
                rpr.push(ev.clone());
                if in_ppr {
                    ppr.push(ev.clone());
                }
            }
            _ if in_ppr => ppr.push(ev.clone()),
            _ => {}
        }
        // The paragraph properties are COPIED for the second half and written
        // for the first, so collecting them must not stop them being written --
        // gating the writer on `in_ppr` dropped the opening tag and let the
        // closing one through, which produced `<a:p></a:pPr>`.
        {
            match ev {
                Event::Text(t) if in_text && !done => {
                    let text = t.unescape().unwrap_or_default().to_string();
                    let n = text.chars().count();
                    if seen + n >= at_char {
                        let cut = at_char.saturating_sub(seen);
                        let head: String = text.chars().take(cut).collect();
                        let tail: String = text.chars().skip(cut).collect();
                        put(writer, Event::Text(BytesText::new(&head)))?;
                        // Close the run and the paragraph, then open a new one
                        // with the same properties and carry on.
                        if let Some(t0) = &cur_t {
                            put(writer, Event::End(BytesEnd::new(
                                String::from_utf8_lossy(t0.name().as_ref()).into_owned())))?;
                        }
                        if let Some(r0) = &cur_r {
                            put(writer, Event::End(BytesEnd::new(
                                String::from_utf8_lossy(r0.name().as_ref()).into_owned())))?;
                        }
                        if let Some(p0) = &p_start {
                            put(writer, Event::End(BytesEnd::new(
                                String::from_utf8_lossy(p0.name().as_ref()).into_owned())))?;
                            put(writer, Event::Start(p0.clone()))?;
                        }
                        for pe in &ppr {
                            put(writer, pe.clone())?;
                        }
                        if let Some(r0) = &cur_r {
                            put(writer, Event::Start(r0.clone()))?;
                            for re in &rpr {
                                put(writer, re.clone())?;
                            }
                        }
                        if let Some(t0) = &cur_t {
                            put(writer, Event::Start(t0.clone()))?;
                        }
                        put(writer, Event::Text(BytesText::new(&tail)))?;
                        done = true;
                    } else {
                        put(writer, Event::Text(t.clone()))?;
                    }
                    seen += n;
                }
                other => put(writer, other.clone())?,
            }
        }
    }
    if !done {
        // The cut is at or past the end: the tail is an empty paragraph, which
        // is what Enter at the end of a line asks for.
        if let Some(p0) = &p_start {
            put(writer, Event::Start(p0.clone()))?;
            for pe in &ppr {
                put(writer, pe.clone())?;
            }
            put(writer, Event::End(BytesEnd::new(
                String::from_utf8_lossy(p0.name().as_ref()).into_owned())))?;
        }
    }
    Ok(())
}

/// Rewrite one `a:rPr`'s attributes to carry `format`.
///
/// Attributes the format leaves alone are copied through, so a run that
/// already says `sz="1800"` keeps it when only the weight is being changed.
fn apply_format(tag: &BytesStart<'_>, format: &RunFormat) -> BytesStart<'static> {
    let name = String::from_utf8_lossy(tag.name().as_ref()).into_owned();
    let mut out = BytesStart::new(name);
    let touched = |k: &str| {
        matches!(
            (k, format.bold.is_some(), format.italic.is_some(),
             format.underline.is_some(), format.font_size.is_some()),
            ("b", true, _, _, _) | ("i", _, true, _, _)
                | ("u", _, _, true, _) | ("sz", _, _, _, true)
        )
    };
    for attr in tag.attributes().flatten() {
        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
        if touched(&key) {
            continue;
        }
        out.push_attribute((key.as_str(), attr.unescape_value().unwrap_or_default().as_ref()));
    }
    if let Some(b) = format.bold {
        out.push_attribute(("b", if b { "1" } else { "0" }));
    }
    if let Some(i) = format.italic {
        out.push_attribute(("i", if i { "1" } else { "0" }));
    }
    if let Some(u) = format.underline {
        // `u` names an underline STYLE, not a flag: "sng" for a single rule and
        // "none" for none at all.
        out.push_attribute(("u", if u { "sng" } else { "none" }));
    }
    if let Some(sz) = format.font_size {
        // The file counts in hundredths of a point.
        out.push_attribute(("sz", (sz * 100.0).round().to_string().as_str()));
    }
    out
}

/// Points to EMU, the unit the file stores geometry in.
fn pt_to_emu(pt: f32) -> i64 {
    (f64::from(pt) * 12700.0).round() as i64
}

/// Emit a fresh `<a:xfrm><a:off/><a:ext/></a:xfrm>` for a shape that carried
/// none. A shape without its own transform inherits it, so it has no rotation
/// or flip to preserve here.
fn emit_xfrm(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    buffer: &mut Option<(usize, Vec<Event<'static>>)>,
    g: ShapeGeometry,
) -> Result<(), PptxError> {
    emit(writer, buffer, Event::Start(BytesStart::new("a:xfrm")))?;
    let mut off = BytesStart::new("a:off");
    off.push_attribute(("x", pt_to_emu(g.x).to_string().as_str()));
    off.push_attribute(("y", pt_to_emu(g.y).to_string().as_str()));
    emit(writer, buffer, Event::Empty(off))?;
    let mut ext = BytesStart::new("a:ext");
    ext.push_attribute(("cx", pt_to_emu(g.width).max(0).to_string().as_str()));
    ext.push_attribute(("cy", pt_to_emu(g.height).max(0).to_string().as_str()));
    emit(writer, buffer, Event::Empty(ext))?;
    emit(writer, buffer, Event::End(BytesEnd::new("a:xfrm")))?;
    Ok(())
}

/// Rewrite an `a:off` / `a:ext` tag's coordinates to `g`, keeping any other
/// attributes (there are none of note, but this stays uniform with the rest).
fn geom_tag(tag: &BytesStart<'_>, g: ShapeGeometry, is_off: bool) -> BytesStart<'static> {
    let name = String::from_utf8_lossy(tag.name().as_ref()).into_owned();
    let mut out = BytesStart::new(name);
    let (ka, kb) = if is_off { ("x", "y") } else { ("cx", "cy") };
    for attr in tag.attributes().flatten() {
        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
        if key == ka || key == kb {
            continue;
        }
        out.push_attribute((key.as_str(), attr.unescape_value().unwrap_or_default().as_ref()));
    }
    if is_off {
        out.push_attribute((ka, pt_to_emu(g.x).to_string().as_str()));
        out.push_attribute((kb, pt_to_emu(g.y).to_string().as_str()));
    } else {
        out.push_attribute((ka, pt_to_emu(g.width).max(0).to_string().as_str()));
        out.push_attribute((kb, pt_to_emu(g.height).max(0).to_string().as_str()));
    }
    out
}

/// Route one event to the split buffer when a paragraph is being cut, else
/// straight to the writer -- the pattern this function repeats inline.
fn emit(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    buffer: &mut Option<(usize, Vec<Event<'static>>)>,
    ev: Event<'static>,
) -> Result<(), PptxError> {
    if let Some((_, buf)) = buffer.as_mut() {
        buf.push(ev);
    } else {
        writer
            .write_event(ev)
            .map_err(|e| PptxError::InvalidData(e.to_string()))?;
    }
    Ok(())
}

/// Emit `<a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>` with the run's
/// own namespace prefix.
fn emit_fill(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    buffer: &mut Option<(usize, Vec<Event<'static>>)>,
    prefix: &str,
    color: &str,
) -> Result<(), PptxError> {
    let q = |local: &str| {
        if prefix.is_empty() {
            local.to_string()
        } else {
            format!("{prefix}:{local}")
        }
    };
    let fill = q("solidFill");
    emit(writer, buffer, Event::Start(BytesStart::new(fill.clone())))?;
    let mut clr = BytesStart::new(q("srgbClr"));
    clr.push_attribute(("val", color));
    emit(writer, buffer, Event::Empty(clr))?;
    emit(writer, buffer, Event::End(BytesEnd::new(fill)))?;
    Ok(())
}

/// The namespace prefix of a qualified tag name (`"a"` for `a:rPr`, `""` for a
/// bare one).
fn tag_prefix(tag: &BytesStart<'_>) -> String {
    String::from_utf8_lossy(tag.name().as_ref())
        .rsplit_once(':')
        .map(|(p, _)| p.to_string())
        .unwrap_or_default()
}

/// A fresh `a:rPr` for a run that carries none, named with the run's own prefix.
fn new_rpr(run_tag: &BytesStart<'_>, format: &RunFormat) -> BytesStart<'static> {
    let run_name = String::from_utf8_lossy(run_tag.name().as_ref()).into_owned();
    let prefix = run_name.rsplit_once(':').map(|(p, _)| p).unwrap_or("");
    let name = if prefix.is_empty() { "rPr".to_string() } else { format!("{prefix}:rPr") };
    apply_format(&BytesStart::new(name), format)
}

fn patch_slide_xml(
    xml: &str,
    edits: &HashMap<(usize, usize, usize), String>,
    splits: &HashMap<(usize, usize), usize>,
    merges: &std::collections::HashSet<(usize, usize)>,
    formats: &HashMap<(usize, usize, usize), RunFormat>,
    geoms: &HashMap<usize, ShapeGeometry>,
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
    // Whether the `<a:t>` currently open carried any text. An empty one emits
    // no `Text` event, so the edit that replaces its text has to be written
    // when the element CLOSES -- which is the only chance to type into the
    // paragraph a break just opened.
    let mut text_seen = false;
    // While a paragraph is being split its events are collected here instead of
    // written, because the second half needs the first half's properties and
    // they arrive before the cut is reached.
    let mut buffer: Option<(usize, Vec<Event<'static>>)> = None;
    // A paragraph joined onto the one before it drops its own opening tag, its
    // own properties and the previous paragraph's closing tag, so its runs flow
    // into the paragraph above.
    let mut swallow_ppr = false;
    let mut ppr_depth = 0usize;
    // A run whose look is being changed: the tag is kept so a missing `a:rPr`
    // can be written with the run's own namespace prefix, and the flag says
    // whether one has been seen yet.
    let mut format_run: Option<(BytesStart<'static>, RunFormat)> = None;
    // Colour is a CHILD of `a:rPr`, so it cannot ride `apply_format`'s
    // attribute rewrite. While inside the rPr being recoloured, its existing
    // fill is dropped (`skip_fill_depth`) and the new `<a:solidFill>` is
    // emitted lazily -- before the first child that is not `a:ln`, or at the
    // rPr's close -- so it lands in schema position whether or not the run
    // carries an outline.
    let mut in_fmt_rpr = false;
    let mut pending_fill: Option<(String, String)> = None; // (prefix, "RRGGBB")
    let mut skip_fill_depth: usize = 0;
    // Nesting below the rPr's direct-child level: 0 = a direct child, >0 = inside
    // one (e.g. an `a:ln`, whose own `a:solidFill` is the OUTLINE colour and
    // must not be mistaken for the text fill).
    let mut fmt_depth: usize = 0;
    // Moving/sizing the current shape. `geom_shape` is set while inside the
    // target `p:sp`/`p:pic`; `in_geom_sppr` while inside its `p:spPr`;
    // `pending_xfrm` holds the geometry until an existing `a:xfrm` is found to
    // rewrite (keeping its rot/flip) or the first non-xfrm child forces a fresh
    // one to be injected in schema position (xfrm is spPr's first child);
    // `in_geom_xfrm` while rewriting that transform's `a:off`/`a:ext`.
    let mut geom_shape: Option<ShapeGeometry> = None;
    let mut in_geom_sppr = false;
    let mut pending_xfrm: Option<ShapeGeometry> = None;
    let mut in_geom_xfrm = false;

    loop {
        match reader.read_event().map_err(PptxError::Xml)? {
            Event::Eof => break,
            // Dropping the run's existing `<a:solidFill>` subtree while its
            // colour is being replaced. Set to 1 when that fill opens; any
            // nesting deepens it, its close returns it to 0.
            Event::Start(_) if skip_fill_depth > 0 => skip_fill_depth += 1,
            Event::End(_) if skip_fill_depth > 0 => skip_fill_depth -= 1,
            Event::Empty(_) | Event::Text(_) if skip_fill_depth > 0 => {}
            // Inside the rPr being recoloured. At the direct-child level the new
            // fill is injected (after any `a:ln`) and the old one dropped;
            // deeper down every event is copied through untouched.
            Event::Start(ref e) if in_fmt_rpr => {
                let ln = local_name(e.name().as_ref());
                if fmt_depth == 0 && ln == "solidFill" {
                    skip_fill_depth = 1;
                } else {
                    if fmt_depth == 0 && ln != "ln" {
                        if let Some((pfx, col)) = pending_fill.take() {
                            emit_fill(&mut writer, &mut buffer, &pfx, &col)?;
                        }
                    }
                    fmt_depth += 1;
                    emit(&mut writer, &mut buffer, Event::Start(e.clone().into_owned()))?;
                }
            }
            Event::Empty(ref e) if in_fmt_rpr && fmt_depth == 0 => {
                let ln = local_name(e.name().as_ref());
                if ln != "solidFill" {
                    if ln != "ln" {
                        if let Some((pfx, col)) = pending_fill.take() {
                            emit_fill(&mut writer, &mut buffer, &pfx, &col)?;
                        }
                    }
                    emit(&mut writer, &mut buffer, Event::Empty(e.clone().into_owned()))?;
                }
            }
            Event::End(ref e) if in_fmt_rpr => {
                if fmt_depth > 0 {
                    fmt_depth -= 1;
                    emit(&mut writer, &mut buffer, Event::End(e.clone().into_owned()))?;
                } else {
                    // The rPr closes: flush the fill if no child pulled it in.
                    if let Some((pfx, col)) = pending_fill.take() {
                        emit_fill(&mut writer, &mut buffer, &pfx, &col)?;
                    }
                    in_fmt_rpr = false;
                    emit(&mut writer, &mut buffer, Event::End(e.clone().into_owned()))?;
                }
            }
            Event::Text(ref e) if in_fmt_rpr => {
                emit(&mut writer, &mut buffer, Event::Text(e.clone().into_owned()))?;
            }
            // --- moving/sizing a shape: its `p:spPr` transform ---
            // The first child of the target spPr decides how the transform is
            // written: an existing `a:xfrm` is rewritten in place (its rot/flip
            // kept); anything else means the shape has none, so a fresh one is
            // injected before it -- the correct first-child position.
            Event::Start(ref e) if in_geom_sppr && pending_xfrm.is_some() => {
                if local_name(e.name().as_ref()) == "xfrm" {
                    pending_xfrm = None;
                    in_geom_xfrm = true;
                    emit(&mut writer, &mut buffer, Event::Start(e.clone().into_owned()))?;
                } else {
                    let g = pending_xfrm.take().expect("just checked");
                    emit_xfrm(&mut writer, &mut buffer, g)?;
                    emit(&mut writer, &mut buffer, Event::Start(e.clone().into_owned()))?;
                }
            }
            Event::Empty(ref e) if in_geom_sppr && pending_xfrm.is_some() => {
                let g = pending_xfrm.take().expect("just checked");
                // A self-closing `<a:xfrm/>` is replaced; any other first child
                // gets a fresh transform injected before it.
                emit_xfrm(&mut writer, &mut buffer, g)?;
                if local_name(e.name().as_ref()) != "xfrm" {
                    emit(&mut writer, &mut buffer, Event::Empty(e.clone().into_owned()))?;
                }
            }
            Event::Empty(ref e) if in_geom_xfrm && local_name(e.name().as_ref()) == "off" => {
                let g = geom_shape.expect("geom active");
                emit(&mut writer, &mut buffer, Event::Empty(geom_tag(e, g, true)))?;
            }
            Event::Empty(ref e) if in_geom_xfrm && local_name(e.name().as_ref()) == "ext" => {
                let g = geom_shape.expect("geom active");
                emit(&mut writer, &mut buffer, Event::Empty(geom_tag(e, g, false)))?;
            }
            Event::End(ref e) if in_geom_xfrm && local_name(e.name().as_ref()) == "xfrm" => {
                in_geom_xfrm = false;
                emit(&mut writer, &mut buffer, Event::End(e.clone().into_owned()))?;
            }
            Event::End(ref e) if in_geom_sppr && local_name(e.name().as_ref()) == "spPr" => {
                // An empty spPr with no children still gets its transform.
                if let Some(g) = pending_xfrm.take() {
                    emit_xfrm(&mut writer, &mut buffer, g)?;
                }
                in_geom_sppr = false;
                emit(&mut writer, &mut buffer, Event::End(e.clone().into_owned()))?;
            }
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
                        geom_shape = geoms.get(&shape_idx).copied();
                    }
                    "spPr" if geom_shape.is_some() => {
                        // The shape's own properties: its transform is written
                        // through the geom arms above. Write the opening tag and
                        // arm them.
                        in_geom_sppr = true;
                        pending_xfrm = geom_shape;
                        emit(&mut writer, &mut buffer, Event::Start(e.clone().into_owned()))?;
                        continue;
                    }
                    "p" if in_shape => {
                        in_paragraph = true;
                        run_idx = 0;
                        if let Some(at) = splits.get(&(shape_idx, para_idx)) {
                            buffer = Some((*at, Vec::new()));
                        }
                        if merges.contains(&(shape_idx, para_idx)) {
                            swallow_ppr = true;
                            continue;
                        }
                    }
                    "pPr" if swallow_ppr => {
                        ppr_depth += 1;
                        continue;
                    }
                    _ if ppr_depth > 0 => {
                        ppr_depth += 1;
                        continue;
                    }
                    "rPr" if format_run.is_some() => {
                        // The run's properties as an OPEN tag, which is what a
                        // run carrying a fill or a typeface has. Rewritten in
                        // place; the fresh-rPr path below must not also fire,
                        // or the run ends up with two and the original wins.
                        let (_, f) = format_run.take().expect("just checked");
                        let tag = apply_format(e, &f);
                        emit(&mut writer, &mut buffer, Event::Start(tag))?;
                        // A colour is a child, not an attribute: hand the rest
                        // of this rPr to the recolour arms above.
                        if let Some(col) = f.color.clone() {
                            in_fmt_rpr = true;
                            fmt_depth = 0;
                            pending_fill = Some((tag_prefix(e), col));
                        }
                        continue;
                    }
                    "r" if in_paragraph => {
                        in_run = true;
                        format_run = formats
                            .get(&(shape_idx, para_idx, run_idx))
                            .map(|f| (e.clone().into_owned(), f.clone()));
                    }
                    "t" if in_run => {
                        in_text = true;
                    }
                    _ => {}
                }
                if format_run.is_some() && name != "r" {
                    // The run carried no `a:rPr`, so one is written before
                    // whatever its first child turned out to be.
                    let (run_tag, f) = format_run.take().expect("just checked");
                    let fresh = new_rpr(&run_tag, &f);
                    match f.color.clone() {
                        // A colour needs a child, so the fresh rPr opens rather
                        // than self-closes and carries the fill.
                        Some(col) => {
                            let prefix = tag_prefix(&fresh);
                            let name = String::from_utf8_lossy(fresh.name().as_ref()).into_owned();
                            emit(&mut writer, &mut buffer, Event::Start(fresh))?;
                            emit_fill(&mut writer, &mut buffer, &prefix, &col)?;
                            emit(&mut writer, &mut buffer, Event::End(BytesEnd::new(name)))?;
                        }
                        None => emit(&mut writer, &mut buffer, Event::Empty(fresh))?,
                    }
                }
                if let Some((_, buf)) = buffer.as_mut() {
                    buf.push(Event::Start(e.clone().into_owned()));
                } else {
                    writer
                        .write_event(Event::Start(e.clone()))
                        .map_err(|e| PptxError::InvalidData(e.to_string()))?;
                }
            }
            Event::End(ref e) => {
                let name = local_name(e.name().as_ref());
                if ppr_depth > 0 {
                    ppr_depth -= 1;
                    if ppr_depth == 0 {
                        swallow_ppr = false;
                    }
                    continue;
                }
                match name.as_str() {
                    "spTree" => {
                        in_sp_tree = false;
                    }
                    "sp" | "pic" if in_shape => {
                        in_shape = false;
                        shape_idx += 1;
                        geom_shape = None;
                        in_geom_sppr = false;
                        pending_xfrm = None;
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
                        // An empty run: no `Text` event came, so the edit is
                        // written here, just before the element closes.
                        if !text_seen {
                            if let Some(new_text) = edits.get(&(shape_idx, para_idx, run_idx)) {
                                let filled = Event::Text(BytesText::new(new_text).into_owned());
                                if let Some((_, buf)) = buffer.as_mut() {
                                    buf.push(filled);
                                } else {
                                    writer
                                        .write_event(filled)
                                        .map_err(|e| PptxError::InvalidData(e.to_string()))?;
                                }
                            }
                        }
                        in_text = false;
                        text_seen = false;
                    }
                    _ => {}
                }
                // The closer of the paragraph BEFORE a joined one is dropped, so
                // the two become one. `para_idx` has already advanced past this
                // paragraph, so it names the paragraph that follows -- which is
                // exactly the one asking to be joined.
                if local_name(e.name().as_ref()) == "p"
                    && merges.contains(&(shape_idx, para_idx))
                {
                    continue;
                }
                let closing_split = local_name(e.name().as_ref()) == "p" && buffer.is_some();
                if let Some((_, buf)) = buffer.as_mut() {
                    buf.push(Event::End(e.clone().into_owned()));
                } else {
                    writer
                        .write_event(Event::End(e.clone()))
                        .map_err(|e| PptxError::InvalidData(e.to_string()))?;
                }
                if closing_split {
                    let (at, buf) = buffer.take().expect("just checked");
                    write_split_paragraph(&mut writer, &buf, at)?;
                }
            }
            Event::Text(ref e) if ppr_depth > 0 => {
                let _ = e;
            }
            Event::Text(ref e) => {
                // A text edit is applied BEFORE any split, so the split counts
                // characters of the text the file will actually carry.
                if in_text {
                    text_seen = true;
                }
                let out = if in_text {
                    match edits.get(&(shape_idx, para_idx, run_idx)) {
                        Some(new_text) => Event::Text(BytesText::new(new_text).into_owned()),
                        None => Event::Text(e.clone().into_owned()),
                    }
                } else {
                    Event::Text(e.clone().into_owned())
                };
                if let Some((_, buf)) = buffer.as_mut() {
                    buf.push(out);
                } else {
                    writer
                        .write_event(out)
                        .map_err(|e| PptxError::InvalidData(e.to_string()))?;
                }
            }
            Event::Empty(ref e)
                if format_run.is_some() && local_name(e.name().as_ref()) == "rPr" =>
            {
                let (_, f) = format_run.take().expect("just checked");
                let tag = apply_format(e, &f);
                match f.color.clone() {
                    // A self-closing rPr has to open to carry a fill child.
                    Some(col) => {
                        let prefix = tag_prefix(e);
                        let name = String::from_utf8_lossy(tag.name().as_ref()).into_owned();
                        emit(&mut writer, &mut buffer, Event::Start(tag))?;
                        emit_fill(&mut writer, &mut buffer, &prefix, &col)?;
                        emit(&mut writer, &mut buffer, Event::End(BytesEnd::new(name)))?;
                    }
                    None => emit(&mut writer, &mut buffer, Event::Empty(tag))?,
                }
            }
            Event::Empty(ref e)
                if swallow_ppr && local_name(e.name().as_ref()) == "pPr" =>
            {
                // A self-closing `<a:pPr/>` never opens, so it is dropped here.
                swallow_ppr = false;
            }
            event if ppr_depth > 0 => {
                let _ = event;
            }
            event => {
                if let Some((_, buf)) = buffer.as_mut() {
                    buf.push(event.into_owned());
                } else {
                    writer
                        .write_event(event)
                        .map_err(|e| PptxError::InvalidData(e.to_string()))?;
                }
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

    #[test]
    fn test_editor_run_colour_and_weight() {
        // The fixture's first run is `<a:rPr .../>` -- the self-closing case,
        // which the colour writer must reopen to carry a fill child.
        let data = include_bytes!("../../../tests/fixtures/basic_test.pptx");
        let mut editor = PptxEditor::new(data).expect("should open");
        editor.set_run_format(
            0,
            0,
            0,
            0,
            RunFormat {
                bold: Some(false),
                color: Some("FF0000".to_string()),
                ..Default::default()
            },
        );
        let saved = editor.save().expect("should save");
        // The saved bytes must still be a valid deck the parser reads, and the
        // run must carry the new colour AND the cleared weight together.
        let pres = parse_pptx(&saved).expect("should parse");
        if let crate::ir::ShapeContent::TextBox { paragraphs } = &pres.slides[0].shapes[0].content {
            let run = &paragraphs[0].runs[0];
            assert_eq!(run.color.as_deref(), Some("FF0000"), "colour persisted");
            assert_eq!(run.bold, Some(false), "weight cleared beside the colour");
        } else {
            panic!("Expected TextBox");
        }
    }

    #[test]
    fn test_editor_move_shape() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.pptx");
        let mut editor = PptxEditor::new(data).expect("should open");
        // The first addressable shape.
        let idx = editor.presentation().slides[0].shapes[0]
            .sp_index
            .expect("shape 0 is addressable");
        editor.set_shape_geometry(0, idx, ShapeGeometry { x: 100.0, y: 60.0, width: 300.0, height: 120.0 });
        let saved = editor.save().expect("should save");
        let pres = parse_pptx(&saved).expect("should parse");
        let sh = pres.slides[0].shapes.iter().find(|s| s.sp_index == Some(idx)).expect("shape kept");
        // Points survive the EMU round-trip to within the rounding.
        assert!((sh.x - 100.0).abs() < 0.1, "x = {}", sh.x);
        assert!((sh.y - 60.0).abs() < 0.1, "y = {}", sh.y);
        assert!((sh.width - 300.0).abs() < 0.1, "w = {}", sh.width);
        assert!((sh.height - 120.0).abs() < 0.1, "h = {}", sh.height);
    }
}

#[cfg(test)]
mod split_tests {
    use super::*;
    use std::collections::{HashMap, HashSet};

    const SLIDE: &str = concat!(
        r#"<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>"#,
        r#"<p:sp><p:txBody>"#,
        r#"<a:p><a:pPr lvl="1"><a:buChar char="-"/></a:pPr>"#,
        r#"<a:r><a:rPr sz="1800" b="1"/><a:t>Hello world</a:t></a:r>"#,
        r#"</a:p>"#,
        r#"<a:p><a:r><a:t>Second</a:t></a:r></a:p>"#,
        r#"</p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#,
    );

    fn split_at(at: usize) -> String {
        let mut splits = HashMap::new();
        splits.insert((0usize, 0usize), at);
        patch_slide_xml(SLIDE, &HashMap::new(), &splits, &HashSet::new(), &HashMap::new(), &HashMap::new()).expect("patch")
    }

    fn paragraphs(xml: &str) -> Vec<String> {
        let mut out = Vec::new();
        let mut rest = xml;
        while let Some(i) = rest.find("<a:p>") {
            let after = &rest[i..];
            let end = after.find("</a:p>").map(|e| e + 6).unwrap_or(after.len());
            out.push(after[..end].to_string());
            rest = &after[end..];
        }
        out
    }

    fn texts(p: &str) -> String {
        let mut out = String::new();
        let mut rest = p;
        while let Some(i) = rest.find("<a:t>") {
            let after = &rest[i + 5..];
            let end = after.find("</a:t>").unwrap_or(after.len());
            out.push_str(&after[..end]);
            rest = &after[end..];
        }
        out
    }

    #[test]
    fn a_cut_in_the_middle_gives_two_paragraphs_with_the_text_divided() {
        let out = split_at(5);
        let ps = paragraphs(&out);
        assert_eq!(ps.len(), 3, "two halves plus the untouched second paragraph");
        assert_eq!(texts(&ps[0]), "Hello");
        assert_eq!(texts(&ps[1]), " world");
        assert_eq!(texts(&ps[2]), "Second", "the other paragraph is untouched");
    }

    #[test]
    fn both_halves_keep_the_paragraphs_own_properties() {
        let out = split_at(5);
        let ps = paragraphs(&out);
        for (i, p) in ps.iter().take(2).enumerate() {
            assert!(p.contains(r#"lvl="1""#), "half {i} lost its level: {p}");
            assert!(p.contains("buChar"), "half {i} lost its bullet: {p}");
        }
    }

    #[test]
    fn the_run_that_is_cut_keeps_its_own_properties_on_both_sides() {
        let out = split_at(5);
        let ps = paragraphs(&out);
        for (i, p) in ps.iter().take(2).enumerate() {
            assert!(p.contains(r#"sz="1800""#), "half {i} lost the run size: {p}");
            assert!(p.contains(r#"b="1""#), "half {i} lost the run weight: {p}");
        }
    }

    #[test]
    fn a_cut_at_the_end_adds_an_empty_paragraph() {
        let out = split_at(11);
        let ps = paragraphs(&out);
        assert_eq!(texts(&ps[0]), "Hello world");
        assert_eq!(texts(&ps[1]), "", "the tail is empty");
        assert!(ps[1].contains(r#"lvl="1""#), "and still carries the properties");
    }

    #[test]
    fn a_cut_past_the_end_behaves_like_a_cut_at_the_end() {
        let out = split_at(999);
        let ps = paragraphs(&out);
        assert_eq!(texts(&ps[0]), "Hello world");
        assert_eq!(texts(&ps[1]), "");
    }

    #[test]
    fn a_cut_at_zero_leaves_the_whole_text_on_the_second_half() {
        let out = split_at(0);
        let ps = paragraphs(&out);
        assert_eq!(texts(&ps[0]), "");
        assert_eq!(texts(&ps[1]), "Hello world");
    }

    #[test]
    fn a_text_edit_is_applied_before_the_cut_is_counted() {
        let mut edits = HashMap::new();
        edits.insert((0usize, 0usize, 0usize), "Goodbye now".to_string());
        let mut splits = HashMap::new();
        splits.insert((0usize, 0usize), 7);
        let out = patch_slide_xml(SLIDE, &edits, &splits, &HashSet::new(), &HashMap::new(), &HashMap::new()).expect("patch");
        let ps = paragraphs(&out);
        assert_eq!(texts(&ps[0]), "Goodbye");
        assert_eq!(texts(&ps[1]), " now");
    }

    #[test]
    fn a_slide_with_no_split_is_left_alone() {
        let out = patch_slide_xml(SLIDE, &HashMap::new(), &HashMap::new(), &HashSet::new(), &HashMap::new(), &HashMap::new()).expect("patch");
        assert_eq!(paragraphs(&out).len(), 2);
    }
}

#[cfg(test)]
mod merge_tests {
    use super::*;
    use std::collections::{HashMap, HashSet};

    const SLIDE: &str = concat!(
        r#"<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>"#,
        r#"<p:sp><p:txBody>"#,
        r#"<a:p><a:pPr lvl="1"><a:buChar char="-"/></a:pPr>"#,
        r#"<a:r><a:t>First</a:t></a:r></a:p>"#,
        r#"<a:p><a:pPr lvl="3" algn="ctr"/><a:r><a:t>Second</a:t></a:r></a:p>"#,
        r#"<a:p><a:r><a:t>Third</a:t></a:r></a:p>"#,
        r#"</p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#,
    );

    fn merge(which: usize) -> String {
        let mut m = HashSet::new();
        m.insert((0usize, which));
        patch_slide_xml(SLIDE, &HashMap::new(), &HashMap::new(), &m, &HashMap::new(), &HashMap::new()).expect("patch")
    }

    fn paragraphs(xml: &str) -> Vec<String> {
        let mut out = Vec::new();
        let mut rest = xml;
        while let Some(i) = rest.find("<a:p>") {
            let after = &rest[i..];
            let end = after.find("</a:p>").map(|e| e + 6).unwrap_or(after.len());
            out.push(after[..end].to_string());
            rest = &after[end..];
        }
        out
    }

    fn texts(p: &str) -> String {
        let mut out = String::new();
        let mut rest = p;
        while let Some(i) = rest.find("<a:t>") {
            let after = &rest[i + 5..];
            let end = after.find("</a:t>").unwrap_or(after.len());
            out.push_str(&after[..end]);
            rest = &after[end..];
        }
        out
    }

    #[test]
    fn joining_the_second_onto_the_first_leaves_one_paragraph_with_both() {
        let ps = paragraphs(&merge(1));
        assert_eq!(ps.len(), 2, "three paragraphs became two");
        assert_eq!(texts(&ps[0]), "FirstSecond");
        assert_eq!(texts(&ps[1]), "Third");
    }

    #[test]
    fn the_joined_paragraph_keeps_the_first_ones_properties() {
        let ps = paragraphs(&merge(1));
        assert!(ps[0].contains(r#"lvl="1""#), "kept the first level: {}", ps[0]);
        assert!(ps[0].contains("buChar"), "kept the first bullet: {}", ps[0]);
        assert!(!ps[0].contains(r#"lvl="3""#), "dropped the second's: {}", ps[0]);
        assert!(!ps[0].contains("algn"), "and its alignment: {}", ps[0]);
    }

    #[test]
    fn a_paragraph_with_no_properties_joins_just_as_well() {
        let ps = paragraphs(&merge(2));
        assert_eq!(ps.len(), 2);
        assert_eq!(texts(&ps[1]), "SecondThird");
    }

    #[test]
    fn joining_the_first_paragraph_is_refused_at_the_api() {
        // `merge_paragraph` declines paragraph 0, which has nothing above it.
        let data = include_bytes!("../../../tests/fixtures/basic_test.pptx");
        if let Ok(mut ed) = PptxEditor::new(data) {
            ed.merge_paragraph(0, 0, 0);
            assert!(!ed.has_edits(), "paragraph 0 must not register a join");
        }
    }

    #[test]
    fn a_slide_with_no_join_is_left_alone() {
        let out = patch_slide_xml(SLIDE, &HashMap::new(), &HashMap::new(), &HashSet::new(), &HashMap::new(), &HashMap::new())
            .expect("patch");
        assert_eq!(paragraphs(&out).len(), 3);
    }
}

#[cfg(test)]
mod format_tests {
    use super::*;
    use std::collections::{HashMap, HashSet};

    const SLIDE: &str = concat!(
        r#"<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>"#,
        r#"<p:sp><p:txBody><a:p>"#,
        r#"<a:r><a:rPr sz="1800" lang="en"/><a:t>Styled</a:t></a:r>"#,
        r#"<a:r><a:t>Bare</a:t></a:r>"#,
        r#"</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#,
    );

    fn format(run: usize, f: RunFormat) -> String {
        let mut m = HashMap::new();
        m.insert((0usize, 0usize, run), f);
        patch_slide_xml(SLIDE, &HashMap::new(), &HashMap::new(), &HashSet::new(), &m, &HashMap::new())
            .expect("patch")
    }

    fn rpr_of(xml: &str, run: usize) -> String {
        let mut rest = xml;
        for _ in 0..run {
            let i = rest.find("<a:r>").expect("run");
            rest = &rest[i + 5..];
        }
        let i = rest.find("<a:r>").expect("run");
        let after = &rest[i..];
        let end = after.find("</a:r>").unwrap_or(after.len());
        after[..end].to_string()
    }

    #[test]
    fn making_a_run_bold_leaves_its_other_properties_alone() {
        let out = format(0, RunFormat { bold: Some(true), ..Default::default() });
        let r = rpr_of(&out, 0);
        assert!(r.contains(r#"b="1""#), "{r}");
        assert!(r.contains(r#"sz="1800""#), "the size it already had survives: {r}");
        assert!(r.contains(r#"lang="en""#), "and everything else: {r}");
    }

    #[test]
    fn clearing_bold_writes_it_off_rather_than_dropping_it() {
        // A run inside a bold paragraph needs to SAY it is not bold; removing
        // the attribute would let the level's weight back in.
        let out = format(0, RunFormat { bold: Some(false), ..Default::default() });
        assert!(rpr_of(&out, 0).contains(r#"b="0""#));
    }

    #[test]
    fn a_size_is_written_in_hundredths_of_a_point() {
        let out = format(0, RunFormat { font_size: Some(24.0), ..Default::default() });
        let r = rpr_of(&out, 0);
        assert!(r.contains(r#"sz="2400""#), "{r}");
        assert!(!r.contains(r#"sz="1800""#), "the old size is gone: {r}");
    }

    #[test]
    fn an_underline_names_a_style_not_a_flag() {
        let on = format(0, RunFormat { underline: Some(true), ..Default::default() });
        assert!(rpr_of(&on, 0).contains(r#"u="sng""#));
        let off = format(0, RunFormat { underline: Some(false), ..Default::default() });
        assert!(rpr_of(&off, 0).contains(r#"u="none""#));
    }

    #[test]
    fn a_run_with_no_properties_gets_one() {
        let out = format(1, RunFormat { bold: Some(true), ..Default::default() });
        let r = rpr_of(&out, 1);
        assert!(r.contains("<a:rPr"), "an rPr was inserted: {r}");
        assert!(r.contains(r#"b="1""#), "{r}");
        assert!(r.find("<a:rPr").unwrap() < r.find("<a:t>").unwrap(),
                "and it comes before the text: {r}");
    }

    /// ★The run whose properties are an OPEN tag -- one carrying a fill or a
    /// typeface, which is most real runs. The self-closing shape the other
    /// tests use went down a different path, and that path was the only one
    /// wired: a real deck came back with TWO `a:rPr`, the inserted one first
    /// and the original second, so the file kept its old weight.
    #[test]
    fn a_run_whose_properties_have_children_is_rewritten_in_place() {
        const OPEN: &str = concat!(
            r#"<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>"#,
            r#"<p:sp><p:txBody><a:p><a:r>"#,
            r#"<a:rPr b="0" sz="5000"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>"#,
            r#"</a:rPr><a:t>Text</a:t></a:r></a:p></p:txBody></p:sp>"#,
            r#"</p:spTree></p:cSld></p:sld>"#,
        );
        let mut m = HashMap::new();
        m.insert((0usize, 0usize, 0usize), RunFormat { bold: Some(true), ..Default::default() });
        let out = patch_slide_xml(OPEN, &HashMap::new(), &HashMap::new(), &HashSet::new(), &m, &HashMap::new())
            .expect("patch");
        assert_eq!(out.matches("<a:rPr").count(), 1, "exactly one rPr: {out}");
        assert!(out.contains(r#"b="1""#), "{out}");
        assert!(!out.contains(r#"b="0""#), "the old weight is gone: {out}");
        assert!(out.contains("srgbClr"), "and its children survive: {out}");
    }

    #[test]
    fn the_other_runs_are_untouched() {
        let out = format(0, RunFormat { bold: Some(true), ..Default::default() });
        assert!(!rpr_of(&out, 1).contains("rPr"), "run 1 stays bare");
    }

    #[test]
    fn a_slide_with_no_format_change_is_left_alone() {
        let out = patch_slide_xml(SLIDE, &HashMap::new(), &HashMap::new(),
                                  &HashSet::new(), &HashMap::new(), &HashMap::new()).expect("patch");
        assert!(out.contains(r#"<a:rPr sz="1800" lang="en"/>"#), "{out}");
    }
}
