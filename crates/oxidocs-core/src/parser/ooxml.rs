// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use std::collections::HashMap;
use std::io::{Cursor, Read};

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use zip::ZipArchive;

use super::numbering::{parse_numbering, NumberingDefinitions};
use super::relationships::{parse_relationships, Relationship};
use super::styles::parse_styles;
use super::theme::{parse_theme, ThemeColors};
use super::ParseError;
use crate::ir::{*, VerticalAlign};

pub struct OoxmlParser {
    archive: ZipArchive<Cursor<Vec<u8>>>,
}

/// Context passed through parsing functions for resource resolution
struct ParseContext {
    /// Relationship ID -> Relationship mapping
    _rels: HashMap<String, Relationship>,
    /// Relationship ID -> binary data (images, etc.)
    media: HashMap<String, Vec<u8>>,
    /// Relationship ID -> content type (e.g., "image/png")
    media_types: HashMap<String, String>,
    /// Relationship ID -> hyperlink URL (external links)
    hyperlinks: HashMap<String, String>,
    /// Numbering definitions from word/numbering.xml
    numbering: NumberingDefinitions,
    /// Counters for numbered lists: (numId, ilvl) -> current count
    list_counters: std::cell::RefCell<HashMap<(String, u8), u32>>,
    /// Footnote ID -> paragraphs (from word/footnotes.xml)
    footnotes: HashMap<String, Vec<Block>>,
    /// Endnote ID -> paragraphs (from word/endnotes.xml)
    endnotes: HashMap<String, Vec<Block>>,
    /// Comment ID -> Comment (from word/comments.xml)
    comments: HashMap<String, Comment>,
    /// Theme colors from theme1.xml
    theme: ThemeColors,
}

impl OoxmlParser {
    pub fn new(data: &[u8]) -> Result<Self, ParseError> {
        let cursor = Cursor::new(data.to_vec());
        let archive = ZipArchive::new(cursor)?;
        Ok(Self { archive })
    }

    pub fn parse(mut self) -> Result<Document, ParseError> {
        // Parse theme first — needed for font resolution in styles
        let theme = match self.read_part("word/theme/theme1.xml") {
            Ok(xml) => parse_theme(&xml),
            Err(_) => ThemeColors::default(),
        };
        let mut styles = self.parse_styles_with_theme(&theme)?;
        self.parse_font_table(&mut styles);
        let numbering = self.parse_numbering()?;
        let ctx = self.build_context_with_theme(numbering, theme, &styles)?;
        let sections = self.parse_document_xml(&ctx, &styles)?;
        let metadata = self.parse_metadata();
        let adjust_line_height_in_table = self.parse_adjust_line_height_in_table();
        let default_tab_stop = self.parse_default_tab_stop();
        let (compat_mode, compat_mode_explicit) = self.parse_compat_mode();
        let fn_special_declared = self.parse_fn_special_declared();
        let compress_punctuation = self.parse_compress_punctuation();
        let do_not_expand_shift_return = self.parse_compat_bool_flag("doNotExpandShiftReturn");
        let balance_single_byte_double_byte_width =
            self.parse_compat_bool_flag("balanceSingleByteDoubleByteWidth");
        let people = self.parse_people();

        let mut pages = Vec::new();
        let mut page_index = 0usize;
        // OOXML: sections without explicit header/footer inherit from previous section
        let mut prev_header_refs: Vec<HdrFtrRef> = Vec::new();
        let mut prev_footer_refs: Vec<HdrFtrRef> = Vec::new();
        // S755: doc-level even/odd header flag (settings.xml).
        let even_odd_hf = self.parse_even_odd_headers();
        for mut section in sections {
            let effective_header_refs = if section.properties.header_refs.is_empty() {
                &prev_header_refs
            } else {
                &section.properties.header_refs
            };
            let effective_footer_refs = if section.properties.footer_refs.is_empty() {
                &prev_footer_refs
            } else {
                &section.properties.footer_refs
            };
            // Determine which header/footer type to use.
            // S755 (2026-07-06): `header`/`footer` always carry the DEFAULT
            // type; the "first" (titlePg) and "even" (evenAndOddHeaders)
            // variants are parsed into header_first/header_even etc. and the
            // LAYOUT selects per rendered page (the old code baked the
            // first-type header into the whole section → a tall first-page
            // header was applied to every page, probextitlepg +1×6).
            // OXI_S755_DISABLE restores the legacy bake-first-into-section.
            // WATERMARK: collected from any of this section's header parts.
            let mut watermark_found: Option<crate::ir::Watermark> = None;
            let hdr_type = if std::env::var("OXI_S755_DISABLE").is_ok()
                && section.properties.title_pg && page_index == 0 { "first" } else { "default" };
            let use_headers: Vec<HdrFtrRef> = effective_header_refs.iter()
                .filter(|r| r.ref_type == hdr_type)
                .cloned()
                .collect();
            let use_footers: Vec<HdrFtrRef> = effective_footer_refs.iter()
                .filter(|r| r.ref_type == hdr_type)
                .cloned()
                .collect();
            // Fall back: if no type-matched reference, try "default"; if none, choose the
            // fallback carefully. ECMA-376 §17.10.2: `type="first"` is only active when
            // titlePg is set. Without titlePg, falling back to a `first` reference produces
            // phantom headers/footers (observed on 34140). `type="even"` is gated by the
            // global `evenAndOddHeaders` flag, but legacy docs reference it without the flag
            // and relied on Oxi's prior all-refs fallback for body-area reservation — keep
            // that fallback for "even" to avoid pagination regressions.
            fn pick_fallback(refs: &[HdrFtrRef]) -> Vec<HdrFtrRef> {
                let defaults: Vec<HdrFtrRef> = refs.iter()
                    .filter(|r| r.ref_type == "default").cloned().collect();
                if !defaults.is_empty() {
                    return defaults;
                }
                // Keep "even" refs as a last-resort fallback (legacy behavior).
                // Drop "first" refs when titlePg is absent (spec §17.10.2).
                refs.iter().filter(|r| r.ref_type != "first").cloned().collect()
            }
            let header = if !use_headers.is_empty() {
                self.parse_header_footer_blocks(&use_headers, &ctx, &styles, &mut watermark_found)
            } else {
                let fallback = pick_fallback(effective_header_refs);
                if !fallback.is_empty() {
                    self.parse_header_footer_blocks(&fallback, &ctx, &styles, &mut watermark_found)
                } else {
                    Vec::new()
                }
            };
            let footer = if !use_footers.is_empty() {
                self.parse_header_footer_blocks(&use_footers, &ctx, &styles, &mut watermark_found)
            } else {
                let fallback = pick_fallback(effective_footer_refs);
                if !fallback.is_empty() {
                    self.parse_header_footer_blocks(&fallback, &ctx, &styles, &mut watermark_found)
                } else {
                    Vec::new()
                }
            };
            // S755: parse the first/even variants (no fallback — an absent
            // variant means the layout falls back to the default blocks).
            let hf_variant = |refs: &[HdrFtrRef], t: &str| -> Vec<HdrFtrRef> {
                refs.iter().filter(|r| r.ref_type == t).cloned().collect()
            };
            let fh = hf_variant(effective_header_refs, "first");
            let header_first = if section.properties.title_pg && !fh.is_empty() {
                self.parse_header_footer_blocks(&fh, &ctx, &styles, &mut watermark_found)
            } else { Vec::new() };
            let ff = hf_variant(effective_footer_refs, "first");
            let footer_first = if section.properties.title_pg && !ff.is_empty() {
                self.parse_header_footer_blocks(&ff, &ctx, &styles, &mut watermark_found)
            } else { Vec::new() };
            let eh = hf_variant(effective_header_refs, "even");
            let header_even = if even_odd_hf && !eh.is_empty() {
                self.parse_header_footer_blocks(&eh, &ctx, &styles, &mut watermark_found)
            } else { Vec::new() };
            let ef = hf_variant(effective_footer_refs, "even");
            let footer_even = if even_odd_hf && !ef.is_empty() {
                self.parse_header_footer_blocks(&ef, &ctx, &styles, &mut watermark_found)
            } else { Vec::new() };

            // Collect referenced footnotes and endnotes for this section
            let mut footnotes_list = Vec::new();
            let mut endnotes_list = Vec::new();
            collect_note_refs(&section.blocks, &ctx, &mut footnotes_list, &mut endnotes_list);
            footnotes_list.sort_by_key(|f| f.number);
            endnotes_list.sort_by_key(|f| f.number);

            // Round 29 (2026-04-08): renumber footnote/endnote references
            // sequentially per section. The OOXML <w:footnoteReference w:id="N"/>
            // values are NOT necessarily 1,2,3... — they can start from 2 (id=1
            // is reserved for the separator) and have arbitrary gaps. Word
            // displays them as 1,2,3..., so we need to remap. The parser had
            // set Run.text = "[id]" which is wrong; rewrite it here to "[seq]".
            let mut fn_id_to_seq: std::collections::HashMap<u32, u32> = std::collections::HashMap::new();
            for (i, f) in footnotes_list.iter().enumerate() {
                fn_id_to_seq.insert(f.number, (i as u32) + 1);
            }
            let mut en_id_to_seq: std::collections::HashMap<u32, u32> = std::collections::HashMap::new();
            for (i, f) in endnotes_list.iter().enumerate() {
                en_id_to_seq.insert(f.number, (i as u32) + 1);
            }
            renumber_note_refs(&mut section.blocks, &fn_id_to_seq, &en_id_to_seq);

            // S833: append the SPECIAL footnote paragraphs (separator /
            // continuationNotice, kept under sentinel keys by parse_notes_xml)
            // with sentinel numbers AFTER the real notes — positions/seq of
            // real notes are unchanged; the layout's declared-separator
            // reservation looks them up by the sentinel numbers.
            if !footnotes_list.is_empty() {
                if let Some(blocks) = ctx.footnotes.get("__sep__") {
                    footnotes_list.push(Footnote { number: u32::MAX, blocks: blocks.clone() });
                }
                if let Some(blocks) = ctx.footnotes.get("__notice__") {
                    footnotes_list.push(Footnote { number: u32::MAX - 1, blocks: blocks.clone() });
                }
            }

            // Continuous section: merge into previous page instead of creating a new one
            if section.properties.section_type.as_deref() == Some("continuous") && !pages.is_empty() {
                let last: &mut Page = pages.last_mut().unwrap();
                // S394: also update LRPB count when extending existing section
                fn count_lrpbs_in_blocks_local(blocks: &[crate::ir::Block]) -> usize {
                    use crate::ir::Block;
                    let mut n = 0;
                    for b in blocks {
                        match b {
                            Block::Paragraph(p) => {
                                n += p.runs.iter()
                                    .filter(|r| r.has_last_rendered_page_break)
                                    .count();
                            }
                            Block::Table(t) => {
                                for row in &t.rows {
                                    for cell in &row.cells {
                                        n += count_lrpbs_in_blocks_local(&cell.blocks);
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                    n
                }
                last.total_lrpb_count += count_lrpbs_in_blocks_local(&section.blocks);
                // S560: record this continuous section's column layout at the
                // block index where its blocks begin, so layout can switch
                // column geometry per section. MUST capture the offset BEFORE
                // extending. The old `last.columns = ...` overwrite collapsed
                // the whole merged page to the LAST section's column count,
                // mis-laying-out earlier sections (kyotei36spec: a 1-col form
                // table rendered in the trailing 2-col 記載心得 context).
                last.column_runs.push((last.blocks.len(), section.properties.columns.clone()));
                // S729: parallel per-section margin run (left, right).
                last.margin_runs.push((last.blocks.len(),
                    section.properties.margin.left, section.properties.margin.right));
                last.vertical_runs.push((last.blocks.len(),
                    section.properties.margin.top, section.properties.margin.bottom,
                    section.properties.header_distance, section.properties.footer_distance));
                // S735: parallel per-section grid pitch run.
                last.grid_runs.push((last.blocks.len(), section.properties.grid_line_pitch));
                // S730: the paragraph that ENDED the previous section (it
                // carries the in-body sectPr and is the last block merged so
                // far) is a CONTINUOUS section-break mark — Word renders it
                // at zero height when empty. Mark it for the layout skip.
                if let Some(crate::ir::Block::Paragraph(bp)) = last.blocks.last_mut() {
                    bp.style.continuous_section_break = true;
                }
                last.blocks.extend(section.blocks);
                last.floating_images.extend(section.floating_images);
                last.text_boxes.extend(section.text_boxes);
                last.shapes.extend(section.shapes);
                last.footnotes.extend(footnotes_list);
                last.endnotes.extend(endnotes_list);
            } else {
                // S394 (2026-05-27): count total LRPBs in section blocks
                // (body + nested tables). Used by S391 per-line LRPB gate.
                fn count_lrpbs_in_blocks(blocks: &[crate::ir::Block]) -> usize {
                    use crate::ir::Block;
                    let mut n = 0;
                    for b in blocks {
                        match b {
                            Block::Paragraph(p) => {
                                n += p.runs.iter()
                                    .filter(|r| r.has_last_rendered_page_break)
                                    .count();
                            }
                            Block::Table(t) => {
                                for row in &t.rows {
                                    for cell in &row.cells {
                                        n += count_lrpbs_in_blocks(&cell.blocks);
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                    n
                }
                let total_lrpb = count_lrpbs_in_blocks(&section.blocks);
                // S560: seed the per-section column-run list with this (first)
                // section's column layout at block 0. Continuous sections merged
                // later append their own (block_start, columns) entries.
                let column_runs = vec![(0usize, section.properties.columns.clone())];
                // S729: seed the margin-run list with the first section's margins.
                let margin_runs = vec![(0usize,
                    section.properties.margin.left, section.properties.margin.right)];
                let vertical_runs = vec![(0usize,
                    section.properties.margin.top, section.properties.margin.bottom,
                    section.properties.header_distance, section.properties.footer_distance)];
                // S735: seed the grid-run list with the first section's pitch.
                let grid_runs = vec![(0usize, section.properties.grid_line_pitch)];
                pages.push(Page {
                    blocks: section.blocks,
                    size: section.properties.page_size,
                    margin: section.properties.margin,
                    grid_line_pitch: section.properties.grid_line_pitch,
                    grid_char_pitch: section.properties.grid_char_pitch,
                    grid_char_space_raw: section.properties.grid_char_space_raw,
                    grid_char_cw_ratio: None,
                    doc_grid_no_type: section.properties.doc_grid_no_type,
                    doc_grid_lines_and_chars: section.properties.doc_grid_lines_and_chars,
                    header,
                    footer,
                    watermark: watermark_found.clone(),
                    header_first,
                    footer_first,
                    header_even,
                    footer_even,
                    title_pg: section.properties.title_pg,
                    even_odd_hf,
                    footnotes: footnotes_list,
                    endnotes: endnotes_list,
                    floating_images: section.floating_images,
                    text_boxes: section.text_boxes,
                    shapes: section.shapes,
                    columns: section.properties.columns,
                    column_runs,
                    margin_runs,
                    vertical_runs,
                    grid_runs,
                    section_start_type: section.properties.section_type.clone(),
                    header_distance: section.properties.header_distance,
                    footer_distance: section.properties.footer_distance,
                    page_number_format: section.properties.page_number_format,
                    page_number_start: section.properties.page_number_start,
                    page_borders: section.properties.page_borders,
                    total_lrpb_count: total_lrpb,
                    bidi_columns: section.properties.bidi,
                    vertical_section: section.properties.text_direction.as_deref()
                        .map(|d| d == "tbRl" || d == "tbRlV")
                        .unwrap_or(false),
                });
            }
            // Update previous refs for inheritance
            if !section.properties.header_refs.is_empty() {
                prev_header_refs = section.properties.header_refs;
            }
            if !section.properties.footer_refs.is_empty() {
                prev_footer_refs = section.properties.footer_refs;
            }
            page_index += 1;
        }

        // Post-process: fix charGrid pitch using actual default font size.
        // parse_section_properties uses 10.5pt placeholder; correct with the
        // document's actual default font size.
        //
        // S340 (2026-05-27): rPrDefault is PRIORITIZED over Normal style.
        // Per S339 minimal repro (`tools/metrics/_s339_repro_default_fs.py`):
        // built 2 docx variants with identical docGrid (linesAndChars,
        // charSpace=-2714) + DIFFERENT Normal style sz (24 vs 21). Word's
        // per-char Information(5) measurements were IDENTICAL (avg 8.325pt
        // both). Word uses rPrDefault as the docGrid reference, IGNORING
        // Normal style overrides. Prior priority (Normal first) was wrong:
        // for b35123 (Normal sz=24=12pt, rPrDefault sz=21=10.5pt) it produced
        // pitch=11.337 / default_fs=12.0 when Word actually computes the
        // grid with default_fs=10.5. The negative-branch formula
        // `cw = font_size + char_space_pt` is default_fs-independent so the
        // cw output is unchanged for b35123; this fix preserves the
        // SectionProperties invariants required by future L2 (per-run fs
        // inheritance) and L3 (formula recalibration) work.
        let default_font_size = styles.doc_default_run_style.as_ref()
            .and_then(|rs| rs.font_size)
            .or_else(|| {
                styles.styles.get("Normal")
                    .or_else(|| styles.styles.get("a"))
                    .and_then(|s| s.paragraph.default_run_style.as_ref())
                    .and_then(|rs| rs.font_size)
            })
            .unwrap_or(10.5);
        for page in &mut pages {
            if let Some(pitch) = page.grid_char_pitch {
                // Recalculate: the pitch was computed with 10.5pt; redo with actual default size
                let content_w = page.size.width - page.margin.left - page.margin.right;
                if content_w > 0.0 {
                    // Reverse-engineer the charSpace from the stored pitch
                    // pitch = contentW / floor(contentW / raw_pitch)
                    // We need to recompute with correct default_font_size
                    // Original raw_pitch = 10.5 + charSpace_pt; we need default_font_size + charSpace_pt
                    // charSpace_pt = original_raw_pitch - 10.5 = contentW/charsLine_old - 10.5... complex
                    // Simpler: just recalculate from scratch using the page margins
                    // Since we can't easily get charSpace here, use a different approach:
                    // If the pitch was based on 10.5 but should be based on default_font_size,
                    // and both use the same formula, we can just redo it.
                    // But we lost the charSpace value. Store it in SectionProperties.
                    // For now: if default != 10.5, redo with approximate raw_pitch
                    if (default_font_size - 10.5).abs() > 0.01 {
                        // 2026-04-18: Use saved charSpace to recompute correctly.
                        // ECMA-376 §17.6.5: raw_pitch = default_font_size + charSpace/4096
                        let char_space_pt = page.grid_char_space_raw
                            .map(|cs| cs as f32 / 4096.0)
                            .unwrap_or(0.0);
                        let raw_pitch = default_font_size + char_space_pt;
                        if raw_pitch > 0.0 {
                            // S466: un-stretched raw_pitch (see section-parse comment).
                            if std::env::var("OXI_S466_DISABLE").is_err() {
                                page.grid_char_pitch = Some(raw_pitch);
                            } else {
                                let chars_line = (content_w / raw_pitch).floor().max(1.0);
                                page.grid_char_pitch = Some(content_w / chars_line);
                            }
                        }
                        let _ = pitch;
                    }
                }
                // 2026-04-19: Compute per-fontsize char-width ratio.
                // Word formula: cw(fs) = fs * pitch / default_font_size.
                // Applies for all default_font_size.
                //
                // 2026-04-19 historical claim "COM-verified b35 (default=12)"
                // was retracted by S339 (2026-05-27): minimal repro showed
                // Word uses rPrDefault sz=21 (10.5pt) for b35123's docGrid
                // reference, NOT Normal style sz=24 (12pt). See S340 fix above.
                if let Some(p) = page.grid_char_pitch {
                    if default_font_size > 0.0 {
                        page.grid_char_cw_ratio = Some(p / default_font_size);
                    }
                }
                // S390 (2026-05-27): temporary instrumentation — print
                // per-section grid_char_pitch / default_font_size /
                // grid_char_space_raw. Env-gated.
                if std::env::var("OXI_S390_DUMP_PITCH").is_ok() {
                    eprintln!("[S390] default_fs={} grid_char_pitch={:?} grid_char_space_raw={:?} cw_ratio={:?} content_w={}",
                        default_font_size, page.grid_char_pitch, page.grid_char_space_raw,
                        page.grid_char_cw_ratio,
                        page.size.width - page.margin.left - page.margin.right);
                }
            }
        }

        // Collect all comments referenced in the document
        let all_comments: Vec<Comment> = ctx.comments.values().cloned().collect();

        // Build the author palette: people.xml first (Word writes reviewer-
        // first-seen order), then any author surfaced by comments or tracked
        // changes. Color index = position in this Vec.
        let authors = build_author_palette(&people, &all_comments, &pages);

        Ok(Document {
            pages,
            styles,
            metadata,
            comments: all_comments,
            people,
            authors,
            adjust_line_height_in_table,
            default_tab_stop,
            compat_mode,
            compat_mode_explicit,
            fn_special_declared,
            compress_punctuation,
            do_not_expand_shift_return,
            balance_single_byte_double_byte_width,
        })
    }

    /// S833: settings.xml `<w:footnotePr>` with at least one `<w:footnote>`
    /// declaration — Word switches to the CUSTOM special-footnote model
    /// (separator/continuation paragraphs at their full styled heights).
    fn parse_fn_special_declared(&mut self) -> bool {
        let xml = match self.read_part("word/settings.xml") {
            Ok(x) => x,
            Err(_) => return false,
        };
        if let Some(i) = xml.find("<w:footnotePr>") {
            if let Some(j) = xml[i..].find("</w:footnotePr>") {
                return xml[i..i + j].contains("<w:footnote ");
            }
        }
        false
    }

    /// Parse `word/people.xml` (MS-DOCX w15). Missing part → empty list.
    fn parse_people(&mut self) -> Vec<Person> {
        match self.read_part("word/people.xml") {
            Ok(xml) => parse_people_xml(&xml).unwrap_or_default(),
            Err(_) => Vec::new(),
        }
    }

    /// Parse header or footer XML parts referenced by relationship IDs.
    /// `watermark_out`: a VML WordArt watermark found in any part is stored
    /// here (first one wins) — see extract_vml_watermark.
    fn parse_header_footer_blocks(
        &mut self,
        refs: &[HdrFtrRef],
        ctx: &ParseContext,
        styles: &StyleSheet,
        watermark_out: &mut Option<crate::ir::Watermark>,
    ) -> Vec<Block> {
        let mut blocks = Vec::new();
        for hdr_ref in refs {
            // Look up the relationship target path
            let target = ctx._rels.get(&hdr_ref.rel_id)
                .map(|r| r.target.clone());
            if let Some(target) = target {
                let part_path = if target.starts_with('/') {
                    target[1..].to_string()
                } else {
                    format!("word/{}", target)
                };
                if let Ok(xml) = self.read_part(&part_path) {
                    if watermark_out.is_none() && std::env::var("OXI_WATERMARK_DISABLE").is_err() {
                        *watermark_out = extract_vml_watermark(&xml);
                    }
                    if let Ok(parsed) = parse_header_footer_xml(&xml, ctx, styles) {
                        blocks.extend(parsed);
                    }
                }
            }
        }
        blocks
    }

    fn read_part(&mut self, name: &str) -> Result<String, ParseError> {
        let mut file = self
            .archive
            .by_name(name)
            .map_err(|_| ParseError::MissingPart(name.to_string()))?;
        let mut contents = String::new();
        file.read_to_string(&mut contents)?;
        Ok(contents)
    }

    fn read_binary_part(&mut self, name: &str) -> Result<Vec<u8>, ParseError> {
        let mut file = self
            .archive
            .by_name(name)
            .map_err(|_| ParseError::MissingPart(name.to_string()))?;
        let mut contents = Vec::new();
        file.read_to_end(&mut contents)?;
        Ok(contents)
    }

    fn parse_numbering(&mut self) -> Result<NumberingDefinitions, ParseError> {
        match self.read_part("word/numbering.xml") {
            Ok(xml) => parse_numbering(&xml),
            Err(ParseError::MissingPart(_)) => Ok(NumberingDefinitions::default()),
            Err(e) => Err(e),
        }
    }

    fn build_context_with_theme(&mut self, numbering: NumberingDefinitions, theme: ThemeColors, styles: &StyleSheet) -> Result<ParseContext, ParseError> {
        // Parse relationships
        let rels = match self.read_part("word/_rels/document.xml.rels") {
            Ok(xml) => parse_relationships(&xml)?,
            Err(ParseError::MissingPart(_)) => HashMap::new(),
            Err(e) => return Err(e),
        };

        // Pre-load media and hyperlinks from relationships
        let mut media = HashMap::new();
        let mut media_types = HashMap::new();
        let mut hyperlinks = HashMap::new();
        let image_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        let hyperlink_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        for (id, rel) in &rels {
            if rel.rel_type == image_rel_type {
                // Skip external targets (e.g., file:/// URLs)
                if rel.target.starts_with("file:") || rel.target.starts_with("http:") || rel.target.starts_with("https:") {
                    continue;
                }
                // Validate relationship target path against traversal attacks
                let path = match oxidocs_common::security::sanitize_rel_target("word", &rel.target) {
                    Ok(p) => p,
                    Err(_e) => {
                        continue;
                    }
                };
                if let Ok(data) = self.read_binary_part(&path) {
                    let ct = match rel.target.rsplit('.').next().map(|s| s.to_lowercase()).as_deref() {
                        Some("png") => "image/png",
                        Some("jpg") | Some("jpeg") => "image/jpeg",
                        Some("gif") => "image/gif",
                        Some("bmp") => "image/bmp",
                        Some("svg") => "image/svg+xml",
                        Some("tiff") | Some("tif") => "image/tiff",
                        Some("wmf") => "image/x-wmf",
                        Some("emf") => "image/x-emf",
                        _ => "application/octet-stream",
                    };
                    media_types.insert(id.clone(), ct.to_string());
                    media.insert(id.clone(), data);
                }
            } else if rel.rel_type == hyperlink_rel_type {
                hyperlinks.insert(id.clone(), rel.target.clone());
            }
        }

        // S759 (2026-07-09): pre-load HEADER/FOOTER part media. A header/footer
        // image (e.g. uk_health_form's Ofsted logo) uses the PART's own rels
        // (word/_rels/header1.xml.rels), NOT the document rels, so its r:embed
        // resolves to empty data unless we load them. The rId namespace is
        // per-part; add only image rIds not already present (document images
        // win — a header image rId rarely collides with a document image rId).
        let header_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
        let footer_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
        let hf_targets: Vec<String> = rels.values()
            .filter(|r| r.rel_type == header_rel_type || r.rel_type == footer_rel_type)
            .map(|r| r.target.clone())
            .collect();
        for target in hf_targets {
            let part_name = target.rsplit('/').next().unwrap_or(&target);
            let rels_path = format!("word/_rels/{}.rels", part_name);
            let hrels = match self.read_part(&rels_path) {
                Ok(xml) => match parse_relationships(&xml) { Ok(r) => r, Err(_) => continue },
                Err(_) => continue,
            };
            for (hid, hrel) in &hrels {
                if hrel.rel_type != image_rel_type || media.contains_key(hid) { continue; }
                if hrel.target.starts_with("file:") || hrel.target.starts_with("http:") || hrel.target.starts_with("https:") { continue; }
                let path = match oxidocs_common::security::sanitize_rel_target("word", &hrel.target) {
                    Ok(p) => p,
                    Err(_) => continue,
                };
                if let Ok(data) = self.read_binary_part(&path) {
                    let ct = match hrel.target.rsplit('.').next().map(|s| s.to_lowercase()).as_deref() {
                        Some("png") => "image/png",
                        Some("jpg") | Some("jpeg") => "image/jpeg",
                        Some("gif") => "image/gif",
                        Some("bmp") => "image/bmp",
                        Some("svg") => "image/svg+xml",
                        Some("tiff") | Some("tif") => "image/tiff",
                        Some("wmf") => "image/x-wmf",
                        Some("emf") => "image/x-emf",
                        _ => "application/octet-stream",
                    };
                    media_types.insert(hid.clone(), ct.to_string());
                    media.insert(hid.clone(), data);
                }
            }
        }

        // Parse footnotes (Round 29: pass styles for pStyle inheritance —
        // footnote text style "a8" sets snapToGrid=0, which without
        // inheritance leaves footnote paragraphs grid-snapped to body
        // pitch and causes line wrap to be ~5 chars too narrow.)
        let footnotes = match self.read_part("word/footnotes.xml") {
            Ok(xml) => parse_notes_xml(&xml, styles)?,
            Err(ParseError::MissingPart(_)) => HashMap::new(),
            Err(e) => return Err(e),
        };

        // Parse endnotes
        let endnotes = match self.read_part("word/endnotes.xml") {
            Ok(xml) => parse_notes_xml(&xml, styles)?,
            Err(ParseError::MissingPart(_)) => HashMap::new(),
            Err(e) => return Err(e),
        };

        // Parse comments
        let mut comments = match self.read_part("word/comments.xml") {
            Ok(xml) => parse_comments_xml(&xml)?,
            Err(ParseError::MissingPart(_)) => HashMap::new(),
            Err(e) => return Err(e),
        };

        // Parse commentsExtended.xml (optional) and merge parent-pointer +
        // resolved flag onto the already-parsed Comment objects via paraId.
        if !comments.is_empty() {
            match self.read_part("word/commentsExtended.xml") {
                Ok(xml) => {
                    let ext = parse_comments_extended_xml(&xml)?;
                    for c in comments.values_mut() {
                        if let Some(pid) = c.para_id.as_deref() {
                            if let Some(info) = ext.get(pid) {
                                c.parent_para_id = info.parent_para_id.clone();
                                c.resolved = info.resolved;
                            }
                        }
                    }
                }
                Err(ParseError::MissingPart(_)) => {}
                Err(e) => return Err(e),
            }

            // Parse commentsIds.xml (Word 2019+, optional) and merge durable
            // ids onto Comments via paraId.
            match self.read_part("word/commentsIds.xml") {
                Ok(xml) => {
                    let ids = parse_comments_ids_xml(&xml)?;
                    for c in comments.values_mut() {
                        if let Some(pid) = c.para_id.as_deref() {
                            if let Some(did) = ids.get(pid) {
                                c.durable_id = Some(did.clone());
                            }
                        }
                    }
                }
                Err(ParseError::MissingPart(_)) => {}
                Err(e) => return Err(e),
            }
        }

        // Theme already parsed and passed in
        Ok(ParseContext {
            _rels: rels,
            media,
            media_types,
            hyperlinks,
            numbering,
            list_counters: std::cell::RefCell::new(HashMap::new()),
            footnotes,
            endnotes,
            comments,
            theme,
        })
    }

    fn parse_styles_with_theme(&mut self, theme: &ThemeColors) -> Result<StyleSheet, ParseError> {
        match self.read_part("word/styles.xml") {
            Ok(xml) => parse_styles(&xml, theme),
            Err(ParseError::MissingPart(_)) => Ok(StyleSheet::default()),
            Err(e) => Err(e),
        }
    }

    /// Parse word/fontTable.xml — extract font info (panose1, charset, family, pitch).
    fn parse_font_table(&mut self, styles: &mut StyleSheet) {
        let xml = match self.read_part("word/fontTable.xml") {
            Ok(x) => x,
            Err(_) => return,
        };
        let mut reader = Reader::from_str(&xml);
        let mut current_font: Option<String> = None;
        let mut current_info = crate::ir::FontInfo::default();

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    let local = local_name(e.name().as_ref());
                    if local == "font" {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "name" {
                                current_font = Some(String::from_utf8_lossy(&attr.value).to_string());
                                current_info = crate::ir::FontInfo::default();
                            }
                        }
                    }
                }
                Ok(Event::Empty(e)) => {
                    let local = local_name(e.name().as_ref());
                    match local.as_str() {
                        "panose1" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    current_info.panose1 = Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        "charset" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    current_info.charset = Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        "family" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    current_info.family = Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        "pitch" => {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "val" {
                                    current_info.pitch = Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(e)) => {
                    if local_name(e.name().as_ref()) == "font" {
                        if let Some(name) = current_font.take() {
                            styles.font_table.insert(name, std::mem::take(&mut current_info));
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
        }
    }

    fn parse_metadata(&self) -> DocumentMetadata {
        DocumentMetadata::default()
    }

    /// Parse word/settings.xml for adjustLineHeightInTable compatibility flag.
    /// COM measurement (2026-03-27): Compatibility(12) = False for ALL tested documents,
    /// regardless of whether <w:adjustLineHeightInTable/> is present in XML.
    /// 151 documents tested across compatibilityMode 14 and 15.
    /// Therefore: always return false (= table cells snap to grid like normal paragraphs).
    fn parse_adjust_line_height_in_table(&mut self) -> bool {
        let xml = match self.read_part("word/settings.xml") {
            Ok(x) => x,
            Err(_) => return false,
        };
        xml.contains("adjustLineHeightInTable")
    }

    /// S755: settings.xml `<w:evenAndOddHeaders/>` (CT_OnOff: bare/omitted
    /// val = true; val="0"/"false"/"off" = false). Even pages then use the
    /// type="even" header/footer references.
    fn parse_even_odd_headers(&mut self) -> bool {
        let xml = match self.read_part("word/settings.xml") {
            Ok(x) => x,
            Err(_) => return false,
        };
        if let Some(i) = xml.find("evenAndOddHeaders") {
            let tail = &xml[i..(i + 60).min(xml.len())];
            let tag = &tail[..tail.find('>').unwrap_or(tail.len())];
            !(tag.contains("\"0\"") || tag.contains("\"false\"") || tag.contains("\"off\""))
        } else {
            false
        }
    }

    /// Parse word/settings.xml for compatibilityMode.
    /// Returns (mode, explicit). `explicit=false` = no compatibilityMode
    /// compatSetting in settings.xml (legacy Word ≤2010 document; Word lays
    /// these out with ≤14 behavior — S545). The mode still DEFAULTS to 15 so
    /// the shipped `compat_mode >= 15` gates keep their corpus-validated
    /// behavior; legacy-sensitive consumers check the explicit flag.
    fn parse_compat_mode(&mut self) -> (u32, bool) {
        let xml = match self.read_part("word/settings.xml") {
            Ok(x) => x,
            Err(_) => return (15, false), // default to Word 2013+
        };
        let mut reader = Reader::from_str(&xml);
        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) => {
                    if local_name(e.name().as_ref()) == "compatSetting" {
                        let mut is_compat_mode = false;
                        let mut val = String::new();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let v = String::from_utf8_lossy(&attr.value).to_string();
                            if key == "name" && v == "compatibilityMode" { is_compat_mode = true; }
                            if key == "val" { val = v; }
                        }
                        if is_compat_mode {
                            return (val.parse().unwrap_or(15), true);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
        }
        (15, false)
    }

    /// Parse word/settings.xml for w:characterSpacingControl value.
    /// Returns true if value is "compressPunctuation" or "compressPunctuationAndJapaneseKana"
    /// (enables yakumono compression). False for "doNotCompress" or absent.
    fn parse_compress_punctuation(&mut self) -> bool {
        let xml = match self.read_part("word/settings.xml") {
            Ok(x) => x,
            Err(_) => return false,
        };
        let mut reader = Reader::from_str(&xml);
        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) => {
                    if local_name(e.name().as_ref()) == "characterSpacingControl" {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                return val == "compressPunctuation"
                                    || val == "compressPunctuationAndJapaneseKana";
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
        }
        false
    }

    /// Parse a boolean compat flag from word/settings.xml.
    /// Looks for <w:flagName/> or <w:flagName w:val="..."/> inside <w:compat>.
    /// Present with no val or val="1"/"true" → true; val="0"/"false" → false; absent → false.
    fn parse_compat_bool_flag(&mut self, flag_name: &str) -> bool {
        let xml = match self.read_part("word/settings.xml") {
            Ok(x) => x,
            Err(_) => return false,
        };
        let mut reader = Reader::from_str(&xml);
        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) => {
                    if local_name(e.name().as_ref()) == flag_name {
                        // Present → default true unless val="0"
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                return val.as_ref() != "0" && val.as_ref() != "false";
                            }
                        }
                        return true; // self-closing with no val = enabled
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
        }
        false
    }

    /// Parse word/settings.xml for defaultTabStop value.
    fn parse_default_tab_stop(&mut self) -> Option<f32> {
        let xml = match self.read_part("word/settings.xml") {
            Ok(x) => x,
            Err(_) => return None,
        };
        // Quick search for w:defaultTabStop
        let mut reader = Reader::from_str(&xml);
        let mut buf: Vec<u8> = Vec::new();
        loop {
            match reader.read_event() {
                Ok(Event::Empty(e)) => {
                    if local_name(e.name().as_ref()) == "defaultTabStop" {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                if let Ok(twips) = val.parse::<f32>() {
                                    return Some(twips / 20.0);
                                }
                            }
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
        None
    }

    fn parse_document_xml(
        &mut self,
        ctx: &ParseContext,
        styles: &StyleSheet,
    ) -> Result<Vec<ParsedSection>, ParseError> {
        let xml = self.read_part("word/document.xml")?;
        parse_body(&xml, ctx, styles)
    }
}

/// A section: blocks + properties. Multiple sections make multiple pages.
struct ParsedSection {
    blocks: Vec<Block>,
    properties: SectionProperties,
    floating_images: Vec<Image>,
    text_boxes: Vec<TextBox>,
    shapes: Vec<Shape>,
}

/// S704 (2026-06-30): compute the EFFECTIVE shading colour from w:shd val/fill/color.
/// val="clear"/"nil" → fill; "solid" → color; "pctN" → N% color blended over fill
/// (e.g. pct15 black/white = #D9D9D9); stripe/grid patterns → ~25% color (approx).
/// Returns None when the result is transparent (auto) or pure white (no visible bg).
fn effective_shading_color(val: &str, fill: &str, color: &str) -> Option<String> {
    let parse = |h: &str, default: (u8, u8, u8)| -> (u8, u8, u8) {
        let h = h.trim_start_matches('#');
        if h.eq_ignore_ascii_case("auto") || h.len() != 6 {
            return default;
        }
        match (
            u8::from_str_radix(&h[0..2], 16),
            u8::from_str_radix(&h[2..4], 16),
            u8::from_str_radix(&h[4..6], 16),
        ) {
            (Ok(r), Ok(g), Ok(b)) => (r, g, b),
            _ => default,
        }
    };
    let fill_rgb = parse(fill, (255, 255, 255)); // default white (page)
    let color_rgb = parse(color, (0, 0, 0)); // default black (pattern ink)
    let pct: f32 = match val {
        "clear" | "nil" | "" => {
            if fill.eq_ignore_ascii_case("auto") {
                return None;
            }
            let (r, g, b) = fill_rgb;
            if (r, g, b) == (255, 255, 255) {
                return None;
            }
            return Some(format!("{:02X}{:02X}{:02X}", r, g, b));
        }
        "solid" => 1.0,
        v if v.starts_with("pct") => v[3..].parse::<f32>().unwrap_or(0.0) / 100.0,
        _ => 0.25, // stripe/grid/cross patterns ≈ 25% ink
    };
    let blend = |c: u8, f: u8| -> u8 { (pct * c as f32 + (1.0 - pct) * f as f32).round() as u8 };
    let (r, g, b) = (
        blend(color_rgb.0, fill_rgb.0),
        blend(color_rgb.1, fill_rgb.1),
        blend(color_rgb.2, fill_rgb.2),
    );
    if (r, g, b) == (255, 255, 255) {
        return None;
    }
    Some(format!("{:02X}{:02X}{:02X}", r, g, b))
}

/// Parse the w:body content of document.xml into sections
fn parse_body(xml: &str, ctx: &ParseContext, styles: &StyleSheet) -> Result<Vec<ParsedSection>, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut sections: Vec<ParsedSection> = Vec::new();
    let mut current_blocks = Vec::new();
    let mut current_floating_images: Vec<Image> = Vec::new();
    let mut current_text_boxes: Vec<TextBox> = Vec::new();
    let mut current_shapes: Vec<Shape> = Vec::new();
    let mut final_sect_pr = None;
    let mut depth = 0;
    let mut in_body = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "body" => {
                        in_body = true;
                        depth = 0;
                    }
                    "p" if in_body && depth == 0 => {
                        let pr = parse_paragraph(&mut reader, ctx, styles, true)?;
                        // S525 (coverage): a paragraph whose ONLY content is display
                        // math (oMathPara, no runs/images/shapes) must NOT emit an
                        // empty paragraph line before the math — the Math block IS
                        // the line (Word renders the equation as the paragraph's
                        // content). Otherwise the math sat ~1 line (~14pt) too low.
                        let math_only = !pr.math_blocks.is_empty()
                            && pr.paragraph.runs.is_empty()
                            && pr.inline_images.is_empty()
                            && pr.floating_images.is_empty()
                            && pr.shapes.is_empty()
                            && pr.text_boxes.is_empty();
                        // S537 (2026-06-10): image-only paragraphs — the inline image
                        // IS the paragraph's line in Word (COM repro _s537_inline_line:
                        // image-only para line = extent EXACTLY, 120.00pt for a 120pt
                        // image, both default and atLeast spacing; no extra text line).
                        // Emitting the empty host Block::Paragraph before the sibling
                        // Block::Image double-counted one line (~14.5pt) per image
                        // paragraph. Body twin of the S536 cell suppression.
                        // opt-out OXI_S537_DISABLE.
                        let image_only = !pr.inline_images.is_empty()
                            && pr.paragraph.runs.iter().all(|r| r.text.is_empty())
                            && pr.math_blocks.is_empty()
                            && std::env::var("OXI_S537_DISABLE").is_err();
                        if !math_only && !image_only {
                            // S784 (2026-07-11, opt-out OXI_S784_DISABLE): a NON-EMPTY
                            // paragraph whose ¶ MARK is hidden (`<w:pPr><w:rPr>
                            // <w:vanish/>`) JOINS the following paragraph — Word hides
                            // the mark so the next paragraph continues on the SAME line
                            // and the hidden para's spacing vanishes (nyserda Section
                            // 4.01 '…performed.]In consideration…' renders merged with
                            // NO space; Oxi rendered 2 paras + after=48pt = a 61.8pt
                            // phantom gap). The merged paragraph takes the FOLLOWING
                            // para's properties (the surviving ¶ mark's — the deleted-
                            // mark revision rule). The EMPTY-vanish case stays S673v
                            // (layout skip). Corpus scan: exactly 1 such para exists
                            // (nyserda); the S673v docs' hidden marks are all EMPTY.
                            let s784_join = std::env::var("OXI_S784_DISABLE").is_err()
                                && matches!(current_blocks.last(), Some(Block::Paragraph(prev))
                                    if prev.style.ppr_rpr.as_ref().map_or(false, |r| r.vanish)
                                        && prev.runs.iter().any(|r| !r.text.is_empty()));
                            if s784_join {
                                if let Some(Block::Paragraph(prev)) = current_blocks.pop() {
                                    let mut joined = pr.paragraph;
                                    let mut runs = prev.runs;
                                    runs.extend(joined.runs);
                                    joined.runs = runs;
                                    // S784b (2026-07-11): the merged paragraph's FIRST
                                    // line is physically the HIDDEN para's first line —
                                    // its first-line indent survives (Word renders the
                                    // nyserda joined 'Section 4.01' line at x=126 =
                                    // margin + para1's firstLine=720; para2's pPr wins
                                    // for everything else).
                                    if joined.style.indent_first_line.is_none() {
                                        joined.style.indent_first_line = prev.style.indent_first_line;
                                    }
                                    current_blocks.push(Block::Paragraph(joined));
                                }
                            } else {
                                current_blocks.push(Block::Paragraph(pr.paragraph));
                            }
                        }
                        // OMML math blocks become sibling Block::Math entries.
                        for mb in pr.math_blocks {
                            current_blocks.push(Block::Math(mb));
                        }
                        // Inline images become separate blocks after the paragraph
                        current_blocks.extend(pr.inline_images);
                        // Set anchor_block_index for floating images
                        let anchor_idx = current_blocks.len().saturating_sub(1);
                        for mut img in pr.floating_images {
                            img.anchor_block_index = anchor_idx;
                            current_floating_images.push(img);
                        }
                        current_shapes.extend(pr.shapes);
                        for mut tb in pr.text_boxes {
                            tb.anchor_block_index = anchor_idx;
                            current_text_boxes.push(tb);
                        }
                        // If this paragraph contained a section break, start a new section
                        if let Some(sp) = pr.sect_pr {
                            // S794 (2026-07-12, opt-out OXI_S794_DISABLE): an explicit
                            // trailing page break IMMEDIATELY before a page-starting
                            // section boundary is redundant — Word emits ONE boundary
                            // (ukframework front matter: a para ending with
                            // <w:br type="page"/> is followed directly by the empty
                            // sectPr(nextPage) para; Word 7 front pages, Oxi rendered a
                            // phantom blank = the whole 470-para body cascaded +1).
                            // STRICT adjacency: the sectPr carrier must be EMPTY and the
                            // \x0C must be the very last content before it — intervening
                            // empty paragraphs mean Word DOES render the page (the same
                            // doc's other transition keeps 3 spacer paras on their page).
                            if std::env::var("OXI_S794_DISABLE").is_err()
                                && sp.section_type.as_deref() != Some("continuous")
                            {
                                let n = current_blocks.len();
                                let s_empty = matches!(current_blocks.last(), Some(Block::Paragraph(p))
                                    if p.runs.iter().all(|r| r.text.trim().is_empty()));
                                if s_empty && n >= 2 {
                                    if let Some(Block::Paragraph(x)) = current_blocks.get_mut(n - 2) {
                                        if let Some(run) = x.runs.iter_mut().rev().find(|r| !r.text.trim().is_empty()) {
                                            let t = run.text.trim_end();
                                            if t.ends_with('\x0C') {
                                                run.text = t[..t.len() - 1].to_string();
                                            }
                                        }
                                        // the br-only-para conversion (page_break_after)
                                        if x.runs.iter().all(|r| r.text.trim().is_empty())
                                            && x.style.page_break_after
                                        {
                                            x.style.page_break_after = false;
                                        }
                                    }
                                }
                            }
                            sections.push(ParsedSection {
                                blocks: std::mem::take(&mut current_blocks),
                                properties: sp,
                                floating_images: std::mem::take(&mut current_floating_images),
                                text_boxes: std::mem::take(&mut current_text_boxes),
                                shapes: std::mem::take(&mut current_shapes),
                            });
                        }
                    }
                    "tbl" if in_body && depth == 0 => {
                        let table = parse_table(&mut reader, ctx, styles)?;
                        current_blocks.push(Block::Table(table));
                    }
                    "sdt" if in_body && depth == 0 => {
                        // Structured Document Tag — skip sdtPr, process sdtContent
                        let mut sdt_depth = 1u32;
                        let mut in_sdt_content = false;
                        loop {
                            match reader.read_event()? {
                                Event::Start(se) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "sdtContent" {
                                        in_sdt_content = true;
                                    } else if in_sdt_content {
                                        match sl.as_str() {
                                            "p" => {
                                                let pr = parse_paragraph(&mut reader, ctx, styles, true)?;
                                                current_blocks.push(Block::Paragraph(pr.paragraph));
                                                for mb in pr.math_blocks {
                                                    current_blocks.push(Block::Math(mb));
                                                }
                                                current_blocks.extend(pr.inline_images);
                                                let anchor_idx2 = current_blocks.len().saturating_sub(1);
                                                for mut img in pr.floating_images {
                                                    img.anchor_block_index = anchor_idx2;
                                                    current_floating_images.push(img);
                                                }
                                                current_shapes.extend(pr.shapes);
                                                for mut tb in pr.text_boxes {
                                                    tb.anchor_block_index = anchor_idx2;
                                                    current_text_boxes.push(tb);
                                                }
                                                if let Some(sp) = pr.sect_pr {
                                                    sections.push(ParsedSection {
                                                        blocks: std::mem::take(&mut current_blocks),
                                                        properties: sp,
                                                        floating_images: std::mem::take(&mut current_floating_images),
                                                        text_boxes: std::mem::take(&mut current_text_boxes),
                                                        shapes: std::mem::take(&mut current_shapes),
                                                    });
                                                }
                                            }
                                            "tbl" => {
                                                let table = parse_table(&mut reader, ctx, styles)?;
                                                current_blocks.push(Block::Table(table));
                                            }
                                            _ => { sdt_depth += 1; }
                                        }
                                    } else {
                                        sdt_depth += 1;
                                    }
                                }
                                Event::End(ee) => {
                                    let sl = local_name(ee.name().as_ref());
                                    if sl == "sdtContent" {
                                        in_sdt_content = false;
                                    } else if sl == "sdt" {
                                        break;
                                    } else if sdt_depth > 0 {
                                        sdt_depth -= 1;
                                    }
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                    }
                    // mc:AlternateContent at body level (e.g., SmartArt diagrams)
                    "AlternateContent" if in_body && depth == 0 => {
                        let ac = parse_alternate_content(&mut reader, &ctx, &styles)?;
                        if let Some(drawing) = ac {
                            if let Some(image) = drawing.image {
                                if image.position.is_some() {
                                    current_floating_images.push(image);
                                } else {
                                    current_blocks.push(Block::Image(image));
                                }
                            }
                            if let Some(mut shape) = drawing.shape {
                                shape.anchor_block_index = current_blocks.len().saturating_sub(1);
                                current_shapes.push(shape);
                            }
                            if let Some(mut tb) = drawing.text_box {
                                tb.anchor_block_index = current_blocks.len().saturating_sub(1);
                                current_text_boxes.push(tb);
                            }
                        }
                    }
                    "sectPr" if in_body && depth == 0 => {
                        // Final section properties (for the last section)
                        final_sect_pr = Some(parse_section_properties(&mut reader)?);
                    }
                    _ if in_body => {
                        depth += 1;
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "body" {
                    in_body = false;
                } else if in_body && depth > 0 {
                    depth -= 1;
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    // Self-closing <w:p .../> — empty paragraph with no children.
                    // Apply default Normal style + docDefaults, matching parse_paragraph().
                    "p" if in_body && depth == 0 => {
                        current_blocks.push(Block::Paragraph(empty_para_with_defaults(styles)));
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // Remaining blocks form the last section
    let last_sp = final_sect_pr.unwrap_or(SectionProperties {
        page_size: PageSize::default(),
        margin: Margin::default(),
        grid_line_pitch: None,
        grid_char_pitch: None,
        grid_char_space_raw: None,
        doc_grid_no_type: false,
        doc_grid_lines_and_chars: false,
        header_refs: Vec::new(),
        footer_refs: Vec::new(),
        columns: None,
        title_pg: false,
        section_type: None,
        page_number_format: None,
        page_number_start: None,
        page_borders: None,
        header_distance: None,
        footer_distance: None,
        bidi: false,
        text_direction: None,
    });
    sections.push(ParsedSection {
        blocks: current_blocks,
        properties: last_sp,
        floating_images: current_floating_images,
        text_boxes: current_text_boxes,
        shapes: current_shapes,
    });

    // §17.2.2 / Round 10 (2026-04-08, COM-confirmed):
    // Word implicitly creates a single empty body paragraph when the
    // <w:body> contains only <w:sectPr> with no <w:p> or <w:tbl> elements
    // (e.g., header_page_number_01, footer_complex_01). Without this,
    // Oxi produces 0 body paragraphs and any tooling that expects a
    // body paragraph (dml_diff, layout cursor placement) breaks.
    if sections.iter().all(|s| s.blocks.is_empty()) {
        if let Some(sec) = sections.last_mut() {
            sec.blocks.push(Block::Paragraph(Paragraph {
                runs: Vec::new(),
                style: ParagraphStyle::default(),
                alignment: Alignment::Left,
                shapes: Vec::new(),
                ppr_change: None,
                paragraph_mark_revision: None,
            }));
        }
    }

    Ok(sections)
}

/// Parse a w:p element (paragraph).
/// Returns (Paragraph, optional SectionProperties if this paragraph ends a section).
/// Parsed paragraph plus any floating elements found inside it
struct ParagraphResult {
    paragraph: Paragraph,
    sect_pr: Option<SectionProperties>,
    shapes: Vec<Shape>,
    text_boxes: Vec<TextBox>,
    inline_images: Vec<Block>,
    floating_images: Vec<Image>,
    /// OMML math blocks extracted from inside this paragraph. Each becomes
    /// a sibling `Block::Math` in the page's block list.
    math_blocks: Vec<crate::ir::MathBlock>,
}

fn parse_paragraph(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet, allow_inline_flow: bool) -> Result<ParagraphResult, ParseError> {
    let mut runs = Vec::new();
    let mut images = Vec::new();
    // S854: (run_index, image) for INLINE (non-positioned) images. A paragraph
    // with MULTIPLE inline images flows them HORIZONTALLY (chord charts, icon
    // rows) via the run's inline_object mechanism instead of extracting each to
    // its own Block::Image line; a single inline image keeps the block path.
    let mut inline_img_runs: Vec<(usize, Image)> = Vec::new();
    let mut found_shapes: Vec<Shape> = Vec::new();
    let mut found_text_boxes: Vec<TextBox> = Vec::new();
    let mut math_blocks: Vec<crate::ir::MathBlock> = Vec::new();
    let mut style = ParagraphStyle::default();
    let mut alignment = Alignment::default();
    // S540 (2026-06-11): explicit `<w:jc w:val="left"/>` parses to Left ==
    // Alignment::default(), which the style/docDefaults inheritance below
    // treated as "unset" and overrode with the style chain's jc (3a4f ① paras:
    // explicit left + default style Normal jc=both → Oxi justified them while
    // Word honors the explicit left). Track explicitness separately.
    let mut has_explicit_jc = false;
    let mut style_id: Option<String> = None;
    let mut num_pr_ref: Option<NumPrRef> = None;
    let mut para_sect_pr: Option<SectionProperties> = None;
    let mut ppr_change: Option<PropertyChange> = None;
    let mut paragraph_mark_revision: Option<TrackedChange> = None;
    let mut depth = 0;
    // Field state: tracks fldChar begin/separate/end across runs.
    // Runs between "separate" and "end" contain cached field results
    // that should be suppressed when the field is evaluated (e.g. PAGE).
    let mut field_result_depth: i32 = 0; // >0 = inside field result region
    let mut current_field_type: Option<FieldType> = None; // tracks the active field across its result runs

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "pPr" if depth == 0 => {
                        let (s, explicit_align, sid, npr, spr, ppr_change_parsed, pmark_rev) = parse_paragraph_properties(reader)?;
                        style = s;
                        if let Some(a) = explicit_align {
                            alignment = a;
                            has_explicit_jc = true;
                        }
                        style_id = sid;
                        num_pr_ref = npr;
                        para_sect_pr = spr;
                        ppr_change = ppr_change_parsed;
                        paragraph_mark_revision = pmark_rev;
                    }
                    // OMML math inside paragraph. Both inline (<m:oMath>) and
                    // display (<m:oMathPara>) are collected and emitted as
                    // sibling Block::Math after the paragraph closes.
                    "oMath" if depth == 0 => {
                        let mb = crate::parser::omml::parse_omath_inline(reader)?;
                        math_blocks.push(mb);
                    }
                    "oMathPara" if depth == 0 => {
                        let mb = crate::parser::omml::parse_omath_para(reader)?;
                        math_blocks.push(mb);
                    }
                    "r" if depth == 0 => {
                        let (mut run, dr) = parse_run(reader, ctx, styles, None)?;
                        // Track field state: fldChar begin/separate/end spans across runs.
                        // Remember a CrossRef field so its cached result run is KEPT.
                        if run.field_type.is_some() {
                            current_field_type = run.field_type.clone();
                        }
                        if run.text.contains('\u{FFFE}') {
                            // Marker for fldChar separate (set in parse_run)
                            run.text = run.text.replace('\u{FFFE}', "");
                            field_result_depth += 1;
                        }
                        if run.text.contains('\u{FFFF}') {
                            // Marker for fldChar end
                            run.text = run.text.replace('\u{FFFF}', "");
                            field_result_depth -= 1;
                            if field_result_depth <= 0 {
                                current_field_type = None;
                            }
                        }
                        // Suppress cached field result text (between separate and end)
                        // when the field was already evaluated (e.g. PAGE → "#"). KEEP the
                        // cached result for CrossRef (REF/NOTEREF/PAGEREF) — Oxi can't
                        // re-resolve the bookmark, so the cache («第１９条») is the display
                        // value; dropping it (old "#") shifted wrapping doc-wide.
                        // KEEP the cache for CrossRef (S685) and Cached (S708,
                        // DATE/TIME/AUTHOR/…); only PAGE/NUMPAGES results are suppressed
                        // (Oxi computes and substitutes those in the layout post-pass).
                        if field_result_depth > 0 && run.field_type.is_none()
                            && !matches!(current_field_type,
                                Some(FieldType::CrossRef) | Some(FieldType::Cached))
                        {
                            run.text.clear();
                        }
                        // S839: an INLINE visual vector group (wpg without
                        // txbxContent — hmrc's checkbox strips) marks its host
                        // run as a width-bearing atomic object; break_into_lines
                        // makes it a U+FFFC fragment of the drawing's extent and
                        // the emit loop draws tb.vector_shapes at the fragment x.
                        if let Some(tb) = dr.as_ref().and_then(|d| d.text_box.as_ref()) {
                            if !tb.vector_shapes.is_empty() && tb.blocks.is_empty()
                                && matches!(tb.wrap_type, Some(crate::ir::WrapType::None))
                                && tb.position.as_ref().map_or(false, |tp| tp.x == 0.0 && tp.y == 0.0
                                    && tp.h_relative.as_deref() == Some("column")
                                    && tp.v_relative.as_deref() == Some("paragraph"))
                            {
                                run.style.inline_object_extent = Some((tb.width, tb.height));
                            }
                        }
                        runs.push(run);
                        if let Some(drawing) = dr {
                            if let Some(image) = drawing.image {
                                // S854: record an INLINE (non-positioned) image with
                                // its run index so ≥2-inline-image paragraphs can flow
                                // horizontally (decided after the run loop). Positioned
                                // images stay on the floating path.
                                if std::env::var("OXI_S854_DISABLE").is_err()
                                    && image.position.is_none()
                                {
                                    inline_img_runs.push((runs.len() - 1, image));
                                } else {
                                    images.push(image);
                                }
                            }
                            if let Some(shape) = drawing.shape {
                                found_shapes.push(shape);
                            }
                            if let Some(tb) = drawing.text_box {
                                found_text_boxes.push(tb);
                            }
                        }
                    }
                    "hyperlink" if depth == 0 => {
                        // w:hyperlink r:id="rIdN" or w:anchor="bookmarkName"
                        let mut link_url: Option<String> = None;
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            if key == "r:id" || key.ends_with(":id") {
                                if let Some(url) = ctx.hyperlinks.get(&val) {
                                    link_url = Some(url.clone());
                                }
                            } else if key == "w:anchor" || key == "anchor" {
                                link_url = Some(format!("#{}", val));
                            }
                        }
                        let hyperlink_runs = parse_hyperlink_runs(reader, ctx, styles, link_url)?;
                        // Word does NOT apply the "Hyperlink" character style inside a
                        // TOC entry: ToC links render in the paragraph text colour
                        // (black), no underline — while body hyperlinks with the SAME
                        // rStyle="Hyperlink" stay blue+underlined (measured: this doc's
                        // body cross-refs are #0000FF, its ToC entries #000000).
                        // Discriminator = the paragraph is a TOC style (TOC1-9).
                        let toc_para = style_id.as_deref()
                            .map_or(false, |s| s.to_ascii_lowercase().starts_with("toc"));
                        // Apply the same field-boundary handling as top-level runs so
                        // fields INSIDE a hyperlink (ToC entries wrap the PAGEREF page
                        // number in <w:hyperlink>) get their fldChar sentinels stripped
                        // and their cached result kept/suppressed correctly. Without
                        // this the ToC page numbers rendered with U+FFFE/U+FFFF tofu.
                        for mut run in hyperlink_runs {
                            if run.field_type.is_some() {
                                current_field_type = run.field_type.clone();
                            }
                            if run.text.contains('\u{FFFE}') {
                                run.text = run.text.replace('\u{FFFE}', "");
                                field_result_depth += 1;
                            }
                            if run.text.contains('\u{FFFF}') {
                                run.text = run.text.replace('\u{FFFF}', "");
                                field_result_depth -= 1;
                                if field_result_depth <= 0 {
                                    current_field_type = None;
                                }
                            }
                            if field_result_depth > 0 && run.field_type.is_none()
                                && !matches!(current_field_type,
                                    Some(FieldType::CrossRef) | Some(FieldType::Cached))
                            {
                                run.text.clear();
                            }
                            if toc_para && run.url.is_some() {
                                run.style.color = None;
                                run.style.underline = false;
                                run.style.underline_style = None;
                            }
                            runs.push(run);
                        }
                    }
                    // Track changes: inserted / deleted / moved content.
                    // ECMA-376 §17.13.5. Each element wraps runs; w:author, w:date,
                    // w:id are attributes on the wrapper. For moves, w:id pairs a
                    // moveFrom with its companion moveTo.
                    // S744 (2026-07-04): <w:smartTag> is a TRANSPARENT wrapper —
                    // its child runs belong to this paragraph (ECMA-376 §17.5.1).
                    // It was falling into the unknown-element catch-all (depth+=1),
                    // so every run inside a smart tag was DROPPED (probezsmarttag:
                    // 「...本規程の趣[20 chars]扱うものとし...」 — the wrapped chunk
                    // vanished; one of the probe hunt's gate-masked content losses).
                    // No depth change: children (w:r, nested smartTags) parse at
                    // depth 0; the matching </w:smartTag> is a no-op in the End arm
                    // (which only decrements when depth > 0). smartTagPr's own
                    // subtree stays skipped via the normal unknown-element path.
                    "smartTag" if depth == 0 => {}
                    "ins" | "del" | "moveFrom" | "moveTo" if depth == 0 => {
                        let change_type = match local.as_str() {
                            "ins" => "insert",
                            "del" => "delete",
                            "moveFrom" => "moveFrom",
                            "moveTo" => "moveTo",
                            _ => unreachable!(),
                        };
                        let mut author = None;
                        let mut date = None;
                        let mut pair_id = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "author" => author = Some(val),
                                "date" => date = Some(val),
                                "id" => pair_id = Some(val),
                                _ => {}
                            }
                        }
                        let tc = TrackedChange {
                            change_type: change_type.into(),
                            author,
                            date,
                            pair_id,
                        };
                        let end_tag = local.clone();
                        let tracked_runs = parse_tracked_change_runs(reader, ctx, styles, &end_tag, tc)?;
                        runs.extend(tracked_runs);
                    }
                    // mc:AlternateContent at paragraph level
                    // OOXML spec (ECMA-376 Part 3): process mc:Choice only, skip mc:Fallback.
                    // mc:Choice may contain both drawings AND text runs (w:r) that belong to this paragraph.
                    "AlternateContent" if depth == 0 => {
                        let mut ac_depth = 0u32;
                        let mut in_choice = false;
                        let mut in_fallback = false;
                        loop {
                            match reader.read_event()? {
                                Event::Start(se) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "Choice" && ac_depth == 0 {
                                        in_choice = true;
                                        ac_depth += 1;
                                    } else if sl == "Fallback" && ac_depth == 0 {
                                        in_fallback = true;
                                        ac_depth += 1;
                                    } else if in_choice && sl == "drawing" && ac_depth == 1 {
                                        let dr = parse_drawing(reader, ctx, styles)?;
                                        if let Some(image) = dr.image {
                                            images.push(image);
                                        }
                                        if let Some(shape) = dr.shape {
                                            found_shapes.push(shape);
                                        }
                                        if let Some(tb) = dr.text_box {
                                            found_text_boxes.push(tb);
                                        }
                                    } else if in_choice && sl == "r" && ac_depth == 1 {
                                        // Text runs inside mc:Choice belong to this paragraph
                                        let (run, dr) = parse_run(reader, ctx, styles, None)?;
                                        runs.push(run);
                                        if let Some(drawing) = dr {
                                            if let Some(image) = drawing.image {
                                                images.push(image);
                                            }
                                            if let Some(shape) = drawing.shape {
                                                found_shapes.push(shape);
                                            }
                                            if let Some(tb) = drawing.text_box {
                                                found_text_boxes.push(tb);
                                            }
                                        }
                                    } else {
                                        ac_depth += 1;
                                    }
                                }
                                Event::End(ee) => {
                                    let sl = local_name(ee.name().as_ref());
                                    if sl == "AlternateContent" && ac_depth == 0 {
                                        break;
                                    }
                                    if sl == "Choice" && in_choice { in_choice = false; }
                                    if sl == "Fallback" && in_fallback { in_fallback = false; }
                                    if ac_depth > 0 { ac_depth -= 1; }
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                    }
                    // Inline Structured Document Tag (w:sdt inside w:p)
                    // ECMA-376: sdtContent contains runs that are part of the paragraph
                    "sdt" if depth == 0 => {
                        let mut sdt_depth = 1u32;
                        let mut in_sdt_content = false;
                        loop {
                            match reader.read_event()? {
                                Event::Start(se) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "sdtContent" && sdt_depth == 1 {
                                        in_sdt_content = true;
                                    } else if in_sdt_content && sl == "r" {
                                        let (run, dr) = parse_run(reader, ctx, styles, None)?;
                                        runs.push(run);
                                        if let Some(drawing) = dr {
                                            if let Some(image) = drawing.image {
                                                images.push(image);
                                            }
                                            if let Some(shape) = drawing.shape {
                                                found_shapes.push(shape);
                                            }
                                            if let Some(tb) = drawing.text_box {
                                                found_text_boxes.push(tb);
                                            }
                                        }
                                    } else {
                                        sdt_depth += 1;
                                    }
                                }
                                Event::End(ee) => {
                                    let sl = local_name(ee.name().as_ref());
                                    if sl == "sdtContent" {
                                        in_sdt_content = false;
                                    } else if sl == "sdt" && sdt_depth == 1 {
                                        break;
                                    } else if sdt_depth > 1 {
                                        sdt_depth -= 1;
                                    }
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                    }
                    // OMML math expressions
                    "oMathPara" | "oMath" if depth == 0 => {
                        let math_text = parse_omml(reader, &local)?;
                        if !math_text.is_empty() {
                            runs.push(Run {
                                text: math_text,
                                style: RunStyle { font_family: Some("Cambria Math".to_string()), ..RunStyle::default() },
                                url: None,
                                footnote_ref: None,
                                endnote_ref: None,
                                comment_range_start: Vec::new(),
                                comment_range_end: Vec::new(),
                                comment_references: Vec::new(),
                                rpr_change: None,
                                tracked_change: None,
                                ruby: None,
                                bookmark_name: None,
                                is_math: true,
                                field_type: None,
                                has_last_rendered_page_break: false,
                            });
                        }
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::Empty(e) => {
                if depth == 0 {
                    let local = local_name(e.name().as_ref());
                    match local.as_str() {
                        "commentRangeStart" => {
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                if key == "id" {
                                    let id = String::from_utf8_lossy(&attr.value).to_string();
                                    // Attach to the previous run when possible; when the
                                    // marker appears before any real run (e.g., at the
                                    // start of a paragraph whose first element is the
                                    // commentRangeStart), create an empty anchor run so
                                    // the id survives to the IR. Without the anchor the
                                    // id is silently dropped — that breaks any
                                    // range-aware renderer pass (R-04 highlight, etc.)
                                    // whenever a comment range begins at a paragraph
                                    // boundary.
                                    if let Some(last_run) = runs.last_mut() {
                                        last_run.comment_range_start.push(id);
                                    } else {
                                        runs.push(Run {
                                            text: String::new(),
                                            style: RunStyle::default(),
                                            url: None,
                                            footnote_ref: None,
                                            endnote_ref: None,
                                            comment_range_start: vec![id],
                                            comment_range_end: Vec::new(),
                                            comment_references: Vec::new(),
                                            tracked_change: None,
                                            rpr_change: None,
                                            ruby: None,
                                            bookmark_name: None,
                                            is_math: false,
                                            field_type: None,
                                            has_last_rendered_page_break: false,
                                        });
                                    }
                                }
                            }
                        }
                        "commentRangeEnd" => {
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                if key == "id" {
                                    let id = String::from_utf8_lossy(&attr.value).to_string();
                                    if let Some(last_run) = runs.last_mut() {
                                        last_run.comment_range_end.push(id);
                                    } else {
                                        // Symmetric with commentRangeStart — create an
                                        // anchor run if there is no prior run to stamp.
                                        // This is rare (usually rangeEnd comes after some
                                        // content) but covers the "empty paragraph that
                                        // only contains the close marker" case.
                                        runs.push(Run {
                                            text: String::new(),
                                            style: RunStyle::default(),
                                            url: None,
                                            footnote_ref: None,
                                            endnote_ref: None,
                                            comment_range_start: Vec::new(),
                                            comment_range_end: vec![id],
                                            comment_references: Vec::new(),
                                            tracked_change: None,
                                            rpr_change: None,
                                            ruby: None,
                                            bookmark_name: None,
                                            is_math: false,
                                            field_type: None,
                                            has_last_rendered_page_break: false,
                                        });
                                    }
                                }
                            }
                        }
                        "bookmarkStart" => {
                            let mut bk_name = None;
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                if key == "name" {
                                    let val = String::from_utf8_lossy(&attr.value).to_string();
                                    if val != "_GoBack" {
                                        bk_name = Some(val);
                                    }
                                }
                            }
                            if let Some(name) = bk_name {
                                // Create an empty anchor run for the bookmark
                                runs.push(Run {
                                    text: String::new(),
                                    style: RunStyle::default(),
                                    url: None,
                                    footnote_ref: None,
                                    endnote_ref: None,
                                    comment_range_start: Vec::new(),
                                    comment_range_end: Vec::new(),
                                    comment_references: Vec::new(),
                                    tracked_change: None,
                                    rpr_change: None,
                                    ruby: None,
                                    bookmark_name: Some(name),
                                    is_math: false,
                                    field_type: None,
                                    has_last_rendered_page_break: false,
                                });
                            }
                        }
                        "bookmarkEnd" => {
                            // End marker; anchor is already placed at bookmarkStart
                        }
                        _ => {}
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "p" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // 2026-05-01 R12: snapshot pPr-explicit indent BEFORE style inheritance.
    // LibreOffice algorithm (ndtxt.cxx:AreListLevelIndentsApplicable, line 4820)
    // applies numbering's indent ONLY when paragraph has NO hard-set indent (in
    // pPr) AND no style in the basedOn chain has hard-set indent. This snapshot
    // captures pPr-explicit; the style chain check is implicit in the order of
    // operations below.
    let ppr_explicit_indent_left =
        style.indent_left.is_some() || style.indent_left_chars.is_some();
    let ppr_explicit_first_line =
        style.indent_first_line.is_some() || style.indent_first_line_chars.is_some();

    // Apply style inheritance from StyleSheet (basedOn already resolved)
    // ECMA-376: paragraph with no pStyle implicitly uses the default paragraph style (w:default="1")
    let effective_style_id = style_id.clone()
        .or_else(|| styles.default_paragraph_style_id.clone());
    if let Some(ref sid) = effective_style_id {
        if let Some(defined) = styles.styles.get(sid) {
            // Inherit alignment from style if not explicitly set in paragraph
            // (S540: explicit jc=left == default Left, so gate on the flag,
            // not the value)
            if !has_explicit_jc && alignment == Alignment::default() {
                if let Some(style_align) = defined.alignment {
                    alignment = style_align;
                }
            }
            let ds = &defined.paragraph;
            if style.space_before.is_none() {
                style.space_before = ds.space_before;
            }
            if style.space_after.is_none() {
                style.space_after = ds.space_after;
            }
            if style.line_spacing.is_none() {
                style.line_spacing = ds.line_spacing;
                style.line_spacing_rule = ds.line_spacing_rule.clone();
            }
            // Bug B Day 18 (2026-05-08): replace hand-rolled field merge with
            // shared merge_run_style helper from parser/styles.rs. The hand-
            // rolled version inherited only 7 fields (font_size, font_family,
            // font_family_east_asia, has_explicit_east_asia, color, bold,
            // italic) plus character_spacing added at Day 16. The helper
            // additionally covers highlight, shading, font_size_cs, kern,
            // text_scale, underline, strikethrough, vertical_align.
            //
            // Day 4 attempted this same broad replacement and was reverted
            // due to "cb8be 0.91→0.45 catastrophic" SSIM result. Day 13's
            // drift discovery showed that result was likely drift artifact;
            // re-evaluating on drift-free baseline.
            if let Some(ref style_rs) = ds.default_run_style {
                if let Some(ref mut para_rs) = style.default_run_style {
                    super::styles::merge_run_style(para_rs, style_rs);
                } else {
                    style.default_run_style = ds.default_run_style.clone();
                }
            }
            // Inherit keepNext, keepLines, contextualSpacing, widowControl from style
            if ds.keep_next { style.keep_next = true; }
            if ds.keep_lines { style.keep_lines = true; }
            // S782 (2026-07-11): a DIRECT `<w:contextualSpacing w:val="0"/>`
            // (explicit OFF) wins over the style's contextualSpacing — the
            // merge used to clobber it back to true (CT_OnOff val=0 class,
            // same shape as has_explicit_widow_control / snapToGrid S606b).
            // nyserda '(b)/(c)' cost items: ListParagraph style contextual ON,
            // direct val=0 + before=240 → Word renders the 12pt before (gap
            // 25.8 vs Oxi 13.8). Only nyserda carries the pattern corpus-wide.
            if ds.contextual_spacing && !style.has_explicit_contextual_spacing {
                style.contextual_spacing = true;
            }
            // S675: before/afterAutospacing is NOT inherited from the paragraph style
            // (Word applies HTML autospacing only on the direct paragraph pPr).
            // Inherit widow_control: style's explicit setting takes precedence
            if !style.has_explicit_widow_control && ds.has_explicit_widow_control {
                style.widow_control = ds.widow_control;
            }
            // Inherit numPr from style definition
            if style.num_id.is_none() {
                if let Some(ref nid) = ds.num_id {
                    style.num_id = Some(nid.clone());
                    style.num_ilvl = ds.num_ilvl;
                }
            }
            // Inherit tab stops from style
            if style.tab_stops.is_empty() && !ds.tab_stops.is_empty() {
                style.tab_stops = ds.tab_stops.clone();
            }
            // Inherit paragraph borders from style
            if style.borders.is_none() {
                style.borders = ds.borders.clone();
            }
            // Inherit indents from style
            if style.indent_left.is_none() && style.indent_left_chars.is_none() {
                style.indent_left = ds.indent_left;
                style.indent_left_chars = ds.indent_left_chars;
            }
            if style.indent_right.is_none() && style.indent_right_chars.is_none() {
                style.indent_right = ds.indent_right;
                style.indent_right_chars = ds.indent_right_chars;
            }
            if style.indent_first_line.is_none() && style.indent_first_line_chars.is_none() {
                style.indent_first_line = ds.indent_first_line;
                style.indent_first_line_chars = ds.indent_first_line_chars;
            }
            // Inherit shading from style
            if style.shading.is_none() {
                style.shading = ds.shading.clone();
            }
            // Inherit page_break_before from style
            if ds.page_break_before {
                style.page_break_before = true;
            }
            // Inherit snap_to_grid from style (false overrides struct default true).
            // Round 29: footnote text style "a8" sets snapToGrid=0; without this
            // inheritance, footnote paragraphs were grid-snapped to body line
            // pitch, causing wide line spacing in the footnote area.
            // S606b (2026-06-20, Word COM-confirmed): a DIRECT paragraph
            // snapToGrid (has_explicit_snap_to_grid) overrides the style — a
            // no-val `<w:snapToGrid/>` (CT_OnOff = true) re-enables grid snap
            // despite the style's snapToGrid=0 (ohnoikuji a4 "header" list items).
            let s606b = std::env::var("OXI_S606B_DISABLE").is_err();
            if !ds.snap_to_grid && !(s606b && style.has_explicit_snap_to_grid) {
                style.snap_to_grid = false;
            }
            // Session 85 fix: inherit auto_space_de from style (false overrides
            // default true). Mirrors snap_to_grid pattern. Confirmed via CR9
            // minimal repro: when paragraph has BOTH direct pPr autoSpaceDE=0
            // AND pStyle ac (also autoSpaceDE=0), Oxi fits 36 chars matching
            // Word. CR6 (pStyle ac only) fits 35 chars → pStyle inheritance
            // for auto_space_de was broken. tokumei_08_01 series (a1d6/d4d126/
            // de6e/etc, 22 baseline docs) uses style "ac" with autoSpaceDE=0.
            if !ds.auto_space_de {
                style.auto_space_de = false;
            }
            // Session 95 (2026-05-18) symmetric fix: inherit auto_space_dn
            // (East Asian ↔ digit auto-spacing). S85 only handled auto_space_de
            // because the layout code at that point gated digits on the same
            // flag (is_ascii_alphanumeric). S95 split alpha vs digit at 4 call
            // sites in mod.rs; without dn inheritance, a1d6/d4d126/de6e "ac"
            // paragraphs would have dn=true (default) and over-space digits.
            if !ds.auto_space_dn {
                style.auto_space_dn = false;
            }
            // S301 (2026-05-26) symmetric fix: inherit word_wrap from style
            // (false overrides default true). Mirrors snap_to_grid /
            // auto_space_de / auto_space_dn. Style "ac" (一太郎) used by
            // 29dc6e / d4d126 / etc. sets `<w:wordWrap w:val="0"/>` to enable
            // CJK-aware break-anywhere wrap. Without this inheritance, the
            // S301 wrap-padding discriminator (`!para.style.word_wrap`) could
            // not distinguish 29dc6e (style "ac", wordWrap=0) from 191cb
            // (no pStyle, wordWrap=true) — both ended up `word_wrap=true` in
            // the IR despite Word's actual behavior diverging on exactly
            // that property.
            if !ds.word_wrap {
                style.word_wrap = false;
            }
        }
    }

    // Apply docDefaults fallback (field-by-field merge per ECMA-376 priority)
    if let Some(ref doc_rs) = styles.doc_default_run_style {
        if let Some(ref mut para_rs) = style.default_run_style {
            if para_rs.font_size.is_none() { para_rs.font_size = doc_rs.font_size; }
            if para_rs.font_family.is_none() { para_rs.font_family = doc_rs.font_family.clone(); }
            if para_rs.font_family_east_asia.is_none() { para_rs.font_family_east_asia = doc_rs.font_family_east_asia.clone(); }
            if !para_rs.has_explicit_east_asia && doc_rs.has_explicit_east_asia { para_rs.has_explicit_east_asia = true; }
            if para_rs.color.is_none() { para_rs.color = doc_rs.color.clone(); }
            // S547 (2026-06-12): w:kern gates the yakumono pair halving — must
            // survive the docDefaults merge like the font fields.
            if para_rs.kern.is_none() { para_rs.kern = doc_rs.kern; }
        } else {
            style.default_run_style = styles.doc_default_run_style.clone();
        }
    }
    if let Some(ref doc_para) = styles.doc_default_para_style {
        if style.space_before.is_none() {
            style.space_before = doc_para.space_before;
        }
        if style.space_after.is_none() {
            style.space_after = doc_para.space_after;
        }
        if style.before_lines.is_none() {
            style.before_lines = doc_para.before_lines;
        }
        if style.after_lines.is_none() {
            style.after_lines = doc_para.after_lines;
        }
        if style.line_spacing.is_none() {
            style.line_spacing = doc_para.line_spacing;
            style.line_spacing_rule = doc_para.line_spacing_rule.clone();
            style.line_spacing_from_doc_defaults = true;
        }
        if style.indent_left.is_none() && style.indent_left_chars.is_none() {
            style.indent_left = doc_para.indent_left;
            style.indent_left_chars = doc_para.indent_left_chars;
        }
        if style.indent_right.is_none() && style.indent_right_chars.is_none() {
            style.indent_right = doc_para.indent_right;
            style.indent_right_chars = doc_para.indent_right_chars;
        }
        if style.indent_first_line.is_none() && style.indent_first_line_chars.is_none() {
            style.indent_first_line = doc_para.indent_first_line;
            style.indent_first_line_chars = doc_para.indent_first_line_chars;
        }
        // Only override widow_control from docDefaults if docDefaults explicitly
        // sets widowControl. When pPrDefault is empty, doc_para has the struct
        // default (true), which would incorrectly override Normal style's false.
        if !style.has_explicit_widow_control && doc_para.has_explicit_widow_control {
            style.widow_control = doc_para.widow_control;
        }
        // textAlignment (§17.3.1.35) inheritance from pPrDefault.
        if style.text_alignment.is_none() {
            style.text_alignment = doc_para.text_alignment.clone();
            // R7.63: mark as inherited from pPrDefault so layout can distinguish
            // document-wide baseline (e3c545: all paras get offset=0) from
            // per-paragraph baseline (ed025c wi=827: only one para, must NOT
            // suppress centering or its gap to wi=826 collapses by 4pt).
            if style.text_alignment.is_some() {
                style.text_alignment_from_pprdefault = true;
            }
        }
    }
    // Inherit alignment from docDefaults pPrDefault (jc)
    // (S540: explicit jc=left must not be overridden — see has_explicit_jc)
    if !has_explicit_jc && alignment == Alignment::default() {
        if let Some(doc_align) = styles.doc_default_alignment {
            alignment = doc_align;
        }
    }

    // S771: was this numPr set DIRECTLY on the paragraph's pPr, or inherited
    // from the paragraph style? The R12 numbering-indent-override below is only
    // correct for a DIRECT numPr (d77a). When the numPr is STYLE-inherited AND
    // the style also defines its own `w:ind`, the style indent wins (ECMA-376 /
    // LibreOffice AreListLevelIndentsApplicable). nyserda ListBullet2 =
    // style-inherited numPr(9) with style ind left=1152tw=57.6pt but numbering
    // level ind left=5490tw=274.5pt; Word uses 57.6, Oxi was applying 274.5 →
    // sub-bullets over-indented ~200pt → wrap to a narrow column → +2 pages.
    let num_pr_is_direct = num_pr_ref.is_some();
    // Inherit numPr from style definition if paragraph doesn't have its own
    if num_pr_ref.is_none() {
        if let Some(ref nid) = style.num_id {
            if !nid.is_empty() && nid != "0" {
                num_pr_ref = Some(NumPrRef {
                    num_id: nid.clone(),
                    ilvl: style.num_ilvl,
                });
            }
        }
    }

    // Resolve list marker from numbering definitions
    if let Some(npr) = num_pr_ref {
        if !npr.num_id.is_empty() && npr.num_id != "0" {
            let resolved = ctx.numbering.resolve_marker_full(
                &npr.num_id,
                npr.ilvl,
                &mut ctx.list_counters.borrow_mut(),
            );
            // S777: an EMPTY resolved marker (numFmt=none) sets no marker
            // element; keep suff/indent resolution unchanged.
            if !resolved.text.is_empty() {
                style.list_marker = Some(resolved.text);
            }
            style.list_suff = Some(resolved.suff);
            style.list_tab_stop = resolved.tab_stop;
            style.list_marker_size = resolved.marker_size;
            // S778 gate = the S771 discriminator: the level ind acts as the
            // suffix-tab stop only for a DIRECT numPr (style-inherited lists
            // keep the style ind — nyserda ListBullet2 level left=274.5pt as
            // a stop re-created the S771 over-indent, Exhibit E +1 page).
            style.list_level_left = if num_pr_is_direct { resolved.level_left } else { None };
            if let Some(ind) = resolved.hanging {
                // Paragraph's explicit hanging indent overrides numbering level's hanging.
                // COM-confirmed (LOD_Handbook P3: XML hanging=426tw=21.3pt overrides
                // numbering hanging=720tw=36pt).
                if let Some(first) = style.indent_first_line {
                    if first < 0.0 {
                        style.list_indent = Some(-first);
                    } else {
                        style.list_indent = Some(ind);
                    }
                } else {
                    style.list_indent = Some(ind);
                }
                if style.indent_left.is_none() {
                    if let Some(left) = ctx.numbering.get_level_indent(&npr.num_id, npr.ilvl) {
                        style.indent_left = Some(left);
                    }
                }
            }
            // 2026-05-01 R12: numbering's left/hanging indent OVERRIDES
            // style-inherited indent (but NOT pPr-explicit indent). Implements
            // LibreOffice's `AreListLevelIndentsApplicable` algorithm
            // (sw/source/core/txtnode/ndtxt.cxx:4820): when paragraph has direct
            // numPr in pPr and no hard-set indent in pPr, the list-level indents
            // apply (i.e., numbering's `ind left/hanging` wins over the style
            // chain's indent). Pixel-verified on d77a58485f16 p10 p5: numbering
            // num=1 ilvl=0 has ind left=780tw=39pt hanging=360tw=18pt. Word
            // visually renders L1 body at margin+21pt, L2+ at margin+39pt
            // (standard hanging-indent), NOT at style a9's ind left=720tw=36pt.
            // Note: COM Format.LeftIndent / Information(5) for Range.Characters
            // are unreliable for hanging-indent paragraphs — pixel measurement
            // of Word PNG required to verify visual position.
            // S771 (2026-07-09, ★default ON, opt-out OXI_S771_DISABLE): only
            // override the style/inherited indent with the numbering level's indent
            // when the numPr is DIRECT on the paragraph. A style-inherited numPr
            // whose style carries its own w:ind should keep that style ind
            // (ECMA-376 / LibreOffice AreListLevelIndentsApplicable). nyserda
            // ListBullet2 = style-inherited numPr(9), style ind left=57.6pt but
            // numbering level ind left=274.5pt; Word uses 57.6, Oxi applies 274.5 →
            // sub-bullets over-indented ~200pt → narrow wrap → +2 pages in Exhibit E.
            // ★ECMA-CORRECT + JP byte-identical (238 word_png, +0.0000). Fixes the
            // Exhibit E +2 drift; the separate pre-Exhibit-C −1 drift it exposed is now
            // closed by LATINEM (no-kern Latin em break), so S771 + LATINEM = nyserda
            // 56 = Word (Exhibit boundaries aligned). framework neutral.
            // See [[english_corpus_bug_mine]].
            let s771_on = std::env::var("OXI_S771_DISABLE").is_err();
            let s771_apply_num_ind = !s771_on
                || num_pr_is_direct
                || style.indent_left.is_none();
            if !ppr_explicit_indent_left && s771_apply_num_ind {
                if let Some(left) = ctx.numbering.get_level_indent(&npr.num_id, npr.ilvl) {
                    style.indent_left = Some(left);
                    style.indent_left_chars = None;
                }
            }
            let s771_apply_num_hang = !s771_on
                || num_pr_is_direct
                || style.indent_first_line.is_none();
            if !ppr_explicit_first_line && s771_apply_num_hang {
                if let Some(hanging) = ctx.numbering.get_level_hanging(&npr.num_id, npr.ilvl) {
                    style.indent_first_line = Some(-hanging);
                    style.indent_first_line_chars = None;
                } else if let Some(fl) = ctx.numbering.get_level_first_line(&npr.num_id, npr.ilvl) {
                    // S781 (2026-07-11): the level may carry a POSITIVE w:firstLine
                    // instead of w:hanging (mutually exclusive; hanging wins).
                    // nyserda "(b) Direct Charges" numId=18 lvl0 = left=0
                    // firstLine=720 + suff=nothing: Word renders the marker at
                    // margin+36 (x=126) with continuation at the margin; dropping
                    // firstLine placed the whole first line at the margin →
                    // 36pt-wider wrap → the p12+ phase drift. Corpus scan: level
                    // firstLine exists ONLY in 5 real_en English docs, 0 JP docs.
                    if std::env::var("OXI_S781_DISABLE").is_err() {
                        style.indent_first_line = Some(fl);
                        style.indent_first_line_chars = None;
                    }
                }
            }
        }
    }

    // Convert inline page break (w:br type="page" as \x0C in first run).
    // COM-confirmed 2026-04-17: when the paragraph is empty except for the br,
    // Word renders the paragraph's mark as a stub line on the CURRENT page
    // then breaks. That maps to `page_break_after`. When the br is followed by
    // other non-empty content, it remains `page_break_before` (existing behavior).
    // See `project_empty_br_para_stub.md`.
    if let Some(first_run) = runs.first() {
        if first_run.text.trim() == "\x0C" || first_run.text == "\x0C" {
            let only_br = runs.len() == 1
                || runs.iter().skip(1).all(|r| r.text.is_empty());
            if only_br {
                style.page_break_after = true;
            } else {
                style.page_break_before = true;
            }
            runs.remove(0); // Remove the break-only run
        }
    }

    // Store style ID for contextual spacing comparison
    style.style_id = style_id;

    // S854: resolve the recorded inline images. ≥2 in one paragraph → flow
    // HORIZONTALLY by carrying each on its run as an inline object (the S839/
    // S851 U+FFFC mechanism), matching Word's inline row (chord charts:
    // educational__0003daa8 stacked 8× 47.7pt images vertically = +330pt over
    // Word's single 48pt row → +1 page; hmrc "Male [☐] Female [☐]" checkboxes).
    // 0 or 1 → restore the block path (byte-identical for single-figure paras).
    // Scoped to BODY paragraphs (allow_inline_flow): a textbox/table-cell/
    // header/footnote paragraph keeps the block path — its inline_object_image
    // is NOT drawn by those emit paths (kyotei36spec's 18 form-marker images
    // live in a textbox at the page-right edge; flowing them there dropped all
    // 18 from the render). Body emit (mod.rs:11135) draws inline_object_image.
    if allow_inline_flow && inline_img_runs.len() >= 2 {
        for (ridx, image) in inline_img_runs {
            if let Some(run) = runs.get_mut(ridx) {
                run.style.inline_object_extent = Some((image.width, image.height));
                run.style.inline_object_image = Some(Box::new(image));
            }
        }
    } else {
        for (_ridx, image) in inline_img_runs {
            images.push(image);
        }
    }

    // Separate inline images (no position) from floating images
    let mut inline_images: Vec<Block> = Vec::new();
    let mut floating_imgs: Vec<Image> = Vec::new();
    for img in images {
        if img.position.is_some() {
            floating_imgs.push(img);
        } else {
            inline_images.push(Block::Image(img));
        }
    }

    // Bug B Day 18: propagate paragraph's default_run_style into individual
    // runs via merge_run_style (parser/styles.rs). Layout reads run-level
    // rPr fields directly with no fallback to para.default_run_style for
    // many fields (e.g. character_spacing). Without this step pStyle-level
    // properties are silently dropped.
    //
    // Day 16 introduced this for character_spacing only; Day 18 broadens
    // to all merge_run_style-covered fields (cs + highlight + shading +
    // font_size_cs + kern + text_scale + underline + strikethrough +
    // vertical_align + font_size + font_family + font_family_east_asia +
    // color + bold + italic).
    if let Some(ref para_drs) = style.default_run_style {
        for run in runs.iter_mut() {
            super::styles::merge_run_style(&mut run.style, para_drs);
        }
    }

    // RUN-PRESENCE rule — ★FALSIFIED (2026-07-07 level 10; kept as a
    // standalone opt-in tombstone). Hypothesis: a paragraph whose runs are
    // ALL text-empty but non-empty (a text-less anchor-run holder) sizes
    // its line from the RUN's inherited properties instead of the ¶-mark
    // rPr. The 2ea81a "1-cell stack" reading that motivated it was an
    // Information(6) CENTERING mis-read (the documented Info6 trap): the
    // real Word stack is pi28 = TWO plain ¶-sz28 grid cells (32.3) with
    // the line box CENTERED in the 2-cell space (Info6 reports 43.85 =
    // 36.85 + (32.3−18.16)/2 = 43.92, and pi29 starts exactly at the cell
    // boundary 69.15). Controlled proof: _anchorclamp_sweep.py's anchor
    // paras (¶sz28 + anchor run inheriting 10.5pt) advance 32.6 ≈ 2
    // RELATIVE cells of the ¶ size — the ¶-mark rPr WINS even with an
    // anchor run present; no special rule exists. 2ea81a's real residual
    // is a ~10pt upstream p1 accumulation (pi27 cursor Oxi 789.95 vs Word
    // 779.5). OXI_RUNPRESENCE=1 keeps the falsified behavior for A/B.
    if std::env::var("OXI_RUNPRESENCE").is_ok()
        && std::env::var("OXI_RUNPRESENCE_DISABLE").is_err()
        && !runs.is_empty()
        && runs.iter().all(|r| r.text.is_empty())
    {
        style.ppr_rpr = Some(runs[0].style.clone());
    }

    Ok(ParagraphResult {
        paragraph: Paragraph {
            runs,
            style,
            alignment,
            shapes: found_shapes.clone(),
            ppr_change,
            paragraph_mark_revision,
        },
        sect_pr: para_sect_pr,
        shapes: Vec::new(), // shapes are now in Paragraph.shapes, not page-level
        text_boxes: found_text_boxes,
        inline_images,
        floating_images: floating_imgs,
        math_blocks,
    })
}

/// Numbering reference parsed from w:numPr
struct NumPrRef {
    num_id: String,
    ilvl: u8,
}

/// Parse w:pPr (paragraph properties).
/// Returns: (style, alignment, style_id, numPr, optional section properties for section break)
fn parse_paragraph_properties(
    reader: &mut Reader<&[u8]>,
) -> Result<
    (
        ParagraphStyle,
        Option<Alignment>,
        Option<String>,
        Option<NumPrRef>,
        Option<SectionProperties>,
        Option<PropertyChange>,
        Option<TrackedChange>,
    ),
    ParseError,
> {
    let mut style = ParagraphStyle::default();
    let mut alignment: Option<Alignment> = None;
    let mut style_id: Option<String> = None;
    let mut num_pr: Option<NumPrRef> = None;
    let mut sect_pr: Option<SectionProperties> = None;
    let mut ppr_change: Option<PropertyChange> = None;
    let mut paragraph_mark_revision: Option<TrackedChange> = None;
    let mut has_explicit_widow_control = false;
    let mut has_explicit_snap_to_grid = false;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "rPr" if depth == 0 => {
                        // pPr/rPr: paragraph-level default run properties (empty para font)
                        let mut ppr_rpr = RunStyle::default();
                        let mut rd = 0;
                        loop {
                            match reader.read_event()? {
                                Event::Empty(e2) => {
                                    let l = local_name(e2.name().as_ref());
                                    if l == "sz" {
                                        for a in e2.attributes().flatten() {
                                            if local_name(a.key.as_ref()) == "val" {
                                                if let Ok(v) = std::str::from_utf8(&a.value) {
                                                    if let Ok(hp) = v.parse::<f32>() { ppr_rpr.font_size = Some(hp / 2.0); }
                                                }
                                            }
                                        }
                                    } else if l == "rFonts" {
                                        for a in e2.attributes().flatten() {
                                            let k = local_name(a.key.as_ref());
                                            let v = String::from_utf8_lossy(&a.value).to_string();
                                            match k.as_str() {
                                                // S583 (2026-06-16): the paragraph-mark / empty-para
                                                // line height is governed by the ASCII font (Word
                                                // measures the ¶ as a Latin glyph). When ascii≠hAnsi
                                                // (e.g. kojin's 様式 spacer: ascii=HGPｺﾞｼｯｸM hAnsi=Century)
                                                // ascii must win — the old `ascii|hAnsi => set` made
                                                // hAnsi (parsed last) overwrite ascii → Century → the
                                                // 14pt empty wrongly snapped to 1 cell. Opt-out
                                                // OXI_S583_DISABLE restores the old last-wins behavior
                                                // (keeps the layout-side gate's canary complete).
                                                "ascii" => { ppr_rpr.font_family = Some(v); }
                                                "hAnsi" => {
                                                    if ppr_rpr.font_family.is_none()
                                                        || std::env::var("OXI_S583_DISABLE").is_ok() {
                                                        ppr_rpr.font_family = Some(v);
                                                    }
                                                }
                                                "eastAsia" => {
                                                    ppr_rpr.font_family_east_asia = Some(v);
                                                    ppr_rpr.has_explicit_east_asia = true;
                                                }
                                                _ => {}
                                            }
                                        }
                                    } else if l == "b" { ppr_rpr.bold = true; }
                                    else if l == "vanish" {
                                        // S673v (2026-06-26): the ¶ MARK is hidden. An empty
                                        // paragraph with a hidden mark COLLAPSES to 0 height in
                                        // Word (invisible separator before a table idiom —
                                        // 3a4f/model/tokyoshugyo). The pPr/rPr parser handled
                                        // only font/bold; vanish was dropped. layout_paragraph
                                        // reads ppr_rpr.vanish to skip the para.
                                        // NOTE: webHidden is web-only (rendered in print), so it
                                        // does NOT trigger the collapse — only true w:vanish does.
                                        ppr_rpr.vanish = true;
                                    }
                                    else if (l == "ins" || l == "del") && paragraph_mark_revision.is_none() {
                                        // Paragraph-mark revision: `<w:pPr>/<w:rPr>/<w:ins>` or
                                        // `<w:pPr>/<w:rPr>/<w:del>` marks the pilcrow (¶) itself
                                        // as inserted or deleted (revisions_notes.md §2). Empty
                                        // element, attrs only.
                                        let change_type = if l == "ins" { "insert" } else { "delete" };
                                        let mut author = None;
                                        let mut date = None;
                                        let mut pair_id = None;
                                        for a in e2.attributes().flatten() {
                                            let k = local_name(a.key.as_ref());
                                            let v = String::from_utf8_lossy(&a.value).to_string();
                                            match k.as_str() {
                                                "author" => author = Some(v),
                                                "date" => date = Some(v),
                                                "id" => pair_id = Some(v),
                                                _ => {}
                                            }
                                        }
                                        paragraph_mark_revision = Some(TrackedChange {
                                            change_type: change_type.into(),
                                            author,
                                            date,
                                            pair_id,
                                        });
                                    }
                                }
                                Event::Start(_) => { rd += 1; }
                                Event::End(_) => { if rd == 0 { break; } rd -= 1; }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                        style.ppr_rpr = Some(ppr_rpr);
                    }
                    "numPr" if depth == 0 => {
                        let npr = parse_num_pr(reader)?;
                        // R95 (2026-04-30): also mirror the parsed inline
                        // numPr onto style.num_id/num_ilvl so describe_ppr_diff
                        // (R-12 v3 ppr-axis) can compare prior vs current
                        // numbering when a pPrChange records an inline numPr
                        // toggle. Layout reads via num_pr_ref directly (line
                        // ~1540-1552 reconstructs num_pr_ref from style.num_id
                        // only when num_pr_ref is None — never the other way
                        // around) so this is read-only from the perspective
                        // of layout / list rendering.
                        if !npr.num_id.is_empty() {
                            style.num_id = Some(npr.num_id.clone());
                            style.num_ilvl = npr.ilvl;
                        }
                        num_pr = Some(npr);
                    }
                    "spacing" if depth == 0 => {
                        style.has_direct_spacing = true;
                        let mut line_val: Option<f32> = None;
                        let mut line_rule: Option<String> = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "before" => {
                                    style.space_before =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                    style.has_direct_before_after = true;
                                }
                                "after" => {
                                    style.space_after =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                    style.has_direct_before_after = true;
                                }
                                "beforeLines" => {
                                    style.before_lines = val.parse::<f32>().ok();
                                    style.has_direct_before_after = true;
                                }
                                "afterLines" => {
                                    style.after_lines = val.parse::<f32>().ok();
                                    style.has_direct_before_after = true;
                                }
                                "beforeAutospacing" => {
                                    // CT_OnOff attr: "1"/"true"/"on" => true (S675)
                                    let v = val.as_ref();
                                    style.before_autospacing = v == "1" || v == "true" || v == "on";
                                    style.has_direct_before_after = true;
                                }
                                "afterAutospacing" => {
                                    let v = val.as_ref();
                                    style.after_autospacing = v == "1" || v == "true" || v == "on";
                                    style.has_direct_before_after = true;
                                }
                                "line" => {
                                    line_val = val.parse::<f32>().ok();
                                }
                                "lineRule" => {
                                    line_rule = Some(val.to_string());
                                }
                                _ => {}
                            }
                        }
                        if let Some(lv) = line_val {
                            match line_rule.as_deref() {
                                Some("exact") => {
                                    style.line_spacing = Some(lv / 20.0);
                                    style.line_spacing_rule = Some("exact".to_string());
                                }
                                Some("atLeast") => {
                                    style.line_spacing = Some(lv / 20.0);
                                    style.line_spacing_rule = Some("atLeast".to_string());
                                }
                                _ if lv < 0.0 => {
                                    // Word quirk (COM-confirmed 2026-05-15 on d4d126):
                                    // negative w:line with lineRule="auto" (or missing) is
                                    // treated as wdLineSpaceExactly with |val|/20 pt.
                                    style.line_spacing = Some(lv.abs() / 20.0);
                                    style.line_spacing_rule = Some("exact".to_string());
                                }
                                _ => {
                                    style.line_spacing = Some(lv / 240.0);
                                }
                            }
                        }
                        depth += 1;
                    }
                    "tabs" if depth == 0 => {
                        style.tab_stops = parse_tab_stops(reader)?;
                    }
                    "pBdr" if depth == 0 => {
                        style.borders = Some(parse_paragraph_borders(reader)?);
                    }
                    "sectPr" if depth == 0 => {
                        sect_pr = Some(parse_section_properties(reader)?);
                    }
                    // Paragraph-property change (`<w:pPrChange>`): body contains
                    // a prior `<w:pPr>` with the pre-edit paragraph style.
                    // Drain it explicitly so its inner Empty children (jc, ind,
                    // spacing…) don't silently overwrite the current style —
                    // the outer handlers don't gate on depth for Empty events.
                    "pPrChange" if depth == 0 => {
                        let mut pc = PropertyChange::default();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "id" => pc.id = Some(val),
                                "author" => pc.author = Some(val),
                                "date" => pc.date = Some(val),
                                _ => {}
                            }
                        }
                        loop {
                            match reader.read_event()? {
                                Event::Start(inner) => {
                                    if local_name(inner.name().as_ref()) == "pPr" {
                                        let (prior, prior_explicit_align, _sid, _npr, _spr, _nested, _pmr) =
                                            parse_paragraph_properties(reader)?;
                                        pc.prior_paragraph_style = Some(Box::new(prior));
                                        // R72: capture prior <w:jc> if the inner pPr
                                        // declared one. Only set when explicit — a
                                        // missing jc means "inherit from style", not
                                        // "Left" (parse_paragraph_properties returns
                                        // None when no <w:jc> child was seen).
                                        if let Some(a) = prior_explicit_align {
                                            pc.prior_alignment = Some(a);
                                        }
                                    }
                                }
                                Event::Empty(inner) => {
                                    if local_name(inner.name().as_ref()) == "pPr"
                                        && pc.prior_paragraph_style.is_none()
                                    {
                                        pc.prior_paragraph_style =
                                            Some(Box::new(ParagraphStyle::default()));
                                    }
                                }
                                Event::End(inner) => {
                                    if local_name(inner.name().as_ref()) == "pPrChange" {
                                        break;
                                    }
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                        ppr_change = Some(pc);
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "pStyle" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val.starts_with("Heading") {
                                    style.heading_level =
                                        val.trim_start_matches("Heading").parse().ok();
                                }
                                style_id = Some(val);
                            }
                        }
                    }
                    "jc" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                alignment = Some(match val.as_ref() {
                                    "left" | "start" => Alignment::Left,
                                    "center" => Alignment::Center,
                                    "right" | "end" => Alignment::Right,
                                    "both" => Alignment::Justify,
                                    "distribute" => Alignment::Distribute,
                                    _ => Alignment::Left,
                                });
                            }
                        }
                    }
                    "framePr" => {
                        let mut fp = crate::ir::FrameProperties::default();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "dropCap" => fp.drop_cap = Some(val.to_string()),
                                "lines" => fp.lines = val.parse().unwrap_or(1),
                                "w" => fp.width = val.parse::<f32>().ok().map(|v| v / 20.0),
                                "h" => fp.height = val.parse::<f32>().ok().map(|v| v / 20.0),
                                "hRule" => fp.height_rule = Some(val.to_string()),
                                "hAnchor" => fp.h_anchor = Some(val.to_string()),
                                "vAnchor" => fp.v_anchor = Some(val.to_string()),
                                "x" => fp.x = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                "y" => fp.y = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                "hSpace" => fp.h_space = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                "vSpace" => fp.v_space = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                "wrap" => fp.wrap = Some(val.to_string()),
                                "xAlign" => fp.x_align = Some(val.to_string()),
                                "yAlign" => fp.y_align = Some(val.to_string()),
                                _ => {}
                            }
                        }
                        style.frame_pr = Some(fp);
                    }
                    "snapToGrid" => {
                        // CT_OnOff: presence alone (no val) = true. A direct
                        // no-val `<w:snapToGrid/>` re-enables grid snap, overriding
                        // a style's snapToGrid=0 (S606b, Word COM-confirmed).
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0"
                                    && val.as_ref() != "false"
                                    && val.as_ref() != "off";
                            }
                        }
                        style.snap_to_grid = enabled;
                        has_explicit_snap_to_grid = true;
                    }
                    "contextualSpacing" => {
                        // w:contextualSpacing: presence alone means true,
                        // or explicit val="1"/"true"
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0" && val.as_ref() != "false";
                            }
                        }
                        style.contextual_spacing = enabled;
                        // S782: remember the DIRECT setting so the style merge
                        // can't clobber an explicit val="0" back to true.
                        style.has_explicit_contextual_spacing = true;
                    }
                    "spacing" => {
                        // S178 (2026-05-22): mirror the Start branch's
                        // has_direct_spacing assignment. Self-closing
                        // `<w:spacing .../>` (common form for cell pPr that
                        // only carries beforeLines/afterLines) was leaving
                        // has_direct_spacing=false → cell-render `should_reset`
                        // zeroed effective_space_before/after for the para
                        // even though OOXML had the spacing. 191cb 厚生労働大臣
                        // cpi=4: bl=Some(50), sb_dir=Some(6.8), pitch=Some(13.6)
                        // → esb=0 in dump (should be 6.8). Set unconditionally;
                        // S237 (2026-05-23): removed OXI_LEGACY_NO_EMPTY_SPACING_HDS
                        // legacy env-var fallback during hardening pass.
                        style.has_direct_spacing = true;
                        let mut line_val: Option<f32> = None;
                        let mut line_rule: Option<String> = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "before" => {
                                    style.space_before =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                    style.has_direct_before_after = true;
                                }
                                "after" => {
                                    style.space_after =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                    style.has_direct_before_after = true;
                                }
                                "beforeLines" => {
                                    style.before_lines = val.parse::<f32>().ok();
                                    style.has_direct_before_after = true;
                                }
                                "afterLines" => {
                                    style.after_lines = val.parse::<f32>().ok();
                                    style.has_direct_before_after = true;
                                }
                                "beforeAutospacing" => {
                                    // CT_OnOff attr: "1"/"true"/"on" => true (S675)
                                    let v = val.as_ref();
                                    style.before_autospacing = v == "1" || v == "true" || v == "on";
                                    style.has_direct_before_after = true;
                                }
                                "afterAutospacing" => {
                                    let v = val.as_ref();
                                    style.after_autospacing = v == "1" || v == "true" || v == "on";
                                    style.has_direct_before_after = true;
                                }
                                "line" => {
                                    line_val = val.parse::<f32>().ok();
                                }
                                "lineRule" => {
                                    line_rule = Some(val.to_string());
                                }
                                _ => {}
                            }
                        }
                        if let Some(lv) = line_val {
                            match line_rule.as_deref() {
                                Some("exact") => {
                                    // Exact: value in twips, convert to points
                                    style.line_spacing = Some(lv / 20.0);
                                    style.line_spacing_rule = Some("exact".to_string());
                                }
                                Some("atLeast") => {
                                    // At least: value in twips, convert to points
                                    style.line_spacing = Some(lv / 20.0);
                                    style.line_spacing_rule = Some("atLeast".to_string());
                                }
                                _ if lv < 0.0 => {
                                    // Word quirk (COM-confirmed 2026-05-15 on d4d126):
                                    // negative w:line with lineRule="auto" (or missing) is
                                    // treated as wdLineSpaceExactly with |val|/20 pt.
                                    style.line_spacing = Some(lv.abs() / 20.0);
                                    style.line_spacing_rule = Some("exact".to_string());
                                }
                                _ => {
                                    // Auto: proportional, divide by 240
                                    style.line_spacing = Some(lv / 240.0);
                                }
                            }
                        }
                    }
                    "ind" => {
                        // Track leftChars/rightChars for override logic
                        // Must be outside attr loop: leftChars overrides left regardless of XML attr order
                        let mut left_chars: Option<f32> = None;
                        let mut right_chars: Option<f32> = None;
                        let mut first_line_chars: Option<f32> = None;
                        let mut hanging_chars: Option<f32> = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "left" | "start" => {
                                    style.indent_left =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                "right" | "end" => {
                                    style.indent_right =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                "leftChars" | "startChars" => {
                                    // COM-confirmed: leftChars overrides left (not additive)
                                    // charWidth = 10.5pt (Japanese default)
                                    left_chars = val.parse::<f32>().ok();
                                }
                                "rightChars" | "endChars" => {
                                    right_chars = val.parse::<f32>().ok();
                                }
                                "firstLine" => {
                                    style.indent_first_line =
                                        val.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                "firstLineChars" => {
                                    first_line_chars = val.parse::<f32>().ok();
                                }
                                "hanging" => {
                                    // Hanging indent: negative first-line indent
                                    style.indent_first_line =
                                        val.parse::<f32>().ok().map(|v| -(v / 20.0));
                                }
                                "hangingChars" => {
                                    // Hanging indent in character units (hundredths)
                                    hanging_chars = val.parse::<f32>().ok();
                                }
                                _ => {}
                            }
                        }
                        // *Chars attributes: store raw values; resolved at layout time.
                        // Pre-2026-05-01: when hanging coexisted with leftChars, the
                        // skip rationale assumed an authoritative twip `left`. But for
                        // docx that have ONLY chars-based indent (no twip), skipping
                        // leftChars left subsequent-line wrap width ~1 char too wide,
                        // causing Oxi to fit 2 extra chars on line 1 vs Word
                        // (wrap_point_diff num_hang_chars test, 2026-05-01).
                        // Now: only skip when twip `left` is also present (twip wins).
                        let has_twip_left = style.indent_left.is_some();
                        if let Some(lc) = left_chars {
                            if !has_twip_left {
                                style.indent_left_chars = Some(lc);
                            }
                        }
                        if let Some(rc) = right_chars {
                            style.indent_right_chars = Some(rc);
                        }
                        // hangingChars overrides firstLineChars (negative = hanging)
                        if let Some(hc) = hanging_chars {
                            style.indent_first_line_chars = Some(-hc);
                        } else if let Some(fc) = first_line_chars {
                            style.indent_first_line_chars = Some(fc);
                        }
                    }
                    "shd" => {
                        // S705 (2026-06-30): paragraph-level shd → effective background
                        // colour (was: raw fill only, so pctN para shading mis-coloured).
                        // val="clear" → fill; "solid" → color; "pctN" → N% color over fill.
                        // Rendered as a full-line paragraph background in layout.
                        let mut shd_val = String::new();
                        let mut shd_fill = String::new();
                        let mut shd_color = String::new();
                        for attr in e.attributes().flatten() {
                            match local_name(attr.key.as_ref()).as_str() {
                                "val" => shd_val = String::from_utf8_lossy(&attr.value).to_string(),
                                "fill" => shd_fill = String::from_utf8_lossy(&attr.value).to_string(),
                                "color" => shd_color = String::from_utf8_lossy(&attr.value).to_string(),
                                _ => {}
                            }
                        }
                        style.shading = effective_shading_color(&shd_val, &shd_fill, &shd_color);
                    }
                    "pageBreakBefore" => {
                        // CT_OnOff: respect w:val="0"/"false"/"off" (S597). Direct
                        // paragraph-level <w:pageBreakBefore w:val="0"/> must NOT
                        // force a break (mirrors the style-level fix in styles.rs).
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0" && val.as_ref() != "false" && val.as_ref() != "off";
                            }
                        }
                        style.page_break_before = enabled;
                    }
                    "keepNext" => {
                        // CT_OnOff: presence alone = true; val="0"/"false"/"off"
                        // = false (S633, Word-confirmed: ailitguide/mysignaiguide
                        // headings ship `<w:keepNext w:val="0"/>` to disable keep).
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0"
                                    && val.as_ref() != "false"
                                    && val.as_ref() != "off";
                            }
                        }
                        style.keep_next = enabled;
                    }
                    "keepLines" => {
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0"
                                    && val.as_ref() != "false"
                                    && val.as_ref() != "off";
                            }
                        }
                        style.keep_lines = enabled;
                    }
                    "widowControl" => {
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0" && val.as_ref() != "false";
                            }
                        }
                        style.widow_control = enabled;
                        has_explicit_widow_control = true;
                    }
                    "bidi" => {
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0" && val.as_ref() != "false";
                            }
                        }
                        style.bidi = enabled;
                    }
                    "adjustRightInd" => {
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0" && val.as_ref() != "false";
                            }
                        }
                        style.adjust_right_ind = enabled;
                    }
                    "wordWrap" => {
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0" && val.as_ref() != "false";
                            }
                        }
                        style.word_wrap = enabled;
                    }
                    "autoSpaceDE" => {
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0" && val.as_ref() != "false";
                            }
                        }
                        style.auto_space_de = enabled;
                    }
                    "autoSpaceDN" => {
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                enabled = val.as_ref() != "0" && val.as_ref() != "false";
                            }
                        }
                        style.auto_space_dn = enabled;
                    }
                    "outlineLvl" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                // outlineLvl is for TOC, not layout font size
                                style.outline_level = val.parse::<u8>().ok();
                            }
                        }
                    }
                    "textAlignment" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                style.text_alignment = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "pPr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    style.has_explicit_widow_control = has_explicit_widow_control;
    style.has_explicit_snap_to_grid = has_explicit_snap_to_grid;
    Ok((style, alignment, style_id, num_pr, sect_pr, ppr_change, paragraph_mark_revision))
}

/// Parse w:numPr element
fn parse_num_pr(reader: &mut Reader<&[u8]>) -> Result<NumPrRef, ParseError> {
    let mut num_id = String::new();
    let mut ilvl: u8 = 0;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                if local_name(e.name().as_ref()) == "numberingChange" {
                    drain_element(reader, "numberingChange")?;
                    continue;
                }
                depth += 1;
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "ilvl" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                ilvl = val.parse().unwrap_or(0);
                            }
                        }
                    }
                    "numId" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                num_id = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "numPr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(NumPrRef { num_id, ilvl })
}

/// Parse w:pBdr element containing border children (top, bottom, left, right, between)
pub(crate) fn parse_paragraph_borders(reader: &mut Reader<&[u8]>) -> Result<ParagraphBorders, ParseError> {
    let mut borders = ParagraphBorders {
        top: None, bottom: None, left: None, right: None, between: None,
    };

    loop {
        match reader.read_event()? {
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                let bdr = parse_border_attrs(&e);
                match local.as_str() {
                    "top" => borders.top = bdr,
                    "bottom" => borders.bottom = bdr,
                    "left" | "start" => borders.left = bdr,
                    "right" | "end" => borders.right = bdr,
                    "between" => borders.between = bdr,
                    _ => {}
                }
            }
            Event::End(e) => {
                if local_name(e.name().as_ref()) == "pBdr" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(borders)
}

/// Parse border attributes from an element (w:sz, w:color, w:val)
fn parse_border_attrs(e: &quick_xml::events::BytesStart) -> Option<BorderDef> {
    let mut style = String::new();
    let mut width: f32 = 0.0;
    let mut color = None;
    let mut space: f32 = 0.0;

    for attr in e.attributes().flatten() {
        let key = local_name(attr.key.as_ref());
        let val = String::from_utf8_lossy(&attr.value).to_string();
        match key.as_str() {
            "val" => {
                if val == "none" || val == "nil" {
                    return None;
                }
                style = val;
            }
            "sz" => {
                width = val.parse::<f32>().unwrap_or(0.0) / 8.0;
            }
            "color" => {
                if val == "auto" {
                    color = Some("000000".to_string());
                } else {
                    color = Some(val);
                }
            }
            "space" => {
                space = val.parse::<f32>().unwrap_or(0.0);
            }
            _ => {}
        }
    }

    if style.is_empty() {
        return None;
    }

    Some(BorderDef { style, width, color, space })
}

/// Parse w:tabs element containing w:tab children
pub(crate) fn parse_tab_stops(reader: &mut Reader<&[u8]>) -> Result<Vec<TabStop>, ParseError> {
    let mut stops = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tab" {
                    let mut position: f32 = 0.0;
                    let mut alignment = TabStopAlignment::Left;
                    let mut leader: Option<String> = None;

                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value);
                        match key.as_str() {
                            "pos" => {
                                // Position in twips (1/20 pt)
                                position = val.parse::<f32>().unwrap_or(0.0) / 20.0;
                            }
                            "val" => {
                                alignment = match val.as_ref() {
                                    "center" => TabStopAlignment::Center,
                                    "right" | "end" => TabStopAlignment::Right,
                                    "decimal" => TabStopAlignment::Decimal,
                                    _ => TabStopAlignment::Left,
                                };
                            }
                            "leader" => {
                                leader = match val.as_ref() {
                                    "none" => None,
                                    _ => Some(val.to_string()),
                                };
                            }
                            _ => {}
                        }
                    }

                    stops.push(TabStop {
                        position,
                        alignment,
                        leader,
                    });
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tabs" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // Sort by position
    stops.sort_by(|a, b| a.position.partial_cmp(&b.position).unwrap_or(std::cmp::Ordering::Equal));
    Ok(stops)
}

/// Parse a w:r element (run). Returns the Run, optionally an Image, and field info.
/// `url` is set when this run is inside a w:hyperlink element.
fn parse_run(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet, url: Option<String>) -> Result<(Run, Option<DrawingResult>), ParseError> {
    let mut text = String::new();
    let mut style = RunStyle::default();
    let mut drawing_result: Option<DrawingResult> = None;
    let mut depth = 0;
    let mut in_text = false;
    let mut in_instr_text = false;
    let mut instr_text = String::new();
    let mut footnote_ref: Option<u32> = None;
    let mut endnote_ref: Option<u32> = None;
    let mut ruby: Option<Ruby> = None;
    let mut comment_references: Vec<String> = Vec::new();
    let mut rpr_change: Option<PropertyChange> = None;
    let mut has_last_rendered_page_break = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "rPr" if depth == 0 => {
                        let (s, c) = parse_run_properties(reader, ctx, styles)?;
                        style = s;
                        rpr_change = c;
                    }
                    "t" | "delText" if depth == 0 => {
                        in_text = true;
                    }
                    "instrText" if depth == 0 => {
                        in_instr_text = true;
                    }
                    "drawing" if depth == 0 => {
                        drawing_result = Some(parse_drawing(reader, ctx, styles)?);
                    }
                    // VML legacy picture/shape
                    "pict" if depth == 0 => {
                        let vml = parse_vml_pict(reader, ctx, styles)?;
                        if drawing_result.is_none() {
                            // S852: an inline horizontal rule (<v:rect o:hr="t"/>)
                            // is routed to the run style — its own line is reserved
                            // via inline_object_extent (the S851/S839 object
                            // mechanism) and the emit draws a full-width gray rule.
                            // A generic inline Shape (position None) is otherwise
                            // dropped by the layout (mod.rs shape loop only draws
                            // positioned shapes). forms only; JP byte-identical.
                            let is_hr_shape = vml.shape.as_ref()
                                .map_or(false, |s| s.shape_type == "hr");
                            if is_hr_shape && std::env::var("OXI_S852_DISABLE").is_err() {
                                let sh = vml.shape.as_ref().unwrap();
                                let hr_w = if sh.width > 1.0 { sh.width } else { 468.0 };
                                let thickness = sh.height.max(0.75);
                                let color = sh.fill.clone()
                                    .filter(|c| c.chars().all(|ch| ch.is_ascii_hexdigit()) && c.len() == 6)
                                    .unwrap_or_else(|| "A6A6A6".to_string());
                                // The rule occupies its own line ~= a default text
                                // line (rule + vertical margins). 13.8pt matches
                                // Word's o:hr line box (measured on forms' title HR).
                                style.inline_object_extent = Some((hr_w, 13.8));
                                style.hr_rule = Some((thickness, color));
                            } else {
                                drawing_result = Some(vml);
                            }
                        }
                    }
                    // mc:AlternateContent — prefer Choice (DrawingML)
                    "AlternateContent" if depth == 0 => {
                        let ac = parse_alternate_content(reader, ctx, styles)?;
                        if drawing_result.is_none() {
                            drawing_result = ac;
                        }
                    }
                    // OLE object — extract preview image from embedded VML shape
                    "object" if depth == 0 => {
                        let (ole, saw_ole) = parse_ole_object(reader, ctx)?;
                        if drawing_result.is_none() {
                            // S851 (2026-07-14, default ON, opt-out OXI_S851_DISABLE):
                            // an inline OLEObject-LESS <w:object> (a bare form-field
                            // picture shape, e.g. the MassHealth PA-form field
                            // underlines "Last name [___] First name [___]") flows
                            // INLINE in its host line in Word. Route it as a run-level
                            // inline object (inline_object_extent + inline_object_image)
                            // so break_into_lines makes a width-bearing U+FFFC fragment
                            // and the emit draws the image there — NOT extracted to a
                            // separate Block::Image, which stacked each field on its own
                            // ~18pt line (forms__00042714 over-reserved ~54pt/row, 8pg
                            // vs Word 6pg). Real OLE (Equation.3 / Visio, saw_ole) keeps
                            // the block path → the JP canaries 3a4f/model/tokyoshugyo
                            // (Equation) + uklocalspending (Visio) are byte-identical.
                            let s851_ole_less = std::env::var("OXI_S851_DISABLE").is_err()
                                && !saw_ole
                                && ole.shape.is_none() && ole.text_box.is_none()
                                && ole.image.as_ref().map_or(false, |i|
                                    i.position.is_none() && i.width > 0.0 && i.height > 0.0);
                            if s851_ole_less {
                                let img = ole.image.unwrap();
                                style.inline_object_extent = Some((img.width, img.height));
                                style.inline_object_image = Some(Box::new(img));
                            } else {
                                drawing_result = Some(ole);
                            }
                        }
                    }
                    "ruby" if depth == 0 => {
                        ruby = Some(parse_ruby(reader)?);
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::Text(e) => {
                let content = e.unescape().unwrap_or_default();
                if in_text {
                    // S747 (2026-07-05): U+00AD SOFT HYPHEN is an invisible
                    // hyphenation OPPORTUNITY — Word renders it zero-width
                    // (a "-" appears only when a line actually breaks there).
                    // Oxi measured/drew it as a visible hyphen-width char ->
                    // phantom width shifted wraps (probeqbrkchars +1). Strip it
                    // from the run text (v1: no break-at-SHY hyphenation —
                    // rendering the hyphen on break needs the hyphenation
                    // feature; zero-width is the dominant fidelity term).
                    // Opt-out OXI_S747_DISABLE.
                    if content.contains('\u{ad}') && std::env::var("OXI_S747_DISABLE").is_err() {
                        text.push_str(&content.replace('\u{ad}', ""));
                    } else {
                        text.push_str(&content);
                    }
                } else if in_instr_text {
                    instr_text.push_str(&content);
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "t" || local == "delText" {
                    in_text = false;
                    // S848 (2026-07-14, opt-out OXI_S848_DISABLE): a LITERAL
                    // trailing newline inside `<w:t>` content
                    // (`<w:t xml:space="preserve">text\n</w:t>`, common in
                    // HTML-export docx) is collapsed by Word — the paragraph
                    // mark is the real break. Oxi's soft-break handling
                    // rendered it as a phantom EMPTY 2nd line (+~one line
                    // height per paragraph): a bulleted list where every item
                    // carries a trailing \n doubled each item's advance (30pt
                    // vs Word 15pt), over-flowing the page (docx-corpus reports
                    // doc 2pg->1pg=Word). Strip it HERE (at </w:t>) so only the
                    // `<w:t>` text content is trimmed — a `<w:br/>` pushes its
                    // \n AFTER </w:t>, so genuine wrap breaks are preserved
                    // (the first cut, which stripped the merged run text, ate
                    // trailing `<w:br/>` breaks — creative doc regressed).
                    if local == "t" && std::env::var("OXI_S848_DISABLE").is_err() {
                        while text.ends_with('\n') || text.ends_with('\r') {
                            text.pop();
                        }
                    }
                } else if local == "instrText" {
                    in_instr_text = false;
                } else if local == "r" && depth == 0 {
                    break;
                } else if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "br" => {
                        let mut br_type = None;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "type" {
                                br_type = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                        match br_type.as_deref() {
                            Some("page") => text.push('\x0C'),   // form feed = page break
                            Some("column") => text.push('\x0B'), // vertical tab = column break
                            _ => text.push('\n'),                 // text wrap break
                        }
                    }
                    "tab" => text.push('\t'),
                    // S747: <w:noBreakHyphen/> renders as a hyphen that FORBIDS
                    // a line break — map to U+2011 NON-BREAKING HYPHEN (the
                    // breaker's is_break_after covers '-' only, so U+2011 adds
                    // no break opportunity). Previously unparsed -> the visible
                    // hyphen was silently DROPPED.
                    "noBreakHyphen" => {
                        if std::env::var("OXI_S747_DISABLE").is_err() {
                            text.push('\u{2011}');
                        }
                    }
                    "lastRenderedPageBreak" => {
                        has_last_rendered_page_break = true;
                    }
                    "fldChar" => {
                        // Track field state via marker characters
                        // (parsed by parent parse_paragraph to manage field_result_depth)
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "fldCharType" {
                                let val = String::from_utf8_lossy(&attr.value);
                                match val.as_ref() {
                                    "separate" => text.push('\u{FFFE}'),
                                    "end" => text.push('\u{FFFF}'),
                                    _ => {} // "begin" — no marker needed
                                }
                            }
                        }
                    }
                    "footnoteReference" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "id" {
                                let val = String::from_utf8_lossy(&attr.value);
                                if let Ok(id) = val.parse::<u32>() {
                                    if id > 0 { // Skip separator/continuation notes (id=0)
                                        footnote_ref = Some(id);
                                        // Word renders just the number (e.g. "1"),
                                        // not "[1]". renumber_note_refs in parse_body
                                        // will rewrite this to the section-local seq.
                                        text = format!("{}", id);
                                    }
                                }
                            }
                        }
                    }
                    "endnoteReference" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "id" {
                                let val = String::from_utf8_lossy(&attr.value);
                                if let Ok(id) = val.parse::<u32>() {
                                    if id > 0 {
                                        endnote_ref = Some(id);
                                        text = format!("[{}]", id);
                                    }
                                }
                            }
                        }
                    }
                    // Comment balloon anchor — zero-width marker inside the run.
                    // The enclosing run is what the renderer projects to the right
                    // margin; one run may legally carry multiple references.
                    "commentReference" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "id" {
                                let id = String::from_utf8_lossy(&attr.value).to_string();
                                if !id.is_empty() {
                                    comment_references.push(id);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // If this run contains a field instruction, convert to placeholder text
    let mut field_type: Option<FieldType> = None;
    if !instr_text.is_empty() {
        let field = instr_text.trim();
        if field.contains("PAGE") && !field.contains("NUMPAGES") && !field.contains("PAGEREF") {
            text = "#".to_string();
            field_type = Some(FieldType::Page);
        } else if field.contains("NUMPAGES") || field.contains("SECTIONPAGES") {
            text = "#".to_string();
            field_type = Some(FieldType::NumPages);
        } else if field.contains("TOC") || field.contains("HYPERLINK") {
            // Table of contents / hyperlink fields — Oxi can't regenerate them, so
            // KEEP the cached result (Word's rendered ToC / link text). Marking the
            // field Cached stops the result-suppression pass from dropping the FIRST
            // ToC entry title, which sits inside the TOC field result region BEFORE
            // any nested PAGEREF sets current_field_type (later entries survived only
            // because current_field_type stayed stuck at CrossRef).
            field_type = Some(FieldType::Cached);
        } else if field.contains("REF") || field.contains("NOTEREF") || field.contains("PAGEREF") {
            // S685 (2026-06-28, default ON, opt-out OXI_CROSSREF_DISABLE): cross-reference
            // fields render the CACHED RESULT (the run between fldChar separate and end),
            // NOT a "#" placeholder. The instruction run itself has no display text (leave
            // empty); FieldType::CrossRef tells parse_paragraph to KEEP (not suppress) the
            // cached result run. Word shows «第１９条»; the old "#" dropped chars → shifted
            // wrapping doc-wide (tokyoshugyo 0.9759→0.9792, deltas reduced).
            if std::env::var("OXI_CROSSREF_DISABLE").is_err() {
                field_type = Some(FieldType::CrossRef);
            } else if text.is_empty() {
                text = "#".to_string();
            }
        } else if field.contains("DATE") || field.contains("TIME")
            || field.contains("AUTHOR") || field.contains("TITLE") || field.contains("SUBJECT")
            || field.contains("FILENAME") || field.contains("DOCPROPERTY")
            || field.contains("USERNAME") || field.contains("LASTSAVEDBY")
            || field.contains("COMMENTS") || field.contains("KEYWORDS")
            || field.contains("STYLEREF")
        {
            // S708 (2026-06-30, default ON, opt-out OXI_FIELDCACHE_DISABLE): fields whose
            // value Word stores as a CACHED RESULT (the run between fldChar separate and
            // end) and re-displays on open — DATE/TIME/CREATEDATE/SAVEDATE/AUTHOR/TITLE/
            // FILENAME/… Oxi can't re-evaluate them, so the cache is the only source.
            // Like CrossRef: KEEP the cached result, leave the instruction run empty.
            // The old code showed the raw instruction («DATE \@ "yyyy/MM/dd"») or a
            // «[AUTHOR]» placeholder AND dropped the cache → garbage text + shifted wrap.
            if std::env::var("OXI_FIELDCACHE_DISABLE").is_err() {
                field_type = Some(FieldType::Cached);
            } else if text.is_empty() {
                // legacy A/B path: show field name placeholder (old AUTHOR behaviour)
                text = format!("[{}]", field.split_whitespace().next().unwrap_or(field));
            }
        }
    }

    // If ruby was parsed, use its base text as the run text
    if let Some(ref r) = ruby {
        if text.is_empty() {
            text = r.base.clone();
        }
    }

    Ok((Run {
        text, style, url, footnote_ref, endnote_ref,
        comment_range_start: Vec::new(),
        comment_range_end: Vec::new(),
        comment_references,
        tracked_change: None,
        rpr_change,
        ruby,
        bookmark_name: None,
        is_math: false,
        field_type,
        has_last_rendered_page_break,
    }, drawing_result))
}

/// Parse runs inside a w:hyperlink element
fn parse_hyperlink_runs(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet, url: Option<String>) -> Result<Vec<Run>, ParseError> {
    let mut runs = Vec::new();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "r" && depth == 0 {
                    let (run, _dr) = parse_run(reader, ctx, styles, url.clone())?;
                    runs.push(run);
                } else {
                    depth += 1;
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "hyperlink" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(runs)
}

/// Parse word/footnotes.xml or word/endnotes.xml into a map of id -> blocks
fn parse_notes_xml(xml: &str, styles: &StyleSheet) -> Result<HashMap<String, Vec<Block>>, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut notes: HashMap<String, Vec<Block>> = HashMap::new();
    let mut current_id: Option<String> = None;
    let mut current_type: Option<String> = None;
    let mut current_blocks: Vec<Block> = Vec::new();
    let mut depth = 0;
    let mut in_note = false;

    // Create a minimal context for parsing (no media/hyperlinks in notes)
    let note_ctx = ParseContext {
        _rels: HashMap::new(),
        media: HashMap::new(),
        media_types: HashMap::new(),
        hyperlinks: HashMap::new(),
        numbering: super::numbering::NumberingDefinitions::default(),
        list_counters: std::cell::RefCell::new(HashMap::new()),
        footnotes: HashMap::new(),
        endnotes: HashMap::new(),
        comments: HashMap::new(),
        theme: ThemeColors::default(),
    };

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "footnote" | "endnote" if !in_note => {
                        in_note = true;
                        depth = 0;
                        current_blocks.clear();
                        current_id = None;
                        current_type = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "id" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                current_id = Some(val);
                            }
                            // S833: capture the special-footnote type so the
                            // separator / continuationNotice paragraphs are
                            // available for the styled-height reservation.
                            if key == "type" {
                                current_type = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "p" if in_note && depth == 0 => {
                        let pr = parse_paragraph(&mut reader, &note_ctx, styles, false)?;
                        let para = pr.paragraph;
                        current_blocks.push(Block::Paragraph(para));
                    }
                    _ if in_note => {
                        depth += 1;
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "footnote" | "endnote" if in_note && depth == 0 => {
                        let typ = current_type.take();
                        if let Some(id) = current_id.take() {
                            // S833: keep the SPECIAL paragraphs under sentinel
                            // keys — their styled heights drive the declared-
                            // separator reservation model.
                            match typ.as_deref() {
                                Some("separator") => {
                                    notes.insert("__sep__".to_string(), std::mem::take(&mut current_blocks));
                                }
                                Some("continuationNotice") => {
                                    notes.insert("__notice__".to_string(), std::mem::take(&mut current_blocks));
                                }
                                Some("continuationSeparator") => { current_blocks.clear(); }
                                _ => {
                                    // Skip separator notes (id 0 and -1)
                                    if id != "0" && id != "-1" {
                                        notes.insert(id, std::mem::take(&mut current_blocks));
                                    }
                                }
                            }
                        }
                        in_note = false;
                    }
                    _ if in_note && depth > 0 => {
                        depth -= 1;
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(notes)
}

/// Result from parsing a w:drawing element — may contain image, shape, and/or text box
struct DrawingResult {
    image: Option<Image>,
    shape: Option<Shape>,
    text_box: Option<TextBox>,
}

impl DrawingResult {
    /// Returns true if at least one component (image, shape, or text_box) is present
    fn has_content(&self) -> bool {
        self.image.is_some() || self.shape.is_some() || self.text_box.is_some()
    }
}

/// Parse DrawingML color modifier child elements (lumMod, lumOff, tint, shade, etc.)
/// and apply them to a base hex color. Consumes elements until the closing tag.
fn parse_color_modifiers(reader: &mut Reader<&[u8]>, base_hex: &str, end_tag: &str) -> String {
    let mut lum_mod: Option<f32> = None;
    let mut lum_off: Option<f32> = None;
    let mut tint: Option<f32> = None;
    let mut shade: Option<f32> = None;
    let mut sat_mod: Option<f32> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Empty(e)) => {
                let local = local_name(e.name().as_ref());
                for attr in e.attributes().flatten() {
                    if local_name(attr.key.as_ref()) == "val" {
                        let val = String::from_utf8_lossy(&attr.value);
                        let v = val.parse::<f32>().unwrap_or(0.0) / 100000.0;
                        match local.as_str() {
                            "lumMod" => lum_mod = Some(v),
                            "lumOff" => lum_off = Some(v),
                            "tint" => tint = Some(v),
                            "shade" => shade = Some(v),
                            "satMod" => sat_mod = Some(v),
                            "alpha" => {} // Recognized but not yet applied (opacity)
                            _ => {}
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                if local_name(e.name().as_ref()) == end_tag {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            _ => {}
        }
    }

    // Apply modifiers to base color
    let r = u8::from_str_radix(&base_hex[0..2], 16).unwrap_or(0) as f32 / 255.0;
    let g = u8::from_str_radix(&base_hex[2..4], 16).unwrap_or(0) as f32 / 255.0;
    let b = u8::from_str_radix(&base_hex[4..6], 16).unwrap_or(0) as f32 / 255.0;

    let (mut r, mut g, mut b) = (r, g, b);

    // Apply tint (move towards white)
    if let Some(t) = tint {
        r = r * t + (1.0 - t);
        g = g * t + (1.0 - t);
        b = b * t + (1.0 - t);
    }

    // Apply shade (move towards black)
    if let Some(s) = shade {
        r *= s;
        g *= s;
        b *= s;
    }

    // Apply HLS modifiers: lumMod/lumOff on L, satMod on S (matches Word output)
    if lum_mod.is_some() || lum_off.is_some() || sat_mod.is_some() {
        // RGB -> HSL
        let max_c = r.max(g).max(b);
        let min_c = r.min(g).min(b);
        let d = max_c - min_c;
        let l_orig = (max_c + min_c) / 2.0;
        let mut l = l_orig;
        // Apply lumMod/lumOff to L
        if lum_mod.is_some() || lum_off.is_some() {
            let m = lum_mod.unwrap_or(1.0);
            let o = lum_off.unwrap_or(0.0);
            l = (l * m + o).clamp(0.0, 1.0);
        }
        if d.abs() < 1e-6 {
            // Achromatic — satMod has no effect
            r = l; g = l; b = l;
        } else {
            let s_orig = if l_orig > 0.5 { d / (2.0 - max_c - min_c).max(1e-6) } else { d / (max_c + min_c).max(1e-6) };
            let h = if (max_c - r).abs() < 1e-6 {
                (g - b) / d + if g < b { 6.0 } else { 0.0 }
            } else if (max_c - g).abs() < 1e-6 {
                (b - r) / d + 2.0
            } else {
                (r - g) / d + 4.0
            } / 6.0;
            // Apply satMod to S
            let s = if let Some(sm) = sat_mod {
                (s_orig * sm).clamp(0.0, 1.0)
            } else {
                s_orig
            };
            // HSL -> RGB with modified L and S, original H
            let hue_to_rgb = |p: f32, q: f32, mut t: f32| -> f32 {
                if t < 0.0 { t += 1.0; }
                if t > 1.0 { t -= 1.0; }
                if t < 1.0/6.0 { return p + (q - p) * 6.0 * t; }
                if t < 1.0/2.0 { return q; }
                if t < 2.0/3.0 { return p + (q - p) * (2.0/3.0 - t) * 6.0; }
                p
            };
            let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
            let p = 2.0 * l - q;
            r = hue_to_rgb(p, q, h + 1.0/3.0);
            g = hue_to_rgb(p, q, h);
            b = hue_to_rgb(p, q, h - 1.0/3.0);
        }
    }

    format!("{:02X}{:02X}{:02X}",
        (r * 255.0).round() as u8,
        (g * 255.0).round() as u8,
        (b * 255.0).round() as u8,
    )
}

/// Parse a w:drawing element to extract image, shape, or text box info
fn parse_drawing(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<DrawingResult, ParseError> {
    let mut width: f32 = 0.0;
    let mut height: f32 = 0.0;
    let mut alt_text = None;
    let mut rel_id = None;
    let mut depth = 0;
    // Floating image info
    let mut is_anchor = false;
    let mut dist_l: Option<f32> = None;
    let mut dist_r: Option<f32> = None;
    // S478: wp:anchor z-order. Word draws floating objects in ascending
    // relativeHeight (highest = on top). behindDoc=1 places the object
    // behind body text. Default 0 (in front, ordered by relativeHeight).
    let mut relative_height: u32 = 0;
    let mut behind_doc: bool = false;
    let mut pos_x: f32 = 0.0;
    let mut pos_y: f32 = 0.0;
    let mut h_relative: Option<String> = None;
    let mut v_relative: Option<String> = None;
    let mut h_align: Option<String> = None;
    let mut v_align: Option<String> = None;
    let mut in_align = false;
    let mut wrap_type: Option<WrapType> = None;
    let mut crop: Option<ImageCrop> = None;
    let mut in_pos_h = false;
    let mut in_pos_v = false;
    let mut in_pos_offset = false;
    // Shape properties
    let mut shape_type: Option<String> = None;
    let mut shape_fill: Option<String> = None;
    let mut stroke_color: Option<String> = None;
    let mut stroke_width: Option<f32> = None;
    let mut shape_text_blocks: Vec<Block> = Vec::new();
    let mut rotation: Option<f32> = None;
    let mut flip_h = false;
    let mut flip_v = false;
    // S493i: connector arrowheads (a:ln/a:headEnd|a:tailEnd type≠"none"). head=start, tail=end.
    let mut arrow_head = false;
    let mut arrow_tail = false;
    let mut has_no_fill = false;
    let mut has_no_stroke = false;
    let mut corner_radius_adj: Option<f32> = None; // avLst adj value (0-100000 scale)
    let mut gradient_stops: Vec<GradientStop> = Vec::new();
    let mut gradient_angle: Option<f32> = None;
    // Text body insets (from bodyPr lIns/rIns/tIns/bIns, in EMU)
    let mut text_inset_left: Option<f32> = None;
    let mut text_inset_right: Option<f32> = None;
    let mut text_inset_top: Option<f32> = None;
    let mut text_inset_bottom: Option<f32> = None;
    let mut text_body_anchor: Option<String> = None;
    // S481: bodyPr@vertOverflow ("overflow" default / "clip" / "ellipsis").
    let mut text_vert_overflow: Option<String> = None;
    // S662: bodyPr@compatLnSpc="1" (legacy "compatible line spacing").
    let mut text_compat_ln_spc = false;
    // S537b: wordprocessingCanvas marker (wpc:wpc child of graphicData).
    let mut is_canvas = false;
    // S839 (2026-07-14): wpg vector-group per-shape extraction. The main
    // loop keeps ALL its existing (last-win, global) prop semantics; the
    // S839 state only PIGGYBACKS per-shape records so a visual-only group
    // (hmrc's checkbox strips / writing boxes / heavy rules) yields
    // drawable primitives. Text-bearing groups (framework cover pages)
    // parse exactly as before and get NO vector_shapes attached.
    struct S839Wsp {
        off: (f32, f32),
        ext: (f32, f32),
        prst: Option<String>,
        pts: Vec<(f32, f32)>,
        path_wh: Option<(f32, f32)>,
        n_paths: u32,
        has_curve: bool,
        ln_w: Option<f32>,
        ln_color: Option<String>,
        ln_nofill: bool,
        fill: Option<String>,
        no_fill: bool,
        lnref: (i32, Option<String>),
        fillref: (i32, Option<String>),
    }
    let mut s839_vector_shapes: Vec<crate::ir::VectorShape> = Vec::new();
    let mut s839_in_wgp = false;
    let mut s839_in_grpsppr = false;
    // group xfrm: off, ext, chOff, chExt (EMU)
    let mut s839_grp: ((f32, f32), (f32, f32), (f32, f32), (f32, f32)) =
        ((0.0, 0.0), (0.0, 0.0), (0.0, 0.0), (1.0, 1.0));
    let mut s839_grp_seen = false;
    let mut s839_wsp: Option<S839Wsp> = None;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                depth += 1;
                match local.as_str() {
                    "wpc" => {
                        is_canvas = true;
                    }
                    // S839: wpg vector-group boundaries + per-shape state.
                    // These arms only toggle piggyback state (the events fell
                    // to the default `_ => {}` before) — every existing global
                    // extraction arm still runs unchanged.
                    "wgp" => { s839_in_wgp = true; }
                    "grpSpPr" if s839_in_wgp => { s839_in_grpsppr = true; }
                    "wsp" if s839_in_wgp => {
                        s839_wsp = Some(S839Wsp {
                            off: (0.0, 0.0), ext: (0.0, 0.0), prst: None,
                            pts: Vec::new(), path_wh: None, n_paths: 0,
                            has_curve: false, ln_w: None, ln_color: None,
                            ln_nofill: false, fill: None, no_fill: false,
                            lnref: (-1, None), fillref: (-1, None),
                        });
                    }
                    "path" if s839_wsp.is_some() => {
                        let w = s839_wsp.as_mut().unwrap();
                        w.n_paths += 1;
                        let (mut pw, mut ph) = (0.0f32, 0.0f32);
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            if key == "w" { pw = val.parse().unwrap_or(0.0); }
                            else if key == "h" { ph = val.parse().unwrap_or(0.0); }
                        }
                        if w.path_wh.is_none() { w.path_wh = Some((pw, ph)); }
                    }
                    "cubicBezTo" | "quadBezTo" | "arcTo" if s839_wsp.is_some() => {
                        s839_wsp.as_mut().unwrap().has_curve = true;
                    }
                    "anchor" => {
                        is_anchor = true;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "relativeHeight" => { relative_height = val.parse::<u32>().unwrap_or(0); }
                                "behindDoc" => { behind_doc = val == "1" || val == "true"; }
                                "distL" => { dist_l = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "distR" => { dist_r = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                _ => {}
                            }
                        }
                    }
                    "docPr" => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "descr" {
                                alt_text = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    // Shape line as Start element — may contain srgbClr child
                    "ln" => {
                        // S839: per-shape capture of THIS a:ln's values (the
                        // globals stay last-win as before).
                        let mut s839_ln_color: Option<String> = None;
                        let mut s839_ln_nofill = false;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "w" {
                                let val = String::from_utf8_lossy(&attr.value);
                                stroke_width = val.parse::<f32>().ok().map(|v| v / 12700.0);
                                if let Some(wsp) = s839_wsp.as_mut() { wsp.ln_w = stroke_width; }
                            }
                        }
                        // Parse children for stroke color
                        let mut ln_depth = 1;
                        loop {
                            match reader.read_event() {
                                Ok(Event::Start(se)) => {
                                    let sl = local_name(se.name().as_ref());
                                    match sl.as_str() {
                                        "schemeClr" => {
                                            let mut base_hex: Option<String> = None;
                                            for attr in se.attributes().flatten() {
                                                if local_name(attr.key.as_ref()) == "val" {
                                                    let val = String::from_utf8_lossy(&attr.value).to_string();
                                                    base_hex = ctx.theme.resolve(&val).cloned();
                                                }
                                            }
                                            if let Some(hex) = base_hex {
                                                stroke_color = Some(parse_color_modifiers(reader, &hex, "schemeClr"));
                                            }
                                            // parse_color_modifiers consumed End tag, no depth change
                                        }
                                        "srgbClr" => {
                                            let mut hex = String::new();
                                            for attr in se.attributes().flatten() {
                                                if local_name(attr.key.as_ref()) == "val" {
                                                    hex = String::from_utf8_lossy(&attr.value).to_string();
                                                }
                                            }
                                            if !hex.is_empty() {
                                                let c = parse_color_modifiers(reader, &hex, "srgbClr");
                                                s839_ln_color = Some(c.clone());
                                                stroke_color = Some(c);
                                            }
                                        }
                                        "sysClr" => {
                                            let mut hex = String::new();
                                            for attr in se.attributes().flatten() {
                                                if local_name(attr.key.as_ref()) == "lastClr" {
                                                    hex = String::from_utf8_lossy(&attr.value).to_string();
                                                }
                                            }
                                            if !hex.is_empty() {
                                                let c = parse_color_modifiers(reader, &hex, "sysClr");
                                                s839_ln_color = Some(c.clone());
                                                stroke_color = Some(c);
                                            }
                                        }
                                        _ => ln_depth += 1,
                                    }
                                }
                                Ok(Event::Empty(se)) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "srgbClr" {
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "val" {
                                                let c = String::from_utf8_lossy(&attr.value).to_string();
                                                s839_ln_color = Some(c.clone());
                                                stroke_color = Some(c);
                                            }
                                        }
                                    } else if sl == "sysClr" {
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "lastClr" {
                                                let c = String::from_utf8_lossy(&attr.value).to_string();
                                                s839_ln_color = Some(c.clone());
                                                stroke_color = Some(c);
                                            }
                                        }
                                    } else if sl == "schemeClr" {
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "val" {
                                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                                if let Some(resolved) = ctx.theme.resolve(&val) {
                                                    s839_ln_color = Some(resolved.clone());
                                                    stroke_color = Some(resolved.clone());
                                                }
                                            }
                                        }
                                    } else if sl == "noFill" {
                                        s839_ln_nofill = true;
                                        has_no_stroke = true;
                                    } else if sl == "headEnd" || sl == "tailEnd" {
                                        // S493i: connector arrowhead (type≠none). head=start, tail=end.
                                        let mut has = false;
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "type" {
                                                let v = String::from_utf8_lossy(&attr.value);
                                                has = !v.is_empty() && v != "none";
                                            }
                                        }
                                        if sl == "headEnd" { arrow_head = has; } else { arrow_tail = has; }
                                    }
                                }
                                Ok(Event::End(_se)) => {
                                    ln_depth -= 1;
                                    if ln_depth == 0 {
                                        break;
                                    }
                                }
                                Ok(Event::Eof) => break,
                                _ => {}
                            }
                        }
                        if let Some(wsp) = s839_wsp.as_mut() {
                            if s839_ln_color.is_some() { wsp.ln_color = s839_ln_color; }
                            wsp.ln_nofill |= s839_ln_nofill;
                        }
                        // We consumed ln's End, decrement outer depth
                        depth -= 1;
                    }
                    // Gradient fill
                    "gradFill" => {
                        let mut gf_depth = 1;
                        let mut current_gs_pos: Option<f32> = None;
                        loop {
                            match reader.read_event() {
                                Ok(Event::Start(se)) => {
                                    let sl = local_name(se.name().as_ref());
                                    gf_depth += 1;
                                    if sl == "gs" {
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "pos" {
                                                let val = String::from_utf8_lossy(&attr.value);
                                                current_gs_pos = val.parse::<f32>().ok().map(|v| v / 1000.0);
                                            }
                                        }
                                    } else if sl == "lin" {
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "ang" {
                                                let val = String::from_utf8_lossy(&attr.value);
                                                gradient_angle = val.parse::<f32>().ok().map(|v| v / 60000.0);
                                            }
                                        }
                                    }
                                }
                                Ok(Event::Empty(se)) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "srgbClr" {
                                        if let Some(pos) = current_gs_pos {
                                            for attr in se.attributes().flatten() {
                                                if local_name(attr.key.as_ref()) == "val" {
                                                    gradient_stops.push(GradientStop {
                                                        position: pos,
                                                        color: String::from_utf8_lossy(&attr.value).to_string(),
                                                    });
                                                }
                                            }
                                        }
                                    } else if sl == "lin" {
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "ang" {
                                                let val = String::from_utf8_lossy(&attr.value);
                                                gradient_angle = val.parse::<f32>().ok().map(|v| v / 60000.0);
                                            }
                                        }
                                    } else if sl == "gs" {
                                        // Empty gs element — unlikely but handle gracefully
                                    }
                                }
                                Ok(Event::End(se)) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "gs" {
                                        current_gs_pos = None;
                                    }
                                    gf_depth -= 1;
                                    if gf_depth == 0 { break; }
                                }
                                Ok(Event::Eof) => break,
                                _ => {}
                            }
                        }
                        depth -= 1; // consumed gradFill End
                    }
                    // schemeClr with child modifiers (lumMod, lumOff, tint, shade)
                    "schemeClr" => {
                        let mut base_hex: Option<String> = None;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                base_hex = ctx.theme.resolve(&val).cloned();
                            }
                        }
                        if let Some(hex) = base_hex {
                            let final_color = parse_color_modifiers(reader, &hex, "schemeClr");
                            if shape_fill.is_none() && !has_no_fill {
                                shape_fill = Some(final_color);
                            }
                        }
                        depth -= 1; // consumed schemeClr End
                    }
                    // srgbClr with child modifiers (alpha, tint, shade, lumMod, etc.)
                    "srgbClr" => {
                        let mut hex = String::new();
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                hex = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        if !hex.is_empty() {
                            let final_color = parse_color_modifiers(reader, &hex, "srgbClr");
                            if let Some(wsp) = s839_wsp.as_mut() {
                                if wsp.fill.is_none() && !wsp.no_fill { wsp.fill = Some(final_color.clone()); }
                            }
                            if shape_fill.is_none() && !has_no_fill {
                                shape_fill = Some(final_color);
                            }
                        }
                        depth -= 1;
                    }
                    // sysClr with child modifiers
                    "sysClr" => {
                        let mut hex = String::new();
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "lastClr" {
                                hex = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        if !hex.is_empty() {
                            let final_color = parse_color_modifiers(reader, &hex, "sysClr");
                            if shape_fill.is_none() && !has_no_fill {
                                shape_fill = Some(final_color);
                            }
                        }
                        depth -= 1;
                    }
                    // Shape transform rotation (xfrm rot attribute in 60000ths of a degree)
                    "xfrm" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "rot" => { rotation = val.parse::<f32>().ok().map(|v| v / 60000.0); }
                                // flipH/flipV: connector diagonal direction (S493h).
                                "flipH" => { flip_h = val == "1" || val == "true"; }
                                "flipV" => { flip_v = val == "1" || val == "true"; }
                                _ => {}
                            }
                        }
                    }
                    "positionH" => {
                        in_pos_h = true;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "relativeFrom" {
                                h_relative = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "positionV" => {
                        in_pos_v = true;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "relativeFrom" {
                                v_relative = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "posOffset" => {
                        in_pos_offset = true;
                    }
                    "align" => {
                        in_align = true;
                    }
                    // Text body properties — text insets (bodyPr as start element with children)
                    "bodyPr" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "lIns" => { text_inset_left = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "rIns" => { text_inset_right = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "tIns" => { text_inset_top = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "bIns" => { text_inset_bottom = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "anchor" => { text_body_anchor = Some(val.to_string()); }
                                "vertOverflow" => { text_vert_overflow = Some(val.to_string()); }
                                "compatLnSpc" => { text_compat_ln_spc = val == "1" || val == "true"; }
                                _ => {}
                            }
                        }
                    }
                    // DrawingML shape text content
                    "txbxContent" => {
                        // Parse text blocks inside shape
                        loop {
                            match reader.read_event() {
                                Ok(Event::Start(se)) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "p" {
                                        if let Ok(pr) = parse_paragraph(reader, ctx, styles, false) {
                                            shape_text_blocks.push(Block::Paragraph(pr.paragraph));
                                        }
                                    }
                                }
                                Ok(Event::End(se)) => {
                                    if local_name(se.name().as_ref()) == "txbxContent" {
                                        break;
                                    }
                                }
                                Ok(Event::Eof) => break,
                                _ => {}
                            }
                        }
                        // We consumed the txbxContent end tag, so decrement depth
                        depth -= 1;
                    }
                    // a:blip as Start element (has child elements like a:extLst)
                    "blip" => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:embed" || key.ends_with(":embed") || key == "embed" {
                                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    // Shape preset geometry as Start element (has avLst child)
                    "prstGeom" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "prst" {
                                let v = String::from_utf8_lossy(&attr.value).to_string();
                                if let Some(wsp) = s839_wsp.as_mut() { wsp.prst = Some(v.clone()); }
                                shape_type = Some(v);
                            }
                        }
                    }
                    // wps:style — parse fillRef/lnRef properly.
                    // fillRef idx="0" means NO fill (ignore child color).
                    // lnRef idx="0" means NO stroke (ignore child color).
                    // idx > 0 means use the child color as fill/stroke.
                    "style" => {
                        let mut style_depth = 1u32;
                        let mut in_fill_ref = false;
                        let mut fill_ref_idx: i32 = -1;
                        let mut in_ln_ref = false;
                        let mut ln_ref_idx: i32 = -1;
                        loop {
                            match reader.read_event() {
                                Ok(Event::Start(se)) => {
                                    let sl = local_name(se.name().as_ref());
                                    style_depth += 1;
                                    match sl.as_str() {
                                        "fillRef" => {
                                            in_fill_ref = true;
                                            for attr in se.attributes().flatten() {
                                                if local_name(attr.key.as_ref()) == "idx" {
                                                    fill_ref_idx = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                                                    if let Some(wsp) = s839_wsp.as_mut() { wsp.fillref.0 = fill_ref_idx; }
                                                }
                                            }
                                        }
                                        "lnRef" => {
                                            in_ln_ref = true;
                                            for attr in se.attributes().flatten() {
                                                if local_name(attr.key.as_ref()) == "idx" {
                                                    ln_ref_idx = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                                                    if let Some(wsp) = s839_wsp.as_mut() { wsp.lnref.0 = ln_ref_idx; }
                                                }
                                            }
                                        }
                                        // S839: srgbClr style refs feed ONLY the per-shape
                                        // record (the legacy globals never read srgbClr here).
                                        "srgbClr" => {
                                            for attr in se.attributes().flatten() {
                                                if local_name(attr.key.as_ref()) == "val" {
                                                    let c = String::from_utf8_lossy(&attr.value).to_string();
                                                    if let Some(wsp) = s839_wsp.as_mut() {
                                                        if in_fill_ref { wsp.fillref.1 = Some(c.clone()); }
                                                        if in_ln_ref { wsp.lnref.1 = Some(c.clone()); }
                                                    }
                                                }
                                            }
                                        }
                                        "schemeClr" => {
                                            for attr in se.attributes().flatten() {
                                                if local_name(attr.key.as_ref()) == "val" {
                                                    let val = String::from_utf8_lossy(&attr.value).to_string();
                                                    if let Some(resolved) = ctx.theme.resolve(&val) {
                                                        let color = parse_color_modifiers(reader, resolved, "schemeClr");
                                                        if let Some(wsp) = s839_wsp.as_mut() {
                                                            if in_fill_ref { wsp.fillref.1 = Some(color.clone()); }
                                                            if in_ln_ref { wsp.lnref.1 = Some(color.clone()); }
                                                        }
                                                        if in_fill_ref && fill_ref_idx > 0 && shape_fill.is_none() {
                                                            shape_fill = Some(color.clone());
                                                        }
                                                        if in_ln_ref && ln_ref_idx > 0 && stroke_color.is_none() {
                                                            stroke_color = Some(color);
                                                        }
                                                    }
                                                    style_depth -= 1;
                                                }
                                            }
                                        }
                                        _ => {}
                                    }
                                }
                                Ok(Event::Empty(se)) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "schemeClr" {
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "val" {
                                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                                if let Some(resolved) = ctx.theme.resolve(&val) {
                                                    if let Some(wsp) = s839_wsp.as_mut() {
                                                        if in_fill_ref { wsp.fillref.1 = Some(resolved.clone()); }
                                                        if in_ln_ref { wsp.lnref.1 = Some(resolved.clone()); }
                                                    }
                                                    if in_fill_ref && fill_ref_idx > 0 && shape_fill.is_none() {
                                                        shape_fill = Some(resolved.clone());
                                                    }
                                                    if in_ln_ref && ln_ref_idx > 0 && stroke_color.is_none() {
                                                        stroke_color = Some(resolved.clone());
                                                    }
                                                }
                                            }
                                        }
                                    } else if sl == "srgbClr" {
                                        // S839: per-shape only (legacy globals unchanged).
                                        for attr in se.attributes().flatten() {
                                            if local_name(attr.key.as_ref()) == "val" {
                                                let c = String::from_utf8_lossy(&attr.value).to_string();
                                                if let Some(wsp) = s839_wsp.as_mut() {
                                                    if in_fill_ref { wsp.fillref.1 = Some(c.clone()); }
                                                    if in_ln_ref { wsp.lnref.1 = Some(c.clone()); }
                                                }
                                            }
                                        }
                                    }
                                }
                                Ok(Event::End(_)) => {
                                    style_depth -= 1;
                                    if style_depth == 1 {
                                        in_fill_ref = false;
                                        in_ln_ref = false;
                                    }
                                    if style_depth == 0 { break; }
                                }
                                Ok(Event::Eof) => break,
                                _ => {}
                            }
                        }
                        depth -= 1; // consumed the style end tag
                    }
                    _ => {}
                }
            }
            Event::Text(e) => {
                if in_pos_offset {
                    let content = e.unescape().unwrap_or_default();
                    if let Ok(v) = content.parse::<f32>() {
                        let pt = v / 12700.0; // EMU to points
                        if in_pos_h { pos_x = pt; }
                        else if in_pos_v { pos_y = pt; }
                    }
                } else if in_align {
                    let content = e.unescape().unwrap_or_default().to_string();
                    if in_pos_h { h_align = Some(content); }
                    else if in_pos_v { v_align = Some(content); }
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    // S493i: connector arrowheads. a:headEnd (line start) / a:tailEnd (line end);
                    // type attr e.g. "triangle"/"arrow"/"stealth"/"oval" → draw arrowhead; absent
                    // or "none" → none.
                    "headEnd" | "tailEnd" => {
                        let mut has = false;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "type" {
                                let v = String::from_utf8_lossy(&attr.value);
                                has = !v.is_empty() && v != "none";
                            }
                        }
                        if local == "headEnd" { arrow_head = has; } else { arrow_tail = has; }
                    }
                    "wrapNone" => { wrap_type = Some(WrapType::None); }
                    "wrapSquare" => { wrap_type = Some(WrapType::Square); }
                    "wrapTight" => { wrap_type = Some(WrapType::Tight); }
                    "wrapTopAndBottom" => { wrap_type = Some(WrapType::TopAndBottom); }
                    "extent" => {
                        // wp:extent cx/cy are in EMUs (English Metric Units)
                        // 1 inch = 914400 EMUs, 1 point = 12700 EMUs
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val = String::from_utf8_lossy(&attr.value);
                            match key {
                                "cx" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        width = v / 12700.0; // EMU to points
                                    }
                                }
                                "cy" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        height = v / 12700.0;
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                    "blip" => {
                        // a:blip r:embed="rId1"
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:embed" || key.ends_with(":embed") || key == "embed" {
                                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "srcRect" => {
                        // a:srcRect — image crop percentages (in 1/1000th percent)
                        let mut c = ImageCrop { top: 0.0, right: 0.0, bottom: 0.0, left: 0.0 };
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val = String::from_utf8_lossy(&attr.value);
                            if let Ok(v) = val.parse::<f32>() {
                                let pct = v / 1000.0; // Convert from 1/1000th percent to percent
                                match key {
                                    "t" => c.top = pct,
                                    "r" => c.right = pct,
                                    "b" => c.bottom = pct,
                                    "l" => c.left = pct,
                                    _ => {}
                                }
                            }
                        }
                        if c.top > 0.0 || c.right > 0.0 || c.bottom > 0.0 || c.left > 0.0 {
                            crop = Some(c);
                        }
                    }
                    "ext" => {
                        // a:ext cx/cy fallback for size
                        let (mut s839_cx, mut s839_cy) = (0.0f32, 0.0f32);
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            let val = String::from_utf8_lossy(&attr.value);
                            match key {
                                "cx" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        s839_cx = v;
                                        if width == 0.0 { width = v / 12700.0; }
                                    }
                                }
                                "cy" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        s839_cy = v;
                                        if height == 0.0 { height = v / 12700.0; }
                                    }
                                }
                                _ => {}
                            }
                        }
                        // S839: per-shape / group extent capture (EMU).
                        if s839_in_wgp {
                            if s839_in_grpsppr { s839_grp.1 = (s839_cx, s839_cy); }
                            else if let Some(wsp) = s839_wsp.as_mut() { wsp.ext = (s839_cx, s839_cy); }
                        }
                    }
                    "docPr" => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "descr" {
                                alt_text = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    // Shape preset geometry (e.g. rect, ellipse, roundRect, triangle, etc.)
                    "prstGeom" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "prst" {
                                let v = String::from_utf8_lossy(&attr.value).to_string();
                                if let Some(wsp) = s839_wsp.as_mut() { wsp.prst = Some(v.clone()); }
                                shape_type = Some(v);
                            }
                        }
                    }
                    // Adjustment values inside prstGeom (e.g. roundRect corner radius)
                    "gd" => {
                        let mut is_adj = false;
                        let mut adj_val: Option<f32> = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "name" && String::from_utf8_lossy(&attr.value) == "adj" {
                                is_adj = true;
                            }
                            if key == "fmla" {
                                let fmla = String::from_utf8_lossy(&attr.value).to_string();
                                if let Some(val_str) = fmla.strip_prefix("val ") {
                                    adj_val = val_str.trim().parse::<f32>().ok();
                                }
                            }
                        }
                        if is_adj {
                            if let Some(val) = adj_val {
                                corner_radius_adj = Some(val);
                            }
                        }
                    }
                    // Shape solid fill color
                    "srgbClr" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if let Some(wsp) = s839_wsp.as_mut() {
                                    if wsp.fill.is_none() && !wsp.no_fill { wsp.fill = Some(val.clone()); }
                                }
                                if shape_fill.is_none() && !has_no_fill {
                                    shape_fill = Some(val);
                                }
                            }
                        }
                    }
                    // System color — use lastClr attribute as fallback
                    "sysClr" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "lastClr" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if shape_fill.is_none() && !has_no_fill {
                                    shape_fill = Some(val);
                                }
                            }
                        }
                    }
                    // Theme color reference — resolve via theme color map
                    "schemeClr" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if shape_fill.is_none() && !has_no_fill {
                                    if let Some(resolved) = ctx.theme.resolve(&val) {
                                        shape_fill = Some(resolved.clone());
                                    }
                                }
                            }
                        }
                    }
                    // Text body properties — text insets (bodyPr as empty element)
                    "bodyPr" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "lIns" => { text_inset_left = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "rIns" => { text_inset_right = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "tIns" => { text_inset_top = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "bIns" => { text_inset_bottom = val.parse::<f32>().ok().map(|v| v / 12700.0); }
                                "anchor" => { text_body_anchor = Some(val.to_string()); }
                                "vertOverflow" => { text_vert_overflow = Some(val.to_string()); }
                                "compatLnSpc" => { text_compat_ln_spc = val == "1" || val == "true"; }
                                _ => {}
                            }
                        }
                    }
                    "noFill" => {
                        if let Some(wsp) = s839_wsp.as_mut() { wsp.no_fill = true; }
                        has_no_fill = true;
                    }
                    "noLn" => { has_no_stroke = true; }
                    // S839: wpg group / member-shape geometry (Empty events;
                    // gated on the group state so non-group drawings are
                    // untouched — a:off/a:ext were unhandled here before).
                    "off" | "chOff" if s839_in_wgp => {
                        let (mut x, mut y) = (0.0f32, 0.0f32);
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            if key == "x" { x = val.parse().unwrap_or(0.0); }
                            else if key == "y" { y = val.parse().unwrap_or(0.0); }
                        }
                        if s839_in_grpsppr {
                            if local == "off" { s839_grp.0 = (x, y); } else { s839_grp.2 = (x, y); }
                        } else if local == "off" {
                            if let Some(wsp) = s839_wsp.as_mut() { wsp.off = (x, y); }
                        }
                    }
                    "chExt" if s839_in_wgp => {
                        let (mut cx, mut cy) = (0.0f32, 0.0f32);
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            if key == "cx" { cx = val.parse().unwrap_or(0.0); }
                            else if key == "cy" { cy = val.parse().unwrap_or(0.0); }
                        }
                        if s839_in_grpsppr { s839_grp.3 = (cx, cy); s839_grp_seen = true; }
                    }
                    "pt" if s839_wsp.is_some() => {
                        let (mut x, mut y) = (0.0f32, 0.0f32);
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            if key == "x" { x = val.parse().unwrap_or(0.0); }
                            else if key == "y" { y = val.parse().unwrap_or(0.0); }
                        }
                        s839_wsp.as_mut().unwrap().pts.push((x, y));
                    }
                    // Shape line/outline properties (as empty element)
                    "ln" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "w" {
                                let val = String::from_utf8_lossy(&attr.value);
                                stroke_width = val.parse::<f32>().ok().map(|v| v / 12700.0);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "positionH" => { in_pos_h = false; }
                    "positionV" => { in_pos_v = false; }
                    "posOffset" => { in_pos_offset = false; }
                    "align" => { in_align = false; }
                    // S839: classify + record the finished group member shape.
                    "wsp" if s839_in_wgp => {
                        if let Some(w) = s839_wsp.take() {
                            let stroke = if w.ln_nofill { None } else {
                                w.ln_color.clone().or_else(|| if w.lnref.0 > 0 { w.lnref.1.clone() } else { None })
                            };
                            let fill = if w.no_fill { None } else {
                                w.fill.clone().or_else(|| if w.fillref.0 > 0 { w.fillref.1.clone() } else { None })
                            };
                            // line: degenerate extent or prst="line".
                            let is_line = w.prst.as_deref() == Some("line")
                                || w.ext.0 == 0.0 || w.ext.1 == 0.0;
                            // rect: preset rect, or a curve-free single closed
                            // path whose points all sit on the extent corners.
                            let is_rect = !is_line && (w.prst.as_deref() == Some("rect")
                                || (w.prst.is_none() && !w.has_curve && w.n_paths <= 1 && {
                                    let (pw, ph) = w.path_wh.unwrap_or(w.ext);
                                    let tx = (pw * 0.02).max(1.0);
                                    let ty = (ph * 0.02).max(1.0);
                                    !w.pts.is_empty() && w.pts.len() <= 6
                                        && w.pts.iter().all(|&(px, py)|
                                            (px.abs() < tx || (px - pw).abs() < tx)
                                            && (py.abs() < ty || (py - ph).abs() < ty))
                                }));
                            if std::env::var("OXI_DBG839").is_ok() {
                                eprintln!("[S839wsp] off={:?} ext={:?} prst={:?} pts={} curve={} paths={} ln_w={:?} ln_c={:?} lnref={:?} fillref={:?} nofill={} -> stroke={:?} fill={:?} line={} rect={}",
                                    w.off, w.ext, w.prst, w.pts.len(), w.has_curve, w.n_paths,
                                    w.ln_w, w.ln_color, w.lnref, w.fillref, w.no_fill,
                                    stroke, fill, is_line, is_rect);
                            }
                            if (stroke.is_some() || fill.is_some()) && (is_line || is_rect) {
                                let ((gx, gy), (gex, gey), (cx0, cy0), (cex, cey)) = s839_grp;
                                let (sx, sy) = if s839_grp_seen && cex > 0.0 && cey > 0.0 {
                                    (gex / cex, gey / cey)
                                } else { (1.0, 1.0) };
                                let e2p = 1.0 / 12700.0;
                                s839_vector_shapes.push(crate::ir::VectorShape {
                                    x: (gx + (w.off.0 - cx0) * sx) * e2p,
                                    y: (gy + (w.off.1 - cy0) * sy) * e2p,
                                    w: w.ext.0 * sx * e2p,
                                    h: w.ext.1 * sy * e2p,
                                    fill,
                                    stroke,
                                    stroke_width: w.ln_w.unwrap_or(0.75).max(0.25),
                                    is_line,
                                });
                            }
                        }
                    }
                    "wgp" => { s839_in_wgp = false; s839_wsp = None; }
                    "grpSpPr" => { s839_in_grpsppr = false; }
                    _ => {}
                }
                if local == "drawing" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    let position = if is_anchor {
        Some(FloatingPosition { x: pos_x, y: pos_y, h_relative, v_relative, h_align, v_align, dist_l, dist_r })
    } else {
        None
    };

    // Build image if we have a blip reference
    let image = if let Some(rid) = rel_id {
        let data = ctx.media.get(&rid).cloned().unwrap_or_default();
        let content_type = ctx.media_types.get(&rid).cloned();
        Some(Image {
            data,
            width,
            height,
            alt_text,
            content_type,
            position: position.clone(),
            wrap_type,
            crop,
            anchor_block_index: 0,
            relative_height,
            behind_doc,
        })
    } else {
        None
    };

    // Save stroke info for TextBox before Shape takes ownership
    let stroke_color_saved = stroke_color.clone();
    let stroke_width_saved = stroke_width;

    // Build shape if we detected a preset geometry
    let shape = if let Some(ref st) = shape_type {
        Some(Shape {
            shape_type: st.clone(),
            width,
            height,
            position: position.clone(),
            fill: if has_no_fill { None } else { shape_fill.clone() },
            stroke_color: if has_no_stroke { None } else { stroke_color },
            stroke_width: if has_no_stroke { None } else { stroke_width },
            text_blocks: Vec::new(), // text goes to text_box
            rotation,
            gradient_stops: gradient_stops.clone(),
            gradient_angle,
            anchor_block_index: 0,
            v_text_anchor: None,
            flip_h,
            flip_v,
            arrow_head,
            arrow_tail,
            is_vml: false, // DrawingML
            escapes_cell: false,
        })
    } else {
        None
    };

    // Compute corner radius for roundRect shapes
    let corner_radius = if shape_type.as_deref() == Some("roundRect") {
        // adj value in 0-100000 scale; default roundRect adj = 16667 (1/6)
        let adj = corner_radius_adj.unwrap_or(16667.0);
        let min_dim = width.min(height);
        Some(min_dim * adj / 100000.0)
    } else {
        None
    };

    // Build text box if we have text content OR visual appearance (fill/border) in a shape.
    // Preset shapes without text content (e.g. bracketPair outlines) should NOT generate
    // a TextBox — they are rendered via PresetShape only. A TextBox with white fill would
    // cover the underlying body text that Word renders behind/through the shape.
    let is_outline_shape = shape.is_some() && shape_text_blocks.is_empty();
    let has_visual = !has_no_fill || !has_no_stroke;
    // S535 (2026-06-10): an INLINE drawing (wp:inline, no wp:anchor position)
    // that carries a text box — e.g. a wordprocessingCanvas (wpc) figure with
    // label textboxes (3a4f's 代替休暇 figures, "キャンバス 3", extent
    // 389×137pt) — occupies flow space like an inline image (Word reserves
    // wp:extent in the host line) and renders its content AT that reserved
    // area. Previously such a text_box had position=None, so the layout's
    // resolve fell back to (margin.left, margin.top): every inline canvas
    // rendered at the page TOP-LEFT (3a4f's two figures overlapped there,
    // duplicating 20 label elements exactly) and NO flow height was reserved
    // (the word-p63/64 7-paragraph delta=-1 cluster).
    // Fix: give the text_box a synthetic paragraph-relative position (y=0 at
    // its host block) and emit a data-less placeholder Image of the drawing
    // extent so the flow reserves the canvas height (renderers skip empty
    // data; Block::Image height handling exists for body and cells per S533).
    let inline_tb = !is_anchor;
    // S741: does this drawing carry REAL textbox text (wps:txbx content)?
    // Distinguishes an inline TEXTBOX (Word grows the host line to the
    // wp:extent — flow reservation needed) from a merely-visual inline shape
    // (line/bracket decorations — S535b's revert case: reserving those
    // regressed 3a4f by 282 paras).
    let s741_has_txbx_text = !shape_text_blocks.is_empty();
    let tb_position = position.clone().or_else(|| if inline_tb {
        Some(FloatingPosition {
            x: 0.0,
            y: 0.0,
            h_relative: Some("column".to_string()),
            v_relative: Some("paragraph".to_string()),
            h_align: None,
            v_align: None,
            dist_l: None, dist_r: None })
    } else { None });
    // S839: a VISUAL-ONLY wpg group (no txbxContent — hmrc's checkbox strips /
    // writing boxes / heavy rules) carries its drawable primitives on the
    // textbox and suppresses the legacy whole-extent outline (border came from
    // the group members' a:ln via the generic last-win extraction). Text-
    // bearing groups (framework cover pages) attach nothing → byte-identical.
    let s839_attach = !s839_vector_shapes.is_empty() && shape_text_blocks.is_empty()
        && std::env::var("OXI_S839_DISABLE").is_err();
    let text_box = if !is_outline_shape && (!shape_text_blocks.is_empty() || has_visual) {
        Some(TextBox {
            blocks: shape_text_blocks,
            width,
            height,
            position: tb_position,
            border: !has_no_stroke && !s839_attach,
            stroke_color: if has_no_stroke || s839_attach { None } else { stroke_color_saved.clone() },
            stroke_width: if has_no_stroke || s839_attach { None } else { stroke_width_saved },
            fill: if has_no_fill || s839_attach { None } else { shape_fill.clone().or_else(|| shape_type.as_ref().map(|_| "FFFFFF".to_string())) },
            anchor_block_index: 0, // set by caller in parse_body
            corner_radius,
            inset_left: text_inset_left,
            inset_right: text_inset_right,
            inset_top: text_inset_top,
            inset_bottom: text_inset_bottom,
            // S535: an inline canvas text_box must stay RENDER-ONLY — without
            // an explicit WrapType::None the table wrap-below rule (layout
            // ~2415) fires on its new synthetic position and pushes following
            // tables below the canvas band (3a4f pagination 0.9931 -> 0.7434,
            // 395 paras +1). Real anchored textboxes keep their parsed wrap.
            wrap_type: if inline_tb && wrap_type.is_none() {
                Some(crate::ir::WrapType::None)
            } else {
                wrap_type
            },
            v_text_anchor: text_body_anchor,
            relative_height,
            behind_doc,
            vert_overflow: text_vert_overflow,
            compat_line_spacing: text_compat_ln_spc,
            vector_shapes: if s839_attach { std::mem::take(&mut s839_vector_shapes) } else { Vec::new() },
        })
    } else {
        None
    };

    // S535b: flow reservation for ALL inline text-bearing drawings TRIED +
    // REVERTED (3a4f pagination 0.9931 -> 0.8076, 282 paras +1..+3): the
    // placeholder fired on every inline has_visual drawing AND the empty host
    // paragraph line was still emitted (double count).
    // S537b (2026-06-10): retried SCOPED to wordprocessingCanvas (wpc) only,
    // now that S537/S536 suppress the empty host paragraph: an image-only
    // paragraph's line = extent EXACTLY per the pinned spec
    // (_s537_inline_line.py), so a canvas-only paragraph reserves exactly
    // wp:extent — Word's model. The canvas content textbox (synthetic
    // paragraph-relative position above) anchors at this placeholder block
    // and renders ON the reserved area.
    // S741 (2026-07-04, default ON, opt-out OXI_S741_DISABLE): extend the
    // S537b flow-reservation placeholder from wordprocessingCanvas to any
    // INLINE drawing with REAL textbox text (modern wps inline textboxes).
    // Word reserves the wp:extent in the host flow (probeqwps: object para
    // consumes 108.3 = text 54 + object 54); Oxi rendered the box overlay-
    // only -> packed ~1 para more per page (-1x6). The S535b danger (visual
    // shapes) is excluded by requiring shape_text_blocks non-empty.
    let s741_reserve = s741_has_txbx_text && !is_canvas
        && std::env::var("OXI_S741_DISABLE").is_err();
    let image = if image.is_none() && inline_tb && (is_canvas || s741_reserve) && text_box.is_some()
        && width > 0.0 && height > 0.0
    {
        Some(Image {
            data: Vec::new(),
            width,
            height,
            alt_text: None,
            content_type: None,
            position: None,
            wrap_type: None,
            crop: None,
            anchor_block_index: 0,
            relative_height,
            behind_doc,
        })
    } else {
        image
    };

    Ok(DrawingResult { image, shape, text_box })
}

/// Parse VML w:pict element (legacy shapes/images)
fn parse_vml_pict(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<DrawingResult, ParseError> {
    let mut shape_type = None;
    let mut width: f32 = 0.0;
    let mut height: f32 = 0.0;
    let mut fill_color: Option<String> = None;
    let mut stroke_color_val: Option<String> = None;
    let mut stroke_width_val: Option<f32> = None;
    let mut no_stroke = false;
    let mut no_fill = false;
    let mut rel_id: Option<String> = None;
    let mut text_blocks: Vec<Block> = Vec::new();
    let mut v_text_anchor: Option<String> = None;
    let mut margin_left: f32 = 0.0;
    let mut margin_top: f32 = 0.0;
    let mut is_absolute = false;
    let mut escapes_cell = false; // o:allowincell="f" (S711b)
    // VML drawing canvas (<v:group editas="canvas">). Word reserves the
    // group's DECLARED height (style height:Npt) in the inline text flow; the
    // inner position:absolute shapes are canvas-internal (relative to the
    // canvas coordinate system), NOT page-absolute. parse_vml_pict previously
    // had no "group" arm, so the canvas dims were lost and the inner shapes
    // overwrote width/height (with canvas-unit coords) -> the figure reserved
    // ZERO flow height (tokyoshugyo 図２ 代替休暇 flowchart, 136.7pt, dropped).
    let mut group_width: f32 = 0.0;
    let mut group_height: f32 = 0.0;
    let mut is_canvas_group = false;
    let mut depth = 0;
    // S850 (2026-07-14, default ON, opt-out OXI_S850_DISABLE): inside a
    // <v:group> the inner shapes' style dims/position are GROUP-INTERNAL
    // (canvas coordinate space, e.g. width:4940;height:4853 within
    // coordsize=4960,5123) and must NOT overwrite the outer function's
    // width/height/is_absolute/margins. An absolutely-positioned image-bearing
    // v:group (reference poster icons: <v:group style="width:247pt;
    // height:255.65pt;mso-position-*-relative:page" ...><v:shape
    // style="position:absolute;width:4252;height:1602"><v:imagedata/>) was
    // being emitted as an INLINE image with the inner shape's VML-unit dims
    // (4252x1602pt) reserving ~1602pt of flow -> each icon forced its own page
    // (Word floats them -> 2 pages, Oxi -> 10). Track group nesting; when
    // inside a group, skip the inner-shape dim/position writes and take the
    // outermost group's declared pt dims + margin position instead.
    let s850 = std::env::var("OXI_S850_DISABLE").is_err();
    let mut group_depth: i32 = 0;
    // S852: an inline VML horizontal rule (<v:rect o:hr="t" .../>). Word draws
    // it on its OWN line (a full-width gray divider). Marked here so parse_run
    // routes it to the run style (reserve a line + draw the rule) instead of a
    // generic inline Shape (which the layout skips because position is None).
    let mut is_hr = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                depth += 1;
                match local.as_str() {
                    // VML text box content — consume all paragraphs inside
                    "txbxContent" => {
                        loop {
                            match reader.read_event() {
                                Ok(Event::Start(se)) => {
                                    if local_name(se.name().as_ref()) == "p" {
                                        if let Ok(pr) = parse_paragraph(reader, ctx, styles, false) {
                                            text_blocks.push(Block::Paragraph(pr.paragraph));
                                        }
                                    }
                                }
                                Ok(Event::End(se)) => {
                                    if local_name(se.name().as_ref()) == "txbxContent" {
                                        break;
                                    }
                                }
                                Ok(Event::Eof) => break,
                                _ => {}
                            }
                        }
                        depth -= 1; // consumed the txbxContent end tag
                    }
                    // VML drawing canvas group — capture its declared pt dims so
                    // the figure reserves height in the inline flow (see field decls).
                    // Inline canvas only (mso-position-*-relative:line/char); a
                    // position:absolute group floats (handled by is_absolute, S566).
                    "group" => {
                        let outer = group_depth == 0;
                        group_depth += 1;
                        // Only the OUTERMOST group's style carries the real pt
                        // display dims + page position; nested groups are
                        // canvas-internal (S850).
                        if outer || !s850 {
                            let mut g_abs = false;
                            let mut g_ml = 0.0;
                            let mut g_mt = 0.0;
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "style" {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    for part in val.split(';') {
                                        let part = part.trim();
                                        if let Some(w) = part.strip_prefix("width:") {
                                            group_width = parse_css_length(w.trim());
                                        } else if let Some(h) = part.strip_prefix("height:") {
                                            group_height = parse_css_length(h.trim());
                                        } else if let Some(ml) = part.strip_prefix("margin-left:") {
                                            g_ml = parse_css_length(ml.trim());
                                        } else if let Some(mt) = part.strip_prefix("margin-top:") {
                                            g_mt = parse_css_length(mt.trim());
                                        } else if part.starts_with("position:absolute") {
                                            g_abs = true;
                                        }
                                    }
                                }
                            }
                            if g_abs {
                                is_absolute = true;
                            } else if group_height > 0.0 {
                                is_canvas_group = true;
                            }
                            if s850 {
                                // Take the group's declared pt dims + margin
                                // position so the inner canvas-unit shapes can be
                                // ignored below.
                                if group_width > 0.0 { width = group_width; }
                                if group_height > 0.0 { height = group_height; }
                                if g_ml != 0.0 { margin_left = g_ml; }
                                if g_mt != 0.0 { margin_top = g_mt; }
                            }
                        }
                    }
                    // VML shape types
                    "shape" | "rect" | "oval" | "roundrect" | "line" => {
                        // Check VML type attribute for preset shape identification
                        let vml_type_attr = e.attributes().flatten()
                            .find(|a| local_name(a.key.as_ref()) == "type")
                            .map(|a| String::from_utf8_lossy(&a.value).to_string());
                        shape_type = Some(match local.as_str() {
                            "shape" => {
                                // Map VML shapetype IDs to OOXML preset names
                                match vml_type_attr.as_deref() {
                                    Some(t) if t.contains("t185") => "bracketPair", // double bracket 〔〕
                                    _ => "rect",
                                }
                            }
                            "roundrect" => "roundRect",
                            other => other,
                        }.to_string());
                        // Parse style attribute for width/height
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "style" => {
                                    let in_grp = s850 && group_depth > 0;
                                    for part in val.split(';') {
                                        let part = part.trim();
                                        if let Some(w) = part.strip_prefix("width:") {
                                            if !in_grp { width = parse_css_length(w.trim()); }
                                        } else if let Some(h) = part.strip_prefix("height:") {
                                            if !in_grp { height = parse_css_length(h.trim()); }
                                        } else if let Some(anchor) = part.strip_prefix("v-text-anchor:") {
                                            v_text_anchor = Some(anchor.trim().to_string());
                                        } else if let Some(ml) = part.strip_prefix("margin-left:") {
                                            if !in_grp { margin_left = parse_css_length(ml.trim()); }
                                        } else if let Some(mt) = part.strip_prefix("margin-top:") {
                                            if !in_grp { margin_top = parse_css_length(mt.trim()); }
                                        } else if part.starts_with("position:absolute") {
                                            if !in_grp { is_absolute = true; }
                                        }
                                    }
                                }
                                "filled" => { if val == "f" || val == "false" { no_fill = true; } }
                                "fillcolor" => fill_color = Some(val.trim_start_matches('#').to_string()),
                                "strokecolor" => stroke_color_val = Some(val.trim_start_matches('#').to_string()),
                                "strokeweight" => stroke_width_val = parse_css_length_opt(&val),
                                "stroked" => { if val == "f" || val == "false" { no_stroke = true; } }
                                "allowincell" => { if val == "f" || val == "false" { escapes_cell = true; } }
                                // S852: o:hr="t" marks this rect as a horizontal rule.
                                "hr" => { if val == "t" || val == "true" { is_hr = true; } }
                                _ => {}
                            }
                        }
                        if is_hr && std::env::var("OXI_S852_DISABLE").is_err() {
                            shape_type = Some("hr".to_string());
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "imagedata" => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:id" || key.ends_with(":id") {
                                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "fill" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "color" {
                                fill_color = Some(String::from_utf8_lossy(&attr.value).trim_start_matches('#').to_string());
                            }
                        }
                    }
                    "stroke" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "color" => stroke_color_val = Some(val.trim_start_matches('#').to_string()),
                                "weight" => stroke_width_val = parse_css_length_opt(&val),
                                "on" => { if val == "f" || val == "false" { no_stroke = true; } }
                                _ => {}
                            }
                        }
                    }
                    // S711 (2026-07-01): a SELF-CLOSING VML shape (<v:rect .../>,
                    // <v:oval/>, <v:roundrect/>, <v:line/>) arrives as Event::Empty,
                    // NOT Event::Start — the Start arm above (4261) only fires for
                    // shapes WITH children. Without this arm a childless filled rect
                    // (tokyoshugyo/3a4f/model (注) note: <v:rect fillcolor="silver"
                    // .../>, the gray reference box) set no shape_type -> the Shape
                    // was dropped entirely. Mirror the Start-arm attribute parse.
                    // Opt-out OXI_VMLRECT_DISABLE.
                    "shape" | "rect" | "oval" | "roundrect" | "line"
                        if std::env::var("OXI_VMLRECT_DISABLE").is_err() =>
                    {
                        let vml_type_attr = e.attributes().flatten()
                            .find(|a| local_name(a.key.as_ref()) == "type")
                            .map(|a| String::from_utf8_lossy(&a.value).to_string());
                        shape_type = Some(match local.as_str() {
                            "shape" => match vml_type_attr.as_deref() {
                                Some(t) if t.contains("t185") => "bracketPair",
                                _ => "rect",
                            },
                            "roundrect" => "roundRect",
                            other => other,
                        }.to_string());
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "style" => {
                                    // S850: inside a v:group the inner shape style
                                    // dims/margins/position are canvas-internal.
                                    let in_grp = s850 && group_depth > 0;
                                    for part in val.split(';') {
                                        let part = part.trim();
                                        if let Some(w) = part.strip_prefix("width:") {
                                            if !in_grp { width = parse_css_length(w.trim()); }
                                        } else if let Some(h) = part.strip_prefix("height:") {
                                            if !in_grp { height = parse_css_length(h.trim()); }
                                        } else if let Some(anchor) = part.strip_prefix("v-text-anchor:") {
                                            v_text_anchor = Some(anchor.trim().to_string());
                                        } else if let Some(ml) = part.strip_prefix("margin-left:") {
                                            if !in_grp { margin_left = parse_css_length(ml.trim()); }
                                        } else if let Some(mt) = part.strip_prefix("margin-top:") {
                                            if !in_grp { margin_top = parse_css_length(mt.trim()); }
                                        } else if part.starts_with("position:absolute") {
                                            if !in_grp { is_absolute = true; }
                                        }
                                    }
                                }
                                "filled" => { if val == "f" || val == "false" { no_fill = true; } }
                                "fillcolor" => fill_color = Some(val.trim_start_matches('#').to_string()),
                                "strokecolor" => stroke_color_val = Some(val.trim_start_matches('#').to_string()),
                                "strokeweight" => stroke_width_val = parse_css_length_opt(&val),
                                "stroked" => { if val == "f" || val == "false" { no_stroke = true; } }
                                "allowincell" => { if val == "f" || val == "false" { escapes_cell = true; } }
                                // S852: o:hr="t" marks this rect as a horizontal rule.
                                "hr" => { if val == "t" || val == "true" { is_hr = true; } }
                                _ => {}
                            }
                        }
                        if is_hr && std::env::var("OXI_S852_DISABLE").is_err() {
                            shape_type = Some("hr".to_string());
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "pict" && depth == 0 {
                    break;
                }
                if local == "group" && group_depth > 0 {
                    group_depth -= 1;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // Inline VML drawing canvas (<v:group editas="canvas">): reserve the
    // group's declared height in the body flow as an inline (position:None)
    // height-only Image. The inner vector shapes/textboxes are not yet
    // rendered (Phase-3) but Word reserves the full canvas height (tokyoshugyo
    // 図２ 136.7pt was being dropped, cascading the 賃金 chapter -1 page).
    // S640 (2026-06-22, default ON, opt-out OXI_VMLCANVAS_DISABLE): legacy-VML
    // canvas figures reserve their declared height. 3a4f/model are
    // template-twins but render their canvases via the DrawingML (w:drawing)
    // path, so this VML-only fix leaves them byte-identical (canary-verified).
    if is_canvas_group
        && group_height > 0.0
        && std::env::var("OXI_VMLCANVAS_DISABLE").is_err()
    {
        return Ok(DrawingResult {
            image: Some(Image {
                data: Vec::new(),
                width: group_width,
                height: group_height,
                alt_text: None,
                content_type: None,
                position: None,
                wrap_type: None,
                crop: None,
                anchor_block_index: 0,
                relative_height: 0,
                behind_doc: false,
            }),
            shape: None,
            text_box: None,
        });
    }

    // VML absolute-positioned shapes get a FloatingPosition. Computed BEFORE
    // the image (S566) so an absolutely-positioned VML picture carries it and
    // is classified as FLOATING by the IR builder, not inline.
    let vml_position = if is_absolute && (margin_left != 0.0 || margin_top != 0.0) {
        Some(FloatingPosition {
            x: margin_left,
            y: margin_top,
            h_relative: Some("text".to_string()),
            v_relative: Some("text".to_string()),
            h_align: None,
            v_align: None,
            dist_l: None, dist_r: None })
    } else {
        None
    };

    // Build image if we have a blip reference
    let image = if let Some(rid) = rel_id {
        let data = ctx.media.get(&rid).cloned().unwrap_or_default();
        let content_type = ctx.media_types.get(&rid).cloned();
        // S566 (2026-06-14): an absolutely-positioned VML picture (harassmanual's
        // 509.75x726.75pt flowchart, z-index<0 behind text) must carry its
        // FloatingPosition so the IR builder (`Separate inline images ... from
        // floating images`) classifies it as FLOATING and does NOT reserve its
        // height in the body flow. Previously hardcoded position:None -> treated
        // as INLINE -> ~2 pages of flow reserved -> spurious empty pages p5/p6.
        // wrap_type None = behind-text overlay, no text reservation.
        // Opt-out OXI_S566_DISABLE restores the old inline (position:None) path.
        let s566 = std::env::var("OXI_S566_DISABLE").is_err();
        let img_position = if s566 { vml_position.clone() } else { None };
        let img_wrap = if s566 && vml_position.is_some() { Some(WrapType::None) } else { None };
        Some(Image {
            data,
            width,
            height,
            alt_text: None,
            content_type,
            position: img_position,
            wrap_type: img_wrap,
            crop: None,
            anchor_block_index: 0,
            relative_height: 0,
            behind_doc: false,
        })
    } else {
        None
    };
    // S746 (2026-07-04, default ON, opt-out OXI_S746_DISABLE): an INLINE
    // (non-absolute) VML shape with v:textbox CONTENT — Word grows the host
    // line to the shape extent and renders the box + inner text; Oxi captured
    // text_blocks into the Shape (no layout consumer -> invisible) and
    // reserved nothing (probezinlinetb: 120x44 boxes -> {-1:7}). Return a
    // TextBox (renders box + text, paragraph-relative, wrap None) + the S741
    // placeholder Image (flow reservation) instead of the dead Shape. ALL 5
    // corpus docs with standalone VML textboxes (tokyoshugyo x35 / roudoujoken
    // / parttime / a1d6 / de6e32) are position:absolute -> is_absolute -> this
    // branch never fires on the corpus (byte-identical by construction);
    // absolute VML textboxes keep the existing (dropped) behavior = the
    // separate S746b render task.
    let s746_inline_txbx = !is_absolute && !text_blocks.is_empty()
        && width > 0.0 && height > 0.0 && image.is_none()
        && std::env::var("OXI_S746_DISABLE").is_err();
    if s746_inline_txbx {
        let text_box = Some(TextBox {
            blocks: text_blocks,
            width,
            height,
            position: Some(FloatingPosition {
                x: 0.0,
                y: 0.0,
                h_relative: Some("column".to_string()),
                v_relative: Some("paragraph".to_string()),
                h_align: None,
                v_align: None,
            dist_l: None, dist_r: None }),
            border: !no_stroke,
            stroke_color: if no_stroke { None } else { stroke_color_val },
            stroke_width: if no_stroke { None } else { stroke_width_val.or(Some(0.75)) },
            fill: if no_fill { None } else { fill_color.or(Some("FFFFFF".to_string())) },
            anchor_block_index: 0,
            corner_radius: if shape_type.as_deref() == Some("roundRect") { Some(3.0) } else { None },
            inset_left: None,
            inset_right: None,
            inset_top: None,
            inset_bottom: None,
            wrap_type: Some(WrapType::None),
            v_text_anchor: v_text_anchor.clone(),
            relative_height: 0,
            behind_doc: false,
            vert_overflow: None,
            compat_line_spacing: false,
            vector_shapes: Vec::new(),
        });
        let placeholder = Some(Image {
            data: Vec::new(),
            width,
            height,
            alt_text: None,
            content_type: None,
            position: None,
            wrap_type: None,
            crop: None,
            anchor_block_index: 0,
            relative_height: 0,
            behind_doc: false,
        });
        return Ok(DrawingResult { image: placeholder, shape: None, text_box });
    }
    let shape = shape_type.as_ref().map(|st| Shape {
        shape_type: st.clone(),
        width,
        height,
        position: vml_position,
        fill: if no_fill { None } else { fill_color.clone() },
        stroke_color: if no_stroke { None } else { stroke_color_val },
        stroke_width: if no_stroke { None } else { stroke_width_val.or(Some(0.75)) },
        text_blocks,
        rotation: None,
        gradient_stops: Vec::new(),
        gradient_angle: None,
        anchor_block_index: 0,
        v_text_anchor,
        flip_h: false,
        flip_v: false,
        arrow_head: false,
        arrow_tail: false,
        is_vml: true, // legacy VML <w:pict> shape
        escapes_cell,
    });

    Ok(DrawingResult { image, shape, text_box: None })
}

/// Parse CSS-like length value (e.g. "200pt", "2in", "100.5px")
fn parse_css_length(s: &str) -> f32 {
    let s = s.trim();
    if let Some(v) = s.strip_suffix("pt") {
        v.trim().parse().unwrap_or(0.0)
    } else if let Some(v) = s.strip_suffix("in") {
        v.trim().parse::<f32>().unwrap_or(0.0) * 72.0
    } else if let Some(v) = s.strip_suffix("cm") {
        v.trim().parse::<f32>().unwrap_or(0.0) * 28.3465
    } else if let Some(v) = s.strip_suffix("mm") {
        v.trim().parse::<f32>().unwrap_or(0.0) * 2.83465
    } else if let Some(v) = s.strip_suffix("px") {
        v.trim().parse::<f32>().unwrap_or(0.0) * 0.75 // 96dpi → 72pt
    } else {
        s.parse().unwrap_or(0.0)
    }
}

fn parse_css_length_opt(s: &str) -> Option<f32> {
    let v = parse_css_length(s);
    if v > 0.0 { Some(v) } else { None }
}

/// Parse w:object (OLE embedded object) — extract preview image from VML shape inside
fn parse_ole_object(reader: &mut Reader<&[u8]>, ctx: &ParseContext) -> Result<(DrawingResult, bool), ParseError> {
    let mut rel_id: Option<String> = None;
    let mut width: f32 = 0.0;
    let mut height: f32 = 0.0;
    let mut depth = 0;
    // S851: whether an <o:OLEObject> child was seen. A real OLE embed
    // (Equation.3, Visio, …) has one; a bare form-field picture (the
    // MassHealth PA-form field underlines) does NOT — the discriminator for
    // routing the inline picture as a run-level inline object vs a block.
    let mut saw_ole_object = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                depth += 1;
                if local == "OLEObject" { saw_ole_object = true; }
                match local.as_str() {
                    // VML shape inside OLE object — parse style for dimensions
                    "shape" | "rect" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "style" {
                                let val = String::from_utf8_lossy(&attr.value);
                                for part in val.split(';') {
                                    let part = part.trim();
                                    if let Some(w) = part.strip_prefix("width:") {
                                        width = parse_css_length(w.trim());
                                    } else if let Some(h) = part.strip_prefix("height:") {
                                        height = parse_css_length(h.trim());
                                    }
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    // v:imagedata — the preview image of the OLE object
                    "imagedata" => {
                        for attr in e.attributes().flatten() {
                            let key = std::str::from_utf8(attr.key.as_ref()).unwrap_or("");
                            if key == "r:id" || key.ends_with(":id") {
                                rel_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    // OLEObject element — skip gracefully (S851: note presence)
                    "OLEObject" => { saw_ole_object = true; }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "object" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // Build image from the preview if available
    let image = if let Some(rid) = rel_id {
        let data = ctx.media.get(&rid).cloned().unwrap_or_default();
        let content_type = ctx.media_types.get(&rid).cloned();
        Some(Image {
            data,
            width,
            height,
            alt_text: Some("OLE Object".to_string()),
            content_type,
            position: None,
            wrap_type: None,
            crop: None,
            anchor_block_index: 0,
            relative_height: 0,
            behind_doc: false,
        })
    } else {
        None
    };

    Ok((DrawingResult { image, shape: None, text_box: None }, saw_ole_object))
}

/// Parse mc:AlternateContent — prefer mc:Choice (DrawingML), fall back to mc:Fallback (VML)
fn parse_alternate_content(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<Option<DrawingResult>, ParseError> {
    let mut result: Option<DrawingResult> = None;
    let mut depth = 0;
    let mut in_choice = false;
    let mut in_fallback = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "Choice" if depth == 0 => {
                        in_choice = true;
                        depth += 1;
                    }
                    "Fallback" if depth == 0 => {
                        in_fallback = true;
                        depth += 1;
                    }
                    "drawing" if in_choice && depth == 1 => {
                        let dr = parse_drawing(reader, ctx, styles)?;
                        if std::env::var("OXI_DEBUG_AC").is_ok() {
                            let alltxt: String = dr.text_box.as_ref().map(|t| t.blocks.iter()
                                .filter_map(|b| if let crate::ir::Block::Paragraph(p)=b { Some(p.runs.iter().flat_map(|r| r.text.chars()).collect::<String>()) } else { None })
                                .collect::<Vec<_>>().join("|")).unwrap_or_default();
                            let pos = dr.text_box.as_ref().map(|t| t.position.is_some()).unwrap_or(false);
                            eprintln!("[AC] Choice drawing: tb={} tb_paras={} tb_pos={} alltxt={:?}",
                                dr.text_box.is_some(),
                                dr.text_box.as_ref().map(|t| t.blocks.len()).unwrap_or(0),
                                pos, alltxt.chars().take(40).collect::<String>());
                        }
                        // Only keep if it produced something useful (image, shape, or text box)
                        if result.is_none() && dr.has_content() {
                            result = Some(dr);
                        }
                    }
                    "pict" if (in_choice || in_fallback) && depth == 1 && result.is_none() => {
                        let dr = parse_vml_pict(reader, ctx, styles)?;
                        if std::env::var("OXI_DEBUG_AC").is_ok() {
                            eprintln!("[AC] pict (in_choice={} in_fallback={}): img={} shape={} tb={} shape_blocks={}",
                                in_choice, in_fallback,
                                dr.image.is_some(), dr.shape.is_some(), dr.text_box.is_some(),
                                dr.shape.as_ref().map(|s| s.text_blocks.len()).unwrap_or(0));
                        }
                        if dr.has_content() {
                            result = Some(dr);
                        }
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "AlternateContent" && depth == 0 {
                    break;
                }
                if local == "Choice" && in_choice {
                    in_choice = false;
                }
                if local == "Fallback" && in_fallback {
                    in_fallback = false;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(result)
}

/// Parse OMML math element (m:oMath or m:oMathPara) into a text representation
fn parse_omml(reader: &mut Reader<&[u8]>, end_tag: &str) -> Result<String, ParseError> {
    let mut result = String::new();
    let mut depth = 0;
    let mut in_text = false;
    // Track context for proper rendering
    let mut context_stack: Vec<String> = Vec::new();

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                depth += 1;
                match local.as_str() {
                    "t" => in_text = true,
                    "f" => context_stack.push("frac".to_string()),
                    "rad" => {
                        result.push('\u{221A}'); // √
                        context_stack.push("rad".to_string());
                    }
                    "sSup" => context_stack.push("sup".to_string()),
                    "sSub" => context_stack.push("sub".to_string()),
                    "d" => {
                        // Delimiter (parentheses)
                        let mut beg = '(';
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "begChr" {
                                let v = String::from_utf8_lossy(&attr.value);
                                beg = v.chars().next().unwrap_or('(');
                            }
                        }
                        result.push(beg);
                        context_stack.push("delim".to_string());
                    }
                    "nary" => {
                        // N-ary (summation, product, integral)
                        let mut chr = '\u{2211}'; // default: summation ∑
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "chr" {
                                let v = String::from_utf8_lossy(&attr.value);
                                chr = v.chars().next().unwrap_or('\u{2211}');
                            }
                        }
                        result.push(chr);
                        context_stack.push("nary".to_string());
                    }
                    "num" if context_stack.last().map_or(false, |c| c == "frac") => {
                        // Numerator of fraction — will add / separator after
                    }
                    "den" if context_stack.last().map_or(false, |c| c == "frac") => {
                        result.push('/');
                    }
                    "sup" if context_stack.last().map_or(false, |c| c == "sup") => {
                        result.push('^');
                    }
                    "sub" if context_stack.last().map_or(false, |c| c == "sub") => {
                        result.push('_');
                    }
                    _ => {}
                }
            }
            Event::Text(e) => {
                if in_text {
                    let content = e.unescape().unwrap_or_default();
                    result.push_str(&content);
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "t" {
                    in_text = false;
                }
                if local == end_tag && depth == 0 {
                    break;
                }
                match local.as_str() {
                    "f" | "rad" | "sSup" | "sSub" | "nary" => {
                        context_stack.pop();
                    }
                    "d" => {
                        context_stack.pop();
                        result.push(')');
                    }
                    _ => {}
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    // m:dPr begChr/endChr as empty element
                    "begChr" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                // We already pushed '(' — if different, fix it
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(result)
}

/// Parse w:rPr (run properties)
fn parse_run_properties(
    reader: &mut Reader<&[u8]>,
    ctx: &ParseContext,
    styles: &StyleSheet,
) -> Result<(RunStyle, Option<PropertyChange>), ParseError> {
    let mut style = RunStyle::default();
    let mut depth = 0;
    let mut rstyle_id: Option<String> = None;
    let mut rpr_change: Option<PropertyChange> = None;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                // rPrChange carries a full prior <w:rPr> body. If we fell through
                // to the normal handlers, every prior property (bold, italic,
                // color, font) would silently merge into the *current* style.
                // Handle it inline: capture attrs, recursively parse the nested
                // <w:rPr> to get the prior RunStyle, then consume the close tag.
                if depth == 0 && local == "rPrChange" {
                    let mut pc = PropertyChange::default();
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value).to_string();
                        match key.as_str() {
                            "id" => pc.id = Some(val),
                            "author" => pc.author = Some(val),
                            "date" => pc.date = Some(val),
                            _ => {}
                        }
                    }
                    // Drain until </w:rPrChange>, recursing into inner <w:rPr>.
                    loop {
                        match reader.read_event()? {
                            Event::Start(inner) => {
                                if local_name(inner.name().as_ref()) == "rPr" {
                                    let (prior, _nested) =
                                        parse_run_properties(reader, ctx, styles)?;
                                    pc.prior_run_style = Some(Box::new(prior));
                                }
                            }
                            Event::Empty(inner) => {
                                // <w:rPr/> with no children — prior style is default.
                                if local_name(inner.name().as_ref()) == "rPr"
                                    && pc.prior_run_style.is_none()
                                {
                                    pc.prior_run_style = Some(Box::new(RunStyle::default()));
                                }
                            }
                            Event::End(inner) => {
                                if local_name(inner.name().as_ref()) == "rPrChange" {
                                    break;
                                }
                            }
                            Event::Eof => break,
                            _ => {}
                        }
                    }
                    rpr_change = Some(pc);
                    continue;
                }
                depth += 1;
                if local == "rFonts" {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == "ascii" || key == "hAnsi" {
                            style.font_family =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                        } else if key == "eastAsia" {
                            style.font_family_east_asia =
                                Some(String::from_utf8_lossy(&attr.value).to_string());
                            style.has_explicit_east_asia = true;
                        } else if key == "asciiTheme" || key == "hAnsiTheme" {
                            if style.font_family.is_none() {
                                let val = String::from_utf8_lossy(&attr.value);
                                let font = super::styles::resolve_theme_font_pub(&val, &ctx.theme);
                                if let Some(f) = font {
                                    style.font_family = Some(f);
                                }
                            }
                        } else if key == "eastAsiaTheme" {
                            if style.font_family_east_asia.is_none() {
                                let val = String::from_utf8_lossy(&attr.value);
                                let font = super::styles::resolve_theme_font_pub(&val, &ctx.theme);
                                if let Some(f) = font {
                                    style.font_family_east_asia = Some(f);
                                }
                            }
                        }
                    }
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "b" => style.bold = true,
                    "i" => style.italic = true,
                    "u" => {
                        style.underline = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val == "none" {
                                    style.underline = false;
                                } else {
                                    style.underline_style = Some(val);
                                }
                            }
                        }
                    }
                    "strike" => style.strikethrough = true,
                    "dstrike" => {
                        style.strikethrough = true;
                        style.double_strikethrough = true;
                    }
                    "highlight" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                style.highlight =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "vertAlign" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                style.vertical_align = match val.as_ref() {
                                    "superscript" => Some(VerticalAlign::Superscript),
                                    "subscript" => Some(VerticalAlign::Subscript),
                                    _ => Some(VerticalAlign::Baseline),
                                };
                            }
                        }
                    }
                    "sz" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                // OOXML sz is in half-points
                                style.font_size = val.parse::<f32>().ok().map(|v| v / 2.0);
                            }
                        }
                    }
                    "rFonts" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "ascii" || key == "hAnsi" {
                                style.font_family =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                            } else if key == "eastAsia" {
                                style.font_family_east_asia =
                                    Some(String::from_utf8_lossy(&attr.value).to_string());
                                style.has_explicit_east_asia = true;
                            } else if key == "asciiTheme" || key == "hAnsiTheme" {
                                if style.font_family.is_none() {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    let font = super::styles::resolve_theme_font_pub(&val, &ctx.theme);
                                    if let Some(f) = font {
                                        style.font_family = Some(f);
                                    }
                                }
                            } else if key == "eastAsiaTheme" {
                                if style.font_family_east_asia.is_none() {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    let font = super::styles::resolve_theme_font_pub(&val, &ctx.theme);
                                    if let Some(f) = font {
                                        style.font_family_east_asia = Some(f);
                                    }
                                }
                            }
                        }
                    }
                    "color" => {
                        let mut color_val = None;
                        let mut theme_color = None;
                        let mut theme_tint = None;
                        let mut theme_shade = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "val" => color_val = Some(val),
                                "themeColor" => theme_color = Some(val),
                                "themeTint" => theme_tint = Some(val),
                                "themeShade" => theme_shade = Some(val),
                                _ => {}
                            }
                        }
                        // Resolve theme color if available
                        if let Some(ref tc) = theme_color {
                            if let Some(resolved) = ctx.theme.resolve(tc) {
                                let mut hex: String = resolved.clone();
                                // Apply tint/shade
                                if let Some(ref tint) = theme_tint {
                                    if let Ok(t) = u8::from_str_radix(tint, 16) {
                                        let tint_val = t as f64 / 255.0;
                                        hex = ThemeColors::apply_tint_shade(&hex, tint_val);
                                    }
                                }
                                if let Some(ref shade) = theme_shade {
                                    if let Ok(s) = u8::from_str_radix(shade, 16) {
                                        let shade_val = -(1.0 - s as f64 / 255.0);
                                        hex = ThemeColors::apply_tint_shade(&hex, shade_val);
                                    }
                                }
                                style.color = Some(hex);
                            } else if let Some(ref cv) = color_val {
                                if cv != "auto" {
                                    style.color = Some(cv.clone());
                                }
                            }
                        } else if let Some(ref cv) = color_val {
                            if cv != "auto" {
                                style.color = Some(cv.clone());
                            }
                        }
                    }
                    "spacing" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                // w:spacing w:val is in twips (1/20 pt)
                                style.character_spacing =
                                    val.parse::<f32>().ok().map(|v| v / 20.0);
                            }
                        }
                    }
                    "w" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                style.text_scale = val.parse::<f32>().ok();
                            }
                        }
                    }
                    "smallCaps" => {
                        style.small_caps = true;
                    }
                    "caps" => {
                        style.all_caps = true;
                    }
                    "shd" => {
                        // S704 (2026-06-30): compute the EFFECTIVE run-shading colour from
                        // val + fill + color (was: stored only `fill`, so a `pct15`/`FFFFFF`
                        // pattern stored white → invisible). val="clear" → fill; "solid" →
                        // color; "pctN" → N% color blended over fill (e.g. pct15 black/white
                        // = #D9D9D9). Stored in RunStyle.shading; the emit draws it as a
                        // run background.
                        let mut shd_val = String::new();
                        let mut shd_fill = String::new();
                        let mut shd_color = String::new();
                        for attr in e.attributes().flatten() {
                            match local_name(attr.key.as_ref()).as_str() {
                                "val" => shd_val = String::from_utf8_lossy(&attr.value).to_string(),
                                "fill" => shd_fill = String::from_utf8_lossy(&attr.value).to_string(),
                                "color" => shd_color = String::from_utf8_lossy(&attr.value).to_string(),
                                _ => {}
                            }
                        }
                        // opt-out OXI_S704_DISABLE → no run-shading background (pre-S704 pixels)
                        style.shading = if std::env::var("OXI_S704_DISABLE").is_ok() {
                            None
                        } else {
                            effective_shading_color(&shd_val, &shd_fill, &shd_color)
                        };
                    }
                    "bdr" => {
                        // S706 (2026-06-30): run/character border (w:bdr) — a box
                        // around the run's text. w:sz in 1/8 pt, w:space in pt.
                        let mut b_style = String::from("single");
                        let mut b_width = 0.5_f32;
                        let mut b_color: Option<String> = None;
                        let mut b_space = 0.0_f32;
                        for attr in e.attributes().flatten() {
                            let v = String::from_utf8_lossy(&attr.value).to_string();
                            match local_name(attr.key.as_ref()).as_str() {
                                "val" => b_style = v,
                                "sz" => b_width = v.parse::<f32>().unwrap_or(4.0) / 8.0,
                                "space" => b_space = v.parse::<f32>().unwrap_or(0.0),
                                "color" => if v != "auto" { b_color = Some(v) },
                                _ => {}
                            }
                        }
                        if b_style != "none" && b_style != "nil" {
                            style.run_border = Some(BorderDef {
                                style: b_style, width: b_width, color: b_color, space: b_space,
                            });
                        }
                    }
                    "rtl" => {
                        style.rtl = true;
                    }
                    "vanish" => {
                        style.vanish = true;
                    }
                    // w:webHidden (ECMA-376 §17.3.2.44) hides text ONLY in Web
                    // Layout view — print/PDF renders it normally. ToC tab-leaders
                    // and PAGEREF page-number runs are marked webHidden; treating
                    // it as vanish dropped every ToC page number.
                    "webHidden" => {}
                    "outline" => {
                        style.outline = true;
                    }
                    "shadow" => {
                        style.shadow = true;
                    }
                    "emboss" => {
                        style.emboss = true;
                    }
                    "imprint" => {
                        style.imprint = true;
                    }
                    "szCs" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                style.font_size_cs = val.parse::<f32>().ok().map(|v| v / 2.0);
                            }
                        }
                    }
                    "bCs" => {
                        style.bold_cs = true;
                    }
                    "iCs" => {
                        style.italic_cs = true;
                    }
                    "kern" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                style.kern = val.parse::<f32>().ok().map(|v| v / 2.0);
                            }
                        }
                    }
                    "fitText" => {
                        for attr in e.attributes().flatten() {
                            let ln = local_name(attr.key.as_ref());
                            if ln == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                style.fit_text = val.parse::<f32>().ok().map(|v| v / 20.0);
                            } else if ln == "id" {
                                let val = String::from_utf8_lossy(&attr.value);
                                style.fit_text_id = val.parse::<i64>().ok();
                            }
                        }
                    }
                    "rStyle" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                rstyle_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "eastAsianLayout" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            match key.as_str() {
                                "combine" => {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    style.combine = val.as_ref() != "0" && val.as_ref() != "false";
                                }
                                "combineBrackets" => {
                                    style.combine_brackets =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                }
                                "vert" => {
                                    let val = String::from_utf8_lossy(&attr.value);
                                    style.vert_in_horz = val.as_ref() != "0" && val.as_ref() != "false";
                                }
                                _ => {}
                            }
                        }
                    }
                    "position" => {
                        // Vertical position offset in half-points
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                style.position = val.parse::<f32>().ok().map(|v| v / 2.0);
                            }
                        }
                    }
                    "em" => {
                        // Emphasis mark / 圏点 (w:em)
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val != "none" {
                                    style.emphasis_mark = Some(val);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "rPr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // Apply character style (rStyle): matches Word output
    // char style properties are base, direct formatting overrides
    if let Some(ref id) = rstyle_id {
        if let Some(sdef) = styles.styles.get(id) {
            if let Some(ref rs) = sdef.paragraph.default_run_style {
                super::styles::merge_run_style(&mut style, rs);
            }
        }
    }

    Ok((style, rpr_change))
}

/// Parse a w:tbl element (table)
fn parse_table(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<Table, ParseError> {
    let mut rows = Vec::new();
    let mut style = TableStyle::default();
    let mut grid_columns = Vec::new();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "tblPr" if depth == 0 => {
                        style = parse_table_properties(reader)?;
                    }
                    "tblGrid" if depth == 0 => {
                        grid_columns = parse_table_grid(reader)?;
                    }
                    "tr" if depth == 0 => {
                        let row = parse_table_row(reader, ctx, styles)?;
                        rows.push(row);
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tbl" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // ECMA-376: Apply table style borders if table doesn't have explicit tblBorders.
    // Priority: tcBorders > tblStylePr > tblStyle tblBorders > direct tblBorders
    if !style.border {
        if let Some(ref style_id) = style.style_id {
            if let Some(tbl_style) = styles.table_styles.get(style_id) {
                if tbl_style.border {
                    style.border = true;
                    if tbl_style.has_inside_h {
                        style.has_inside_h = true;
                    }
                    if style.border_color.is_none() {
                        style.border_color = tbl_style.border_color.clone();
                    }
                    if style.border_width.is_none() {
                        style.border_width = tbl_style.border_width;
                    }
                    if style.border_style.is_none() {
                        style.border_style = tbl_style.border_style.clone();
                    }
                }
            }
        }
    }

    // S531 (2026-06-09): cell-margin (and table-style paragraph) inheritance is
    // INDEPENDENT of whether the table declares its own borders. It was wrongly
    // nested inside `if !style.border` above, so any bordered table that takes
    // its cellMar from a tblStyle (e.g. 683f's single-cell `af`-styled 解説 table:
    // af defines tblCellMar left/right=108tw, the table itself has explicit
    // tblBorders) dropped the inherited 108tw cellMar and fell back to the 4.95pt
    // default — making the cell wrap budget ~10.8pt too wide (+1 char/line).
    if let Some(ref style_id) = style.style_id {
        if let Some(tbl_style) = styles.table_styles.get(style_id) {
            if style.default_cell_margins.is_none() {
                style.default_cell_margins = tbl_style.default_cell_margins.clone();
            }
            if style.para_style.is_none() {
                style.para_style = tbl_style.para_style.clone();
            }
        }
    }

    // Apply tblStylePr conditional formatting to cells
    if let Some(ref style_id) = style.style_id {
        if let Some(cond_fmts) = styles.table_conditional_formats.get(style_id) {
            let look = style.tbl_look.unwrap_or_default();
            let num_rows = rows.len();
            for (row_idx, row) in rows.iter_mut().enumerate() {
                let num_cols = row.cells.len();
                for (col_idx, cell) in row.cells.iter_mut().enumerate() {
                    // Determine which conditional format applies (priority: corner > row/col > band)
                    let cond_key = resolve_conditional_type(
                        row_idx, col_idx, num_rows, num_cols, &look,
                    );
                    if let Some(key) = cond_key {
                        if let Some(fmt) = cond_fmts.get(key) {
                            // Apply shading if cell doesn't have explicit shading
                            if cell.shading.is_none() {
                                cell.shading = fmt.shading.clone();
                            }
                            // Apply borders if cell doesn't have explicit borders
                            if cell.borders.is_none() {
                                cell.borders = fmt.borders.clone();
                            }
                            // Apply run-level properties (bold, color) to cell paragraphs
                            if fmt.bold.is_some() || fmt.color.is_some() {
                                for block in &mut cell.blocks {
                                    if let Block::Paragraph(para) = block {
                                        for run in &mut para.runs {
                                            if let Some(b) = fmt.bold {
                                                if !run.style.bold { run.style.bold = b; }
                                            }
                                            if let Some(ref c) = fmt.color {
                                                if run.style.color.is_none() {
                                                    run.style.color = Some(c.clone());
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    Ok(Table { rows, style, grid_columns })
}

/// Determine which tblStylePr condition type applies to a cell.
/// Returns the highest-priority condition key (corners > first/last row/col > bands).
fn resolve_conditional_type(
    row_idx: usize, col_idx: usize,
    num_rows: usize, num_cols: usize,
    look: &crate::ir::TableLook,
) -> Option<&'static str> {
    let is_first_row = row_idx == 0 && look.first_row;
    let is_last_row = row_idx == num_rows - 1 && look.last_row;
    let is_first_col = col_idx == 0 && look.first_column;
    let is_last_col = col_idx == num_cols - 1 && look.last_column;

    // Corner cells (highest priority)
    if is_first_row && is_last_col { return Some("neCell"); }
    if is_first_row && is_first_col { return Some("nwCell"); }
    if is_last_row && is_last_col { return Some("seCell"); }
    if is_last_row && is_first_col { return Some("swCell"); }

    // First/last row (higher than column)
    if is_first_row { return Some("firstRow"); }
    if is_last_row { return Some("lastRow"); }

    // First/last column
    if is_first_col { return Some("firstCol"); }
    if is_last_col { return Some("lastCol"); }

    // Banded rows/columns
    if look.banded_rows {
        let band_size = look.row_band_size.max(1) as usize;
        // Adjust row index: skip header row for banding count
        let banding_row = if look.first_row { row_idx.saturating_sub(1) } else { row_idx };
        let band_index = banding_row / band_size;
        if band_index % 2 == 0 {
            return Some("band1Horz");
        } else {
            return Some("band2Horz");
        }
    }
    if look.banded_columns {
        let band_size = look.col_band_size.max(1) as usize;
        let banding_col = if look.first_column { col_idx.saturating_sub(1) } else { col_idx };
        let band_index = banding_col / band_size;
        if band_index % 2 == 0 {
            return Some("band1Vert");
        } else {
            return Some("band2Vert");
        }
    }

    None
}

/// Parse w:tblGrid element — extract gridCol widths (twips → points)
fn parse_table_grid(reader: &mut Reader<&[u8]>) -> Result<Vec<f32>, ParseError> {
    let mut columns = Vec::new();
    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tblGridChange" {
                    drain_element(reader, "tblGridChange")?;
                    continue;
                }
                // gridCol Start (rare but legal) — handled in next branch by also matching Start.
                if local == "gridCol" {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == "w" {
                            if let Ok(val) = std::str::from_utf8(&attr.value) {
                                if let Ok(twips) = val.parse::<f32>() {
                                    columns.push(twips / 20.0);
                                }
                            }
                        }
                    }
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                if local == "gridCol" {
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        if key == "w" {
                            if let Ok(val) = std::str::from_utf8(&attr.value) {
                                if let Ok(twips) = val.parse::<f32>() {
                                    columns.push(twips / 20.0); // twips to points
                                }
                            }
                        }
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tblGrid" {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(columns)
}

/// Parse w:tblPr (table properties)
fn parse_table_properties(reader: &mut Reader<&[u8]>) -> Result<TableStyle, ParseError> {
    let mut style = TableStyle::default();
    let mut depth = 0;
    let mut in_borders = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tblPrChange" {
                    // Revision history — drain so the prior tblPr inside
                    // doesn't leak into the current style.
                    drain_element(reader, "tblPrChange")?;
                    continue;
                }
                if local == "tblBorders" {
                    // Don't set border=true here; individual border elements check val!=none
                    in_borders = true;
                    style.explicit_borders = true;
                } else if local == "tblCellMar" {
                    // Parse default cell margins
                    let mut margins = CellMargins { top: None, bottom: None, left: None, right: None };
                    loop {
                        match reader.read_event() {
                            Ok(Event::Empty(me)) => {
                                let ml = local_name(me.name().as_ref());
                                let mut w_val: Option<f32> = None;
                                for attr in me.attributes().flatten() {
                                    if local_name(attr.key.as_ref()) == "w" {
                                        w_val = String::from_utf8_lossy(&attr.value).parse::<f32>().ok().map(|v| v / 20.0);
                                    }
                                }
                                match ml.as_str() {
                                    "top" => margins.top = w_val,
                                    "bottom" => margins.bottom = w_val,
                                    "left" | "start" => margins.left = w_val,
                                    "right" | "end" => margins.right = w_val,
                                    _ => {}
                                }
                            }
                            Ok(Event::End(ee)) => {
                                if local_name(ee.name().as_ref()) == "tblCellMar" { break; }
                            }
                            Ok(Event::Eof) => break,
                            _ => {}
                        }
                    }
                    style.default_cell_margins = Some(margins);
                    // S418: this tblCellMar is directly in THIS table's tblPr
                    // (author-declared), not inherited from a tblStyle — the
                    // precise S412 cellMar wrap-budget discriminator. Mirrors
                    // explicit_borders above.
                    style.has_explicit_cellmar = true;
                    // Don't increment depth — we consumed the End event
                    continue;
                }
                depth += 1;
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tblBorders" {
                    in_borders = false;
                }
                if local == "tblPr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "top" | "left" | "bottom" | "right" | "insideH" | "insideV" | "start" | "end"
                        if in_borders || depth == 0 =>
                    {
                        // Check if border val is "none" or "nil" — if so, don't set border=true
                        let mut is_none = false;
                        let mut border_color_val = None;
                        let mut border_sz_val = None;
                        // S480: capture the actual w:val (single/dashed/dotted/
                        // dashDotStroked/...) instead of hardcoding "single", so
                        // table-level decorative border styles reach the renderer.
                        let mut border_style_val: Option<String> = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "val" => {
                                    if val == "none" || val == "nil" {
                                        is_none = true;
                                    } else {
                                        border_style_val = Some(val.to_string());
                                    }
                                }
                                "color" => {
                                    if val != "auto" {
                                        border_color_val = Some(val.to_string());
                                    }
                                }
                                "sz" => {
                                    border_sz_val = val.parse::<f32>().ok().map(|v| v / 8.0);
                                }
                                _ => {}
                            }
                        }
                        if !is_none {
                            style.border = true;
                            if local == "insideH" {
                                style.has_inside_h = true;
                            }
                            if style.border_color.is_none() {
                                style.border_color = border_color_val;
                            }
                            if style.border_width.is_none() {
                                style.border_width = border_sz_val;
                            }
                            // Prefer a non-"single" decorative style if any edge
                            // declares one (uniform-border tables — the common
                            // case + a1d6e4 dashDotStroked); else keep "single".
                            let v = border_style_val.unwrap_or_else(|| "single".to_string());
                            if style.border_style.is_none()
                                || style.border_style.as_deref() == Some("single")
                            {
                                style.border_style = Some(v);
                            }
                        }
                    }
                    "tblW" => {
                        let mut w_val = None;
                        let mut w_type = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "w" => w_val = Some(val),
                                "type" => w_type = Some(val),
                                _ => {}
                            }
                        }
                        style.width_type = w_type.clone();
                        if let Some(ref wv) = w_val {
                            match w_type.as_deref() {
                                Some("dxa") => {
                                    style.width = wv.parse::<f32>().ok().map(|v| v / 20.0);
                                }
                                Some("pct") => {
                                    // Percentage stored as 50ths of a percent
                                    style.width = wv.parse::<f32>().ok().map(|v| v / 50.0);
                                }
                                _ => {}
                            }
                        }
                    }
                    "jc" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                style.alignment = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "tblStyle" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                style.style_id = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "tblStyleRowBandSize" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                if let Some(ref mut look) = style.tbl_look {
                                    look.row_band_size = String::from_utf8_lossy(&attr.value).parse().unwrap_or(1);
                                }
                            }
                        }
                    }
                    "tblStyleColBandSize" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                if let Some(ref mut look) = style.tbl_look {
                                    look.col_band_size = String::from_utf8_lossy(&attr.value).parse().unwrap_or(1);
                                }
                            }
                        }
                    }
                    "tblInd" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            if key == "w" {
                                style.indent = val.parse::<f32>().ok().map(|v| v / 20.0);
                            }
                        }
                    }
                    "tblCellSpacing" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            if key == "w" {
                                style.cell_spacing = val.parse::<f32>().ok().map(|v| v / 20.0);
                            }
                        }
                    }
                    "tblLook" => {
                        let mut look = TableLook::default();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "firstRow" => look.first_row = val.as_ref() == "1",
                                "lastRow" => look.last_row = val.as_ref() == "1",
                                "firstColumn" => look.first_column = val.as_ref() == "1",
                                "lastColumn" => look.last_column = val.as_ref() == "1",
                                "noHBand" => look.banded_rows = val.as_ref() != "1",
                                "noVBand" => look.banded_columns = val.as_ref() != "1",
                                "val" => {
                                    // Hex bitmask fallback (e.g. "04A0")
                                    if let Ok(v) = u32::from_str_radix(&val, 16) {
                                        look.first_row = v & 0x0020 != 0;
                                        look.last_row = v & 0x0040 != 0;
                                        look.first_column = v & 0x0080 != 0;
                                        look.last_column = v & 0x0100 != 0;
                                        look.banded_rows = v & 0x0200 == 0; // noHBand inverted
                                        look.banded_columns = v & 0x0400 == 0; // noVBand inverted
                                    }
                                }
                                _ => {}
                            }
                        }
                        style.tbl_look = Some(look);
                    }
                    "tblLayout" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "type" {
                                style.layout = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "tblpPr" => {
                        let mut tp = TablePosition {
                            x: 0.0, y: 0.0,
                            h_anchor: None, v_anchor: None, h_align: None,
                            left_from_text: 0.0, right_from_text: 0.0,
                            top_from_text: 0.0, bottom_from_text: 0.0,
                        };
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "tblpX" => tp.x = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                "tblpY" => tp.y = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                "tblpXSpec" => tp.h_align = Some(val.to_string()),
                                "horzAnchor" => tp.h_anchor = Some(val.to_string()),
                                "vertAnchor" => tp.v_anchor = Some(val.to_string()),
                                "leftFromText" => tp.left_from_text = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                "rightFromText" => tp.right_from_text = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                "topFromText" => tp.top_from_text = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                "bottomFromText" => tp.bottom_from_text = val.parse::<f32>().unwrap_or(0.0) / 20.0,
                                _ => {}
                            }
                        }
                        style.position = Some(tp);
                    }
                    "top" | "bottom" | "left" | "right" | "start" | "end"
                        if !in_borders =>
                    {
                        // tblCellMar children (when not inside tblBorders)
                        // These appear inside <w:tblCellMar> which is a Start element
                        // but we handle them as Empty within depth tracking
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(style)
}

/// Parse a w:tr element (table row)
fn parse_table_row(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<TableRow, ParseError> {
    let mut cells = Vec::new();
    let mut height: Option<f32> = None;
    let mut height_rule: Option<String> = None;
    let mut header = false;
    let mut cant_split = false;
    let mut grid_before: u32 = 0;
    let mut cell_margins_override: Option<CellMargins> = None;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "trPrChange" => {
                        drain_element(reader, "trPrChange")?;
                        continue;
                    }
                    "tc" if depth == 0 => {
                        let cell = parse_table_cell(reader, ctx, styles)?;
                        cells.push(cell);
                    }
                    "tblPrEx" if depth == 0 => {
                        // Row-level table property exceptions — parse tblCellMar override
                        let mut ex_depth = 0u32;
                        loop {
                            match reader.read_event()? {
                                Event::Start(se) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "tblCellMar" && ex_depth == 0 {
                                        let mut margins = CellMargins { top: None, bottom: None, left: None, right: None };
                                        loop {
                                            match reader.read_event() {
                                                Ok(Event::Empty(me)) => {
                                                    let ml = local_name(me.name().as_ref());
                                                    let mut w_val: Option<f32> = None;
                                                    for attr in me.attributes().flatten() {
                                                        if local_name(attr.key.as_ref()) == "w" {
                                                            w_val = String::from_utf8_lossy(&attr.value).parse::<f32>().ok().map(|v| v / 20.0);
                                                        }
                                                    }
                                                    match ml.as_str() {
                                                        "top" => margins.top = w_val,
                                                        "bottom" => margins.bottom = w_val,
                                                        "left" | "start" => margins.left = w_val,
                                                        "right" | "end" => margins.right = w_val,
                                                        _ => {}
                                                    }
                                                }
                                                Ok(Event::End(ee)) => {
                                                    if local_name(ee.name().as_ref()) == "tblCellMar" { break; }
                                                }
                                                Ok(Event::Eof) => break,
                                                _ => {}
                                            }
                                        }
                                        cell_margins_override = Some(margins);
                                        continue;
                                    }
                                    ex_depth += 1;
                                }
                                Event::End(se) => {
                                    if local_name(se.name().as_ref()) == "tblPrEx" && ex_depth == 0 { break; }
                                    if ex_depth > 0 { ex_depth -= 1; }
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "trHeight" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            if key == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                height = val.parse::<f32>().ok().map(|v| v / 20.0);
                            }
                            if key == "hRule" {
                                height_rule = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "tblHeader" => { header = true; }
                    "cantSplit" => { cant_split = true; }
                    "gridBefore" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                grid_before = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(TableRow { cells, height, height_rule, header, cant_split, grid_before, cell_margins_override })
}

/// Parse a w:tc element (table cell)
fn parse_table_cell(reader: &mut Reader<&[u8]>, ctx: &ParseContext, styles: &StyleSheet) -> Result<TableCell, ParseError> {
    let mut blocks = Vec::new();
    let mut cell_props = CellProperties::default();
    // S486: collect floating text boxes/shapes anchored inside the cell (was
    // dropped — see TableCell::cell_text_boxes). Foundational data-preservation;
    // the in-cell anchor-resolution render is the deferred follow-up step.
    let mut cell_text_boxes: Vec<TextBox> = Vec::new();
    let mut cell_shapes: Vec<Shape> = Vec::new();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "p" if depth == 0 => {
                        let pr = parse_paragraph(reader, ctx, styles, false)?;
                        // S486: preserve in-cell floating text boxes/shapes
                        // (previously discarded). S488: stamp each preserved
                        // text box with its anchor paragraph's index within THIS
                        // cell's block list (= blocks.len() now, before the
                        // paragraph is pushed below). The cell render loop uses
                        // this to resolve relV="paragraph" against the specific
                        // anchoring paragraph's top (COM-confirmed 1636d28: the
                        // 3 cell text boxes anchor to paragraphs at y=431/710/737
                        // inside one cell — a single cell-top origin mis-places
                        // all three). pr.text_boxes carry a body-relative index
                        // from parse_paragraph; overwrite with the cell-relative one.
                        let cell_para_block_idx = blocks.len();
                        // S838 (2026-07-14): a CELL paragraph hosting an INLINE
                        // visual drawing (the S535 synthetic-inline signature —
                        // hmrc's NI/DOB checkbox strips inside the personal-
                        // details table) must reserve the drawing's extent like
                        // the body S773 rule. Collect the max cy while moving
                        // the tbs; an EMPTY host para is then replaced by a
                        // height-only Block::Image (the S536/S537b image-only
                        // path: cell line = extent EXACTLY = the pinned Word
                        // rule), instead of contributing one empty text line.
                        let mut s838_cy: f32 = 0.0;
                        let mut s838_cx: f32 = 0.0;
                        for mut tb in pr.text_boxes {
                            if tb.blocks.is_empty()
                                && matches!(tb.wrap_type, Some(crate::ir::WrapType::None))
                                && tb.position.as_ref().map_or(false, |p| p.x == 0.0 && p.y == 0.0
                                    && p.h_relative.as_deref() == Some("column")
                                    && p.v_relative.as_deref() == Some("paragraph"))
                                && tb.height > s838_cy
                            {
                                s838_cy = tb.height;
                                s838_cx = tb.width;
                            }
                            tb.anchor_block_index = cell_para_block_idx;
                            cell_text_boxes.push(tb);
                        }
                        cell_shapes.extend(pr.shapes);
                        // S536 (2026-06-10): suppress the EMPTY host paragraph of a
                        // cell inline image — in Word the image IS the paragraph's
                        // line, so pushing both the (empty) Block::Paragraph and the
                        // forwarded Block::Image double-counts one line. COM-measured
                        // on 3a4f's calendar table: Word rendered height 448.5pt
                        // (Tables(33) cell top 120.0 -> after 568.5) vs Oxi row
                        // 466.25 = +17.75 ≈ exactly the pict para's 17.5pt line
                        // (spacing line=350 atLeast). Mirrors the S525 math_only
                        // empty-paragraph suppression (ooxml.rs ~823) for the
                        // image-in-cell case.
                        let image_only = std::env::var("OXI_S331_DISABLE").is_err()
                            && !pr.inline_images.is_empty()
                            && pr.paragraph.runs.iter().all(|r| r.text.is_empty())
                            && pr.math_blocks.is_empty();
                        // S838: visual-only cell para (inline vector group, no
                        // text/images/math) -> height-only placeholder Image.
                        let s838_visual_only = std::env::var("OXI_S838_DISABLE").is_err()
                            && s838_cy > 0.0
                            && pr.inline_images.is_empty()
                            && pr.paragraph.runs.iter().all(|r| r.text.is_empty())
                            && pr.math_blocks.is_empty();
                        if s838_visual_only {
                            blocks.push(Block::Image(crate::ir::Image {
                                data: Vec::new(),
                                width: s838_cx,
                                height: s838_cy,
                                alt_text: None,
                                content_type: None,
                                position: None,
                                wrap_type: None,
                                crop: None,
                                anchor_block_index: 0,
                                relative_height: 0,
                                behind_doc: false,
                            }));
                        } else if !image_only {
                            blocks.push(Block::Paragraph(pr.paragraph));
                        }
                        // S331 (2026-05-26): forward inline images from cell
                        // paragraphs so cell height includes drawing. Body-level
                        // parser does this at line 736/781; cell parser was
                        // dropping pr.inline_images, pr.math_blocks etc.,
                        // causing cells with inline drawings to be shorter than
                        // Word renders them.
                        // S533 (2026-06-10): default ON (opt-out OXI_S331_DISABLE).
                        // The S331 forward had stayed env-gated OFF; combined with
                        // the missing layout arms (cell placement + row-height
                        // estimate, added in S533) an image-bearing cell collapsed
                        // to its text height — 3a4f p34's 321.75pt calendar EMF
                        // cell rendered ~28pt, the Phase-1 sole-FAIL root cause.
                        if std::env::var("OXI_S331_DISABLE").is_err() {
                            for mb in pr.math_blocks {
                                blocks.push(Block::Math(mb));
                            }
                            blocks.extend(pr.inline_images);
                            // S331b (2026-05-26): also forward floating (anchored)
                            // drawings as Block::Image so cell height accounts
                            // for them. Word wraps anchored drawings inside
                            // cells differently from body (wrap=topAndBottom
                            // is dominant inside cells), so treating them as
                            // inline-flow images is a reasonable approximation.
                            for fimg in pr.floating_images {
                                blocks.push(Block::Image(fimg));
                            }
                        }
                    }
                    "tbl" if depth == 0 => {
                        let table = parse_table(reader, ctx, styles)?;
                        blocks.push(Block::Table(table));
                    }
                    "tcPr" if depth == 0 => {
                        cell_props = parse_cell_properties(reader)?;
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tc" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            // S864: preserve a self-closing empty paragraph inside
            // a table cell. The Start/End path above handles `<w:p>...</w:p>`,
            // but quick-xml reports `<w:p/>` as Event::Empty. Dropping it makes
            // row height depend on the producer's equivalent XML spelling.
            Event::Empty(e) if depth == 0
                && local_name(e.name().as_ref()) == "p"
                && std::env::var("OXI_S864_DISABLE").is_err() =>
            {
                blocks.push(Block::Paragraph(empty_para_with_defaults(styles)));
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // S864: Word applies NormalWeb's style-level autospacing to a table cell
    // containing a long manual-break list, while S675 keeps style-level HTML
    // spacing inert in ordinary body/cell content. At the boundary immediately
    // after the long list, Word suppresses it again (explicit spacing remains).
    // The >=10 w:br signature keeps this correction isolated from generic Web
    // style tables, which have deliberately different S675 behaviour.
    if std::env::var("OXI_S864_DISABLE").is_err() {
        let has_long_manual_breaks = blocks.iter().any(|b| matches!(b, Block::Paragraph(p)
            if p.runs.iter().map(|r| r.text.matches('\n').count()).sum::<usize>() >= 10));
        if has_long_manual_breaks {
            for block in &mut blocks {
                if let Block::Paragraph(p) = block {
                    if let Some(style_id) = p.style.style_id.as_deref() {
                        if let Some(defined) = styles.styles.get(style_id) {
                            p.style.before_autospacing = defined.paragraph.before_autospacing;
                            p.style.after_autospacing = defined.paragraph.after_autospacing;
                        }
                    }
                }
            }
        }
        for i in 1..blocks.len() {
            let long_manual_breaks = matches!(&blocks[i - 1], Block::Paragraph(p)
                if p.runs.iter().map(|r| r.text.matches('\n').count()).sum::<usize>() >= 10);
            if long_manual_breaks {
                let (before, after) = blocks.split_at_mut(i);
                if let Block::Paragraph(prev) = &mut before[i - 1] {
                    prev.style.after_autospacing = false;
                }
                if let Block::Paragraph(next) = &mut after[0] {
                    next.style.before_autospacing = false;
                }
            }
        }
    }

    Ok(TableCell {
        blocks,
        width: cell_props.width,
        grid_span: cell_props.grid_span,
        v_merge: cell_props.v_merge,
        shading: cell_props.shading,
        v_align: cell_props.v_align,
        borders: cell_props.borders,
        margins: cell_props.margins,
        text_direction: cell_props.text_direction,
        cell_text_boxes,
        cell_shapes,
        hide_mark: cell_props.hide_mark,
    })
}

#[derive(Default)]
struct CellProperties {
    width: Option<f32>,
    grid_span: u32,
    v_merge: Option<String>,
    shading: Option<String>,
    v_align: Option<String>,
    borders: Option<CellBorders>,
    margins: Option<CellMargins>,
    text_direction: Option<String>,
    hide_mark: bool, // S751: w:hideMark
}

/// Parse w:tcPr (table cell properties)
fn parse_cell_properties(reader: &mut Reader<&[u8]>) -> Result<CellProperties, ParseError> {
    let mut props = CellProperties { grid_span: 1, ..Default::default() };
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "tcPrChange" => {
                        drain_element(reader, "tcPrChange")?;
                        continue;
                    }
                    "vMerge" => {
                        let mut val = "continue".to_string();
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                val = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        props.v_merge = Some(val);
                        depth += 1;
                    }
                    "tcBorders" if depth == 0 => {
                        props.borders = Some(parse_cell_borders(reader)?);
                    }
                    "tcMar" if depth == 0 => {
                        props.margins = Some(parse_cell_margins(reader)?);
                    }
                    _ => {
                        depth += 1;
                    }
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "tcW" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "w" {
                                let val = String::from_utf8_lossy(&attr.value);
                                props.width = val.parse::<f32>().ok().map(|v| v / 20.0);
                            }
                        }
                    }
                    "gridSpan" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let val = String::from_utf8_lossy(&attr.value);
                                props.grid_span = val.parse().unwrap_or(1);
                            }
                        }
                    }
                    "vMerge" => {
                        let mut val = "continue".to_string();
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                val = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        props.v_merge = Some(val);
                    }
                    "shd" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "fill" {
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                if val != "auto" {
                                    props.shading = Some(val);
                                }
                            }
                        }
                    }
                    "vAlign" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                props.v_align = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    // S751: w:hideMark — the end-of-cell mark is excluded from
                    // row-height (an empty hideMark cell = zero content height).
                    "hideMark" => { props.hide_mark = true; }
                    "textDirection" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                props.text_direction = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "tcPr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(props)
}

/// Parse w:tcBorders
fn parse_cell_borders(reader: &mut Reader<&[u8]>) -> Result<CellBorders, ParseError> {
    let mut borders = CellBorders {
        top: None, bottom: None, left: None, right: None,
    };
    loop {
        match reader.read_event()? {
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                let bdr = parse_border_attrs(&e);
                // S482: a cell's EXPLICIT w:val="nil"/"none" edge SUPPRESSES the
                // table-level border on that edge — it is NOT the same as an absent
                // edge (which inherits the table border). parse_border_attrs returns
                // None for both; preserve explicit-nil as a {style:"none"} sentinel
                // (cell path ONLY — pBdr/pgBorders unaffected) so resolve_border can
                // suppress instead of falling through to the table border (31420af:
                // tcBorders with nil top/left over an all-single tblBorders drew
                // spurious top/left rules). Opt-out OXI_S482_DISABLE.
                let bdr = if bdr.is_none() && std::env::var("OXI_S482_DISABLE").is_err() {
                    let explicit_none = e.attributes().flatten().any(|a| {
                        local_name(a.key.as_ref()) == "val" && {
                            let v = String::from_utf8_lossy(&a.value);
                            v == "nil" || v == "none"
                        }
                    });
                    if explicit_none {
                        Some(BorderDef { style: "none".to_string(), width: 0.0, color: None, space: 0.0 })
                    } else { None }
                } else { bdr };
                match local.as_str() {
                    "top" => borders.top = bdr,
                    "bottom" => borders.bottom = bdr,
                    "left" | "start" => borders.left = bdr,
                    "right" | "end" => borders.right = bdr,
                    _ => {}
                }
            }
            Event::End(e) => {
                if local_name(e.name().as_ref()) == "tcBorders" { break; }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(borders)
}

/// Parse w:tcMar
fn parse_cell_margins(reader: &mut Reader<&[u8]>) -> Result<CellMargins, ParseError> {
    let mut margins = CellMargins {
        top: None, bottom: None, left: None, right: None,
    };
    loop {
        match reader.read_event()? {
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                let val = e.attributes().flatten()
                    .find(|a| local_name(a.key.as_ref()) == "w")
                    .and_then(|a| {
                        String::from_utf8_lossy(&a.value).parse::<f32>().ok().map(|v| v / 20.0)
                    });
                match local.as_str() {
                    "top" => margins.top = val,
                    "bottom" => margins.bottom = val,
                    "left" | "start" => margins.left = val,
                    "right" | "end" => margins.right = val,
                    _ => {}
                }
            }
            Event::End(e) => {
                if local_name(e.name().as_ref()) == "tcMar" { break; }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(margins)
}

/// A header/footer reference with its type
#[derive(Debug, Clone)]
struct HdrFtrRef {
    rel_id: String,
    ref_type: String, // "default", "first", "even"
}

/// Parsed section properties
struct SectionProperties {
    page_size: PageSize,
    margin: Margin,
    /// Document grid line pitch in points (from w:docGrid w:linePitch, twips/20)
    grid_line_pitch: Option<f32>,
    /// Character grid pitch in points (for linesAndChars mode)
    grid_char_pitch: Option<f32>,
    /// Raw charSpace value (1/4096 point units) for post-process recompute
    grid_char_space_raw: Option<i32>,
    /// docGrid exists but has no type attribute
    doc_grid_no_type: bool,
    /// docGrid type == "linesAndChars" (character grid)
    doc_grid_lines_and_chars: bool,
    /// Reference IDs for header parts (with type)
    header_refs: Vec<HdrFtrRef>,
    /// Reference IDs for footer parts (with type)
    footer_refs: Vec<HdrFtrRef>,
    /// Column layout
    columns: Option<ColumnLayout>,
    /// Whether this section has a different first page header/footer
    title_pg: bool,
    /// Section break type: "nextPage" (default), "continuous", "evenPage", "oddPage"
    section_type: Option<String>,
    /// Page number format (e.g. "decimal", "lowerRoman", "upperRoman", "lowerLetter", "upperLetter")
    page_number_format: Option<String>,
    /// Starting page number for this section
    page_number_start: Option<u32>,
    /// Page borders
    page_borders: Option<PageBorders>,
    /// Header distance from page top edge (w:pgMar header attr, in points)
    header_distance: Option<f32>,
    /// Footer distance from page bottom edge (w:pgMar footer attr, in points)
    footer_distance: Option<f32>,
    /// Bidirectional (RTL) section (`<w:bidi/>`). Drives right-to-left
    /// multi-column flow (first reading column = rightmost).
    bidi: bool,
    /// Section text direction (`<w:textDirection w:val="tbRl"/>`). "tbRl" =
    /// vertical writing (tategaki/縦書き): chars stack top-to-bottom, lines
    /// advance right-to-left.
    text_direction: Option<String>,
}

/// Parse w:sectPr (section properties - page size, margins, document grid)
fn parse_section_properties(
    reader: &mut Reader<&[u8]>,
) -> Result<SectionProperties, ParseError> {
    let mut page_size = PageSize::default();
    let mut margin = Margin::default();
    let mut grid_line_pitch: Option<f32> = None;
    let mut grid_char_pitch: Option<f32> = None;
    let mut char_space_section: Option<i32> = None;
    let mut doc_grid_no_type = false;
    let mut doc_grid_lines_and_chars = false;
    let mut header_refs: Vec<HdrFtrRef> = Vec::new();
    let mut footer_refs: Vec<HdrFtrRef> = Vec::new();
    let mut columns: Option<ColumnLayout> = None;
    let mut title_pg = false;
    let mut section_type: Option<String> = None;
    let mut page_number_format: Option<String> = None;
    let mut page_number_start: Option<u32> = None;
    let mut page_borders: Option<PageBorders> = None;
    let mut header_distance: Option<f32> = None;
    let mut footer_distance: Option<f32> = None;
    let mut bidi = false;
    let mut text_direction: Option<String> = None;
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "sectPrChange" {
                    drain_element(reader, "sectPrChange")?;
                    continue;
                }
                if local == "pgBorders" && depth == 0 {
                    // Parse page borders - child elements: top, bottom, left, right
                    let mut pb = PageBorders { top: None, bottom: None, left: None, right: None };
                    loop {
                        match reader.read_event() {
                            Ok(Event::Empty(be)) => {
                                let bl = local_name(be.name().as_ref());
                                let mut bdr_style = String::new();
                                let mut bdr_width = 0.0f32;
                                let mut bdr_color: Option<String> = None;
                                let mut bdr_space = 0.0f32;
                                for attr in be.attributes().flatten() {
                                    let key = local_name(attr.key.as_ref());
                                    let val = String::from_utf8_lossy(&attr.value);
                                    match key.as_str() {
                                        "val" => bdr_style = val.to_string(),
                                        "sz" => { bdr_width = val.parse::<f32>().unwrap_or(0.0) / 8.0; }
                                        "color" => {
                                            let c = val.to_string();
                                            if c != "auto" { bdr_color = Some(c); }
                                        }
                                        "space" => { bdr_space = val.parse::<f32>().unwrap_or(0.0); }
                                        _ => {}
                                    }
                                }
                                if bdr_style != "none" && bdr_style != "nil" && bdr_width > 0.0 {
                                    let def = BorderDef { style: bdr_style, width: bdr_width, color: bdr_color, space: bdr_space };
                                    match bl.as_str() {
                                        "top" => pb.top = Some(def),
                                        "bottom" => pb.bottom = Some(def),
                                        "left" => pb.left = Some(def),
                                        "right" => pb.right = Some(def),
                                        _ => {}
                                    }
                                }
                            }
                            Ok(Event::End(ee)) => {
                                if local_name(ee.name().as_ref()) == "pgBorders" { break; }
                            }
                            Ok(Event::Eof) => break,
                            _ => {}
                        }
                    }
                    if pb.top.is_some() || pb.bottom.is_some() || pb.left.is_some() || pb.right.is_some() {
                        page_borders = Some(pb);
                    }
                } else if local == "cols" && depth == 0 {
                    // w:cols as Start element — has child w:col elements
                    let mut num = 1u32;
                    let mut space: Option<f32> = None;
                    let mut equal_width = true;
                    let mut separator = false;
                    for attr in e.attributes().flatten() {
                        let key = local_name(attr.key.as_ref());
                        let val = String::from_utf8_lossy(&attr.value);
                        match key.as_str() {
                            "num" => { num = val.parse().unwrap_or(1); }
                            "space" => { space = val.parse::<f32>().ok().map(|v| v / 20.0); }
                            "equalWidth" => { equal_width = val.as_ref() != "0" && val.as_ref() != "false"; }
                            "sep" => { separator = val.as_ref() != "0" && val.as_ref() != "false"; }
                            _ => {}
                        }
                    }
                    // Parse child w:col elements
                    let mut col_defs = Vec::new();
                    loop {
                        match reader.read_event() {
                            Ok(Event::Empty(ce)) => {
                                let cl = local_name(ce.name().as_ref());
                                if cl == "col" {
                                    let mut col_w = 0.0f32;
                                    let mut col_space: Option<f32> = None;
                                    for attr in ce.attributes().flatten() {
                                        let key = local_name(attr.key.as_ref());
                                        let val = String::from_utf8_lossy(&attr.value);
                                        match key.as_str() {
                                            "w" => { col_w = val.parse::<f32>().unwrap_or(0.0) / 20.0; }
                                            "space" => { col_space = val.parse::<f32>().ok().map(|v| v / 20.0); }
                                            _ => {}
                                        }
                                    }
                                    col_defs.push(ColumnDef { width: col_w, space: col_space });
                                }
                            }
                            Ok(Event::End(ee)) => {
                                if local_name(ee.name().as_ref()) == "cols" {
                                    break;
                                }
                            }
                            Ok(Event::Eof) => break,
                            _ => {}
                        }
                    }
                    if num > 1 {
                        columns = Some(ColumnLayout { num, space, equal_width, separator, columns: col_defs });
                    }
                } else {
                    depth += 1;
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "pgSz" => {
                        let mut orient = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "w" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        page_size.width = v / 20.0;
                                    }
                                }
                                "h" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        page_size.height = v / 20.0;
                                    }
                                }
                                "orient" => orient = Some(val.to_string()),
                                _ => {}
                            }
                        }
                        // Landscape: ensure width > height
                        if orient.as_deref() == Some("landscape") && page_size.width < page_size.height {
                            std::mem::swap(&mut page_size.width, &mut page_size.height);
                        }
                    }
                    "pgMar" => {
                        // Top/bottom margins: round to 10tw (0.5pt).
                        // COM-confirmed (0e7a P208 y=56.50 = round_10tw(1134)=1130tw).
                        // Left/right margins: exact twips (no rounding).
                        // COM-confirmed (0e7a LeftMargin=53.85pt = 1077tw/20).
                        let to_pt = |tw: f32| -> f32 { tw / 20.0 };
                        let to_pt_round10 = |tw: f32| -> f32 { (tw / 10.0).round() * 10.0 / 20.0 };
                        let mut gutter = 0.0f32;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "top" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        margin.top = to_pt_round10(v);
                                    }
                                }
                                "bottom" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        // Bottom margin: exact twips (no 10tw rounding).
                                        // Word rounds top margin to 10tw for content start Y,
                                        // but uses exact bottom margin for page break limit.
                                        // COM-confirmed (0e7a): bottom=1134tw=56.7pt, limit=785.2pt
                                        margin.bottom = to_pt(v);
                                    }
                                }
                                "left" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        margin.left = to_pt(v);
                                    }
                                }
                                "right" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        margin.right = to_pt(v);
                                    }
                                }
                                "gutter" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        gutter = to_pt(v);
                                    }
                                }
                                "header" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        header_distance = Some(to_pt(v));
                                    }
                                }
                                "footer" => {
                                    if let Ok(v) = val.parse::<f32>() {
                                        footer_distance = Some(to_pt(v));
                                    }
                                }
                                _ => {}
                            }
                        }
                        // Gutter adds to left margin (default) or top margin (gutterTop not implemented yet)
                        if gutter > 0.0 {
                            margin.left += gutter;
                        }
                    }
                    "docGrid" => {
                        let mut grid_type = String::new();
                        let mut line_pitch = 0u32;
                        // §17.6.5 (Round 15, COM-confirmed): w:charSpace is a SIGNED
                        // integer (can be negative for compressed grid). u32 silently
                        // dropped negative values to 0, causing wrong pitch on docs
                        // with charSpace<0.
                        let mut char_space: Option<i32> = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "type" => grid_type = val.to_string(),
                                "linePitch" => {
                                    line_pitch = val.parse().unwrap_or(0);
                                }
                                "charSpace" => {
                                    char_space = Some(val.parse().unwrap_or(0));
                                }
                                _ => {}
                            }
                        }
                        // Only apply grid for "lines" or "linesAndChars" types
                        if (grid_type == "lines" || grid_type == "linesAndChars")
                            && line_pitch > 0
                        {
                            grid_line_pitch = Some(line_pitch as f32 / 20.0);
                        } else if grid_type.is_empty() && line_pitch > 0 {
                            // docGrid exists with linePitch but no type attribute.
                            // S571 (2026-06-14): Word STILL uses the linePitch for LINE
                            // SPACING even with no type attribute. RENDER-TRUTH (ikujidetail
                            // 育児介護 詳細版, compat=11, docGrid linePitch=286 no type):
                            // Word's body line pitch = 14.3pt (= 286tw); Oxi used the
                            // font-natural ~14.5pt → +0.2pt/line accumulating → each page
                            // fits ~1 fewer para → a per-page +1 over-count (+1×73). The
                            // doc_grid_no_type flag was DEAD (set but never consumed in
                            // layout) so the linePitch was simply dropped. Treat a no-type
                            // docGrid with linePitch as a LINE grid (use its linePitch).
                            doc_grid_no_type = true;
                            // ★S571 REFINED (2026-06-23): apply the no-type docGrid linePitch
                            // ONLY when it is NON-DEFAULT (≠360). linePitch=360 (18pt) is
                            // Word's AUTOMATIC default no-type docGrid (= "no real grid") —
                            // Word uses the NATURAL line height, NOT the 18pt pitch. The
                            // original S571 applied to ALL no-type linePitch, inflating the
                            // ~14pt natural body of every default-360 doc to 18pt → regressed
                            // ALL 112 word_png no-type-docGrid docs (the gen2 family) by
                            // -0.05..-0.08 SSIM (the corpus "mole-whacking", masked by the
                            // broken ssim_ab tool — git-bisect pinned b589873b; see
                            // [[ssim_ab_tool_was_broken]]). A CUSTOM linePitch (ikujidetail
                            // 286=14.3pt ≈ its MS Mincho natural) IS a real grid Word honors
                            // (needed for its Phase-1 pagination — full disable = ikujidetail
                            // PASS→FAIL). Opt-out OXI_S571_DISABLE; force-all OXI_S571_ALL=1.
                            if std::env::var("OXI_S571_DISABLE").is_err()
                                && (line_pitch != 360
                                    || std::env::var("OXI_S571_ALL").ok().as_deref() == Some("1"))
                            {
                                grid_line_pitch = Some(line_pitch as f32 / 20.0);
                            }
                        }
                        // linesAndChars: compute character grid pitch
                        // COM-confirmed (2026-04-03): charGrid is active even without charSpace.
                        // Formula: raw_pitch = default_font_size + charSpace/4096
                        //          charsLine = floor(contentWidth / raw_pitch)
                        //          actual_pitch = contentWidth / charsLine
                        // charSpace unit: 1/4096 of a point (ECMA-376 §17.6.5)
                        if grid_type == "linesAndChars" {
                            doc_grid_lines_and_chars = true;
                            // charGrid raw_pitch uses the document's default font size.
                            // This comes from Normal style's sz, or rPrDefault sz, or 10.5pt fallback.
                            // Stored in SectionProperties and resolved by the caller post-parse.
                            let default_font_size = 10.5_f32; // placeholder; overridden by caller
                            let char_space_pt = char_space.map(|cs| cs as f32 / 4096.0).unwrap_or(0.0);
                            let raw_pitch = default_font_size + char_space_pt;
                            let content_w = page_size.width - margin.left - margin.right;
                            if raw_pitch > 0.0 && content_w > 0.0 {
                                // S466 (2026-05-31): Word uses the UN-stretched raw_pitch
                                // as the char advance and leaves the line-end remainder as
                                // trailing space; COM (cg_mincho_10.5 charSpace=1453) shows
                                // rendered advance 10.875 ≈ raw 10.855, NOT the stretched
                                // content_w/chars_line=11.075. The stretch over-expands when
                                // the remainder is large (default=10.5, charSpace=1453: 9.7pt
                                // remainder), causing over-wrap (b837 7->9 pages). Only the
                                // small-remainder cases (10/9pt) made stretch≈raw, which is
                                // why the 2026-04-03 stretch "COM-confirmation" held there.
                                if std::env::var("OXI_S466_DISABLE").is_err() {
                                    grid_char_pitch = Some(raw_pitch);
                                } else {
                                    let chars_line = (content_w / raw_pitch).floor().max(1.0);
                                    grid_char_pitch = Some(content_w / chars_line);
                                }
                            }
                            // Save raw charSpace so the caller's post-process can recompute
                            // with the correct default_font_size while preserving charSpace.
                            char_space_section = char_space;
                        }
                    }
                    "headerReference" => {
                        let mut rel_id = String::new();
                        let mut ref_type = "default".to_string();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "id" => rel_id = val,
                                "type" => ref_type = val,
                                _ => {}
                            }
                        }
                        if !rel_id.is_empty() {
                            header_refs.push(HdrFtrRef { rel_id, ref_type });
                        }
                    }
                    "footerReference" => {
                        let mut rel_id = String::new();
                        let mut ref_type = "default".to_string();
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "id" => rel_id = val,
                                "type" => ref_type = val,
                                _ => {}
                            }
                        }
                        if !rel_id.is_empty() {
                            footer_refs.push(HdrFtrRef { rel_id, ref_type });
                        }
                    }
                    "titlePg" => {
                        title_pg = true;
                    }
                    "textDirection" => {
                        // Section-level text direction (e.g. "tbRl" = vertical
                        // top-to-bottom, right-to-left = tategaki/縦書き).
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                text_direction = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "bidi" => {
                        // CT_OnOff: absent val (or "1"/"true"/"on") => true
                        let mut enabled = true;
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                let v = String::from_utf8_lossy(&attr.value);
                                enabled = v != "0" && v != "false" && v != "off";
                            }
                        }
                        bidi = enabled;
                    }
                    "type" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "val" {
                                section_type = Some(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    "pgNumType" => {
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "fmt" => {
                                    page_number_format = Some(val.to_string());
                                }
                                "start" => {
                                    page_number_start = val.parse::<u32>().ok();
                                }
                                _ => {}
                            }
                        }
                    }
                    "cols" => {
                        let mut num = 1u32;
                        let mut space: Option<f32> = None;
                        let mut equal_width = true;
                        let mut separator = false;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value);
                            match key.as_str() {
                                "num" => { num = val.parse().unwrap_or(1); }
                                "space" => { space = val.parse::<f32>().ok().map(|v| v / 20.0); }
                                "equalWidth" => { equal_width = val.as_ref() != "0" && val.as_ref() != "false"; }
                                "sep" => { separator = val.as_ref() != "0" && val.as_ref() != "false"; }
                                _ => {}
                            }
                        }
                        if num > 1 {
                            columns = Some(ColumnLayout { num, space, equal_width, separator, columns: Vec::new() });
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "sectPr" && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(SectionProperties {
        page_size,
        margin,
        grid_line_pitch,
        grid_char_pitch,
        grid_char_space_raw: char_space_section,
        doc_grid_no_type,
        doc_grid_lines_and_chars,
        header_refs,
        footer_refs,
        columns,
        title_pg,
        section_type,
        page_number_format,
        page_number_start,
        page_borders,
        header_distance,
        footer_distance,
        bidi,
        text_direction,
    })
}

/// WATERMARK (2026-07-08): extract a VML WordArt watermark from a header
/// part — the Word idiom `<v:shape type="#_x0000_t136" style="…rotation:315…"
/// fillcolor="gray [1629]"><v:textpath … string="SAMPLE"/></v:shape>`
/// (PowerPlusWaterMarkObject). Only a textpath WITH a `string` attribute
/// counts (the shapetype defines a bare `<v:textpath on="t"/>`). String-level
/// scan: the markup is machine-generated and the JP corpus carries ZERO
/// t136/textpath docs (scanned), so this is corpus-inert by construction.
fn extract_vml_watermark(xml: &str) -> Option<crate::ir::Watermark> {
    // find a v:textpath with a string attribute
    let mut search_from = 0usize;
    let (tp_start, tp_tag) = loop {
        let i = xml[search_from..].find("<v:textpath ")? + search_from;
        let end = xml[i..].find('>')? + i;
        let tag = &xml[i..=end];
        if tag.contains("string=\"") {
            break (i, tag.to_string());
        }
        search_from = end + 1;
    };
    let attr = |tag: &str, name: &str| -> Option<String> {
        let pat = format!("{}=\"", name);
        let s = tag.find(&pat)? + pat.len();
        let e = tag[s..].find('"')? + s;
        Some(tag[s..e].to_string())
    };
    let text = attr(&tp_tag, "string")?;
    if text.trim().is_empty() {
        return None;
    }
    // the owning v:shape is the last one opened before the textpath
    let shape_start = xml[..tp_start].rfind("<v:shape ")?;
    let shape_end = xml[shape_start..].find('>')? + shape_start;
    let shape_tag = &xml[shape_start..=shape_end];
    let style = attr(shape_tag, "style").unwrap_or_default();
    let style_val = |key: &str| -> Option<String> {
        for part in style.split(';') {
            let mut kv = part.splitn(2, ':');
            if kv.next()?.trim() == key {
                return kv.next().map(|v| v.trim().to_string());
            }
        }
        None
    };
    let pt = |v: Option<String>| -> Option<f32> {
        v.and_then(|s| s.trim_end_matches("pt").parse::<f32>().ok())
    };
    let width = pt(style_val("width"))?;
    let height = pt(style_val("height"))?;
    let rotation = style_val("rotation")
        .and_then(|s| s.parse::<f32>().ok())
        .unwrap_or(0.0);
    // fillcolor: "#RRGGBB" | "name" | "name [themeidx]"
    let color = attr(shape_tag, "fillcolor").map(|c| {
        let base = c.split_whitespace().next().unwrap_or("").trim_start_matches('#');
        match base.to_ascii_lowercase().as_str() {
            "gray" | "grey" => "808080".to_string(),
            "silver" => "C0C0C0".to_string(),
            "black" => "000000".to_string(),
            "red" => "FF0000".to_string(),
            "blue" => "0000FF".to_string(),
            h if h.len() == 6 && h.chars().all(|c| c.is_ascii_hexdigit()) => h.to_uppercase(),
            _ => "C0C0C0".to_string(),
        }
    });
    // textpath style font-family:"CG Times" (entities already decoded? raw
    // part text has &quot; — strip both)
    let font_family = attr(&tp_tag, "style").and_then(|s| {
        for part in s.split(';') {
            let mut kv = part.splitn(2, ':');
            if kv.next().map(|k| k.trim()) == Some("font-family") {
                let v = kv.next().unwrap_or("").trim()
                    .replace("&quot;", "").replace('"', "");
                return Some(v);
            }
        }
        None
    });
    Some(crate::ir::Watermark { text, width, height, rotation, color, font_family })
}

// Parse a header or footer XML part (w:hdr or w:ftr element)
/// Self-closing `<w:p/>` — empty paragraph with no children. Applies the
/// default Normal style + docDefaults, matching parse_paragraph(). Shared by
/// the body parser and parse_header_footer_xml (S806p: the header/footer
/// parser previously DROPPED Event::Empty paragraphs).
fn empty_para_with_defaults(styles: &StyleSheet) -> Paragraph {
    let mut style = ParagraphStyle::default();
    // Apply Normal style (common IDs: "a" for Japanese, "Normal" for English)
    if let Some(defined) = styles.styles.get("a")
        .or_else(|| styles.styles.get("Normal"))
    {
        let ds = &defined.paragraph;
        if style.space_before.is_none() { style.space_before = ds.space_before; }
        if style.space_after.is_none() { style.space_after = ds.space_after; }
        if style.line_spacing.is_none() {
            style.line_spacing = ds.line_spacing;
            style.line_spacing_rule = ds.line_spacing_rule.clone();
        }
        if let Some(ref drs) = ds.default_run_style {
            style.default_run_style = Some(drs.clone());
        }
        if ds.keep_next { style.keep_next = true; }
        if ds.keep_lines { style.keep_lines = true; }
    }
    // Apply docDefaults fallback
    if style.default_run_style.is_none() {
        style.default_run_style = styles.doc_default_run_style.clone();
    }
    if let Some(ref doc_para) = styles.doc_default_para_style {
        if style.space_before.is_none() { style.space_before = doc_para.space_before; }
        if style.space_after.is_none() { style.space_after = doc_para.space_after; }
        if style.line_spacing.is_none() {
            style.line_spacing = doc_para.line_spacing;
            style.line_spacing_rule = doc_para.line_spacing_rule.clone();
            style.line_spacing_from_doc_defaults = true;
        }
        if style.indent_left.is_none() && style.indent_left_chars.is_none() {
            style.indent_left = doc_para.indent_left;
            style.indent_left_chars = doc_para.indent_left_chars;
        }
        if style.indent_right.is_none() && style.indent_right_chars.is_none() {
            style.indent_right = doc_para.indent_right;
            style.indent_right_chars = doc_para.indent_right_chars;
        }
        if style.indent_first_line.is_none() && style.indent_first_line_chars.is_none() {
            style.indent_first_line = doc_para.indent_first_line;
            style.indent_first_line_chars = doc_para.indent_first_line_chars;
        }
        // Empty paragraphs: only override if docDefaults explicitly sets widowControl
        if doc_para.has_explicit_widow_control {
            style.widow_control = doc_para.widow_control;
        }
    }
    Paragraph {
        runs: vec![],
        style,
        alignment: Alignment::default(),
        shapes: vec![],
        ppr_change: None,
        paragraph_mark_revision: None,
    }
}

fn parse_header_footer_xml(xml: &str, ctx: &ParseContext, styles: &StyleSheet) -> Result<Vec<Block>, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut blocks = Vec::new();
    let mut depth = 0;
    let mut in_root = false;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "hdr" | "ftr" => {
                        in_root = true;
                        depth = 0;
                    }
                    "p" if in_root && depth == 0 => {
                        let pr = parse_paragraph(&mut reader, ctx, styles, false)?;
                        // S742 (2026-07-04): keep header/footer inline images —
                        // they were silently DROPPED (only pr.paragraph was
                        // pushed), so a header logo contributed no height and
                        // never rendered (probeqhdrimg: Word body top 143.35 =
                        // header_y + image 85 + text line; Oxi started ~85pt
                        // high -> {-1:12}). Mirrors the body path: image-only
                        // host paragraphs are suppressed (S537 twin) and the
                        // images become sibling blocks.
                        let image_only = !pr.inline_images.is_empty()
                            && pr.paragraph.runs.iter().all(|r| r.text.is_empty())
                            && std::env::var("OXI_S742_DISABLE").is_err();
                        if !image_only {
                            blocks.push(Block::Paragraph(pr.paragraph));
                        }
                        if std::env::var("OXI_S742_DISABLE").is_err() {
                            blocks.extend(pr.inline_images);
                        }
                        // S759 (2026-07-09): keep FLOATING header/footer images
                        // (wp:anchor) — like inline (S742) they were DROPPED
                        // (only pr.inline_images was kept). A logo header
                        // (uk_health_form Ofsted) is a wp:anchor image
                        // positioned relative to the PAGE. Push it as a
                        // Block::Image carrying its FloatingPosition; the header
                        // render draws it at the absolute page position and
                        // excludes it from header height (position.is_some()).
                        if std::env::var("OXI_HDRFLOAT_DISABLE").is_err() {
                            for img in pr.floating_images {
                                blocks.push(Block::Image(img));
                            }
                        }
                    }
                    "tbl" if in_root && depth == 0 => {
                        let table = parse_table(&mut reader, ctx, styles)?;
                        blocks.push(Block::Table(table));
                    }
                    "sdt" if in_root && depth == 0 => {
                        // Structured Document Tag: skip sdtPr, process sdtContent children
                        let mut sdt_depth = 1u32;
                        let mut in_sdt_content = false;
                        loop {
                            match reader.read_event()? {
                                Event::Start(se) => {
                                    let sl = local_name(se.name().as_ref());
                                    if sl == "sdtContent" && sdt_depth == 1 {
                                        in_sdt_content = true;
                                    } else if in_sdt_content && sl == "p" {
                                        let pr = parse_paragraph(&mut reader, ctx, styles, false)?;
                                        blocks.push(Block::Paragraph(pr.paragraph));
                                    } else if in_sdt_content && sl == "tbl" {
                                        let table = parse_table(&mut reader, ctx, styles)?;
                                        blocks.push(Block::Table(table));
                                    } else {
                                        sdt_depth += 1;
                                    }
                                }
                                Event::End(ee) => {
                                    let sl = local_name(ee.name().as_ref());
                                    if sl == "sdtContent" {
                                        in_sdt_content = false;
                                    } else if sl == "sdt" && sdt_depth == 1 {
                                        break;
                                    } else if sdt_depth > 1 {
                                        sdt_depth -= 1;
                                    }
                                }
                                Event::Eof => break,
                                _ => {}
                            }
                        }
                    }
                    _ if in_root => {
                        depth += 1;
                    }
                    _ => {}
                }
            }
            Event::Empty(e) => {
                // S806p (2026-07-12): a SELF-CLOSING <w:p/> (Event::Empty, not
                // Start) was silently DROPPED — uklocalspending footer1.xml ends
                // with <w:p/> (a Normal-style empty line Word reserves ~24.6pt
                // for in the footer stack); its loss under-reserved the footer
                // and let the body pack into Word's footer zone. Corpus scan:
                // only ukhealthform/header1 + uklocalspending/footer1 carry the
                // pattern (JP untouched by construction). Push an empty
                // default-style paragraph (same as parse_paragraph on an empty
                // <w:p> with no pPr).
                let local = local_name(e.name().as_ref());
                if local == "p" && in_root && depth == 0 {
                    blocks.push(Block::Paragraph(empty_para_with_defaults(styles)));
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "hdr" || local == "ftr" {
                    in_root = false;
                } else if in_root && depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(blocks)
}

/// Round 29: Rewrite footnoteReference / endnoteReference run text from
/// the OOXML id ("[2]") to the section-local sequential number ("[1]"),
/// using the supplied id→seq maps. Walks paragraphs and recursively into
/// table cells. Run.footnote_ref / Run.endnote_ref still hold the original
/// OOXML id so the rendering loop can look up the body by id.
fn renumber_note_refs(
    blocks: &mut Vec<Block>,
    fn_map: &std::collections::HashMap<u32, u32>,
    en_map: &std::collections::HashMap<u32, u32>,
) {
    for block in blocks.iter_mut() {
        match block {
            Block::Paragraph(para) => {
                for run in para.runs.iter_mut() {
                    if let Some(id) = run.footnote_ref {
                        if let Some(seq) = fn_map.get(&id) {
                            run.text = format!("{}", seq);
                        }
                    }
                    if let Some(id) = run.endnote_ref {
                        if let Some(seq) = en_map.get(&id) {
                            run.text = format!("{}", seq);
                        }
                    }
                }
            }
            Block::Table(table) => {
                for row in table.rows.iter_mut() {
                    for cell in row.cells.iter_mut() {
                        renumber_note_refs(&mut cell.blocks, fn_map, en_map);
                    }
                }
            }
            Block::Image(_) | Block::UnsupportedElement(_) | Block::Math(_) => {}
        }
    }
}

/// Recursively collect footnote/endnote references from blocks
fn collect_note_refs(blocks: &[Block], ctx: &ParseContext, footnotes: &mut Vec<Footnote>, endnotes: &mut Vec<Footnote>) {
    for block in blocks {
        match block {
            Block::Paragraph(para) => {
                for run in &para.runs {
                    if let Some(fn_id) = run.footnote_ref {
                        let id_str = fn_id.to_string();
                        if let Some(note_blocks) = ctx.footnotes.get(&id_str) {
                            if !footnotes.iter().any(|f| f.number == fn_id) {
                                footnotes.push(Footnote {
                                    number: fn_id,
                                    blocks: note_blocks.clone(),
                                });
                            }
                        }
                    }
                    if let Some(en_id) = run.endnote_ref {
                        let id_str = en_id.to_string();
                        if let Some(note_blocks) = ctx.endnotes.get(&id_str) {
                            if !endnotes.iter().any(|f| f.number == en_id) {
                                endnotes.push(Footnote {
                                    number: en_id,
                                    blocks: note_blocks.clone(),
                                });
                            }
                        }
                    }
                }
            }
            Block::Table(table) => {
                for row in &table.rows {
                    for cell in &row.cells {
                        collect_note_refs(&cell.blocks, ctx, footnotes, endnotes);
                    }
                }
            }
            Block::Image(_) | Block::UnsupportedElement(_) | Block::Math(_) => {}
        }
    }
}

/// Parse runs inside w:ins or w:del (tracked changes)
fn parse_tracked_change_runs(
    reader: &mut Reader<&[u8]>,
    ctx: &ParseContext,
    styles: &StyleSheet,
    end_tag: &str,
    tc: TrackedChange,
) -> Result<Vec<Run>, ParseError> {
    let mut runs = Vec::new();
    let mut depth = 0;

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "r" && depth == 0 {
                    let (mut run, _dr) = parse_run(reader, ctx, styles, None)?;
                    // R62 (2026-04-29): parser stores tracked_change ONLY.
                    // Visual styling (underline/strikethrough + author-palette
                    // color) is applied at layout time by R-01
                    // apply_revision_styling_to_run (mod.rs:732), using the
                    // author palette and ShowRevisions mode. Pre-applying
                    // hardcoded FF0000 here was legacy from before R-01
                    // landed and required compensating strip helpers.
                    run.tracked_change = Some(tc.clone());
                    runs.push(run);
                } else {
                    depth += 1;
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == end_tag && depth == 0 {
                    break;
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(runs)
}

/// Parse w:ruby element (furigana). Populates all rubyPr children:
/// rubyAlign, hps, hpsRaise, hpsBaseText, lid (per ECMA-376 §17.3.3.25).
/// Geometry rules in spec/word_layout_spec_ra.md §18.
fn parse_ruby(reader: &mut Reader<&[u8]>) -> Result<Ruby, ParseError> {
    let mut base_text = String::new();
    let mut ruby_text = String::new();
    let mut ruby_font_size: Option<f32> = None;
    let mut align: Option<RubyAlign> = None;
    let mut hps_halfpt: Option<u32> = None;
    let mut hps_raise_halfpt: Option<u32> = None;
    let mut hps_base_text_halfpt: Option<u32> = None;
    let mut lang: Option<String> = None;
    let mut depth = 0;
    let mut in_rt = false;
    let mut in_ruby_base = false;
    let mut in_ruby_pr = false;
    let mut in_t = false;

    fn read_val_attr(e: &quick_xml::events::BytesStart) -> Option<String> {
        for attr in e.attributes().flatten() {
            if local_name(attr.key.as_ref()) == "val" {
                return Some(String::from_utf8_lossy(&attr.value).into_owned());
            }
        }
        None
    }

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "rt" if depth == 0 => { in_rt = true; }
                    "rubyBase" if depth == 0 => { in_ruby_base = true; }
                    "rubyPr" if depth == 0 => { in_ruby_pr = true; }
                    "t" => { in_t = true; }
                    _ => {}
                }
                depth += 1;
            }
            Event::Text(e) => {
                if in_t {
                    let content = e.unescape().unwrap_or_default();
                    if in_rt {
                        ruby_text.push_str(&content);
                    } else if in_ruby_base {
                        base_text.push_str(&content);
                    }
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                if !in_ruby_pr {
                    continue;
                }
                match local.as_str() {
                    "sz" => {
                        if let Some(val) = read_val_attr(&e) {
                            ruby_font_size = val.parse::<f32>().ok().map(|v| v / 2.0);
                        }
                    }
                    "rubyAlign" => {
                        if let Some(val) = read_val_attr(&e) {
                            align = match val.as_str() {
                                "center" => Some(RubyAlign::Center),
                                "distributeLetter" => Some(RubyAlign::DistributeLetter),
                                "distributeSpace" => Some(RubyAlign::DistributeSpace),
                                "left" => Some(RubyAlign::Left),
                                "right" => Some(RubyAlign::Right),
                                "rightVertical" => Some(RubyAlign::RightVertical),
                                _ => None,
                            };
                        }
                    }
                    "hps" => {
                        if let Some(val) = read_val_attr(&e) {
                            hps_halfpt = val.parse::<u32>().ok();
                        }
                    }
                    "hpsRaise" => {
                        if let Some(val) = read_val_attr(&e) {
                            hps_raise_halfpt = val.parse::<u32>().ok();
                        }
                    }
                    "hpsBaseText" => {
                        if let Some(val) = read_val_attr(&e) {
                            hps_base_text_halfpt = val.parse::<u32>().ok();
                        }
                    }
                    "lid" => {
                        lang = read_val_attr(&e);
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    // parse_ruby is invoked AFTER the parent loop consumed
                    // <w:ruby> Start, so we are positioned INSIDE ruby with
                    // depth=0 representing the ruby's own level. We break on
                    // </w:ruby> at that level. (Previous condition used
                    // depth==1 which never matched, causing the reader to
                    // consume past ruby and eat following Runs — the bug
                    // was dormant pre-Round-5 because no baseline doc had
                    // a ruby followed by additional Runs.)
                    "ruby" if depth == 0 => {
                        break;
                    }
                    "rt" => { in_rt = false; }
                    "rubyBase" => { in_ruby_base = false; }
                    "rubyPr" => { in_ruby_pr = false; }
                    "t" => { in_t = false; }
                    _ => {}
                }
                if depth > 0 { depth -= 1; }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    // If hps was parsed and font_size wasn't set from <sz>, derive from hps.
    if ruby_font_size.is_none() {
        ruby_font_size = hps_halfpt.map(|h| h as f32 / 2.0);
    }

    Ok(Ruby {
        base: base_text,
        text: ruby_text,
        font_size: ruby_font_size,
        align,
        hps_halfpt,
        hps_raise_halfpt,
        hps_base_text_halfpt,
        lang,
    })
}

/// Parse word/comments.xml
fn parse_comments_xml(xml: &str) -> Result<HashMap<String, Comment>, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut comments: HashMap<String, Comment> = HashMap::new();
    let mut depth = 0;
    let mut in_comment = false;
    let mut current_id = String::new();
    let mut current_author: Option<String> = None;
    let mut current_date: Option<String> = None;
    let mut current_initials: Option<String> = None;
    let mut current_para_id: Option<String> = None;
    let mut current_blocks: Vec<Block> = Vec::new();

    let note_ctx = ParseContext {
        _rels: HashMap::new(),
        media: HashMap::new(),
        media_types: HashMap::new(),
        hyperlinks: HashMap::new(),
        numbering: super::numbering::NumberingDefinitions::default(),
        list_counters: std::cell::RefCell::new(HashMap::new()),
        footnotes: HashMap::new(),
        endnotes: HashMap::new(),
        comments: HashMap::new(),
        theme: ThemeColors::default(),
    };
    let empty_styles = StyleSheet::default();

    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    "comment" if !in_comment => {
                        in_comment = true;
                        depth = 0;
                        current_blocks.clear();
                        current_id.clear();
                        current_author = None;
                        current_date = None;
                        current_initials = None;
                        current_para_id = None;
                        for attr in e.attributes().flatten() {
                            let key = local_name(attr.key.as_ref());
                            let val = String::from_utf8_lossy(&attr.value).to_string();
                            match key.as_str() {
                                "id" => current_id = val,
                                "author" => current_author = Some(val),
                                "date" => current_date = Some(val),
                                "initials" => current_initials = Some(val),
                                _ => {}
                            }
                        }
                    }
                    "p" if in_comment && depth == 0 => {
                        // The first paragraph's w14:paraId is the join key used
                        // by commentsExtended.xml (MS-DOCX w15).
                        if current_para_id.is_none() {
                            for attr in e.attributes().flatten() {
                                if local_name(attr.key.as_ref()) == "paraId" {
                                    current_para_id =
                                        Some(String::from_utf8_lossy(&attr.value).to_string());
                                    break;
                                }
                            }
                        }
                        let pr = parse_paragraph(&mut reader, &note_ctx, &empty_styles, false)?;
                        let para = pr.paragraph;
                        current_blocks.push(Block::Paragraph(para));
                    }
                    _ if in_comment => {
                        depth += 1;
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                let local = local_name(e.name().as_ref());
                if local == "comment" && in_comment && depth == 0 {
                    if !current_id.is_empty() {
                        comments.insert(current_id.clone(), Comment {
                            id: current_id.clone(),
                            author: current_author.take(),
                            date: current_date.take(),
                            initials: current_initials.take(),
                            para_id: current_para_id.take(),
                            parent_para_id: None,
                            resolved: false,
                            durable_id: None,
                            blocks: std::mem::take(&mut current_blocks),
                        });
                    }
                    in_comment = false;
                } else if in_comment && depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(comments)
}

/// Build the document-level author palette in deterministic first-seen order.
///
/// Sources, in priority:
/// 1. `Document.people` (people.xml — Word writes reviewer-first-seen order)
/// 2. Each `Comment.author`
/// 3. Each `Run.tracked_change.author` (walked in document order, plus
///    rpr_change.author, ppr_change.author, paragraph_mark_revision.author)
///
/// Authors that already appear earlier in the list are skipped, so a single
/// reviewer always gets the same `color_index` regardless of how many times
/// they appear.
fn build_author_palette(
    people: &[Person],
    comments: &[Comment],
    pages: &[Page],
) -> Vec<Author> {
    fn push_unique(seen: &mut Vec<String>, name: &str) {
        if !name.is_empty() && !seen.iter().any(|s| s == name) {
            seen.push(name.to_string());
        }
    }
    let mut seen: Vec<String> = Vec::new();
    for p in people {
        push_unique(&mut seen, &p.author);
    }
    for c in comments {
        if let Some(a) = c.author.as_deref() {
            push_unique(&mut seen, a);
        }
    }
    for page in pages {
        for block in &page.blocks {
            walk_block_authors(block, &mut seen);
        }
    }
    seen.into_iter()
        .enumerate()
        .map(|(color_index, display)| Author { display, color_index })
        .collect()
}

fn walk_block_authors(block: &Block, seen: &mut Vec<String>) {
    fn push(seen: &mut Vec<String>, a: Option<&str>) {
        if let Some(a) = a {
            if !a.is_empty() && !seen.iter().any(|s| s == a) {
                seen.push(a.to_string());
            }
        }
    }
    match block {
        Block::Paragraph(p) => {
            for r in &p.runs {
                if let Some(tc) = r.tracked_change.as_ref() {
                    push(seen, tc.author.as_deref());
                }
                if let Some(pc) = r.rpr_change.as_ref() {
                    push(seen, pc.author.as_deref());
                }
            }
            if let Some(pc) = p.ppr_change.as_ref() {
                push(seen, pc.author.as_deref());
            }
            if let Some(pmark) = p.paragraph_mark_revision.as_ref() {
                push(seen, pmark.author.as_deref());
            }
        }
        Block::Table(t) => {
            for row in &t.rows {
                for cell in &row.cells {
                    for inner in &cell.blocks {
                        walk_block_authors(inner, seen);
                    }
                }
            }
        }
        Block::Image(_) | Block::UnsupportedElement(_) | Block::Math(_) => {}
    }
}

/// Parse `word/people.xml` (MS-DOCX w15).
///
/// Produces one [`Person`] per `<w15:person>`. Preserves document order so
/// downstream code can assign author colours deterministically without a
/// separate sort (Word writes people.xml in reviewer-first-seen order).
fn parse_people_xml(xml: &str) -> Result<Vec<Person>, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut people: Vec<Person> = Vec::new();
    let mut current: Option<Person> = None;
    loop {
        match reader.read_event()? {
            Event::Start(e) => {
                let local = local_name(e.name().as_ref());
                if local == "person" {
                    let mut author = String::new();
                    for attr in e.attributes().flatten() {
                        if local_name(attr.key.as_ref()) == "author" {
                            author = String::from_utf8_lossy(&attr.value).to_string();
                        }
                    }
                    current = Some(Person { author, provider_id: None, user_id: None });
                }
            }
            Event::Empty(e) => {
                let local = local_name(e.name().as_ref());
                match local.as_str() {
                    // Self-closing <w15:person w15:author="..."/> with no presenceInfo.
                    "person" => {
                        let mut author = String::new();
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == "author" {
                                author = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        if !author.is_empty() {
                            people.push(Person { author, provider_id: None, user_id: None });
                        }
                    }
                    "presenceInfo" => {
                        if let Some(p) = current.as_mut() {
                            for attr in e.attributes().flatten() {
                                let key = local_name(attr.key.as_ref());
                                let val = String::from_utf8_lossy(&attr.value).to_string();
                                match key.as_str() {
                                    "providerId" => p.provider_id = Some(val),
                                    "userId" => p.user_id = Some(val),
                                    _ => {}
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::End(e) => {
                if local_name(e.name().as_ref()) == "person" {
                    if let Some(p) = current.take() {
                        if !p.author.is_empty() {
                            people.push(p);
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(people)
}

/// Parse `word/commentsIds.xml` (Word 2019+ w16cid extension).
///
/// Returns a `paraId → durableId` map. The durable id survives save-as
/// roundtrips and is the stable key to use across sessions (the local
/// `w:id` is renumbered freely on save — see `comments_notes.md` §7).
///
/// Accepts `w16cid:durableId` (canonical) and `w16cid:id` (older draft
/// spelling). Namespace stripped via `local_name`.
fn parse_comments_ids_xml(xml: &str) -> Result<HashMap<String, String>, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut map: HashMap<String, String> = HashMap::new();
    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) => {
                if local_name(e.name().as_ref()) != "commentId" {
                    continue;
                }
                let mut para_id: Option<String> = None;
                let mut durable_id: Option<String> = None;
                for attr in e.attributes().flatten() {
                    let key = local_name(attr.key.as_ref());
                    let val = String::from_utf8_lossy(&attr.value).to_string();
                    match key.as_str() {
                        "paraId" => para_id = Some(val),
                        "durableId" | "id" => durable_id = Some(val),
                        _ => {}
                    }
                }
                if let (Some(p), Some(d)) = (para_id, durable_id) {
                    map.insert(p, d);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(map)
}

/// Metadata merged onto a [`Comment`] from `word/commentsExtended.xml`.
///
/// The join key is the `w14:paraId` of the comment body's first paragraph.
#[derive(Debug, Default, Clone)]
struct CommentExtendedInfo {
    parent_para_id: Option<String>,
    resolved: bool,
}

/// Parse `word/commentsExtended.xml` (MS-DOCX w15 extension).
///
/// Each `<w15:commentEx>` carries the paraId of a comment body paragraph plus
/// optional `w15:paraIdParent` (reply threading) and `w15:done` (resolved
/// state). Accept both `w15:paraIdParent` (canonical) and `w15:parentParaId`
/// (legacy variant) — see `comments_notes.md` §4.
fn parse_comments_extended_xml(
    xml: &str,
) -> Result<HashMap<String, CommentExtendedInfo>, ParseError> {
    let mut reader = Reader::from_str(xml);
    let mut map: HashMap<String, CommentExtendedInfo> = HashMap::new();
    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) => {
                if local_name(e.name().as_ref()) != "commentEx" {
                    continue;
                }
                let mut para_id: Option<String> = None;
                let mut info = CommentExtendedInfo::default();
                for attr in e.attributes().flatten() {
                    let key = local_name(attr.key.as_ref());
                    let val = String::from_utf8_lossy(&attr.value).to_string();
                    match key.as_str() {
                        "paraId" => para_id = Some(val),
                        "paraIdParent" | "parentParaId" => info.parent_para_id = Some(val),
                        "done" => info.resolved = val == "1" || val == "true",
                        _ => {}
                    }
                }
                if let Some(pid) = para_id {
                    map.insert(pid, info);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(map)
}

/// Read and discard XML events up to and including the matching End tag for
/// `tag_name`. Used to drain `*PrChange` revision-history bodies (tblPrChange,
/// trPrChange, tcPrChange, sectPrChange, tblGridChange, numberingChange) so
/// their inner property elements don't silently leak into the current parse
/// state — every such body contains a *prior* copy of the same property
/// element it sits inside (e.g. tcPrChange contains a prior tcPr).
fn drain_element(reader: &mut Reader<&[u8]>, tag_name: &str) -> Result<(), ParseError> {
    let mut depth = 0u32;
    loop {
        match reader.read_event()? {
            Event::Start(_) => depth += 1,
            Event::End(e) => {
                if depth == 0 && local_name(e.name().as_ref()) == tag_name {
                    return Ok(());
                }
                if depth > 0 {
                    depth -= 1;
                }
            }
            Event::Eof => return Ok(()),
            _ => {}
        }
    }
}

/// Extract local name from a potentially namespaced XML tag
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
    use crate::ir::Block;

    #[test]
    fn parse_comments_xml_captures_initials_and_metadata() {
        // Mirrors tools/fixtures/comments_samples/fixture_01_single_comment.docx.
        // Word COM-validated 2026-04-18: Comments.Count=1, Author="Alice Reviewer",
        // Initial="AR", Scope.Text="brown fox".
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:comment w:id="0" w:author="Alice Reviewer" w:date="2026-04-18T10:00:00Z" w:initials="AR">
    <w:p w14:paraId="00000010">
      <w:r><w:t>Is 'brown' needed here?</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>"#;
        let comments = parse_comments_xml(xml).expect("parse");
        assert_eq!(comments.len(), 1);
        let c = comments.get("0").expect("id=0 present");
        assert_eq!(c.id, "0");
        assert_eq!(c.author.as_deref(), Some("Alice Reviewer"));
        assert_eq!(c.date.as_deref(), Some("2026-04-18T10:00:00Z"));
        assert_eq!(c.initials.as_deref(), Some("AR"));
        assert_eq!(c.blocks.len(), 1);
        match &c.blocks[0] {
            Block::Paragraph(p) => {
                let text: String = p.runs.iter().map(|r| r.text.as_str()).collect();
                assert_eq!(text, "Is 'brown' needed here?");
            }
            other => panic!("expected paragraph, got {other:?}"),
        }
    }

    #[test]
    fn parse_pprchange_stores_prior_style_without_merging_into_current() {
        // `<w:pPr>` carries current alignment=left, and a `<w:pPrChange>` with
        // prior alignment=right. Without the explicit drain, the inner `<w:jc
        // val="right"/>` would silently overwrite the current alignment via
        // the outer Empty handler (which doesn't gate on depth).
        let xml = r#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:jc w:val="left"/>
  <w:pPrChange w:id="42" w:author="A" w:date="2026-04-18T10:00:00Z">
    <w:pPr>
      <w:jc w:val="right"/>
    </w:pPr>
  </w:pPrChange>
</w:pPr>"#;
        let mut reader = Reader::from_str(xml);
        // Advance to the outer <w:pPr> Start.
        loop {
            match reader.read_event().expect("read") {
                Event::Start(e) if local_name(e.name().as_ref()) == "pPr" => break,
                Event::Eof => panic!("no <w:pPr>"),
                _ => continue,
            }
        }
        let (_style, alignment, _sid, _npr, _spr, ppr_change, _pmark) =
            parse_paragraph_properties(&mut reader).expect("parse");
        assert_eq!(
            alignment,
            Some(crate::ir::Alignment::Left),
            "current alignment must stay Left; pPrChange body must not leak"
        );
        let pc = ppr_change.expect("ppr_change populated");
        assert_eq!(pc.id.as_deref(), Some("42"));
        assert_eq!(pc.author.as_deref(), Some("A"));
        assert!(pc.prior_paragraph_style.is_some(), "prior style must be captured");
    }

    #[test]
    fn show_revisions_default_is_all_and_round_trips_json() {
        use crate::ir::ShowRevisions;
        // I-04: default = All; serde uses snake_case rename.
        assert_eq!(ShowRevisions::default(), ShowRevisions::All);
        assert_eq!(serde_json::to_string(&ShowRevisions::All).unwrap(), "\"all\"");
        assert_eq!(serde_json::to_string(&ShowRevisions::Simple).unwrap(), "\"simple\"");
        assert_eq!(serde_json::to_string(&ShowRevisions::Original).unwrap(), "\"original\"");
        assert_eq!(serde_json::to_string(&ShowRevisions::Final).unwrap(), "\"final\"");
        let parsed: ShowRevisions = serde_json::from_str("\"original\"").unwrap();
        assert_eq!(parsed, ShowRevisions::Original);
    }

    #[test]
    fn build_author_palette_dedupes_in_first_seen_order() {
        let people = vec![Person {
            author: "Alice".into(),
            provider_id: None,
            user_id: None,
        }];
        // Comments add a new author Bob (after Alice from people).
        let comments = vec![Comment {
            id: "0".into(),
            author: Some("Bob".into()),
            date: None,
            initials: None,
            para_id: None,
            parent_para_id: None,
            resolved: false,
            durable_id: None,
            blocks: vec![],
        }];
        let pages: Vec<Page> = Vec::new();
        let palette = build_author_palette(&people, &comments, &pages);
        assert_eq!(palette.len(), 2);
        assert_eq!(palette[0].display, "Alice");
        assert_eq!(palette[0].color_index, 0);
        assert_eq!(palette[1].display, "Bob");
        assert_eq!(palette[1].color_index, 1);

        // Re-seen author keeps original index — duplicates suppressed.
        let comments_dup = vec![Comment {
            id: "0".into(),
            author: Some("Alice".into()),
            date: None,
            initials: None,
            para_id: None,
            parent_para_id: None,
            resolved: false,
            durable_id: None,
            blocks: vec![],
        }];
        let palette2 = build_author_palette(&people, &comments_dup, &pages);
        assert_eq!(palette2.len(), 1);
        assert_eq!(palette2[0].display, "Alice");
        assert_eq!(palette2[0].color_index, 0);
    }

    #[test]
    fn drain_element_skips_nested_body() {
        let xml = r#"<w:trPrChange xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="1">
  <w:trPr>
    <w:trHeight w:val="500"/>
    <w:cantSplit/>
  </w:trPr>
</w:trPrChange><w:after/>"#;
        let mut reader = Reader::from_str(xml);
        // Advance past the opening tag.
        loop {
            match reader.read_event().expect("read") {
                Event::Start(e) if local_name(e.name().as_ref()) == "trPrChange" => break,
                Event::Eof => panic!("no trPrChange"),
                _ => continue,
            }
        }
        drain_element(&mut reader, "trPrChange").expect("drain");
        // After drain, the next event must be `<w:after/>` Empty — proves we
        // consumed exactly the trPrChange body and its closing tag.
        let next = reader.read_event().expect("read after drain");
        match next {
            Event::Empty(e) => assert_eq!(local_name(e.name().as_ref()), "after"),
            other => panic!("expected <w:after/> Empty, got {other:?}"),
        }
    }

    #[test]
    #[allow(non_snake_case)]
    fn parse_table_grid_ignores_tblGridChange_prior_columns() {
        // Without the drain, the prior <w:gridCol w="999"/> inside tblGridChange
        // would be appended to the columns vector. With the drain, only the
        // current columns are emitted.
        let xml = r#"<w:tblGrid xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:gridCol w:w="2880"/>
  <w:gridCol w:w="2880"/>
  <w:tblGridChange w:id="1">
    <w:tblGrid>
      <w:gridCol w:w="9999"/>
    </w:tblGrid>
  </w:tblGridChange>
</w:tblGrid>"#;
        let mut reader = Reader::from_str(xml);
        // Advance to tblGrid Start.
        loop {
            match reader.read_event().expect("read") {
                Event::Start(e) if local_name(e.name().as_ref()) == "tblGrid" => break,
                Event::Eof => panic!("no tblGrid"),
                _ => continue,
            }
        }
        let columns = parse_table_grid(&mut reader).expect("parse");
        assert_eq!(columns.len(), 2, "prior gridCol must not leak");
        assert_eq!(columns, vec![144.0, 144.0]); // 2880 twips → 144 pt
    }

    #[test]
    #[allow(non_snake_case)]
    fn parse_num_pr_ignores_numberingChange_prior_values() {
        // Without the drain, prior <w:numId val="999"/> inside numberingChange
        // would silently overwrite the current numId.
        let xml = r#"<w:numPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:ilvl w:val="0"/>
  <w:numId w:val="5"/>
  <w:numberingChange w:id="1" w:author="A" w:date="2026-04-18T10:00:00Z">
    <w:numPr>
      <w:numId w:val="999"/>
    </w:numPr>
  </w:numberingChange>
</w:numPr>"#;
        let mut reader = Reader::from_str(xml);
        loop {
            match reader.read_event().expect("read") {
                Event::Start(e) if local_name(e.name().as_ref()) == "numPr" => break,
                Event::Eof => panic!("no numPr"),
                _ => continue,
            }
        }
        let np = parse_num_pr(&mut reader).expect("parse");
        assert_eq!(np.num_id, "5", "prior numId must not leak");
        assert_eq!(np.ilvl, 0);
    }

    #[test]
    fn parse_pmark_ins_via_ppr_rpr() {
        // `<w:pPr>/<w:rPr>/<w:ins>` flags the paragraph-mark (¶) as inserted.
        let xml = r#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:rPr>
    <w:ins w:id="77" w:author="Alice" w:date="2026-04-18T10:00:00Z"/>
  </w:rPr>
</w:pPr>"#;
        let mut reader = Reader::from_str(xml);
        loop {
            match reader.read_event().expect("read") {
                Event::Start(e) if local_name(e.name().as_ref()) == "pPr" => break,
                Event::Eof => panic!("no pPr"),
                _ => continue,
            }
        }
        let (_s, _a, _sid, _npr, _spr, _pc, pmark) =
            parse_paragraph_properties(&mut reader).expect("parse");
        let tc = pmark.expect("paragraph_mark_revision populated");
        assert_eq!(tc.change_type, "insert");
        assert_eq!(tc.pair_id.as_deref(), Some("77"));
        assert_eq!(tc.author.as_deref(), Some("Alice"));
    }

    #[test]
    fn parse_pmark_del_via_ppr_rpr() {
        // `<w:pPr>/<w:rPr>/<w:del>` flags the pilcrow as deleted — the
        // paragraph has been merged with the next one.
        let xml = r#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:rPr>
    <w:del w:id="88" w:author="Bob" w:date="2026-04-18T10:00:00Z"/>
  </w:rPr>
</w:pPr>"#;
        let mut reader = Reader::from_str(xml);
        loop {
            match reader.read_event().expect("read") {
                Event::Start(e) if local_name(e.name().as_ref()) == "pPr" => break,
                Event::Eof => panic!("no pPr"),
                _ => continue,
            }
        }
        let (_s, _a, _sid, _npr, _spr, _pc, pmark) =
            parse_paragraph_properties(&mut reader).expect("parse");
        let tc = pmark.expect("paragraph_mark_revision populated");
        assert_eq!(tc.change_type, "delete");
        assert_eq!(tc.pair_id.as_deref(), Some("88"));
        assert_eq!(tc.author.as_deref(), Some("Bob"));
    }

    #[test]
    fn parse_run_captures_comment_reference() {
        // ECMA-376: <w:commentReference w:id="N"/> is a zero-width marker inside
        // a <w:r>. The enclosing run is what the renderer projects to the right
        // margin — the id must survive to Run::comment_references.
        let xml = r#"<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
  <w:commentReference w:id="0"/>
</w:r>"#;
        let ctx = ParseContext {
            _rels: HashMap::new(),
            media: HashMap::new(),
            media_types: HashMap::new(),
            hyperlinks: HashMap::new(),
            numbering: NumberingDefinitions::default(),
            list_counters: std::cell::RefCell::new(HashMap::new()),
            footnotes: HashMap::new(),
            endnotes: HashMap::new(),
            comments: HashMap::new(),
            theme: ThemeColors::default(),
        };
        let styles = StyleSheet::default();
        let mut reader = Reader::from_str(xml);
        // Advance to the <w:r> start tag the way parse_paragraph does.
        let start = loop {
            match reader.read_event().expect("read") {
                Event::Start(e) if local_name(e.name().as_ref()) == "r" => break e,
                Event::Eof => panic!("no <w:r> in test fixture"),
                _ => continue,
            }
        };
        let _ = start;
        let (run, _dr) = parse_run(&mut reader, &ctx, &styles, None).expect("parse_run");
        assert_eq!(run.comment_references, vec!["0".to_string()]);
    }

    #[test]
    fn parse_comments_xml_captures_first_para_id() {
        // w14:paraId on the comment body's first <w:p> is the join key used by
        // commentsExtended.xml. Must land on Comment.para_id.
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:comment w:id="0" w:author="A" w:date="2026-04-18T10:00:00Z" w:initials="A">
    <w:p w14:paraId="00000010"><w:r><w:t>body</w:t></w:r></w:p>
  </w:comment>
</w:comments>"#;
        let comments = parse_comments_xml(xml).expect("parse");
        assert_eq!(comments.get("0").and_then(|c| c.para_id.as_deref()), Some("00000010"));
    }

    #[test]
    fn parse_comments_extended_reply_and_resolved() {
        // Mirrors fixture_02/03: paraIdParent marks a reply, done="1" marks resolved.
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w15:commentEx w15:paraId="00000010" w15:done="0"/>
  <w15:commentEx w15:paraId="00000011" w15:paraIdParent="00000010" w15:done="0"/>
  <w15:commentEx w15:paraId="00000020" w15:done="1"/>
</w15:commentsEx>"#;
        let ext = parse_comments_extended_xml(xml).expect("parse");
        assert_eq!(ext.len(), 3);
        let root = ext.get("00000010").expect("root");
        assert!(root.parent_para_id.is_none());
        assert!(!root.resolved);
        let reply = ext.get("00000011").expect("reply");
        assert_eq!(reply.parent_para_id.as_deref(), Some("00000010"));
        assert!(!reply.resolved);
        let resolved = ext.get("00000020").expect("resolved");
        assert!(resolved.resolved);
    }

    #[test]
    fn parse_people_xml_two_reviewers_preserves_order() {
        // Mirrors fixture_10_multiple_reviewers.docx
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w15:person w15:author="Alice Reviewer">
    <w15:presenceInfo w15:providerId="None" w15:userId="Alice Reviewer"/>
  </w15:person>
  <w15:person w15:author="Bob Reviewer">
    <w15:presenceInfo w15:providerId="None" w15:userId="Bob Reviewer"/>
  </w15:person>
</w15:people>"#;
        let people = parse_people_xml(xml).expect("parse");
        assert_eq!(people.len(), 2);
        assert_eq!(people[0].author, "Alice Reviewer");
        assert_eq!(people[0].provider_id.as_deref(), Some("None"));
        assert_eq!(people[0].user_id.as_deref(), Some("Alice Reviewer"));
        assert_eq!(people[1].author, "Bob Reviewer");
        assert_eq!(people[1].user_id.as_deref(), Some("Bob Reviewer"));
    }

    #[test]
    fn parse_people_xml_without_presence_info() {
        // <w15:person> without a nested <w15:presenceInfo> — provider_id and
        // user_id must be None, not empty strings.
        let xml = r#"<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w15:person w15:author="Charlie"/>
</w15:people>"#;
        let people = parse_people_xml(xml).expect("parse");
        assert_eq!(people.len(), 1);
        assert_eq!(people[0].author, "Charlie");
        assert!(people[0].provider_id.is_none());
        assert!(people[0].user_id.is_none());
    }

    #[test]
    fn parse_people_xml_drops_blank_author() {
        // Malformed: <w15:person> without w15:author — must not appear in output.
        let xml = r#"<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w15:person><w15:presenceInfo w15:providerId="None" w15:userId="x"/></w15:person>
  <w15:person w15:author="OK"><w15:presenceInfo w15:providerId="None" w15:userId="OK"/></w15:person>
</w15:people>"#;
        let people = parse_people_xml(xml).expect("parse");
        assert_eq!(people.len(), 1);
        assert_eq!(people[0].author, "OK");
    }

    #[test]
    fn parse_comments_ids_durable_id_mapping() {
        let xml = r#"<?xml version="1.0"?>
<w16cid:commentsIds xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">
  <w16cid:commentId w16cid:paraId="00000010" w16cid:durableId="12345"/>
  <w16cid:commentId w16cid:paraId="00000011" w16cid:durableId="67890"/>
</w16cid:commentsIds>"#;
        let map = parse_comments_ids_xml(xml).expect("parse");
        assert_eq!(map.get("00000010").map(String::as_str), Some("12345"));
        assert_eq!(map.get("00000011").map(String::as_str), Some("67890"));
    }

    #[test]
    fn parse_comments_ids_accepts_legacy_id_attribute() {
        // Older drafts used `w16cid:id` before settling on `w16cid:durableId`.
        let xml = r#"<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">
  <w16cid:commentId w16cid:paraId="00000020" w16cid:id="AAA"/>
</w16cid:commentsIds>"#;
        let map = parse_comments_ids_xml(xml).expect("parse");
        assert_eq!(map.get("00000020").map(String::as_str), Some("AAA"));
    }

    #[test]
    fn parse_comments_extended_accepts_legacy_parent_para_id_spelling() {
        // comments_notes.md §4 — some Word versions emit w15:parentParaId.
        let xml = r#"<?xml version="1.0"?>
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w15:commentEx w15:paraId="00000002" w15:parentParaId="00000001" w15:done="0"/>
</w15:commentsEx>"#;
        let ext = parse_comments_extended_xml(xml).expect("parse");
        assert_eq!(
            ext.get("00000002").and_then(|i| i.parent_para_id.as_deref()),
            Some("00000001")
        );
    }

    #[test]
    fn parse_comments_xml_missing_initials_is_none() {
        // Older Word versions sometimes omit w:initials; it must parse as None,
        // not empty string.
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="5" w:author="Bob" w:date="2026-04-18T10:00:00Z">
    <w:p><w:r><w:t>No initials set</w:t></w:r></w:p>
  </w:comment>
</w:comments>"#;
        let comments = parse_comments_xml(xml).expect("parse");
        let c = comments.get("5").expect("id=5 present");
        assert!(c.initials.is_none(), "initials should be None when absent");
        assert_eq!(c.author.as_deref(), Some("Bob"));
    }

    #[test]
    fn parse_ruby_captures_full_ruby_pr() {
        // ECMA-376 §17.3.3.25 ruby with all rubyPr children. Geometry rules
        // verified empirically in spec/word_layout_spec_ra.md §18.
        let xml = r#"<w:ruby xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:rubyPr>
    <w:rubyAlign w:val="distributeLetter"/>
    <w:hps w:val="11"/>
    <w:hpsRaise w:val="18"/>
    <w:hpsBaseText w:val="21"/>
    <w:lid w:val="ja-JP"/>
  </w:rubyPr>
  <w:rt><w:r><w:rPr><w:sz w:val="11"/></w:rPr><w:t>かんじ</w:t></w:r></w:rt>
  <w:rubyBase><w:r><w:t>漢字</w:t></w:r></w:rubyBase>
</w:ruby>"#;
        let mut reader = Reader::from_str(xml);
        // Advance to the <w:ruby> start tag.
        loop {
            match reader.read_event().expect("read") {
                Event::Start(e) if local_name(e.name().as_ref()) == "ruby" => break,
                Event::Eof => panic!("no <w:ruby> in fixture"),
                _ => continue,
            }
        }
        let ruby = parse_ruby(&mut reader).expect("parse_ruby");
        assert_eq!(ruby.base, "漢字");
        assert_eq!(ruby.text, "かんじ");
        assert_eq!(ruby.align, Some(RubyAlign::DistributeLetter));
        assert_eq!(ruby.hps_halfpt, Some(11));
        assert_eq!(ruby.hps_raise_halfpt, Some(18));
        assert_eq!(ruby.hps_base_text_halfpt, Some(21));
        assert_eq!(ruby.lang.as_deref(), Some("ja-JP"));
        // hps=11 half-pt → font_size 5.5pt (derived from hps when no <sz> present)
        assert_eq!(ruby.font_size, Some(5.5));
    }

    #[test]
    fn parse_ruby_minimal_omits_optional_fields() {
        // ECMA-376 allows w:ruby with no rubyPr at all; all rubyPr-derived
        // fields must default to None and not panic.
        let xml = r#"<w:ruby xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:rt><w:r><w:t>ふく</w:t></w:r></w:rt>
  <w:rubyBase><w:r><w:t>含</w:t></w:r></w:rubyBase>
</w:ruby>"#;
        let mut reader = Reader::from_str(xml);
        loop {
            match reader.read_event().expect("read") {
                Event::Start(e) if local_name(e.name().as_ref()) == "ruby" => break,
                Event::Eof => panic!("no <w:ruby> in fixture"),
                _ => continue,
            }
        }
        let ruby = parse_ruby(&mut reader).expect("parse_ruby");
        assert_eq!(ruby.base, "含");
        assert_eq!(ruby.text, "ふく");
        assert!(ruby.align.is_none());
        assert!(ruby.hps_halfpt.is_none());
        assert!(ruby.hps_raise_halfpt.is_none());
        assert!(ruby.hps_base_text_halfpt.is_none());
        assert!(ruby.lang.is_none());
    }

    #[test]
    fn parse_ruby_align_modes_round_trip() {
        // Each w:rubyAlign value parses to its corresponding RubyAlign variant.
        let modes = [
            ("center", RubyAlign::Center),
            ("distributeLetter", RubyAlign::DistributeLetter),
            ("distributeSpace", RubyAlign::DistributeSpace),
            ("left", RubyAlign::Left),
            ("right", RubyAlign::Right),
            ("rightVertical", RubyAlign::RightVertical),
        ];
        for (val, expected) in modes {
            let xml = format!(
                r#"<w:ruby xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:rubyPr><w:rubyAlign w:val="{}"/></w:rubyPr>
  <w:rt><w:r><w:t>a</w:t></w:r></w:rt>
  <w:rubyBase><w:r><w:t>b</w:t></w:r></w:rubyBase>
</w:ruby>"#,
                val
            );
            let mut reader = Reader::from_str(&xml);
            loop {
                match reader.read_event().expect("read") {
                    Event::Start(e) if local_name(e.name().as_ref()) == "ruby" => break,
                    Event::Eof => panic!("no <w:ruby>"),
                    _ => continue,
                }
            }
            let ruby = parse_ruby(&mut reader).expect("parse_ruby");
            assert_eq!(ruby.align, Some(expected), "align for w:val={val:?}");
        }
    }
}
