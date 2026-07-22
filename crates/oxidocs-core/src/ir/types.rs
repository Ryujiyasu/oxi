// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use serde::{Deserialize, Serialize};
use std::collections::HashMap;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Document {
    pub pages: Vec<Page>,
    pub styles: StyleSheet,
    pub metadata: DocumentMetadata,
    /// Comments referenced in the document
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub comments: Vec<Comment>,
    /// Authors known to the document (from `word/people.xml`, MS-DOCX w15).
    /// Seeds the renderer's author-color palette (attack-matrix row R-02).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub people: Vec<Person>,
    /// Author palette derived from people.xml + Comment.author + Run.tracked_change.author.
    /// Order is first-seen across the document, with `people.xml` honoured first
    /// (Word writes people.xml in reviewer-first-seen order). Each entry's
    /// `color_index` is its position in this Vec, so the renderer can map
    /// `color_index` → RGB through any palette without a separate join step.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub authors: Vec<Author>,
    /// Compatibility: adjustLineHeightInTable (compat65).
    /// true = adjust line height in table cells (disable grid snap in cells).
    /// false (default) = table cells snap to document grid like normal paragraphs.
    #[serde(default)]
    pub adjust_line_height_in_table: bool,
    /// Default tab stop interval from w:settings/w:defaultTabStop (in points).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_tab_stop: Option<f32>,
    /// Compatibility mode (from w:settings/w:compat/w:compatSetting w:name="compatibilityMode")
    /// 14=Word 2010, 15=Word 2013+. Affects table cell grid snap behavior.
    #[serde(default)]
    pub compat_mode: u32,
    /// S545 (2026-06-11): whether a compatibilityMode compatSetting was
    /// actually present. ABSENT means a legacy (Word ≤2010) document, which
    /// Word lays out with ≤14 behavior (e.g. jc=left demand oikomi), but
    /// parse_compat_mode reports 15 for backward compatibility with the
    /// shipped >=15 gates (S475/S476/...). Consumers that need Word's
    /// legacy-vs-2013 split must check `compat_mode <= 14 || !compat_mode_explicit`.
    #[serde(default)]
    pub compat_mode_explicit: bool,
    /// S933b (2026-07-18): whether word/settings.xml EXISTS as a part. The
    /// justify-shrink allowance classes measured so far differ on it:
    /// settings present + compatibilityMode ABSENT -> ~0 (legal__000ad039
    /// in-doc: Word wraps 'Plan'/'The' at needs ~<2/4.2pt); settings part
    /// MISSING entirely -> flat fs/4 (the S825 booster-hunt cs3 probes).
    #[serde(default)]
    pub settings_part_exists: bool,
    /// S833 (2026-07-13): settings.xml `<w:footnotePr>` declares CUSTOM special
    /// footnotes (`<w:footnote w:id="-1"/>` etc.). Word then reserves/renders
    /// the custom separator paragraph at its FULL styled height (style-chain
    /// space before/after included) and pre-reserves the continuationNotice
    /// paragraph even on non-continuing pages; without the declaration the
    /// built-in compact separator (~one bare line) applies. Derived via the
    /// _pb_fnres probe family on uklocalspending (P1-P4 decomposition).
    #[serde(default)]
    pub fn_special_declared: bool,
    /// w:characterSpacingControl from settings.xml.
    /// True when value is "compressPunctuation" or "compressPunctuationAndJapaneseKana"
    /// (enables CJK yakumono compression). False (default) for "doNotCompress" or absent.
    #[serde(default)]
    pub compress_punctuation: bool,
    /// w:doNotExpandShiftReturn compat setting.
    /// When true, Shift+Enter (soft break) lines are NOT justified even in jc=both paragraphs.
    #[serde(default)]
    pub do_not_expand_shift_return: bool,
    /// w:balanceSingleByteDoubleByteWidth compat setting.
    /// When true, the effective character_spacing for CJK fullwidth chars is doubled
    /// (Word's "balance single/double byte widths" mode).
    /// Derived from V19 minimal repro vs real 1636 (Session 56 Finding 3).
    #[serde(default)]
    pub balance_single_byte_double_byte_width: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Page {
    pub blocks: Vec<Block>,
    pub size: PageSize,
    pub margin: Margin,
    /// Document grid line pitch in points (from w:docGrid w:linePitch).
    /// When set with grid_type "lines" or "linesAndChars", line spacing
    /// snaps to multiples of this pitch.
    #[serde(default)]
    pub grid_line_pitch: Option<f32>,
    /// Character grid pitch in points (from w:docGrid w:charSpace for linesAndChars).
    /// When set, character widths are expanded to align to this grid.
    #[serde(default)]
    pub grid_char_pitch: Option<f32>,
    /// Raw charSpace value from docGrid (stored so post-process can recompute
    /// grid_char_pitch with correct default_font_size). Unit: 1/4096 of a point.
    #[serde(default)]
    pub grid_char_space_raw: Option<i32>,
    /// Character-width ratio = grid_char_pitch / default_font_size.
    /// Word renders fullwidth CJK chars at `font_size × ratio` in LM2 mode.
    /// COM-confirmed 2026-04-19 on b35 (fs=9, pitch=11.337, default=12): advance=8.5pt.
    #[serde(default)]
    pub grid_char_cw_ratio: Option<f32>,
    /// True when docGrid element exists but has NO type attribute.
    /// CJK 83/64 multiplier is NOT applied; COM-measured Single heights used instead.
    #[serde(default)]
    pub doc_grid_no_type: bool,
    /// True when docGrid type == "linesAndChars" (a CHARACTER grid: Word fixes the
    /// char count per line). Distinguishes it from type=="lines" (line grid only,
    /// width-determined break). S475 yakumono capacity break applies ONLY to the
    /// non-char-grid case; linesAndChars is grid-determined (charGrid mechanism).
    #[serde(default)]
    pub doc_grid_lines_and_chars: bool,
    /// Header content (paragraphs from header part) — the DEFAULT-type header
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub header: Vec<Block>,
    /// Footer content (paragraphs from footer part) — the DEFAULT-type footer
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub footer: Vec<Block>,
    /// S755: first-page header/footer (type="first", active when title_pg)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub header_first: Vec<Block>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub footer_first: Vec<Block>,
    /// S755: even-page header/footer (type="even", active when even_odd_hf)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub header_even: Vec<Block>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub footer_even: Vec<Block>,
    /// WATERMARK (2026-07-08): a VML WordArt watermark from a header part
    /// (`v:shape type="#_x0000_t136"` + `v:textpath string=…`, the Word
    /// PowerPlusWaterMarkObject idiom — «SAMPLE»/«DRAFT» diagonal text on
    /// every page, centered on the margin box, painted BEHIND the body).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub watermark: Option<Watermark>,
    /// S755: sectPr <w:titlePg/> — this section's first page uses the "first" type
    #[serde(default)]
    pub title_pg: bool,
    /// S755: settings.xml <w:evenAndOddHeaders/> — even pages use the "even" type
    #[serde(default)]
    pub even_odd_hf: bool,
    /// Header distance from page top edge in points (w:pgMar header attr)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub header_distance: Option<f32>,
    /// Footer distance from page bottom edge in points (w:pgMar footer attr)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footer_distance: Option<f32>,
    /// Footnotes referenced in this page
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub footnotes: Vec<Footnote>,
    /// Endnotes referenced in this page
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub endnotes: Vec<Footnote>,
    /// Floating images (anchored, not inline)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub floating_images: Vec<Image>,
    /// Text boxes
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub text_boxes: Vec<TextBox>,
    /// Geometric shapes (DrawingML / VML)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub shapes: Vec<Shape>,
    /// Column layout for this section
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub columns: Option<ColumnLayout>,
    /// S560 (2026-06-13): per-section column layouts within a merged
    /// continuous-section page. When `continuous` section breaks merge
    /// multiple sections into one Page (see parser ooxml.rs), each section
    /// may declare a DIFFERENT column count (e.g. kyotei36spec: a 1-col
    /// form table followed by a continuous 2-col 記載心得 instruction
    /// block). The old single `columns` field could only hold ONE layout
    /// (the last section's), so the 1-col content was wrongly laid out in
    /// the 2-col context. `column_runs` records (block_start_index,
    /// section_columns) for each section span so layout can switch the
    /// column geometry at each boundary. Empty = use `columns` for the
    /// whole page (single-section / non-parser constructions).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub column_runs: Vec<(usize, Option<ColumnLayout>)>,
    /// S729 (2026-07-03): per-section HORIZONTAL margins for merged
    /// continuous sections — (block_start_index, margin_left, margin_right).
    /// The S560 merge kept only the first section's margins, so a continuous
    /// section with different left/right margins rendered at the wrong text
    /// width (probexmargins {-1:6}; the documented kyotei S560 residual).
    /// Parallel to `column_runs` (the same parser sites populate both).
    /// Empty = uniform margins (single-section / non-parser constructions).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub margin_runs: Vec<(usize, f32, f32)>,
    /// S863: per-section vertical page geometry for merged continuous
    /// sections: (block_start, top, bottom, header_distance, footer_distance).
    /// A continuous section adopts this geometry on the next physical page;
    /// its blocks continue at the current cursor on the boundary page.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub vertical_runs: Vec<(usize, f32, f32, Option<f32>, Option<f32>)>,
    /// S735 (2026-07-03): per-section docGrid line pitch for merged continuous
    /// sections — (block_start_index, grid_line_pitch). The merged Page kept
    /// only the first section's pitch (probezcontgrid: pitch 360→480 change at
    /// a continuous break was ignored → later section packed at the old pitch,
    /// {-1:10}). Parallel to `margin_runs`.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub grid_runs: Vec<(usize, Option<f32>)>,
    /// S732 (2026-07-03): how this section STARTS relative to the previous
    /// one — "evenPage"/"oddPage" force the section onto the next even/odd
    /// physical page, inserting a BLANK page when the parity mismatches
    /// (probexeo2: Word leaves pages 3 and 6 blank). None/nextPage = normal.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub section_start_type: Option<String>,
    /// Page number format (e.g. "decimal", "lowerRoman", "upperRoman", "lowerLetter", "upperLetter")
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_number_format: Option<String>,
    /// Starting page number for this section
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_number_start: Option<u32>,
    /// Page borders
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub page_borders: Option<PageBorders>,
    /// S394 (2026-05-27): total count of <w:lastRenderedPageBreak/>
    /// occurrences across all runs (body + nested tables/cells) in this
    /// section. Used to discriminate "low-LRPB docs" (clean Word-current
    /// break hints, e.g. b837=6, d77a=11) from "high-LRPB docs" (stale
    /// re-render artifacts, e.g. 3a4f=82). Gates per-LINE LRPB respect
    /// (S391/OXI_S391_PER_LINE_LRPB) when env OXI_S394_LRPB_MAX is set.
    #[serde(default)]
    pub total_lrpb_count: usize,
    /// True when this section is bidirectional (`<w:bidi/>` in sectPr).
    /// In a bidi (RTL) section, a multi-column layout flows columns
    /// RIGHT-to-LEFT: the first reading column is the RIGHTMOST one.
    /// Word-confirmed (minimal repro + albalunaTaidan): first content
    /// fills the right column, subsequent columns proceed leftward.
    #[serde(default)]
    pub bidi_columns: bool,
    /// True when this section is vertical writing (tategaki/縦書き,
    /// `<w:textDirection w:val="tbRl"/>` in sectPr): characters stack
    /// top-to-bottom within a line, lines advance right-to-left, and
    /// multi-column "columns" become horizontal bands stacked top-to-bottom.
    #[serde(default)]
    pub vertical_section: bool,
}

/// Page border definitions
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PageBorders {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<BorderDef>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
    Image(Image),
    /// OMML math block (inline `<m:oMath>` or display `<m:oMathPara>`).
    /// See `crates/oxidocs-core/src/ir/math.rs` for the recursive
    /// expression tree.
    Math(crate::ir::math::MathBlock),
    /// Placeholder for unsupported content (SmartArt, Chart, etc.)
    UnsupportedElement(UnsupportedElement),
}

/// Represents an unsupported OOXML element that was skipped during parsing
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UnsupportedElement {
    /// Type of unsupported element (e.g. "SmartArt", "Chart", "ActiveX")
    pub element_type: String,
    /// Optional fallback image data (base64)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fallback_image: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Paragraph {
    pub runs: Vec<Run>,
    pub style: ParagraphStyle,
    pub alignment: Alignment,
    /// Inline/anchor shapes attached to this paragraph (e.g. bracketPair)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub shapes: Vec<Shape>,
    /// Paragraph-property change (`<w:pPrChange>`): carries the prior pPr so
    /// the renderer can reconstruct "Original" views (attack-matrix R-13).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ppr_change: Option<PropertyChange>,
    /// Revision on the paragraph-mark itself (`<w:pPr>/<w:rPr>/<w:ins>` or
    /// `<w:pPr>/<w:rPr>/<w:del>`). An inserted mark means the ¶ split is
    /// new; a deleted mark means the ¶ was removed (paragraph merged with
    /// the next). See revisions_notes.md §2.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub paragraph_mark_revision: Option<TrackedChange>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Run {
    pub text: String,
    pub style: RunStyle,
    /// Hyperlink URL (external) or anchor (internal bookmark)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub url: Option<String>,
    /// Footnote reference number (1-based)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub footnote_ref: Option<u32>,
    /// Endnote reference number (1-based)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub endnote_ref: Option<u32>,
    /// Comment IDs that start at this run
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub comment_range_start: Vec<String>,
    /// Comment IDs that end at this run
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub comment_range_end: Vec<String>,
    /// Comment IDs whose balloon anchor (`<w:commentReference>`) sits in this run.
    /// Typically the enclosing run carries rStyle="CommentReference"; the reference
    /// marker is zero-width but the renderer projects the Y of this run to the
    /// right margin to position the balloon (ECMA-376 §22.1.2.56 + §17.13.1 Word
    /// display rules).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub comment_references: Vec<String>,
    /// Tracked change info (insertion/deletion)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tracked_change: Option<TrackedChange>,
    /// Run-property change (`<w:rPrChange>`): carries the prior rPr so the
    /// renderer can emit "formatting changed" annotations (attack-matrix R-12).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rpr_change: Option<PropertyChange>,
    /// Ruby (furigana) annotation
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ruby: Option<Ruby>,
    /// Bookmark anchor name (from w:bookmarkStart w:name)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bookmark_name: Option<String>,
    /// Whether this run contains OMML math content
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub is_math: bool,
    /// Field type for dynamic content substitution (PAGE, NUMPAGES, etc.)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub field_type: Option<FieldType>,
    /// `<w:lastRenderedPageBreak/>` hint (ECMA-376 §17.3.1.18). When set,
    /// Word's last-saved render had a page break before this point. Used by
    /// the SOFT LRPB layout rule: force a page break before this paragraph
    /// ONLY if the paragraph would naturally fit on the current page (i.e.,
    /// Oxi has not already overflowed past Word's saved break position).
    /// Naive (always force) caused cascading regressions in over-packed
    /// docs (Session 56 Day 3, 2026-05-07: 0e7af 1.0→0.26, d77a 0.96→0.27).
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub has_last_rendered_page_break: bool,
}

/// Field types for dynamic content that gets resolved during layout
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum FieldType {
    /// Current page number (PAGE field)
    Page,
    /// Total number of pages (NUMPAGES field)
    NumPages,
    /// Cross-reference (REF/NOTEREF/PAGEREF) — render the CACHED RESULT (the text
    /// between fldChar separate and end), NOT a "#" placeholder. Word displays the
    /// cached value (e.g. «第１９条»); Oxi can't re-resolve the bookmark, so the
    /// cache is the only source. Dropping it shifts wrapping doc-wide (tokyoshugyo).
    CrossRef,
    /// Cached-result field (DATE/TIME/CREATEDATE/SAVEDATE/AUTHOR/TITLE/FILENAME/…).
    /// S708 (2026-06-30): Word displays the CACHED RESULT run (the text between
    /// fldChar separate and end, e.g. «2026/06/30»), which it last evaluated on
    /// save. Oxi can't re-evaluate these, so the cache is the only display source.
    /// Behaves like CrossRef: KEEP the cached result, suppress the instruction.
    /// The old code showed the raw instruction («DATE \@ "yyyy/MM/dd"») or a
    /// «[AUTHOR]» placeholder and DROPPED the cache → garbage text + shifted wrapping.
    Cached,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RunStyle {
    pub font_family: Option<String>,
    /// East Asian font family (w:rFonts eastAsia) for CJK characters
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_family_east_asia: Option<String>,
    /// True iff `<w:rFonts w:eastAsia="..."/>` was set as an EXPLICIT attribute
    /// somewhere in the inheritance chain (run / style / docDefault), as
    /// opposed to a theme-fallback `eastAsiaTheme="minorEastAsia"`. Used by
    /// §4.6.3 Latin-space-adjacent-CJK widening (COM-confirmed jfmb vs
    /// runtime-saved equivalent: only docs with explicit eastAsia widen).
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub has_explicit_east_asia: bool,
    /// East-Asian language tag (`<w:lang w:eastAsia="ja-JP">`). Drives Word's
    /// ambiguous-quote font choice: a CJK eastAsia lang (ja/zh/ko) renders
    /// curly quotes in the eastAsia font; a Latin one (en-US) renders them in
    /// the Latin font. S763c.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub east_asia_lang: Option<String>,
    /// Latin language tag (`<w:lang w:val="en-US">`). Normally Latin; a CJK
    /// value (`w:val="ja"`) makes Word resolve Latin text AND the paragraph
    /// mark through the East Asian font chain instead of the ASCII font. S956.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub latin_lang: Option<String>,
    pub font_size: Option<f32>,
    pub bold: bool,
    /// S976: `w:b` was EXPLICITLY set (the element is present, whatever its
    /// `w:val`) — distinguishes an explicit `<w:b w:val="0"/>`, which must beat
    /// a basedOn parent's / paragraph style's ON, from "not set", which
    /// inherits. The keepNext (S955) three-state pattern; without it a run that
    /// switches a bold heading style off is measured with BOLD glyph widths.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub has_explicit_bold: bool,
    pub italic: bool,
    /// S976: `w:i` three-state marker (see has_explicit_bold).
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub has_explicit_italic: bool,
    pub underline: bool,
    /// Underline style (e.g. "single", "double", "wave", "dash", "dotted", "thick")
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline_style: Option<String>,
    pub strikethrough: bool,
    /// S988A: the strike/dstrike/caps/smallCaps CT_OnOff was set explicitly
    /// (incl. `w:val="false"`), so it wins over an inherited ON — the S976
    /// bold/italic three-state extended to these siblings.
    #[serde(default)]
    pub has_explicit_strikethrough: bool,
    /// Double strikethrough (w:dstrike)
    #[serde(default)]
    pub double_strikethrough: bool,
    #[serde(default)]
    pub has_explicit_double_strikethrough: bool,
    pub color: Option<String>,
    pub highlight: Option<String>,
    pub vertical_align: Option<VerticalAlign>,
    /// Character spacing in points (w:spacing w:val in twips, converted to pt)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub character_spacing: Option<f32>,
    /// Small capitals (w:smallCaps)
    #[serde(default)]
    pub small_caps: bool,
    #[serde(default)]
    pub has_explicit_small_caps: bool,
    /// All capitals (w:caps)
    #[serde(default)]
    pub all_caps: bool,
    #[serde(default)]
    pub has_explicit_all_caps: bool,
    /// Character-level shading/background color (w:shd fill, hex)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<String>,
    /// Right-to-left text run (w:rtl)
    #[serde(default)]
    pub rtl: bool,
    /// Hidden text (w:vanish)
    #[serde(default)]
    pub vanish: bool,
    /// Text outline effect (w:outline)
    #[serde(default)]
    pub outline: bool,
    /// Text shadow effect (w:shadow)
    #[serde(default)]
    pub shadow: bool,
    /// Text emboss effect (w:emboss)
    #[serde(default)]
    pub emboss: bool,
    /// Text imprint/engrave effect (w:imprint)
    #[serde(default)]
    pub imprint: bool,
    /// Complex script font size in points (w:szCs, half-points / 2)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_cs: Option<f32>,
    /// Complex script bold (w:bCs)
    #[serde(default)]
    pub bold_cs: bool,
    /// Complex script italic (w:iCs)
    #[serde(default)]
    pub italic_cs: bool,
    /// Character kerning threshold in points (w:kern, half-points / 2)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub kern: Option<f32>,
    /// Fit text width in points (w:fitText, twips / 20)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fit_text: Option<f32>,
    /// Fit text group ID (w:fitText w:id)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fit_text_id: Option<i64>,
    /// Character width scale percentage (w:w, default 100)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_scale: Option<f32>,
    /// East Asian layout: combine (kumimoji / 割注 — two lines in one cell)
    #[serde(default)]
    pub combine: bool,
    /// w:combineBrackets value (none/round/square/angle/curly) for `combine`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub combine_brackets: Option<String>,
    /// East Asian layout: vertical-in-horizontal (tate-chu-yoko)
    #[serde(default)]
    pub vert_in_horz: bool,
    /// Vertical position offset in points (w:position, half-points / 2)
    /// Positive = raised, negative = lowered
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<f32>,
    /// Emphasis mark / 圏点 (w:em): "dot", "comma", "circle", "underDot"
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub emphasis_mark: Option<String>,
    /// S706 (2026-06-30): run/character border (w:bdr) — Word draws a box
    /// around the run's text (e.g. a bordered title banner). Rendered per
    /// line fragment as a BoxRect stroke.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run_border: Option<BorderDef>,
    /// S839 (2026-07-14): (width, height) pt of an INLINE visual drawing
    /// (wpg vector group without textbox text — hmrc's checkbox strips)
    /// hosted by this run. The run becomes a width-bearing atomic line
    /// fragment (Word reserves cx and positions it via the tab machinery:
    /// the NI strip is CENTERED at its tab stop) and the emit loop draws
    /// the group's vector shapes at the fragment position. None everywhere
    /// except the marked runs (wpg = hmrc/framework only by corpus scan).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inline_object_extent: Option<(f32, f32)>,
    /// S851: an inline OLEObject-less `<w:object>` (a bare form-field picture
    /// shape — no `<o:OLEObject>` child) that flows in its host line. The
    /// FFFC object fragment (created from `inline_object_extent`) draws THIS
    /// image at the fragment position instead of a vector group. None
    /// everywhere except inline form-field w:objects; real OLE (Equation /
    /// Visio) keeps the block-extraction path, so the canaries are unaffected.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inline_object_image: Option<Box<Image>>,
    /// S852: an inline VML horizontal rule (`<v:rect o:hr="t" .../>`). Carried
    /// on the same run whose `inline_object_extent` reserves the rule's own
    /// line; the emit draws a full-width gray line instead of an image. Tuple =
    /// (thickness_pt, hex_color). None everywhere except o:hr runs (forms only
    /// by corpus scan; JP has 0 → byte-identical).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hr_rule: Option<(f32, String)>,
}

impl Default for RunStyle {
    fn default() -> Self {
        Self {
            font_family: None,
            font_family_east_asia: None,
            has_explicit_east_asia: false,
            east_asia_lang: None,
            latin_lang: None,
            font_size: None,
            bold: false,
            has_explicit_bold: false,
            italic: false,
            has_explicit_italic: false,
            underline: false,
            underline_style: None,
            strikethrough: false,
            has_explicit_strikethrough: false,
            double_strikethrough: false,
            has_explicit_double_strikethrough: false,
            color: None,
            highlight: None,
            vertical_align: None,
            character_spacing: None,
            small_caps: false,
            has_explicit_small_caps: false,
            all_caps: false,
            has_explicit_all_caps: false,
            shading: None,
            rtl: false,
            vanish: false,
            outline: false,
            shadow: false,
            emboss: false,
            imprint: false,
            font_size_cs: None,
            bold_cs: false,
            italic_cs: false,
            kern: None,
            fit_text: None,
            fit_text_id: None,
            text_scale: None,
            combine: false,
            combine_brackets: None,
            vert_in_horz: false,
            position: None,
            emphasis_mark: None,
            run_border: None,
            inline_object_extent: None,
            inline_object_image: None,
            hr_rule: None,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq)]
pub enum VerticalAlign {
    Baseline,
    Superscript,
    Subscript,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Table {
    pub rows: Vec<TableRow>,
    pub style: TableStyle,
    /// Column widths from tblGrid/gridCol in points
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub grid_columns: Vec<f32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
    /// Row height in points (w:trHeight)
    #[serde(default)]
    pub height: Option<f32>,
    /// Height rule: "exact" or "atLeast" (default)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height_rule: Option<String>,
    /// Repeat as header row at top of each page (w:tblHeader)
    #[serde(default)]
    pub header: bool,
    /// Prevent row from breaking across pages (w:cantSplit)
    #[serde(default)]
    pub cant_split: bool,
    /// Number of grid columns to skip at start of row (w:gridBefore)
    #[serde(default)]
    pub grid_before: u32,
    /// Row-level cell margin override from w:tblPrEx/w:tblCellMar
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_margins_override: Option<CellMargins>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableCell {
    pub blocks: Vec<Block>,
    pub width: Option<f32>,
    /// Horizontal merge span (w:gridSpan), default 1
    #[serde(default = "default_one")]
    pub grid_span: u32,
    /// Vertical merge: "restart" starts a new merged cell, "continue" is merged into above
    #[serde(default)]
    pub v_merge: Option<String>,
    /// Cell shading/background color (hex)
    #[serde(default)]
    pub shading: Option<String>,
    /// Vertical alignment within cell
    #[serde(default)]
    pub v_align: Option<String>,
    /// Cell-specific borders (override table borders)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub borders: Option<CellBorders>,
    /// Cell margins/padding in points
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margins: Option<CellMargins>,
    /// Text direction within cell (w:textDirection): "btLr" (bottom-to-top LR), "tbRl" (top-bottom RL), etc.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_direction: Option<String>,
    /// S486: floating text boxes anchored to a paragraph INSIDE this cell.
    /// parse_table_cell historically discarded these (only `blocks` survived),
    /// so in-cell callouts/annotations (e.g. 1636d28, 664c38, 9a8e8d) never
    /// rendered. Preserved here for the in-cell anchor-resolution layout step.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cell_text_boxes: Vec<TextBox>,
    /// S486: floating shapes anchored inside this cell (same drop as above).
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cell_shapes: Vec<Shape>,
    /// w:hideMark (ECMA-376 17.4.22): ignore the end-of-cell mark for row
    /// height — an EMPTY hideMark cell contributes zero content height
    /// (the thin-spacer-row idiom; S751).
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub hide_mark: bool,
}

/// Cell border definitions
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellBorders {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<BorderDef>,
}

/// Cell margin/padding
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CellMargins {
    #[serde(default)]
    pub top: Option<f32>,
    #[serde(default)]
    pub bottom: Option<f32>,
    #[serde(default)]
    pub left: Option<f32>,
    #[serde(default)]
    pub right: Option<f32>,
}

fn default_one() -> u32 { 1 }

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Image {
    pub data: Vec<u8>,
    pub width: f32,
    pub height: f32,
    pub alt_text: Option<String>,
    pub content_type: Option<String>,
    /// Floating position (None = inline image)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<FloatingPosition>,
    /// Text wrapping mode
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap_type: Option<WrapType>,
    /// Crop percentages (a:srcRect) — top, right, bottom, left as 0-100%
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub crop: Option<ImageCrop>,
    /// Index of the anchor paragraph block (for paragraph-relative positioning)
    #[serde(default)]
    pub anchor_block_index: usize,
    /// S765 (2026-07-08): wp:anchor z-order — Word draws floating objects
    /// (images AND textboxes) in ascending relativeHeight (highest on top);
    /// carried so the layout can interleave images and textboxes in ONE
    /// z-order pass. Default 0.
    #[serde(default)]
    pub relative_height: u32,
    /// behindDoc=1 places the object behind body text. Default false.
    #[serde(default)]
    pub behind_doc: bool,
    /// S965: the resolved before/after spacing of the image-only HOST paragraph
    /// (S537 lowers that paragraph to a bare `Block::Image`, so without this the
    /// spacing at both of its boundaries is lost). Layout collapses
    /// `max(prev.after, before)` above the image and exposes `after` to the next
    /// paragraph — the ordinary paragraph rule. Zero for images that are not the
    /// sole content of their paragraph.
    #[serde(default)]
    pub paragraph_space_before: f32,
    #[serde(default)]
    pub paragraph_space_after: f32,
    /// S971: the image-only HOST paragraph, runs removed. Word's inline-image
    /// line is `max(host paragraph line, image extent)` (measured — see
    /// tools/metrics/_pb_imgline_gen.py), and S536/S537 drop that paragraph, so
    /// without it the line is the extent alone and a small image (a spacer gif)
    /// under-counts by the whole line. Only the STYLE is needed; the parser
    /// cannot compute a line height because it has no font metrics.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub host_paragraph: Option<Box<Paragraph>>,
}

/// Image crop rectangle (percentages from each edge)
#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
pub struct ImageCrop {
    #[serde(default)]
    pub top: f32,
    #[serde(default)]
    pub right: f32,
    #[serde(default)]
    pub bottom: f32,
    #[serde(default)]
    pub left: f32,
}

/// Position for a floating (anchored) element
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FloatingPosition {
    /// Horizontal offset in points from anchor
    pub x: f32,
    /// Vertical offset in points from anchor
    pub y: f32,
    /// Horizontal anchor reference (e.g. "column", "page", "margin")
    #[serde(default)]
    pub h_relative: Option<String>,
    /// Vertical anchor reference
    #[serde(default)]
    pub v_relative: Option<String>,
    /// Horizontal alignment (e.g. "left", "center", "right")
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_align: Option<String>,
    /// Vertical alignment (e.g. "top", "center", "bottom")
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_align: Option<String>,
    /// wrapSquare keep-out distance left of the float (wp:anchor distL, pt)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dist_l: Option<f32>,
    /// wrapSquare keep-out distance right of the float (wp:anchor distR, pt)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub dist_r: Option<f32>,
}

/// Text wrapping mode for floating elements
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum WrapType {
    /// No wrapping (in front/behind text)
    None,
    /// Square wrapping
    Square,
    /// Tight wrapping
    Tight,
    /// Top and bottom only
    TopAndBottom,
}

/// A text box (from w:txbxContent or wps:txbx)
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TextBox {
    /// Content paragraphs inside the text box
    pub blocks: Vec<Block>,
    /// Width in points
    pub width: f32,
    /// Height in points
    pub height: f32,
    /// Position (floating)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<FloatingPosition>,
    /// Border style
    #[serde(default)]
    pub border: bool,
    /// Border stroke color (hex)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stroke_color: Option<String>,
    /// Border stroke width in points
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stroke_width: Option<f32>,
    /// Background color (hex)
    #[serde(default)]
    pub fill: Option<String>,
    /// Index of the anchor block (paragraph) in page.blocks
    #[serde(default)]
    pub anchor_block_index: usize,
    /// Corner radius for rounded rectangles (in points)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub corner_radius: Option<f32>,
    /// Text inset left (in points, default 7.2pt = 91440 EMU)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inset_left: Option<f32>,
    /// Text inset right (in points, default 7.2pt)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inset_right: Option<f32>,
    /// Text inset top (in points, default 3.6pt = 45720 EMU)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inset_top: Option<f32>,
    /// Text inset bottom (in points, default 3.6pt)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inset_bottom: Option<f32>,
    /// Wrap type for text wrapping around this text box
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap_type: Option<WrapType>,
    /// Vertical text anchor: "top" (default), "middle", "bottom"
    /// From VML v-text-anchor or DrawingML bodyPr anchor attribute.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_text_anchor: Option<String>,
    /// S478: wp:anchor relativeHeight — Word draws floating objects in
    /// ascending relativeHeight order (highest on top). Default 0.
    #[serde(default)]
    pub relative_height: u32,
    /// S478: wp:anchor behindDoc — true = behind body text.
    #[serde(default)]
    pub behind_doc: bool,
    /// S481: bodyPr@vertOverflow — "overflow" (ECMA default, do not clip
    /// vertically), "clip", or "ellipsis". None = overflow (default).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vert_overflow: Option<String>,
    /// S662: bodyPr@compatLnSpc="1" — "compatible line spacing" (legacy Word
    /// line-spacing model). With it set, Word places textbox text ~2.5pt LOWER
    /// (the line's leading sits above the first baseline); Oxi's default
    /// placement is ~2.5pt too HIGH. Used to apply a render-only text DY scoped
    /// to compatLnSpc=1 textboxes. Default false (most textboxes lack the attr).
    #[serde(default)]
    pub compat_line_spacing: bool,
    /// S839 (2026-07-14): drawable vector primitives extracted from a wpg
    /// group's rect/line-class member shapes (hmrc-class form furniture:
    /// checkbox strips, writing boxes, heavy rules). Attached only for
    /// VISUAL-ONLY groups (no txbxContent) so text-bearing groups (framework
    /// cover pages) keep their legacy merged-textbox behavior byte-identical.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub vector_shapes: Vec<VectorShape>,
}

/// S839: one drawable vector primitive from a wpg/wps drawing group.
/// Coordinates in pt relative to the drawing extent's top-left (group
/// child-space transform already applied). `is_line` draws a straight
/// segment of `stroke_width` thickness; otherwise a rectangle outline
/// (`stroke`) and/or fill.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct VectorShape {
    pub x: f32,
    pub y: f32,
    pub w: f32,
    pub h: f32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub stroke: Option<String>,
    #[serde(default)]
    pub stroke_width: f32,
    #[serde(default)]
    pub is_line: bool,
}

/// A geometric shape (DrawingML or VML)
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Shape {
    /// Shape type (e.g. "rect", "ellipse", "roundRect", "line", "arrow", etc.)
    pub shape_type: String,
    /// Width in points
    pub width: f32,
    /// Height in points
    pub height: f32,
    /// Position (floating)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<FloatingPosition>,
    /// Fill color (hex)
    #[serde(default)]
    pub fill: Option<String>,
    /// Outline/stroke color (hex)
    #[serde(default)]
    pub stroke_color: Option<String>,
    /// Outline width in points
    #[serde(default)]
    pub stroke_width: Option<f32>,
    /// Text content inside the shape (if any)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub text_blocks: Vec<Block>,
    /// Rotation in degrees
    #[serde(default)]
    pub rotation: Option<f32>,
    /// Gradient fill stops (from a:gradFill)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub gradient_stops: Vec<GradientStop>,
    /// Gradient fill angle in degrees
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub gradient_angle: Option<f32>,
    /// Index of the anchor paragraph block (for positioning)
    #[serde(default)]
    pub anchor_block_index: usize,
    /// Vertical text anchor: "top" (default), "middle", "bottom"
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_text_anchor: Option<String>,
    /// Horizontal flip (a:xfrm flipH) — sets the connector diagonal direction.
    #[serde(default)]
    pub flip_h: bool,
    /// Vertical flip (a:xfrm flipV).
    #[serde(default)]
    pub flip_v: bool,
    /// Connector arrowhead at the start (a:ln/a:headEnd type≠none).
    #[serde(default)]
    pub arrow_head: bool,
    /// Connector arrowhead at the end (a:ln/a:tailEnd type≠none).
    #[serde(default)]
    pub arrow_tail: bool,
    /// S711: true when this shape came from the legacy VML path (`<w:pict>`),
    /// false for DrawingML (`<a:...>`). A filled VML rect renders its fill (the
    /// (注) legend box); a DrawingML rect keeps the outline-only PresetShape —
    /// filling DML form rects regressed 2ea81a/1ec1 (SSIM A/B).
    #[serde(default)]
    pub is_vml: bool,
    /// S711b: VML `o:allowincell="f"` — the shape ESCAPES the table cell and is
    /// positioned relative to the page text column (start_x), NOT the cell content
    /// edge. The (注) gray legend box uses this (Word anchors it at page-margin +
    /// margin-left, not cell_x+pad). Default false = normal cell-relative.
    #[serde(default)]
    pub escapes_cell: bool,
}

/// A gradient color stop
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct GradientStop {
    /// Position as 0-100 percentage
    pub position: f32,
    /// Color hex
    pub color: String,
}

/// A comment annotation
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Comment {
    /// Comment ID
    pub id: String,
    /// Author name
    pub author: Option<String>,
    /// Date string
    pub date: Option<String>,
    /// Author initials (1–6 glyphs; ECMA-376 §17.13.4.2 w:initials)
    pub initials: Option<String>,
    /// `w14:paraId` of the comment body's first paragraph — join key used by
    /// `word/commentsExtended.xml` (ECMA-376 §17.13.1 + MS-DOCX w15 extensions).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub para_id: Option<String>,
    /// `w15:paraIdParent` from commentsExtended.xml — set when this comment is a
    /// reply; points to the `para_id` of the parent comment.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub parent_para_id: Option<String>,
    /// `w15:done="1"` from commentsExtended.xml — Word 2013+ resolved state.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub resolved: bool,
    /// `w16cid:durableId` from `word/commentsIds.xml` — Word 2019+ identifier
    /// that survives save-as round-trips (local `w:id` is freely renumbered).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub durable_id: Option<String>,
    /// Comment text paragraphs
    pub blocks: Vec<Block>,
}

/// Author entry from `word/people.xml` (MS-DOCX w15 extension).
///
/// Word writes one `<w15:person>` per distinct reviewer. `author` is the
/// display name used as the join key to `<w:comment w:author>`, `<w:ins
/// w:author>`, etc. `provider_id` + `user_id` come from the nested
/// `<w15:presenceInfo>` and identify the reviewer across sessions — Word
/// uses the pair to colour-code revisions even when two authors share a
/// display name.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Person {
    /// Display name (`w15:author`).
    pub author: String,
    /// Presence provider id (e.g., "AD" for Active Directory, "None" when absent).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub provider_id: Option<String>,
    /// Provider-specific user id (often an email or the display name again when
    /// `provider_id == "None"`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub user_id: Option<String>,
}

/// A property-change revision (`<w:rPrChange>`, `<w:pPrChange>`, etc.).
///
/// `prior_run_style` stores the run's style as it was before the change, so
/// "Original" / "Simple markup" views can reconstruct the pre-edit document
/// without needing the original XML.
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct PropertyChange {
    /// `w:id` attribute on the change element (document-local, not durable).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub id: Option<String>,
    /// `w:author` attribute.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    /// `w:date` attribute.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub date: Option<String>,
    /// Prior run style (body of `<w:rPrChange>/<w:rPr>`). Boxed to keep `Run`
    /// small in the common (no-change) case.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prior_run_style: Option<Box<RunStyle>>,
    /// Prior paragraph style (body of `<w:pPrChange>/<w:pPr>`). Boxed for the
    /// same reason as `prior_run_style`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prior_paragraph_style: Option<Box<ParagraphStyle>>,
    /// Prior paragraph alignment (`<w:jc>` inside `<w:pPrChange>/<w:pPr>`).
    /// `Paragraph.alignment` is a top-level IR field separate from
    /// `ParagraphStyle`, so a pPrChange that toggles alignment can't ride
    /// on `prior_paragraph_style` alone — the parser captures alignment
    /// here when the prior pPr declares `<w:jc>`. R-12 v3.5 (R72,
    /// 2026-04-29) consumes this to surface "Alignment: …" in the
    /// "Formatted" margin balloon.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prior_alignment: Option<Alignment>,
}

/// Resolved author palette entry — `display` is the join key, `color_index` is
/// the position the renderer uses to look up an RGB swatch.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Author {
    /// Display name (matches `<w:author>` attributes verbatim).
    pub display: String,
    /// 0-based palette index. Stable for a given document — derived from first-seen
    /// order across people.xml + comments + tracked changes.
    pub color_index: usize,
}

/// Reveal mode for revisions, mirroring Word's "Show Markup" / "Display for
/// Review" dropdown. Set on the render config; the renderer uses it to decide
/// which revisions to draw and how (attack-matrix row S-02).
///
/// - `All` — every revision is rendered with markup (default).
/// - `Simple` — vertical change bar in the margin only; in-line text shows
///   the post-edit document.
/// - `Original` — pre-edit document: insertions are hidden, deletions appear
///   as normal text, prior `*PrChange` styles are applied.
/// - `Final` — post-edit document: deletions are hidden, insertions are
///   normal text, `*PrChange` keeps the current style.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "snake_case")]
pub enum ShowRevisions {
    #[default]
    All,
    Simple,
    Original,
    Final,
}

/// A tracked change (insertion, deletion, or move).
///
/// `change_type` is one of: `"insert"` (`<w:ins>`), `"delete"` (`<w:del>`),
/// `"moveFrom"` (`<w:moveFrom>`), `"moveTo"` (`<w:moveTo>`). `pair_id` is
/// the `w:id` attribute, used by the renderer to pair a `moveFrom` with its
/// `moveTo` across paragraphs (ECMA-376 §17.13.5).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TrackedChange {
    /// "insert" | "delete" | "moveFrom" | "moveTo"
    pub change_type: String,
    /// Author of the change
    pub author: Option<String>,
    /// Date of the change
    pub date: Option<String>,
    /// Document-local `w:id` of the revision. For moves, `pair_id` on a
    /// `moveFrom` matches the `pair_id` on its partner `moveTo`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pair_id: Option<String>,
}

/// Alignment of ruby annotation relative to base text.
/// Per ECMA-376 §17.3.3.26 (CT_RubyAlign).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum RubyAlign {
    /// Center: ruby and base both centered in field (default).
    Center,
    /// Distribute ruby letters evenly across the base width.
    DistributeLetter,
    /// Distribute ruby letters with extra padding at both ends.
    DistributeSpace,
    /// Both ruby and base left-aligned.
    Left,
    /// Both ruby and base right-aligned.
    Right,
    /// For vertical writing only.
    RightVertical,
}

/// Ruby (furigana) annotation
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Ruby {
    /// Base text (the main character(s))
    pub base: String,
    /// Annotation text (furigana reading)
    pub text: String,
    /// Font size of the ruby text in points (derived from `hps`).
    #[serde(default)]
    pub font_size: Option<f32>,
    /// Alignment of annotation relative to base (`<w:rubyAlign>`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<RubyAlign>,
    /// Ruby font size in half-points (`<w:hps w:val="...">`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hps_halfpt: Option<u32>,
    /// Distance ruby annotation is raised above base baseline, half-points
    /// (`<w:hpsRaise w:val="...">`). Default is approximately 18 (= 9pt)
    /// per Word's empirical behavior; see spec §18.4.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hps_raise_halfpt: Option<u32>,
    /// Base text font size in half-points (`<w:hpsBaseText w:val="...">`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hps_base_text_halfpt: Option<u32>,
    /// Language code for the ruby annotation (`<w:lid w:val="...">`).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub lang: Option<String>,
}

/// Column layout definition (from w:cols)
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColumnLayout {
    /// Number of columns
    pub num: u32,
    /// Space between columns in points
    #[serde(default)]
    pub space: Option<f32>,
    /// Whether columns have equal width
    #[serde(default)]
    pub equal_width: bool,
    /// Draw a separator line between columns (`w:sep="1"`). Currently
    /// rendered only for vertical-writing sections (between horizontal bands).
    #[serde(default)]
    pub separator: bool,
    /// Individual column definitions (for unequal widths)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<ColumnDef>,
}

/// Individual column definition
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColumnDef {
    /// Column width in points
    pub width: f32,
    /// Space after this column in points
    #[serde(default)]
    pub space: Option<f32>,
}

/// A tab stop definition
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TabStop {
    /// Position in points from the left margin
    pub position: f32,
    /// Alignment at the tab stop
    pub alignment: TabStopAlignment,
    /// Leader character
    #[serde(default)]
    pub leader: Option<String>,
    /// S977: this entry is a `<w:tab w:val="clear"/>` — it REMOVES an inherited
    /// stop at `position` rather than being a stop itself (ECMA-376
    /// §17.3.1.38). Never reaches layout: the merge drops it.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub clear: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum TabStopAlignment {
    Left,
    Center,
    Right,
    Decimal,
}

impl Default for TabStopAlignment {
    fn default() -> Self {
        Self::Left
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
    Distribute,
}

impl Default for Alignment {
    fn default() -> Self {
        Self::Left
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ParagraphStyle {
    /// S730 (2026-07-03): this paragraph carries the sectPr of a section
    /// merged as CONTINUOUS into the same IR Page. Word renders the empty
    /// continuous section-break paragraph at ZERO height (probexmargins
    /// Word COM: the break para's Info(6) y equals the previous paragraph's
    /// last-line row — no new line); Oxi gave it a normal empty-para line
    /// (+18pt drift for everything below the break). Layout skips it.
    #[serde(default)]
    pub continuous_section_break: bool,
    /// S945 (2026-07-19): this paragraph carries an in-body sectPr (it ENDS a
    /// section). An EMPTY section-ending paragraph never opens a new page by
    /// natural overflow in Word — the next section's page break follows
    /// immediately, so its height on a fresh page is unobservable and Oxi's
    /// normal empty-line overflow manufactured a phantom page (NDIS wp41/42).
    #[serde(default)]
    pub page_section_break: bool,
    pub heading_level: Option<u8>,
    /// Outline level from w:outlineLvl (0-8, for TOC generation, NOT for layout)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub outline_level: Option<u8>,
    /// Line spacing value. Interpretation depends on line_spacing_rule:
    /// - "auto": multiplier (w:line / 240, e.g. 1.15 for w:line="276")
    /// - "exact": fixed height in points (w:line / 20)
    /// - "atLeast": minimum height in points (w:line / 20)
    /// None means single spacing (1.0).
    pub line_spacing: Option<f32>,
    /// Line spacing rule: "auto" (default), "exact", or "atLeast"
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_spacing_rule: Option<String>,
    pub space_before: Option<f32>,
    pub space_after: Option<f32>,
    /// True when spacing was directly specified in paragraph's pPr (not just inherited from style).
    /// Word resets inherited spacing to Single/0 inside table cells.
    #[serde(default)]
    pub has_direct_spacing: bool,
    /// S855 (2026-07-15): True when the DIRECT pPr `<w:spacing>` set before/after
    /// (or beforeLines/afterLines/*Autospacing) — as opposed to setting ONLY the
    /// line multiplier. A direct line-only `<w:spacing w:line=…>` sets
    /// `has_direct_spacing` but must NOT prevent the table-cell before/after
    /// reset: Word still resets the docDefaults-inherited before/after to the
    /// cell default (0) when the paragraph does not itself set them.
    #[serde(default)]
    pub has_direct_before_after: bool,
    /// S909: DIRECT pPr carries <w:tabs> / <w:ind> (as opposed to
    /// style-inherited tab stops / indents — ea8f's pStyle "Footer" tabs
    /// must NOT defeat the untouched-blank-footer exemption; bd832's
    /// direct tabs+ind must).
    #[serde(default)]
    pub has_direct_tabs_or_ind: bool,
    /// S865: DIRECT pPr spacing specifies a before-side value
    /// (before/beforeLines/beforeAutospacing). Kept per-side because Word
    /// resets an unspecified opposite side to the table-cell default.
    #[serde(default)]
    pub has_direct_before: bool,
    /// S865: DIRECT pPr spacing specifies an after-side value.
    #[serde(default)]
    pub has_direct_after: bool,
    /// True when line_spacing was inherited from docDefaults pPrDefault (not from Normal style or direct).
    /// Word resets docDefaults lineSpacing to Single inside table cells but keeps Normal style's lineSpacing.
    #[serde(default)]
    pub line_spacing_from_doc_defaults: bool,
    /// S906: space_before/space_after were merged from docDefaults (not a
    /// named/default paragraph style). The table-style cell spacing layer
    /// overrides docDefaults-sourced spacing but NOT style-sourced spacing
    /// (ECMA: docDefaults < table style pPr < paragraph style < direct).
    #[serde(default)]
    pub space_before_from_doc_defaults: bool,
    #[serde(default)]
    pub space_after_from_doc_defaults: bool,
    /// S935: default_run_style.font_size was merged from docDefaults (not a
    /// named/default paragraph style). The table-style rPr sz layer
    /// overrides docDefaults-sourced size but NOT style-sourced size
    /// (ECMA: docDefaults < table style rPr < paragraph style < direct).
    #[serde(default)]
    pub font_size_from_doc_defaults: bool,
    /// w:spacing beforeLines — in 1/100 of a line (raw value from OOXML)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub before_lines: Option<f32>,
    /// w:spacing afterLines — in 1/100 of a line (raw value from OOXML)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub after_lines: Option<f32>,
    pub indent_left: Option<f32>,
    pub indent_right: Option<f32>,
    pub indent_first_line: Option<f32>,
    /// Raw leftChars/startChars value (hundredths of a character width).
    /// Resolved to indent_left at layout time using grid_char_pitch.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent_left_chars: Option<f32>,
    /// Raw rightChars/endChars value (hundredths of a character width).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent_right_chars: Option<f32>,
    /// Raw firstLineChars value (hundredths of a character width).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent_first_line_chars: Option<f32>,
    /// Default run style from style definition (font size, bold, etc.)
    pub default_run_style: Option<RunStyle>,
    /// Pre-resolved list marker text (e.g., "•", "1.", "a)")
    pub list_marker: Option<String>,
    /// Hanging indent for the list marker in points
    pub list_indent: Option<f32>,
    /// Suffix after list number: "tab" (default), "space", or "nothing"
    #[serde(default)]
    pub list_suff: Option<String>,
    /// Tab stop position for list numbering (in points)
    #[serde(default)]
    pub list_tab_stop: Option<f32>,
    /// S801b: the numbering level rPr's marker font size (pt), if declared.
    pub list_marker_size: Option<f32>,
    /// S778: the numbering LEVEL's ind left (pt) — the marker suffix-tab stop.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub list_level_left: Option<f32>,
    /// Whether this paragraph snaps to the document grid (default: true).
    #[serde(default = "default_true")]
    pub snap_to_grid: bool,
    /// Whether snapToGrid was explicitly set in the paragraph's direct pPr.
    /// A direct `<w:snapToGrid/>` (CT_OnOff, no val = true) overrides a style's
    /// snapToGrid=0 — Word COM-confirmed (S606b: ohnoikuji a4 "header"-style
    /// list items carry a direct no-val snapToGrid re-enabling grid snap).
    #[serde(default, skip_serializing)]
    pub has_explicit_snap_to_grid: bool,
    /// w:contextualSpacing: suppress space_before/after between paragraphs of the same style.
    #[serde(default)]
    pub contextual_spacing: bool,
    /// w:spacing@beforeAutospacing — Word applies a flat ~13.75pt auto before-space
    /// (COM-derived 2026-06-26, S675: constant, independent of font/grid/docDefaults;
    /// = Word's hardcoded Normal-Web 11pt×1.25). Overrides explicit before; collapses MAX.
    #[serde(default)]
    pub before_autospacing: bool,
    /// w:spacing@afterAutospacing — flat ~13.75pt auto after-space (see before_autospacing).
    #[serde(default)]
    pub after_autospacing: bool,
    /// S895: the autospacing flags came from the paragraph STYLE (not direct
    /// pPr). Word applies STYLE-level HTML autospacing in Latin docs
    /// (legal__00081e80 Metadata style: measured ~13.95/gap) while the JP
    /// evidence (S675: harassbosi/b837 Web styles render 0) keeps style-level
    /// autospacing inert for CJK docs — the layout gates style-sourced flags
    /// on !doc_body_has_real_cjk.
    #[serde(default, skip_serializing)]
    pub autospacing_from_style: bool,
    /// Style ID (e.g. "Normal", "Heading1") for contextual spacing comparison.
    #[serde(default)]
    pub style_id: Option<String>,
    /// Custom tab stops (w:tabs)
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tab_stops: Vec<TabStop>,
    /// Paragraph background/shading color (hex from w:shd fill)
    #[serde(default)]
    pub shading: Option<String>,
    /// pPr/rPr: paragraph-level default run properties for empty paragraph height.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ppr_rpr: Option<RunStyle>,
    /// Page break before this paragraph (w:pageBreakBefore)
    #[serde(default)]
    pub page_break_before: bool,
    /// Whether pageBreakBefore was explicitly set in the paragraph's direct pPr.
    /// A direct `<w:pageBreakBefore w:val="0"/>` overrides a style's
    /// pageBreakBefore (S884 — same tri-state pattern as has_explicit_snap_to_grid).
    #[serde(default, skip_serializing)]
    pub has_explicit_page_break_before: bool,
    /// Page break AFTER this paragraph (empty-paragraph-with-inline-br pattern).
    /// Word renders the empty paragraph's mark on the CURRENT page then breaks.
    #[serde(default)]
    pub page_break_after: bool,
    /// Paragraph borders
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub borders: Option<ParagraphBorders>,
    /// Keep with next paragraph on same page (w:keepNext)
    #[serde(default)]
    pub keep_next: bool,
    /// S955: keepNext was EXPLICITLY set (val present or element present) —
    /// distinguishes an explicit `<w:keepNext w:val="0"/>` (which must beat a
    /// basedOn parent's / paragraph style's ON) from "not set" (which
    /// inherits). The widow_control three-state pattern.
    #[serde(default)]
    pub has_explicit_keep_next: bool,
    /// Keep all lines of this paragraph together (w:keepLines)
    #[serde(default)]
    pub keep_lines: bool,
    /// S955: keepLines three-state marker (see has_explicit_keep_next).
    #[serde(default)]
    pub has_explicit_keep_lines: bool,
    /// Widow/orphan control (w:widowControl, default true in Word)
    #[serde(default = "default_true")]
    pub widow_control: bool,
    /// Whether widowControl was explicitly set in XML (for docDefaults inheritance)
    #[serde(default, skip_serializing)]
    pub has_explicit_widow_control: bool,
    /// S782: whether contextualSpacing was explicitly set in the DIRECT pPr —
    /// a direct `<w:contextualSpacing w:val="0"/>` must not be clobbered by
    /// the style merge (nyserda ListParagraph items disable the style's
    /// contextualSpacing to restore their direct before=240).
    #[serde(default, skip_serializing)]
    pub has_explicit_contextual_spacing: bool,
    /// Word wrap at CJK character boundaries (w:wordWrap, default true)
    /// When false, lines break only at spaces (no CJK inter-character break)
    #[serde(default = "default_true")]
    pub word_wrap: bool,
    /// Automatically adjust right indent for CJK grid (w:adjustRightInd, default true)
    #[serde(default = "default_true")]
    pub adjust_right_ind: bool,
    /// Auto space between East Asian and Western text (w:autoSpaceDE, default true)
    #[serde(default = "default_true")]
    pub auto_space_de: bool,
    /// Auto space between East Asian and numbers (w:autoSpaceDN, default true)
    #[serde(default = "default_true")]
    pub auto_space_dn: bool,
    /// Text alignment within line (w:textAlignment): "top", "center", "baseline", "bottom", "auto"
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_alignment: Option<String>,
    /// R7.63 (Day 36 part 10, 2026-05-14): true when text_alignment was inherited
    /// from pPrDefault, false when set per-paragraph or via paragraph style chain.
    /// text_y_offset_for_line uses this to gate the "baseline → offset=0" rule:
    /// pPrDefault baseline applies document-wide and suppresses centering for all
    /// paragraphs (e3c545 case); per-paragraph baseline does NOT suppress centering
    /// (ed025c wi=827 case — only that one paragraph has it, breaking gap to wi=826).
    #[serde(default)]
    pub text_alignment_from_pprdefault: bool,
    /// Frame paragraph properties (w:framePr) — for drop caps and positioned paragraphs
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub frame_pr: Option<FrameProperties>,
    /// Bidirectional text / RTL paragraph (w:bidi)
    #[serde(default)]
    pub bidi: bool,
    /// Numbering ID from style definition (w:numPr/w:numId)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub num_id: Option<String>,
    /// Numbering indent level from style definition (w:numPr/w:ilvl)
    #[serde(default)]
    pub num_ilvl: u8,
}

/// Paragraph border definitions
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct ParagraphBorders {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<BorderDef>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub between: Option<BorderDef>,
}

/// A single border definition
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct BorderDef {
    /// Border style (e.g. "single", "double", "dashed", "dotted", "thick")
    pub style: String,
    /// Width in points (w:sz is in 1/8 pt)
    pub width: f32,
    /// Color hex
    pub color: Option<String>,
    /// Distance from text in points (w:space)
    #[serde(default)]
    pub space: f32,
}

fn default_true() -> bool { true }

impl Default for ParagraphStyle {
    fn default() -> Self {
        Self {
            continuous_section_break: false,
            page_section_break: false,
            heading_level: None,
            outline_level: None,
            line_spacing: None,
            line_spacing_rule: None,
            space_before: None,
            space_after: None,
            has_direct_spacing: false,
            has_direct_before_after: false,
            has_direct_tabs_or_ind: false,
            has_direct_before: false,
            has_direct_after: false,
            line_spacing_from_doc_defaults: false,
            space_before_from_doc_defaults: false,
            font_size_from_doc_defaults: false,
            space_after_from_doc_defaults: false,
            before_lines: None,
            after_lines: None,
            indent_left: None,
            indent_right: None,
            indent_first_line: None,
            indent_left_chars: None,
            indent_right_chars: None,
            indent_first_line_chars: None,
            default_run_style: None,
            list_marker: None,
            list_indent: None,
            list_suff: None,
            list_tab_stop: None,
            list_marker_size: None,
            list_level_left: None,
            snap_to_grid: true,
            has_explicit_snap_to_grid: false,
            contextual_spacing: false,
            before_autospacing: false,
            after_autospacing: false,
            autospacing_from_style: false,
            style_id: None,
            tab_stops: Vec::new(),
            shading: None,
            ppr_rpr: None,
            page_break_before: false,
            has_explicit_page_break_before: false,
            page_break_after: false,
            borders: None,
            keep_next: false,
            has_explicit_keep_next: false,
            keep_lines: false,
            has_explicit_keep_lines: false,
            widow_control: true,
            has_explicit_widow_control: false,
            has_explicit_contextual_spacing: false,
            word_wrap: true,
            adjust_right_ind: true,
            text_alignment: None,
            text_alignment_from_pprdefault: false,
            auto_space_de: true,
            auto_space_dn: true,
            frame_pr: None,
            bidi: false,
            num_id: None,
            num_ilvl: 0,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct TableStyle {
    pub border: bool,
    /// True iff `<w:tblBorders>` was directly in this table's `<w:tblPr>`,
    /// not inherited from a `<w:tblStyle>`. Used to gate the b35-style
    /// "border-at-margin-minus-padding" offset (gen2_052 needs it,
    /// 683ff with explicit borders does NOT — Word renders 683ff border
    /// at margin without subtracting padding).
    #[serde(default)]
    pub explicit_borders: bool,
    /// Whether the table has inside horizontal borders (insideH)
    #[serde(default)]
    pub has_inside_h: bool,
    /// Whether the table has inside vertical borders (insideV).
    #[serde(default)]
    pub has_inside_v: bool,
    /// Inside horizontal border, including `style="none"` when suppressed.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inside_horizontal_border: Option<BorderDef>,
    /// Inside vertical border, including `style="none"` when suppressed.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inside_vertical_border: Option<BorderDef>,
    /// Border color (hex), e.g. "000000"
    #[serde(default)]
    pub border_color: Option<String>,
    /// Border width in points (w:sz is in 1/8 pt)
    #[serde(default)]
    pub border_width: Option<f32>,
    /// Border style (e.g. "single", "double", "dashed")
    #[serde(default)]
    pub border_style: Option<String>,
    /// Table width in points (from w:tblW)
    #[serde(default)]
    pub width: Option<f32>,
    /// Table width type: "dxa" (fixed), "pct" (percentage), "auto"
    #[serde(default)]
    pub width_type: Option<String>,
    /// Table alignment (w:jc): "left", "center", "right"
    #[serde(default)]
    pub alignment: Option<String>,
    /// Table style ID reference (w:tblStyle)
    #[serde(default)]
    pub style_id: Option<String>,
    /// Table look flags (from w:tblLook)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tbl_look: Option<TableLook>,
    /// Table indent from left margin in points (w:tblInd)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent: Option<f32>,
    /// Cell spacing in points (w:tblCellSpacing)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_spacing: Option<f32>,
    /// Table layout mode: "fixed" or "autofit" (w:tblLayout)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout: Option<String>,
    /// Default cell margins in points (w:tblCellMar)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_cell_margins: Option<CellMargins>,
    /// True iff `<w:tblCellMar>` was directly in this table's `<w:tblPr>`,
    /// not inherited from a `<w:tblStyle>` or defaulted. Mirrors
    /// `explicit_borders`. Distinguishes author-declared cell margins
    /// (the S412 cellMar wrap-budget discriminator) from style-inherited
    /// or OOXML-default margins. `default_cell_margins.is_some()` is too
    /// broad for that gate (S417e: it fired on 04b88e — which has default
    /// margins but no explicit tblCellMar — and regressed it).
    #[serde(default)]
    pub has_explicit_cellmar: bool,
    /// Paragraph properties from table style (pPr) — applied to cell paragraphs as fallback
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub para_style: Option<ParagraphStyle>,
    /// Paragraph alignment from table style pPr (jc) — applied as fallback
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub para_alignment: Option<Alignment>,
    /// Run font size (pt) from the table style chain's top-level `w:rPr`
    /// (S935): a table style's run properties apply to every run in the
    /// table — above docDefaults, below the paragraph-style chain /
    /// character style / direct rPr.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub run_font_size: Option<f32>,
    /// Table floating position (w:tblpPr)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<TablePosition>,
}

/// Floating table position properties (w:tblpPr)
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TablePosition {
    /// Horizontal offset in points (w:tblpX)
    #[serde(default)]
    pub x: f32,
    /// Vertical offset in points (w:tblpY)
    #[serde(default)]
    pub y: f32,
    /// Horizontal anchor: "text", "margin", "page"
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_anchor: Option<String>,
    /// Vertical anchor: "text", "margin", "page"
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_anchor: Option<String>,
    /// Horizontal alignment spec: "left", "center", "right"
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_align: Option<String>,
    /// Distance from surrounding text (points)
    #[serde(default)]
    pub left_from_text: f32,
    #[serde(default)]
    pub right_from_text: f32,
    #[serde(default)]
    pub top_from_text: f32,
    #[serde(default)]
    pub bottom_from_text: f32,
}

/// Table look conditional formatting flags (w:tblLook)
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Default)]
pub struct TableLook {
    /// Apply first row conditional style
    #[serde(default)]
    pub first_row: bool,
    /// Apply last row conditional style
    #[serde(default)]
    pub last_row: bool,
    /// Apply first column conditional style
    #[serde(default)]
    pub first_column: bool,
    /// Apply last column conditional style
    #[serde(default)]
    pub last_column: bool,
    /// Show horizontal banding (alternating row shading)
    #[serde(default)]
    pub banded_rows: bool,
    /// Show vertical banding (alternating column shading)
    #[serde(default)]
    pub banded_columns: bool,
    /// Row band size (number of rows per band, default 1)
    #[serde(default = "default_one")]
    pub row_band_size: u32,
    /// Column band size (number of columns per band, default 1)
    #[serde(default = "default_one")]
    pub col_band_size: u32,
}

/// Frame paragraph properties (w:framePr) — for drop caps and positioned text frames
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct FrameProperties {
    /// Drop cap type: "drop" (dropped into text), "margin" (in margin)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub drop_cap: Option<String>,
    /// Number of lines to drop (w:lines, default 1)
    #[serde(default)]
    pub lines: u32,
    /// Frame width in points (w:w)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<f32>,
    /// Frame height in points (w:h)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<f32>,
    /// Frame height rule: "auto", "atLeast", or "exact" (w:hRule)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height_rule: Option<String>,
    /// Horizontal anchor: "text", "margin", "page" (w:hAnchor)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub h_anchor: Option<String>,
    /// Vertical anchor: "text", "margin", "page" (w:vAnchor)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub v_anchor: Option<String>,
    /// Horizontal position in points (w:x)
    #[serde(default)]
    pub x: f32,
    /// Vertical position in points (w:y)
    #[serde(default)]
    pub y: f32,
    /// Horizontal space from text in points (w:hSpace)
    #[serde(default)]
    pub h_space: f32,
    /// Vertical space from text in points (w:vSpace)
    #[serde(default)]
    pub v_space: f32,
    /// Text wrapping: "auto", "around", "none", "notBeside", "through", "tight"
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap: Option<String>,
    /// Horizontal alignment: "left", "center", "right" (w:xAlign)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub x_align: Option<String>,
    /// Vertical alignment: "top", "center", "bottom" (w:yAlign)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub y_align: Option<String>,
}

/// Conditional formatting properties from w:tblStylePr
/// Applied to cells based on their position in the table (firstRow, band1Horz, etc.)
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct TableConditionalFormat {
    /// Cell shading/background color (hex)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shading: Option<String>,
    /// Cell borders
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub borders: Option<CellBorders>,
    /// Bold override for runs
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    /// Text color override (hex)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// Paragraph alignment override
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<Alignment>,
    /// Cell margins override
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub cell_margins: Option<CellMargins>,
}

/// A VML WordArt page watermark (see Page.watermark).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Watermark {
    pub text: String,
    /// Shape box size in points (pre-rotation), from the v:shape style attr.
    pub width: f32,
    pub height: f32,
    /// VML rotation in degrees CLOCKWISE (style `rotation:315` = 45° CCW visual).
    #[serde(default)]
    pub rotation: f32,
    /// Resolved fill color hex (e.g. "808080" for fillcolor="gray [...]").
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    /// v:textpath style font-family, if declared.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
pub struct PageSize {
    pub width: f32,
    pub height: f32,
}

impl Default for PageSize {
    fn default() -> Self {
        // A4 in points (210mm x 297mm)
        Self {
            width: 595.0,
            height: 842.0,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
pub struct Margin {
    pub top: f32,
    pub bottom: f32,
    pub left: f32,
    pub right: f32,
}

impl Default for Margin {
    fn default() -> Self {
        // Word default margins in points (1 inch = 72pt)
        Self {
            top: 72.0,
            bottom: 72.0,
            left: 72.0,
            right: 72.0,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct StyleSheet {
    pub styles: HashMap<String, StyleDefinition>,
    /// Default run properties from w:docDefaults/w:rPrDefault
    pub doc_default_run_style: Option<RunStyle>,
    /// Default paragraph properties from w:docDefaults/w:pPrDefault
    pub doc_default_para_style: Option<ParagraphStyle>,
    /// Default paragraph alignment from w:docDefaults/w:pPrDefault/w:jc
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub doc_default_alignment: Option<Alignment>,
    /// Table style borders: style_id -> TableStyle (with border info from tblBorders)
    #[serde(default, skip_serializing_if = "HashMap::is_empty")]
    pub table_styles: HashMap<String, TableStyle>,
    /// Table conditional formats: (style_id, condition_type) -> conditional properties
    /// condition_type: "firstRow", "lastRow", "firstCol", "lastCol", "band1Horz", "band2Horz", etc.
    #[serde(default, skip_serializing_if = "HashMap::is_empty")]
    pub table_conditional_formats: HashMap<String, HashMap<String, TableConditionalFormat>>,
    /// Default paragraph style ID (w:type="paragraph" w:default="1")
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_paragraph_style_id: Option<String>,
    /// Default TABLE style ID (w:type="table" w:default="1", normally
    /// "TableNormal"). Per ECMA-376 a table with no explicit `w:tblStyle`
    /// still inherits the default table style — which is where Word's
    /// "default" 108tw cellMar actually lives. See S871.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_table_style_id: Option<String>,
    /// Font table from word/fontTable.xml: font_name -> FontInfo
    #[serde(default, skip_serializing_if = "HashMap::is_empty")]
    pub font_table: HashMap<String, FontInfo>,
}

/// A named style definition with inheritance
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StyleDefinition {
    /// Style ID
    pub style_id: String,
    /// Parent style ID (w:basedOn)
    #[serde(default)]
    pub based_on: Option<String>,
    /// Paragraph properties defined in this style
    pub paragraph: ParagraphStyle,
    /// Paragraph alignment from this style (w:jc)
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<Alignment>,
    /// Whether inheritance has been resolved
    #[serde(skip, default)]
    pub resolved: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct DocumentMetadata {
    pub title: Option<String>,
    pub author: Option<String>,
    pub description: Option<String>,
}

/// Font information from fontTable.xml
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct FontInfo {
    /// PANOSE-1 classification (10 bytes as hex string, e.g. "020B0604020202020204")
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub panose1: Option<String>,
    /// Character set: "00" (ANSI), "80" (ShiftJIS), "02" (Symbol), etc.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub charset: Option<String>,
    /// Font family: "roman", "swiss", "modern", "decorative", "script", "auto"
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub family: Option<String>,
    /// Pitch: "fixed", "variable", "default"
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub pitch: Option<String>,
}

/// A footnote or endnote
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Footnote {
    /// Note number (1-based, matching the reference in the body)
    pub number: u32,
    /// Content paragraphs of the note
    pub blocks: Vec<Block>,
}
