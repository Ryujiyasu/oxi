// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Editing a model by position: insert text, delete a range, split and join
//! paragraphs, and format a range, keeping every run's formatting.
//!
//! A position is a paragraph and a character offset in it. Paragraphs are
//! numbered in document order through tables, content controls and other
//! containers, per story (body, a header or footer, a note, a comment).
//! Offsets count the way Word numbers its character positions, in UTF-16
//! code units (as JavaScript indexes strings), so a position here is a
//! position in Word: a field counts its begin
//! character, its code, its separator, its result and its end character; a
//! hyperlink counts as the HYPERLINK field Word makes of it; a picture,
//! a floating shape's anchor and a note reference count one each; text a
//! tracked deletion removed still counts; an inline content control's start
//! and end count one each; a complex HYPERLINK field, or one with a `\h`
//! switch, has one more just after its separator. [`paragraph_text`] spells out the characters.
//!
//! What each edit does follows Word (measured with Word COM, 2026-11-09):
//! typed text takes the formatting of the character before it, or the one
//! after it at the start of a paragraph, and text typed at a hyperlink's
//! edge goes outside the link; Enter gives both halves the paragraph's
//! formatting; deleting a range across paragraphs keeps the head of the
//! first and the tail of the last as two paragraphs (the first disappears
//! when the range starts at its beginning); joining paragraphs, and typing
//! over a range across paragraphs, keeps the first paragraph's formatting.

use serde::{Deserialize, Serialize};

use crate::content::{Block, CellItem, Hyperlink, Inline, Object, Paragraph, RowItem, Run, RunItem, Table};
use crate::package::DocxModel;
use crate::props::Props;
use crate::xml::{Element, Node};

/// Which part of the document a position is in.
#[derive(Debug, Clone, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum Story {
    #[default]
    Body,
    /// An index into `headers_footers`.
    HeaderFooter(usize),
    /// An index into `footnotes`.
    Footnote(usize),
    /// An index into `endnotes`.
    Endnote(usize),
    /// An index into `comments`.
    Comment(usize),
}

#[derive(Debug, Clone, PartialEq, Eq, Default, Serialize, Deserialize)]
pub struct Pos {
    #[serde(default)]
    pub story: Story,
    /// The paragraph's number in the story, in document order.
    pub para: usize,
    /// Characters before the position, counted as Word counts them.
    pub offset: usize,
}

impl Pos {
    pub fn body(para: usize, offset: usize) -> Pos {
        Pos { story: Story::Body, para, offset }
    }
}

/// Properties to set and remove. `set` holds property elements (`w:b`,
/// `w:sz w:val="28"`, `w:jc w:val="center"`); each replaces any property of
/// its name. `merge` holds elements whose attributes are set on the
/// property of that name, keeping its other attributes (`w:spacing
/// w:line="360"` keeps `w:before` and `w:after`). `remove` names properties
/// to take away.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct Patch {
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub set: Vec<Element>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub merge: Vec<Element>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub remove: Vec<String>,
    /// Take every property away first, so the bag holds just `set` and
    /// `merge` afterwards.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub clear: bool,
}

/// The kind of story `add_story` appends.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum StoryKind {
    /// A header (`w:hdr`) or footer (`w:ftr`).
    HeaderFooter { kind: String },
    Footnote,
    Endnote,
    Comment,
}

/// The document-wide definitions: styles, numbering, settings, theme, fonts
/// and the parts Word keeps for its own features. Each one given replaces
/// what the document had.
#[derive(Debug, Clone, PartialEq, Default, Serialize, Deserialize)]
pub struct DocumentParts {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub background: Option<Element>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub styles: Option<crate::docparts::Styles>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub styles_with_effects: Option<crate::docparts::Styles>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub numbering: Option<crate::docparts::Numbering>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub settings: Option<Object>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub web_settings: Option<Object>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_table: Option<Object>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub people: Option<Object>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme: Option<crate::content::Part>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub custom_xml: Vec<crate::content::Part>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub glossary: Option<crate::content::Part>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub app_data: Vec<crate::content::Part>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comments_extended: Option<Object>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comments_ids: Option<Object>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comments_extensible: Option<Object>,
    /// Keep a notes or comments part even with nothing in it.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub footnotes: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub endnotes: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub comments: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "op", rename_all = "snake_case")]
pub enum Edit {
    /// Type `text` at a position. Tabs become tab characters; a line or
    /// paragraph break is not text (use `split_paragraph`).
    InsertText { at: Pos, text: String },
    /// Delete what lies between two positions, as the Delete key does on a
    /// selection.
    Delete { from: Pos, to: Pos },
    /// Type `text` over a range: the range is deleted, paragraphs it spans
    /// are joined, and the text goes where the range began.
    Replace { from: Pos, to: Pos, text: String },
    /// Break a paragraph in two, as Enter does.
    SplitParagraph { at: Pos },
    /// Join a paragraph with the next one, as Delete at its end does. The
    /// next paragraph must follow it directly in the same container.
    JoinParagraphs {
        #[serde(default)]
        story: Story,
        para: usize,
    },
    /// Change run properties over a range. A range that runs past the end
    /// of a paragraph also formats that paragraph's mark, as Word does when
    /// a selection takes in the mark.
    FormatText { from: Pos, to: Pos, patch: Patch },
    /// Change paragraph properties of every paragraph the range touches.
    FormatParagraphs { from: Pos, to: Pos, patch: Patch },
    /// Put one thing at a position: a run (its properties and content), a
    /// marker (bookmark, comment range, permission range), an equation, or
    /// a run item written outside any run. It goes after anything of no
    /// width already at the position, inside the innermost container
    /// (hyperlink, field, content control) whose content takes the position
    /// in.
    InsertInline { at: Pos, inline: Inline },
    /// Put what lies between two positions of one paragraph into a
    /// container: a hyperlink, a simple field, a content control, a smart
    /// tag, custom XML, a tracked change, a bidirectional embedding.
    /// `wrapper` is the container with no content. Things of no width at
    /// either end go inside, except the first `outside` at the start.
    Wrap {
        from: Pos,
        to: Pos,
        wrapper: Inline,
        #[serde(default, skip_serializing_if = "is_zero")]
        outside: usize,
    },
    /// Insert a table before (or after) the block that holds a paragraph.
    /// The table's cells must be empty; each has one paragraph to edit
    /// afterwards.
    InsertTable {
        #[serde(default)]
        story: Story,
        para: usize,
        #[serde(default)]
        after: bool,
        table: Table,
    },
    /// Insert a block of no text before the block that holds a paragraph: a
    /// bookmark or comment range between paragraphs, an altChunk.
    InsertBlock {
        #[serde(default)]
        story: Story,
        para: usize,
        block: Block,
    },
    /// Put the `count` blocks just before the block that holds a paragraph
    /// into a container: a content control, custom XML, a tracked change.
    /// `wrapper` is the container with no content.
    WrapBlocks {
        #[serde(default)]
        story: Story,
        para: usize,
        count: usize,
        wrapper: Block,
    },
    /// Take out a paragraph with nothing in it.
    RemoveParagraph {
        #[serde(default)]
        story: Story,
        para: usize,
    },
    /// Set a paragraph's own attributes (w14:paraId, w14:textId, ...).
    SetParagraphAttrs {
        #[serde(default)]
        story: Story,
        para: usize,
        attrs: Vec<(String, String)>,
    },
    /// Set the document's last section (the body's own w:sectPr); `at`
    /// places it among the body's blocks when it is not last.
    SetSection {
        #[serde(default)]
        props: Option<Props>,
        #[serde(default)]
        at: Option<usize>,
    },
    /// Add an empty header, footer, footnote, endnote or comment at the end
    /// of its list; edit it as a story afterwards.
    AddStory {
        story: StoryKind,
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        attrs: Vec<(String, String)>,
    },
    /// Replace document-wide definitions.
    SetDocumentParts { parts: Box<DocumentParts> },
    /// Add a style, or change the one with its w:styleId. `with_effects`
    /// edits the Word 2010 copy of the styles (stylesWithEffects).
    DefineStyle {
        #[serde(default, skip_serializing_if = "std::ops::Not::not")]
        with_effects: bool,
        style: Box<crate::docparts::Style>,
    },
    /// Set the formatting every style starts from (docDefaults) and the
    /// built-in styles' visibility defaults (latentStyles). `declared` keeps
    /// an empty docDefaults.
    SetStyleDefaults {
        #[serde(default, skip_serializing_if = "std::ops::Not::not")]
        with_effects: bool,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        run: Option<Props>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        paragraph: Option<Props>,
        #[serde(default, skip_serializing_if = "std::ops::Not::not")]
        declared: bool,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        latent: Option<Element>,
    },
    /// Add a list definition (w:abstractNum), or change the one with its
    /// w:abstractNumId.
    DefineList { list: Box<crate::docparts::AbstractNum> },
    /// Add a list that paragraphs number with (w:num), or change the one
    /// with its w:numId.
    DefineListInstance { num: crate::docparts::Num },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EditError(pub String);

impl std::fmt::Display for EditError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "edit: {}", self.0)
    }
}

impl std::error::Error for EditError {}

fn is_zero(n: &usize) -> bool {
    *n == 0
}

fn err<T>(msg: impl Into<String>) -> Result<T, EditError> {
    Err(EditError(msg.into()))
}

// ---------------------------------------------------------------------------
// Characters
// ---------------------------------------------------------------------------

const FIELD_BEGIN: char = '\u{13}';
const FIELD_SEPARATE: char = '\u{14}';
const FIELD_END: char = '\u{15}';
/// What stands for a character whose content is not text: a picture, a
/// shape anchor, ruby, an equation.
pub const OBJECT: char = '\u{1}';
/// The start and end of an inline content control (Word gives each a
/// position; its Range.Text shows neither).
pub const CONTROL_START: char = '\u{1c}';
pub const CONTROL_END: char = '\u{1d}';
/// The character Word keeps after a hyperlink-type field's separator (see
/// `field_data_points`); not in the file, and not shown.
pub const FIELD_DATA: char = '\u{1b}';

fn attr<'a>(attrs: &'a [(String, String)], name: &str) -> Option<&'a str> {
    attrs.iter().find(|(k, _)| k == name).map(|(_, v)| v.as_str())
}

fn item_text(it: &RunItem, out: &mut String) {
    match it {
        RunItem::Text(s) | RunItem::DeletedText(s) | RunItem::InstrText(s) | RunItem::DeletedInstrText(s) => out.push_str(s),
        RunItem::Tab => out.push('\t'),
        RunItem::Break { attrs } => out.push(match attr(attrs, "w:type") {
            Some("page") => '\u{c}',
            Some("column") => '\u{e}',
            _ => '\u{b}',
        }),
        RunItem::FieldChar { attrs, data } => match attr(attrs, "w:fldCharType") {
            Some("begin") => {
                out.push(FIELD_BEGIN);
                // A form field's data (w:ffData) is a character of its own
                // in Word's field code.
                if data.is_some() {
                    out.push(OBJECT);
                }
            }
            Some("separate") => out.push(FIELD_SEPARATE),
            _ => out.push(FIELD_END),
        },
        // A comment is a range in Word; its reference takes no position
        // (measured: a comment Word made, live and reopened).
        RunItem::Mark { name, .. } if name == "w:commentReference" => {}
        // A note reference whose mark is the character after it
        // (customMarkFollows): the two are one position in Word.
        RunItem::Mark { name, attrs }
            if matches!(name.as_str(), "w:footnoteReference" | "w:endnoteReference")
                && matches!(attr(attrs, "w:customMarkFollows"), Some("1" | "true" | "on")) => {}
        RunItem::Mark { name, .. } => out.push(match name.as_str() {
            "w:cr" => '\u{b}',
            "w:noBreakHyphen" => '\u{1e}',
            "w:softHyphen" => '\u{1f}',
            "w:ptab" => '\t',
            "w:footnoteReference" | "w:endnoteReference" | "w:footnoteRef" | "w:endnoteRef" => '\u{2}',
            "w:separator" => '\u{3}',
            "w:continuationSeparator" => '\u{4}',
            "w:annotationRef" => '\u{5}',
            _ => OBJECT,
        }),
        RunItem::Object(o) => match embed_prog_id(&o.xml) {
            // An embedded OLE object is an EMBED field in Word: begin, code,
            // separator, the object, end.
            Some(prog) => {
                out.push(FIELD_BEGIN);
                out.push_str(&format!(" EMBED {prog} "));
                out.push(FIELD_SEPARATE);
                out.push(OBJECT);
                out.push(FIELD_END);
            }
            None => out.push(OBJECT),
        },
        // Ruby is an EQ field in Word, with no separator or result.
        RunItem::Ruby(r) => {
            out.push(FIELD_BEGIN);
            out.push_str(&ruby_code(r));
            out.push(FIELD_END);
        }
        RunItem::Props(_) => {}
    }
}

/// The code of the EQ field Word makes of ruby, with no leading space:
/// `EQ \* jc2 \* "Font:ＭＳ 明朝" \* hps10 \o\ad(\s\up 9(ほんこう),本項)`
/// (measured over every rubyAlign, an odd hpsRaise, and rt runs with and
/// without a font, 2026-11-09). The font is the ruby text's East Asian
/// font, else the base text's.
fn ruby_code(r: &crate::content::Ruby) -> String {
    let pr = |name: &str| r.props.child(name).and_then(|e| e.attr("w:val")).map(|s| s.to_string());
    let (jc, align) = match pr("w:rubyAlign").as_deref() {
        Some("center") | None => (0, ""),
        Some("distributeLetter") => (1, "\\ad"),
        Some("distributeSpace") => (2, "\\ad"),
        Some("left") => (3, "\\al"),
        Some("right") => (4, "\\ar"),
        Some("rightVertical") => (5, ""),
        Some(_) => (0, ""),
    };
    fn plain(v: &[Inline]) -> String {
        let mut s = String::new();
        for i in v {
            match i {
                Inline::Run(run) => {
                    for it in &run.content {
                        if let RunItem::Text(t) = it {
                            s.push_str(t);
                        }
                    }
                }
                _ => {
                    if let Some(c) = children(i) {
                        s.push_str(&plain(c));
                    }
                }
            }
        }
        s
    }
    fn east_asia(v: &[Inline]) -> Option<String> {
        v.iter().find_map(|i| match i {
            Inline::Run(run) => run.props.as_ref()?.get("w:rFonts")?.attr("w:eastAsia").map(|s| s.to_string()),
            _ => None,
        })
    }
    let font = east_asia(&r.text).or_else(|| east_asia(&r.base)).unwrap_or_default();
    let hps = pr("w:hps").unwrap_or_default();
    let raise = pr("w:hpsRaise").and_then(|v| v.parse::<i64>().ok()).unwrap_or(0) / 2;
    format!("EQ \\* jc{jc} \\* \"Font:{font}\" \\* hps{hps} \\o{align}(\\s\\up {raise}({}),{})", plain(&r.text), plain(&r.base))
}

/// The ProgID of an inline embedded OLE object (`w:object` holding
/// `o:OLEObject Type="Embed"`). A floating one (its VML shape positioned
/// absolutely) is a shape anchor in Word, one position, not a field.
fn embed_prog_id(e: &Element) -> Option<String> {
    if e.name != "w:object" {
        return None;
    }
    let mut all = Vec::new();
    e.walk(&mut all);
    let floating = all.iter().any(|c| {
        c.name.starts_with("v:") && c.attr("style").is_some_and(|s| s.replace(' ', "").contains("position:absolute"))
    });
    if floating {
        return None;
    }
    all.iter()
        .find(|c| c.name == "o:OLEObject" && c.attr("Type") == Some("Embed"))
        .map(|c| c.attr("ProgID").unwrap_or("").to_string())
}

/// The code of the HYPERLINK field Word makes of a hyperlink element,
/// with its leading space.
fn link_code(h: &Hyperlink) -> String {
    let mut code = String::from(" HYPERLINK");
    if let Some(t) = &h.target {
        // A backslash in a field code is written doubled (Word: a file:///
        // target's backslashes take two positions each).
        code.push_str(&format!(" \"{}\"", t.replace('\\', "\\\\")));
    }
    for (name, switch) in [("w:anchor", 'l'), ("w:tgtFrame", 't'), ("w:tooltip", 'o'), ("w:docLocation", 's')] {
        if let Some(v) = attr(&h.attrs, name) {
            code.push_str(&format!(" \\{switch} \"{v}\""));
        }
    }
    if !matches!(attr(&h.attrs, "w:history"), Some("1" | "true" | "on")) {
        code.push_str(" \\h");
    }
    code
}

/// Characters a container writes before its content.
fn prefix(i: &Inline) -> String {
    match i {
        Inline::Hyperlink(h) => format!("{FIELD_BEGIN}{}{FIELD_SEPARATE}", link_code(h)),
        Inline::SimpleField(f) => format!("{FIELD_BEGIN}{}{FIELD_SEPARATE}", attr(&f.attrs, "w:instr").unwrap_or("")),
        // An inline content control's start and end are a position each.
        Inline::Sdt(_) => CONTROL_START.to_string(),
        _ => String::new(),
    }
}

fn suffix(i: &Inline) -> &'static str {
    match i {
        Inline::Hyperlink(_) | Inline::SimpleField(_) => "\u{15}",
        Inline::Sdt(_) => "\u{1d}",
        _ => "",
    }
}

fn inline_text(i: &Inline, out: &mut String) {
    match i {
        Inline::Run(r) => r.content.iter().for_each(|it| item_text(it, out)),
        Inline::Marker(_) => {}
        Inline::Math(_) => out.push(OBJECT),
        Inline::Loose(it) => item_text(it, out),
        _ => {
            out.push_str(&prefix(i));
            for c in children(i).unwrap() {
                inline_text(c, out);
            }
            out.push_str(suffix(i));
        }
    }
}

fn children(i: &Inline) -> Option<&Vec<Inline>> {
    Some(match i {
        Inline::Hyperlink(h) => &h.content,
        Inline::SimpleField(f) => &f.content,
        Inline::Sdt(s) => &s.content,
        Inline::SmartTag(w) | Inline::CustomXml(w) | Inline::Bidi(w) => &w.content,
        Inline::Revision(r) => &r.content,
        _ => return None,
    })
}

fn children_mut(i: &mut Inline) -> Option<&mut Vec<Inline>> {
    Some(match i {
        Inline::Hyperlink(h) => &mut h.content,
        Inline::SimpleField(f) => &mut f.content,
        Inline::Sdt(s) => &mut s.content,
        Inline::SmartTag(w) | Inline::CustomXml(w) | Inline::Bidi(w) => &mut w.content,
        Inline::Revision(r) => &mut r.content,
        _ => return None,
    })
}

/// Word counts positions in UTF-16 code units: a character outside the
/// Basic Multilingual Plane (an emoji, a rare kanji) takes two. So do
/// JavaScript string indexes.
fn count(s: &str) -> usize {
    s.encode_utf16().count()
}

/// The byte index `units` UTF-16 units into `s`, if that is a character
/// boundary.
fn byte_at(s: &str, units: usize) -> Option<usize> {
    let mut u = 0;
    for (b, c) in s.char_indices() {
        if u == units {
            return Some(b);
        }
        if u > units {
            return None;
        }
        u += c.len_utf16();
    }
    (u == units).then_some(s.len())
}

fn item_len(it: &RunItem) -> usize {
    let mut s = String::new();
    item_text(it, &mut s);
    count(&s)
}

fn inline_len(i: &Inline) -> usize {
    let mut s = String::new();
    inline_text(i, &mut s);
    count(&s)
}

fn list_len(v: &[Inline]) -> usize {
    v.iter().map(inline_len).sum()
}

/// A paragraph's characters, one per Word character position (the
/// paragraph mark excluded).
pub fn paragraph_text(p: &Paragraph) -> String {
    let mut s = String::new();
    for i in &p.content {
        inline_text(i, &mut s);
    }
    s
}

/// The characters a paragraph shows, each with its position: a field shows
/// its result, not its begin, code, separator or end, and content-control
/// edges show nothing. `text` is a paragraph's characters as
/// [`paragraph_texts`] gives them. A renderer's text for the paragraph is
/// these characters, so this maps a place in what is drawn to a position.
pub fn shown_positions(text: &str) -> Vec<(char, usize)> {
    let mut out = Vec::new();
    // For each open field: whether its result has begun.
    let mut fields: Vec<bool> = Vec::new();
    let mut pos = 0;
    for c in text.chars() {
        let here = pos;
        pos += c.len_utf16();
        match c {
            FIELD_BEGIN => fields.push(false),
            FIELD_SEPARATE => {
                if let Some(f) = fields.last_mut() {
                    *f = true;
                }
            }
            FIELD_END => {
                fields.pop();
            }
            CONTROL_START | CONTROL_END | FIELD_DATA => {}
            _ if fields.iter().all(|in_result| *in_result) => out.push((c, here)),
            _ => {}
        }
    }
    out
}

/// Every paragraph of a story, in the order positions number them.
///
/// A story or table cell whose last block is not a paragraph (a body
/// ending in a table, a cell with no `w:p`) has one more, empty: Word ends
/// every story and cell with a paragraph mark, and writes `<w:p/>` there
/// when it saves. Editing the story writes them too.
pub fn paragraphs<'a>(model: &'a DocxModel, story: &Story) -> Result<Vec<&'a Paragraph>, EditError> {
    Ok(entries(model, story)?.into_iter().map(|e| e.para).collect())
}

/// Every paragraph of a story as positions count it: its characters, with
/// the start of any block-level content control it opens and the end of any
/// it closes (Word gives each a position, in the first and last paragraph
/// of the control).
pub fn paragraph_texts(model: &DocxModel, story: &Story) -> Result<Vec<String>, EditError> {
    Ok(entries(model, story)?
        .into_iter()
        .map(|e| {
            let mut s: String = std::iter::repeat_n(CONTROL_START, e.lead).collect();
            let text = paragraph_text(e.para);
            let mut from = 0;
            for p in field_data_points(e.para) {
                let b = byte_at(&text, p).unwrap_or(text.len());
                s.push_str(&text[from..b]);
                s.push(FIELD_DATA);
                from = b;
            }
            s.push_str(&text[from..]);
            s.extend(std::iter::repeat_n(CONTROL_END, e.trail));
            s
        })
        .collect())
}

/// Where Word has a character the file does not write: just after the
/// separator of a complex HYPERLINK field, or of a PAGEREF, REF or NOTEREF
/// field with a `\h` switch (not TOC's), Word keeps the link's data
/// as a character (measured: tools/metrics, 2026-11-09; a w:hyperlink
/// element has none). Offsets into the paragraph's own content, ascending.
fn field_data_points(p: &Paragraph) -> Vec<usize> {
    fn needs_data(code: &str) -> bool {
        let first = code.split_whitespace().next().unwrap_or("").to_ascii_uppercase();
        let linked = code.split_whitespace().any(|t| t.eq_ignore_ascii_case("\\h"));
        // A TOC field's \h makes its entries links, but the TOC field
        // itself has no data character.
        first == "HYPERLINK" || (linked && matches!(first.as_str(), "PAGEREF" | "REF" | "NOTEREF"))
    }
    fn walk(v: &[Inline], pos: &mut usize, stack: &mut Vec<(String, bool)>, out: &mut Vec<usize>) {
        for i in v {
            match i {
                Inline::Run(r) => {
                    for it in &r.content {
                        *pos += item_len(it);
                        match it {
                            RunItem::FieldChar { attrs, .. } => match attr(attrs, "w:fldCharType") {
                                Some("begin") => stack.push((String::new(), false)),
                                Some("separate") => {
                                    if let Some((code, in_result)) = stack.last_mut() {
                                        if !*in_result && needs_data(code) {
                                            out.push(*pos);
                                        }
                                        *in_result = true;
                                    }
                                }
                                _ => {
                                    stack.pop();
                                }
                            },
                            RunItem::InstrText(s) | RunItem::DeletedInstrText(s) => {
                                if let Some((code, false)) = stack.last_mut() {
                                    code.push_str(s);
                                }
                            }
                            _ => {}
                        }
                    }
                }
                _ => match children(i) {
                    Some(c) => {
                        *pos += count(&prefix(i));
                        walk(c, pos, stack, out);
                        *pos += count(suffix(i));
                    }
                    None => *pos += inline_len(i),
                },
            }
        }
    }
    let mut out = Vec::new();
    walk(&p.content, &mut 0, &mut Vec::new(), &mut out);
    out
}

/// A paragraph in a story's numbering, with the content-control starts
/// before its content and ends after it.
struct Entry<'a> {
    para: &'a Paragraph,
    lead: usize,
    trail: usize,
}

/// Whether a story ends without its own final paragraph mark: its last
/// block is not a paragraph, or (the body) its last paragraph carries a
/// section break, which ends a section before the document's last one.
fn story_lacks_final_paragraph(blocks: &[Block], story: &Story) -> bool {
    if lacks_final_paragraph(blocks) {
        return true;
    }
    *story == Story::Body
        && matches!(blocks.iter().rev().find(|b| !matches!(b, Block::Marker(_))),
            Some(Block::Paragraph(p)) if p.props.as_ref().is_some_and(|pp| pp.get("w:sectPr").is_some()))
}

fn entries<'a>(model: &'a DocxModel, story: &Story) -> Result<Vec<Entry<'a>>, EditError> {
    let blocks = story_blocks(model, story)?;
    let mut out = Vec::new();
    collect(blocks, &mut out);
    if story_lacks_final_paragraph(blocks, story) {
        out.push(Entry { para: empty_paragraph(), lead: 0, trail: 0 });
    }
    Ok(out)
}

/// The paragraph Word gives a story or cell that does not end in one.
fn empty_paragraph() -> &'static Paragraph {
    static EMPTY: std::sync::OnceLock<Paragraph> = std::sync::OnceLock::new();
    EMPTY.get_or_init(Paragraph::default)
}

/// Whether the last thing in a list of blocks is a paragraph mark: a
/// paragraph, or a content control (custom XML, revision) ending in one.
fn lacks_final_paragraph(blocks: &[Block]) -> bool {
    match blocks.iter().rev().find(|b| !matches!(b, Block::Marker(_))) {
        Some(Block::Paragraph(_)) => false,
        Some(Block::Sdt(s)) => lacks_final_paragraph(&s.content),
        Some(Block::CustomXml(w)) => lacks_final_paragraph(&w.content),
        Some(Block::Revision(r)) => lacks_final_paragraph(&r.content),
        _ => true,
    }
}

/// Mark the paragraphs collected since `start` as inside one content control.
fn control_edges(out: &mut [Entry<'_>], start: usize) {
    if out.len() > start {
        out[start].lead += 1;
        out.last_mut().unwrap().trail += 1;
    }
}

fn collect<'a>(blocks: &'a [Block], out: &mut Vec<Entry<'a>>) {
    for b in blocks {
        match b {
            Block::Paragraph(p) => out.push(Entry { para: p, lead: 0, trail: 0 }),
            Block::Table(t) => collect_rows(&t.rows, out),
            Block::Sdt(s) => {
                let start = out.len();
                collect(&s.content, out);
                control_edges(out, start);
            }
            Block::CustomXml(w) => collect(&w.content, out),
            Block::Revision(r) => collect(&r.content, out),
            Block::Marker(_) | Block::AltChunk(_) => {}
        }
    }
}

fn collect_rows<'a>(rows: &'a [RowItem], out: &mut Vec<Entry<'a>>) {
    for r in rows {
        match r {
            RowItem::Row(row) => collect_cells(&row.cells, out),
            RowItem::Sdt(s) => {
                let start = out.len();
                collect_rows(&s.content, out);
                control_edges(out, start);
            }
            RowItem::CustomXml(w) => collect_rows(&w.content, out),
            RowItem::Revision(v) => collect_rows(&v.content, out),
            RowItem::Marker(_) => {}
        }
    }
}

fn collect_cells<'a>(cells: &'a [CellItem], out: &mut Vec<Entry<'a>>) {
    for c in cells {
        match c {
            CellItem::Cell(cell) => {
                collect(&cell.content, out);
                if lacks_final_paragraph(&cell.content) {
                    out.push(Entry { para: empty_paragraph(), lead: 0, trail: 0 });
                }
            }
            CellItem::Sdt(s) => {
                let start = out.len();
                collect_cells(&s.content, out);
                control_edges(out, start);
            }
            CellItem::CustomXml(w) => collect_cells(&w.content, out),
            CellItem::Revision(v) => collect_cells(&v.content, out),
            CellItem::Marker(_) => {}
        }
    }
}

/// A position as an offset into the paragraph's own content: a position
/// before a content control's start, or after its end, is the nearest
/// position inside.
fn localize(model: &DocxModel, p: &Pos) -> Result<Pos, EditError> {
    let all = entries(model, &p.story)?;
    let e = all.get(p.para).ok_or_else(|| EditError(format!("no paragraph {}", p.para)))?;
    let len = list_len(&e.para.content);
    let data = field_data_points(e.para);
    let total = e.lead + len + data.len() + e.trail;
    if p.offset > total {
        return err(format!("offset {} past the paragraph's {total} positions", p.offset));
    }
    // Take away the control starts before the content and the field-data
    // characters before the position.
    let mut off = p.offset.saturating_sub(e.lead);
    for (k, at) in data.iter().enumerate() {
        if off > at + k {
            off -= 1;
        }
    }
    Ok(Pos { offset: off.min(len), ..p.clone() })
}

fn localize_edit(model: &DocxModel, edit: &Edit) -> Result<Edit, EditError> {
    let l = |p: &Pos| localize(model, p);
    Ok(match edit {
        Edit::InsertText { at, text } => Edit::InsertText { at: l(at)?, text: text.clone() },
        Edit::Delete { from, to } => Edit::Delete { from: l(from)?, to: l(to)? },
        Edit::Replace { from, to, text } => Edit::Replace { from: l(from)?, to: l(to)?, text: text.clone() },
        Edit::SplitParagraph { at } => Edit::SplitParagraph { at: l(at)? },
        Edit::JoinParagraphs { .. } => edit.clone(),
        Edit::FormatText { from, to, patch } => Edit::FormatText { from: l(from)?, to: l(to)?, patch: patch.clone() },
        Edit::FormatParagraphs { from, to, patch } => Edit::FormatParagraphs { from: l(from)?, to: l(to)?, patch: patch.clone() },
        Edit::InsertInline { at, inline } => Edit::InsertInline { at: l(at)?, inline: inline.clone() },
        Edit::Wrap { from, to, wrapper, outside } => Edit::Wrap { from: l(from)?, to: l(to)?, wrapper: wrapper.clone(), outside: *outside },
        _ => edit.clone(),
    })
}

// ---------------------------------------------------------------------------
// Finding paragraphs
// ---------------------------------------------------------------------------

fn story_blocks<'a>(model: &'a DocxModel, story: &Story) -> Result<&'a Vec<Block>, EditError> {
    let missing = || EditError(format!("no story {story:?}"));
    Ok(match story {
        Story::Body => &model.body,
        Story::HeaderFooter(i) => &model.headers_footers.get(*i).ok_or_else(missing)?.content,
        Story::Footnote(i) => &model.footnotes.as_ref().and_then(|v| v.get(*i)).ok_or_else(missing)?.content,
        Story::Endnote(i) => &model.endnotes.as_ref().and_then(|v| v.get(*i)).ok_or_else(missing)?.content,
        Story::Comment(i) => &model.comments.as_ref().and_then(|v| v.get(*i)).ok_or_else(missing)?.content,
    })
}

fn story_list_mut<'a>(model: &'a mut DocxModel, story: &Story) -> Result<&'a mut Vec<Block>, EditError> {
    let missing = || EditError(format!("no story {story:?}"));
    Ok(match story {
        Story::Body => &mut model.body,
        Story::HeaderFooter(i) => &mut model.headers_footers.get_mut(*i).ok_or_else(missing)?.content,
        Story::Footnote(i) => &mut model.footnotes.as_mut().and_then(|v| v.get_mut(*i)).ok_or_else(missing)?.content,
        Story::Endnote(i) => &mut model.endnotes.as_mut().and_then(|v| v.get_mut(*i)).ok_or_else(missing)?.content,
        Story::Comment(i) => &mut model.comments.as_mut().and_then(|v| v.get_mut(*i)).ok_or_else(missing)?.content,
    })
}

enum Reach {
    No,
    Real,
    /// The paragraph Word gives a story or cell that does not end in one.
    Virtual,
}

/// Find paragraph `n` of a list of blocks, pushing the indexes that lead to
/// it onto `path`. Paragraphs Word implies at the end of a cell count; with
/// `make`, the one sought is written out (Word writes `<w:p/>` there when it
/// saves). Only that one: a cell nobody edits stays as the file had it.
fn reach(blocks: &mut [Block], n: &mut usize, path: &mut Vec<usize>, make: bool) -> Reach {
    for (i, item) in blocks.iter_mut().enumerate() {
        path.push(i);
        let r = match item {
            Block::Paragraph(_) => {
                if *n == 0 {
                    Reach::Real
                } else {
                    *n -= 1;
                    Reach::No
                }
            }
            Block::Table(t) => reach_rows(&mut t.rows, n, path, make),
            Block::Sdt(s) => reach(&mut s.content, n, path, make),
            Block::CustomXml(w) => reach(&mut w.content, n, path, make),
            Block::Revision(r) => reach(&mut r.content, n, path, make),
            Block::Marker(_) | Block::AltChunk(_) => Reach::No,
        };
        if !matches!(r, Reach::No) {
            return r;
        }
        path.pop();
    }
    Reach::No
}

fn reach_rows(rows: &mut [RowItem], n: &mut usize, path: &mut Vec<usize>, make: bool) -> Reach {
    for (i, item) in rows.iter_mut().enumerate() {
        path.push(i);
        let r = match item {
            RowItem::Row(row) => reach_cells(&mut row.cells, n, path, make),
            RowItem::Sdt(s) => reach_rows(&mut s.content, n, path, make),
            RowItem::CustomXml(w) => reach_rows(&mut w.content, n, path, make),
            RowItem::Revision(v) => reach_rows(&mut v.content, n, path, make),
            RowItem::Marker(_) => Reach::No,
        };
        if !matches!(r, Reach::No) {
            return r;
        }
        path.pop();
    }
    Reach::No
}

fn reach_cells(cells: &mut [CellItem], n: &mut usize, path: &mut Vec<usize>, make: bool) -> Reach {
    for (i, item) in cells.iter_mut().enumerate() {
        path.push(i);
        let r = match item {
            CellItem::Cell(c) => match reach(&mut c.content, n, path, make) {
                Reach::No if lacks_final_paragraph(&c.content) => implied_slot(&mut c.content, n, path, make),
                r => r,
            },
            CellItem::Sdt(s) => reach_cells(&mut s.content, n, path, make),
            CellItem::CustomXml(w) => reach_cells(&mut w.content, n, path, make),
            CellItem::Revision(v) => reach_cells(&mut v.content, n, path, make),
            CellItem::Marker(_) => Reach::No,
        };
        if !matches!(r, Reach::No) {
            return r;
        }
        path.pop();
    }
    Reach::No
}

/// The implied final paragraph of `list`, if it is the one sought.
fn implied_slot(list: &mut Vec<Block>, n: &mut usize, path: &mut Vec<usize>, make: bool) -> Reach {
    if *n > 0 {
        *n -= 1;
        return Reach::No;
    }
    path.push(list.len());
    if make {
        list.push(Block::Paragraph(Paragraph::default()));
        Reach::Real
    } else {
        Reach::Virtual
    }
}

/// The path to paragraph `n` of a story: indexes through blocks, rows and
/// cells, the last one into the list of blocks that holds the paragraph
/// (for an implied paragraph not made, that list's length). Says whether
/// the paragraph is in the file.
fn locate(model: &mut DocxModel, story: &Story, n: usize, make: bool) -> Result<(Vec<usize>, bool), EditError> {
    let blocks = story_list_mut(model, story)?;
    let lacks = story_lacks_final_paragraph(blocks, story);
    let mut left = n;
    let mut path = Vec::new();
    let r = match reach(blocks, &mut left, &mut path, make) {
        Reach::No if lacks => implied_slot(blocks, &mut left, &mut path, make),
        r => r,
    };
    match r {
        Reach::No => err(format!("no paragraph {n}")),
        Reach::Real => Ok((path, true)),
        Reach::Virtual => Ok((path, false)),
    }
}

fn path_mut(model: &mut DocxModel, story: &Story, n: usize) -> Result<Vec<usize>, EditError> {
    Ok(locate(model, story, n, true)?.0)
}

/// The list of blocks the path's last index points into.
fn nav<'a>(v: &'a mut Vec<Block>, p: &[usize]) -> &'a mut Vec<Block> {
    if p.len() == 1 {
        return v;
    }
    match &mut v[p[0]] {
        Block::Table(t) => nav_rows(&mut t.rows, &p[1..]),
        Block::Sdt(s) => nav(&mut s.content, &p[1..]),
        Block::CustomXml(w) => nav(&mut w.content, &p[1..]),
        Block::Revision(r) => nav(&mut r.content, &p[1..]),
        _ => unreachable!("path through a block without blocks"),
    }
}

fn nav_rows<'a>(v: &'a mut [RowItem], p: &[usize]) -> &'a mut Vec<Block> {
    match &mut v[p[0]] {
        RowItem::Row(r) => nav_cells(&mut r.cells, &p[1..]),
        RowItem::Sdt(s) => nav_rows(&mut s.content, &p[1..]),
        RowItem::CustomXml(w) => nav_rows(&mut w.content, &p[1..]),
        RowItem::Revision(r) => nav_rows(&mut r.content, &p[1..]),
        RowItem::Marker(_) => unreachable!("path through a marker"),
    }
}

fn nav_cells<'a>(v: &'a mut [CellItem], p: &[usize]) -> &'a mut Vec<Block> {
    match &mut v[p[0]] {
        CellItem::Cell(c) => nav(&mut c.content, &p[1..]),
        CellItem::Sdt(s) => nav_cells(&mut s.content, &p[1..]),
        CellItem::CustomXml(w) => nav_cells(&mut w.content, &p[1..]),
        CellItem::Revision(r) => nav_cells(&mut r.content, &p[1..]),
        CellItem::Marker(_) => unreachable!("path through a marker"),
    }
}

fn para_at<'a>(blocks: &'a mut Vec<Block>, path: &[usize]) -> &'a mut Paragraph {
    let i = *path.last().unwrap();
    match &mut nav(blocks, path)[i] {
        Block::Paragraph(p) => p,
        _ => unreachable!("path to a block that is not a paragraph"),
    }
}

fn para_mut<'a>(model: &'a mut DocxModel, story: &Story, n: usize) -> Result<&'a mut Paragraph, EditError> {
    let path = path_mut(model, story, n)?;
    Ok(para_at(story_list_mut(model, story)?, &path))
}

// ---------------------------------------------------------------------------
// Runs
// ---------------------------------------------------------------------------

/// Cut a run at `k` characters; the run keeps the head, the tail is
/// returned with the same properties.
fn split_run(run: &mut Run, k: usize) -> Run {
    let mut head = Vec::new();
    let mut tail = Vec::new();
    let mut pos = 0;
    for it in std::mem::take(&mut run.content) {
        let l = item_len(&it);
        if pos > k || (pos == k && l > 0) {
            tail.push(it);
        } else if pos + l <= k {
            head.push(it);
        } else {
            let at = k - pos;
            let cut = |s: &str| {
                // check_offset has refused positions inside a character.
                let b = byte_at(s, at).unwrap_or(s.len());
                (s[..b].to_string(), s[b..].to_string())
            };
            let (a, b) = match &it {
                RunItem::Text(s) => {
                    let (a, b) = cut(s);
                    (RunItem::Text(a), RunItem::Text(b))
                }
                RunItem::DeletedText(s) => {
                    let (a, b) = cut(s);
                    (RunItem::DeletedText(a), RunItem::DeletedText(b))
                }
                RunItem::InstrText(s) => {
                    let (a, b) = cut(s);
                    (RunItem::InstrText(a), RunItem::InstrText(b))
                }
                RunItem::DeletedInstrText(s) => {
                    let (a, b) = cut(s);
                    (RunItem::DeletedInstrText(a), RunItem::DeletedInstrText(b))
                }
                // A form field's begin is two positions; a cut between them
                // keeps it whole, before the cut.
                other => (other.clone(), RunItem::Props(Props::new("w:rPr"))),
            };
            head.push(a);
            if !matches!(b, RunItem::Props(_)) {
                tail.push(b);
            }
        }
        pos += l;
    }
    run.content = head;
    Run { props: run.props.clone(), content: tail }
}

/// Make `off` a run boundary at every depth: no run is left straddling it.
fn split_at(v: &mut Vec<Inline>, off: usize) -> Result<(), EditError> {
    let mut pos = 0;
    let mut i = 0;
    while i < v.len() {
        let l = inline_len(&v[i]);
        if off <= pos {
            return Ok(());
        }
        if off < pos + l {
            let k = off - pos;
            match &mut v[i] {
                Inline::Run(r) => {
                    let tail = split_run(r, k);
                    v.insert(i + 1, Inline::Run(tail));
                }
                Inline::Math(_) | Inline::Loose(_) | Inline::Marker(_) => unreachable!("an atom cut inside"),
                other => {
                    let p = count(&prefix(other));
                    if k < p {
                        return err("position inside a field code");
                    }
                    let inner = children_mut(other).unwrap();
                    split_at(inner, k - p)?;
                }
            }
            return Ok(());
        }
        pos += l;
        i += 1;
    }
    Ok(())
}

/// Remove the characters in `[s, e)`, which are run boundaries.
fn delete_in(v: &mut Vec<Inline>, s: usize, e: usize) {
    let mut pos = 0;
    let mut out = Vec::with_capacity(v.len());
    for mut it in std::mem::take(v) {
        let l = inline_len(&it);
        let (a, b) = (pos, pos + l);
        pos = b;
        if l == 0 || b <= s || a >= e {
            out.push(it);
            continue;
        }
        if a >= s && b <= e {
            continue;
        }
        // Part of a container: delete inside its content.
        let p = count(&prefix(&it));
        if let Some(inner) = children_mut(&mut it) {
            let clen = list_len(inner);
            let lo = s.max(a + p).saturating_sub(a + p).min(clen);
            let hi = e.min(a + p + clen).saturating_sub(a + p);
            if hi > lo {
                delete_in(inner, lo, hi);
            }
            if inner.is_empty() && !matches!(it, Inline::Sdt(_)) {
                continue;
            }
        }
        out.push(it);
    }
    *v = out;
}

/// Run properties for text typed into a paragraph with no run to follow:
/// the paragraph mark's, without its revision marks.
fn mark_run_props(p: &Paragraph) -> Option<Props> {
    let mut el = p.props.as_ref()?.get("w:rPr")?.clone();
    el.children.retain(|n| !matches!(n, Node::Element(c) if matches!(c.name.as_str(), "w:ins" | "w:del" | "w:moveFrom" | "w:moveTo")));
    Props::read("w:r", &el).ok().filter(|p| !p.is_empty())
}

/// Items for typed text: text, with tabs as tab characters.
fn text_items(text: &str) -> Vec<RunItem> {
    let mut out = Vec::new();
    for (i, part) in text.split('\t').enumerate() {
        if i > 0 {
            out.push(RunItem::Tab);
        }
        if !part.is_empty() {
            out.push(RunItem::Text(part.to_string()));
        }
    }
    out
}

fn push_text(run: &mut Run, text: &str, at_end: bool) {
    let mut items = text_items(text);
    if at_end {
        // After the last character, before trailing zero-width items.
        let at = run.content.iter().rposition(|it| item_len(it) > 0).map(|i| i + 1).unwrap_or(0);
        if at > 0 {
            if let (RunItem::Text(prev), Some(RunItem::Text(first))) = (&mut run.content[at - 1], items.first()) {
                prev.push_str(first);
                items.remove(0);
            }
        }
        run.content.splice(at..at, items);
    } else {
        let at = run.content.iter().position(|it| item_len(it) > 0).unwrap_or(run.content.len());
        if at < run.content.len() {
            if let (RunItem::Text(next), Some(RunItem::Text(last))) = (&mut run.content[at], items.last()) {
                next.insert_str(0, last);
                items.pop();
            }
        }
        run.content.splice(at..at, items);
    }
}

/// Containers typing passes through at their edges: they only mark text.
fn transparent(i: &Inline) -> bool {
    matches!(i, Inline::SmartTag(_) | Inline::CustomXml(_) | Inline::Bidi(_))
}

fn edge_run_props(i: &Inline, last: bool) -> Option<Option<Props>> {
    match i {
        Inline::Run(r) => Some(r.props.clone()),
        _ => {
            let c = children(i)?;
            let mut it: Box<dyn Iterator<Item = &Inline>> = if last { Box::new(c.iter().rev()) } else { Box::new(c.iter()) };
            it.find_map(|x| edge_run_props(x, last))
        }
    }
}

/// Type `text` at `off` (a run boundary). False when the list holds no
/// character to take formatting from.
fn insert_in(v: &mut Vec<Inline>, off: usize, text: &str, fallback: &Option<Props>) -> Result<bool, EditError> {
    let mut pos = 0;
    let mut prev = None;
    let mut next = None;
    for (i, it) in v.iter().enumerate() {
        let l = inline_len(it);
        if l == 0 {
            continue;
        }
        if pos < off && off < pos + l {
            let p = count(&prefix(it));
            if off - pos < p {
                return err("position inside a field code");
            }
            let k = off - pos - p;
            let props = edge_run_props(it, false).flatten();
            let inner = children_mut(&mut v[i]).unwrap();
            if !insert_in(inner, k, text, fallback)? {
                inner.insert(0, Inline::Run(Run { props, content: text_items(text) }));
            }
            return Ok(true);
        }
        if pos + l == off {
            prev = Some(i);
        }
        if pos == off && next.is_none() {
            next = Some(i);
        }
        pos += l;
    }
    if let Some(i) = prev {
        match &mut v[i] {
            Inline::Run(r) => {
                push_text(r, text, true);
                return Ok(true);
            }
            w if transparent(w) => {
                let inner = children_mut(w).unwrap();
                let end = list_len(inner);
                if insert_in(inner, end, text, fallback)? {
                    return Ok(true);
                }
            }
            _ => {}
        }
    }
    if let Some(i) = next {
        match &mut v[i] {
            Inline::Run(r) => {
                push_text(r, text, false);
                return Ok(true);
            }
            w if transparent(w) => {
                let inner = children_mut(w).unwrap();
                if insert_in(inner, 0, text, fallback)? {
                    return Ok(true);
                }
            }
            _ => {}
        }
    }
    // Beside a hyperlink, field or object: a run of its own, formatted like
    // the neighbouring text but without its character style (text typed
    // after a hyperlink is not link-styled).
    let near = prev.and_then(|i| edge_run_props(&v[i], true)).or_else(|| next.and_then(|i| edge_run_props(&v[i], false)));
    let props = match near {
        Some(p) => p.map(|mut p| {
            p.remove("w:rStyle");
            p
        }).filter(|p| !p.is_empty()),
        None if v.iter().any(|i| inline_len(i) > 0) => None,
        None => {
            if fallback.is_none() && v.is_empty() {
                return Ok(false);
            }
            fallback.clone()
        }
    };
    let at = match (prev, next) {
        (Some(i), _) => i + 1,
        (None, Some(i)) => i,
        (None, None) => v.len(),
    };
    v.insert(at, Inline::Run(Run { props, content: text_items(text) }));
    Ok(true)
}

/// Cut a list at `off` (a run boundary); the list keeps the head and the
/// tail is returned. A container the cut passes through is split in two.
fn partition(v: &mut Vec<Inline>, off: usize) -> Result<Vec<Inline>, EditError> {
    let mut pos = 0;
    let mut head = Vec::new();
    let mut tail = Vec::new();
    for mut it in std::mem::take(v) {
        let l = inline_len(&it);
        let (a, b) = (pos, pos + l);
        pos = b;
        if l == 0 && a == off {
            // A zero-width marker at the cut: ends stay, starts go on.
            let ends = matches!(&it, Inline::Marker(m) if m.name.ends_with("End"));
            if ends { head.push(it) } else { tail.push(it) }
        } else if b <= off {
            head.push(it);
        } else if a >= off {
            tail.push(it);
        } else {
            let p = count(&prefix(&it));
            if off - a < p {
                return err("position inside a field code");
            }
            let mut second = it.clone();
            let inner = children_mut(&mut it).unwrap();
            let rest = partition(inner, off - a - p)?;
            let keep_head = !inner.is_empty();
            *children_mut(&mut second).unwrap() = rest;
            if keep_head {
                head.push(it);
            }
            if !children(&second).unwrap().is_empty() {
                tail.push(second);
            }
        }
    }
    *v = head;
    Ok(tail)
}

fn check_text(text: &str) -> Result<(), EditError> {
    if let Some(c) = text.chars().find(|c| c.is_control() && *c != '\t') {
        return err(format!("text holds control character U+{:04X}; break paragraphs with split_paragraph", c as u32));
    }
    Ok(())
}

fn check_offset(p: &Paragraph, off: usize) -> Result<(), EditError> {
    let len = list_len(&p.content);
    if off > len {
        return err(format!("offset {off} past the paragraph's {len} positions"));
    }
    if byte_at(&paragraph_text(p), off).is_none() {
        return err(format!("offset {off} falls inside a character (between the halves of a surrogate pair)"));
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Edits
// ---------------------------------------------------------------------------

fn insert_text(model: &mut DocxModel, at: &Pos, text: &str) -> Result<(), EditError> {
    check_text(text)?;
    if text.is_empty() {
        return Ok(());
    }
    let p = para_mut(model, &at.story, at.para)?;
    check_offset(p, at.offset)?;
    split_at(&mut p.content, at.offset)?;
    let fallback = mark_run_props(p);
    if !insert_in(&mut p.content, at.offset, text, &fallback)? {
        p.content.push(Inline::Run(Run { props: fallback, content: text_items(text) }));
    }
    Ok(())
}

/// Widen `[s, e)` to whole fields: a range that takes in a field's begin or
/// end character deletes the whole field, as Word does (a hyperlink is a
/// field too).
fn whole_fields(text: &str, mut s: usize, mut e: usize) -> (usize, usize) {
    let mut spans = Vec::new();
    let mut open = Vec::new();
    let mut i = 0;
    for c in text.chars() {
        let here = i;
        i += c.len_utf16();
        let i = here;
        match c {
            FIELD_BEGIN => open.push(i),
            FIELD_END => {
                if let Some(b) = open.pop() {
                    spans.push((b, i));
                }
            }
            _ => {}
        }
    }
    loop {
        let mut changed = false;
        for &(b, end) in &spans {
            let takes_begin = s <= b && b < e;
            let takes_end = s <= end && end < e;
            if (takes_begin || takes_end) && (b < s || end >= e) {
                s = s.min(b);
                e = e.max(end + 1);
                changed = true;
            }
        }
        if !changed {
            return (s, e);
        }
    }
}

fn delete_within(p: &mut Paragraph, s: usize, e: usize) -> Result<(), EditError> {
    check_offset(p, e)?;
    if e <= s {
        return Ok(());
    }
    let (s, e) = whole_fields(&paragraph_text(p), s, e);
    split_at(&mut p.content, e)?;
    split_at(&mut p.content, s)?;
    delete_in(&mut p.content, s, e);
    Ok(())
}

/// Two paragraphs of a story, checked to be in one list of blocks, `a`
/// before `b`. Returns the shared path prefix and the two indexes.
fn siblings(model: &mut DocxModel, story: &Story, a: usize, b: usize) -> Result<(Vec<usize>, usize, usize), EditError> {
    let pa = path_mut(model, story, a)?;
    let pb = path_mut(model, story, b)?;
    if pa.len() != pb.len() || pa[..pa.len() - 1] != pb[..pb.len() - 1] {
        return err(format!("paragraphs {a} and {b} are in different containers"));
    }
    let (ia, ib) = (*pa.last().unwrap(), *pb.last().unwrap());
    Ok((pa, ia, ib))
}

/// The paragraph properties a joined paragraph keeps: the first one's,
/// with the section break (if any) of the second, since a paragraph
/// carrying a section break ends that section.
fn joined_props(first: Option<Props>, second: &Option<Props>) -> Option<Props> {
    let mut props = first;
    if let Some(p) = props.as_mut() {
        p.remove("w:sectPr");
    }
    if let Some(sect) = second.as_ref().and_then(|p| p.get("w:sectPr")) {
        let p = props.get_or_insert_with(|| Props::new("w:pPr"));
        let _ = p.set(sect.clone());
    }
    props.filter(|p| !p.is_empty())
}

/// Join paragraph `ia` of the list with the one at `ib`, dropping every
/// block between them.
fn join_blocks(list: &mut Vec<Block>, ia: usize, ib: usize) {
    let removed: Vec<Block> = list.drain(ia + 1..=ib).collect();
    let Some(Block::Paragraph(second)) = removed.into_iter().last() else { unreachable!() };
    let Block::Paragraph(first) = &mut list[ia] else { unreachable!() };
    first.content.extend(second.content);
    first.props = joined_props(first.props.take(), &second.props);
}

fn delete(model: &mut DocxModel, from: &Pos, to: &Pos, join: bool) -> Result<(), EditError> {
    if from.story != to.story {
        return err("a range across stories");
    }
    let (from, to) = if (to.para, to.offset) < (from.para, from.offset) { (to, from) } else { (from, to) };
    if from.para == to.para {
        return delete_within(para_mut(model, &from.story, from.para)?, from.offset, to.offset);
    }
    let (path, ia, ib) = siblings(model, &from.story, from.para, to.para)?;
    let blocks = story_list_mut(model, &from.story)?;
    let first_len = {
        let p = para_at(blocks, &path);
        check_offset(p, from.offset)?;
        list_len(&p.content)
    };
    let mut pb = path.clone();
    *pb.last_mut().unwrap() = ib;
    delete_within(para_at(blocks, &pb), 0, to.offset)?;
    let list = nav(blocks, &path);
    if join || to.offset == 0 && from.offset > 0 {
        // Typing over the range, or a range that ends where the last
        // paragraph begins (it takes in only the marks): one paragraph.
        let Block::Paragraph(first) = &mut list[ia] else { unreachable!() };
        delete_within(first, from.offset, first_len)?;
        join_blocks(list, ia, ib);
    } else if from.offset == 0 {
        // From the start of the first paragraph: it goes, with everything
        // up to the tail of the last.
        list.drain(ia..ib);
    } else {
        let Block::Paragraph(first) = &mut list[ia] else { unreachable!() };
        delete_within(first, from.offset, first_len)?;
        list.drain(ia + 1..ib);
    }
    Ok(())
}

fn split_paragraph(model: &mut DocxModel, at: &Pos) -> Result<(), EditError> {
    let path = path_mut(model, &at.story, at.para)?;
    let blocks = story_list_mut(model, &at.story)?;
    let p = para_at(blocks, &path);
    check_offset(p, at.offset)?;
    split_at(&mut p.content, at.offset)?;
    let tail = partition(&mut p.content, at.offset)?;
    let mut second = Paragraph {
        props: p.props.clone(),
        content: tail,
        // Paragraph ids name one paragraph; the new one has none yet.
        attrs: p.attrs.iter().filter(|(k, _)| k != "w14:paraId" && k != "w14:textId").cloned().collect(),
    };
    // The section break stays with the paragraph that ends the section.
    if let Some(props) = p.props.as_mut() {
        if props.get("w:sectPr").is_some() {
            props.remove("w:sectPr");
            if props.is_empty() {
                p.props = None;
            }
        }
    }
    if second.props.as_ref().is_some_and(|p| p.is_empty()) {
        second.props = None;
    }
    let i = *path.last().unwrap();
    nav(blocks, &path).insert(i + 1, Block::Paragraph(second));
    Ok(())
}

fn join_paragraphs(model: &mut DocxModel, story: &Story, para: usize) -> Result<(), EditError> {
    let (path, ia, ib) = siblings(model, story, para, para + 1)?;
    let blocks = story_list_mut(model, story)?;
    if ib != ia + 1 {
        return err(format!("paragraph {} does not follow paragraph {para} directly", para + 1));
    }
    join_blocks(nav(blocks, &path), ia, ib);
    Ok(())
}

/// Apply a patch to a property element, then check and order it as `parent`'s bag.
fn patched(current: Option<&Props>, name: &str, parent: &str, patch: &Patch) -> Result<Option<Props>, EditError> {
    let mut el = current.map(|p| p.element().clone()).unwrap_or_else(|| Element::new(name));
    if patch.clear {
        el.children.clear();
        el.attrs.clear();
    }
    let drop: Vec<&str> = patch.remove.iter().map(|s| s.as_str()).chain(patch.set.iter().map(|e| e.name.as_str())).collect();
    el.children.retain(|n| !matches!(n, Node::Element(c) if drop.contains(&c.name.as_str())));
    el.children.extend(patch.set.iter().cloned().map(Node::Element));
    for m in &patch.merge {
        if let Some(c) = el.elements_mut().find(|c| c.name == m.name) {
            for (k, v) in &m.attrs {
                c.set_attr(k.clone(), v.clone());
            }
            continue;
        }
        el.children.push(Node::Element(m.clone()));
    }
    let props = Props::read(parent, &el).map_err(|u| EditError(u.to_string()))?;
    Ok((!props.is_empty()).then_some(props))
}

fn format_runs(v: &mut [Inline], s: usize, e: usize, patch: &Patch) -> Result<(), EditError> {
    let mut pos = 0;
    for it in v.iter_mut() {
        let l = inline_len(it);
        let (a, b) = (pos, pos + l);
        pos = b;
        if l == 0 || b <= s || a >= e {
            continue;
        }
        match it {
            Inline::Run(r) => r.props = patched(r.props.as_ref(), "w:rPr", "w:r", patch)?,
            Inline::Math(_) | Inline::Loose(_) | Inline::Marker(_) => {}
            other => {
                let p = count(&prefix(other));
                let inner = children_mut(other).unwrap();
                let clen = list_len(inner);
                let lo = s.max(a + p).saturating_sub(a + p).min(clen);
                let hi = e.min(a + p + clen).saturating_sub(a + p);
                if hi > lo {
                    format_runs(inner, lo, hi, patch)?;
                }
            }
        }
    }
    Ok(())
}

fn format_text(model: &mut DocxModel, from: &Pos, to: &Pos, patch: &Patch) -> Result<(), EditError> {
    if from.story != to.story {
        return err("a range across stories");
    }
    let (from, to) = if (to.para, to.offset) < (from.para, from.offset) { (to, from) } else { (from, to) };
    for n in from.para..=to.para {
        let p = para_mut(model, &from.story, n)?;
        let len = list_len(&p.content);
        let s = if n == from.para { from.offset } else { 0 };
        let e = if n == to.para { to.offset } else { len };
        check_offset(p, s)?;
        check_offset(p, e)?;
        if e > s {
            split_at(&mut p.content, e)?;
            split_at(&mut p.content, s)?;
            format_runs(&mut p.content, s, e, patch)?;
        }
        if n < to.para {
            // The range runs on past this paragraph's mark.
            let mark = p.props.as_ref().and_then(|pp| pp.get("w:rPr")).cloned();
            let current = mark.map(|m| Props::read("w:pPr", &m)).transpose().map_err(|u| EditError(u.to_string()))?;
            let new = patched(current.as_ref(), "w:rPr", "w:pPr", patch)?;
            let props = p.props.get_or_insert_with(|| Props::new("w:pPr"));
            match new {
                Some(m) => props.set(m.into_element()).map_err(|u| EditError(u.to_string()))?,
                None => props.remove("w:rPr"),
            }
            if props.is_empty() {
                p.props = None;
            }
        }
    }
    Ok(())
}

fn format_paragraphs(model: &mut DocxModel, from: &Pos, to: &Pos, patch: &Patch) -> Result<(), EditError> {
    if from.story != to.story {
        return err("a range across stories");
    }
    let (a, b) = if to.para < from.para { (to.para, from.para) } else { (from.para, to.para) };
    for n in a..=b {
        let p = para_mut(model, &from.story, n)?;
        p.props = patched(p.props.as_ref(), "w:pPr", "w:p", patch)?;
    }
    Ok(())
}

/// Type over a range. The typed text takes the formatting of the range's
/// first character (Word), so it is first typed just after that character,
/// which is then deleted with the rest of the range.
fn replace(model: &mut DocxModel, from: &Pos, to: &Pos, text: &str) -> Result<(), EditError> {
    check_text(text)?;
    let (from, to) = if (to.para, to.offset) < (from.para, from.offset) { (to, from) } else { (from, to) };
    let n = count(text);
    let first_len = list_len(&para_mut(model, &from.story, from.para)?.content);
    let empty = (from.para, from.offset) == (to.para, to.offset);
    if n > 0 && !empty && from.offset < first_len {
        let mut trial = model.clone();
        let after = Pos { offset: from.offset + 1, ..from.clone() };
        let shifted = |p: &Pos, by: usize| Pos { offset: p.offset + by, ..p.clone() };
        let done = insert_text(&mut trial, &after, text)
            .and_then(|_| delete(&mut trial, from, &shifted(from, 1), false))
            .and_then(|_| {
                let end = if to.para == from.para { shifted(to, n - 1) } else { to.clone() };
                delete(&mut trial, &shifted(from, n), &end, true)
            });
        if done.is_ok() {
            *model = trial;
            return Ok(());
        }
    }
    delete(model, from, to, true)?;
    insert_text(model, from, text)
}

/// Put a leaf at `off` (a run boundary): after the zero-width things there,
/// inside the innermost container whose content takes the position in.
fn place_in(v: &mut Vec<Inline>, off: usize, new: Inline) -> Result<(), EditError> {
    let mut pos = 0;
    for i in 0..v.len() {
        let l = inline_len(&v[i]);
        if pos < off && off < pos + l {
            let p = count(&prefix(&v[i]));
            if off - pos < p {
                return err("position inside a field code");
            }
            let inner = children_mut(&mut v[i]).ok_or_else(|| EditError("position inside an object".into()))?;
            return place_in(inner, off - pos - p, new);
        }
        if pos > off || (pos == off && l > 0) {
            v.insert(i, new);
            return Ok(());
        }
        pos += l;
    }
    v.push(new);
    Ok(())
}

fn insert_inline(model: &mut DocxModel, at: &Pos, inline: &Inline) -> Result<(), EditError> {
    if !matches!(inline, Inline::Run(_) | Inline::Marker(_) | Inline::Math(_) | Inline::Loose(_)) {
        return err("insert_inline takes a run, a marker, an equation or a loose item; put content into containers with wrap");
    }
    let p = para_mut(model, &at.story, at.para)?;
    check_offset(p, at.offset)?;
    split_at(&mut p.content, at.offset)?;
    place_in(&mut p.content, at.offset, inline.clone())
}

/// Wrap `[s, e)` (run boundaries) of a list in `shell`, inside the innermost
/// container whose content holds the range.
fn wrap_in(v: &mut Vec<Inline>, s: usize, e: usize, mut shell: Inline, outside: usize) -> Result<(), EditError> {
    let mut pos = 0;
    let mut spans = Vec::with_capacity(v.len());
    for it in v.iter() {
        let l = inline_len(it);
        spans.push((pos, pos + l));
        pos += l;
    }
    for (i, &(a, b)) in spans.iter().enumerate() {
        if children(&v[i]).is_none() {
            continue;
        }
        let p = count(&prefix(&v[i]));
        let q = count(suffix(&v[i]));
        let (cs, ce) = (a + p, b - q);
        if cs <= s && e <= ce && (a < s || e < b) {
            return wrap_in(children_mut(&mut v[i]).unwrap(), s - cs, e - cs, shell, outside);
        }
    }
    if spans.iter().any(|&(a, b)| (a < s && s < b) || (a < e && e < b)) {
        return err("the range takes in part of a container");
    }
    // The first `outside` things of no width at the start stay out.
    let mut left = outside;
    let inside: Vec<bool> = spans
        .iter()
        .map(|&(a, b)| {
            if a == s && b == s && left > 0 {
                left -= 1;
                return false;
            }
            a >= s && b <= e
        })
        .collect();
    if left > 0 {
        return err(format!("fewer than {outside} things of no width at the start of the range"));
    }
    let first = inside
        .iter()
        .position(|x| *x)
        .unwrap_or_else(|| spans.iter().position(|&(a, b)| a >= e && b > a).unwrap_or(v.len()));
    let mut kept = Vec::with_capacity(v.len());
    let inner = children_mut(&mut shell).unwrap();
    for (it, &x) in std::mem::take(v).into_iter().zip(&inside) {
        if x {
            inner.push(it);
        } else {
            kept.push(it);
        }
    }
    let at = first.min(kept.len());
    kept.insert(at, shell);
    *v = kept;
    Ok(())
}

fn wrap(model: &mut DocxModel, from: &Pos, to: &Pos, wrapper: &Inline, outside: usize) -> Result<(), EditError> {
    if from.story != to.story || from.para != to.para {
        return err("wrap takes a range in one paragraph");
    }
    if !children(wrapper).is_some_and(|c| c.is_empty()) {
        return err("wrap takes a container with no content");
    }
    let (s, e) = (from.offset.min(to.offset), from.offset.max(to.offset));
    let p = para_mut(model, &from.story, from.para)?;
    check_offset(p, s)?;
    check_offset(p, e)?;
    split_at(&mut p.content, e)?;
    split_at(&mut p.content, s)?;
    wrap_in(&mut p.content, s, e, wrapper.clone(), outside)
}

/// Positions an inline takes up, as Word counts them.
pub fn inline_width(i: &Inline) -> usize {
    inline_len(i)
}

fn cells_empty(rows: &[RowItem]) -> bool {
    fn cells(c: &[CellItem]) -> bool {
        c.iter().all(|c| match c {
            CellItem::Cell(c) => c.content.is_empty(),
            CellItem::Sdt(s) => cells(&s.content),
            CellItem::CustomXml(w) => cells(&w.content),
            CellItem::Revision(r) => cells(&r.content),
            CellItem::Marker(_) => true,
        })
    }
    rows.iter().all(|r| match r {
        RowItem::Row(r) => cells(&r.cells),
        RowItem::Sdt(s) => cells_empty(&s.content),
        RowItem::CustomXml(w) => cells_empty(&w.content),
        RowItem::Revision(r) => cells_empty(&r.content),
        RowItem::Marker(_) => true,
    })
}

fn insert_table(model: &mut DocxModel, story: &Story, para: usize, after: bool, table: &Table) -> Result<(), EditError> {
    if !cells_empty(&table.rows) {
        return err("insert_table takes a table whose cells are empty; edit their paragraphs afterwards");
    }
    // A table goes after a paragraph in the file, never an implied one.
    let (path, _) = locate(model, story, para, after)?;
    let list = nav(story_list_mut(model, story)?, &path);
    let i = (*path.last().unwrap() + after as usize).min(list.len());
    list.insert(i, Block::Table(table.clone()));
    Ok(())
}

fn insert_block(model: &mut DocxModel, story: &Story, para: usize, block: &Block) -> Result<(), EditError> {
    if !matches!(block, Block::Marker(_) | Block::AltChunk(_)) {
        return err("insert_block takes a marker or an altChunk; paragraphs come from split_paragraph, tables from insert_table");
    }
    let (path, _) = locate(model, story, para, false)?;
    let list = nav(story_list_mut(model, story)?, &path);
    let i = (*path.last().unwrap()).min(list.len());
    list.insert(i, block.clone());
    Ok(())
}

fn wrap_blocks(model: &mut DocxModel, story: &Story, para: usize, count: usize, wrapper: &Block) -> Result<(), EditError> {
    let mut shell = wrapper.clone();
    let content = match &mut shell {
        Block::Sdt(s) => &mut s.content,
        Block::CustomXml(w) => &mut w.content,
        Block::Revision(r) => &mut r.content,
        _ => return err("wrap_blocks takes a content control, custom XML or a tracked change"),
    };
    if !content.is_empty() {
        return err("wrap_blocks takes a container with no content");
    }
    let (path, _) = locate(model, story, para, false)?;
    let list = nav(story_list_mut(model, story)?, &path);
    let end = (*path.last().unwrap()).min(list.len());
    if count == 0 || count > end {
        return err(format!("{count} blocks do not come before paragraph {para} in its list"));
    }
    content.extend(list.drain(end - count..end));
    list.insert(end - count, shell);
    Ok(())
}

fn remove_paragraph(model: &mut DocxModel, story: &Story, para: usize) -> Result<(), EditError> {
    let (path, real) = locate(model, story, para, false)?;
    if !real {
        return err(format!("paragraph {para} is not in the file"));
    }
    let list = nav(story_list_mut(model, story)?, &path);
    let i = *path.last().unwrap();
    match &list[i] {
        Block::Paragraph(p) if p.content.is_empty() => {
            list.remove(i);
            Ok(())
        }
        _ => err(format!("paragraph {para} is not empty")),
    }
}

/// Whether a paragraph is one Word implies at the end of a story or cell
/// that does not end in one (see [`paragraphs`]), not in the file.
pub fn implied(model: &DocxModel, story: &Story, para: usize) -> Result<bool, EditError> {
    let all = entries(model, story)?;
    let e = all.get(para).ok_or_else(|| EditError(format!("no paragraph {para}")))?;
    Ok(std::ptr::eq(e.para, empty_paragraph()))
}

fn add_story(model: &mut DocxModel, story: &StoryKind, attrs: &[(String, String)]) {
    use crate::package::{HeaderFooter, Note};
    let note = || Note { attrs: attrs.to_vec(), content: Vec::new() };
    match story {
        StoryKind::HeaderFooter { kind } => model.headers_footers.push(HeaderFooter { kind: kind.clone(), content: Vec::new() }),
        StoryKind::Footnote => model.footnotes.get_or_insert_with(Vec::new).push(note()),
        StoryKind::Endnote => model.endnotes.get_or_insert_with(Vec::new).push(note()),
        StoryKind::Comment => model.comments.get_or_insert_with(Vec::new).push(note()),
    }
}

fn styles_mut(model: &mut DocxModel, with_effects: bool) -> &mut crate::docparts::Styles {
    let st = if with_effects { &mut model.styles_with_effects } else { &mut model.styles };
    st.get_or_insert_with(Default::default)
}

/// Replace the definition with the same id, or add it at the end.
fn define<T>(list: &mut Vec<T>, new: T, id: impl Fn(&T) -> Option<String>) {
    let key = id(&new);
    match key.and_then(|k| list.iter().position(|x| id(x).as_deref() == Some(k.as_str()))) {
        Some(i) => list[i] = new,
        None => list.push(new),
    }
}

fn set_document_parts(model: &mut DocxModel, parts: &DocumentParts) {
    let p = parts.clone();
    macro_rules! take {
        ($($f:ident),*) => { $( if p.$f.is_some() { model.$f = p.$f; } )* };
    }
    take!(background, styles, styles_with_effects, numbering, settings, web_settings, font_table, people, theme, glossary, comments_extended, comments_ids, comments_extensible);
    if !p.custom_xml.is_empty() {
        model.custom_xml = p.custom_xml;
    }
    if !p.app_data.is_empty() {
        model.app_data = p.app_data;
    }
    if p.footnotes {
        model.footnotes.get_or_insert_with(Vec::new);
    }
    if p.endnotes {
        model.endnotes.get_or_insert_with(Vec::new);
    }
    if p.comments {
        model.comments.get_or_insert_with(Vec::new);
    }
}

/// Where a paragraph's content ends, as a position: what an edit at the end
/// of the paragraph gives as `offset`.
pub fn paragraph_end(model: &DocxModel, story: &Story, para: usize) -> Result<usize, EditError> {
    let all = entries(model, story)?;
    let e = all.get(para).ok_or_else(|| EditError(format!("no paragraph {para}")))?;
    Ok(e.lead + list_len(&e.para.content) + field_data_points(e.para).len())
}

/// Join neighbouring runs with the same properties, as Word does when it
/// saves, so cutting a run to edit it leaves no seam.
fn coalesce(v: &mut Vec<Inline>) {
    let mut out: Vec<Inline> = Vec::with_capacity(v.len());
    for mut it in std::mem::take(v) {
        if let Some(inner) = children_mut(&mut it) {
            coalesce(inner);
        }
        if let (Some(Inline::Run(prev)), Inline::Run(r)) = (out.last_mut(), &it) {
            if prev.props == r.props && !r.content.iter().any(|c| matches!(c, RunItem::Props(_))) {
                let Inline::Run(r) = it else { unreachable!() };
                for c in r.content {
                    match (prev.content.last_mut(), c) {
                        (Some(RunItem::Text(a)), RunItem::Text(b)) => a.push_str(&b),
                        (_, c) => prev.content.push(c),
                    }
                }
                continue;
            }
        }
        out.push(it);
    }
    *v = out;
}

/// The paragraphs an edit touches, to tidy afterwards.
fn touched(edit: &Edit) -> (Story, usize, usize) {
    let span = |a: &Pos, b: &Pos| (a.story.clone(), a.para.min(b.para), a.para.max(b.para) + 1);
    match edit {
        Edit::InsertText { at, .. } => (at.story.clone(), at.para, at.para),
        Edit::SplitParagraph { at } => (at.story.clone(), at.para, at.para + 1),
        Edit::JoinParagraphs { story, para } => (story.clone(), *para, *para),
        Edit::Delete { from, to } | Edit::Replace { from, to, .. } => {
            let (s, a, _) = span(from, to);
            (s, a, a + 1)
        }
        Edit::FormatText { from, to, .. } => span(from, to),
        Edit::InsertInline { at, .. } => (at.story.clone(), at.para, at.para),
        Edit::Wrap { from, .. } => (from.story.clone(), from.para, from.para),
        _ => (Story::Body, 1, 0),
    }
}

/// Apply one edit.
pub fn apply(model: &mut DocxModel, edit: &Edit) -> Result<(), EditError> {
    let edit = &localize_edit(model, edit)?;
    apply_one(model, edit)?;
    let (story, a, b) = touched(edit);
    for n in a..=b {
        // Only paragraphs in the file: tidying does not write out an implied one.
        if let Ok((path, true)) = locate(model, &story, n, false) {
            if let Ok(blocks) = story_list_mut(model, &story) {
                coalesce(&mut para_at(blocks, &path).content);
            }
        }
    }
    Ok(())
}

fn apply_one(model: &mut DocxModel, edit: &Edit) -> Result<(), EditError> {
    match edit {
        Edit::InsertText { at, text } => insert_text(model, at, text),
        Edit::Delete { from, to } => delete(model, from, to, false),
        Edit::Replace { from, to, text } => replace(model, from, to, text),
        Edit::SplitParagraph { at } => split_paragraph(model, at),
        Edit::JoinParagraphs { story, para } => join_paragraphs(model, story, *para),
        Edit::FormatText { from, to, patch } => format_text(model, from, to, patch),
        Edit::FormatParagraphs { from, to, patch } => format_paragraphs(model, from, to, patch),
        Edit::InsertInline { at, inline } => insert_inline(model, at, inline),
        Edit::Wrap { from, to, wrapper, outside } => wrap(model, from, to, wrapper, *outside),
        Edit::InsertTable { story, para, after, table } => insert_table(model, story, *para, *after, table),
        Edit::InsertBlock { story, para, block } => insert_block(model, story, *para, block),
        Edit::WrapBlocks { story, para, count, wrapper } => wrap_blocks(model, story, *para, *count, wrapper),
        Edit::RemoveParagraph { story, para } => remove_paragraph(model, story, *para),
        Edit::SetParagraphAttrs { story, para, attrs } => {
            para_mut(model, story, *para)?.attrs = attrs.clone();
            Ok(())
        }
        Edit::SetSection { props, at } => {
            model.section = props.clone();
            model.section_at = *at;
            Ok(())
        }
        Edit::AddStory { story, attrs } => {
            add_story(model, story, attrs);
            Ok(())
        }
        Edit::SetDocumentParts { parts } => {
            set_document_parts(model, parts);
            Ok(())
        }
        Edit::DefineStyle { with_effects, style } => {
            let list = &mut styles_mut(model, *with_effects).styles;
            define(list, (**style).clone(), |x| attr(&x.attrs, "w:styleId").map(str::to_string));
            Ok(())
        }
        Edit::SetStyleDefaults { with_effects, run, paragraph, declared, latent } => {
            let st = styles_mut(model, *with_effects);
            st.default_run = run.clone();
            st.default_paragraph = paragraph.clone();
            st.has_defaults = *declared;
            st.latent = latent.clone();
            Ok(())
        }
        Edit::DefineList { list } => {
            let n = model.numbering.get_or_insert_with(Default::default);
            define(&mut n.abstracts, (**list).clone(), |x| attr(&x.attrs, "w:abstractNumId").map(str::to_string));
            Ok(())
        }
        Edit::DefineListInstance { num } => {
            let n = model.numbering.get_or_insert_with(Default::default);
            define(&mut n.nums, num.clone(), |x| attr(&x.attrs, "w:numId").map(str::to_string));
            Ok(())
        }
    }
}

/// Read edits from JSON. Edits that carry deeply nested content (a table
/// in a table in a table ...) go deeper than serde_json's default limit of
/// 128 levels, so the limit is lifted, as for a model.
pub fn from_json(json: &str) -> Result<Vec<Edit>, serde_json::Error> {
    let mut de = serde_json::Deserializer::from_str(json);
    de.disable_recursion_limit();
    let edits = Vec::<Edit>::deserialize(&mut de)?;
    de.end()?;
    Ok(edits)
}

/// Apply edits in order. On an error the model is left as it was.
pub fn apply_all(model: &mut DocxModel, edits: &[Edit]) -> Result<(), EditError> {
    let mut work = model.clone();
    for (i, e) in edits.iter().enumerate() {
        apply(&mut work, e).map_err(|x| EditError(format!("edit {i}: {}", x.0)))?;
    }
    *model = work;
    Ok(())
}
