// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Mail merge: fill a main document's merge fields from a table of records,
//! the way Word's "merge to a new document" does it.
//!
//! The main document is edited as XML, in place: only the fields are
//! replaced, so everything else (styles, tables, pictures, sections) is the
//! template's own. What a merge writes was measured on Word's own merges
//! (letters, to a new document):
//!
//! - every record is a copy of the body ending in its own section, which
//!   restarts page numbering and carries its own copy of any header or
//!   footer that holds merge fields; after the last record an empty
//!   paragraph closes a continuous section;
//! - a field becomes plain text with `noProof` added; it is set in the
//!   formatting of the field code's first run, or of the field's last result
//!   when the code says `\* MERGEFORMAT`;
//! - bookmarks are not copied, and the main document's data-source link is
//!   not written to the merged document;
//! - field names match the data's column names without regard to case, and
//!   a space in a column name is read as `_`;
//! - `IF` compares text with case and with `*`/`?` wildcards after `=`/`<>`,
//!   and numbers as numbers; `SKIPIF` drops the record, `NEXT`/`NEXTIF` move
//!   on to the next record inside the same copy;
//! - `\b`/`\f` add their text only around a value that is not empty; `\#` and
//!   `\@` format numbers and dates as Word's pictures do; `\*` changes case;
//! - blank-line suppression (Word's default) drops a paragraph that held a
//!   merge field and is left with nothing but spaces.

use std::collections::{BTreeSet, HashMap};
use std::io::{Cursor, Read, Write};

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// The records to merge: column names and one row of text per record.
#[derive(Debug, Clone, Default, PartialEq, serde::Serialize, serde::Deserialize)]
pub struct MergeData {
    pub columns: Vec<String>,
    pub rows: Vec<Vec<String>>,
}

/// How to merge.
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
#[serde(default)]
pub struct MergeOptions {
    /// Drop a paragraph that held a merge field and is left blank (Word's
    /// "do not print blank lines when data fields are empty", on by default).
    pub suppress_blank_lines: bool,
    /// First record to merge, counting from 1.
    pub first_record: Option<usize>,
    /// Last record to merge, counting from 1.
    pub last_record: Option<usize>,
}

impl Default for MergeOptions {
    fn default() -> Self {
        MergeOptions { suppress_blank_lines: true, first_record: None, last_record: None }
    }
}

/// One merged document of [`mail_merge_each`].
#[derive(Debug, Clone)]
pub struct MergedDocument {
    /// The record it was made from, counting from 1 (the first one, when
    /// `NEXT` fields pulled in more).
    pub record: usize,
    pub bytes: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum MergeError {
    Package(String),
    NoDocument,
}

impl std::fmt::Display for MergeError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            MergeError::Package(e) => write!(f, "cannot read the document: {e}"),
            MergeError::NoDocument => write!(f, "the package has no main document"),
        }
    }
}

impl std::error::Error for MergeError {}

/// What the main document asks of the data.
#[derive(Debug, Clone, Default, PartialEq, serde::Serialize)]
pub struct MergeFieldReport {
    /// The field names the main document uses, in order of first use.
    pub fields: Vec<String>,
}

// ---------------------------------------------------------------------------
// Data
// ---------------------------------------------------------------------------

/// Read comma- or tab-separated text: the first line names the columns.
/// Quotes follow RFC 4180 (a quoted value may hold the separator, a doubled
/// quote and line breaks). The separator is the one the first line uses more.
pub fn parse_delimited(text: &str) -> MergeData {
    let text = text.strip_prefix('\u{feff}').unwrap_or(text);
    let first_line: &str = text.split('\n').next().unwrap_or("");
    let sep = if first_line.matches('\t').count() > first_line.matches(',').count() { '\t' } else { ',' };
    let mut rows: Vec<Vec<String>> = Vec::new();
    let mut row: Vec<String> = Vec::new();
    let mut cell = String::new();
    let mut quoted = false;
    let mut any = false;
    let mut chars = text.chars().peekable();
    while let Some(c) = chars.next() {
        if quoted {
            if c == '"' {
                if chars.peek() == Some(&'"') {
                    chars.next();
                    cell.push('"');
                } else {
                    quoted = false;
                }
            } else {
                cell.push(c);
            }
            continue;
        }
        match c {
            '"' if cell.is_empty() => {
                quoted = true;
                any = true;
            }
            c if c == sep => {
                row.push(std::mem::take(&mut cell));
                any = true;
            }
            '\r' => {}
            '\n' => {
                row.push(std::mem::take(&mut cell));
                rows.push(std::mem::take(&mut row));
                any = false;
            }
            c => {
                cell.push(c);
                any = true;
            }
        }
    }
    if any || !cell.is_empty() || !row.is_empty() {
        row.push(cell);
        rows.push(row);
    }
    let cell_text = |s: &String| s.replace("\r\n", "\n");
    let mut rows = rows.into_iter();
    let columns: Vec<String> = rows.next().unwrap_or_default().iter().map(|c| c.trim().to_string()).collect();
    let rows = rows
        .filter(|r| r.iter().any(|c| !c.is_empty()))
        .map(|r| r.iter().map(cell_text).collect())
        .collect();
    MergeData { columns, rows }
}

/// Word reads a column name with its spaces as `_`, and matches a field's
/// name to it without regard to case.
fn name_key(name: &str) -> String {
    name.trim().chars().map(|c| if c.is_whitespace() { '_' } else { c }).flat_map(char::to_lowercase).collect()
}

// ---------------------------------------------------------------------------
// A light XML element scanner: element names and byte ranges, nothing else.
// ---------------------------------------------------------------------------

#[derive(Debug)]
struct El {
    /// Qualified name.
    name: String,
    start: usize,
    /// Just past the start tag's `>`.
    open_end: usize,
    /// Just past the end tag (or the start tag, when it closes itself).
    end: usize,
    /// Where the end tag starts (== `end` for a self-closing element).
    close_start: usize,
    parent: Option<usize>,
    children: Vec<usize>,
}

fn find_from(s: &str, from: usize, pat: &str) -> Option<usize> {
    s.get(from..)?.find(pat).map(|i| i + from)
}

fn scan(xml: &str) -> Vec<El> {
    let b = xml.as_bytes();
    let mut els: Vec<El> = Vec::new();
    let mut stack: Vec<usize> = Vec::new();
    let mut i = 0;
    while let Some(lt) = find_from(xml, i, "<") {
        let rest = &xml[lt..];
        if rest.starts_with("<?") {
            i = find_from(xml, lt, "?>").map_or(xml.len(), |e| e + 2);
            continue;
        }
        if rest.starts_with("<!--") {
            i = find_from(xml, lt, "-->").map_or(xml.len(), |e| e + 3);
            continue;
        }
        if rest.starts_with("<![CDATA[") {
            i = find_from(xml, lt, "]]>").map_or(xml.len(), |e| e + 3);
            continue;
        }
        if rest.starts_with("<!") {
            i = find_from(xml, lt, ">").map_or(xml.len(), |e| e + 1);
            continue;
        }
        // Find the tag's end, stepping over quoted attribute values.
        let mut j = lt + 1;
        let mut quote = 0u8;
        while j < b.len() {
            let c = b[j];
            if quote != 0 {
                if c == quote {
                    quote = 0;
                }
            } else if c == b'"' || c == b'\'' {
                quote = c;
            } else if c == b'>' {
                break;
            }
            j += 1;
        }
        if j >= b.len() {
            break;
        }
        let gt = j;
        let after = gt + 1;
        if rest.starts_with("</") {
            if let Some(open) = stack.pop() {
                els[open].close_start = lt;
                els[open].end = after;
            }
        } else {
            let name_end = xml[lt + 1..gt]
                .find(|c: char| c.is_whitespace() || c == '/' || c == '>')
                .map_or(gt, |k| lt + 1 + k);
            let self_closing = gt > lt && b[gt - 1] == b'/';
            let id = els.len();
            let parent = stack.last().copied();
            els.push(El {
                name: xml[lt + 1..name_end].to_string(),
                start: lt,
                open_end: after,
                end: after,
                close_start: after,
                parent,
                children: Vec::new(),
            });
            if let Some(p) = parent {
                els[p].children.push(id);
            }
            if !self_closing {
                stack.push(id);
            }
        }
        i = after;
    }
    els
}

fn local(name: &str) -> &str {
    name.rsplit(':').next().unwrap_or(name)
}

fn prefix(name: &str) -> &str {
    name.split_once(':').map_or("", |(p, _)| p)
}

/// The prefix the part binds WordprocessingML's namespace to.
fn w_prefix(xml: &str) -> String {
    let needle = format!("=\"{W_NS}\"");
    if let Some(at) = xml.find(&needle) {
        let head = &xml[..at];
        if let Some(x) = head.rfind("xmlns") {
            let decl = &head[x + 5..];
            return decl.strip_prefix(':').unwrap_or("").to_string();
        }
    }
    "w".to_string()
}

fn attr<'a>(xml: &'a str, el: &El, name: &str) -> Option<&'a str> {
    let tag = &xml[el.start..el.open_end];
    let mut from = 0;
    while let Some(k) = find_from(tag, from, name) {
        let before = tag.as_bytes().get(k.wrapping_sub(1)).copied().unwrap_or(b' ');
        let rest = &tag[k + name.len()..];
        let rest_trim = rest.trim_start();
        if before.is_ascii_whitespace() && rest_trim.starts_with('=') {
            let v = rest_trim[1..].trim_start();
            let q = v.chars().next()?;
            if q == '"' || q == '\'' {
                let body = &v[1..];
                return body.find(q).map(|e| &body[..e]);
            }
        }
        from = k + name.len();
    }
    None
}

fn unescape(s: &str) -> String {
    if !s.contains('&') {
        return s.to_string();
    }
    let mut out = String::with_capacity(s.len());
    let mut rest = s;
    while let Some(a) = rest.find('&') {
        out.push_str(&rest[..a]);
        let tail = &rest[a..];
        if let Some(semi) = tail.find(';') {
            let ent = &tail[1..semi];
            let ch = match ent {
                "amp" => Some('&'),
                "lt" => Some('<'),
                "gt" => Some('>'),
                "quot" => Some('"'),
                "apos" => Some('\''),
                _ if ent.starts_with("#x") || ent.starts_with("#X") => {
                    u32::from_str_radix(&ent[2..], 16).ok().and_then(char::from_u32)
                }
                _ if ent.starts_with('#') => ent[1..].parse::<u32>().ok().and_then(char::from_u32),
                _ => None,
            };
            if let Some(c) = ch {
                out.push(c);
                rest = &tail[semi + 1..];
                continue;
            }
        }
        out.push('&');
        rest = &tail[1..];
    }
    out.push_str(rest);
    out
}

fn escape(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;")
}

fn inner<'a>(xml: &'a str, el: &El) -> &'a str {
    if el.close_start <= el.open_end {
        ""
    } else {
        &xml[el.open_end..el.close_start]
    }
}

// ---------------------------------------------------------------------------
// Fields
// ---------------------------------------------------------------------------

/// One piece of a paragraph's run content, in document order.
#[derive(Debug, Clone, Copy)]
enum Item {
    /// A child of a run (`run`, `child` are element ids).
    RunChild { run: usize, child: usize },
    /// A `w:fldSimple`, taken whole.
    Simple(usize),
}

#[derive(Debug)]
enum Piece {
    Text(String),
    Field(usize),
}

#[derive(Debug)]
struct Field {
    /// Item index of the begin `fldChar` (or the `fldSimple`).
    begin: usize,
    end: Option<usize>,
    separated: bool,
    code: Vec<Piece>,
    code_rpr: String,
    result_rpr: Option<String>,
    parent: Option<usize>,
}

struct Paragraph {
    items: Vec<Item>,
    fields: Vec<Field>,
}

fn rpr_of(xml: &str, els: &[El], run: usize, w: &str) -> String {
    els[run]
        .children
        .iter()
        .find(|&&c| els[c].name == format!("{w}:rPr").as_str() || (w.is_empty() && els[c].name == "rPr"))
        .map(|&c| xml[els[c].start..els[c].end].to_string())
        .unwrap_or_default()
}

fn qn(w: &str, l: &str) -> String {
    if w.is_empty() {
        l.to_string()
    } else {
        format!("{w}:{l}")
    }
}

fn is_w(els: &[El], id: usize, w: &str, l: &str) -> bool {
    let n = &els[id].name;
    local(n) == l && prefix(n) == w
}

fn collect_items(els: &[El], node: usize, w: &str, out: &mut Vec<Item>) {
    for &c in &els[node].children {
        if is_w(els, c, w, "r") {
            for &k in &els[c].children {
                if !is_w(els, k, w, "rPr") {
                    out.push(Item::RunChild { run: c, child: k });
                }
            }
        } else if is_w(els, c, w, "fldSimple") {
            out.push(Item::Simple(c));
        } else if is_w(els, c, w, "p") || is_w(els, c, w, "del") || is_w(els, c, w, "moveFrom")
            || is_w(els, c, w, "pPr") || local(&els[c].name) == "Fallback"
        {
            // A nested paragraph is merged on its own; deleted text does not count.
        } else {
            collect_items(els, c, w, out);
        }
    }
}

fn paragraph_fields(xml: &str, els: &[El], p: usize, w: &str) -> Paragraph {
    let mut items = Vec::new();
    collect_items(els, p, w, &mut items);
    let mut fields: Vec<Field> = Vec::new();
    let mut stack: Vec<usize> = Vec::new();
    for (i, item) in items.iter().enumerate() {
        match *item {
            Item::Simple(s) => {
                let instr = unescape(attr(xml, &els[s], &qn(w, "instr")).unwrap_or(""));
                let first_run = els[s].children.iter().copied().find(|&c| is_w(els, c, w, "r"));
                let id = fields.len();
                fields.push(Field {
                    begin: i,
                    end: Some(i),
                    separated: true,
                    code: vec![Piece::Text(instr)],
                    code_rpr: String::new(),
                    result_rpr: first_run.map(|r| rpr_of(xml, els, r, w)),
                    parent: stack.last().copied(),
                });
                if let Some(&top) = stack.last() {
                    if !fields[top].separated {
                        fields[top].code.push(Piece::Field(id));
                    }
                }
            }
            Item::RunChild { run, child } => {
                let l = local(&els[child].name);
                if prefix(&els[child].name) != w {
                    // Not WordprocessingML content (mc:AlternateContent and kin): content.
                } else if l == "fldChar" {
                    match attr(xml, &els[child], &qn(w, "fldCharType")).unwrap_or("") {
                        "begin" => {
                            let id = fields.len();
                            let parent = stack.last().copied();
                            fields.push(Field {
                                begin: i,
                                end: None,
                                separated: false,
                                code: Vec::new(),
                                code_rpr: rpr_of(xml, els, run, w),
                                result_rpr: None,
                                parent,
                            });
                            if let Some(pid) = parent {
                                if !fields[pid].separated {
                                    fields[pid].code.push(Piece::Field(id));
                                }
                            }
                            stack.push(id);
                        }
                        "separate" => {
                            if let Some(&top) = stack.last() {
                                fields[top].separated = true;
                            }
                        }
                        "end" => {
                            if let Some(top) = stack.pop() {
                                fields[top].end = Some(i);
                            }
                        }
                        _ => {}
                    }
                    continue;
                } else if l == "instrText" {
                    if let Some(&top) = stack.last() {
                        if !fields[top].separated {
                            fields[top].code.push(Piece::Text(unescape(inner(xml, &els[child]))));
                        }
                    }
                    continue;
                } else if l == "t" {
                    if let Some(&top) = stack.last() {
                        if fields[top].separated && fields[top].result_rpr.is_none() {
                            fields[top].result_rpr = Some(rpr_of(xml, els, run, w));
                        }
                    }
                }
            }
        }
    }
    Paragraph { items, fields }
}

/// A token of a field code: quoted text, or a bare word. A nested field's
/// result is part of the token it sits in and never splits it.
#[derive(Debug, Clone, PartialEq)]
struct Tok {
    text: String,
    quoted: bool,
}

fn tokenize(pieces: &[(bool, String)]) -> Vec<Tok> {
    // pieces: (is_value, text); a value is atomic.
    let mut toks = Vec::new();
    let mut cur = String::new();
    let mut in_tok = false;
    let mut quoted = false;
    let flush = |toks: &mut Vec<Tok>, cur: &mut String, in_tok: &mut bool, quoted: bool| {
        if *in_tok {
            toks.push(Tok { text: std::mem::take(cur), quoted });
            *in_tok = false;
        }
    };
    for (is_value, text) in pieces {
        if *is_value {
            cur.push_str(text);
            in_tok = true;
            continue;
        }
        let mut chars = text.chars().peekable();
        while let Some(c) = chars.next() {
            if quoted {
                if c == '"' {
                    flush(&mut toks, &mut cur, &mut in_tok, true);
                    in_tok = false;
                    quoted = false;
                } else {
                    cur.push(c);
                }
                continue;
            }
            match c {
                '"' => {
                    flush(&mut toks, &mut cur, &mut in_tok, false);
                    quoted = true;
                    in_tok = true;
                }
                c if c.is_whitespace() => flush(&mut toks, &mut cur, &mut in_tok, false),
                '=' | '<' | '>' => {
                    flush(&mut toks, &mut cur, &mut in_tok, false);
                    let mut op = c.to_string();
                    if let Some(&n) = chars.peek() {
                        if (c == '<' && (n == '>' || n == '=')) || (c == '>' && n == '=') {
                            op.push(n);
                            chars.next();
                        }
                    }
                    toks.push(Tok { text: op, quoted: false });
                }
                '\\' if !in_tok => {
                    // A switch: `\b`, `\*` … one character after the backslash.
                    cur.push(c);
                    in_tok = true;
                    if let Some(n) = chars.next() {
                        cur.push(n);
                    }
                    flush(&mut toks, &mut cur, &mut in_tok, false);
                }
                c => {
                    cur.push(c);
                    in_tok = true;
                }
            }
        }
    }
    flush(&mut toks, &mut cur, &mut in_tok, quoted);
    toks
}

#[derive(Debug, Default)]
struct Switches {
    before: Option<String>,
    after: Option<String>,
    number: Option<String>,
    date: Option<String>,
    cases: Vec<String>,
    mergeformat: bool,
}

/// Split a field's tokens into its arguments and its switches.
fn split_switches(toks: &[Tok]) -> (Vec<Tok>, Switches) {
    let mut args = Vec::new();
    let mut sw = Switches::default();
    let mut i = 0;
    while i < toks.len() {
        let t = &toks[i];
        if !t.quoted && t.text.starts_with('\\') && t.text.len() >= 2 {
            let key = t.text[1..].to_ascii_lowercase();
            let val = toks.get(i + 1).map(|v| v.text.clone());
            match key.as_str() {
                "b" | "f" | "#" | "@" | "*" => {
                    i += 1;
                    let v = val.unwrap_or_default();
                    match key.as_str() {
                        "b" => sw.before = Some(v),
                        "f" => sw.after = Some(v),
                        "#" => sw.number = Some(v),
                        "@" => sw.date = Some(v),
                        _ => {
                            let lower = v.to_ascii_lowercase();
                            if lower == "mergeformat" {
                                sw.mergeformat = true;
                            } else if lower != "charformat" {
                                sw.cases.push(lower);
                            }
                        }
                    }
                }
                _ => {}
            }
        } else {
            args.push(t.clone());
        }
        i += 1;
    }
    (args, sw)
}

struct Ctx<'a> {
    data: &'a MergeData,
    keys: HashMap<String, usize>,
    /// Current record (index into `data.rows`).
    rec: usize,
    /// Sequence number of the merged copy, from 1.
    seq: usize,
    skip: bool,
    missing: BTreeSet<String>,
}

impl<'a> Ctx<'a> {
    fn new(data: &'a MergeData, rec: usize, seq: usize) -> Self {
        let mut keys = HashMap::new();
        for (i, c) in data.columns.iter().enumerate() {
            keys.entry(name_key(c)).or_insert(i);
        }
        Ctx { data, keys, rec, seq, skip: false, missing: BTreeSet::new() }
    }

    fn value(&mut self, name: &str) -> String {
        match self.keys.get(&name_key(name)) {
            Some(&col) => self
                .data
                .rows
                .get(self.rec)
                .and_then(|r| r.get(col))
                .cloned()
                .unwrap_or_default(),
            None => {
                self.missing.insert(name.to_string());
                String::new()
            }
        }
    }

    fn next_record(&mut self) {
        self.rec += 1;
    }
}

/// The field's type: the first word of its code. The code may come in many
/// `instrText` pieces (even a character each), so they are joined up to the
/// first nested field.
fn field_kind(f: &Field) -> String {
    let mut head = String::new();
    for p in &f.code {
        match p {
            Piece::Text(t) => head.push_str(t),
            Piece::Field(_) => break,
        }
    }
    head.split_whitespace().next().map(|w| w.to_ascii_uppercase()).unwrap_or_default()
}

fn is_merge_kind(kind: &str) -> bool {
    matches!(kind, "MERGEFIELD" | "IF" | "MERGEREC" | "MERGESEQ" | "NEXT" | "NEXTIF" | "SKIPIF")
}

/// Evaluate a field. `None`: it is not a merge field (or depends on one
/// that is not), so it stays as it is.
fn evaluate(fields: &[Field], id: usize, ctx: &mut Ctx) -> Option<String> {
    let f = &fields[id];
    let kind = field_kind(f);
    if !is_merge_kind(&kind) {
        return None;
    }
    let mut pieces: Vec<(bool, String)> = Vec::new();
    for p in &f.code {
        match p {
            Piece::Text(t) => pieces.push((false, t.clone())),
            Piece::Field(n) => pieces.push((true, evaluate(fields, *n, ctx)?)),
        }
    }
    let toks = tokenize(&pieces);
    let (args, sw) = split_switches(&toks);
    let arg = |i: usize| args.get(i).map(|t| t.text.as_str()).unwrap_or("");
    let value = match kind.as_str() {
        "MERGEFIELD" => {
            let v = if ctx.rec < ctx.data.rows.len() { ctx.value(arg(1)) } else { String::new() };
            if v.is_empty() {
                return Some(String::new());
            }
            let v = format_value(&v, &sw);
            return Some(format!("{}{}{}", sw.before.as_deref().unwrap_or(""), v, sw.after.as_deref().unwrap_or("")));
        }
        "MERGEREC" => (ctx.rec + 1).to_string(),
        "MERGESEQ" => ctx.seq.to_string(),
        "NEXT" => {
            ctx.next_record();
            return Some(String::new());
        }
        "NEXTIF" | "SKIPIF" => {
            if compare(arg(1), arg(2), arg(3), args.get(3).is_some_and(|t| t.quoted)) {
                if kind == "NEXTIF" {
                    ctx.next_record();
                } else {
                    ctx.skip = true;
                }
            }
            return Some(String::new());
        }
        "IF" => {
            let yes = compare(arg(1), arg(2), arg(3), args.get(3).is_some_and(|t| t.quoted));
            if yes { arg(4).to_string() } else { arg(5).to_string() }
        }
        _ => return None,
    };
    Some(format_value(&value, &sw))
}

fn format_value(v: &str, sw: &Switches) -> String {
    let mut v = v.to_string();
    if let Some(p) = &sw.number {
        v = format_number(&v, p);
    } else if let Some(p) = &sw.date {
        v = format_date(&v, p);
    }
    for c in &sw.cases {
        v = match c.as_str() {
            "upper" => v.to_uppercase(),
            "lower" => v.to_lowercase(),
            "firstcap" => first_cap(&v),
            "caps" => v.split(' ').map(first_cap).collect::<Vec<_>>().join(" "),
            _ => v,
        };
    }
    v
}

fn first_cap(s: &str) -> String {
    let mut c = s.chars();
    match c.next() {
        Some(f) => f.to_uppercase().collect::<String>() + c.as_str(),
        None => String::new(),
    }
}

fn number_of(s: &str) -> Option<f64> {
    let t = s.trim().replace(',', "");
    if t.is_empty() {
        return None;
    }
    t.parse::<f64>().ok()
}

/// Word's comparison of an `IF`/`SKIPIF`/`NEXTIF`: numbers as numbers, text
/// with case, and `*`/`?` wildcards in a quoted right side of `=`/`<>`.
fn compare(a: &str, op: &str, b: &str, b_quoted: bool) -> bool {
    if let (Some(x), Some(y)) = (number_of(a), number_of(b)) {
        return match op {
            "=" => x == y,
            "<>" => x != y,
            "<" => x < y,
            "<=" => x <= y,
            ">" => x > y,
            ">=" => x >= y,
            _ => false,
        };
    }
    let eq = if b_quoted && (b.contains('*') || b.contains('?')) { wildcard(a, b) } else { a == b };
    match op {
        "=" => eq,
        "<>" => !eq,
        "<" => a < b,
        "<=" => a <= b,
        ">" => a > b,
        ">=" => a >= b,
        _ => false,
    }
}

fn wildcard(s: &str, pat: &str) -> bool {
    let s: Vec<char> = s.chars().collect();
    let p: Vec<char> = pat.chars().collect();
    let (mut i, mut j, mut star, mut mark) = (0usize, 0usize, None::<usize>, 0usize);
    while i < s.len() {
        if j < p.len() && (p[j] == '?' || p[j] == s[i]) {
            i += 1;
            j += 1;
        } else if j < p.len() && p[j] == '*' {
            star = Some(j);
            mark = i;
            j += 1;
        } else if let Some(st) = star {
            j = st + 1;
            mark += 1;
            i = mark;
        } else {
            return false;
        }
    }
    while j < p.len() && p[j] == '*' {
        j += 1;
    }
    j == p.len()
}

// ---------------------------------------------------------------------------
// Number pictures (\#)
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq)]
enum NTok {
    Digit(char),
    Point,
    Group,
    Sign(char),
    Lit(String),
}

fn picture_sections(p: &str) -> Vec<String> {
    let mut out = vec![String::new()];
    let mut q: Option<char> = None;
    for c in p.chars() {
        match q {
            Some(qc) => {
                if c == qc {
                    q = None;
                }
                out.last_mut().unwrap().push(c);
            }
            None => {
                if c == ';' {
                    out.push(String::new());
                } else {
                    if c == '\'' || c == '"' {
                        q = Some(c);
                    }
                    out.last_mut().unwrap().push(c);
                }
            }
        }
    }
    out
}

fn number_tokens(sec: &str) -> Vec<NTok> {
    let mut out = Vec::new();
    let mut chars = sec.chars();
    let mut seen_point = false;
    while let Some(c) = chars.next() {
        match c {
            '0' | '#' | 'x' => out.push(NTok::Digit(c)),
            '.' if !seen_point => {
                seen_point = true;
                out.push(NTok::Point)
            }
            ',' if !seen_point => out.push(NTok::Group),
            '-' | '+' => out.push(NTok::Sign(c)),
            '\'' | '"' => {
                let lit: String = chars.by_ref().take_while(|&k| k != c).collect();
                out.push(NTok::Lit(lit));
            }
            c => out.push(NTok::Lit(c.to_string())),
        }
    }
    out
}

/// Round a decimal string's magnitude to `places` decimals, half away from
/// zero, exactly. Returns (integer digits without leading zeros, fraction
/// digits of length `places`).
fn round_decimal(s: &str, places: usize) -> (String, String) {
    let s = s.trim_start_matches(['-', '+']);
    let (ip, fp) = s.split_once('.').unwrap_or((s, ""));
    let mut digits: Vec<u8> = ip.bytes().chain(fp.bytes()).map(|b| b - b'0').collect();
    let int_len = ip.len();
    let keep = int_len + places;
    let round_up = digits.get(keep).is_some_and(|&d| d >= 5);
    digits.resize(keep.max(digits.len()), 0);
    digits.truncate(keep);
    if round_up {
        let mut k = keep;
        loop {
            if k == 0 {
                digits.insert(0, 1);
                break;
            }
            k -= 1;
            if digits[k] == 9 {
                digits[k] = 0;
            } else {
                digits[k] += 1;
                break;
            }
        }
    }
    let int_count = digits.len() - places;
    let int: String = digits[..int_count].iter().map(|d| (b'0' + d) as char).collect();
    let frac: String = digits[int_count..].iter().map(|d| (b'0' + d) as char).collect();
    (int.trim_start_matches('0').to_string(), frac)
}

/// The number at the start of a value (`"12abc"` is 12, `"1,234"` 1234), as
/// a plain decimal string.
fn leading_number(v: &str) -> Option<String> {
    let t: String = v.trim().chars().filter(|&c| c != ',').collect();
    let mut out = String::new();
    let mut chars = t.chars().peekable();
    if let Some(&c) = chars.peek() {
        if c == '-' || c == '+' {
            out.push(c);
            chars.next();
        }
    }
    let mut digits = 0;
    let mut point = false;
    for c in chars {
        if c.is_ascii_digit() {
            out.push(c);
            digits += 1;
        } else if c == '.' && !point {
            point = true;
            out.push(c);
        } else {
            break;
        }
    }
    if digits == 0 {
        return None;
    }
    if out.ends_with('.') {
        out.pop();
    }
    if out.starts_with('.') || out.starts_with("-.") || out.starts_with("+.") {
        out = out.replacen('.', "0.", 1);
    }
    Some(out)
}

/// Format a value with a Word numeric picture (`\# "#,##0.00"`), as measured:
/// an unused `#` is a space, an unused `0` a zero, `x` drops the digits to its
/// left; a group comma shows only with a digit to its left; with no sign in
/// the picture a minus goes just before the first `#`/`0`; `;` separates the
/// positive, negative and zero pictures; rounding is half away from zero.
pub fn format_number(value: &str, picture: &str) -> String {
    let Some(num) = leading_number(value) else { return value.to_string() };
    let negative = num.starts_with('-') && num.chars().any(|c| c.is_ascii_digit() && c != '0');
    let is_zero = !num.chars().any(|c| c.is_ascii_digit() && c != '0');
    let sections = picture_sections(picture);
    let (sec, explicit_negative) = if negative && sections.len() >= 2 {
        (&sections[1], true)
    } else if is_zero && sections.len() >= 3 {
        (&sections[2], false)
    } else {
        (&sections[0], false)
    };
    let toks = number_tokens(sec);
    let point = toks.iter().position(|t| *t == NTok::Point);
    let int_slots: Vec<usize> = toks
        .iter()
        .enumerate()
        .filter(|(i, t)| matches!(t, NTok::Digit(_)) && point.map_or(true, |p| *i < p))
        .map(|(i, _)| i)
        .collect();
    let frac_slots: Vec<usize> = toks
        .iter()
        .enumerate()
        .filter(|(i, t)| matches!(t, NTok::Digit(_)) && point.is_some_and(|p| *i > p))
        .map(|(i, _)| i)
        .collect();
    let (int_digits, frac_digits) = round_decimal(&num, frac_slots.len());
    let rounded_zero = int_digits.is_empty() && frac_digits.bytes().all(|b| b == b'0');
    let shown_negative = negative && !rounded_zero;
    let grouping = toks.iter().enumerate().any(|(i, t)| *t == NTok::Group && point.map_or(true, |p| i < p));

    // Integer placeholders, filled from the right.
    let mut slot_text: HashMap<usize, String> = HashMap::new();
    let digits: Vec<char> = int_digits.chars().collect();
    let n = int_slots.len();
    let x_slot = int_slots.iter().position(|&i| toks[i] == NTok::Digit('x'));
    for (k, &slot) in int_slots.iter().enumerate() {
        // Slot k (of n) takes digit index (len - n + k).
        let from_right = n - k; // 1-based position from the right
        let text = if from_right <= digits.len() {
            let di = digits.len() - from_right;
            if k == 0 && x_slot.is_none() && digits.len() > n {
                digits[..=di].iter().collect::<String>()
            } else {
                digits[di].to_string()
            }
        } else {
            match toks[slot] {
                NTok::Digit('0') => "0".to_string(),
                _ => " ".to_string(),
            }
        };
        slot_text.insert(slot, text);
    }
    let first_sig_slot = int_slots.iter().copied().find(|&i| matches!(toks[i], NTok::Digit('0') | NTok::Digit('#')));
    let mut out = String::new();
    let has_sign = toks.iter().any(|t| matches!(t, NTok::Sign(_)));
    let mut int_part = String::new();
    let mut in_int = false;
    let mut frac_k = 0;
    let all_frac_zero = frac_digits.bytes().all(|b| b == b'0');
    let flush_int = |out: &mut String, int_part: &mut String| {
        if !int_part.is_empty() {
            out.push_str(&if grouping { group_digits(int_part) } else { std::mem::take(int_part) });
            int_part.clear();
        }
    };
    for (i, t) in toks.iter().enumerate() {
        match t {
            NTok::Digit(_) if point.map_or(true, |p| i < p) => {
                if Some(i) == first_sig_slot && shown_negative && !has_sign && !explicit_negative {
                    flush_int(&mut out, &mut int_part);
                    out.push('-');
                }
                in_int = true;
                int_part.push_str(&slot_text[&i]);
            }
            NTok::Group => {}
            NTok::Digit(c) => {
                flush_int(&mut out, &mut int_part);
                let d = frac_digits.as_bytes()[frac_k] as char;
                let trailing_zero = frac_digits[frac_k..].bytes().all(|b| b == b'0');
                let text = if *c == '0' {
                    d
                } else if trailing_zero {
                    if frac_k == 0 && all_frac_zero && int_digits.is_empty() { '0' } else { ' ' }
                } else {
                    d
                };
                out.push(text);
                frac_k += 1;
            }
            NTok::Point => {
                flush_int(&mut out, &mut int_part);
                out.push('.');
            }
            NTok::Sign(c) => {
                flush_int(&mut out, &mut int_part);
                out.push(if shown_negative {
                    '-'
                } else if *c == '+' && !rounded_zero {
                    '+'
                } else {
                    ' '
                });
            }
            NTok::Lit(l) => {
                flush_int(&mut out, &mut int_part);
                out.push_str(l);
            }
        }
    }
    let _ = in_int;
    flush_int(&mut out, &mut int_part);
    out
}

/// Put a group comma between every three digits, counting from the right;
/// leading spaces are left alone.
fn group_digits(s: &str) -> String {
    let lead = s.len() - s.trim_start().len();
    let digits = &s[lead..];
    let n = digits.chars().count();
    let mut out = s[..lead].to_string();
    for (i, c) in digits.chars().enumerate() {
        if i > 0 && (n - i) % 3 == 0 {
            out.push(',');
        }
        out.push(c);
    }
    out
}

// ---------------------------------------------------------------------------
// Date pictures (\@)
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Copy, PartialEq)]
struct DateTime {
    y: i32,
    m: u32,
    d: u32,
    hh: u32,
    mi: u32,
    ss: u32,
}

fn parse_date(v: &str) -> Option<DateTime> {
    let v = v.trim();
    let (date_part, time_part) = match v.find(' ') {
        Some(k) if !v.contains('年') => (&v[..k], v[k + 1..].trim()),
        _ => (v, ""),
    };
    // Word reads past leading letters: "R8.4.1" is 8/4/01 to it.
    let date_part = date_part.trim_start_matches(|c: char| c.is_ascii_alphabetic());
    let nums: Vec<u32> = date_part
        .split(|c: char| matches!(c, '/' | '-' | '年' | '月' | '日' | '.'))
        .filter(|s| !s.is_empty())
        .map(|s| s.trim().parse::<u32>())
        .collect::<Result<_, _>>()
        .ok()?;
    if nums.len() != 3 {
        return None;
    }
    let (y, m, d) = if date_part.split(['/', '-', '.']).next().is_some_and(|s| s.len() == 4) || date_part.contains('年') {
        (nums[0] as i32, nums[1], nums[2])
    } else {
        // m/d/yyyy (and m/d/yy), as Excel's values come to Word.
        let y = if nums[2] < 100 { 2000 + nums[2] as i32 } else { nums[2] as i32 };
        (y, nums[0], nums[1])
    };
    if !(1..=12).contains(&m) || d == 0 || d > days_in_month(y, m) {
        return None;
    }
    let (mut hh, mut mi, mut ss) = (0, 0, 0);
    if !time_part.is_empty() {
        let upper = time_part.to_ascii_uppercase();
        let pm = upper.ends_with("PM");
        let am = upper.ends_with("AM");
        let t = upper.trim_end_matches("PM").trim_end_matches("AM").trim();
        let parts: Vec<u32> = t.split(':').map(|s| s.trim().parse::<u32>().unwrap_or(0)).collect();
        hh = *parts.first().unwrap_or(&0);
        mi = *parts.get(1).unwrap_or(&0);
        ss = *parts.get(2).unwrap_or(&0);
        if pm && hh < 12 {
            hh += 12;
        }
        if am && hh == 12 {
            hh = 0;
        }
    }
    Some(DateTime { y, m, d, hh, mi, ss })
}

fn days_in_month(y: i32, m: u32) -> u32 {
    match m {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if (y % 4 == 0 && y % 100 != 0) || y % 400 == 0 => 29,
        _ => 28,
    }
}

fn weekday(dt: &DateTime) -> usize {
    // 0 = Sunday (Zeller-like, via days from civil).
    let (y, m) = if dt.m <= 2 { (dt.y - 1, dt.m + 12) } else { (dt.y, dt.m) };
    let k = y % 100;
    let j = y / 100;
    let h = (dt.d as i32 + (13 * (m as i32 + 1)) / 5 + k + k / 4 + j / 4 + 5 * j) % 7;
    ((h + 6) % 7) as usize
}

const MONTHS: [&str; 12] = [
    "January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
    "November", "December",
];
const DAYS: [&str; 7] = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const JA_DAYS: [&str; 7] = ["日", "月", "火", "水", "木", "金", "土"];

/// The Japanese era of a date: (name, initial, year of the era).
fn era(dt: &DateTime) -> (&'static str, &'static str, i32) {
    let key = (dt.y, dt.m, dt.d);
    if key >= (2019, 5, 1) {
        ("令和", "R", dt.y - 2018)
    } else if key >= (1989, 1, 8) {
        ("平成", "H", dt.y - 1988)
    } else if key >= (1926, 12, 25) {
        ("昭和", "S", dt.y - 1925)
    } else if key >= (1912, 7, 30) {
        ("大正", "T", dt.y - 1911)
    } else {
        ("明治", "M", dt.y - 1867)
    }
}

/// Format a value with a Word date picture (`\@ "yyyy年M月d日"`). A value
/// that is not a date is left as it is.
pub fn format_date(value: &str, picture: &str) -> String {
    let Some(dt) = parse_date(value) else { return value.to_string() };
    let p: Vec<char> = picture.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    let run = |i: usize, c: char| p[i..].iter().take_while(|&&k| k == c).count();
    while i < p.len() {
        let c = p[i];
        match c {
            '\'' | '"' => {
                let lit: String = p[i + 1..].iter().take_while(|&&k| k != c).collect();
                i += lit.chars().count() + 2;
                out.push_str(&lit);
                continue;
            }
            'y' | 'Y' => {
                let n = run(i, c);
                if n >= 3 {
                    out.push_str(&format!("{:04}", dt.y));
                } else {
                    out.push_str(&format!("{:02}", dt.y.rem_euclid(100)));
                }
                i += n;
            }
            'M' => {
                let n = run(i, c);
                out.push_str(&match n {
                    1 => dt.m.to_string(),
                    2 => format!("{:02}", dt.m),
                    3 => MONTHS[dt.m as usize - 1][..3].to_string(),
                    _ => MONTHS[dt.m as usize - 1].to_string(),
                });
                i += n;
            }
            'd' | 'D' => {
                let n = run(i, c);
                let wd = weekday(&dt);
                out.push_str(&match n {
                    1 => dt.d.to_string(),
                    2 => format!("{:02}", dt.d),
                    3 => DAYS[wd][..3].to_string(),
                    _ => DAYS[wd].to_string(),
                });
                i += n;
            }
            'a' if run(i, 'a') >= 3 => {
                out.push_str(JA_DAYS[weekday(&dt)]);
                i += 3;
            }
            'A' | 'a' if p[i..].iter().collect::<String>().to_ascii_uppercase().starts_with("AM/PM") => {
                let upper = c == 'A';
                let s = if dt.hh < 12 { "AM" } else { "PM" };
                out.push_str(&if upper { s.to_string() } else { s.to_lowercase() });
                i += 5;
            }
            'H' | 'h' => {
                let n = run(i, c);
                let has_ampm = picture.to_ascii_uppercase().contains("AM/PM");
                let h = if c == 'h' && has_ampm {
                    let h = dt.hh % 12;
                    if h == 0 { 12 } else { h }
                } else if c == 'h' {
                    let h = dt.hh % 12;
                    if h == 0 { 12 } else { h }
                } else {
                    dt.hh
                };
                out.push_str(&if n >= 2 { format!("{:02}", h) } else { h.to_string() });
                i += n;
            }
            'm' => {
                let n = run(i, c);
                out.push_str(&if n >= 2 { format!("{:02}", dt.mi) } else { dt.mi.to_string() });
                i += n;
            }
            's' | 'S' => {
                let n = run(i, c);
                out.push_str(&if n >= 2 { format!("{:02}", dt.ss) } else { dt.ss.to_string() });
                i += n;
            }
            'g' => {
                let n = run(i, c);
                let (name, initial, _) = era(&dt);
                out.push_str(match n {
                    1 => initial,
                    2 => &name[..name.char_indices().nth(1).map_or(name.len(), |(k, _)| k)],
                    _ => name,
                });
                i += n;
            }
            'e' => {
                let n = run(i, c);
                let (_, _, ey) = era(&dt);
                out.push_str(&if n >= 2 { format!("{:02}", ey) } else { ey.to_string() });
                i += n;
            }
            c => {
                out.push(c);
                i += 1;
            }
        }
    }
    out
}

// ---------------------------------------------------------------------------
// Rewriting a part
// ---------------------------------------------------------------------------

/// rPr children in schema order (CT_RPr), to put `noProof` in its place.
const RPR_ORDER: &[&str] = &[
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
    "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
    "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
    "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange",
];

fn rpr_with_no_proof(rpr: &str, w: &str) -> String {
    let np = format!("<{}/>", qn(w, "noProof"));
    if rpr.is_empty() {
        return format!("<{0}>{1}</{0}>", qn(w, "rPr"), np);
    }
    let els = scan(rpr);
    let Some(root) = els.first() else { return rpr.to_string() };
    if root.end == root.open_end {
        // <w:rPr/>
        return format!("<{0}>{1}</{0}>", qn(w, "rPr"), np);
    }
    let pos_np = RPR_ORDER.iter().position(|&n| n == "noProof").unwrap();
    let mut insert_at = root.close_start;
    for &c in &root.children {
        let l = local(&els[c].name);
        if l == "noProof" {
            return rpr.to_string();
        }
        if RPR_ORDER.iter().position(|&n| n == l).is_some_and(|k| k > pos_np) {
            insert_at = els[c].start;
            break;
        }
    }
    format!("{}{}{}", &rpr[..insert_at], np, &rpr[insert_at..])
}

fn value_runs(value: &str, rpr: &str, w: &str) -> String {
    if value.is_empty() {
        return String::new();
    }
    let rpr = rpr_with_no_proof(rpr, w);
    let mut content = String::new();
    let text = value.replace("\r\n", "\n");
    let mut buf = String::new();
    let flush = |content: &mut String, buf: &mut String| {
        if !buf.is_empty() {
            let sp = if buf.starts_with(' ') || buf.ends_with(' ') || buf.contains("  ") {
                " xml:space=\"preserve\""
            } else {
                ""
            };
            content.push_str(&format!("<{0}{1}>{2}</{0}>", qn(w, "t"), sp, escape(buf)));
            buf.clear();
        }
    };
    for c in text.chars() {
        match c {
            '\n' | '\r' | '\u{b}' => {
                flush(&mut content, &mut buf);
                content.push_str(&format!("<{}/>", qn(w, "br")));
            }
            '\t' => {
                flush(&mut content, &mut buf);
                content.push_str(&format!("<{}/>", qn(w, "tab")));
            }
            c => buf.push(c),
        }
    }
    flush(&mut content, &mut buf);
    format!("<{0}>{1}{2}</{0}>", qn(w, "r"), rpr, content)
}

/// Replacements in a part: byte ranges (start, end) and their new text.
type Edits = Vec<(usize, usize, String)>;

fn apply_edits(xml: &str, mut edits: Edits) -> String {
    edits.sort_by(|a, b| a.0.cmp(&b.0).then(b.1.cmp(&a.1)));
    // Drop an edit that sits inside one already taken.
    let mut kept: Edits = Vec::new();
    for e in edits {
        if kept.last().is_some_and(|k| e.0 < k.1) {
            continue;
        }
        kept.push(e);
    }
    let mut out = String::with_capacity(xml.len());
    let mut at = 0;
    for (s, e, t) in kept {
        out.push_str(&xml[at..s]);
        out.push_str(&t);
        at = e;
    }
    out.push_str(&xml[at..]);
    out
}

/// What a merge pass did to a part.
struct PartMerge {
    xml: String,
    /// The part had merge fields.
    merged: bool,
}

/// Merge one part (document, header, footer, notes) for the record `ctx`
/// is on. Paragraphs are taken in document order, so `NEXT` moves the
/// record for everything after it.
fn merge_part(xml: &str, w: &str, ctx: &mut Ctx, opts: &MergeOptions, drop_bookmarks: bool) -> PartMerge {
    let w = w.to_string();
    let els = scan(xml);
    let mut edits: Edits = Vec::new();
    let mut merged = false;
    let paragraphs: Vec<usize> = (0..els.len()).filter(|&i| is_w(&els, i, &w, "p")).collect();
    for &p in &paragraphs {
        if ctx.skip {
            break;
        }
        let para = paragraph_fields(xml, &els, p, &w);
        if para.fields.is_empty() {
            continue;
        }
        // Item index ranges covered by an evaluated field, with the value runs.
        let mut covered: Vec<(usize, usize, String)> = Vec::new();
        let mut merge_fields_here = false;
        for (id, f) in para.fields.iter().enumerate() {
            if f.parent.is_some() {
                continue;
            }
            let Some(end) = f.end else { continue };
            if !is_merge_kind(&field_kind(f)) {
                continue;
            }
            let Some(value) = evaluate(&para.fields, id, ctx) else { continue };
            merge_fields_here = true;
            let toks = {
                let pieces: Vec<(bool, String)> = f
                    .code
                    .iter()
                    .filter_map(|p| if let Piece::Text(t) = p { Some((false, t.clone())) } else { None })
                    .collect();
                tokenize(&pieces)
            };
            let (_, sw) = split_switches(&toks);
            let rpr = if sw.mergeformat {
                f.result_rpr.clone().unwrap_or_else(|| f.code_rpr.clone())
            } else {
                f.code_rpr.clone()
            };
            covered.push((f.begin, end, value_runs(&value, &rpr, &w)));
            if ctx.skip {
                break;
            }
        }
        if covered.is_empty() {
            continue;
        }
        merged = true;
        let begin_of = |i: usize| covered.iter().find(|c| c.0 == i).map(|c| c.2.clone());
        let is_covered = |i: usize| covered.iter().any(|c| c.0 <= i && i <= c.1);

        // Blank-line suppression: the paragraph keeps nothing but spaces.
        if opts.suppress_blank_lines && merge_fields_here && !paragraph_has_section(&els, p, &w)
            && !only_block_in_container(&els, p, &w)
        {
            let mut blank = true;
            for (i, item) in para.items.iter().enumerate() {
                if let Some(v) = begin_of(i) {
                    if !v.is_empty() {
                        blank = false;
                    }
                    continue;
                }
                if is_covered(i) {
                    continue;
                }
                match *item {
                    Item::Simple(_) => blank = false,
                    Item::RunChild { child, .. } => {
                        let l = local(&els[child].name);
                        match l {
                            "t" => {
                                if !unescape(inner(xml, &els[child])).trim().is_empty() {
                                    blank = false;
                                }
                            }
                            "lastRenderedPageBreak" | "tab" | "proofErr" => {}
                            _ => blank = false,
                        }
                    }
                }
            }
            if blank {
                edits.push((els[p].start, els[p].end, String::new()));
                continue;
            }
        }

        // Rewrite each run the covered items touch.
        let mut runs: Vec<usize> = Vec::new();
        for (i, item) in para.items.iter().enumerate() {
            match *item {
                Item::Simple(s) => {
                    if let Some(v) = begin_of(i) {
                        edits.push((els[s].start, els[s].end, v));
                    } else if is_covered(i) {
                        edits.push((els[s].start, els[s].end, String::new()));
                    }
                }
                Item::RunChild { run, .. } => {
                    if is_covered(i) && !runs.contains(&run) {
                        runs.push(run);
                    }
                }
            }
        }
        for run in runs {
            let r = &els[run];
            let open = &xml[r.start..r.open_end];
            let rpr = rpr_of(xml, &els, run, &w);
            let close = format!("</{}>", r.name);
            let mut out = String::new();
            let mut kept = String::new();
            let flush = |out: &mut String, kept: &mut String| {
                if !kept.is_empty() {
                    out.push_str(open);
                    out.push_str(&rpr);
                    out.push_str(kept);
                    out.push_str(&close);
                    kept.clear();
                }
            };
            for (i, item) in para.items.iter().enumerate() {
                let Item::RunChild { run: rr, child } = *item else { continue };
                if rr != run {
                    continue;
                }
                if let Some(v) = begin_of(i) {
                    flush(&mut out, &mut kept);
                    out.push_str(&v);
                } else if !is_covered(i) {
                    kept.push_str(&xml[els[child].start..els[child].end]);
                }
            }
            flush(&mut out, &mut kept);
            edits.push((r.start, r.end, out));
        }
    }
    if drop_bookmarks {
        for (i, e) in els.iter().enumerate() {
            if is_w(&els, i, &w, "bookmarkStart") || is_w(&els, i, &w, "bookmarkEnd") {
                edits.push((e.start, e.end, String::new()));
            }
        }
    }
    PartMerge { xml: apply_edits(xml, edits), merged }
}

fn paragraph_has_section(els: &[El], p: usize, w: &str) -> bool {
    els[p].children.iter().any(|&c| {
        is_w(els, c, w, "pPr") && els[c].children.iter().any(|&k| is_w(els, k, w, "sectPr"))
    })
}

/// A table cell, text box or note needs a paragraph; the last one stays.
fn only_block_in_container(els: &[El], p: usize, w: &str) -> bool {
    let Some(parent) = els[p].parent else { return false };
    if is_w(els, parent, w, "body") {
        return false;
    }
    els[parent].children.iter().filter(|&&c| is_w(els, c, w, "p") || is_w(els, c, w, "tbl")).count() <= 1
}

/// Does a part hold a merge field at all? (A quick look, before any work.)
fn has_merge_fields(xml: &str) -> bool {
    ["MERGEFIELD", "MERGEREC", "MERGESEQ", "NEXT", "SKIPIF", " IF ", "IF "].iter().any(|k| xml.contains(k))
}

// ---------------------------------------------------------------------------
// Packages
// ---------------------------------------------------------------------------

struct Pkg {
    names: Vec<String>,
    parts: HashMap<String, Vec<u8>>,
}

impl Pkg {
    fn read(bytes: &[u8]) -> Result<Pkg, MergeError> {
        let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).map_err(|e| MergeError::Package(e.to_string()))?;
        let mut names = Vec::new();
        let mut parts = HashMap::new();
        for i in 0..zip.len() {
            let mut f = zip.by_index(i).map_err(|e| MergeError::Package(e.to_string()))?;
            if f.is_dir() {
                continue;
            }
            let mut buf = Vec::new();
            f.read_to_end(&mut buf).map_err(|e| MergeError::Package(e.to_string()))?;
            names.push(f.name().to_string());
            parts.insert(f.name().to_string(), buf);
        }
        Ok(Pkg { names, parts })
    }

    fn text(&self, name: &str) -> Option<String> {
        self.parts.get(name).map(|b| String::from_utf8_lossy(b).into_owned())
    }

    fn set(&mut self, name: &str, text: String) {
        if !self.parts.contains_key(name) {
            self.names.push(name.to_string());
        }
        self.parts.insert(name.to_string(), text.into_bytes());
    }

    fn write(&self) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        for n in &self.names {
            if let Some(b) = self.parts.get(n) {
                let _ = zip.start_file(n.as_str(), opts);
                let _ = zip.write_all(b);
            }
        }
        zip.finish().map(|c| c.into_inner()).unwrap_or_default()
    }

    /// The main document part's name.
    fn main_part(&self) -> Option<String> {
        let rels = self.text("_rels/.rels")?;
        let els = scan(&rels);
        for e in &els {
            if local(&e.name) == "Relationship"
                && attr(&rels, e, "Type").is_some_and(|t| t.ends_with("/officeDocument"))
            {
                let t = attr(&rels, e, "Target")?;
                return Some(t.trim_start_matches('/').to_string());
            }
        }
        None
    }
}

fn rels_name(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((dir, file)) => format!("{dir}/_rels/{file}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

fn resolve(base_part: &str, target: &str) -> String {
    if let Some(t) = target.strip_prefix('/') {
        return t.to_string();
    }
    let mut segs: Vec<&str> = base_part.rsplit_once('/').map_or(vec![], |(d, _)| d.split('/').collect());
    for s in target.split('/') {
        match s {
            ".." => {
                segs.pop();
            }
            "." | "" => {}
            s => segs.push(s),
        }
    }
    segs.join("/")
}

#[derive(Debug, Clone)]
struct Rel {
    id: String,
    kind: String,
    target: String,
}

fn read_rels(xml: &str) -> Vec<Rel> {
    let els = scan(xml);
    els.iter()
        .filter(|e| local(&e.name) == "Relationship")
        .filter(|e| attr(xml, e, "TargetMode") != Some("External"))
        .map(|e| Rel {
            id: attr(xml, e, "Id").unwrap_or("").to_string(),
            kind: attr(xml, e, "Type").unwrap_or("").rsplit('/').next().unwrap_or("").to_string(),
            target: attr(xml, e, "Target").unwrap_or("").to_string(),
        })
        .collect()
}

/// The document's body: everything between `<w:body>` and its last
/// `sectPr`, and that `sectPr`.
struct Body {
    head: String,
    content: String,
    sect: String,
    tail: String,
}

fn split_document(xml: &str) -> Option<Body> {
    let w = w_prefix(xml);
    let els = scan(xml);
    let body = (0..els.len()).find(|&i| is_w(&els, i, &w, "body"))?;
    let b = &els[body];
    let last_sect = b.children.iter().copied().rev().find(|&c| is_w(&els, c, &w, "sectPr"));
    let (content_end, sect, after) = match last_sect {
        Some(s) => (els[s].start, xml[els[s].start..els[s].end].to_string(), els[s].end),
        None => (b.close_start, String::new(), b.close_start),
    };
    Some(Body {
        head: xml[..b.open_end].to_string(),
        content: xml[b.open_end..content_end].to_string(),
        sect,
        tail: xml[after..].to_string(),
    })
}

/// sectPr children in schema order (CT_SectPr).
const SECT_ORDER: &[&str] = &[
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
    "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
    "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
];

/// Put `new_el` in `sect` in schema order, replacing an element of the same
/// name (or, with `merge_attr`, setting one attribute on it).
fn sect_set(sect: &str, w: &str, name: &str, new_el: &str, merge_attr: Option<(&str, &str)>) -> String {
    let els = scan(sect);
    let Some(root) = els.first() else { return sect.to_string() };
    let order = |l: &str| SECT_ORDER.iter().position(|&n| n == l).unwrap_or(usize::MAX - 1);
    let mine = order(name);
    for &c in &root.children {
        if local(&els[c].name) == name && prefix(&els[c].name) == w {
            if let Some((a, v)) = merge_attr {
                let tag = &sect[els[c].start..els[c].open_end];
                let qa = qn(w, a);
                let new_tag = if let Some(old) = attr(sect, &els[c], &qa) {
                    tag.replacen(&format!("{qa}=\"{old}\""), &format!("{qa}=\"{v}\""), 1)
                } else {
                    let cut = if tag.ends_with("/>") { tag.len() - 2 } else { tag.len() - 1 };
                    format!("{} {qa}=\"{v}\"{}", &tag[..cut], &tag[cut..])
                };
                return format!("{}{}{}", &sect[..els[c].start], new_tag, &sect[els[c].open_end..]);
            }
            return format!("{}{}{}", &sect[..els[c].start], new_el, &sect[els[c].end..]);
        }
    }
    let mut at = root.close_start;
    if root.end == root.open_end {
        // <w:sectPr/>
        let open = sect.trim_end_matches("/>");
        return format!("{open}>{new_el}</{}>", root.name);
    }
    for &c in &root.children {
        if order(local(&els[c].name)) > mine {
            at = els[c].start;
            break;
        }
    }
    format!("{}{}{}", &sect[..at], new_el, &sect[at..])
}

/// Swap the header/footer relationship ids a section points at.
fn remap_refs(xml: &str, map: &HashMap<String, String>) -> String {
    if map.is_empty() {
        return xml.to_string();
    }
    let w = w_prefix(xml);
    let els = scan(xml);
    let mut edits: Edits = Vec::new();
    let rid = "r:id";
    for (i, e) in els.iter().enumerate() {
        if is_w(&els, i, &w, "headerReference") || is_w(&els, i, &w, "footerReference") {
            if let Some(old) = attr(xml, e, rid) {
                if let Some(new) = map.get(old) {
                    let tag = &xml[e.start..e.open_end];
                    edits.push((e.start, e.open_end, tag.replacen(&format!("{rid}=\"{old}\""), &format!("{rid}=\"{new}\""), 1)));
                }
            }
        }
    }
    apply_edits(xml, edits)
}

/// Give drawings' `wp:docPr` ids that are unique across the merged copies,
/// and drop `w14:paraId`/`w14:textId`, which must be unique too.
fn renumber_ids(xml: &str, next_id: &mut u64) -> String {
    let els = scan(xml);
    let mut edits: Edits = Vec::new();
    for e in &els {
        let l = local(&e.name);
        if l == "docPr" || l == "cNvPr" && prefix(&e.name) == "pic" {
            if let Some(old) = attr(xml, e, "id") {
                let tag = &xml[e.start..e.open_end];
                *next_id += 1;
                let new_tag = tag.replacen(&format!(" id=\"{old}\""), &format!(" id=\"{next_id}\""), 1);
                edits.push((e.start, e.open_end, new_tag));
            }
        } else if l == "p" {
            let tag = &xml[e.start..e.open_end];
            if tag.contains("w14:paraId") || tag.contains("w14:textId") {
                let mut t = tag.to_string();
                for a in ["w14:paraId", "w14:textId"] {
                    if let Some(v) = attr(xml, e, a) {
                        t = t.replacen(&format!(" {a}=\"{v}\""), "", 1);
                    }
                }
                edits.push((e.start, e.open_end, t));
            }
        }
    }
    apply_edits(xml, edits)
}

/// Put a section break (`sect`) into the copy's last paragraph, or after
/// the copy in a paragraph of its own when the copy ends in a table.
fn close_with_section(content: &str, sect: &str, w: &str) -> String {
    let wrapper = format!("<x xmlns:{}=\"{}\">{}</x>", if w.is_empty() { "w" } else { w }, W_NS, content);
    let off = wrapper.find('>').unwrap() + 1;
    let els = scan(&wrapper);
    let root = &els[0];
    let last = root.children.iter().copied().rev().find(|&c| {
        let l = local(&els[c].name);
        l == "p" || l == "tbl" || l == "sdt"
    });
    let ppr = qn(w, "pPr");
    match last {
        Some(p) if is_w(&els, p, w, "p") && els[p].end == els[p].open_end => {
            // <w:p/>
            let mut s = wrapper.clone();
            s.replace_range(els[p].start..els[p].end, &format!("<{0}><{ppr}>{sect}</{ppr}></{0}>", qn(w, "p")));
            s[off..s.len() - 4].to_string()
        }
        Some(p) if is_w(&els, p, w, "p") => {
            let e = &els[p];
            let existing = e.children.iter().copied().find(|&c| is_w(&els, c, w, "pPr"));
            let (at, text) = match existing {
                Some(pp) if els[pp].end == els[pp].open_end => {
                    // <w:pPr/>
                    (els[pp].start, format!("<{ppr}>{sect}</{ppr}>"))
                }
                Some(pp) => {
                    let change = els[pp].children.iter().copied().find(|&c| is_w(&els, c, w, "pPrChange"));
                    let replace_old = els[pp].children.iter().copied().find(|&c| is_w(&els, c, w, "sectPr"));
                    if let Some(old) = replace_old {
                        let mut s = wrapper.clone();
                        s.replace_range(els[old].start..els[old].end, sect);
                        return s[off..s.len() - 4].to_string();
                    }
                    (change.map_or(els[pp].close_start, |c| els[c].start), sect.to_string())
                }
                None => (e.open_end, format!("<{ppr}>{sect}</{ppr}>")),
            };
            let mut s = wrapper.clone();
            let remove_to = match existing {
                Some(pp) if els[pp].end == els[pp].open_end => els[pp].end,
                _ => at,
            };
            s.replace_range(at..remove_to, &text);
            s[off..s.len() - 4].to_string()
        }
        _ => format!("{content}<{0}><{ppr}>{sect}</{ppr}></{0}>", qn(w, "p")),
    }
}

/// Header and footer parts the document points at: (rel id, part name).
fn header_footer_parts(pkg: &Pkg, main: &str) -> Vec<(String, String, String)> {
    let Some(rels) = pkg.text(&rels_name(main)) else { return vec![] };
    read_rels(&rels)
        .into_iter()
        .filter(|r| r.kind == "header" || r.kind == "footer")
        .map(|r| (r.id, r.kind, resolve(main, &r.target)))
        .collect()
}

fn content_type_of(ct: &str, part: &str) -> Option<String> {
    let els = scan(ct);
    let want = format!("/{part}");
    els.iter()
        .find(|e| local(&e.name) == "Override" && attr(ct, e, "PartName") == Some(want.as_str()))
        .and_then(|e| attr(ct, e, "ContentType").map(str::to_string))
}

fn add_override(ct: &str, part: &str, kind: &str) -> String {
    match ct.rfind("</Types>") {
        Some(at) => format!(
            "{}<Override PartName=\"/{}\" ContentType=\"{}\"/>{}",
            &ct[..at],
            part,
            kind,
            &ct[at..]
        ),
        None => ct.to_string(),
    }
}

fn add_rel(rels: &str, id: &str, kind_url: &str, target: &str) -> String {
    match rels.rfind("</Relationships>") {
        Some(at) => format!(
            "{}<Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"/>{}",
            &rels[..at],
            id,
            kind_url,
            target,
            &rels[at..]
        ),
        None => rels.to_string(),
    }
}

fn strip_mail_merge_settings(pkg: &mut Pkg, main: &str) {
    let Some(rels) = pkg.text(&rels_name(main)) else { return };
    for r in read_rels(&rels) {
        if r.kind == "settings" {
            let name = resolve(main, &r.target);
            if let Some(s) = pkg.text(&name) {
                let w = w_prefix(&s);
                let els = scan(&s);
                let edits: Edits = els
                    .iter()
                    .enumerate()
                    .filter(|(i, _)| is_w(&els, *i, &w, "mailMerge"))
                    .map(|(_, e)| (e.start, e.end, String::new()))
                    .collect();
                if !edits.is_empty() {
                    pkg.set(&name, apply_edits(&s, edits));
                }
            }
        }
    }
}

/// The field names a main document uses (body, headers, footers, notes),
/// in order of first use.
pub fn merge_fields(docx: &[u8]) -> Result<MergeFieldReport, MergeError> {
    let pkg = Pkg::read(docx)?;
    let main = pkg.main_part().ok_or(MergeError::NoDocument)?;
    let mut parts = vec![main.clone()];
    parts.extend(header_footer_parts(&pkg, &main).into_iter().map(|p| p.2));
    let mut fields: Vec<String> = Vec::new();
    for part in parts {
        let Some(xml) = pkg.text(&part) else { continue };
        let w = w_prefix(&xml);
        let els = scan(&xml);
        for p in (0..els.len()).filter(|&i| is_w(&els, i, &w, "p")) {
            let para = paragraph_fields(&xml, &els, p, &w);
            for f in &para.fields {
                if field_kind(f) != "MERGEFIELD" {
                    continue;
                }
                let pieces: Vec<(bool, String)> = f
                    .code
                    .iter()
                    .filter_map(|p| if let Piece::Text(t) = p { Some((false, t.clone())) } else { None })
                    .collect();
                let toks = tokenize(&pieces);
                let (args, _) = split_switches(&toks);
                if let Some(name) = args.get(1) {
                    if !fields.iter().any(|f| name_key(f) == name_key(&name.text)) {
                        fields.push(name.text.clone());
                    }
                }
            }
        }
    }
    Ok(MergeFieldReport { fields })
}

struct Copy {
    record: usize,
    body: String,
    /// Merged header/footer parts: rel id → xml.
    parts: Vec<(String, String)>,
}

fn record_range(data: &MergeData, opts: &MergeOptions) -> (usize, usize) {
    let first = opts.first_record.unwrap_or(1).max(1) - 1;
    let last = opts.last_record.map_or(data.rows.len(), |l| l.min(data.rows.len()));
    (first, last)
}

fn copies(pkg: &Pkg, main: &str, body: &Body, data: &MergeData, opts: &MergeOptions) -> Vec<Copy> {
    let hf = header_footer_parts(pkg, main);
    let (mut rec, last) = record_range(data, opts);
    // Only a header or footer that holds a merge field gets a copy per record.
    let hf_merge: Vec<(String, String)> = hf
        .iter()
        .filter_map(|(id, _, part)| pkg.text(part).filter(|x| has_merge_fields(x)).map(|x| (id.clone(), x)))
        .filter(|(_, x)| merge_part(x, &w_prefix(x), &mut Ctx::new(data, rec, 1), opts, false).merged)
        .collect();
    let w = w_prefix(&pkg.text(main).unwrap_or_default());
    let mut seq = 0;
    let mut out = Vec::new();
    // A document without a field still makes one copy per record.
    while rec < last {
        let mut ctx = Ctx::new(data, rec, seq + 1);
        let merged = merge_part(&body.content, &w, &mut ctx, opts, true);
        let next = ctx.rec + 1;
        if ctx.skip {
            rec = next;
            continue;
        }
        seq += 1;
        let mut parts = Vec::new();
        for (id, xml) in &hf_merge {
            let mut hctx = Ctx::new(data, rec, seq);
            parts.push((id.clone(), merge_part(xml, &w_prefix(xml), &mut hctx, opts, true).xml));
        }
        out.push(Copy { record: rec + 1, body: merged.xml, parts });
        rec = next;
    }
    out
}

/// Merge every record into one document, as Word's "merge to new document".
pub fn mail_merge(docx: &[u8], data: &MergeData, opts: &MergeOptions) -> Result<Vec<u8>, MergeError> {
    let mut pkg = Pkg::read(docx)?;
    let main = pkg.main_part().ok_or(MergeError::NoDocument)?;
    let doc = pkg.text(&main).ok_or(MergeError::NoDocument)?;
    let body = split_document(&doc).ok_or(MergeError::NoDocument)?;
    let w = w_prefix(&doc);
    let copies = copies(&pkg, &main, &body, data, opts);
    let hf = header_footer_parts(&pkg, &main);
    let mut rels = pkg.text(&rels_name(&main)).unwrap_or_default();
    let mut ct = pkg.text("[Content_Types].xml").unwrap_or_default();
    let mut content = String::new();
    let mut next_id = 0u64;
    let mut last_map: HashMap<String, String> = HashMap::new();
    for (n, copy) in copies.iter().enumerate() {
        // This record's own headers and footers.
        let mut map = HashMap::new();
        for (id, xml) in &copy.parts {
            let Some((_, kind, part)) = hf.iter().find(|h| &h.0 == id) else { continue };
            let dir = part.rsplit_once('/').map_or("", |(d, _)| d);
            let file = part.rsplit_once('/').map_or(part.as_str(), |(_, f)| f);
            let stem = file.trim_end_matches(".xml");
            let new_file = format!("{stem}_merge{}.xml", n + 1);
            let new_part = if dir.is_empty() { new_file.clone() } else { format!("{dir}/{new_file}") };
            pkg.set(&new_part, xml.clone());
            if let Some(r) = pkg.text(&rels_name(part)) {
                pkg.set(&rels_name(&new_part), r);
            }
            if let Some(k) = content_type_of(&ct, part) {
                ct = add_override(&ct, &new_part, &k);
            }
            let new_id = format!("rIdMerge{}{}", n + 1, id);
            let url = format!("http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}");
            let target = match main.rsplit_once('/') {
                Some((mdir, _)) if mdir == dir => new_file.clone(),
                _ => format!("/{new_part}"),
            };
            rels = add_rel(&rels, &new_id, &url, &target);
            map.insert(id.clone(), new_id);
        }
        let sect = remap_refs(&body.sect, &map);
        let sect = sect_set(&sect, &w, "pgNumType", &format!("<{} {}=\"1\"/>", qn(&w, "pgNumType"), qn(&w, "start")), Some(("start", "1")));
        let copy_body = remap_refs(&copy.body, &map);
        let copy_body = renumber_ids(&copy_body, &mut next_id);
        content.push_str(&close_with_section(&copy_body, &sect, &w));
        last_map = map;
    }
    // After the last record, an empty paragraph in a continuous section.
    let final_sect = remap_refs(&body.sect, &last_map);
    let final_sect = sect_set(&final_sect, &w, "type", &format!("<{} {}=\"continuous\"/>", qn(&w, "type"), qn(&w, "val")), None);
    if !copies.is_empty() {
        content.push_str(&format!("<{}/>", qn(&w, "p")));
    }
    let sect_out = if body.sect.is_empty() { String::new() } else { final_sect };
    pkg.set(&main, format!("{}{}{}{}", body.head, content, sect_out, body.tail));
    pkg.set(&rels_name(&main), rels);
    pkg.set("[Content_Types].xml", ct);
    strip_mail_merge_settings(&mut pkg, &main);
    Ok(pkg.write())
}

/// Merge each record into a document of its own.
pub fn mail_merge_each(docx: &[u8], data: &MergeData, opts: &MergeOptions) -> Result<Vec<MergedDocument>, MergeError> {
    let base = Pkg::read(docx)?;
    let main = base.main_part().ok_or(MergeError::NoDocument)?;
    let doc = base.text(&main).ok_or(MergeError::NoDocument)?;
    let body = split_document(&doc).ok_or(MergeError::NoDocument)?;
    let hf = header_footer_parts(&base, &main);
    let mut out = Vec::new();
    for copy in copies(&base, &main, &body, data, opts) {
        let mut pkg = Pkg { names: base.names.clone(), parts: base.parts.clone() };
        pkg.set(&main, format!("{}{}{}{}", body.head, copy.body, body.sect, body.tail));
        for (id, xml) in &copy.parts {
            if let Some((_, _, part)) = hf.iter().find(|h| &h.0 == id) {
                pkg.set(part, xml.clone());
            }
        }
        strip_mail_merge_settings(&mut pkg, &main);
        out.push(MergedDocument { record: copy.record, bytes: pkg.write() });
    }
    Ok(out)
}

/// A merge field as Word writes one: a complex field `MERGEFIELD name`
/// showing «name» until a merge fills it, every run in `rpr`. A space in the
/// name is written as `_`, as Word names such a column.
pub fn merge_field_xml(name: &str, rpr: &str) -> String {
    let name: String = name.trim().chars().filter(|&c| c != '"').map(|c| if c.is_whitespace() { '_' } else { c }).collect();
    let name = escape(&name);
    format!(
        "<w:r>{rpr}<w:fldChar w:fldCharType=\"begin\"/></w:r>\
         <w:r>{rpr}<w:instrText xml:space=\"preserve\"> MERGEFIELD {name} </w:instrText></w:r>\
         <w:r>{rpr}<w:fldChar w:fldCharType=\"separate\"/></w:r>\
         <w:r>{rpr}<w:t>\u{ab}{name}\u{bb}</w:t></w:r>\
         <w:r>{rpr}<w:fldChar w:fldCharType=\"end\"/></w:r>"
    )
}

/// One piece of a paragraph an editor writes back: text, or a merge field.
#[derive(Debug, Clone, Default, PartialEq, serde::Deserialize)]
pub struct ParagraphInline {
    #[serde(default)]
    pub text: String,
    #[serde(default)]
    pub merge_field: Option<String>,
}

/// Rewrite a paragraph's runs as `inlines`, keeping the paragraph's own
/// start tag and properties, every run in the formatting of its first text
/// run. A merge field is written as Word writes one ([`merge_field_xml`]).
pub(crate) fn paragraph_with_inlines(p_xml: &str, inlines: &[ParagraphInline], text_xml: fn(&str) -> String) -> String {
    let w = w_prefix(p_xml);
    let w = if w.is_empty() { "w".to_string() } else { w };
    let els = scan(p_xml);
    let Some(p) = els.first() else { return p_xml.to_string() };
    let open = if p.end == p.open_end {
        format!("{}>", p_xml[p.start..p.open_end].trim_end_matches("/>").trim_end())
    } else {
        p_xml[p.start..p.open_end].to_string()
    };
    let ppr = p
        .children
        .iter()
        .find(|&&c| is_w(&els, c, &w, "pPr"))
        .map(|&c| p_xml[els[c].start..els[c].end].to_string())
        .unwrap_or_default();
    let mut items = Vec::new();
    collect_items(&els, 0, &w, &mut items);
    let rpr = items
        .iter()
        .find_map(|it| match *it {
            Item::RunChild { run, child } if is_w(&els, child, &w, "t") => Some(rpr_of(p_xml, &els, run, &w)),
            _ => None,
        })
        .or_else(|| {
            items.iter().find_map(|it| match *it {
                Item::RunChild { run, .. } => Some(rpr_of(p_xml, &els, run, &w)),
                _ => None,
            })
        })
        .unwrap_or_default();
    let mut body = String::new();
    for inline in inlines {
        match inline.merge_field.as_deref().filter(|n| !n.trim().is_empty()) {
            Some(name) => body.push_str(&merge_field_xml(name, &rpr)),
            None if inline.text.is_empty() => {}
            None => body.push_str(&format!("<w:r>{rpr}{}</w:r>", text_xml(&inline.text))),
        }
    }
    if body.is_empty() {
        body = format!("<w:r>{rpr}{}</w:r>", text_xml(""));
    }
    format!("{open}{ppr}{body}</{}>", p.name)
}

/// Field names the main document asks for that the data does not have.
pub fn missing_fields(docx: &[u8], data: &MergeData) -> Result<Vec<String>, MergeError> {
    let report = merge_fields(docx)?;
    let have: Vec<String> = data.columns.iter().map(|c| name_key(c)).collect();
    Ok(report.fields.into_iter().filter(|f| !have.contains(&name_key(f))).collect())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn data() -> MergeData {
        MergeData {
            columns: vec!["Name".into(), "氏名".into(), "住所".into(), "Amount".into(), "Kind".into(), "Blank".into(), "Date".into(), "First Name".into()],
            rows: vec![
                vec!["alice smith".into(), "山田 太郎".into(), "東京都, 港区".into(), "1234.5".into(), "A".into(), "".into(), "2026/01/05".into(), "Ann".into()],
                vec!["bob".into(), "佐藤 花子".into(), "".into(), "7".into(), "B".into(), "".into(), "2026/12/31".into(), "Bo".into()],
                vec!["CAROL ann".into(), "\"引用\" 次郎".into(), "大阪府".into(), "-30".into(), "a".into(), "x".into(), "".into(), "".into()],
            ],
        }
    }

    fn fld(code: &str, result: &str, code_rpr: &str, res_rpr: &str) -> String {
        let rp = if code_rpr.is_empty() { String::new() } else { format!("<w:rPr>{code_rpr}</w:rPr>") };
        let rr = if res_rpr.is_empty() { String::new() } else { format!("<w:rPr>{res_rpr}</w:rPr>") };
        format!(
            "<w:r>{rp}<w:fldChar w:fldCharType=\"begin\"/></w:r><w:r>{rp}<w:instrText xml:space=\"preserve\">{code}</w:instrText></w:r><w:r>{rp}<w:fldChar w:fldCharType=\"separate\"/></w:r><w:r>{rr}<w:t>{result}</w:t></w:r><w:r>{rp}<w:fldChar w:fldCharType=\"end\"/></w:r>"
        )
    }

    fn text_of(xml: &str) -> Vec<String> {
        let els = scan(xml);
        let w = "w";
        (0..els.len())
            .filter(|&i| is_w(&els, i, w, "p"))
            .map(|p| {
                let mut s = String::new();
                fn walk(xml: &str, els: &[El], n: usize, s: &mut String) {
                    for &c in &els[n].children {
                        match local(&els[c].name) {
                            "t" => s.push_str(&unescape(inner(xml, &els[c]))),
                            "br" => s.push('\n'),
                            _ => walk(xml, els, c, s),
                        }
                    }
                }
                walk(xml, &els, p, &mut s);
                s
            })
            .collect()
    }

    fn merge_one(body: &str, rec: usize) -> String {
        let doc = format!("<w:document xmlns:w=\"{W_NS}\"><w:body>{body}</w:body></w:document>");
        let d = data();
        let mut ctx = Ctx::new(&d, rec, rec + 1);
        merge_part(&doc, "w", &mut ctx, &MergeOptions { suppress_blank_lines: false, ..Default::default() }, true).xml
    }

    #[test]
    fn fields_take_values_switches_and_conditions_as_word_does() {
        // Word's own merge of the same fields (main_merged.docx, 2026-11-26).
        let body = [
            format!("<w:p><w:r><w:t xml:space=\"preserve\">P1 </w:t></w:r>{}<w:r><w:t xml:space=\"preserve\"> 様</w:t></w:r></w:p>", fld(" MERGEFIELD 氏名 ", "«氏名»", "", "")),
            format!("<w:p>{}</w:p>", fld(" MERGEFIELD 住所 \\b \"〒\" \\f \"!\" ", "«住所»", "", "")),
            format!("<w:p>{}|{}|{}|{}</w:p>", fld(" MERGEFIELD name \\* Upper ", "x", "", ""), fld(" MERGEFIELD Name \\* Caps ", "x", "", ""), fld(" MERGEFIELD Name \\* FirstCap ", "x", "", ""), fld(" MERGEFIELD Name \\* Lower ", "x", "", "")),
            format!("<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> IF </w:instrText></w:r>{}<w:r><w:instrText xml:space=\"preserve\"> = \"A\" \"会員\" \"一般\" </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>会員</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>", fld(" MERGEFIELD Kind ", "«Kind»", "", "")),
            format!("<w:p>{}|{}</w:p>", fld(" MERGEFIELD Amount \\# \"#,##0\" ", "x", "", ""), fld(" MERGEFIELD Amount ", "x", "", "")),
            format!("<w:p>{}|{}</w:p>", fld(" MERGEREC ", "x", "", ""), fld(" MERGESEQ ", "x", "", "")),
            format!("<w:p>[{}]</w:p>", fld(" MERGEFIELD Blank \\b \"X\" \\f \"Y\" ", "x", "", "")),
            format!("<w:p>{}</w:p>", fld(" MERGEFIELD Date \\@ \"yyyy年M月d日\" ", "x", "", "")),
            format!("<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> IF </w:instrText></w:r>{}<w:r><w:instrText xml:space=\"preserve\"> = \"b*\" \"B-name\" \"other\" </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>x</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>", fld(" MERGEFIELD Name ", "«Name»", "", "")),
            "<w:p><w:fldSimple w:instr=\" MERGEFIELD 氏名 \"><w:r><w:t>«氏名»</w:t></w:r></w:fldSimple><w:r><w:t xml:space=\"preserve\"> simple</w:t></w:r></w:p>".to_string(),
            format!("<w:p>{}</w:p>", fld(" MERGEFIELD First_Name ", "x", "", "")),
        ]
        .concat();
        // The literal `|`/`[` text above is not in runs; wrap them.
        let body = body.replace("|", "<w:r><w:t>|</w:t></w:r>").replace("<w:p>[", "<w:p><w:r><w:t>[</w:t></w:r>").replace("]</w:p>", "<w:r><w:t>]</w:t></w:r></w:p>");
        let want = [
            vec!["P1 山田 太郎 様", "〒東京都, 港区!", "ALICE SMITH|Alice Smith|Alice smith|alice smith", "会員", "1,235|1234.5", "1|1", "[]", "2026年1月5日", "other", "山田 太郎 simple", "Ann"],
            vec!["P1 佐藤 花子 様", "", "BOB|Bob|Bob|bob", "一般", "   7|7", "2|2", "[]", "2026年12月31日", "B-name", "佐藤 花子 simple", "Bo"],
            vec!["P1 \"引用\" 次郎 様", "〒大阪府!", "CAROL ANN|CAROL Ann|CAROL ann|carol ann", "一般", "-  30|-30", "3|3", "[XxY]", "", "other", "\"引用\" 次郎 simple", ""],
        ];
        for (rec, want) in want.iter().enumerate() {
            let got = text_of(&merge_one(&body, rec));
            assert_eq!(&got, want, "record {}", rec + 1);
        }
    }

    #[test]
    fn result_formatting_follows_the_code_or_mergeformat() {
        let body = format!(
            "<w:p>{}{}</w:p>",
            fld(" MERGEFIELD Name \\* MERGEFORMAT ", "«Name»", "<w:b/>", "<w:i/>"),
            fld(" MERGEFIELD Name ", "«Name»", "<w:b/>", "<w:i/>")
        );
        let xml = merge_one(&body, 1);
        assert!(xml.contains("<w:r><w:rPr><w:i/><w:noProof/></w:rPr><w:t>bob</w:t></w:r><w:r><w:rPr><w:b/><w:noProof/></w:rPr><w:t>bob</w:t></w:r>"), "{xml}");
        assert!(!xml.contains("fldChar"), "{xml}");
    }

    #[test]
    fn number_pictures_match_word() {
        // (picture, [values…]) → Word's results, measured 2026-11-26 (fmtn_merged.docx).
        let vals = ["1234.5", "7", "-30", "0", "0.125", "1234567.891", "12abc", " 42 ", "1,234", "-0.004"];
        let table: &[(&str, [&str; 10])] = &[
            ("0", ["1235", "7", "-30", "0", "0", "1234568", "12", "42", "1234", "0"]),
            ("0.00", ["1234.50", "7.00", "-30.00", "0.00", "0.13", "1234567.89", "12.00", "42.00", "1234.00", "0.00"]),
            ("#,##0", ["1,235", "   7", "-  30", "   0", "   0", "1,234,568", "  12", "  42", "1,234", "   0"]),
            ("#,##0.00", ["1,234.50", "   7.00", "-  30.00", "   0.00", "   0.13", "1,234,567.89", "  12.00", "  42.00", "1,234.00", "   0.00"]),
            ("¥#,##0", ["¥1,235", "¥   7", "¥-  30", "¥   0", "¥   0", "¥1,234,568", "¥  12", "¥  42", "¥1,234", "¥   0"]),
            ("0%", ["1235%", "7%", "-30%", "0%", "0%", "1234568%", "12%", "42%", "1234%", "0%"]),
            ("#,##0;(#,##0)", ["1,235", "   7", "(  30)", "   0", "   0", "1,234,568", "  12", "  42", "1,234", "(   0)"]),
            ("00000", ["01235", "00007", "-00030", "00000", "00000", "1234568", "00012", "00042", "01234", "00000"]),
            ("#.##", ["1234.5 ", "7.  ", "-30.  ", " .0 ", " .13", "1234567.89", "12.  ", "42.  ", "1234.  ", " .0 "]),
            ("$###", ["$1235", "$  7", "$- 30", "$   ", "$   ", "$1234568", "$ 12", "$ 42", "$1234", "$   "]),
            ("x##", ["235", "  7", " -30", "   ", "   ", "568", " 12", " 42", "234", "   "]),
            ("#,##0円", ["1,235円", "   7円", "-  30円", "   0円", "   0円", "1,234,568円", "  12円", "  42円", "1,234円", "   0円"]),
            ("'Total' 0", ["Total 1235", "Total 7", "Total -30", "Total 0", "Total 0", "Total 1234568", "Total 12", "Total 42", "Total 1234", "Total 0"]),
            ("+0;-0", ["+1235", "+7", "-30", " 0", " 0", "+1234568", "+12", "+42", "+1234", " 0"]),
            ("#,###", ["1,235", "   7", "-  30", "    ", "    ", "1,234,568", "  12", "  42", "1,234", "    "]),
            ("##0.0#", ["1234.5 ", "  7.0 ", "- 30.0 ", "  0.0 ", "  0.13", "1234567.89", " 12.0 ", " 42.0 ", "1234.0 ", "  0.0 "]),
        ];
        for (pic, want) in table {
            for (v, w) in vals.iter().zip(want.iter()) {
                assert_eq!(format_number(v, pic), *w, "{v} with {pic}");
            }
        }
        // A three-part picture whose zero part is empty (Word reads `"#,##0;-#,##0;"zero""` so).
        assert_eq!(format_number("0", "#,##0;-#,##0;"), "");
        assert_eq!(format_number("-0.004", "#,##0;-#,##0;"), "    0");
        assert_eq!(format_number("abc", "#,##0"), "abc");
    }

    #[test]
    fn date_pictures_match_word() {
        // Measured 2026-11-26 (fmtd_merged.docx); a value that is not a date stays as it is.
        let vals = ["2026/01/05", "2026-12-31", "1/5/2026", "2026年3月4日", "abc", "2026/1/5 13:07:09", "45658", ""];
        let table: &[(&str, [&str; 8])] = &[
            ("yyyy/MM/dd", ["2026/01/05", "2026/12/31", "2026/01/05", "2026/03/04", "abc", "2026/01/05", "45658", ""]),
            ("yyyy年M月d日", ["2026年1月5日", "2026年12月31日", "2026年1月5日", "2026年3月4日", "abc", "2026年1月5日", "45658", ""]),
            ("M/d/yy", ["1/5/26", "12/31/26", "1/5/26", "3/4/26", "abc", "1/5/26", "45658", ""]),
            ("dddd, MMMM d, yyyy", ["Monday, January 5, 2026", "Thursday, December 31, 2026", "Monday, January 5, 2026", "Wednesday, March 4, 2026", "abc", "Monday, January 5, 2026", "45658", ""]),
            ("ddd MMM", ["Mon Jan", "Thu Dec", "Mon Jan", "Wed Mar", "abc", "Mon Jan", "45658", ""]),
            ("ggge年M月d日", ["令和8年1月5日", "令和8年12月31日", "令和8年1月5日", "令和8年3月4日", "abc", "令和8年1月5日", "45658", ""]),
            ("aaa", ["月", "木", "月", "水", "abc", "月", "45658", ""]),
            ("aaaa", ["月a", "木a", "月a", "水a", "abc", "月a", "45658", ""]),
            ("H:mm", ["0:00", "0:00", "0:00", "0:00", "abc", "13:07", "45658", ""]),
            ("hh:mm AM/PM", ["12:00 AM", "12:00 AM", "12:00 AM", "12:00 AM", "abc", "01:07 PM", "45658", ""]),
            ("'d:' d", ["d: 5", "d: 31", "d: 5", "d: 4", "abc", "d: 5", "45658", ""]),
            ("yy-M-d", ["26-1-5", "26-12-31", "26-1-5", "26-3-4", "abc", "26-1-5", "45658", ""]),
            ("MMMM", ["January", "December", "January", "March", "abc", "January", "45658", ""]),
        ];
        for (pic, want) in table {
            for (v, w) in vals.iter().zip(want.iter()) {
                assert_eq!(format_date(v, pic), *w, "{v} with {pic}");
            }
        }
        assert_eq!(format_date("2001/08/04", "ggge年M月d日"), "平成13年8月4日");
        // Word reads "R8.4.1" as 8/4/01 (fmtd_merged.docx, record 5).
        assert_eq!(format_date("R8.4.1", "yyyy/MM/dd"), "2001/08/04");
        assert_eq!(format_date("R8.4.1", "dddd"), "Saturday");
    }

    #[test]
    fn delimited_text_reads_quotes_tabs_and_line_breaks() {
        let d = parse_delimited("\u{feff}Name,住所,Note\r\n\"Smith, Ann\",\"東京都 \"\"港区\"\"\",\"two\nlines\"\r\nBob,,\r\n\r\n");
        assert_eq!(d.columns, vec!["Name", "住所", "Note"]);
        assert_eq!(d.rows, vec![vec!["Smith, Ann", "東京都 \"港区\"", "two\nlines"], vec!["Bob", "", ""]]);
        let t = parse_delimited("a\tb\n1,5\t2\n");
        assert_eq!(t.columns, vec!["a", "b"]);
        assert_eq!(t.rows, vec![vec!["1,5", "2"]]);
    }

    #[test]
    fn skipif_drops_a_record_and_next_pulls_in_the_next() {
        let body = format!(
            "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> SKIPIF </w:instrText></w:r>{}<w:r><w:instrText xml:space=\"preserve\"> = \"B\" </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r><w:r><w:t xml:space=\"preserve\">S1 </w:t></w:r>{}</w:p>",
            fld(" MERGEFIELD Kind ", "«Kind»", "", ""),
            fld(" MERGEFIELD Name ", "«Name»", "", "")
        );
        let docx = test_docx(&body, "");
        let each = mail_merge_each(&docx, &data(), &MergeOptions::default()).unwrap();
        let texts: Vec<Vec<String>> = each.iter().map(|d| text_of(&Pkg::read(&d.bytes).unwrap().text("word/document.xml").unwrap())).collect();
        assert_eq!(texts, vec![vec!["S1 alice smith".to_string()], vec!["S1 CAROL ann".to_string()]]);
        assert_eq!(each.iter().map(|d| d.record).collect::<Vec<_>>(), vec![1, 3]);

        let labels = format!(
            "<w:p><w:r><w:t xml:space=\"preserve\">L1 </w:t></w:r>{}</w:p><w:p>{}<w:r><w:t xml:space=\"preserve\">L2 </w:t></w:r>{}</w:p>",
            fld(" MERGEFIELD Name ", "x", "", ""),
            fld(" NEXT ", "", "", ""),
            fld(" MERGEFIELD Name ", "x", "", "")
        );
        let merged = mail_merge(&test_docx(&labels, ""), &data(), &MergeOptions::default()).unwrap();
        let doc = Pkg::read(&merged).unwrap().text("word/document.xml").unwrap();
        // Word: "L1 alice smith","L2 bob" | "L1 CAROL ann","L2 " (+ the closing empty paragraph).
        assert_eq!(text_of(&doc), vec!["L1 alice smith", "L2 bob", "L1 CAROL ann", "L2 ", ""]);
    }

    fn test_docx(body: &str, header: &str) -> Vec<u8> {
        let mut pkg = Pkg { names: vec![], parts: HashMap::new() };
        let hdr_rel = if header.is_empty() { "" } else { "<Relationship Id=\"rIdH1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>" };
        let hdr_ref = if header.is_empty() { "" } else { "<w:headerReference w:type=\"default\" r:id=\"rIdH1\"/>" };
        pkg.set("[Content_Types].xml", format!("<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>{}</Types>", if header.is_empty() { "" } else { "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>" }));
        pkg.set("_rels/.rels", "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>".into());
        pkg.set("word/_rels/document.xml.rels", format!("<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{hdr_rel}</Relationships>"));
        pkg.set("word/document.xml", format!("<?xml version=\"1.0\"?><w:document xmlns:w=\"{W_NS}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>{body}<w:sectPr>{hdr_ref}<w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>"));
        if !header.is_empty() {
            pkg.set("word/header1.xml", format!("<?xml version=\"1.0\"?><w:hdr xmlns:w=\"{W_NS}\">{header}</w:hdr>"));
        }
        pkg.write()
    }

    #[test]
    fn one_document_has_a_section_per_record_and_headers_of_its_own() {
        let body = format!(
            "<w:p><w:bookmarkStart w:id=\"0\" w:name=\"mark1\"/><w:r><w:t xml:space=\"preserve\">To </w:t></w:r>{}<w:bookmarkEnd w:id=\"0\"/></w:p>",
            fld(" MERGEFIELD 氏名 ", "«氏名»", "", "")
        );
        let header = format!("<w:p><w:r><w:t xml:space=\"preserve\">H: </w:t></w:r>{}</w:p>", fld(" MERGEFIELD 氏名 ", "«氏名»", "", ""));
        let out = mail_merge(&test_docx(&body, &header), &data(), &MergeOptions::default()).unwrap();
        let pkg = Pkg::read(&out).unwrap();
        let doc = pkg.text("word/document.xml").unwrap();
        assert_eq!(text_of(&doc), vec!["To 山田 太郎", "To 佐藤 花子", "To \"引用\" 次郎", ""]);
        assert!(!doc.contains("bookmarkStart"), "{doc}");
        // Three record sections restarting page numbers, then the continuous one.
        assert_eq!(doc.matches("<w:pgNumType w:start=\"1\"/>").count(), 3, "{doc}");
        assert_eq!(doc.matches("<w:sectPr>").count(), 4);
        assert!(doc.contains("<w:headerReference w:type=\"default\" r:id=\"rIdMerge3rIdH1\"/><w:type w:val=\"continuous\"/><w:pgSz"), "{doc}");
        for (n, name) in ["山田 太郎", "佐藤 花子", "\"引用\" 次郎"].iter().enumerate() {
            let h = pkg.text(&format!("word/header1_merge{}.xml", n + 1)).unwrap();
            assert_eq!(text_of(&h), vec![format!("H: {name}")]);
            assert!(doc.contains(&format!("r:id=\"rIdMerge{}rIdH1\"", n + 1)));
        }
        let ct = pkg.text("[Content_Types].xml").unwrap();
        assert!(ct.contains("PartName=\"/word/header1_merge2.xml\""), "{ct}");
        let rels = pkg.text("word/_rels/document.xml.rels").unwrap();
        assert!(rels.contains("Id=\"rIdMerge2rIdH1\"") && rels.contains("Target=\"header1_merge2.xml\""), "{rels}");
        assert_eq!(merge_fields(&test_docx(&body, &header)).unwrap().fields, vec!["氏名"]);
    }

    #[test]
    fn blank_lines_go_when_a_field_leaves_nothing() {
        // Word with SuppressBlankLines (fmtn_merged1.docx): the three paragraphs
        // holding only empty fields (and spaces) go; an empty paragraph stays.
        let body = format!(
            "<w:p>{0}</w:p><w:p><w:r><w:t xml:space=\"preserve\">  </w:t></w:r>{0}<w:r><w:t xml:space=\"preserve\"> </w:t></w:r></w:p><w:p>{0}{1}</w:p><w:p/><w:p><w:r><w:t>END</w:t></w:r></w:p>",
            fld(" MERGEFIELD Blank ", "«Blank»", "", ""),
            fld(" MERGEFIELD Blank \\b \"x\" ", "«Blank»", "", "")
        );
        let out = mail_merge_each(&test_docx(&body, ""), &data(), &MergeOptions { last_record: Some(1), ..Default::default() }).unwrap();
        let doc = Pkg::read(&out[0].bytes).unwrap().text("word/document.xml").unwrap();
        assert_eq!(text_of(&doc), vec!["", "END"]);
        let kept = mail_merge_each(&test_docx(&body, ""), &data(), &MergeOptions { last_record: Some(1), suppress_blank_lines: false, ..Default::default() }).unwrap();
        let doc = Pkg::read(&kept[0].bytes).unwrap().text("word/document.xml").unwrap();
        assert_eq!(text_of(&doc), vec!["", "   ", "", "", "END"]);
    }

    #[test]
    fn an_editors_field_is_written_as_word_writes_it_and_read_back() {
        let blocks: Vec<crate::ContentBlock> = serde_json::from_str(
            r#"[{"type":"paragraph","runs":[{"text":"Dear "},{"text":"«First Name»","merge_field":"First Name","bold":true},{"text":","}]}]"#,
        )
        .unwrap();
        let docx = crate::build_docx(&blocks);
        let doc = Pkg::read(&docx).unwrap().text("word/document.xml").unwrap();
        assert!(doc.contains("<w:instrText xml:space=\"preserve\"> MERGEFIELD First_Name </w:instrText>"), "{doc}");
        assert_eq!(merge_fields(&docx).unwrap().fields, vec!["First_Name"]);
        // The parser marks the field's result run with its name.
        let parsed = serde_json::to_string(&crate::parser::parse_docx(&docx).unwrap()).unwrap();
        assert!(parsed.contains("\"merge_field\":\"First_Name\""), "{parsed}");
        // And it merges, in the field's own (bold) formatting.
        let out = mail_merge_each(&docx, &data(), &MergeOptions::default()).unwrap();
        let xml = Pkg::read(&out[0].bytes).unwrap().text("word/document.xml").unwrap();
        assert_eq!(text_of(&xml), vec!["Dear Ann,"]);
        assert!(xml.contains("<w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr><w:t>Ann</w:t>"), "{xml}");
    }

    #[test]
    fn a_paragraph_written_again_keeps_its_properties_and_takes_fields() {
        let p = "<w:p w:rsidR=\"00A1\"><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>old</w:t></w:r><w:r><w:t>more</w:t></w:r></w:p>";
        let inlines = vec![
            ParagraphInline { text: "To ".into(), merge_field: None },
            ParagraphInline { text: "«氏名»".into(), merge_field: Some("氏名".into()) },
            ParagraphInline { text: " 様".into(), merge_field: None },
        ];
        let out = paragraph_with_inlines(p, &inlines, |t| format!("<w:t xml:space=\"preserve\">{}</w:t>", escape(t)));
        assert!(out.starts_with("<w:p w:rsidR=\"00A1\"><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t xml:space=\"preserve\">To </w:t></w:r>"), "{out}");
        assert!(out.contains(" MERGEFIELD 氏名 ") && out.ends_with("</w:p>"), "{out}");
        assert_eq!(text_of(&out), vec!["To «氏名» 様"]);
        let mut editor = crate::DocxEditor::new(&test_docx(p, "")).unwrap();
        editor.add_edit(crate::editor::DocxEdit::SetParagraphInlines { paragraph_index: 0, inlines });
        let saved = editor.save().unwrap();
        assert_eq!(merge_fields(&saved).unwrap().fields, vec!["氏名"]);
    }

    /// Merge files named by the environment (OXI_MM_MAIN, OXI_MM_DATA as
    /// CSV, OXI_MM_OUT), to put a merge next to Word's own.
    #[test]
    #[ignore]
    fn merge_files_from_env() {
        let main = std::fs::read(std::env::var("OXI_MM_MAIN").unwrap()).unwrap();
        let data = parse_delimited(&std::fs::read_to_string(std::env::var("OXI_MM_DATA").unwrap()).unwrap());
        let opts = MergeOptions { suppress_blank_lines: std::env::var("OXI_MM_SB").is_ok(), ..Default::default() };
        let out = mail_merge(&main, &data, &opts).unwrap();
        std::fs::write(std::env::var("OXI_MM_OUT").unwrap(), out).unwrap();
    }

    #[test]
    fn a_field_inside_one_run_and_a_main_documents_data_link() {
        let body = "<w:p><w:r><w:t xml:space=\"preserve\">A </w:t><w:fldChar w:fldCharType=\"begin\"/><w:instrText> MERGEFIELD Name </w:instrText><w:fldChar w:fldCharType=\"separate\"/><w:t>x</w:t><w:fldChar w:fldCharType=\"end\"/><w:t xml:space=\"preserve\"> B</w:t></w:r></w:p>";
        let xml = merge_one(body, 0);
        assert_eq!(text_of(&xml), vec!["A alice smith B"]);
        assert!(xml.contains("<w:r><w:t xml:space=\"preserve\">A </w:t></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t xml:space=\"preserve\">alice smith</w:t></w:r><w:r><w:t xml:space=\"preserve\"> B</w:t></w:r>") || xml.contains("<w:t>alice smith</w:t>"), "{xml}");
    }
}
