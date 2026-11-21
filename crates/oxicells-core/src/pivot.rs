// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Pivot tables: what a workbook's pivot definitions say, and the grid Excel
//! fills the pivot's range with when it refreshes it from its source.
//!
//! The definition lives in three parts: `pivotTableDefinition` (where the
//! table stands, which fields sit on rows, columns, pages and values, and in
//! what order each field's items come), `pivotCacheDefinition` (the source
//! range and each field's distinct items) and the workbook's `pivotCaches`
//! list that joins the two by `cacheId`. The grid is computed afresh from the
//! source cells, so a refresh after the data changes gives Excel's answer and
//! not the cached one.
//!
//! Layout is Excel's default compact form: row fields stacked in one column,
//! each parent item on its own row carrying its subtotal above its children,
//! a grand total row and column, captions in the UI language. Measured
//! against Excel's own refreshed pivots (`tools/metrics/_xlsx_pivot_gen.py`).

use std::collections::HashMap;
use std::io::Read;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::ir::{CellValue, Workbook};

/// How a value field combines the source cells that land on one pivot cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Aggregate {
    Sum,
    Count,
    Average,
    Max,
    Min,
    Product,
    CountNums,
    StdDev,
    StdDevP,
    Var,
    VarP,
}

impl Aggregate {
    fn from_attr(value: Option<&str>) -> Self {
        match value.unwrap_or("sum") {
            "count" => Self::Count,
            "average" => Self::Average,
            "max" => Self::Max,
            "min" => Self::Min,
            "product" => Self::Product,
            "countNums" => Self::CountNums,
            "stdDev" => Self::StdDev,
            "stdDevp" => Self::StdDevP,
            "var" => Self::Var,
            "varp" => Self::VarP,
            _ => Self::Sum,
        }
    }
}

#[derive(Debug, Clone)]
pub struct DataField {
    pub name: String,
    pub field: usize,
    pub aggregate: Aggregate,
}

/// How the row fields are laid out: Excel's three report layouts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RowLayout {
    /// All row fields in one column, subtotals on the parent's row.
    Compact,
    /// A column per field, subtotals on the parent's row.
    Outline,
    /// A column per field, the parent beside its first child and its
    /// subtotal on a row of its own under the group.
    Tabular,
}

/// One pivot table, as its parts define it.
#[derive(Debug, Clone)]
pub struct PivotSpec {
    pub name: String,
    /// The sheet the table stands on, as the workbook names it.
    pub sheet: String,
    /// Top-left of `location ref`, zero-based (row, column).
    pub origin: (u32, u32),
    /// Bottom-right of `location ref` as last saved, zero-based: with the page
    /// rows above `origin`, the range a refresh clears before writing.
    pub extent: (u32, u32),
    pub source_sheet: String,
    /// Source range, zero-based inclusive (r0, c0, r1, c1), header row first.
    pub source: (u32, u32, u32, u32),
    pub field_names: Vec<String>,
    /// Each field's items in the order the table shows them, as cache values.
    pub item_order: Vec<Vec<CellValue>>,
    /// Items a field hides.
    pub hidden: Vec<Vec<CellValue>>,
    /// Row fields; `None` stands for the values pseudo-field (`x="-2"`).
    pub rows: Vec<Option<usize>>,
    /// Column fields; `None` stands for the values pseudo-field (`x="-2"`).
    pub cols: Vec<Option<usize>>,
    /// Page fields, each with the one item it shows when it shows one.
    pub pages: Vec<(usize, Option<CellValue>)>,
    pub data: Vec<DataField>,
    pub row_header_caption: Option<String>,
    pub col_header_caption: Option<String>,
    pub grand_total_caption: Option<String>,
    pub data_caption: Option<String>,
    pub row_grand_totals: bool,
    pub col_grand_totals: bool,
    pub row_layout: RowLayout,
}

/// The words a pivot shows that its file does not spell out, by UI language.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Ui {
    Japanese,
    English,
}

impl Ui {
    fn row_labels(self) -> &'static str {
        match self {
            Ui::Japanese => "行ラベル",
            Ui::English => "Row Labels",
        }
    }
    fn col_labels(self) -> &'static str {
        match self {
            Ui::Japanese => "列ラベル",
            Ui::English => "Column Labels",
        }
    }
    fn grand_total(self) -> &'static str {
        match self {
            Ui::Japanese => "総計",
            Ui::English => "Grand Total",
        }
    }
    fn all(self) -> &'static str {
        match self {
            Ui::Japanese => "(すべて)",
            Ui::English => "(All)",
        }
    }
    fn blank(self) -> &'static str {
        match self {
            Ui::Japanese => "(空白)",
            Ui::English => "(blank)",
        }
    }
    fn subtotal(self, item: &str) -> String {
        match self {
            Ui::Japanese => format!("{item} 集計"),
            Ui::English => format!("{item} Total"),
        }
    }
    fn data_total(self, name: &str) -> String {
        match self {
            Ui::Japanese => format!("全体の {name}"),
            Ui::English => format!("Total {name}"),
        }
    }
}

/// A pivot's computed range: the cell values from `origin` (zero-based), one
/// row of the grid per sheet row, with `Empty` where Excel leaves a blank.
#[derive(Debug, Clone, PartialEq)]
pub struct PivotGrid {
    pub origin: (u32, u32),
    pub cells: Vec<Vec<CellValue>>,
}

fn attr(e: &BytesStart<'_>, name: &str) -> Option<String> {
    e.attributes().flatten().find(|a| {
        let key = a.key.as_ref();
        key == name.as_bytes() || key.rsplit(|b| *b == b':').next() == Some(name.as_bytes())
    }).and_then(|a| a.unescape_value().ok().map(|v| v.into_owned()))
}

fn local(name: &[u8]) -> String {
    String::from_utf8_lossy(name.rsplit(|b| *b == b':').next().unwrap_or(name)).into_owned()
}

fn read_part(zip: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>, path: &str) -> Option<String> {
    let mut file = zip.by_name(path).ok()?;
    let mut out = String::new();
    file.read_to_string(&mut out).ok()?;
    Some(out)
}

/// `../pivotTables/pivotTable1.xml` against `xl/worksheets/` -> `xl/pivotTables/pivotTable1.xml`.
fn resolve(base_dir: &str, target: &str) -> String {
    if let Some(rest) = target.strip_prefix('/') {
        return rest.to_string();
    }
    let mut parts: Vec<&str> = base_dir.split('/').filter(|p| !p.is_empty()).collect();
    for piece in target.split('/') {
        match piece {
            ".." => {
                parts.pop();
            }
            "." | "" => {}
            other => parts.push(other),
        }
    }
    parts.join("/")
}

fn rels(xml: &str) -> HashMap<String, (String, String)> {
    let mut out = HashMap::new();
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Empty(e) | Event::Start(e)) if local(e.name().as_ref()) == "Relationship" => {
                if let (Some(id), Some(target)) = (attr(&e, "Id"), attr(&e, "Target")) {
                    out.insert(id, (target, attr(&e, "Type").unwrap_or_default()));
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    out
}

/// "A3:E9" -> zero-based (r0, c0, r1, c1).
fn range(reference: &str) -> Option<(u32, u32, u32, u32)> {
    let one = |cell: &str| -> Option<(u32, u32)> {
        let cell = cell.replace('$', "");
        let letters: String = cell.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
        let digits = &cell[letters.len()..];
        let mut col = 0u32;
        for c in letters.chars() {
            col = col * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        }
        Some((digits.parse::<u32>().ok()?.checked_sub(1)?, col.checked_sub(1)?))
    };
    let (a, b) = reference.split_once(':').unwrap_or((reference, reference));
    let (r0, c0) = one(a)?;
    let (r1, c1) = one(b)?;
    Some((r0.min(r1), c0.min(c1), r0.max(r1), c0.max(c1)))
}

struct Cache {
    source_sheet: String,
    source: (u32, u32, u32, u32),
    names: Vec<String>,
    shared: Vec<Vec<CellValue>>,
}

fn parse_cache(xml: &str) -> Option<Cache> {
    let mut reader = Reader::from_str(xml);
    let mut cache = Cache { source_sheet: String::new(), source: (0, 0, 0, 0), names: Vec::new(), shared: Vec::new() };
    let mut in_shared = false;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match local(e.name().as_ref()).as_str() {
                "worksheetSource" => {
                    cache.source_sheet = attr(&e, "sheet").unwrap_or_default();
                    cache.source = range(&attr(&e, "ref")?)?;
                }
                "cacheField" => {
                    cache.names.push(attr(&e, "name").unwrap_or_default());
                    cache.shared.push(Vec::new());
                }
                "sharedItems" => in_shared = true,
                "s" if in_shared => cache.shared.last_mut()?.push(CellValue::String(attr(&e, "v").unwrap_or_default())),
                "n" if in_shared => cache.shared.last_mut()?.push(CellValue::Number(attr(&e, "v").and_then(|v| v.parse().ok()).unwrap_or(0.0))),
                "b" if in_shared => cache.shared.last_mut()?.push(CellValue::Boolean(attr(&e, "v").as_deref() == Some("1"))),
                "m" if in_shared => cache.shared.last_mut()?.push(CellValue::Empty),
                "e" if in_shared => cache.shared.last_mut()?.push(CellValue::Error(attr(&e, "v").unwrap_or_default())),
                _ => {}
            },
            Ok(Event::End(e)) if local(e.name().as_ref()) == "sharedItems" => in_shared = false,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    Some(cache)
}

fn is_true(value: Option<String>, default: bool) -> bool {
    match value.as_deref() {
        Some("1") | Some("true") => true,
        Some("0") | Some("false") => false,
        _ => default,
    }
}

fn parse_table(xml: &str, sheet: &str, cache: &Cache) -> Option<PivotSpec> {
    let mut reader = Reader::from_str(xml);
    let fields = cache.names.len();
    let mut spec = PivotSpec {
        name: String::new(),
        sheet: sheet.to_string(),
        origin: (0, 0),
        extent: (0, 0),
        source_sheet: cache.source_sheet.clone(),
        source: cache.source,
        field_names: cache.names.clone(),
        item_order: vec![Vec::new(); fields],
        hidden: vec![Vec::new(); fields],
        rows: Vec::new(),
        cols: Vec::new(),
        pages: Vec::new(),
        data: Vec::new(),
        row_header_caption: None,
        col_header_caption: None,
        grand_total_caption: None,
        data_caption: None,
        row_grand_totals: true,
        col_grand_totals: true,
        row_layout: RowLayout::Compact,
    };
    let mut field = None::<usize>;
    let mut next_field = 0usize;
    let mut table_compact = true;
    let mut field_layout: Vec<(bool, bool)> = Vec::new();
    let mut area = "";
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match local(e.name().as_ref()).as_str() {
                "pivotTableDefinition" => {
                    spec.name = attr(&e, "name").unwrap_or_default();
                    spec.row_header_caption = attr(&e, "rowHeaderCaption");
                    spec.col_header_caption = attr(&e, "colHeaderCaption");
                    spec.grand_total_caption = attr(&e, "grandTotalCaption");
                    spec.data_caption = attr(&e, "dataCaption");
                    spec.row_grand_totals = is_true(attr(&e, "rowGrandTotals"), true);
                    table_compact = is_true(attr(&e, "compact"), true);
                    spec.col_grand_totals = is_true(attr(&e, "colGrandTotals"), true);
                }
                "location" => {
                    let (r0, c0, r1, c1) = range(&attr(&e, "ref")?)?;
                    spec.origin = (r0, c0);
                    spec.extent = (r1, c1);
                }
                "pivotField" => {
                    field = Some(next_field);
                    next_field += 1;
                    field_layout.push((is_true(attr(&e, "compact"), table_compact), is_true(attr(&e, "outline"), true)));
                }
                "item" => {
                    if let Some(f) = field {
                        if let Some(x) = attr(&e, "x").and_then(|v| v.parse::<usize>().ok()) {
                            if let Some(value) = cache.shared.get(f).and_then(|s| s.get(x)) {
                                spec.item_order[f].push(value.clone());
                                if is_true(attr(&e, "h"), false) {
                                    spec.hidden[f].push(value.clone());
                                }
                            }
                        }
                    }
                }
                "rowFields" => area = "rows",
                "colFields" => area = "cols",
                "pageFields" => area = "pages",
                "field" => {
                    let x = attr(&e, "x").and_then(|v| v.parse::<i64>().ok()).unwrap_or(0);
                    match area {
                        "rows" => spec.rows.push((x >= 0).then_some(x as usize)),
                        "cols" => spec.cols.push((x >= 0).then_some(x as usize)),
                        _ => {}
                    }
                }
                "pageField" => {
                    if let Some(f) = attr(&e, "fld").and_then(|v| v.parse::<usize>().ok()) {
                        // `item` counts the field's own item list, not the cache's.
                        let pick = attr(&e, "item")
                            .and_then(|v| v.parse::<usize>().ok())
                            .and_then(|i| spec.item_order.get(f).and_then(|items| items.get(i)).cloned());
                        spec.pages.push((f, pick));
                    }
                }
                "dataField" => {
                    let f = attr(&e, "fld").and_then(|v| v.parse::<usize>().ok())?;
                    spec.data.push(DataField {
                        name: attr(&e, "name").unwrap_or_else(|| cache.names.get(f).cloned().unwrap_or_default()),
                        field: f,
                        aggregate: Aggregate::from_attr(attr(&e, "subtotal").as_deref()),
                    });
                }
                _ => {}
            },
            Ok(Event::End(e)) => match local(e.name().as_ref()).as_str() {
                "pivotField" => field = None,
                "rowFields" | "colFields" | "pageFields" => area = "",
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    if let Some((compact, outline)) = spec.rows.iter().flatten().next().and_then(|f| field_layout.get(*f)) {
        spec.row_layout = match (compact, outline) {
            (true, _) => RowLayout::Compact,
            (false, true) => RowLayout::Outline,
            (false, false) => RowLayout::Tabular,
        };
    }
    Some(spec)
}

/// Every pivot table a workbook package defines.
pub fn read_pivots(data: &[u8]) -> Vec<PivotSpec> {
    let Ok(mut zip) = zip::ZipArchive::new(std::io::Cursor::new(data)) else {
        return Vec::new();
    };
    let Some(workbook) = read_part(&mut zip, "xl/workbook.xml") else {
        return Vec::new();
    };
    let wb_rels = read_part(&mut zip, "xl/_rels/workbook.xml.rels").map(|x| rels(&x)).unwrap_or_default();
    // cacheId -> cache part; sheet name -> sheet part.
    let mut caches = HashMap::new();
    let mut sheets = Vec::new();
    let mut reader = Reader::from_str(&workbook);
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match local(e.name().as_ref()).as_str() {
                "pivotCache" => {
                    if let (Some(id), Some(rid)) = (attr(&e, "cacheId"), attr(&e, "id")) {
                        if let Some((target, _)) = wb_rels.get(&rid) {
                            caches.insert(id, resolve("xl", target));
                        }
                    }
                }
                "sheet" => {
                    if let (Some(name), Some(rid)) = (attr(&e, "name"), attr(&e, "id")) {
                        if let Some((target, _)) = wb_rels.get(&rid) {
                            sheets.push((name, resolve("xl", target)));
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    let mut out = Vec::new();
    for (name, part) in sheets {
        let (dir, file) = part.rsplit_once('/').unwrap_or(("", &part));
        let Some(sheet_rels) = read_part(&mut zip, &format!("{dir}/_rels/{file}.rels")).map(|x| rels(&x)) else {
            continue;
        };
        for (target, kind) in sheet_rels.values() {
            if !kind.ends_with("/pivotTable") {
                continue;
            }
            let table_path = resolve(dir, target);
            let Some(table_xml) = read_part(&mut zip, &table_path) else { continue };
            let cache_id = {
                let mut reader = Reader::from_str(&table_xml);
                let mut id = None;
                loop {
                    match reader.read_event() {
                        Ok(Event::Start(e)) | Ok(Event::Empty(e)) if local(e.name().as_ref()) == "pivotTableDefinition" => {
                            id = attr(&e, "cacheId");
                            break;
                        }
                        Ok(Event::Eof) | Err(_) => break,
                        _ => {}
                    }
                }
                id
            };
            let Some(cache_part) = cache_id.and_then(|id| caches.get(&id).cloned()) else { continue };
            let Some(cache) = read_part(&mut zip, &cache_part).and_then(|x| parse_cache(&x)) else { continue };
            if let Some(spec) = parse_table(&table_xml, &name, &cache) {
                out.push(spec);
            }
        }
    }
    out
}

/// A key a value is grouped under: text compares case-insensitively, the way
/// Excel folds "east" and "East" into one item.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum Key {
    Blank,
    Number(u64),
    Text(String),
    Bool(bool),
    Error(String),
}

fn key(value: &CellValue) -> Key {
    match value {
        CellValue::Empty => Key::Blank,
        CellValue::Number(n) => Key::Number(n.to_bits()),
        CellValue::String(s) => Key::Text(s.to_lowercase()),
        CellValue::Utf16(u) => Key::Text(String::from_utf16_lossy(u).to_lowercase()),
        CellValue::Boolean(b) => Key::Bool(*b),
        CellValue::Error(e) => Key::Error(e.clone()),
    }
}

const MONTHS_SHORT: [&str; 12] = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
const MONTHS_LONG: [&str; 12] = [
    "january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november",
    "december",
];
const DAYS_SHORT: [&str; 7] = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];
const DAYS_LONG: [&str; 7] = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];

/// Excel's ascending item order: numbers, then text (built-in custom lists
/// -- month and weekday names -- in their own order first), then booleans,
/// errors, and the blank last. `p03`: Jan, Feb, Mar, not Feb, Jan, Mar.
fn ascending(a: &CellValue, b: &CellValue) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    let rank = |v: &CellValue| match v {
        CellValue::Number(_) => 0,
        CellValue::String(_) | CellValue::Utf16(_) => 1,
        CellValue::Boolean(_) => 2,
        CellValue::Error(_) => 3,
        CellValue::Empty => 4,
    };
    let text = |v: &CellValue| match v {
        CellValue::String(s) => s.to_lowercase(),
        CellValue::Utf16(u) => String::from_utf16_lossy(u).to_lowercase(),
        _ => String::new(),
    };
    let listed = |t: &str| -> Option<(usize, usize)> {
        for (list, names) in [(0usize, &MONTHS_SHORT[..]), (1, &MONTHS_LONG[..]), (2, &DAYS_SHORT[..]), (3, &DAYS_LONG[..])] {
            if let Some(i) = names.iter().position(|n| *n == t) {
                return Some((list, i));
            }
        }
        None
    };
    match rank(a).cmp(&rank(b)) {
        Ordering::Equal => {}
        other => return other,
    }
    match (a, b) {
        (CellValue::Number(x), CellValue::Number(y)) => x.partial_cmp(y).unwrap_or(Ordering::Equal),
        (CellValue::Boolean(x), CellValue::Boolean(y)) => x.cmp(y),
        _ => {
            let (ta, tb) = (text(a), text(b));
            match (listed(&ta), listed(&tb)) {
                (Some(x), Some(y)) if x.0 == y.0 => x.1.cmp(&y.1),
                (Some(_), None) => Ordering::Less,
                (None, Some(_)) => Ordering::Greater,
                _ => ta.cmp(&tb),
            }
        }
    }
}

#[derive(Default, Clone)]
struct Acc {
    sum: f64,
    product: f64,
    count: usize,
    nums: usize,
    max: f64,
    min: f64,
    sq: f64,
}

impl Acc {
    fn add(&mut self, value: &CellValue) {
        if matches!(value, CellValue::Empty) {
            return;
        }
        self.count += 1;
        if let CellValue::Number(n) = value {
            if self.nums == 0 {
                self.max = *n;
                self.min = *n;
                self.product = 1.0;
            }
            self.nums += 1;
            self.sum += n;
            self.sq += n * n;
            self.product *= n;
            self.max = self.max.max(*n);
            self.min = self.min.min(*n);
        }
    }
    fn value(&self, how: Aggregate) -> CellValue {
        let n = self.nums as f64;
        let var = |sample: bool| {
            let d = if sample { n - 1.0 } else { n };
            if d <= 0.0 {
                return None;
            }
            Some(((self.sq - self.sum * self.sum / n) / d).max(0.0))
        };
        let number = match how {
            Aggregate::Count => return CellValue::Number(self.count as f64),
            Aggregate::CountNums => return CellValue::Number(n),
            Aggregate::Sum => Some(self.sum),
            Aggregate::Average => (self.nums > 0).then(|| self.sum / n),
            Aggregate::Max => Some(if self.nums > 0 { self.max } else { 0.0 }),
            Aggregate::Min => Some(if self.nums > 0 { self.min } else { 0.0 }),
            Aggregate::Product => Some(if self.nums > 0 { self.product } else { 0.0 }),
            Aggregate::StdDev => var(true).map(f64::sqrt),
            Aggregate::StdDevP => var(false).map(f64::sqrt),
            Aggregate::Var => var(true),
            Aggregate::VarP => var(false),
        };
        match number {
            Some(v) => CellValue::Number(v),
            None => CellValue::Error("#DIV/0!".into()),
        }
    }
}

/// One heading position on an axis: the items it stands for (one per real
/// field above it), the value field it carries when the values pseudo-field
/// sits on this axis, and what kind of heading it is.
#[derive(Debug, Clone)]
struct Slot {
    path: Vec<Key>,
    data: Option<usize>,
    kind: SlotKind,
}

#[derive(Debug, Clone, PartialEq)]
enum SlotKind {
    /// An item at `level` of the axis, labelled `label`.
    Item { level: usize, label: CellValue },
    /// A value field's own line under an item (values on this axis).
    Data { level: usize },
    /// The total of an outer item over the items under it.
    Subtotal { level: usize, label: CellValue },
    /// The grand total (one per value field when values are on the axis).
    Grand,
}

/// The value a source cell holds, read through `get(row, col)` (zero-based).
pub fn compute(spec: &PivotSpec, get: &dyn Fn(u32, u32) -> CellValue, ui: Ui) -> PivotGrid {
    let (r0, c0, r1, c1) = spec.source;
    let width = (c1 - c0 + 1) as usize;
    let all: Vec<Vec<CellValue>> =
        (r0 + 1..=r1).map(|r| (0..width).map(|j| get(r, c0 + j as u32)).collect()).collect();

    let text = |s: &str| CellValue::String(s.to_string());
    let shown = |v: &CellValue| match v {
        CellValue::Empty => text(ui.blank()),
        other => other.clone(),
    };
    let label_text = |v: &CellValue| match v {
        CellValue::Empty => ui.blank().to_string(),
        CellValue::String(s) => s.clone(),
        CellValue::Number(n) => {
            if n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{}", *n as i64)
            } else {
                format!("{n}")
            }
        }
        CellValue::Boolean(b) => if *b { "TRUE".into() } else { "FALSE".into() },
        CellValue::Error(e) => e.clone(),
        CellValue::Utf16(u) => String::from_utf16_lossy(u),
    };

    // Records the table counts: hidden items and the chosen page items drop out.
    let field_of = |rec: &Vec<CellValue>, f: usize| key(rec.get(f).unwrap_or(&CellValue::Empty));
    let axis_fields: Vec<usize> = spec.rows.iter().chain(spec.cols.iter()).flatten().copied().collect();
    let records: Vec<&Vec<CellValue>> = all
        .iter()
        .filter(|rec| {
            axis_fields.iter().chain(spec.pages.iter().map(|(f, _)| f)).all(|f| {
                let k = field_of(rec, *f);
                !spec.hidden[*f].iter().any(|h| key(h) == k)
            }) && spec.pages.iter().all(|(f, pick)| match pick {
                Some(v) => field_of(rec, *f) == key(v),
                None => true,
            })
        })
        .collect();

    // Each field's items, in the table's order, without the hidden ones.
    let items_of = |f: usize| -> Vec<CellValue> {
        let mut order: Vec<CellValue> = Vec::new();
        let mut seen = std::collections::HashSet::new();
        for v in &spec.item_order[f] {
            if seen.insert(key(v)) {
                order.push(v.clone());
            }
        }
        let mut fresh: Vec<CellValue> = Vec::new();
        for rec in &all {
            let v = rec.get(f).cloned().unwrap_or(CellValue::Empty);
            if seen.insert(key(&v)) {
                fresh.push(v);
            }
        }
        fresh.sort_by(ascending);
        order.extend(fresh);
        let hidden: std::collections::HashSet<Key> = spec.hidden[f].iter().map(key).collect();
        order.into_iter().filter(|v| !hidden.contains(&key(v))).collect()
    };

    let data_count = spec.data.len().max(1);
    // An axis laid out: depth-first items, subtotals, value lines, grand totals.
    let lay_axis = |fields: &[Option<usize>], subtotal_after: bool| -> Vec<Slot> {
        let real: Vec<usize> = fields.iter().flatten().copied().collect();
        let values_level = fields.iter().position(Option::is_none);
        let mut out = Vec::new();
        #[allow(clippy::too_many_arguments)]
        fn walk(
            fields: &[Option<usize>],
            real: &[usize],
            level: usize,
            path: Vec<Key>,
            records: &[&Vec<CellValue>],
            items_of: &dyn Fn(usize) -> Vec<CellValue>,
            data_count: usize,
            subtotal_after: bool,
            out: &mut Vec<Slot>,
        ) {
            if level >= fields.len() {
                return;
            }
            match fields[level] {
                None => {
                    for d in 0..data_count {
                        out.push(Slot { path: path.clone(), data: Some(d), kind: SlotKind::Data { level } });
                        walk(fields, real, level + 1, path.clone(), records, items_of, data_count, subtotal_after, out);
                    }
                }
                Some(f) => {
                    let depth = path.len();
                    for item in items_of(f) {
                        let mut next = path.clone();
                        next.push(key(&item));
                        let present = records.iter().any(|rec| {
                            real[..=depth]
                                .iter()
                                .zip(&next)
                                .all(|(g, k)| key(rec.get(*g).unwrap_or(&CellValue::Empty)) == *k)
                        });
                        if !present {
                            continue;
                        }
                        let below_real = fields[level + 1..].iter().any(Option::is_some);
                        out.push(Slot {
                            path: next.clone(),
                            data: None,
                            kind: SlotKind::Item { level, label: item.clone() },
                        });
                        walk(fields, real, level + 1, next.clone(), records, items_of, data_count, subtotal_after, out);
                        if subtotal_after && below_real {
                            out.push(Slot {
                                path: next.clone(),
                                data: None,
                                kind: SlotKind::Subtotal { level, label: item.clone() },
                            });
                        }
                    }
                }
            }
        }
        walk(fields, &real, 0, Vec::new(), &records, &items_of, data_count, subtotal_after, &mut out);
        let _ = values_level;
        out
    };

    let row_fields = &spec.rows;
    let col_fields = &spec.cols;
    let values_on_rows = row_fields.contains(&None);
    let values_on_cols = col_fields.contains(&None) || (!values_on_rows && spec.data.len() > 1);
    let col_fields: Vec<Option<usize>> = if values_on_cols && !col_fields.contains(&None) {
        let mut c = col_fields.clone();
        c.push(None);
        c
    } else {
        col_fields.clone()
    };
    let row_real: Vec<usize> = row_fields.iter().flatten().copied().collect();
    let col_real: Vec<usize> = col_fields.iter().flatten().copied().collect();

    let matches = |rec: &&Vec<CellValue>, fields: &[usize], path: &[Key]| {
        fields.iter().zip(path).all(|(f, k)| key(rec.get(*f).unwrap_or(&CellValue::Empty)) == *k)
    };
    let cell = |row_path: &[Key], col_path: &[Key], d: usize| -> CellValue {
        let Some(data) = spec.data.get(d) else { return CellValue::Empty };
        let mut acc = Acc::default();
        let mut hit = false;
        for rec in &records {
            if matches(rec, &row_real, row_path) && matches(rec, &col_real, col_path) {
                hit = true;
                acc.add(rec.get(data.field).unwrap_or(&CellValue::Empty));
            }
        }
        if hit {
            acc.value(data.aggregate)
        } else {
            CellValue::Empty
        }
    };

    let row_header = spec.row_header_caption.clone().unwrap_or_else(|| ui.row_labels().to_string());
    let col_header = spec.col_header_caption.clone().unwrap_or_else(|| ui.col_labels().to_string());
    let grand = spec.grand_total_caption.clone().unwrap_or_else(|| ui.grand_total().to_string());

    // Columns: the value columns left to right.
    let mut columns: Vec<Slot> = Vec::new();
    if col_fields.is_empty() {
        columns.push(Slot { path: Vec::new(), data: Some(0), kind: SlotKind::Data { level: 0 } });
    } else {
        let laid = lay_axis(&col_fields, true);
        // A heading column stands only where the axis bottoms out: a leaf
        // item, a value line with nothing below it, or a subtotal.
        let last = col_fields.len() - 1;
        for slot in laid {
            let leaf = match &slot.kind {
                SlotKind::Item { level, .. } | SlotKind::Data { level } => *level == last,
                SlotKind::Subtotal { .. } | SlotKind::Grand => true,
            };
            if leaf {
                columns.push(slot);
            }
        }
        // Subtotal columns of an item stand once per value field when values
        // sit below it on the axis.
        if spec.col_grand_totals && !col_real.is_empty() {
            if values_on_cols {
                for d in 0..data_count {
                    columns.push(Slot { path: Vec::new(), data: Some(d), kind: SlotKind::Grand });
                }
            } else {
                columns.push(Slot { path: Vec::new(), data: None, kind: SlotKind::Grand });
            }
        }
    }
    // Which data field a column carries when the row says nothing: its own,
    // or the only one.
    let column_data = |c: &Slot| c.data.unwrap_or(0);
    // The path a column carries, for the data value lookup.
    let col_path_of = |c: &Slot| c.path.clone();

    let mut grid: Vec<Vec<CellValue>> = Vec::new();
    let has_rows = !row_real.is_empty();
    // Column headings.
    if col_real.is_empty() {
        let mut header = vec![if has_rows || values_on_rows { text(&row_header) } else { CellValue::Empty }];
        if values_on_cols {
            for d in &spec.data {
                header.push(text(&d.name));
            }
        } else if !values_on_rows {
            header.push(text(&spec.data.first().map(|d| d.name.clone()).unwrap_or_default()));
        } else {
            header.push(CellValue::Empty);
        }
        if has_rows || values_on_rows || spec.data.len() > 1 {
            grid.push(header);
        }
    } else {
        let corner = if spec.data.len() == 1 && !values_on_cols && !values_on_rows {
            text(&spec.data[0].name)
        } else {
            CellValue::Empty
        };
        let corner = if has_rows || values_on_rows { corner } else { CellValue::Empty };
        let mut first = vec![corner, text(&col_header)];
        first.resize(1 + columns.len(), CellValue::Empty);
        grid.push(first);
        let levels = col_fields.len();
        for level in 0..levels {
            let mut row = vec![if level == levels - 1 && (has_rows || values_on_rows) {
                text(&row_header)
            } else {
                CellValue::Empty
            }];
            for (i, c) in columns.iter().enumerate() {
                let label = match &c.kind {
                    SlotKind::Grand if level == 0 => {
                        if values_on_cols {
                            text(&ui.data_total(&spec.data[column_data(c)].name))
                        } else {
                            text(&grand)
                        }
                    }
                    SlotKind::Subtotal { level: l, label } if *l == level => text(&ui.subtotal(&label_text(label))),
                    SlotKind::Item { .. } | SlotKind::Data { .. } => {
                        // A heading for `level` stands on the first column of
                        // its group: where the items (and value field) down to
                        // this level differ from the column before.
                        let reals = col_fields[..=level].iter().flatten().count();
                        let with_values = col_fields[..=level].contains(&None);
                        let prefix = |c: &Slot| {
                            (c.path.get(..reals).map(<[Key]>::to_vec), if with_values { c.data } else { None })
                        };
                        let opens = i == 0
                            || !matches!(columns[i - 1].kind, SlotKind::Item { .. } | SlotKind::Data { .. })
                            || prefix(&columns[i - 1]) != prefix(c);
                        if !opens {
                            CellValue::Empty
                        } else {
                            match col_fields[level] {
                                Some(f) => {
                                    let mine = c.path.get(reals - 1);
                                    items_of(f)
                                        .into_iter()
                                        .find(|v| Some(&key(v)) == mine)
                                        .map(|v| shown(&v))
                                        .unwrap_or(CellValue::Empty)
                                }
                                None => text(&spec.data[column_data(c)].name),
                            }
                        }
                    }
                    _ => CellValue::Empty,
                };
                row.push(label);
            }
            grid.push(row);
        }
    }

    // Body rows.
    let value_row = |row_path: &[Key], row_data: Option<usize>| -> Vec<CellValue> {
        columns
            .iter()
            .map(|c| cell(row_path, &col_path_of(c), row_data.unwrap_or_else(|| column_data(c))))
            .collect()
    };
    if has_rows || values_on_rows {
        let laid = lay_axis(row_fields, false);
        for slot in &laid {
            match &slot.kind {
                SlotKind::Item { level, label } => {
                    // Values below this item on the axis: the item's own row
                    // carries nothing, its value lines carry the numbers.
                    let values_below = row_fields[level + 1..].contains(&None);
                    let mut row = vec![shown(label)];
                    if values_below {
                        row.extend(std::iter::repeat_n(CellValue::Empty, columns.len()));
                    } else {
                        row.extend(value_row(&slot.path, None));
                    }
                    grid.push(row);
                }
                SlotKind::Data { .. } => {
                    let d = slot.data.unwrap_or(0);
                    let mut row = vec![text(&spec.data[d].name)];
                    row.extend(value_row(&slot.path, Some(d)));
                    grid.push(row);
                }
                _ => {}
            }
        }
        if spec.row_grand_totals {
            if values_on_rows {
                for (d, data) in spec.data.iter().enumerate() {
                    let mut row = vec![text(&ui.data_total(&data.name))];
                    row.extend(value_row(&[], Some(d)));
                    grid.push(row);
                }
            } else {
                let mut row = vec![text(&grand)];
                row.extend(value_row(&[], None));
                grid.push(row);
            }
        }
    } else {
        // Values only across the top: one line per value field.
        if col_real.is_empty() {
            let row: Vec<CellValue> = spec.data.iter().enumerate().map(|(d, _)| cell(&[], &[], d)).collect();
            if spec.data.len() > 1 {
                let mut r = vec![CellValue::Empty];
                r.extend(row);
                grid.push(r);
            } else {
                grid.push(row);
            }
        } else {
            let mut row = vec![if values_on_cols { CellValue::Empty } else { text(&spec.data[0].name) }];
            row.extend(value_row(&[], None));
            grid.push(row);
        }
    }

    // Outline and tabular layouts: a label column per row field, field names
    // for headings. The compact headings are reshaped; the body is rebuilt.
    if spec.row_layout != RowLayout::Compact && has_rows && !values_on_rows {
        let lc = row_real.len();
        let body_lines = lay_axis(row_fields, false).len() + usize::from(spec.row_grand_totals);
        let header_rows = grid.len() - body_lines;
        let mut reshaped: Vec<Vec<CellValue>> = Vec::new();
        let names: Vec<CellValue> = row_real.iter().map(|f| text(&spec.field_names[*f])).collect();
        let col_name = col_real.first().map(|f| text(&spec.field_names[*f]));
        for (i, line) in grid.iter().take(header_rows).enumerate() {
            let mut out: Vec<CellValue> = Vec::new();
            let last = i + 1 == header_rows;
            if last {
                out.extend(names.iter().cloned());
            } else {
                out.push(line.first().cloned().unwrap_or(CellValue::Empty));
                out.extend(std::iter::repeat_n(CellValue::Empty, lc - 1));
            }
            for (j, v) in line.iter().enumerate().skip(1) {
                let v = if i == 0 && j == 1 && !col_real.is_empty() && !last {
                    col_name.clone().unwrap_or_else(|| v.clone())
                } else {
                    v.clone()
                };
                out.push(v);
            }
            reshaped.push(out);
        }
        let laid: Vec<Slot> = lay_axis(row_fields, false)
            .into_iter()
            .filter(|s| matches!(s.kind, SlotKind::Item { .. }))
            .collect();
        let leaf = lc - 1;
        let blank = || vec![CellValue::Empty; lc];
        match spec.row_layout {
            RowLayout::Outline => {
                for slot in &laid {
                    if let SlotKind::Item { level, label } = &slot.kind {
                        let mut line = blank();
                        line[*level] = shown(label);
                        line.extend(value_row(&slot.path, None));
                        reshaped.push(line);
                    }
                }
            }
            _ => {
                // Open groups: (level, label, path, label still to show).
                type Group = (usize, CellValue, Vec<Key>, bool);
                let close = |open: &mut Vec<Group>, down_to: usize, out: &mut Vec<Vec<CellValue>>| {
                    while open.last().is_some_and(|g| g.0 >= down_to) {
                        let (level, label, path, _) = open.pop().unwrap();
                        let mut line = blank();
                        line[level] = text(&ui.subtotal(&label_text(&label)));
                        line.extend(value_row(&path, None));
                        out.push(line);
                    }
                };
                let mut open: Vec<Group> = Vec::new();
                for slot in &laid {
                    if let SlotKind::Item { level, label } = &slot.kind {
                        close(&mut open, *level, &mut reshaped);
                        if *level < leaf {
                            open.push((*level, label.clone(), slot.path.clone(), true));
                        } else {
                            let mut line = blank();
                            for g in open.iter_mut() {
                                if g.3 {
                                    line[g.0] = shown(&g.1);
                                    g.3 = false;
                                }
                            }
                            line[leaf] = shown(label);
                            line.extend(value_row(&slot.path, None));
                            reshaped.push(line);
                        }
                    }
                }
                close(&mut open, 0, &mut reshaped);
            }
        }
        if spec.row_grand_totals {
            let mut line = blank();
            line[0] = text(&grand);
            line.extend(value_row(&[], None));
            reshaped.push(line);
        }
        grid = reshaped;
    }

    // Page fields stand above the table: one row each, then a blank row.
    let mut origin = spec.origin;
    if !spec.pages.is_empty() {
        let mut above: Vec<Vec<CellValue>> = spec
            .pages
            .iter()
            .map(|(f, pick)| {
                vec![
                    text(&spec.field_names[*f]),
                    match pick {
                        Some(v) => shown(v),
                        None => text(ui.all()),
                    },
                ]
            })
            .collect();
        above.push(Vec::new());
        origin.0 = origin.0.saturating_sub(above.len() as u32);
        above.extend(grid);
        grid = above;
    }
    PivotGrid { origin, cells: grid }
}

/// The source value at (row, col) of a parsed workbook's named sheet.
pub fn source_reader<'a>(workbook: &'a Workbook, sheet: &str) -> impl Fn(u32, u32) -> CellValue + 'a {
    let held = workbook.sheets.iter().find(|s| s.name.eq_ignore_ascii_case(sheet));
    move |row, col| {
        held.and_then(|s| s.rows.iter().find(|r| r.index == row + 1))
            .and_then(|r| r.cells.iter().find(|c| c.col == col))
            .map(|c| c.value.clone())
            .unwrap_or(CellValue::Empty)
    }
}

/// A pivot table to be created: which source columns go where. Field indices
/// count the source range's columns from zero.
#[derive(Debug, Clone)]
pub struct NewPivot {
    pub name: String,
    pub source_sheet: String,
    /// Source range, zero-based inclusive (r0, c0, r1, c1), header row first.
    pub source: (u32, u32, u32, u32),
    pub target_sheet: String,
    /// Top-left of the table proper, zero-based; page rows stand above it.
    pub at: (u32, u32),
    pub rows: Vec<usize>,
    pub cols: Vec<usize>,
    pub pages: Vec<usize>,
    pub data: Vec<(usize, Aggregate)>,
}

impl Aggregate {
    fn attr(self) -> Option<&'static str> {
        match self {
            Aggregate::Sum => None,
            Aggregate::Count => Some("count"),
            Aggregate::Average => Some("average"),
            Aggregate::Max => Some("max"),
            Aggregate::Min => Some("min"),
            Aggregate::Product => Some("product"),
            Aggregate::CountNums => Some("countNums"),
            Aggregate::StdDev => Some("stdDev"),
            Aggregate::StdDevP => Some("stdDevp"),
            Aggregate::Var => Some("var"),
            Aggregate::VarP => Some("varp"),
        }
    }
    /// The caption Excel gives a new value field: `合計 / Sales`, `Sum of Sales`.
    pub fn caption(self, ui: Ui, field: &str) -> String {
        let (ja, en) = match self {
            Aggregate::Sum => ("合計", "Sum of"),
            Aggregate::Count => ("データの個数", "Count of"),
            Aggregate::Average => ("平均", "Average of"),
            Aggregate::Max => ("最大", "Max of"),
            Aggregate::Min => ("最小", "Min of"),
            Aggregate::Product => ("積", "Product of"),
            Aggregate::CountNums => ("数値の個数", "Count of"),
            Aggregate::StdDev => ("標本標準偏差", "StdDev of"),
            Aggregate::StdDevP => ("標準偏差", "StdDevp of"),
            Aggregate::Var => ("標本分散", "Var of"),
            Aggregate::VarP => ("分散", "Varp of"),
        };
        match ui {
            Ui::Japanese => format!("{ja} / {field}"),
            Ui::English => format!("{en} {field}"),
        }
    }
}

/// The definition a new pivot amounts to, ready for `compute`.
pub fn spec_for_new(new: &NewPivot, get: &dyn Fn(u32, u32) -> CellValue, ui: Ui) -> PivotSpec {
    let (r0, c0, r1, c1) = new.source;
    let width = (c1 - c0 + 1) as usize;
    let names: Vec<String> = (0..width)
        .map(|j| match get(r0, c0 + j as u32) {
            CellValue::String(s) => s,
            CellValue::Number(n) => format!("{n}"),
            other => format!("{other:?}"),
        })
        .collect();
    // A new table lists every field's items in ascending order.
    let mut item_order = vec![Vec::new(); width];
    for f in new.rows.iter().chain(&new.cols).chain(&new.pages) {
        let mut seen = std::collections::HashSet::new();
        let mut items: Vec<CellValue> = Vec::new();
        for r in r0 + 1..=r1 {
            let v = get(r, c0 + *f as u32);
            if seen.insert(key(&v)) {
                items.push(v);
            }
        }
        items.sort_by(ascending);
        item_order[*f] = items;
    }
    let mut cols: Vec<Option<usize>> = new.cols.iter().map(|c| Some(*c)).collect();
    if new.data.len() > 1 {
        cols.push(None);
    }
    PivotSpec {
        name: new.name.clone(),
        sheet: new.target_sheet.clone(),
        origin: new.at,
        extent: new.at,
        source_sheet: new.source_sheet.clone(),
        source: new.source,
        field_names: names.clone(),
        item_order,
        hidden: vec![Vec::new(); width],
        rows: new.rows.iter().map(|r| Some(*r)).collect(),
        cols,
        pages: new.pages.iter().map(|p| (*p, None)).collect(),
        data: new
            .data
            .iter()
            .map(|(f, how)| DataField { name: how.caption(ui, &names[*f]), field: *f, aggregate: *how })
            .collect(),
        row_header_caption: None,
        col_header_caption: None,
        grand_total_caption: None,
        data_caption: None,
        row_grand_totals: true,
        col_grand_totals: true,
        row_layout: RowLayout::Compact,
    }
}

fn esc(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;").replace('"', "&quot;")
}

fn column_name(mut col: u32) -> String {
    let mut out = Vec::new();
    col += 1;
    while col > 0 {
        let rem = (col - 1) % 26;
        out.push((b'A' + rem as u8) as char);
        col = (col - 1) / 26;
    }
    out.iter().rev().collect()
}

fn a1(row: u32, col: u32) -> String {
    format!("{}{}", column_name(col), row + 1)
}

fn shared_item(v: &CellValue) -> String {
    match v {
        CellValue::Empty => "<m/>".into(),
        CellValue::Number(n) => format!("<n v=\"{n}\"/>"),
        CellValue::Boolean(b) => format!("<b v=\"{}\"/>", u8::from(*b)),
        CellValue::Error(e) => format!("<e v=\"{}\"/>", esc(e)),
        CellValue::String(s) => format!("<s v=\"{}\"/>", esc(s)),
        CellValue::Utf16(u) => format!("<s v=\"{}\"/>", esc(&String::from_utf16_lossy(u))),
    }
}

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_TYPE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn cache_xml(spec: &PivotSpec) -> String {
    let (r0, c0, r1, c1) = spec.source;
    let mut x = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<pivotCacheDefinition xmlns=\"{MAIN_NS}\" xmlns:r=\"{REL_NS}\" saveData=\"0\" refreshOnLoad=\"1\" createdVersion=\"8\" refreshedVersion=\"8\" minRefreshableVersion=\"3\" recordCount=\"{}\"><cacheSource type=\"worksheet\"><worksheetSource ref=\"{}:{}\" sheet=\"{}\"/></cacheSource><cacheFields count=\"{}\">",
        r1 - r0,
        a1(r0, c0),
        a1(r1, c1),
        esc(&spec.source_sheet),
        spec.field_names.len()
    );
    for (f, name) in spec.field_names.iter().enumerate() {
        x.push_str(&format!("<cacheField name=\"{}\" numFmtId=\"0\">", esc(name)));
        let items = &spec.item_order[f];
        if items.is_empty() {
            x.push_str("<sharedItems/>");
        } else {
            let blank = items.iter().any(|v| matches!(v, CellValue::Empty));
            let all_numbers = items.iter().all(|v| matches!(v, CellValue::Number(_) | CellValue::Empty));
            let mut attrs = String::new();
            if blank {
                attrs.push_str(" containsBlank=\"1\"");
            }
            if all_numbers {
                let nums: Vec<f64> = items.iter().filter_map(|v| if let CellValue::Number(n) = v { Some(*n) } else { None }).collect();
                let (lo, hi) = nums.iter().fold((f64::MAX, f64::MIN), |(a, b), n| (a.min(*n), b.max(*n)));
                attrs.push_str(&format!(
                    " containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\"{} minValue=\"{lo}\" maxValue=\"{hi}\"",
                    if nums.iter().all(|n| n.fract() == 0.0) { " containsInteger=\"1\"" } else { "" }
                ));
            }
            x.push_str(&format!("<sharedItems{attrs} count=\"{}\">", items.len()));
            for v in items {
                x.push_str(&shared_item(v));
            }
            x.push_str("</sharedItems>");
        }
        x.push_str("</cacheField>");
    }
    x.push_str("</cacheFields></pivotCacheDefinition>");
    x
}

fn table_xml(spec: &PivotSpec, cache_id: u32, grid: &PivotGrid, header_rows: usize) -> String {
    let pages = spec.pages.len();
    let top = grid.origin.0 + if pages > 0 { pages as u32 + 1 } else { 0 };
    let height = grid.cells.len() as u32 - if pages > 0 { pages as u32 + 1 } else { 0 };
    let width = grid.cells.iter().skip(if pages > 0 { pages + 1 } else { 0 }).map(Vec::len).max().unwrap_or(1) as u32;
    let first_header = if spec.cols.iter().all(Option::is_none) && !spec.cols.is_empty() { 0 } else { 1 };
    let mut x = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<pivotTableDefinition xmlns=\"{MAIN_NS}\" name=\"{}\" cacheId=\"{cache_id}\" applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\" applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\" dataCaption=\"{}\" updatedVersion=\"8\" minRefreshableVersion=\"3\" useAutoFormatting=\"1\" itemPrintTitles=\"1\" createdVersion=\"8\" indent=\"0\" outline=\"1\" outlineData=\"1\" multipleFieldFilters=\"0\"><location ref=\"{}:{}\" firstHeaderRow=\"{first_header}\" firstDataRow=\"{header_rows}\" firstDataCol=\"1\"{}/>",
        esc(&spec.name),
        esc(spec.data_caption.as_deref().unwrap_or("値")),
        a1(top, spec.origin.1),
        a1(top + height.max(1) - 1, spec.origin.1 + width.max(1) - 1),
        if pages > 0 { format!(" rowPageCount=\"{pages}\" colPageCount=\"1\"") } else { String::new() }
    );
    x.push_str(&format!("<pivotFields count=\"{}\">", spec.field_names.len()));
    for f in 0..spec.field_names.len() {
        let axis = if spec.rows.contains(&Some(f)) {
            Some("axisRow")
        } else if spec.cols.contains(&Some(f)) {
            Some("axisCol")
        } else if spec.pages.iter().any(|(p, _)| *p == f) {
            Some("axisPage")
        } else {
            None
        };
        let data = spec.data.iter().any(|d| d.field == f);
        let mut attrs = String::new();
        if let Some(a) = axis {
            attrs.push_str(&format!(" axis=\"{a}\""));
        }
        if data {
            attrs.push_str(" dataField=\"1\"");
        }
        if axis.is_some() {
            let n = spec.item_order[f].len();
            x.push_str(&format!("<pivotField{attrs} showAll=\"0\"><items count=\"{}\">", n + 1));
            for i in 0..n {
                x.push_str(&format!("<item x=\"{i}\"/>"));
            }
            x.push_str("<item t=\"default\"/></items></pivotField>");
        } else {
            x.push_str(&format!("<pivotField{attrs} showAll=\"0\"/>"));
        }
    }
    x.push_str("</pivotFields>");
    let fields = |list: &[Option<usize>], tag: &str| -> String {
        if list.is_empty() {
            return String::new();
        }
        let mut s = format!("<{tag} count=\"{}\">", list.len());
        for f in list {
            s.push_str(&format!("<field x=\"{}\"/>", f.map(|v| v as i64).unwrap_or(-2)));
        }
        s.push_str(&format!("</{tag}>"));
        s
    };
    x.push_str(&fields(&spec.rows, "rowFields"));
    x.push_str(&fields(&spec.cols, "colFields"));
    if pages > 0 {
        x.push_str(&format!("<pageFields count=\"{pages}\">"));
        for (f, _) in &spec.pages {
            x.push_str(&format!("<pageField fld=\"{f}\" hier=\"-1\"/>"));
        }
        x.push_str("</pageFields>");
    }
    x.push_str(&format!("<dataFields count=\"{}\">", spec.data.len()));
    for d in &spec.data {
        x.push_str(&format!(
            "<dataField name=\"{}\" fld=\"{}\"{} baseField=\"0\" baseItem=\"0\"/>",
            esc(&d.name),
            d.field,
            d.aggregate.attr().map(|a| format!(" subtotal=\"{a}\"")).unwrap_or_default()
        ));
    }
    x.push_str("</dataFields><pivotTableStyleInfo name=\"PivotStyleLight16\" showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"0\" showColStripes=\"0\" showLastColumn=\"1\"/><extLst><ext uri=\"{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:pivotTableDefinition hideValuesRow=\"1\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"/></ext></extLst></pivotTableDefinition>");
    x
}

/// How many heading rows a computed grid starts with (page rows aside): the
/// rows before the first body row, which is `firstDataRow`.
pub fn header_rows(spec: &PivotSpec) -> usize {
    let real_cols = spec.cols.iter().flatten().count();
    let values_on_cols = spec.cols.contains(&None) || (spec.data.len() > 1 && !spec.rows.contains(&None));
    if real_cols == 0 {
        usize::from(!spec.rows.is_empty() || values_on_cols)
    } else {
        spec.cols.len() + usize::from(values_on_cols && !spec.cols.contains(&None)) + 1
    }
}

fn read_all(data: &[u8]) -> Result<Vec<(String, Vec<u8>)>, String> {
    let mut zip = zip::ZipArchive::new(std::io::Cursor::new(data)).map_err(|e| e.to_string())?;
    let mut out = Vec::new();
    for i in 0..zip.len() {
        let mut f = zip.by_index(i).map_err(|e| e.to_string())?;
        let mut buf = Vec::new();
        f.read_to_end(&mut buf).map_err(|e| e.to_string())?;
        out.push((f.name().to_string(), buf));
    }
    Ok(out)
}

/// Writes the parts of new pivot tables into a saved package: a cache
/// definition (no records, refreshed on load) and a table definition each,
/// wired into the workbook, the target sheet and the content types.
pub fn add_pivot_parts(data: &[u8], added: &[(PivotSpec, PivotGrid)]) -> Result<Vec<u8>, String> {
    use std::io::Write;
    if added.is_empty() {
        return Ok(data.to_vec());
    }
    let mut parts = read_all(data)?;
    let get = |parts: &Vec<(String, Vec<u8>)>, name: &str| parts.iter().find(|(n, _)| n == name).map(|(_, b)| String::from_utf8_lossy(b).into_owned());
    let set = |parts: &mut Vec<(String, Vec<u8>)>, name: &str, body: String| {
        if let Some(slot) = parts.iter_mut().find(|(n, _)| n == name) {
            slot.1 = body.into_bytes();
        } else {
            parts.push((name.to_string(), body.into_bytes()));
        }
    };
    let mut workbook = get(&parts, "xl/workbook.xml").ok_or("no workbook part")?;
    let mut wb_rels = get(&parts, "xl/_rels/workbook.xml.rels").ok_or("no workbook rels")?;
    let mut types = get(&parts, "[Content_Types].xml").ok_or("no content types")?;
    let mut next_cache_id = 1u32;
    for m in workbook.match_indices("cacheId=\"") {
        let rest = &workbook[m.0 + 9..];
        if let Some(id) = rest.split('"').next().and_then(|v| v.parse::<u32>().ok()) {
            next_cache_id = next_cache_id.max(id + 1);
        }
    }
    let free = |parts: &Vec<(String, Vec<u8>)>, stem: &str| -> u32 {
        (1..).find(|n| !parts.iter().any(|(name, _)| name == &format!("{stem}{n}.xml"))).unwrap_or(1)
    };
    let next_rid = |rels: &str| -> String {
        let mut n = 1;
        while rels.contains(&format!("Id=\"rId{n}\"")) {
            n += 1;
        }
        format!("rId{n}")
    };
    let sheet_parts: Vec<(String, String)> = {
        let rel_map = rels(&wb_rels);
        let mut reader = Reader::from_str(&workbook);
        let mut out = Vec::new();
        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) if local(e.name().as_ref()) == "sheet" => {
                    if let (Some(name), Some(rid)) = (attr(&e, "name"), attr(&e, "id")) {
                        if let Some((target, _)) = rel_map.get(&rid) {
                            out.push((name, resolve("xl", target)));
                        }
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        out
    };
    let mut caches_xml = String::new();
    for (spec, grid) in added {
        let sheet_part = sheet_parts
            .iter()
            .find(|(n, _)| n.eq_ignore_ascii_case(&spec.sheet))
            .map(|(_, p)| p.clone())
            .ok_or("pivot target sheet not found")?;
        let cache_n = free(&parts, "xl/pivotCache/pivotCacheDefinition");
        let cache_part = format!("xl/pivotCache/pivotCacheDefinition{cache_n}.xml");
        set(&mut parts, &cache_part, cache_xml(spec));
        let table_n = free(&parts, "xl/pivotTables/pivotTable");
        let table_part = format!("xl/pivotTables/pivotTable{table_n}.xml");
        let cache_id = next_cache_id;
        next_cache_id += 1;
        set(&mut parts, &table_part, table_xml(spec, cache_id, grid, header_rows(spec)));
        set(
            &mut parts,
            &format!("xl/pivotTables/_rels/pivotTable{table_n}.xml.rels"),
            format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{REL_TYPE}/pivotCacheDefinition\" Target=\"../pivotCache/pivotCacheDefinition{cache_n}.xml\"/></Relationships>"),
        );
        // workbook -> cache
        let rid = next_rid(&wb_rels);
        wb_rels = wb_rels.replace(
            "</Relationships>",
            &format!("<Relationship Id=\"{rid}\" Type=\"{REL_TYPE}/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition{cache_n}.xml\"/></Relationships>"),
        );
        caches_xml.push_str(&format!("<pivotCache cacheId=\"{cache_id}\" r:id=\"{rid}\"/>"));
        // sheet -> table
        let (dir, file) = sheet_part.rsplit_once('/').unwrap_or(("xl/worksheets", &sheet_part));
        let sheet_rels_name = format!("{dir}/_rels/{file}.rels");
        let mut sheet_rels = get(&parts, &sheet_rels_name).unwrap_or_else(|| {
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>".to_string()
        });
        let srid = next_rid(&sheet_rels);
        sheet_rels = sheet_rels.replace(
            "</Relationships>",
            &format!("<Relationship Id=\"{srid}\" Type=\"{REL_TYPE}/pivotTable\" Target=\"../pivotTables/pivotTable{table_n}.xml\"/></Relationships>"),
        );
        set(&mut parts, &sheet_rels_name, sheet_rels);
        types = types.replace(
            "</Types>",
            &format!("<Override PartName=\"/{cache_part}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml\"/><Override PartName=\"/{table_part}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml\"/></Types>"),
        );
    }
    // The workbook lists its caches after its calculation settings.
    // The `r` prefix has to be in scope on the root: openpyxl declares it on
    // each <sheet> instead, which leaves <pivotCaches> without it.
    let root_end = workbook.find("<workbook").and_then(|at| workbook[at..].find('>').map(|end| at + end)).unwrap_or(0);
    if !workbook[..root_end].contains("xmlns:r=") {
        workbook = workbook.replacen("<workbook", &format!("<workbook xmlns:r=\"{REL_NS}\""), 1);
    }
    if let Some(at) = workbook.find("</pivotCaches>") {
        workbook.insert_str(at, &caches_xml);
    } else {
        let block = format!("<pivotCaches>{caches_xml}</pivotCaches>");
        let anchor = ["<smartTagPr", "<smartTagTypes", "<webPublishing", "<fileRecoveryPr", "<webPublishObjects", "<extLst", "</workbook>"]
            .iter()
            .filter_map(|tag| workbook.find(tag))
            .min()
            .ok_or("workbook has no end")?;
        workbook.insert_str(anchor, &block);
    }
    set(&mut parts, "xl/workbook.xml", workbook);
    set(&mut parts, "xl/_rels/workbook.xml.rels", wb_rels);
    set(&mut parts, "[Content_Types].xml", types);

    let mut out = std::io::Cursor::new(Vec::new());
    {
        let mut writer = zip::ZipWriter::new(&mut out);
        let options = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        for (name, body) in &parts {
            writer.start_file(name.as_str(), options).map_err(|e| e.to_string())?;
            writer.write_all(body).map_err(|e| e.to_string())?;
        }
        writer.finish().map_err(|e| e.to_string())?;
    }
    Ok(out.into_inner())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn s(v: &str) -> CellValue {
        CellValue::String(v.to_string())
    }

    #[test]
    fn compact_rows_with_subtotals_columns_and_grand_totals() {
        let data = vec![
            vec![s("Region"), s("Product"), s("Month"), s("Sales")],
            vec![s("West"), s("Apple"), s("Feb"), CellValue::Number(5.0)],
            vec![s("East"), s("Apple"), s("Jan"), CellValue::Number(2.0)],
            vec![s("East"), s("Pear"), s("Feb"), CellValue::Number(3.0)],
            vec![s("West"), s("Apple"), s("Jan"), CellValue::Number(1.0)],
        ];
        let get = |r: u32, c: u32| data[r as usize][c as usize].clone();
        let spec = PivotSpec {
            name: "PT".into(),
            sheet: "P".into(),
            origin: (2, 0),
            extent: (9, 3),
            source_sheet: "D".into(),
            source: (0, 0, 4, 3),
            field_names: vec!["Region".into(), "Product".into(), "Month".into(), "Sales".into()],
            item_order: vec![Vec::new(); 4],
            hidden: vec![Vec::new(); 4],
            rows: vec![Some(0), Some(1)],
            cols: vec![Some(2)],
            pages: Vec::new(),
            data: vec![DataField { name: "Sum of Sales".into(), field: 3, aggregate: Aggregate::Sum }],
            row_header_caption: None,
            col_header_caption: None,
            grand_total_caption: None,
            data_caption: None,
            row_grand_totals: true,
            col_grand_totals: true,
            row_layout: RowLayout::Compact,
        };
        let grid = compute(&spec, &get, Ui::English);
        let n = CellValue::Number;
        let e = CellValue::Empty;
        assert_eq!(grid.origin, (2, 0));
        assert_eq!(
            grid.cells,
            vec![
                vec![s("Sum of Sales"), s("Column Labels"), e.clone(), e.clone()],
                // Month names sort by the built-in list, not alphabetically.
                vec![s("Row Labels"), s("Jan"), s("Feb"), s("Grand Total")],
                vec![s("East"), n(2.0), n(3.0), n(5.0)],
                vec![s("Apple"), n(2.0), e.clone(), n(2.0)],
                vec![s("Pear"), e.clone(), n(3.0), n(3.0)],
                vec![s("West"), n(1.0), n(5.0), n(6.0)],
                vec![s("Apple"), n(1.0), n(5.0), n(6.0)],
                vec![s("Grand Total"), n(3.0), n(8.0), n(11.0)],
            ]
        );
    }
}
