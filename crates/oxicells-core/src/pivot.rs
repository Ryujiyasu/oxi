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

/// One pivot table, as its parts define it.
#[derive(Debug, Clone)]
pub struct PivotSpec {
    pub name: String,
    /// The sheet the table stands on, as the workbook names it.
    pub sheet: String,
    /// Top-left of `location ref`, zero-based (row, column).
    pub origin: (u32, u32),
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
    };
    let mut field = None::<usize>;
    let mut next_field = 0usize;
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
                    spec.col_grand_totals = is_true(attr(&e, "colGrandTotals"), true);
                }
                "location" => {
                    let (r0, c0, _, _) = range(&attr(&e, "ref")?)?;
                    spec.origin = (r0, c0);
                }
                "pivotField" => {
                    field = Some(next_field);
                    next_field += 1;
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
