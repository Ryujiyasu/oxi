// SPDX-License-Identifier: MIT OR Apache-2.0

//! Solver's VBA functions: SolverReset, SolverOk, SolverAdd, SolverChange,
//! SolverDelete, SolverOptions, SolverGet, SolverSolve and SolverFinish,
//! as a macro with a reference to Solver.xlam calls them.
//!
//! They keep the model where Solver keeps it, in the active sheet's hidden
//! `solver_*` names, and SolverSolve runs `oxicells_core::solver` on them.
//! What each leaves behind was read off Excel 16.0 on a Japanese Office
//! through COM and Application.Run (`tools/metrics/_xlsx_vba_solver_probe.py`,
//! 2026-11-22):
//! - SolverReset leaves eng 1, neg 1, num 0, typ 1, val 0, ver 3 and nothing
//!   else;
//! - SolverOk writes opt, typ, val, adj and eng ("M!$A$3", "M!$A$1:$A$2"),
//!   keeping the constraints;
//! - SolverAdd appends lhsN / relN / rhsN: a number as written ("4"), a
//!   formula as its reference ("=$B$1" -> M!$B$1), an integer constraint
//!   as the word "整数", AllDifferent as "AllDifferent";
//! - SolverChange finds the constraint by its cells and relation and puts
//!   the new right-hand side;
//! - SolverDelete takes it out and lowers num, leaving the last names where
//!   they were;
//! - SolverGet 1..5 answer opt (no "="), typ, val as text, adj and num; 6, 7
//!   and 8 the constraints' sides and relations, as text, in an array whose
//!   first element is empty; 9 and 10 answer 100;
//! - SolverSolve answers Solver's return code (0 found, 7 not linear for
//!   Simplex LP); SolverFinish 1 keeps the answer, 2 puts the cells back.

use super::*;

const SOLVER_DEFAULTS: [(&str, &str); 6] =
    [("solver_eng", "1"), ("solver_neg", "1"), ("solver_num", "0"), ("solver_typ", "1"), ("solver_val", "0"), ("solver_ver", "3")];

/// A reference as Solver writes it: every part absolute, on its sheet.
fn absolute_reference(home: &str, text: &str) -> Option<String> {
    let text = text.trim().trim_start_matches('=');
    let mut parts = Vec::new();
    for area in text.split(',') {
        let (sheet, cells) = match area.rsplit_once('!') {
            Some((sheet, cells)) => (sheet.trim_matches('\'').replace("''", "'"), cells),
            None => (home.to_string(), area),
        };
        let mut ends = Vec::new();
        for end in cells.split(':') {
            let end = end.trim().replace('$', "").to_ascii_uppercase();
            let letters: String = end.chars().take_while(char::is_ascii_alphabetic).collect();
            let digits = &end[letters.len()..];
            if letters.is_empty() || letters.len() > 3 || digits.is_empty() || !digits.chars().all(|c| c.is_ascii_digit()) {
                return None;
            }
            ends.push(format!("${letters}${digits}"));
        }
        parts.push(format!("{}!{}", quoted_sheet(&sheet), ends.join(":")));
    }
    Some(parts.join(","))
}

impl WorkbookHost<'_> {
    fn solver_sheet(&self) -> String {
        self.workbook.sheets[self.active_sheet].name.clone()
    }

    fn solver_get_name(&self, leaf: &str) -> Option<String> {
        let key = oxicells_core::ir::scoped_defined_name(&self.solver_sheet(), leaf);
        self.workbook.defined_names.iter().find(|(held, _)| same_defined_name(held, &key)).map(|(_, v)| v.clone())
    }

    fn solver_set_name(&mut self, leaf: &str, body: String) {
        let key = oxicells_core::ir::scoped_defined_name(&self.solver_sheet(), leaf);
        match self.workbook.defined_names.iter().position(|(held, _)| same_defined_name(held, &key)) {
            Some(at) => self.workbook.defined_names[at].1 = body,
            None => self.workbook.defined_names.push((key.clone(), body)),
        }
        // Solver's names are hidden.
        self.workbook.defined_name_attributes.entry(key).or_default().insert("hidden".into(), "1".into());
    }

    fn solver_count(&self) -> usize {
        self.solver_get_name("solver_num").and_then(|v| v.trim().parse::<f64>().ok()).unwrap_or(0.0) as usize
    }

    /// A cell or range argument: a Range, or its address as text.
    fn solver_reference(&mut self, held: &Value) -> Result<String, String> {
        if let Value::Object(object) = held {
            if let Some(range) = self.range(object) {
                let sheet = self.workbook.sheets[range.sheet].name.clone();
                return Ok(format!("{}!{}", quoted_sheet(&sheet), format_range_address(range, true, true)));
            }
        }
        match held {
            Value::String(text) => absolute_reference(&self.solver_sheet(), text).ok_or_else(|| host_error(1004, "not a reference")),
            _ => Err(host_error(13, "Type mismatch")),
        }
    }

    /// A constraint's right-hand side as Solver writes it.
    fn solver_rhs(&mut self, relation: i64, held: Option<&Value>) -> Result<String, String> {
        Ok(match relation {
            4 => "\"整数\"".into(),
            5 => "\"バイナリ\"".into(),
            6 => "\"AllDifferent\"".into(),
            _ => match held {
                Some(Value::Object(_)) => self.solver_reference(held.unwrap())?,
                Some(Value::String(text)) => {
                    let body = text.trim();
                    match body.strip_prefix('=') {
                        Some(formula) => absolute_reference(&self.solver_sheet(), formula).unwrap_or_else(|| formula.trim().to_string()),
                        None => body.to_string(),
                    }
                }
                Some(other) => any_number(other).map(|n| shown_text(&Value::Double(n), None)).unwrap_or_default(),
                None => String::new(),
            },
        })
    }

    fn solver_find(&mut self, lhs: &str, relation: i64) -> Option<usize> {
        let count = self.solver_count();
        (1..=count).find(|k| {
            self.solver_get_name(&format!("solver_lhs{k}")).is_some_and(|held| held.eq_ignore_ascii_case(lhs))
                && self.solver_get_name(&format!("solver_rel{k}")).and_then(|v| v.trim().parse::<i64>().ok()) == Some(relation)
        })
    }

    /// The Solver function a macro calls by name, or `None` for another name.
    pub(super) fn solver_call(&mut self, name: &str, args: &[Value]) -> Result<Option<Value>, String> {
        let lower = name.to_ascii_lowercase();
        if !lower.starts_with("solver") {
            return Ok(None);
        }
        let arg = |at: usize| args.get(at).filter(|v| !matches!(v, Value::Missing));
        let number = |at: usize| arg(at).and_then(any_number);
        Ok(Some(match lower.as_str() {
            "solverreset" => {
                let sheet = self.solver_sheet();
                self.workbook.defined_names.retain(|(held, _)| {
                    !matches!(oxicells_core::ir::defined_name_scope(held), Some((owner, leaf))
                        if owner.eq_ignore_ascii_case(&sheet) && leaf.to_ascii_lowercase().starts_with("solver_"))
                });
                for (leaf, body) in SOLVER_DEFAULTS {
                    self.solver_set_name(leaf, body.into());
                }
                Value::Integer(0)
            }
            "solverok" => {
                if self.solver_get_name("solver_ver").is_none() {
                    for (leaf, body) in SOLVER_DEFAULTS {
                        self.solver_set_name(leaf, body.into());
                    }
                }
                if let Some(cell) = arg(0) {
                    let cell = cell.clone();
                    let body = self.solver_reference(&cell)?;
                    self.solver_set_name("solver_opt", body);
                }
                if let Some(goal) = number(1) {
                    self.solver_set_name("solver_typ", format!("{}", goal as i64));
                }
                if let Some(value) = number(2) {
                    self.solver_set_name("solver_val", shown_text(&Value::Double(value), None));
                }
                if let Some(cells) = arg(3) {
                    let cells = cells.clone();
                    let body = self.solver_reference(&cells)?;
                    self.solver_set_name("solver_adj", body);
                }
                if let Some(engine) = number(4) {
                    self.solver_set_name("solver_eng", format!("{}", engine as i64));
                }
                Value::Integer(0)
            }
            "solveradd" => {
                let Some(cells) = arg(0).cloned() else { return Err(host_error(450, "SolverAdd needs CellRef")) };
                let lhs = self.solver_reference(&cells)?;
                let relation = number(1).unwrap_or(1.0) as i64;
                let rhs = self.solver_rhs(relation, arg(2).cloned().as_ref())?;
                let k = self.solver_count() + 1;
                self.solver_set_name(&format!("solver_lhs{k}"), lhs);
                self.solver_set_name(&format!("solver_rel{k}"), relation.to_string());
                self.solver_set_name(&format!("solver_rhs{k}"), rhs);
                self.solver_set_name("solver_num", k.to_string());
                Value::Integer(0)
            }
            "solverchange" => {
                let Some(cells) = arg(0).cloned() else { return Err(host_error(450, "SolverChange needs CellRef")) };
                let lhs = self.solver_reference(&cells)?;
                let relation = number(1).unwrap_or(1.0) as i64;
                let Some(k) = self.solver_find(&lhs, relation) else { return Ok(Some(Value::Integer(1))) };
                let rhs = self.solver_rhs(relation, arg(2).cloned().as_ref())?;
                self.solver_set_name(&format!("solver_rhs{k}"), rhs);
                Value::Integer(0)
            }
            "solverdelete" => {
                let Some(cells) = arg(0).cloned() else { return Err(host_error(450, "SolverDelete needs CellRef")) };
                let lhs = self.solver_reference(&cells)?;
                let relation = number(1).unwrap_or(1.0) as i64;
                let Some(k) = self.solver_find(&lhs, relation) else { return Ok(Some(Value::Integer(1))) };
                let count = self.solver_count();
                for j in k..count {
                    for part in ["lhs", "rel", "rhs"] {
                        if let Some(next) = self.solver_get_name(&format!("solver_{part}{}", j + 1)) {
                            self.solver_set_name(&format!("solver_{part}{j}"), next);
                        }
                    }
                }
                self.solver_set_name("solver_num", (count - 1).to_string());
                Value::Integer(0)
            }
            "solveroptions" => {
                // AssumeNonNeg is the twelfth argument.
                if let Some(Value::Boolean(on)) = arg(11) {
                    self.solver_set_name("solver_neg", if *on { "1" } else { "2" }.into());
                }
                Value::Integer(0)
            }
            "solverget" => {
                let kind = number(0).unwrap_or(0.0) as i64;
                let text = |v: Option<String>| Value::String(v.unwrap_or_default());
                let count = self.solver_count();
                let list = |host: &Self, part: &str| -> Value {
                    let mut values = vec![Value::String(String::new())];
                    values.extend((1..=count).map(|k| Value::String(host.solver_get_name(&format!("solver_{part}{k}")).unwrap_or_default().trim_matches('"').to_string())));
                    let length = values.len();
                    Value::Array(ArrayValue {
                        dimensions: vec![ArrayDimension { lower_bound: 0, length }],
                        values,
                        element_default: Box::new(Value::Empty),
                        resizable: true,
                    })
                };
                match kind {
                    1 => text(self.solver_get_name("solver_opt")),
                    2 => Value::Double(self.solver_get_name("solver_typ").and_then(|v| v.parse().ok()).unwrap_or(1.0)),
                    3 => text(self.solver_get_name("solver_val")),
                    4 => text(self.solver_get_name("solver_adj")),
                    5 => Value::Double(count as f64),
                    6 => list(self, "lhs"),
                    7 => list(self, "rel"),
                    8 => list(self, "rhs"),
                    9 | 10 => Value::Integer(100),
                    _ => Value::Empty,
                }
            }
            "solversolve" => {
                let Some(model) = oxicells_core::solver::read_model(&self.workbook, self.active_sheet) else {
                    return Ok(Some(Value::Integer(oxicells_core::solver::Outcome::ModelError as i64)));
                };
                self.solver_before = model
                    .variables
                    .iter()
                    .map(|&(s, r, c)| {
                        let value = self.workbook.sheets[s]
                            .rows
                            .iter()
                            .find(|row| row.index == r + 1)
                            .and_then(|row| row.cells.iter().find(|cell| cell.col == c))
                            .map(|cell| cell.value.clone())
                            .unwrap_or(CellValue::Empty);
                        ((s, r, c), value)
                    })
                    .collect();
                let code = oxicells_core::solver::solve(&mut self.workbook, self.active_sheet) as i64;
                self.recalculate();
                Value::Integer(code)
            }
            "solverfinish" => {
                if number(0) == Some(2.0) {
                    for ((s, r, c), value) in std::mem::take(&mut self.solver_before) {
                        let at = CellAddress { sheet: s, row: r + 1, column: c };
                        self.set_cell_value_with_revision(at, value, true)?;
                    }
                    self.recalculate();
                } else {
                    self.solver_before.clear();
                }
                Value::Integer(0)
            }
            _ => return Ok(None),
        }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The probe's model, through the same calls, against what Excel left.
    #[test]
    fn solver_functions_leave_what_excel_leaves() {
        let original = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut book = oxicells_core::parse_xlsx(original).unwrap();
        let module = parse_module(r#"Public Function Run() As String
            Dim s As String, k As Long, v As Variant
            Worksheets(1).Cells.Clear
            Worksheets(1).Name = "M"
            Range("A1").Value = 0
            Range("A2").Value = 0
            Range("B1").Value = 6
            Range("A3").Formula = "=3*A1+2*A2"
            Range("A4").Formula = "=A1+A2"
            Range("A5").Formula = "=A1+3*A2"
            SolverReset
            SolverOk SetCell:="$A$3", MaxMinVal:=1, ValueOf:=0, ByChange:="$A$1:$A$2", Engine:=2
            SolverAdd CellRef:="$A$4", Relation:=1, FormulaText:="4"
            SolverAdd CellRef:="$A$5", Relation:=1, FormulaText:="=$B$1"
            SolverAdd CellRef:="$A$1:$A$2", Relation:=4
            SolverChange CellRef:="$A$4", Relation:=1, FormulaText:="5"
            SolverDelete CellRef:="$A$1:$A$2", Relation:=4
            s = SolverGet(1) & "|" & SolverGet(2) & "|" & SolverGet(3) & "|" & SolverGet(4) & "|" & SolverGet(5) & "|"
            v = SolverGet(6)
            For k = LBound(v) To UBound(v)
                s = s & "[" & v(k) & "]"
            Next
            v = SolverGet(8)
            For k = LBound(v) To UBound(v)
                s = s & "[" & v(k) & "]"
            Next
            s = s & "|" & SolverSolve(True)
            SolverFinish KeepFinal:=1
            s = s & "|" & Range("A1").Value & "," & Range("A2").Value & "," & Range("A3").Value
            SolverOk SetCell:="$A$3", MaxMinVal:=3, ValueOf:=12, ByChange:="$A$1:$A$2", Engine:=1
            s = s & "|" & SolverSolve(True)
            SolverFinish KeepFinal:=2
            s = s & "|" & Range("A1").Value & "," & Range("A2").Value & "," & Range("A3").Value
            Run = s
        End Function"#).unwrap();
        let answer = execute_with_host(&module, "Run", vec![], &mut WorkbookHost::new(&mut book, 0).unwrap()).unwrap();
        assert_eq!(
            answer,
            Value::String("M!$A$3|1|0|M!$A$1:$A$2|2|[][M!$A$4][M!$A$5][][5][M!$B$1]|0|5,0,15|0|5,0,15".into())
        );
        let names: std::collections::BTreeMap<String, String> = book
            .defined_names
            .iter()
            .filter_map(|(n, v)| n.strip_prefix("M!").map(|leaf| (leaf.to_string(), v.clone())))
            .collect();
        let want: std::collections::BTreeMap<String, String> = [
            ("solver_adj", "M!$A$1:$A$2"),
            ("solver_eng", "1"),
            ("solver_lhs1", "M!$A$4"),
            ("solver_lhs2", "M!$A$5"),
            ("solver_lhs3", "M!$A$1:$A$2"),
            ("solver_neg", "1"),
            ("solver_num", "2"),
            ("solver_opt", "M!$A$3"),
            ("solver_rel1", "1"),
            ("solver_rel2", "1"),
            ("solver_rel3", "4"),
            ("solver_rhs1", "5"),
            ("solver_rhs2", "M!$B$1"),
            ("solver_rhs3", "\"整数\""),
            ("solver_typ", "3"),
            ("solver_val", "12"),
            ("solver_ver", "3"),
        ]
        .into_iter()
        .map(|(a, b)| (a.to_string(), b.to_string()))
        .collect();
        assert_eq!(names, want);
    }

    /// The editor's add-in bridge (docs/sheets.html `addinMacro`) turns an
    /// Office.js batch into this macro; the engine must take every line and
    /// the saved book must carry the sheet, formats, name and table.
    #[test]
    fn office_addin_batch_runs_as_a_macro() {
        let original = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut book = oxicells_core::parse_xlsx(original).unwrap();
        let module = parse_module(r##"Private Function OxiTable(name As String) As Object
  Dim ws As Object, lo As Object
  For Each ws In Worksheets
    For Each lo In ws.ListObjects
      If LCase(lo.Name) = LCase(name) Then Set OxiTable = lo: Exit Function
    Next
  Next
End Function
Public Sub OxiAddin()
  Dim w10 As Double, w20 As Double
  Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Report"
  Worksheets(1).Cells(1, 1).Interior.Color = RGB(255, 255, 0)
  Worksheets(1).Cells(1, 1).Font.Bold = True
  Worksheets(1).Cells(1, 2).HorizontalAlignment = -4108
  Worksheets(1).Columns(2).ColumnWidth = 10
  w10 = Worksheets(1).Columns(2).Width
  Worksheets(1).Columns(2).ColumnWidth = 20
  w20 = Worksheets(1).Columns(2).Width
  Worksheets(1).Columns(2).ColumnWidth = Application.Max(0, 10 + (96 - w10) * 10 / (w20 - w10))
  Worksheets(1).Cells(2, 1).NumberFormat = "0.00""x"""
  ThisWorkbook.Names.Add Name:="Rate", RefersTo:="=Sheet1!$A$1"
  Worksheets(1).ListObjects.Add(1, Worksheets(1).Range("$A$1:$B$3"), , 1).Name = "Table1"
  OxiTable("Table1").ListRows.Add.Range.Value = Array("Pad", 7)
End Sub
Public Function W() As Double
  W = Worksheets(1).Columns(2).Width
End Function"##).unwrap();
        {
            let mut host = WorkbookHost::new(&mut book, 0).unwrap();
            host.set_cell_value_with_revision(CellAddress { sheet: 0, row: 1, column: 0 }, CellValue::String("Item".into()), true).unwrap();
            host.set_cell_value_with_revision(CellAddress { sheet: 0, row: 1, column: 1 }, CellValue::String("Qty".into()), true).unwrap();
            execute_with_host(&module, "OxiAddin", vec![], &mut host).unwrap();
            // Office.js asked for 96 points; Excel rounds a column to whole pixels.
            let width = execute_with_host(&module, "W", vec![], &mut host).unwrap();
            assert!(matches!(width, Value::Double(w) if (w - 96.0).abs() <= 0.75), "{width:?}");
        }
        let mut editor = oxicells_core::XlsxEditor::new(original).unwrap();
        editor.apply_workbook(&book).unwrap();
        let saved = editor.save().unwrap();
        let back = oxicells_core::parse_xlsx(&saved).unwrap();
        assert!(back.sheets.iter().any(|s| s.name == "Report"));
        let first = &back.sheets[0];
        let cell = |r: u32, c: u32| first.rows.iter().find(|row| row.index == r).and_then(|row| row.cells.iter().find(|x| x.col == c)).unwrap();
        assert!(cell(1, 0).style.bold);
        assert_eq!(cell(1, 0).style.bg_color.as_deref(), Some("FFFF00"));
        assert_eq!(cell(1, 1).style.horizontal_align.as_deref(), Some("center"));
        assert_eq!(cell(2, 0).style.number_format.as_deref(), Some("0.00\"x\""));
        assert!(back.defined_names.iter().any(|(n, v)| n == "Rate" && v == "Sheet1!$A$1"), "{:?}", back.defined_names);
        let table = first.tables.iter().find(|t| t.name == "Table1").expect("table");
        assert_eq!((table.start_row, table.end_row, table.start_col, table.end_col), (1, 4, 0, 1));
        assert!(matches!(&cell(4, 0).value, CellValue::String(s) if s == "Pad"));
    }

    /// AllDifferent: Excel's GRG answer for the probe's four jobs.
    #[test]
    fn all_different_finds_the_assignment_excel_finds() {
        let original = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let mut book = oxicells_core::parse_xlsx(original).unwrap();
        let module = parse_module(r#"Public Function Run() As String
            Dim i As Long, s As String
            Worksheets(1).Cells.Clear
            For i = 1 To 4
                Cells(i, 1).Value = i
            Next
            Range("B1:B4").Value = Application.Transpose(Array(4, 9, 2, 7))
            Range("C1").Formula = "=SUMPRODUCT(A1:A4,B1:B4)+ABS(A1-A3)*5"
            SolverReset
            SolverOk SetCell:="$C$1", MaxMinVal:=2, ValueOf:=0, ByChange:="$A$1:$A$4", Engine:=2
            SolverAdd CellRef:="$A$1:$A$4", Relation:=6
            s = SolverSolve(True) & "|" & Range("A1").Value & Range("A2").Value & Range("A3").Value & Range("A4").Value & "|"
            SolverOk SetCell:="$C$1", MaxMinVal:=2, ValueOf:=0, ByChange:="$A$1:$A$4", Engine:=1
            s = s & SolverSolve(True) & "|" & Range("A1").Value & Range("A2").Value & Range("A3").Value & Range("A4").Value & "|" & Range("C1").Value
            Run = s
        End Function"#).unwrap();
        let answer = execute_with_host(&module, "Run", vec![], &mut WorkbookHost::new(&mut book, 0).unwrap()).unwrap();
        // Simplex LP refuses the ABS() (7) and leaves 1..4; GRG finds 3,1,4,2 at 48.
        assert_eq!(answer, Value::String("7|1234|0|3142|48".into()));
    }
}
