// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
use oxicells_calc::{set_user_function_hook, Value, Workbook};
use oxicells_calc::functions::Arg;
use oxicells_calc::reference::{parse_a1, RangeRef};
use std::{cell::RefCell, collections::BTreeSet, rc::Rc};

fn book() -> Workbook { let mut book = Workbook::new(); book.add_sheet("S"); book }
fn change(address: &str) -> Vec<(String, RangeRef)> {
    vec![("S".into(), RangeRef::single(parse_a1(address).unwrap()))]
}
struct Reset;
impl Drop for Reset { fn drop(&mut self) { set_user_function_hook(None); } }

#[test]
fn dependency_closure_resolves_names_without_invoking_callbacks() {
    let _reset = Reset;
    set_user_function_hook(Some(Box::new(|_, _, _| panic!("dependency discovery evaluated a UDF"))));
    let mut book = book();
    // This test follows one fixed input, independently of the caller column.
    book.define_name("InputValue", "S!$D$1").unwrap();
    for (address, formula) in [("M1", "Tick(InputValue)"), ("N1", "Tick(M1)"), ("P1", "1")] {
        book.set_formula("S", address, formula).unwrap();
    }
    assert_eq!(book.dependent_formula_ranges(&change("D1")), vec![("S".into(), (12, 0)), ("S".into(), (13, 0))]);
}

#[test]
fn let_binding_values_see_outer_names_but_the_body_uses_locals() {
    let mut book = book();
    book.define_name("tax", "$D$1").unwrap();
    book.set_formula("S", "M1", "LET(tax,1,tax)").unwrap();
    book.set_formula("S", "N1", "LET(tax,tax+1,tax)").unwrap();
    assert_eq!(book.dependent_formula_ranges(&change("D1")), vec![("S".into(), (13, 0))]);
}

#[test]
fn independent_callback_order_is_retained_and_dependencies_take_precedence() {
    let _reset = Reset;
    let calls = Rc::new(RefCell::new(Vec::new())); let seen = calls.clone();
    set_user_function_hook(Some(Box::new(move |_, _, caller| {
        seen.borrow_mut().push(caller.unwrap().1);
        Some(Arg::Value(Value::Number(seen.borrow().len() as f64)))
    })));
    let mut book = book();
    book.set_formula("S", "M1", "Tick()").unwrap();
    book.set_formula("S", "N1", "Tick()").unwrap();
    book.set_calculation_order(&[("S".into(), (13, 0)), ("S".into(), (12, 0))]);
    book.recalculate();
    assert_eq!(*calls.borrow(), [13, 12]);
    assert_eq!(book.value("S", "M1"), Value::Number(2.));
    assert_eq!(book.value("S", "N1"), Value::Number(1.));
    book.set_formula("S", "N1", "M1+1").unwrap(); calls.borrow_mut().clear();
    book.recalculate();
    assert_eq!(*calls.borrow(), [12]);
    assert_eq!(book.value("S", "N1"), Value::Number(2.));
}

#[test]
fn implicit_intersection_calls_scalar_array_and_reference_udfs_once() {
    use oxicells_calc::functions::{BlockOrigin, RangeData};
    let _reset = Reset;
    let calls = Rc::new(RefCell::new(Vec::new())); let seen = calls.clone();
    set_user_function_hook(Some(Box::new(move |name, _, _| {
        seen.borrow_mut().push(name.to_string());
        Some(match name {
            "SCALARVALUE" => Arg::Value(Value::Number(5.)),
            "ARRAYVALUE" => Arg::Range(RangeData { origin: BlockOrigin::Array,
                width: 1, height: 2, cells: vec![Value::Number(7.), Value::Number(8.)] }),
            "REFERENCEVALUE" => Arg::Range(RangeData { origin: BlockOrigin::Cells {
                sheet: "S".into(), range: RangeRef::normalised(parse_a1("B1").unwrap(), parse_a1("B3").unwrap()) },
                width: 1, height: 3, cells: vec![Value::Number(10.), Value::Number(20.), Value::Number(30.)] }),
            _ => panic!("unexpected function {name}"),
        })
    })));
    let mut book = book();
    for (address, value) in [("B1", 10.), ("B2", 20.), ("B3", 30.)] { book.set_value("S", address, Value::Number(value)).unwrap(); }
    for (address, formula) in [("D1", "=@ScalarValue()"), ("D2", "=@ReferenceValue()"), ("D3", "=@ArrayValue()")] {
        book.set_formula("S", address, formula).unwrap();
    }
    book.recalculate();
    assert_eq!(*calls.borrow(), ["SCALARVALUE", "REFERENCEVALUE", "ARRAYVALUE"]);
    for (address, value) in [("D1", 5.), ("D2", 20.), ("D3", 7.)] { assert_eq!(book.value("S", address), Value::Number(value)); }
}

#[test]
fn volatile_roots_include_names_and_mismatched_sumif_shapes() {
    let mut book = book(); book.define_name("Clock", "NOW()").unwrap();
    for (address, formula) in [("M1", "N(Clock)"), ("N1", "SUMIF(D1:D3,\">0\",F1:F2)"), ("P1", "SUMIF(D1:D3,\">0\",F1:F3)"), ("R1", "CELL(\"row\",D1)"), ("S1", "INFO(\"release\")")] {
        book.set_formula("S", address, formula).unwrap();
    }
    assert_eq!(book.volatile_formulas().into_iter().map(|(_, pos)| pos.0).collect::<BTreeSet<_>>(), BTreeSet::from([12, 13, 17, 18]));
}

#[test]
fn sumif_reads_the_resized_sum_range_and_recalculates_for_its_extra_cell() {
    let mut book = book();
    for row in 1..=3 {
        book.set_value("S", &format!("D{row}"), Value::Number(row as f64)).unwrap();
        book.set_value("S", &format!("F{row}"), Value::Number(row as f64 * 10.)).unwrap();
    }
    book.set_formula("S", "M1", "SUMIF(D1:D3,\">0\",F1:F2)").unwrap();
    book.recalculate(); assert_eq!(book.value("S", "M1"), Value::Number(60.));
    book.set_value("S", "F3", Value::Number(50.)).unwrap();
    assert_eq!(book.dependent_formula_ranges(&change("F3")), vec![("S".into(), (12, 0))]);
    book.recalculate_after(&[("S".into(), (5, 2))]);
    assert_eq!(book.value("S", "M1"), Value::Number(80.));
}

#[test]
fn partial_calculation_reports_visited_cells_and_preserves_other_caches() {
    let mut book = book();
    for (address, value) in [("D1", 1.), ("D2", 2.)] { book.set_value("S", address, Value::Number(value)).unwrap(); }
    book.set_formula("S", "M1", "D1+1").unwrap(); book.set_formula("S", "N1", "D2+1").unwrap(); book.recalculate();
    book.set_value("S", "D1", Value::Number(7.)).unwrap(); book.set_value("S", "D2", Value::Number(8.)).unwrap();
    let report = book.recalculate_these(&[("S".into(), (12, 0))]);
    assert_eq!(report.calculated, vec![("S".into(), (12, 0))]);
    assert_eq!(book.value("S", "M1"), Value::Number(8.));
    assert_eq!(book.value("S", "N1"), Value::Number(3.));
}

#[test]
fn clearing_a_spill_obstruction_reaches_the_blocked_formula() {
    let mut book = book();
    book.set_formula("S", "A1", "SEQUENCE(3)").unwrap();
    book.set_dynamic("S", "A1").unwrap();
    book.set_value("S", "A2", Value::Number(99.)).unwrap();
    book.recalculate();
    assert_eq!(book.value("S", "A1"), Value::Error(oxicells_calc::ExcelError::Spill));
    assert_eq!(book.blocked_spill_ranges(), vec![(("S".into(), (0, 0)), (1, 3))]);
    book.clear("S", "A2").unwrap();
    book.recalculate_after(&[("S".into(), (0, 1))]);
    assert_eq!(book.value("S", "A1"), Value::Number(1.));
    assert_eq!(book.value("S", "A2"), Value::Number(2.));
    assert!(book.blocked_spill_ranges().is_empty());
}
