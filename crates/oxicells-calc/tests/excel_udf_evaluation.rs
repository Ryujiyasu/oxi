// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.
use oxicells_calc::{defer_user_function, set_user_function_hook, ExcelError, UserArg, Value, Workbook};
use oxicells_calc::functions::Arg;
use std::{cell::RefCell, rc::Rc};

struct Reset;
impl Drop for Reset {
    fn drop(&mut self) { set_user_function_hook(None); }
}

fn book() -> Workbook {
    let mut book = Workbook::new();
    book.add_sheet("S");
    for (address, value) in [("A1", 1.), ("A2", 2.)] {
        book.set_value("S", address, Value::Number(value)).unwrap();
    }
    book
}

fn tagged(log: Rc<RefCell<Vec<String>>>) {
    set_user_function_hook(Some(Box::new(move |name, args, _| {
        log.borrow_mut().push(name.to_string());
        Some(Arg::Value(match name {
            "BROKEN" => Value::Error(ExcelError::Value),
            "A" => Value::Number(1.),
            "B" => Value::Number(2.),
            _ => match args.first() { Some(UserArg::Value(value)) => value.clone(), _ => Value::Number(3.) },
        }))
    })));
}

#[test]
fn measured_scalar_branches_preserve_lazy_and_eager_callback_order() {
    let _reset = Reset;
    let log = Rc::new(RefCell::new(Vec::new()));
    tagged(log.clone());
    for (formula, expected, calls) in [
        ("CHOOSE(2,A(),B())", 2., vec!["B"]),
        ("IF(TRUE,A(),B())", 1., vec!["A"]),
        ("IFERROR(A(),B())", 1., vec!["A"]),
        ("IFERROR(BROKEN(),B())", 2., vec!["BROKEN", "B"]),
        ("IFS(TRUE,A(),TRUE,B())", 1., vec!["A", "B"]),
        ("SWITCH(1,1,A(),2,B())", 1., vec!["A", "B"]),
    ] {
        log.borrow_mut().clear();
        assert_eq!(book().evaluate("S", formula).unwrap(), Value::Number(expected), "{formula}");
        assert_eq!(*log.borrow(), calls, "{formula}");
    }
}

#[test]
fn measured_selection_roots_keep_the_selected_reference() {
    let book = book();
    for formula in ["CHOOSE(2,A2,A1)", "IFS(TRUE,A1,TRUE,A2)", "SWITCH(1,1,A1,2,A2)"] {
        let Arg::Range(block) = book.evaluate_argument("S", formula).unwrap() else { panic!("{formula}") };
        let (sheet, range) = block.cell_reference().expect(formula);
        assert_eq!(sheet, "S");
        assert_eq!(range.to_string(), "A1");
    }
}

#[test]
fn measured_array_branch_selectors_broadcast_without_repeating_callbacks() {
    let _reset = Reset;
    let log = Rc::new(RefCell::new(Vec::new()));
    tagged(log.clone());
    for formula in ["CHOOSE({1,2},A(),B())", "IFS({TRUE,FALSE},A(),TRUE,B())", "SWITCH({1,2},1,A(),2,B())"] {
        log.borrow_mut().clear();
        let Arg::Range(block) = book().evaluate_argument("S", formula).unwrap() else { panic!("{formula}") };
        assert!(block.is_array(), "{formula}");
        assert_eq!((block.width, block.height), (2, 1));
        assert_eq!(block.cells, vec![Value::Number(1.), Value::Number(2.)], "{formula}");
        assert_eq!(*log.borrow(), ["A", "B"], "{formula}");
    }
}

#[test]
fn deferred_callback_stops_fallback_siblings_and_outer_udfs() {
    let _reset = Reset;
    for formula in ["IFERROR(A(),B())+C()", "OUTER(A())+C()"] {
        let log = Rc::new(RefCell::new(Vec::new()));
        let seen = log.clone();
        set_user_function_hook(Some(Box::new(move |name, _, _| {
            seen.borrow_mut().push(name.to_string());
            defer_user_function();
            Some(Arg::Value(Value::Error(ExcelError::Value)))
        })));
        let _ = book().evaluate("S", formula).unwrap();
        assert_eq!(*log.borrow(), ["A"], "{formula}");
        set_user_function_hook(None);
        assert_eq!(book().evaluate("S", "1+2").unwrap(), Value::Number(3.));
    }
}

#[test]
fn array_builtins_are_known_and_never_enter_the_udf_hook() {
    let _reset = Reset;
    set_user_function_hook(Some(Box::new(|name, _, _| panic!("builtin dispatched as UDF: {name}"))));
    for formula in ["FILTER(A1:A2,A1:A2>0)", "TRANSPOSE(A1:A2)", "TREND({1;2},{1;2},{3;4})"] {
        let Arg::Range(block) = book().evaluate_argument("S", formula).unwrap() else { panic!("{formula}") };
        assert_eq!(block.cells.len(), 2, "{formula}");
    }
}

#[test]
fn subset_calculation_preserves_cached_spill_values_and_ownership() {
    let mut book = book();
    book.set_formula("S", "D1", "SEQUENCE(3)").unwrap();
    book.set_dynamic("S", "D1").unwrap();
    book.set_cached("S", "D1", Value::Number(1.)).unwrap();
    book.set_spilled_value("S", "D2", Value::Number(2.)).unwrap();
    book.set_spilled_value("S", "D3", Value::Number(3.)).unwrap();
    book.set_cached_spill("S", "D1", 1, 3).unwrap();
    book.set_formula("S", "G1", "SUM(D1#)").unwrap();
    book.recalculate_these(&[("S".to_string(), (6, 0))]);
    assert_eq!(book.value("S", "G1"), Value::Number(6.));
    assert_eq!(book.spilled_values("S").len(), 2);
    book.set_formula("S", "D1", "SEQUENCE(1)").unwrap();
    book.recalculate_these(&[("S".to_string(), (3, 0))]);
    assert_eq!(book.value("S", "D2"), Value::Blank);
    assert!(book.spilled_values("S").is_empty());
}
