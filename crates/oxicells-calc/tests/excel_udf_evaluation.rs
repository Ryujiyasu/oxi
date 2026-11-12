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


#[test]
fn legacy_array_group_shares_each_callback_result() {
    let _reset = Reset;
    for addresses in [vec!["M1", "M2", "M3"], vec!["M1", "N1", "O1"]] {
        let count = Rc::new(RefCell::new(0u32));
        let seen = count.clone();
        set_user_function_hook(Some(Box::new(move |_, _, _| {
            *seen.borrow_mut() += 1;
            Some(Arg::Value(Value::Number(*seen.borrow() as f64)))
        })));
        let mut held = book();
        for address in &addresses {
            let cell = oxicells_calc::reference::parse_a1(address).unwrap();
            held.set_array_member("S", address, "COUNTER()+{0;0;0}", (cell.col - 12, cell.row)).unwrap();
        }
        for expected in [1u32, 2] {
            held.recalculate();
            assert_eq!(*count.borrow(), expected);
            for address in &addresses { assert_eq!(held.value("S", address), Value::Number(expected as f64)); }
        }
    }
}

#[test]
fn legacy_array_group_passes_row_vectors_to_variant_callbacks() {
    let _reset = Reset;
    for (addresses, rows) in [
        (vec!["M1", "M2", "M3"], 3usize),
        (vec!["M1", "N1", "O1"], 1),
        (vec!["M1", "N1", "M2", "N2"], 2),
    ] {
        let log = Rc::new(RefCell::new(Vec::new()));
        let seen = log.clone();
        set_user_function_hook(Some(Box::new(move |_, args, caller| {
            let Some(UserArg::Block(block)) = args.first() else {
                return Some(Arg::Value(Value::Error(ExcelError::Value)));
            };
            seen.borrow_mut().push((caller, block.width, block.height, block.is_array()));
            Some(Arg::Range(block.clone()))
        })));
        let mut held = book();
        for address in &addresses {
            let cell = oxicells_calc::reference::parse_a1(address).unwrap();
            held.set_array_member("S", address, "ECHO(ROW())", (cell.col - 12, cell.row)).unwrap();
        }
        held.recalculate();
        assert_eq!(*log.borrow(), vec![(Some(("S".into(), 12, 0)), 1, rows, true)]);
        for address in &addresses {
            let cell = oxicells_calc::reference::parse_a1(address).unwrap();
            assert_eq!(held.value("S", address), Value::Number(cell.row as f64 + 1.0));
        }
    }
}


#[test]
fn legacy_array_index_contexts_match_native_excel() {
    let mut mismatches = Vec::new();
    for (formula, expected) in [
        ("INDEX(D1:D3,{1;2;3},1)", [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]),
        ("SUM(INDEX(D1:D3,{1;2;3},1))", [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]),
        ("SUM(INDEX({1;2;3},{1;2;3},1))", [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]),
        ("SUM(INDEX(D1:D3,{1;2;3},1)+0)", [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]),
        ("SUM(INDEX(D1:D3,N(IF(1,{1;2;3})),1))", [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]),
        ("SUM(INDEX(D1:D3,ROW(A1:A3),1))", [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]),
        ("SUM(INDEX(D1:D3,TRANSPOSE({1,2,3}),1))", [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]),
        ("SUM(INDEX(D1:D3,IF(TRUE,{1;2;3}),1))", [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]),
        ("SUM(INDEX(D1:D3,{1;2;3},1)*1)", [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]),
        ("SUMPRODUCT(INDEX(D1:D3,{1;2;3},1))", [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]),
        ("COUNT(INDEX(D1:D3,{1;2;3},1))", [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]),
        ("AVERAGE(INDEX(D1:D3,{1;2;3},1))", [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]),
        ("ROWS(INDEX(D1:D3,{1;2;3},1))", [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]),
        ("SUM(LET(r,D1:D3,INDEX(r,{1;2;3},1)))", [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]),
        ("LET(r,INDEX(D1:D3,{1;2;3},1),SUM(r))", [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]),
        ("INDEX(D1:D3,{1;2;3},1)+0", [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]),
        ("SUM(INDEX(D1:D3,{1;3},1))", [Value::Number(1.0), Value::Number(3.0), Value::Error(ExcelError::NA)]),
        ("SUM(INDEX(D1:D3,{1,3},1))", [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]),
        ("SUM(INDEX({1;2;3},{1;3},1))", [Value::Number(1.0), Value::Number(3.0), Value::Error(ExcelError::NA)]),
        ("SUMPRODUCT(INDEX({1;2;3},{1;3},1))", [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]),
        ("SUM(INDEX(D1:D3,{1;2;3},1),10)", [Value::Number(11.0), Value::Number(12.0), Value::Number(13.0)]),
        ("SUM(INDEX(D1:D3,{1;2;3},1),INDEX(D1:D3,{3;2;1},1))", [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]),
        ("SUM(INDEX(D1:D3,0,1))", [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]),
        ("SUM(INDEX(D1:D3,{0;1;2},1))", [Value::Number(6.0), Value::Number(1.0), Value::Number(2.0)]),
    ] {
        let mut held = book();
        for (address, value) in [("D1", 1.), ("D2", 2.), ("D3", 3.)] {
            held.set_value("S", address, Value::Number(value)).unwrap();
        }
        for (row, address) in ["M1", "M2", "M3"].iter().enumerate() {
            held.set_array_member("S", address, formula, (0, row as u32)).unwrap();
        }
        for _ in 0..2 {
            held.recalculate();
            let actual = ["M1", "M2", "M3"].map(|address| held.value("S", address));
            if actual != expected { mismatches.push(format!("{formula}: {actual:?} != {expected:?}")); }
        }
    }
    assert!(mismatches.is_empty(), "{}", mismatches.join("\n"));
}


#[test]
fn legacy_array_entry_and_recalculation_match_native_callback_stages() {
    let _reset = Reset;
    let mut mismatches = Vec::new();
    for (formula, stages) in [
        ("SelectionTick(5)", [(1u32, [Value::Number(5.0), Value::Number(5.0), Value::Number(5.0)]), (2u32, [Value::Number(5.0), Value::Number(5.0), Value::Number(5.0)]), (3u32, [Value::Number(5.0), Value::Number(5.0), Value::Number(5.0)]), (4u32, [Value::Number(5.0), Value::Number(5.0), Value::Number(5.0)])]),
        ("SelectionTick(5)+{1;2;3}", [(1u32, [Value::Number(6.0), Value::Number(7.0), Value::Number(8.0)]), (2u32, [Value::Number(6.0), Value::Number(7.0), Value::Number(8.0)]), (3u32, [Value::Number(6.0), Value::Number(7.0), Value::Number(8.0)]), (4u32, [Value::Number(6.0), Value::Number(7.0), Value::Number(8.0)])]),
        ("SelectionTick(5)+{1;2}", [(1u32, [Value::Number(6.0), Value::Number(7.0), Value::Error(ExcelError::NA)]), (2u32, [Value::Number(6.0), Value::Number(7.0), Value::Error(ExcelError::NA)]), (3u32, [Value::Number(6.0), Value::Number(7.0), Value::Error(ExcelError::NA)]), (4u32, [Value::Number(6.0), Value::Number(7.0), Value::Error(ExcelError::NA)])]),
        ("SelectionTick(5)+{1,2}", [(1u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (2u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (3u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (4u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("LET(unused,SelectionTick(5),INDEX(D1:D3,1,1))", [(1u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(SelectionTick(5)+{1;2;3})", [(1u32, [Value::Number(21.0), Value::Number(21.0), Value::Number(21.0)]), (2u32, [Value::Number(21.0), Value::Number(21.0), Value::Number(21.0)]), (3u32, [Value::Number(21.0), Value::Number(21.0), Value::Number(21.0)]), (4u32, [Value::Number(21.0), Value::Number(21.0), Value::Number(21.0)])]),
        ("SUM(INDEX(D1:D3,SelectionTick(1)+{0;1;2},1))", [(1u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUMPRODUCT(INDEX(D1:D3,{1;2;3},SelectionTick(1)))", [(1u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (2u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (3u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (4u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("IF({TRUE;FALSE;TRUE},SelectionTick(5),0)", [(1u32, [Value::Number(5.0), Value::Number(0.0), Value::Number(5.0)]), (2u32, [Value::Number(5.0), Value::Number(0.0), Value::Number(5.0)]), (3u32, [Value::Number(5.0), Value::Number(0.0), Value::Number(5.0)]), (4u32, [Value::Number(5.0), Value::Number(0.0), Value::Number(5.0)])]),
        ("IF(TRUE,SelectionTick(5),{1;2;3})", [(1u32, [Value::Number(5.0), Value::Number(5.0), Value::Number(5.0)]), (2u32, [Value::Number(5.0), Value::Number(5.0), Value::Number(5.0)]), (3u32, [Value::Number(5.0), Value::Number(5.0), Value::Number(5.0)]), (4u32, [Value::Number(5.0), Value::Number(5.0), Value::Number(5.0)])]),
        ("LET(r,SelectionTick(5)+{1;2;3},SUM(r))", [(1u32, [Value::Number(21.0), Value::Number(21.0), Value::Number(21.0)]), (2u32, [Value::Number(21.0), Value::Number(21.0), Value::Number(21.0)]), (3u32, [Value::Number(21.0), Value::Number(21.0), Value::Number(21.0)]), (4u32, [Value::Number(21.0), Value::Number(21.0), Value::Number(21.0)])]),
        ("LET(r,SelectionTick(5),r+{1;2;3})", [(1u32, [Value::Number(6.0), Value::Number(7.0), Value::Number(8.0)]), (2u32, [Value::Number(6.0), Value::Number(7.0), Value::Number(8.0)]), (3u32, [Value::Number(6.0), Value::Number(7.0), Value::Number(8.0)]), (4u32, [Value::Number(6.0), Value::Number(7.0), Value::Number(8.0)])]),
        ("CounterTick()", [(1u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2u32, [Value::Number(2.0), Value::Number(2.0), Value::Number(2.0)]), (3u32, [Value::Number(3.0), Value::Number(3.0), Value::Number(3.0)]), (4u32, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)])]),
        ("CounterTick()+{0;0;0}", [(1u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2u32, [Value::Number(2.0), Value::Number(2.0), Value::Number(2.0)]), (3u32, [Value::Number(3.0), Value::Number(3.0), Value::Number(3.0)]), (4u32, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)])]),
        ("SUM(INDEX(D1:D3,N({1;2;3}),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,N(IF(1,{1;2;3})),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("SUM(INDEX(D1:D3,N(ROW(A1:A3)),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,N(TRANSPOSE({1,2,3})),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,IF(1,N({1;2;3})),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,IF(1,N(ROW(A1:A3))),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,IFERROR({1;2;3},0),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,ABS({1;2;3}),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,INT({1;2;3}),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,--{1;2;3},1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,N(IF(1,{1;2;3}))+0,1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("SUM(INDEX(D1:D3,N(D1:D3),1))", [(0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(D1:D3,N(IF(1,{1;2;3}))*1,1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("LET(n,N(IF(1,{1;2;3})),SUM(INDEX(D1:D3,n,1)))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,N(IFERROR({1;2;3},0)),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,N(CHOOSE(1,{1;2;3})),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("SUM(INDEX(D1:D3,N({1;2;3})+ROW(A1:A3)*0,1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUMPRODUCT(INDEX(D1:D3,N(IF(1,{1;2;3})),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("SUM(INDEX(D1:D3,LET(n,N(IF(1,{1;2;3})),n),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,LET(n,N(IF(1,{3;2;1})),{1;2;3}),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,IF(FALSE,N(IF(1,{3;2;1})),{1;2;3}),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,IF(TRUE,N(IF(1,{1;2;3})),{3;2;1}),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,ABS(IF(1,{1;2;3})),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,INT(IF(1,{1;2;3})),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,IF(1,{1;2;3})+0,1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,--IF(1,{1;2;3}),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,TRANSPOSE(N(IF(1,{1,2,3}))),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,VALUE(IF(1,{\"1\";\"2\";\"3\"})),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,N(IF(1,{1;2;3})),{1;1;1}))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("SUM(INDEX(D1:D3,N(IF(1,{1;2;3}))+{0;0;0},1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("SUM(INDEX(D1:D3,N(INDEX({1;2;3},0,1)),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("SUM(INDEX(D1:D3,N(INDEX(D1:D3,0,1)),1))", [(0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(n,{1;2;3},SUM(INDEX(D1:D3,N(n),1)))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("LET(n,D1:D3,SUM(INDEX(D1:D3,N(n),1)))", [(0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(D1:D3,N(LET(n,{1;2;3},n)),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("SUM(INDEX(D1:D3,N(IF(TRUE,D1:D3)),1))", [(0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(D1:D3,N(CHOOSE(1,D1:D3)),1))", [(0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(D1:D3,N(OFFSET(D1,0,0,3,1)),1))", [(0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(D1:D3,N(INDIRECT(\"D1:D3\")),1))", [(0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(D1:D3,N(TAKE({1;2;3},3)),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("SUM(INDEX(D1:D3,N(LET(n,D1:D3,n)),1))", [(0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(D1:D3,N(IFNA({1;2;3},0)),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,TAKE({1;2;3},3),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,TAKE(D1:D3,3),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,N(TAKE(D1:D3,3)),1))", [(0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(D1:D3,SEQUENCE(3),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,N(SEQUENCE(3)),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,FILTER({1;2;3},TRUE),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,N(FILTER({1;2;3},TRUE)),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,CHOOSECOLS({1;2;3},1),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,N(CHOOSECOLS({1;2;3},1)),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,TRANSPOSE({1,2,3}),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,LET(r,TAKE({1;2;3},3),r),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,N(ABS({1;2;3})),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,N(DROP({1;2;3;4},-1)),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("SUM(INDEX(D1:D3,N(DROP(D1:D3,-1)),1))", [(0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1,2},{1,4;2,5;3,6})),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1,2},D1:E3)),1))", [(0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0u32, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(D1:D3,N(INDEX(D1:D3,{1;2;3},1)),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,N(IFERROR({1;2;3},0)),1))", [(0u32, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (0u32, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
    ] {
        let count = Rc::new(RefCell::new(0u32));
        let seen = count.clone();
        set_user_function_hook(Some(Box::new(move |name, args, _| {
            *seen.borrow_mut() += 1;
            Some(Arg::Value(if name == "COUNTERTICK" {
                Value::Number(*seen.borrow() as f64)
            } else {
                match args.first() { Some(UserArg::Value(value)) => value.clone(), _ => Value::Error(ExcelError::Value) }
            }))
        })));
        let mut held = book();
        for (address, value) in [("D1", 1.), ("D2", 2.), ("D3", 3.)] {
            held.set_value("S", address, Value::Number(value)).unwrap();
        }
        let addresses = ["M1", "M2", "M3"];
        for (row, address) in addresses.iter().enumerate() {
            held.set_array_member("S", address, formula, (0, row as u32)).unwrap();
        }
        let wanted = (0..3).map(|row| ("S".to_string(), (12, row))).collect::<Vec<_>>();
        for (stage, (calls, expected)) in stages.into_iter().enumerate() {
            match stage {
                0 => { held.recalculate_array_entry(&wanted); }
                1 | 2 => { held.recalculate_these(&wanted); }
                _ => { held.recalculate(); }
            }
            let actual_calls = *count.borrow();
            let actual = addresses.map(|address| held.value("S", address));
            if actual_calls != calls || actual != expected {
                mismatches.push(format!("{formula}, stage {stage}: calls {actual_calls} != {calls}, {actual:?} != {expected:?}"));
            }
        }
    }
    assert!(mismatches.is_empty(), "{}", mismatches.join("\n"));
}


#[test]
fn index_selector_coercion_matches_native_evaluate_and_cell_modes() {
    for (formula, expected) in [
        ("SUM(INDEX(D1:D3,N({1;2;3}),1))", [1.0, 1.0, 6.0]),
        ("SUM(INDEX(D1:D3,N(IF(1,{1;2;3})),1))", [6.0, 6.0, 6.0]),
        ("SUM(INDEX(D1:D3,N(ROW(A1:A3)),1))", [1.0, 1.0, 6.0]),
        ("SUM(INDEX(D1:D3,N(TRANSPOSE({1,2,3})),1))", [1.0, 1.0, 6.0]),
        ("SUM(INDEX(D1:D3,IF(1,N({1;2;3})),1))", [1.0, 1.0, 6.0]),
        ("SUM(INDEX(D1:D3,IF(1,N(ROW(A1:A3))),1))", [1.0, 1.0, 6.0]),
        ("SUM(INDEX(D1:D3,IFERROR({1;2;3},0),1))", [1.0, 1.0, 6.0]),
        ("SUM(INDEX(D1:D3,ABS({1;2;3}),1))", [1.0, 1.0, 6.0]),
        ("SUM(INDEX(D1:D3,INT({1;2;3}),1))", [1.0, 1.0, 6.0]),
        ("SUM(INDEX(D1:D3,--{1;2;3},1))", [1.0, 1.0, 6.0]),
        ("SUM(INDEX(D1:D3,N(IF(1,{1;2;3}))+0,1))", [6.0, 6.0, 6.0]),
        ("SUM(INDEX(D1:D3,N(D1:D3),1))", [1.0, 1.0, 1.0]),
        ("SUM(INDEX(D1:D3,N(IF(1,{1;2;3}))*1,1))", [6.0, 6.0, 6.0]),
        ("LET(n,N(IF(1,{1;2;3})),SUM(INDEX(D1:D3,n,1)))", [1.0, 1.0, 6.0]),
        ("SUM(INDEX(D1:D3,N(IFERROR({1;2;3},0)),1))", [1.0, 1.0, 6.0]),
        ("SUM(INDEX(D1:D3,N(CHOOSE(1,{1;2;3})),1))", [6.0, 6.0, 6.0]),
        ("SUM(INDEX(D1:D3,N({1;2;3})+ROW(A1:A3)*0,1))", [1.0, 1.0, 6.0]),
        ("SUMPRODUCT(INDEX(D1:D3,N(IF(1,{1;2;3})),1))", [6.0, 6.0, 6.0]),
    ] {
        for (mode, value) in expected.into_iter().enumerate() {
            let mut held = book();
            for (address, value) in [("D1", 1.), ("D2", 2.), ("D3", 3.)] {
                held.set_value("S", address, Value::Number(value)).unwrap();
            }
            let actual = if mode == 0 {
                held.evaluate("S", formula).unwrap()
            } else {
                held.set_formula("S", "M1", formula).unwrap();
                if mode == 2 { held.set_dynamic("S", "M1").unwrap(); }
                held.recalculate();
                held.value("S", "M1")
            };
            assert_eq!(actual, Value::Number(value), "{formula}, mode {mode}");
        }
    }
}


#[test]
fn legacy_array_lifted_callbacks_resume_without_losing_the_probe() {
    let _reset = Reset;
    let mut held = book();
    for (address, value) in [("D1", 1.), ("D2", 2.), ("D3", 3.)] {
        held.set_value("S", address, Value::Number(value)).unwrap();
    }
    let formula = "SUM(INDEX(D1:D3,TICK(1)+{0;1;2},1))";
    for (row, address) in ["M1", "M2", "M3"].iter().enumerate() {
        held.set_array_member("S", address, formula, (0, row as u32)).unwrap();
    }
    let wanted = (0..3).map(|row| ("S".to_string(), (12, row))).collect::<Vec<_>>();
    let mut calls = 0;
    for (stage, expected_calls) in [1, 5, 9, 13].into_iter().enumerate() {
        held.set_calculation_order(&wanted.iter().cloned().rev().collect::<Vec<_>>());
        held.set_chain_probes(stage != 0);
        let answers = Rc::new(RefCell::new(std::collections::BTreeSet::<String>::new()));
        let mut complete = false;
        for _ in 0..20 {
            let queued = Rc::new(RefCell::new(Vec::new()));
            let queue = queued.clone();
            let cached = answers.clone();
            set_user_function_hook(Some(Box::new(move |name, args, caller| {
                let context = oxicells_calc::engine::user_function_call_context();
                let key = format!("{name}|{args:?}|{caller:?}|{context:?}");
                if cached.borrow().contains(&key) {
                    return Some(Arg::Value(Value::Number(1.)));
                }
                if !queue.borrow().contains(&key) { queue.borrow_mut().push(key); }
                defer_user_function();
                Some(Arg::Value(Value::Error(ExcelError::Value)))
            })));
            match stage {
                0 => { held.recalculate_array_entry(&wanted); }
                1 | 2 => { held.recalculate_these(&wanted); }
                _ => { held.recalculate(); }
            }
            set_user_function_hook(None);
            if queued.borrow().is_empty() { complete = true; break; }
            for key in queued.borrow_mut().drain(..) {
                assert!(answers.borrow_mut().insert(key));
                calls += 1;
            }
        }
        assert!(complete, "stage {stage} did not settle");
        assert_eq!(calls, expected_calls, "stage {stage}");
        let expected = if stage == 0 { [6., 6., 6.] } else { [1., 2., 3.] };
        assert_eq!(["M1", "M2", "M3"].map(|address| held.value("S", address)), expected.map(Value::Number));
    }
}


#[test]
fn array_function_values_shapes_and_rank_match_native_excel() {
    for (formula, expected) in [
        ("FILTER({1,2;3,4;5,6},{TRUE,FALSE})", (2u8, 1usize, 3usize, vec![Value::Number(1.0), Value::Number(3.0), Value::Number(5.0)])),
        ("FILTER({1,2;3,4;5,6},{FALSE,TRUE})", (2u8, 1usize, 3usize, vec![Value::Number(2.0), Value::Number(4.0), Value::Number(6.0)])),
        ("FILTER({1,2;3,4;5,6},TRUE)", (0u8, 1usize, 1usize, vec![Value::Error(ExcelError::Value)])),
        ("FILTER({1,2;3,4;5,6},FALSE,99)", (0u8, 1usize, 1usize, vec![Value::Error(ExcelError::Value)])),
        ("FILTER({1;2;3},TRUE)", (2u8, 1usize, 3usize, vec![Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])),
        ("FILTER({1;2;3},FALSE,99)", (0u8, 1usize, 1usize, vec![Value::Number(99.0)])),
        ("FILTER({1,2,3},TRUE)", (1u8, 3usize, 1usize, vec![Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])),
        ("FILTER({1,2,3},{TRUE,FALSE,TRUE})", (1u8, 2usize, 1usize, vec![Value::Number(1.0), Value::Number(3.0)])),
        ("FILTER({1,2;3,4;5,6},{TRUE;FALSE;TRUE})", (2u8, 2usize, 2usize, vec![Value::Number(1.0), Value::Number(2.0), Value::Number(5.0), Value::Number(6.0)])),
        ("FILTER({1,2;3,4;5,6},{TRUE;FALSE})", (0u8, 1usize, 1usize, vec![Value::Error(ExcelError::Value)])),
        ("XLOOKUP(1,{1,2},{1,4;2,5;3,6})", (2u8, 1usize, 3usize, vec![Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])),
        ("XLOOKUP(2,{1,2},{1,4;2,5;3,6})", (2u8, 1usize, 3usize, vec![Value::Number(4.0), Value::Number(5.0), Value::Number(6.0)])),
        ("XLOOKUP(2,{1;2;3},{1,4;2,5;3,6})", (1u8, 2usize, 1usize, vec![Value::Number(2.0), Value::Number(5.0)])),
        ("XLOOKUP(5,{1,2},{1,4;2,5;3,6},99)", (0u8, 1usize, 1usize, vec![Value::Number(99.0)])),
        ("XLOOKUP(5,{1,2},{1,4;2,5;3,6},{98;99})", (2u8, 1usize, 2usize, vec![Value::Number(98.0), Value::Number(99.0)])),
        ("XLOOKUP(1,{1,2},{1;2;3})", (0u8, 1usize, 1usize, vec![Value::Error(ExcelError::Value)])),
        ("FILTER({1},TRUE)", (1u8, 1usize, 1usize, vec![Value::Number(1.0)])),
        ("FILTER({1;2},{TRUE;FALSE})", (1u8, 1usize, 1usize, vec![Value::Number(1.0)])),
        ("FILTER({1,2},{TRUE,FALSE})", (1u8, 1usize, 1usize, vec![Value::Number(1.0)])),
        ("XLOOKUP(1,{1,2},{1,4})", (0u8, 1usize, 1usize, vec![Value::Number(1.0)])),
        ("XLOOKUP(1,{1;2},{1;2})", (0u8, 1usize, 1usize, vec![Value::Number(1.0)])),
        ("XLOOKUP(1,{1},{1,2})", (1u8, 2usize, 1usize, vec![Value::Number(1.0), Value::Number(2.0)])),
        ("XLOOKUP(1,{1},{1;2})", (2u8, 1usize, 2usize, vec![Value::Number(1.0), Value::Number(2.0)])),
        ("XLOOKUP(9,{1,2},{1,4;2,5},{99})", (1u8, 1usize, 1usize, vec![Value::Number(99.0)])),
        ("XLOOKUP(1,{1},{7},{8})", (0u8, 1usize, 1usize, vec![Value::Number(7.0)])),
        ("XLOOKUP(1,{1},{7},{8,9})", (0u8, 1usize, 1usize, vec![Value::Number(7.0)])),
        ("XLOOKUP(1,{1},{7},{8;9})", (0u8, 1usize, 1usize, vec![Value::Number(7.0)])),
        ("XLOOKUP(1,{1},{7})", (0u8, 1usize, 1usize, vec![Value::Number(7.0)])),
        ("XLOOKUP(9,{1},{7},{8})", (1u8, 1usize, 1usize, vec![Value::Number(8.0)])),
        ("XLOOKUP(1,{1},{7,8},{9})", (1u8, 2usize, 1usize, vec![Value::Number(7.0), Value::Number(8.0)])),
        ("XLOOKUP(1,{1},{7;8},{9})", (2u8, 1usize, 2usize, vec![Value::Number(7.0), Value::Number(8.0)])),
    ] {
        let actual = match book().evaluate_argument("S", formula).unwrap() {
            Arg::Value(value) => (0u8, 1usize, 1usize, vec![value]),
            Arg::Range(block) => (if block.is_vector() { 1u8 } else { 2u8 }, block.width, block.height, block.cells),
        };
        assert_eq!(actual, expected, "{formula}");
    }
}

#[test]
fn single_key_xlookup_keeps_the_returned_reference() {
    for (formula, expected) in [
        ("ISREF(XLOOKUP(1,D1,D1:E1))", Value::Logical(true)),
        ("ISREF(XLOOKUP(1,D1,D1:D3))", Value::Logical(true)),
        ("N(XLOOKUP(1,D1,D1:E1))", Value::Number(1.0)),
        ("N(XLOOKUP(1,D1,D1:D3))", Value::Number(1.0)),
        ("ROWS(XLOOKUP(1,D1,D1:D3))", Value::Number(3.0)),
        ("COLUMNS(XLOOKUP(1,D1,D1:E1))", Value::Number(2.0)),
    ] {
        let mut held = book();
        for (address, value) in [("D1", 1.), ("D2", 2.), ("D3", 3.)] {
            held.set_value("S", address, Value::Number(value)).unwrap();
        }
        let actual = held.evaluate_argument("S", formula).unwrap();
        match actual {
            Arg::Value(value) => assert_eq!(value, expected, "{formula}"),
            other => panic!("{formula}: expected a scalar, got {other:?}"),
        }
    }
}

#[test]
fn legacy_array_lifted_callbacks_resume_through_the_core_bridge() {
    use oxicells_core::ir::{Cell, CellValue, Row, Spill};
    let _reset = Reset;
    let mut held = oxicells_core::parse_xlsx(include_bytes!("../../oxicells-core/tests/fixtures/hidden_rows_cols.xlsx")).unwrap();
    held.sheets.truncate(1);
    held.sheets[0].name = "S".into();
    held.sheets[0].rows.clear();
    held.sheets[0].cached_spills = None;
    held.defined_names.clear();
    held.external_books.clear();
    held.recalculation = Default::default();
    for row in 1..=3 {
        let cell = |col, value, formula, array_block| Cell {
            col, value, formula, array_block, style: Default::default(),
            formula_utf16: None, runs: Vec::new(), spill: Spill::Scalar,
        };
        held.sheets[0].rows.push(Row {
            index: row, height: None, custom_height: false, format_recorded: None, style_font: None,
            thick_top: false, thick_bottom: false, hidden: false,
            cells: vec![
                cell(3, CellValue::Number(row as f64), None, None),
                cell(12, CellValue::Empty, Some("SUM(INDEX(D1:D3,TICK(1)+{0;1;2},1))".into()), Some((1,12,3,12))),
            ],
        });
    }
    let wanted = (0..3).map(|row| ("S".to_string(), (12, row))).collect::<Vec<_>>();
    let mut calls = 0;
    for (stage, expected_calls) in [1, 5, 9, 13].into_iter().enumerate() {
        let answers = Rc::new(RefCell::new(std::collections::BTreeSet::<String>::new()));
        let mut complete = false;
        for _ in 0..20 {
            let queued = Rc::new(RefCell::new(Vec::new()));
            let queue = queued.clone();
            let cached = answers.clone();
            set_user_function_hook(Some(Box::new(move |name, args, caller| {
                let context = oxicells_calc::engine::user_function_call_context();
                let key = format!("{name}|{args:?}|{caller:?}|{context:?}");
                if cached.borrow().contains(&key) { return Some(Arg::Value(Value::Number(1.))); }
                if !queue.borrow().contains(&key) { queue.borrow_mut().push(key); }
                defer_user_function();
                Some(Arg::Value(Value::Error(ExcelError::Value)))
            })));
            match stage {
                0 => { oxicells_core::formula::evaluate_workbook_formula_array_entry(&mut held, &wanted, None); }
                1 | 2 => { oxicells_core::formula::evaluate_workbook_formula_subset(&mut held, &wanted, false, None); }
                _ => { oxicells_core::formula::evaluate_workbook_formulas(&mut held); }
            }
            set_user_function_hook(None);
            if queued.borrow().is_empty() { complete = true; break; }
            for key in queued.borrow_mut().drain(..) {
                assert!(answers.borrow_mut().insert(key));
                calls += 1;
            }
        }
        assert!(complete, "stage {stage} did not settle");
        assert_eq!(calls, expected_calls, "stage {stage}");
        let actual = held.sheets[0].rows.iter().map(|row| match &row.cells.iter().find(|cell| cell.col == 12).unwrap().value { CellValue::Number(number) => *number, other => panic!("unexpected cell value: {other:?}") }).collect::<Vec<_>>();
        let expected = if stage == 0 { [6., 6., 6.] } else { [1., 2., 3.] };
        assert_eq!(actual, expected.to_vec());
    }
}

#[test]
fn legacy_array_let_reference_dependencies_and_cycles_match_native_excel() {
    struct Restore;
    impl Drop for Restore {
        fn drop(&mut self) {
            set_user_function_hook(None);
            oxicells_calc::engine::set_user_function_wait_for_dependencies(false);
            oxicells_calc::engine::set_user_function_full_rebuild(false);
        }
    }
    let _restore = Restore;
    let cases = [
        ("LET(r,A1:O3,5)",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,A1:O3,ISREF(r))",1u32,3u32,[(0,vec![Value::Logical(true),Value::Logical(true),Value::Logical(true)]),(0,vec![Value::Logical(true),Value::Logical(true),Value::Logical(true)]),(0,vec![Value::Logical(true),Value::Logical(true),Value::Logical(true)]),(0,vec![Value::Logical(true),Value::Logical(true),Value::Logical(true)])]),
        ("LET(r,A1:O3,ROWS(r))",1u32,3u32,[(0,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)]),(0,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)]),(0,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)]),(0,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)])]),
        ("LET(r,A1:O3,COLUMNS(r))",1u32,3u32,[(0,vec![Value::Number(15.0),Value::Number(15.0),Value::Number(15.0)]),(0,vec![Value::Number(15.0),Value::Number(15.0),Value::Number(15.0)]),(0,vec![Value::Number(15.0),Value::Number(15.0),Value::Number(15.0)]),(0,vec![Value::Number(15.0),Value::Number(15.0),Value::Number(15.0)])]),
        ("LET(r,A1:O3,INDEX(r,1,4))",1u32,3u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),A1:O3,D1:D3),5)",1u32,3u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(4,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,CHOOSE(SelectionTick(1),A1:O3,D1:D3),5)",1u32,3u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(4,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,INDEX(A1:O3,0,0),5)",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,OFFSET(A1,0,0,3,15),5)",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,INDIRECT(\"A1:O3\"),5)",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SelectionTick(5),5)",1u32,3u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(4,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SUM(M1:M3),5)",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,M1:M3,SUM(r))",1u32,3u32,[(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,A1:O3,SUM(INDEX(r,1,13)))",1u32,3u32,[(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,A1:O3,LET(r,D1:D3,SUM(r)))",1u32,3u32,[(0,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(0,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(0,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(0,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)])]),
        ("LET(r,A1:O3,LET(r,r,INDEX(r,{1;2;3},4)))",1u32,3u32,[(0,vec![Value::Number(1.0),Value::Number(2.0),Value::Number(3.0)]),(0,vec![Value::Number(1.0),Value::Number(2.0),Value::Number(3.0)]),(0,vec![Value::Number(1.0),Value::Number(2.0),Value::Number(3.0)]),(0,vec![Value::Number(1.0),Value::Number(2.0),Value::Number(3.0)])]),
        ("LET(r,SelectionTick(SUM(M1:M3)),5)",1u32,3u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SUM(M1:M3)+SelectionTick(5),5)",1u32,3u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(FALSE,SelectionTick(5),SUM(M1:M3)),5)",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(TRUE,SelectionTick(5),SUM(M1:M3)),5)",1u32,3u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(4,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(FALSE,SelectionTick(SUM(M1:M3)),1),5)",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(TRUE,1,SelectionTick(SUM(M1:M3))),5)",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SelectionTick(5)/0,5)",1u32,3u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(4,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,1/0,5)",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SUM(M1:M3),SelectionTick(5))",1u32,3u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SUM(M1:M3),ISREF(r))",1u32,3u32,[(0,vec![Value::Logical(false),Value::Logical(false),Value::Logical(false)]),(0,vec![Value::Logical(false),Value::Logical(false),Value::Logical(false)]),(0,vec![Value::Logical(false),Value::Logical(false),Value::Logical(false)]),(0,vec![Value::Logical(false),Value::Logical(false),Value::Logical(false)])]),
        ("LET(r,SUM(M1:M3),N(r))",1u32,3u32,[(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,SUM(M1:M3),LET(r,5,r))",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SelectionTick(5),LET(r,6,r))",1u32,3u32,[(1,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(2,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(3,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(4,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)])]),
        ("LET(r,IF(SelectionTick(1),M1:M3,D1:D3),ROWS(r))",1u32,3u32,[(1,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)]),(2,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)]),(3,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)]),(4,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)])]),
        ("LET(r,INDEX(A1:O3,SelectionTick(1),0),5)",1u32,3u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(4,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,OFFSET(A1,SelectionTick(0),0,3,15),5)",1u32,3u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(4,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("SUM(M1:M3)+1",1u32,3u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,SUM(M1:M3),r+1)",1u32,3u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,SUM(M1:M3),r-r+5)",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SUM(M1:M3),TYPE(r))",1u32,3u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,SUM(M1:M3),IF(FALSE,r,5))",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SUM(M1:M3),ISNUMBER(r))",1u32,3u32,[(0,vec![Value::Logical(true),Value::Logical(true),Value::Logical(true)]),(0,vec![Value::Logical(true),Value::Logical(true),Value::Logical(true)]),(0,vec![Value::Logical(true),Value::Logical(true),Value::Logical(true)]),(0,vec![Value::Logical(true),Value::Logical(true),Value::Logical(true)])]),
        ("LET(r,SUM(M1:M3),r*0+5)",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,M1:M3,INDEX(r,1,1))",1u32,3u32,[(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,M1:M3,INDEX(r,3,1))",1u32,3u32,[(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,IF(SelectionTick(1),A1:O3,D1:D3),INDEX(r,1,4))",1u32,3u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(0),A1:O3,D1:D3),INDEX(r,1,1))",1u32,3u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,A1:O3,s,r,INDEX(s,1,4))",1u32,3u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,A1:O3,s,SUM(r),INDEX(r,1,4))",1u32,3u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,A1:O3,s,INDEX(r,1,4),s)",1u32,3u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,A1:O3,s,INDEX(r,1,13),5)",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,A1:O3,s,INDEX(r,1,13),N(s))",1u32,3u32,[(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,A1:O3,ROWS(INDEX(r,0,13)))",1u32,3u32,[(0,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)]),(0,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)]),(0,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)]),(0,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)])]),
        ("LET(r,A1:O3,ISREF(INDEX(r,1,13)))",1u32,3u32,[(0,vec![Value::Logical(true),Value::Logical(true),Value::Logical(true)]),(0,vec![Value::Logical(true),Value::Logical(true),Value::Logical(true)]),(0,vec![Value::Logical(true),Value::Logical(true),Value::Logical(true)]),(0,vec![Value::Logical(true),Value::Logical(true),Value::Logical(true)])]),
        ("LET(r,A1:O3,LET(r,D1:D3,SUM(r))+SelectionTick(5))",1u32,3u32,[(1,vec![Value::Number(11.0),Value::Number(11.0),Value::Number(11.0)]),(2,vec![Value::Number(11.0),Value::Number(11.0),Value::Number(11.0)]),(3,vec![Value::Number(11.0),Value::Number(11.0),Value::Number(11.0)]),(4,vec![Value::Number(11.0),Value::Number(11.0),Value::Number(11.0)])]),
        ("LET(r,SUM(M1:M3),SelectionTick(5)+{1;2;3})",1u32,3u32,[(1,vec![Value::Number(6.0),Value::Number(7.0),Value::Number(8.0)]),(1,vec![Value::Number(6.0),Value::Number(7.0),Value::Number(8.0)]),(1,vec![Value::Number(6.0),Value::Number(7.0),Value::Number(8.0)]),(1,vec![Value::Number(6.0),Value::Number(7.0),Value::Number(8.0)])]),
        ("LET(r,SelectionTick(SUM(M1:M3)),r*0+5)",1u32,3u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,M1:M3,IF(FALSE,SUM(r),5))",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,M1:M3,IF(TRUE,SUM(r),5))",1u32,3u32,[(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,IF(TRUE,M1:M3,D1:D3),5)",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SUM(M1),SelectionTick(5))",1u32,1u32,[(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)])]),
        ("LET(r,SelectionTick(SUM(M1)),5)",1u32,1u32,[(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)])]),
        ("SUM(M1)+1",1u32,1u32,[(0,vec![Value::Number(1.0)]),(0,vec![Value::Number(1.0)]),(0,vec![Value::Number(1.0)]),(0,vec![Value::Number(1.0)])]),
        ("LET(r,M1,5)",1u32,1u32,[(0,vec![Value::Number(5.0)]),(0,vec![Value::Number(5.0)]),(0,vec![Value::Number(5.0)]),(0,vec![Value::Number(5.0)])]),
        ("LET(r,M1,SelectionTick(5))",1u32,1u32,[(1,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(3,vec![Value::Number(5.0)]),(4,vec![Value::Number(5.0)])]),
        ("LET(r,M1,SUM(r))",1u32,1u32,[(0,vec![Value::Number(0.0)]),(0,vec![Value::Number(0.0)]),(0,vec![Value::Number(0.0)]),(0,vec![Value::Number(0.0)])]),
        ("LET(r,M1,ROWS(r))",1u32,1u32,[(0,vec![Value::Number(1.0)]),(0,vec![Value::Number(1.0)]),(0,vec![Value::Number(1.0)]),(0,vec![Value::Number(1.0)])]),
        ("LET(r,A1:O3,INDEX(r,1,4))",1u32,1u32,[(0,vec![Value::Number(1.0)]),(0,vec![Value::Number(1.0)]),(0,vec![Value::Number(1.0)]),(0,vec![Value::Number(1.0)])]),
        ("LET(r,SUM(M1:M3),SelectionTick(5))",1u32,3u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SelectionTick(SUM(M1:M3)),5)",1u32,3u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("SUM(M1:M3)+1",1u32,3u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,M1:M3,5)",1u32,3u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,M1:M3,SelectionTick(5))",1u32,3u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(4,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,M1:M3,SUM(r))",1u32,3u32,[(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,M1:M3,ROWS(r))",1u32,3u32,[(0,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)]),(0,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)]),(0,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)]),(0,vec![Value::Number(3.0),Value::Number(3.0),Value::Number(3.0)])]),
        ("LET(r,A1:O3,INDEX(r,1,4))",1u32,3u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,SUM(M1:O1),SelectionTick(5))",3u32,1u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SelectionTick(SUM(M1:O1)),5)",3u32,1u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("SUM(M1:O1)+1",3u32,1u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,M1:O1,5)",3u32,1u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,M1:O1,SelectionTick(5))",3u32,1u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(4,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,M1:O1,SUM(r))",3u32,1u32,[(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,M1:O1,ROWS(r))",3u32,1u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,A1:O3,INDEX(r,1,4))",3u32,1u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,SUM(M1:N2),SelectionTick(5))",2u32,2u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SelectionTick(SUM(M1:N2)),5)",2u32,2u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("SUM(M1:N2)+1",2u32,2u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,M1:N2,5)",2u32,2u32,[(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(0,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,M1:N2,SelectionTick(5))",2u32,2u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(4,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,M1:N2,SUM(r))",2u32,2u32,[(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(0,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,M1:N2,ROWS(r))",2u32,2u32,[(0,vec![Value::Number(2.0),Value::Number(2.0),Value::Number(2.0),Value::Number(2.0)]),(0,vec![Value::Number(2.0),Value::Number(2.0),Value::Number(2.0),Value::Number(2.0)]),(0,vec![Value::Number(2.0),Value::Number(2.0),Value::Number(2.0),Value::Number(2.0)]),(0,vec![Value::Number(2.0),Value::Number(2.0),Value::Number(2.0),Value::Number(2.0)])]),
        ("LET(r,A1:O3,INDEX(r,1,4))",2u32,2u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(0),M1,D1),SUM(r))",1u32,1u32,[(1,vec![Value::Number(1.0)]),(2,vec![Value::Number(1.0)]),(3,vec![Value::Number(1.0)]),(4,vec![Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),M1,D1),SUM(r))",1u32,1u32,[(1,vec![Value::Number(0.0)]),(2,vec![Value::Number(0.0)]),(3,vec![Value::Number(0.0)]),(6,vec![Value::Number(0.0)])]),
        ("SUM(IF(SelectionTick(0),M1,D1))",1u32,1u32,[(1,vec![Value::Number(1.0)]),(2,vec![Value::Number(1.0)]),(3,vec![Value::Number(1.0)]),(4,vec![Value::Number(1.0)])]),
        ("SUM(IF(SelectionTick(1),M1,D1))",1u32,1u32,[(1,vec![Value::Number(0.0)]),(2,vec![Value::Number(0.0)]),(3,vec![Value::Number(0.0)]),(6,vec![Value::Number(0.0)])]),
        ("LET(r,IF(SelectionTick(0),M1:M3,D1),SUM(r))",1u32,3u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),M1:M3,D1),SUM(r))",1u32,3u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(3,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(6,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("SUM(IF(SelectionTick(0),M1:M3,D1))",1u32,3u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("SUM(IF(SelectionTick(1),M1:M3,D1))",1u32,3u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(3,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(6,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,IF(SelectionTick(0),M1:O1,D1),SUM(r))",3u32,1u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),M1:O1,D1),SUM(r))",3u32,1u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(3,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(6,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("SUM(IF(SelectionTick(0),M1:O1,D1))",3u32,1u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("SUM(IF(SelectionTick(1),M1:O1,D1))",3u32,1u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(3,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(6,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,IF(SelectionTick(0),M1:N2,D1),SUM(r))",2u32,2u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),M1:N2,D1),SUM(r))",2u32,2u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(3,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(6,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("SUM(IF(SelectionTick(0),M1:N2,D1))",2u32,2u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("SUM(IF(SelectionTick(1),M1:N2,D1))",2u32,2u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(3,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(6,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
    ];
    let mut mismatches = Vec::new();
    for (formula, width, height, stages) in cases {
        for version in [1u8, 2] {
            let count = Rc::new(RefCell::new(0u32));
            let counted = count.clone();
            set_user_function_hook(Some(Box::new(move |name, args, _| {
                if !name.eq_ignore_ascii_case("SelectionTick") { return None; }
                *counted.borrow_mut() += 1;
                let value = match args.first() {
                    Some(UserArg::Value(value)) => value.clone(),
                    Some(UserArg::Cells { values, .. }) => values.first().cloned().unwrap_or(Value::Blank),
                    Some(UserArg::Block(block)) => block.cells.first().cloned().unwrap_or(Value::Blank),
                    _ => Value::Blank,
                };
                Some(Arg::Value(value))
            })));
            let mut held = book();
            held.set_compatibility_version(version);
            for (address, value) in [("D1", 1.), ("D2", 2.), ("D3", 3.)] {
                held.set_value("S", address, Value::Number(value)).unwrap();
            }
            let mut wanted = Vec::new();
            for row in 0..height { for col in 0..width {
                let coord = (12 + col, row);
                let address = oxicells_calc::CellRef::new(coord.0, coord.1).to_a1();
                held.set_array_member("S", &address, formula, (col, row)).unwrap();
                wanted.push(("S".to_string(), coord));
            }}
            let mut reverse = wanted.clone(); reverse.reverse();
            held.set_calculation_order(&reverse);
            for (stage, (expected_calls, expected_values)) in stages.iter().enumerate() {
                held.set_chain_probes(stage != 0);
                oxicells_calc::engine::set_user_function_wait_for_dependencies(true);
                oxicells_calc::engine::set_user_function_full_rebuild(stage == 3);
                match stage {
                    0 => { held.recalculate_array_entry(&wanted); }
                    1 | 2 => { held.recalculate_these(&wanted); }
                    _ => { held.recalculate(); }
                }
                let values: Vec<_> = wanted.iter().map(|(_, coord)|
                    held.value("S", &oxicells_calc::CellRef::new(coord.0, coord.1).to_a1())).collect();
                let calls = *count.borrow();
                if values != *expected_values || calls != *expected_calls {
                    mismatches.push(format!("{formula}, {width}x{height}, version {version}, stage {stage}: calls {calls} != {expected_calls}, {values:?} != {expected_values:?}"));
                }
            }
        }
    }
    assert!(mismatches.is_empty(), "{} Native stage mismatches:\n{}", mismatches.len(), mismatches.join("\n"));
}

#[test]
fn index_selector_return_and_fallback_storage_match_native_excel() {
    for (formula, expected) in [
        ("SUM(INDEX(D1:D3,N({1;2;3}),1))",true),
        ("SUM(INDEX(D1:D3,N(IF(1,{1;2;3})),1))",true),
        ("SUM(INDEX(D1:D3,N(ROW(A1:A3)),1))",true),
        ("SUM(INDEX(D1:D3,N(TRANSPOSE({1,2,3})),1))",true),
        ("SUM(INDEX(D1:D3,IF(1,N({1;2;3})),1))",true),
        ("SUM(INDEX(D1:D3,IF(1,N(ROW(A1:A3))),1))",true),
        ("SUM(INDEX(D1:D3,IFERROR({1;2;3},0),1))",true),
        ("SUM(INDEX(D1:D3,ABS({1;2;3}),1))",true),
        ("SUM(INDEX(D1:D3,INT({1;2;3}),1))",true),
        ("SUM(INDEX(D1:D3,--{1;2;3},1))",true),
        ("SUM(INDEX(D1:D3,N(IF(1,{1;2;3}))+0,1))",true),
        ("SUM(INDEX(D1:D3,N(D1:D3),1))",false),
        ("SUM(INDEX(D1:D3,N(IF(1,{1;2;3}))*1,1))",true),
        ("LET(n,N(IF(1,{1;2;3})),SUM(INDEX(D1:D3,n,1)))",true),
        ("SUM(INDEX(D1:D3,N(IFERROR({1;2;3},0)),1))",true),
        ("SUM(INDEX(D1:D3,N(CHOOSE(1,{1;2;3})),1))",true),
        ("SUM(INDEX(D1:D3,N({1;2;3})+ROW(A1:A3)*0,1))",true),
        ("SUMPRODUCT(INDEX(D1:D3,N(IF(1,{1;2;3})),1))",true),
        ("SUM(INDEX(D1:D3,N(INDEX(D1:D3,0,1)),1))",false),
        ("SUM(INDEX(D1:D3,N(INDEX({1;2;3},0,1)),1))",true),
        ("SUM(INDEX(D1:D3,N(IF(1,D1:D3)),1))",false),
        ("SUM(INDEX(D1:D3,N(CHOOSE(1,D1:D3)),1))",false),
        ("LET(n,{1;2;3},SUM(INDEX(D1:D3,N(n),1)))",true),
        ("LET(n,D1:D3,SUM(INDEX(D1:D3,N(n),1)))",false),
        ("SUM(INDEX(D1:D3,N(OFFSET(D1,0,0,3,1)),1))",false),
        ("SUM(INDEX(D1:D3,N(INDIRECT(\"D1:D3\")),1))",false),
        ("SUM(INDEX(D1:D3,N(D1:D3+0),1))",true),
        ("SUM(INDEX(D1:D3,N(IFERROR(D1:D3,0)),1))",true),
        ("SUM(INDEX(D1:D3,N(TAKE(D1:D3,3)),1))",true),
        ("SUM(INDEX(D1:D3,N(TAKE({1;2;3},3)),1))",true),
        ("SUM(INDEX(D1:D3,N(DROP(D1:D3,1)),1))",true),
        ("SUM(INDEX(D1:D3,N(DROP({0;1;2;3},1)),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1},{1;2;3})),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,D1,D1:D3)),1))",false),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1},D1:D3)),1))",false),
        ("SUM(INDEX(D1:D3,N(LET(r,D1:D3,r)),1))",false),
        ("SUM(INDEX(D1:D3,N(LET(r,{1;2;3},r)),1))",true),
        ("SUM(INDEX(D1:D3,N(IF(1,TAKE(D1:D3,3))),1))",true),
        ("SUM(INDEX(D1:D3,N(IFERROR(TAKE(D1:D3,3),0)),1))",true),
        ("SUM(INDEX(D1:D3,N(TOROW(D1:D3)),1))",true),
        ("SUM(INDEX(D1:D3,N(FILTER(D1:D3,D1:D3>0)),1))",true),
        ("SUM(INDEX(D1:D3,N(CHOOSECOLS(D1:D3,1)),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(9,{1},{1;2;3},{1;2;3})),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1,2,3},{1,2,3})),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1},{1})),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1},{1},{2})),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(9,{1},{1},{2})),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1},D1)),1))",false),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1},D1,{2})),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1},D1,{2;3})),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(9,{1},D1,{2})),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(9,{1},D1,{2;3})),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,D1,D1:D3,{2;3})),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(9,D1,D1:D3,{2;3})),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1},{1},D1:D3)),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(9,{1},{1},D1:D3)),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1},D1:D3+0)),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1},IF(1,D1:D3))),1))",false),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1},IFERROR(D1:D3,0))),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1},TAKE(D1:D3,3))),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1},DROP(D1:D3,0))),1))",true),
        ("SUM(INDEX(D1:D3,N(XLOOKUP(1,{1},LET(r,D1:D3,r))),1))",false),
    ] {
        assert_eq!(oxicells_calc::lexer::requires_array_storage(formula, &|_| false), expected, "{formula}");
    }
}

#[test]
fn legacy_array_conditional_cycles_match_native_calculation_scopes() {
    struct Restore;
    impl Drop for Restore {
        fn drop(&mut self) {
            set_user_function_hook(None);
            oxicells_calc::engine::set_user_function_wait_for_dependencies(false);
            oxicells_calc::engine::set_user_function_full_rebuild(false);
            oxicells_calc::engine::set_user_function_full_calculation(false);
        }
    }
    let _restore = Restore;
    let cases = [
        ("LET(r,IF(SelectionTick(0),M1,D1),SUM(r))",1u32,1u32,0u32,[(1,vec![Value::Number(1.0)]),(2,vec![Value::Number(1.0)]),(3,vec![Value::Number(1.0)]),(4,vec![Value::Number(1.0)]),(5,vec![Value::Number(1.0)]),(6,vec![Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),M1,D1),SUM(r))",1u32,1u32,0u32,[(1,vec![Value::Number(0.0)]),(2,vec![Value::Number(0.0)]),(5,vec![Value::Number(0.0)]),(8,vec![Value::Number(0.0)]),(11,vec![Value::Number(0.0)]),(14,vec![Value::Number(0.0)])]),
        ("SUM(IF(SelectionTick(0),M1,D1))",1u32,1u32,0u32,[(1,vec![Value::Number(1.0)]),(2,vec![Value::Number(1.0)]),(3,vec![Value::Number(1.0)]),(4,vec![Value::Number(1.0)]),(5,vec![Value::Number(1.0)]),(6,vec![Value::Number(1.0)])]),
        ("SUM(IF(SelectionTick(1),M1,D1))",1u32,1u32,0u32,[(1,vec![Value::Number(0.0)]),(2,vec![Value::Number(0.0)]),(5,vec![Value::Number(0.0)]),(8,vec![Value::Number(0.0)]),(11,vec![Value::Number(0.0)]),(14,vec![Value::Number(0.0)])]),
        ("LET(r,SUM(M1),SelectionTick(5))",1u32,1u32,0u32,[(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)])]),
        ("SUM(M1)+1",1u32,1u32,0u32,[(0,vec![Value::Number(1.0)]),(0,vec![Value::Number(1.0)]),(0,vec![Value::Number(1.0)]),(0,vec![Value::Number(1.0)]),(0,vec![Value::Number(1.0)]),(0,vec![Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(0),M1:M3,D1),SUM(r))",1u32,3u32,0u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(5,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(6,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),M1:M3,D1),SUM(r))",1u32,3u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("SUM(IF(SelectionTick(0),M1:M3,D1))",1u32,3u32,0u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(5,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(6,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("SUM(IF(SelectionTick(1),M1:M3,D1))",1u32,3u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,SUM(M1:M3),SelectionTick(5))",1u32,3u32,0u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("SUM(M1:M3)+1",1u32,3u32,0u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(0),M1:O1,D1),SUM(r))",3u32,1u32,0u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(5,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(6,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),M1:O1,D1),SUM(r))",3u32,1u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("SUM(IF(SelectionTick(0),M1:O1,D1))",3u32,1u32,0u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(5,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(6,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("SUM(IF(SelectionTick(1),M1:O1,D1))",3u32,1u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,SUM(M1:O1),SelectionTick(5))",3u32,1u32,0u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("SUM(M1:O1)+1",3u32,1u32,0u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(0),M1:N2,D1),SUM(r))",2u32,2u32,0u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(5,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(6,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),M1:N2,D1),SUM(r))",2u32,2u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("SUM(IF(SelectionTick(0),M1:N2,D1))",2u32,2u32,0u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(5,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(6,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("SUM(IF(SelectionTick(1),M1:N2,D1))",2u32,2u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,SUM(M1:N2),SelectionTick(5))",2u32,2u32,0u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("SUM(M1:N2)+1",2u32,2u32,0u32,[(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(0,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),M1,M1),SUM(r))",1u32,1u32,0u32,[(1,vec![Value::Number(0.0)]),(2,vec![Value::Number(0.0)]),(5,vec![Value::Number(0.0)]),(8,vec![Value::Number(0.0)]),(11,vec![Value::Number(0.0)]),(14,vec![Value::Number(0.0)])]),
        ("LET(r,CHOOSE(SelectionTick(1),M1,D1),SUM(r))",1u32,1u32,0u32,[(1,vec![Value::Number(0.0)]),(2,vec![Value::Number(0.0)]),(5,vec![Value::Number(0.0)]),(8,vec![Value::Number(0.0)]),(11,vec![Value::Number(0.0)]),(14,vec![Value::Number(0.0)])]),
        ("LET(r,CHOOSE(SelectionTick(2),M1,D1),SUM(r))",1u32,1u32,0u32,[(1,vec![Value::Number(1.0)]),(2,vec![Value::Number(1.0)]),(3,vec![Value::Number(1.0)]),(4,vec![Value::Number(1.0)]),(5,vec![Value::Number(1.0)]),(6,vec![Value::Number(1.0)])]),
        ("SUM(IF(SelectionTick(1),M1,D1))+SelectionTick(5)",1u32,1u32,0u32,[(2,vec![Value::Number(5.0)]),(3,vec![Value::Number(5.0)]),(6,vec![Value::Number(5.0)]),(9,vec![Value::Number(5.0)]),(12,vec![Value::Number(5.0)]),(15,vec![Value::Number(5.0)])]),
        ("SUM(IF(SelectionTick(0),M1,D1))+SelectionTick(5)",1u32,1u32,0u32,[(2,vec![Value::Number(6.0)]),(4,vec![Value::Number(6.0)]),(6,vec![Value::Number(6.0)]),(8,vec![Value::Number(6.0)]),(10,vec![Value::Number(6.0)]),(12,vec![Value::Number(6.0)])]),
        ("LET(r,SUM(M1),IF(SelectionTick(0),5,5))",1u32,1u32,0u32,[(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)]),(1,vec![Value::Number(5.0)])]),
        ("LET(r,SUM(M1),IF(SelectionTick(1),M1,D1))",1u32,1u32,0u32,[(1,vec![Value::Number(0.0)]),(1,vec![Value::Number(0.0)]),(1,vec![Value::Number(0.0)]),(1,vec![Value::Number(0.0)]),(1,vec![Value::Number(0.0)]),(1,vec![Value::Number(0.0)])]),
        ("LET(r,IF(SelectionTick(0),SUM(M1),1),5)",1u32,1u32,0u32,[(1,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(3,vec![Value::Number(5.0)]),(4,vec![Value::Number(5.0)]),(5,vec![Value::Number(5.0)]),(6,vec![Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(1),SUM(M1),1),5)",1u32,1u32,0u32,[(1,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(5,vec![Value::Number(5.0)]),(8,vec![Value::Number(5.0)]),(11,vec![Value::Number(5.0)]),(14,vec![Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(1),M1,D1),SUM(r)+SelectionTick(5))",1u32,1u32,0u32,[(2,vec![Value::Number(5.0)]),(3,vec![Value::Number(5.0)]),(6,vec![Value::Number(5.0)]),(9,vec![Value::Number(5.0)]),(12,vec![Value::Number(5.0)]),(15,vec![Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(1),M1:M3,M1:M3),SUM(r))",1u32,3u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,CHOOSE(SelectionTick(1),M1:M3,D1),SUM(r))",1u32,3u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,CHOOSE(SelectionTick(2),M1:M3,D1),SUM(r))",1u32,3u32,0u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(5,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(6,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("SUM(IF(SelectionTick(1),M1:M3,D1))+SelectionTick(5)",1u32,3u32,0u32,[(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(6,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(9,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(12,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(15,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("SUM(IF(SelectionTick(0),M1:M3,D1))+SelectionTick(5)",1u32,3u32,0u32,[(2,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(4,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(6,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(8,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(10,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(12,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)])]),
        ("LET(r,SUM(M1:M3),IF(SelectionTick(0),5,5))",1u32,3u32,0u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SUM(M1:M3),IF(SelectionTick(1),M1:M3,D1))",1u32,3u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,IF(SelectionTick(0),SUM(M1:M3),1),5)",1u32,3u32,0u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(4,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(5,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(6,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(1),SUM(M1:M3),1),5)",1u32,3u32,0u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(5,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(8,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(11,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(14,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(1),M1:M3,D1),SUM(r)+SelectionTick(5))",1u32,3u32,0u32,[(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(6,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(9,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(12,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(15,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(1),M1:O1,M1:O1),SUM(r))",3u32,1u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,CHOOSE(SelectionTick(1),M1:O1,D1),SUM(r))",3u32,1u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,CHOOSE(SelectionTick(2),M1:O1,D1),SUM(r))",3u32,1u32,0u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(5,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(6,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("SUM(IF(SelectionTick(1),M1:O1,D1))+SelectionTick(5)",3u32,1u32,0u32,[(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(6,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(9,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(12,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(15,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("SUM(IF(SelectionTick(0),M1:O1,D1))+SelectionTick(5)",3u32,1u32,0u32,[(2,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(4,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(6,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(8,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(10,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(12,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)])]),
        ("LET(r,SUM(M1:O1),IF(SelectionTick(0),5,5))",3u32,1u32,0u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SUM(M1:O1),IF(SelectionTick(1),M1:O1,D1))",3u32,1u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,IF(SelectionTick(0),SUM(M1:O1),1),5)",3u32,1u32,0u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(4,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(5,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(6,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(1),SUM(M1:O1),1),5)",3u32,1u32,0u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(5,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(8,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(11,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(14,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(1),M1:O1,D1),SUM(r)+SelectionTick(5))",3u32,1u32,0u32,[(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(6,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(9,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(12,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(15,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(1),M1:N2,M1:N2),SUM(r))",2u32,2u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,CHOOSE(SelectionTick(1),M1:N2,D1),SUM(r))",2u32,2u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,CHOOSE(SelectionTick(2),M1:N2,D1),SUM(r))",2u32,2u32,0u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(5,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(6,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("SUM(IF(SelectionTick(1),M1:N2,D1))+SelectionTick(5)",2u32,2u32,0u32,[(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(6,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(9,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(12,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(15,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("SUM(IF(SelectionTick(0),M1:N2,D1))+SelectionTick(5)",2u32,2u32,0u32,[(2,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(4,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(6,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(8,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(10,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)]),(12,vec![Value::Number(6.0),Value::Number(6.0),Value::Number(6.0),Value::Number(6.0)])]),
        ("LET(r,SUM(M1:N2),IF(SelectionTick(0),5,5))",2u32,2u32,0u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,SUM(M1:N2),IF(SelectionTick(1),M1:N2,D1))",2u32,2u32,0u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,IF(SelectionTick(0),SUM(M1:N2),1),5)",2u32,2u32,0u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(4,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(5,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(6,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(1),SUM(M1:N2),1),5)",2u32,2u32,0u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(5,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(8,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(11,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(14,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(1),M1:N2,D1),SUM(r)+SelectionTick(5))",2u32,2u32,0u32,[(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(3,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(6,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(9,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(12,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(15,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(0),T10,D1),SUM(r))",1u32,1u32,1u32,[(1,vec![Value::Number(1.0)]),(2,vec![Value::Number(1.0)]),(3,vec![Value::Number(1.0)]),(4,vec![Value::Number(1.0)]),(5,vec![Value::Number(1.0)]),(6,vec![Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),T10,D1),SUM(r))",1u32,1u32,1u32,[(1,vec![Value::Number(0.0)]),(2,vec![Value::Number(0.0)]),(5,vec![Value::Number(0.0)]),(8,vec![Value::Number(0.0)]),(11,vec![Value::Number(0.0)]),(14,vec![Value::Number(0.0)])]),
        ("LET(r,SUM(T10),SelectionTick(5))",1u32,1u32,1u32,[(1,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(0),T10,D1),SUM(r))",1u32,1u32,2u32,[(1,vec![Value::Number(1.0)]),(2,vec![Value::Number(1.0)]),(3,vec![Value::Number(1.0)]),(4,vec![Value::Number(1.0)]),(5,vec![Value::Number(1.0)]),(6,vec![Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),T10,D1),SUM(r))",1u32,1u32,2u32,[(1,vec![Value::Number(0.0)]),(2,vec![Value::Number(0.0)]),(5,vec![Value::Number(0.0)]),(8,vec![Value::Number(0.0)]),(11,vec![Value::Number(0.0)]),(14,vec![Value::Number(0.0)])]),
        ("LET(r,SUM(T10),SelectionTick(5))",1u32,1u32,2u32,[(1,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(0),T10,D1),SUM(r))",1u32,1u32,4u32,[(1,vec![Value::Number(1.0)]),(2,vec![Value::Number(1.0)]),(3,vec![Value::Number(1.0)]),(4,vec![Value::Number(1.0)]),(5,vec![Value::Number(1.0)]),(6,vec![Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),T10,D1),SUM(r))",1u32,1u32,4u32,[(1,vec![Value::Number(0.0)]),(2,vec![Value::Number(0.0)]),(5,vec![Value::Number(0.0)]),(8,vec![Value::Number(0.0)]),(11,vec![Value::Number(0.0)]),(14,vec![Value::Number(0.0)])]),
        ("LET(r,SUM(T10),SelectionTick(5))",1u32,1u32,4u32,[(1,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(0),T10,D1),SUM(r))",1u32,1u32,8u32,[(1,vec![Value::Number(1.0)]),(2,vec![Value::Number(1.0)]),(3,vec![Value::Number(1.0)]),(4,vec![Value::Number(1.0)]),(5,vec![Value::Number(1.0)]),(6,vec![Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),T10,D1),SUM(r))",1u32,1u32,8u32,[(1,vec![Value::Number(0.0)]),(2,vec![Value::Number(0.0)]),(5,vec![Value::Number(0.0)]),(8,vec![Value::Number(0.0)]),(11,vec![Value::Number(0.0)]),(14,vec![Value::Number(0.0)])]),
        ("LET(r,SUM(T10),SelectionTick(5))",1u32,1u32,8u32,[(1,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)]),(2,vec![Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(0),T10,D1),SUM(r))",1u32,3u32,1u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(5,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(6,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),T10,D1),SUM(r))",1u32,3u32,1u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,SUM(T10),SelectionTick(5))",1u32,3u32,1u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(0),T10,D1),SUM(r))",1u32,3u32,2u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(5,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(6,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),T10,D1),SUM(r))",1u32,3u32,2u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,SUM(T10),SelectionTick(5))",1u32,3u32,2u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(0),T10,D1),SUM(r))",1u32,3u32,4u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(5,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(6,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),T10,D1),SUM(r))",1u32,3u32,4u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,SUM(T10),SelectionTick(5))",1u32,3u32,4u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
        ("LET(r,IF(SelectionTick(0),T10,D1),SUM(r))",1u32,3u32,8u32,[(1,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(2,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(3,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(4,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(5,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)]),(6,vec![Value::Number(1.0),Value::Number(1.0),Value::Number(1.0)])]),
        ("LET(r,IF(SelectionTick(1),T10,D1),SUM(r))",1u32,3u32,8u32,[(1,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(2,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(5,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(8,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(11,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)]),(14,vec![Value::Number(0.0),Value::Number(0.0),Value::Number(0.0)])]),
        ("LET(r,SUM(T10),SelectionTick(5))",1u32,3u32,8u32,[(1,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)]),(2,vec![Value::Number(5.0),Value::Number(5.0),Value::Number(5.0)])]),
    ];
    let mut mismatches = Vec::new();
    for (formula, width, height, chain, stages) in cases {
        for version in [1u8, 2] { for dirty in [false, true] {
            let count = Rc::new(RefCell::new(0u32));
            let counted = count.clone();
            set_user_function_hook(Some(Box::new(move |name, args, _| {
                if !name.eq_ignore_ascii_case("SelectionTick") { return None; }
                *counted.borrow_mut() += 1;
                let value = match args.first() {
                    Some(UserArg::Value(value)) => value.clone(),
                    Some(UserArg::Cells { values, .. }) => values.first().cloned().unwrap_or(Value::Blank),
                    Some(UserArg::Block(block)) => block.cells.first().cloned().unwrap_or(Value::Blank),
                    _ => Value::Blank,
                };
                Some(Arg::Value(value))
            })));
            oxicells_calc::engine::set_user_function_wait_for_dependencies(false);
            oxicells_calc::engine::set_user_function_full_rebuild(false);
            oxicells_calc::engine::set_user_function_full_calculation(false);
            let mut held = book(); held.set_compatibility_version(version);
            for (address, value) in [("D1", 1.), ("D2", 2.), ("D3", 3.)] {
                held.set_value("S", address, Value::Number(value)).unwrap();
            }
            let mut external = Vec::new();
            for link in 0..chain {
                let address = format!("T{}", 10 + link);
                let source = if link + 1 < chain { format!("T{}", 11 + link) } else { "M1".to_string() };
                held.set_formula("S", &address, &source).unwrap();
                external.push(("S".to_string(), (19, 9 + link)));
            }
            if !external.is_empty() { held.recalculate_these(&external); }
            let mut wanted = Vec::new();
            for row in 0..height { for col in 0..width {
                let coord = (12 + col, row);
                let address = oxicells_calc::CellRef::new(coord.0, coord.1).to_a1();
                held.set_array_member("S", &address, formula, (col, row)).unwrap();
                wanted.push(("S".to_string(), coord));
            }}
            let mut reverse = wanted.clone(); reverse.reverse();
            held.set_calculation_order(&reverse);
            let all: Vec<_> = wanted.iter().chain(external.iter()).cloned().collect();
            for (stage, (expected_calls, expected_values)) in stages.iter().enumerate() {
                held.set_chain_probes(stage != 0);
                oxicells_calc::engine::set_user_function_wait_for_dependencies(true);
                oxicells_calc::engine::set_user_function_full_calculation(stage >= 3);
                oxicells_calc::engine::set_user_function_full_rebuild(stage >= 4);
                match stage {
                    0 => { held.recalculate_array_entry(&wanted); }
                    1 => { held.recalculate_these(&wanted); }
                    2 if dirty => { held.recalculate_dirty(&all, None); }
                    _ => { held.recalculate(); }
                }
                let values: Vec<_> = wanted.iter().map(|(_, coord)|
                    held.value("S", &oxicells_calc::CellRef::new(coord.0, coord.1).to_a1())).collect();
                let calls = *count.borrow();
                if values != *expected_values || calls != *expected_calls {
                    mismatches.push(format!("{formula}, {width}x{height}, chain {chain}, version {version}, dirty {dirty}, stage {stage}: calls {calls} != {expected_calls}, {values:?} != {expected_values:?}"));
                }
            }
        }}
    }
    assert!(mismatches.is_empty(), "{} Native calculation-scope mismatches:\n{}", mismatches.len(), mismatches.join("\n"));
}

#[test]
fn conditional_legacy_array_cycles_resume_through_the_core_bridge() {
    use oxicells_core::ir::{Cell, CellValue, Row, Spill};
    struct Restore;
    impl Drop for Restore {
        fn drop(&mut self) {
            set_user_function_hook(None);
            oxicells_calc::engine::set_user_function_wait_for_dependencies(false);
            oxicells_calc::engine::set_user_function_full_rebuild(false);
            oxicells_calc::engine::set_user_function_full_calculation(false);
        }
    }
    let _restore = Restore;
    let cases = [
        ("LET(r,IF(SelectionTick(0),M1,D1),SUM(r))",1u32,1u32,[(1,vec![1.0]),(2,vec![1.0]),(3,vec![1.0]),(4,vec![1.0]),(5,vec![1.0]),(6,vec![1.0])]),
        ("LET(r,IF(SelectionTick(1),M1,D1),SUM(r))",1u32,1u32,[(1,vec![0.0]),(2,vec![0.0]),(5,vec![0.0]),(8,vec![0.0]),(11,vec![0.0]),(14,vec![0.0])]),
        ("LET(r,SUM(M1),SelectionTick(5))",1u32,1u32,[(1,vec![5.0]),(1,vec![5.0]),(1,vec![5.0]),(1,vec![5.0]),(1,vec![5.0]),(1,vec![5.0])]),
        ("LET(r,IF(SelectionTick(0),M1:M3,D1),SUM(r))",1u32,3u32,[(1,vec![1.0,1.0,1.0]),(2,vec![1.0,1.0,1.0]),(3,vec![1.0,1.0,1.0]),(4,vec![1.0,1.0,1.0]),(5,vec![1.0,1.0,1.0]),(6,vec![1.0,1.0,1.0])]),
        ("LET(r,IF(SelectionTick(1),M1:M3,D1),SUM(r))",1u32,3u32,[(1,vec![0.0,0.0,0.0]),(2,vec![0.0,0.0,0.0]),(5,vec![0.0,0.0,0.0]),(8,vec![0.0,0.0,0.0]),(11,vec![0.0,0.0,0.0]),(14,vec![0.0,0.0,0.0])]),
        ("LET(r,SUM(M1:M3),SelectionTick(5))",1u32,3u32,[(1,vec![5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0])]),
        ("LET(r,IF(SelectionTick(0),M1:O1,D1),SUM(r))",3u32,1u32,[(1,vec![1.0,1.0,1.0]),(2,vec![1.0,1.0,1.0]),(3,vec![1.0,1.0,1.0]),(4,vec![1.0,1.0,1.0]),(5,vec![1.0,1.0,1.0]),(6,vec![1.0,1.0,1.0])]),
        ("LET(r,IF(SelectionTick(1),M1:O1,D1),SUM(r))",3u32,1u32,[(1,vec![0.0,0.0,0.0]),(2,vec![0.0,0.0,0.0]),(5,vec![0.0,0.0,0.0]),(8,vec![0.0,0.0,0.0]),(11,vec![0.0,0.0,0.0]),(14,vec![0.0,0.0,0.0])]),
        ("LET(r,SUM(M1:O1),SelectionTick(5))",3u32,1u32,[(1,vec![5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0])]),
        ("LET(r,IF(SelectionTick(0),M1:N2,D1),SUM(r))",2u32,2u32,[(1,vec![1.0,1.0,1.0,1.0]),(2,vec![1.0,1.0,1.0,1.0]),(3,vec![1.0,1.0,1.0,1.0]),(4,vec![1.0,1.0,1.0,1.0]),(5,vec![1.0,1.0,1.0,1.0]),(6,vec![1.0,1.0,1.0,1.0])]),
        ("LET(r,IF(SelectionTick(1),M1:N2,D1),SUM(r))",2u32,2u32,[(1,vec![0.0,0.0,0.0,0.0]),(2,vec![0.0,0.0,0.0,0.0]),(5,vec![0.0,0.0,0.0,0.0]),(8,vec![0.0,0.0,0.0,0.0]),(11,vec![0.0,0.0,0.0,0.0]),(14,vec![0.0,0.0,0.0,0.0])]),
        ("LET(r,SUM(M1:N2),SelectionTick(5))",2u32,2u32,[(1,vec![5.0,5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0,5.0]),(1,vec![5.0,5.0,5.0,5.0])]),
        ("SUM(IF(SelectionTick(1),M1,D1))+SelectionTick(5)",1u32,1u32,[(2,vec![5.0]),(3,vec![5.0]),(6,vec![5.0]),(9,vec![5.0]),(12,vec![5.0]),(15,vec![5.0])]),
        ("SUM(IF(SelectionTick(1),M1:M3,D1))+SelectionTick(5)",1u32,3u32,[(2,vec![5.0,5.0,5.0]),(3,vec![5.0,5.0,5.0]),(6,vec![5.0,5.0,5.0]),(9,vec![5.0,5.0,5.0]),(12,vec![5.0,5.0,5.0]),(15,vec![5.0,5.0,5.0])]),
        ("SUM(IF(SelectionTick(1),M1:O1,D1))+SelectionTick(5)",3u32,1u32,[(2,vec![5.0,5.0,5.0]),(3,vec![5.0,5.0,5.0]),(6,vec![5.0,5.0,5.0]),(9,vec![5.0,5.0,5.0]),(12,vec![5.0,5.0,5.0]),(15,vec![5.0,5.0,5.0])]),
        ("SUM(IF(SelectionTick(1),M1:N2,D1))+SelectionTick(5)",2u32,2u32,[(2,vec![5.0,5.0,5.0,5.0]),(3,vec![5.0,5.0,5.0,5.0]),(6,vec![5.0,5.0,5.0,5.0]),(9,vec![5.0,5.0,5.0,5.0]),(12,vec![5.0,5.0,5.0,5.0]),(15,vec![5.0,5.0,5.0,5.0])]),
    ];
    let mut mismatches = Vec::new();
    for (formula, width, height, stages) in cases {
        for version in [1u8, 2] {
            let mut held = oxicells_core::parse_xlsx(include_bytes!("../../oxicells-core/tests/fixtures/hidden_rows_cols.xlsx")).unwrap();
            held.sheets.truncate(1); held.sheets[0].name = "S".into();
            held.sheets[0].rows.clear(); held.sheets[0].cached_spills = None;
            held.defined_names.clear(); held.external_books.clear();
            held.recalculation = Default::default();
            held.calculation.compatibility_version = version;
            let mut wanted = Vec::new();
            for row in 1..=height.max(3) {
                let mut cells = vec![Cell { col: 3, value: CellValue::Number(row as f64), formula: None,
                    array_block: None, style: Default::default(), formula_utf16: None, runs: Vec::new(), spill: Spill::Scalar }];
                if row <= height { for col in 0..width {
                    cells.push(Cell { col: 12 + col, value: CellValue::Empty, formula: Some(formula.into()),
                        array_block: Some((1,12,height,12+width-1)), style: Default::default(),
                        formula_utf16: None, runs: Vec::new(), spill: Spill::Scalar });
                    wanted.push(("S".to_string(), (12 + col, row - 1)));
                }}
                held.sheets[0].rows.push(Row { index: row, height: None, custom_height: false, format_recorded: None, style_font: None,
                    thick_top: false, thick_bottom: false, hidden: false, cells });
            }
            let mut calls = 0u32;
            for (stage, (expected_calls, expected_values)) in stages.iter().enumerate() {
                let mut reverse = wanted.clone(); reverse.reverse(); held.recalculation.order = reverse;
                let baseline = held.clone();
                let answers = Rc::new(RefCell::new(std::collections::BTreeMap::<String, Value>::new()));
                let mut complete = false;
                oxicells_calc::engine::set_user_function_wait_for_dependencies(true);
                oxicells_calc::engine::set_user_function_full_calculation(stage >= 3);
                oxicells_calc::engine::set_user_function_full_rebuild(stage >= 4);
                for _ in 0..32 {
                    held = baseline.clone();
                    let queued = Rc::new(RefCell::new(std::collections::BTreeMap::<String, Value>::new()));
                    let queue = queued.clone(); let cached = answers.clone();
                    set_user_function_hook(Some(Box::new(move |name, args, caller| {
                        let context = oxicells_calc::engine::user_function_call_context();
                        let key = format!("{name}|{args:?}|{caller:?}|{context:?}");
                        if let Some(value) = cached.borrow().get(&key) { return Some(Arg::Value(value.clone())); }
                        let value = match args.first() {
                            Some(UserArg::Value(value)) => value.clone(),
                            Some(UserArg::Cells { values, .. }) => values.first().cloned().unwrap_or(Value::Blank),
                            Some(UserArg::Block(block)) => block.cells.first().cloned().unwrap_or(Value::Blank),
                            _ => Value::Blank,
                        };
                        queue.borrow_mut().insert(key, value);
                        defer_user_function(); Some(Arg::Value(Value::Error(ExcelError::Value)))
                    })));
                    match stage {
                        0 => { oxicells_core::formula::evaluate_workbook_formula_array_entry(&mut held, &wanted, None); }
                        1 => { oxicells_core::formula::evaluate_workbook_formula_subset(&mut held, &wanted, false, None); }
                        _ => { oxicells_core::formula::evaluate_workbook_formulas(&mut held); }
                    }
                    set_user_function_hook(None);
                    if queued.borrow().is_empty() { complete = true; break; }
                    for (key, value) in std::mem::take(&mut *queued.borrow_mut()) {
                        assert!(answers.borrow_mut().insert(key, value).is_none()); calls += 1;
                    }
                }
                let values: Vec<_> = wanted.iter().map(|(_, (col, row))|
                    held.sheets[0].rows.iter().find(|r| r.index == row + 1).unwrap().cells.iter().find(|cell| cell.col == *col).unwrap().value.clone()).collect();
                let expected: Vec<_> = expected_values.iter().map(|v| CellValue::Number(*v)).collect();
                let same = values.iter().zip(&expected).all(|(a,b)| matches!((a,b), (CellValue::Number(a),CellValue::Number(b)) if a==b));
                if !complete || calls != *expected_calls || !same {
                    mismatches.push(format!("{formula}, {width}x{height}, version {version}, stage {stage}: complete {complete}, calls {calls} != {expected_calls}, {values:?} != {expected:?}"));
                }
            }
        }
    }
    assert!(mismatches.is_empty(), "{} Native async/core mismatches:\n{}", mismatches.len(), mismatches.join("\n"));
}

#[test]
fn dynamic_index_dependencies_and_selector_stages_match_native_excel() {
    use oxicells_calc::functions::{BlockOrigin, RangeData};
    struct Restore;
    impl Drop for Restore {
        fn drop(&mut self) {
            set_user_function_hook(None);
            oxicells_calc::engine::set_user_function_wait_for_dependencies(false);
            oxicells_calc::engine::set_user_function_full_rebuild(false);
        }
    }
    let _restore = Restore;
    const CASES: &[(&str, [(u32, [Value; 3]); 4])] = &[
        ("INDEX(A1:O3,SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(D1:D3,SelectionTick(1),1)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(r,A1:O3,INDEX(r,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(r,D1:D3,INDEX(r,SelectionTick(1),1))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(A1:O3,{1;2;3},SelectionTick(4)))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,{1;2;3},SelectionTick(1)))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(A1:O3,SelectionTick({1;2;3}),4))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("SUM(INDEX(D1:D3,SelectionTick({1;2;3}),1))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("INDEX(A1:O3,SelectionTick(1),13)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("SUM(INDEX(A1:O3,{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(r,A1:O3,INDEX(r,SelectionTick(1),13))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(A1:O3,IF(SelectionTick(1),1,2),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(A1:O3,SelectionTick(1),SelectionTick(4))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(A1:O3,SelectionTick({1;2;3}),SelectionTick(4)))", [(2, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (10, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (18, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (26, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("INDEX(A1:O3+0,SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(r,A1:O3+0,INDEX(r,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(D1:D3+0,SelectionTick(1),1)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(r,D1:D3+0,INDEX(r,SelectionTick(1),1))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(IF(1,A1:O3),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(IFERROR(A1:O3,0),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(r,IF(1,A1:O3),INDEX(r,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(r,IFERROR(A1:O3,0),INDEX(r,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(XLOOKUP(1,{1},A1:O3),SelectionTick(1),4)", [(1, [Value::Error(ExcelError::Value), Value::Error(ExcelError::Value), Value::Error(ExcelError::Value)]), (2, [Value::Error(ExcelError::Value), Value::Error(ExcelError::Value), Value::Error(ExcelError::Value)]), (3, [Value::Error(ExcelError::Value), Value::Error(ExcelError::Value), Value::Error(ExcelError::Value)]), (4, [Value::Error(ExcelError::Value), Value::Error(ExcelError::Value), Value::Error(ExcelError::Value)])]),
        ("INDEX(OFFSET(A1,0,0,3,15),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(r,OFFSET(A1,0,0,3,15),INDEX(r,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(INDIRECT(\"A1:O3\"),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(r,INDIRECT(\"A1:O3\"),INDEX(r,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(A1:O3+0,{1;2;3},SelectionTick(4)))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("LET(r,A1:O3+0,SUM(INDEX(r,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("INDEX(SUM(M1:M3),SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(r,SUM(M1:M3),INDEX(r,SelectionTick(1),1))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(SelectionTick(SUM(M1:M3)),SelectionTick(1),1)", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(SUM(D1:D3),SelectionTick(1),1)", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (2, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (3, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (4, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("LET(r,SUM(D1:D3),INDEX(r,SelectionTick(1),1))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (2, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (3, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (4, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)])]),
        ("INDEX(A1:O3*2,SelectionTick(1),4)", [(1, [Value::Number(2.0), Value::Number(2.0), Value::Number(2.0)]), (2, [Value::Number(2.0), Value::Number(2.0), Value::Number(2.0)]), (3, [Value::Number(2.0), Value::Number(2.0), Value::Number(2.0)]), (4, [Value::Number(2.0), Value::Number(2.0), Value::Number(2.0)])]),
        ("INDEX(SelectionTick(1)+(A1:O3+0),SelectionTick(1),4)", [(2, [Value::Number(2.0), Value::Number(2.0), Value::Number(2.0)]), (4, [Value::Number(2.0), Value::Number(2.0), Value::Number(2.0)]), (6, [Value::Number(2.0), Value::Number(2.0), Value::Number(2.0)]), (8, [Value::Number(2.0), Value::Number(2.0), Value::Number(2.0)])]),
        ("LET(r,A1:O3+SelectionTick(1),INDEX(r,SelectionTick(1),4))", [(2, [Value::Number(2.0), Value::Number(2.0), Value::Number(2.0)]), (4, [Value::Number(2.0), Value::Number(2.0), Value::Number(2.0)]), (6, [Value::Number(2.0), Value::Number(2.0), Value::Number(2.0)]), (8, [Value::Number(2.0), Value::Number(2.0), Value::Number(2.0)])]),
        ("INDEX(CHOOSE(1,A1:O3),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(r,CHOOSE(1,A1:O3),INDEX(r,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(r,A1:O3+0,SelectionTick(7))", [(1, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)]), (2, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)]), (3, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)]), (4, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)])]),
        ("LET(r,SUM(M1:M3),SelectionTick(7))", [(1, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)]), (1, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)]), (1, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)]), (1, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)])]),
        ("LET(r,IFERROR(A1:O3,0),SelectionTick(7))", [(1, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)]), (2, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)]), (3, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)]), (4, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)])]),
        ("LET(r,A1:O3+0,ROWS(r)+SelectionTick(7))", [(1, [Value::Number(10.0), Value::Number(10.0), Value::Number(10.0)]), (2, [Value::Number(10.0), Value::Number(10.0), Value::Number(10.0)]), (3, [Value::Number(10.0), Value::Number(10.0), Value::Number(10.0)]), (4, [Value::Number(10.0), Value::Number(10.0), Value::Number(10.0)])]),
        ("INDEX(SUM(M1:M3)+0,SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(-SUM(M1:M3),SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(IF(1,SUM(M1:M3),0),SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(IFERROR(SUM(M1:M3),0),SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(r,A1:O3+0,INDEX(r,1,4))", [(0, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(A1:O3+0,1,4)", [(0, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(r,IFERROR(A1:O3,0),INDEX(r,1,4))", [(0, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (0, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(A1:O1+0,SelectionTick(4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(A1:O1+0,SelectionTick(13))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(r,A1:O1+0,INDEX(r,SelectionTick(4)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(-A1:O3,SelectionTick(1),4)", [(1, [Value::Number(-1.0), Value::Number(-1.0), Value::Number(-1.0)]), (2, [Value::Number(-1.0), Value::Number(-1.0), Value::Number(-1.0)]), (3, [Value::Number(-1.0), Value::Number(-1.0), Value::Number(-1.0)]), (4, [Value::Number(-1.0), Value::Number(-1.0), Value::Number(-1.0)])]),
        ("INDEX(N(A1:O3),SelectionTick(1),4)", [(1, [Value::Error(ExcelError::Ref), Value::Error(ExcelError::Ref), Value::Error(ExcelError::Ref)]), (2, [Value::Error(ExcelError::Ref), Value::Error(ExcelError::Ref), Value::Error(ExcelError::Ref)]), (3, [Value::Error(ExcelError::Ref), Value::Error(ExcelError::Ref), Value::Error(ExcelError::Ref)]), (4, [Value::Error(ExcelError::Ref), Value::Error(ExcelError::Ref), Value::Error(ExcelError::Ref)])]),
        ("INDEX(IF({1;1;1},A1:O3,0),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(IFERROR(A1:O3,M1:M3),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(IFERROR(1/(A1:O3+0),M1:M3),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(ABS(A1:O3),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(TRANSPOSE(A1:O3),SelectionTick(4),1)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(LET(r,A1:O3+0,r),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(A1:O3,{1,2,3},SelectionTick(4)))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (9, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (13, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(A1:O3,SelectionTick({1;2;3}),SelectionTick(13)))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (16, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (30, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (50, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(IFERROR(A1:O3,SUM(M1:M3)),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(IF(1,A1:O3,SUM(M1:M3)),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(r,IFERROR(A1:O3,SUM(M1:M3)),INDEX(r,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(A1:O3+0,SelectionTick(1),4))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(A1:O3+0,SelectionTick(1),13))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(IF({1;1;1},A1:O3,0),SelectionTick(1),13)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ABS(A1:O3),SelectionTick(1),13)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(TRANSPOSE(A1:O3),SelectionTick(13),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(N(A1:O3+0),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(N(A1:O3+0),SelectionTick(1),13)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("SelectionTick(INDEX(A1:O3+0,{1;2;3},4))", [(1, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (4, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (7, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (10, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("INDEX(IFERROR(A1:O3+0,0),SelectionTick(1),13)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(A1:O3+0,SelectionTick(1.9),4.9)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(A1:O3+0,SelectionTick(3.9),13.9)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(IF({1;0;1},A1:O3,0),SelectionTick(2),13)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(IF(A1:O3>0,A1:O3,0),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(IF(A1:O3>0,1,0),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(IF(A1:O3>0,A1:O3,0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("INDEX(CHOOSE({1;1;1},A1:O3),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(CHOOSE({1;1;1},A1:O3),SelectionTick(1),13)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ABS(IFERROR(A1:O3,0)),SelectionTick(1),13)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(IFNA(A1:O3,0),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(t,A1:O3>0,INDEX(IF(t,A1:O3,0),SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(IF(M1,1,0),SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(IF(M1+1,1,0),SelectionTick(1),1)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(IF(SUM(M1:M3),1,0),SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(t,M1,INDEX(IF(t,1,0),SelectionTick(1),1))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(t,IF(M1,1,0),SelectionTick(7))", [(1, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)]), (1, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)]), (1, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)]), (1, [Value::Number(7.0), Value::Number(7.0), Value::Number(7.0)])]),
        ("INDEX(CHOOSE(M1+1,1,0),SelectionTick(1),1)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(IF(A1:O3>0,A1:O3,0),SelectionTick(1),13)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("SUM(INDEX(CHOOSE({1;1;1},A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(IF(M1+{1},1,0),SelectionTick(1),1)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(CHOOSE(M1+{1},1,0),SelectionTick(1),1)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(IF(D1:D3,A1:O3,0),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(IFERROR(IF(A1:O3>0,A1:O3,0),0),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("TYPE(INDEX(A1:O3+0,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("TYPE(INDEX(A1:O3+0,SelectionTick(1),13))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("TYPE(INDEX(A1:O3,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(A1:O3,SelectionTick(1),4))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("TYPE(INDEX(A1:O3,{1;2;3},SelectionTick(4)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,A1:O3,INDEX(src_,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,A1:O3,TYPE(INDEX(src_,SelectionTick(1),4)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,A1:O3,SelectionTick(INDEX(src_,SelectionTick(1),4)))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,A1:O3,SUM(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("LET(src_,A1:O3,TYPE(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(A1:O3,SelectionTick(1),13))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(A1:O3,SelectionTick(1),13))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (12, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("TYPE(INDEX(A1:O3,{1;2;3},SelectionTick(13)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,A1:O3,INDEX(src_,SelectionTick(1),13))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,A1:O3,TYPE(INDEX(src_,SelectionTick(1),13)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,A1:O3,SelectionTick(INDEX(src_,SelectionTick(1),13)))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (12, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,A1:O3,SUM(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,A1:O3,TYPE(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(A1:O3+0,{1;2;3},SelectionTick(4)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,A1:O3+0,INDEX(src_,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,A1:O3+0,TYPE(INDEX(src_,SelectionTick(1),4)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,A1:O3+0,SelectionTick(INDEX(src_,SelectionTick(1),4)))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,A1:O3+0,SUM(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("LET(src_,A1:O3+0,TYPE(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("INDEX(A1:O3+0,SelectionTick(1),13)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("SUM(INDEX(A1:O3+0,{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("TYPE(INDEX(A1:O3+0,{1;2;3},SelectionTick(13)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,A1:O3+0,INDEX(src_,SelectionTick(1),13))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,A1:O3+0,TYPE(INDEX(src_,SelectionTick(1),13)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,A1:O3+0,SelectionTick(INDEX(src_,SelectionTick(1),13)))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,A1:O3+0,SUM(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,A1:O3+0,TYPE(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(IF(A1:O3>0,A1:O3,0),SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(IF(A1:O3>0,A1:O3,0),SelectionTick(1),4))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("TYPE(INDEX(IF(A1:O3>0,A1:O3,0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,IF(A1:O3>0,A1:O3,0),INDEX(src_,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,IF(A1:O3>0,A1:O3,0),TYPE(INDEX(src_,SelectionTick(1),4)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,IF(A1:O3>0,A1:O3,0),SelectionTick(INDEX(src_,SelectionTick(1),4)))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,IF(A1:O3>0,A1:O3,0),SUM(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("LET(src_,IF(A1:O3>0,A1:O3,0),TYPE(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(IF(A1:O3>0,A1:O3,0),SelectionTick(1),13))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(IF(A1:O3>0,A1:O3,0),SelectionTick(1),13))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("SUM(INDEX(IF(A1:O3>0,A1:O3,0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("TYPE(INDEX(IF(A1:O3>0,A1:O3,0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,IF(A1:O3>0,A1:O3,0),INDEX(src_,SelectionTick(1),13))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,IF(A1:O3>0,A1:O3,0),TYPE(INDEX(src_,SelectionTick(1),13)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,IF(A1:O3>0,A1:O3,0),SelectionTick(INDEX(src_,SelectionTick(1),13)))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,IF(A1:O3>0,A1:O3,0),SUM(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,IF(A1:O3>0,A1:O3,0),TYPE(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("INDEX(IF(A1:O3,A1:O3,0),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("TYPE(INDEX(IF(A1:O3,A1:O3,0),SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(IF(A1:O3,A1:O3,0),SelectionTick(1),4))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(IF(A1:O3,A1:O3,0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("TYPE(INDEX(IF(A1:O3,A1:O3,0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,IF(A1:O3,A1:O3,0),INDEX(src_,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,IF(A1:O3,A1:O3,0),TYPE(INDEX(src_,SelectionTick(1),4)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,IF(A1:O3,A1:O3,0),SelectionTick(INDEX(src_,SelectionTick(1),4)))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,IF(A1:O3,A1:O3,0),SUM(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("LET(src_,IF(A1:O3,A1:O3,0),TYPE(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("INDEX(IF(A1:O3,A1:O3,0),SelectionTick(1),13)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("TYPE(INDEX(IF(A1:O3,A1:O3,0),SelectionTick(1),13))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(IF(A1:O3,A1:O3,0),SelectionTick(1),13))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("SUM(INDEX(IF(A1:O3,A1:O3,0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("TYPE(INDEX(IF(A1:O3,A1:O3,0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,IF(A1:O3,A1:O3,0),INDEX(src_,SelectionTick(1),13))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,IF(A1:O3,A1:O3,0),TYPE(INDEX(src_,SelectionTick(1),13)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,IF(A1:O3,A1:O3,0),SelectionTick(INDEX(src_,SelectionTick(1),13)))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,IF(A1:O3,A1:O3,0),SUM(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,IF(A1:O3,A1:O3,0),TYPE(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(CHOOSE({1;1;1},A1:O3),SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(CHOOSE({1;1;1},A1:O3),SelectionTick(1),4))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(CHOOSE({1;1;1},A1:O3),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("TYPE(INDEX(CHOOSE({1;1;1},A1:O3),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,CHOOSE({1;1;1},A1:O3),INDEX(src_,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,CHOOSE({1;1;1},A1:O3),TYPE(INDEX(src_,SelectionTick(1),4)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,CHOOSE({1;1;1},A1:O3),SelectionTick(INDEX(src_,SelectionTick(1),4)))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,CHOOSE({1;1;1},A1:O3),SUM(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("LET(src_,CHOOSE({1;1;1},A1:O3),TYPE(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(CHOOSE({1;1;1},A1:O3),SelectionTick(1),13))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(CHOOSE({1;1;1},A1:O3),SelectionTick(1),13))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("TYPE(INDEX(CHOOSE({1;1;1},A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,CHOOSE({1;1;1},A1:O3),INDEX(src_,SelectionTick(1),13))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,CHOOSE({1;1;1},A1:O3),TYPE(INDEX(src_,SelectionTick(1),13)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,CHOOSE({1;1;1},A1:O3),SelectionTick(INDEX(src_,SelectionTick(1),13)))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,CHOOSE({1;1;1},A1:O3),SUM(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,CHOOSE({1;1;1},A1:O3),TYPE(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(ABS(A1:O3),SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(ABS(A1:O3),SelectionTick(1),4))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(ABS(A1:O3),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("TYPE(INDEX(ABS(A1:O3),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,ABS(A1:O3),INDEX(src_,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,ABS(A1:O3),TYPE(INDEX(src_,SelectionTick(1),4)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,ABS(A1:O3),SelectionTick(INDEX(src_,SelectionTick(1),4)))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,ABS(A1:O3),SUM(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("LET(src_,ABS(A1:O3),TYPE(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(ABS(A1:O3),SelectionTick(1),13))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(ABS(A1:O3),SelectionTick(1),13))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("SUM(INDEX(ABS(A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("TYPE(INDEX(ABS(A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,ABS(A1:O3),INDEX(src_,SelectionTick(1),13))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,ABS(A1:O3),TYPE(INDEX(src_,SelectionTick(1),13)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,ABS(A1:O3),SelectionTick(INDEX(src_,SelectionTick(1),13)))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,ABS(A1:O3),SUM(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,ABS(A1:O3),TYPE(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("INDEX(ISNUMBER(A1:O3),SelectionTick(1),4)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (4, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (7, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (10, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("TYPE(INDEX(ISNUMBER(A1:O3),SelectionTick(1),4))", [(1, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (4, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (7, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (10, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)])]),
        ("SelectionTick(INDEX(ISNUMBER(A1:O3),SelectionTick(1),4))", [(2, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (8, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (14, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (20, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("SUM(INDEX(ISNUMBER(A1:O3),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (7, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (10, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("TYPE(INDEX(ISNUMBER(A1:O3),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (7, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (10, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)])]),
        ("LET(src_,ISNUMBER(A1:O3),INDEX(src_,SelectionTick(1),4))", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (4, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (7, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (10, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("LET(src_,ISNUMBER(A1:O3),TYPE(INDEX(src_,SelectionTick(1),4)))", [(1, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (4, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (7, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (10, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)])]),
        ("LET(src_,ISNUMBER(A1:O3),SelectionTick(INDEX(src_,SelectionTick(1),4)))", [(2, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (8, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (14, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (20, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("LET(src_,ISNUMBER(A1:O3),SUM(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (7, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (10, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,ISNUMBER(A1:O3),TYPE(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (7, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (10, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)])]),
        ("INDEX(ISNUMBER(A1:O3),SelectionTick(1),13)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (7, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (13, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (22, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("TYPE(INDEX(ISNUMBER(A1:O3),SelectionTick(1),13))", [(1, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (7, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (13, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (22, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)])]),
        ("SelectionTick(INDEX(ISNUMBER(A1:O3),SelectionTick(1),13))", [(2, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (8, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (14, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (23, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("SUM(INDEX(ISNUMBER(A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("TYPE(INDEX(ISNUMBER(A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (7, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (13, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (22, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,ISNUMBER(A1:O3),INDEX(src_,SelectionTick(1),13))", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (7, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (13, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (22, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("LET(src_,ISNUMBER(A1:O3),TYPE(INDEX(src_,SelectionTick(1),13)))", [(1, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (7, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (13, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (22, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)])]),
        ("LET(src_,ISNUMBER(A1:O3),SelectionTick(INDEX(src_,SelectionTick(1),13)))", [(2, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (8, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (14, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (23, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("LET(src_,ISNUMBER(A1:O3),SUM(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,ISNUMBER(A1:O3),TYPE(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (7, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (13, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (22, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(N(A1:O3+0),SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(N(A1:O3+0),SelectionTick(1),4))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(N(A1:O3+0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("TYPE(INDEX(N(A1:O3+0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,N(A1:O3+0),INDEX(src_,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,N(A1:O3+0),TYPE(INDEX(src_,SelectionTick(1),4)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,N(A1:O3+0),SelectionTick(INDEX(src_,SelectionTick(1),4)))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,N(A1:O3+0),SUM(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("LET(src_,N(A1:O3+0),TYPE(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(N(A1:O3+0),SelectionTick(1),13))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(N(A1:O3+0),SelectionTick(1),13))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("SUM(INDEX(N(A1:O3+0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("TYPE(INDEX(N(A1:O3+0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,N(A1:O3+0),INDEX(src_,SelectionTick(1),13))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,N(A1:O3+0),TYPE(INDEX(src_,SelectionTick(1),13)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,N(A1:O3+0),SelectionTick(INDEX(src_,SelectionTick(1),13)))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,N(A1:O3+0),SUM(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,N(A1:O3+0),TYPE(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(IFERROR(A1:O3,0),SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(IFERROR(A1:O3,0),SelectionTick(1),4))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(IFERROR(A1:O3,0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("TYPE(INDEX(IFERROR(A1:O3,0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,IFERROR(A1:O3,0),INDEX(src_,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,IFERROR(A1:O3,0),TYPE(INDEX(src_,SelectionTick(1),4)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,IFERROR(A1:O3,0),SelectionTick(INDEX(src_,SelectionTick(1),4)))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,IFERROR(A1:O3,0),SUM(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("LET(src_,IFERROR(A1:O3,0),TYPE(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("INDEX(IFERROR(A1:O3,0),SelectionTick(1),13)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("TYPE(INDEX(IFERROR(A1:O3,0),SelectionTick(1),13))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(IFERROR(A1:O3,0),SelectionTick(1),13))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("SUM(INDEX(IFERROR(A1:O3,0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("TYPE(INDEX(IFERROR(A1:O3,0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,IFERROR(A1:O3,0),INDEX(src_,SelectionTick(1),13))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,IFERROR(A1:O3,0),TYPE(INDEX(src_,SelectionTick(1),13)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,IFERROR(A1:O3,0),SelectionTick(INDEX(src_,SelectionTick(1),13)))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,IFERROR(A1:O3,0),SUM(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,IFERROR(A1:O3,0),TYPE(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("INDEX(ROUND(A1:O3,0),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("TYPE(INDEX(ROUND(A1:O3,0),SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(ROUND(A1:O3,0),SelectionTick(1),4))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUM(INDEX(ROUND(A1:O3,0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("TYPE(INDEX(ROUND(A1:O3,0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,ROUND(A1:O3,0),INDEX(src_,SelectionTick(1),4))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,ROUND(A1:O3,0),TYPE(INDEX(src_,SelectionTick(1),4)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,ROUND(A1:O3,0),SelectionTick(INDEX(src_,SelectionTick(1),4)))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (8, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,ROUND(A1:O3,0),SUM(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(6.0), Value::Number(6.0), Value::Number(6.0)]), (5, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (9, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)]), (13, [Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)])]),
        ("LET(src_,ROUND(A1:O3,0),TYPE(INDEX(src_,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("INDEX(ROUND(A1:O3,0),SelectionTick(1),13)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("TYPE(INDEX(ROUND(A1:O3,0),SelectionTick(1),13))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(INDEX(ROUND(A1:O3,0),SelectionTick(1),13))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("SUM(INDEX(ROUND(A1:O3,0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("TYPE(INDEX(ROUND(A1:O3,0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,ROUND(A1:O3,0),INDEX(src_,SelectionTick(1),13))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,ROUND(A1:O3,0),TYPE(INDEX(src_,SelectionTick(1),13)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (6, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(src_,ROUND(A1:O3,0),SelectionTick(INDEX(src_,SelectionTick(1),13)))", [(2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,ROUND(A1:O3,0),SUM(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(src_,ROUND(A1:O3,0),TYPE(INDEX(src_,{1;2;3},SelectionTick(13))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("INDEX(TYPE(A1:O3+0),SelectionTick(1),1)", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("LET(src_,TYPE(A1:O3+0),INDEX(src_,SelectionTick(1),1))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(A1:O3+0))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(M1))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(TYPE(M1+0))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(TYPE(A1:O3))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(IF(M1,1,0)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(TYPE(IF(M1,A1:O3,0)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(TYPE(IF(A1:O3>0,A1:O3,0)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(CHOOSE(M1+1,A1:O3,0)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(SUM(M1:M3)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(TYPE(IFERROR(M1:M3,0)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(SelectionTick(M1)))", [(2, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (3, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (7, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SUMPRODUCT(TYPE(INDEX(A1:O3,{1;2;3},SelectionTick(4))))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(IFERROR(A1:O3+0,0)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(INDEX(M1:M3,1,1)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(TYPE(INDEX(M1:M3,{1;2;3},1)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(INDEX(M1:M3+0,{1;2;3},1)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(A1:O3,{1},SelectionTick(4)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(A1:O3+0,{1},SelectionTick(4)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(ABS(M1)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(TYPE(ABS(A1:O3)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(M1:O3+0))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(L1:O3+0))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(M1:N3))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(M1:M3))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(L1:M3))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(M1:O3+0,{1;2;3},SelectionTick(1)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(L1:O3+0,{1;2;3},SelectionTick(2)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(M1:O3,{1;2;3},SelectionTick(1)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(M1:O3+0,{1;2;3},SelectionTick(2)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(IFERROR(L1:M3,0)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(IFERROR(M1:O3,0)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(IF(M1:O3>0,M1:O3,0)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(A1:O3,IF(M1,{1;2;3},{1;2;3}),SelectionTick(13)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(IF(M1,A1:O3,A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(IFERROR(M1:O3,0),{1;2;3},SelectionTick(1)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(ABS(M1:O3),{1;2;3},SelectionTick(1)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("INDEX(ISNUMBER(A1:O3+0),SelectionTick(1),4)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (4, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (7, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (10, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("SUM(INDEX(ISNUMBER(A1:O3+0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (7, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (10, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(ISNUMBER(A1:O3+0),SelectionTick(1),13)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (7, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (13, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (22, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("SUM(INDEX(ISNUMBER(A1:O3+0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISERROR(A1:O3),SelectionTick(1),4)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (10, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISERROR(A1:O3),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (10, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISERROR(A1:O3),SelectionTick(1),13)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (13, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (22, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISERROR(A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISERROR(A1:O3+0),SelectionTick(1),4)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (10, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISERROR(A1:O3+0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (10, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISERROR(A1:O3+0),SelectionTick(1),13)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (13, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (22, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISERROR(A1:O3+0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISTEXT(A1:O3),SelectionTick(1),4)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (10, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISTEXT(A1:O3),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (10, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISTEXT(A1:O3),SelectionTick(1),13)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (13, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (22, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISTEXT(A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISTEXT(A1:O3+0),SelectionTick(1),4)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (10, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISTEXT(A1:O3+0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (10, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISTEXT(A1:O3+0),SelectionTick(1),13)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (13, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (22, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISTEXT(A1:O3+0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISBLANK(A1:O3),SelectionTick(1),4)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (10, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISBLANK(A1:O3),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (10, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISBLANK(A1:O3),SelectionTick(1),13)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (13, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (22, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISBLANK(A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISBLANK(A1:O3+0),SelectionTick(1),4)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (10, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISBLANK(A1:O3+0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (10, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISBLANK(A1:O3+0),SelectionTick(1),13)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (13, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (22, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISBLANK(A1:O3+0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISLOGICAL(A1:O3),SelectionTick(1),4)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (10, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISLOGICAL(A1:O3),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (10, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISLOGICAL(A1:O3),SelectionTick(1),13)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (13, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (22, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISLOGICAL(A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISLOGICAL(A1:O3+0),SelectionTick(1),4)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (10, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISLOGICAL(A1:O3+0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (10, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISLOGICAL(A1:O3+0),SelectionTick(1),13)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (13, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (22, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISLOGICAL(A1:O3+0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISNA(A1:O3),SelectionTick(1),4)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (10, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISNA(A1:O3),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (10, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISNA(A1:O3),SelectionTick(1),13)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (13, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (22, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISNA(A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISNA(A1:O3+0),SelectionTick(1),4)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (10, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISNA(A1:O3+0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (10, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISNA(A1:O3+0),SelectionTick(1),13)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (13, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (22, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISNA(A1:O3+0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISERR(A1:O3),SelectionTick(1),4)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (10, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISERR(A1:O3),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (10, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISERR(A1:O3),SelectionTick(1),13)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (13, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (22, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISERR(A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISERR(A1:O3+0),SelectionTick(1),4)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (10, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISERR(A1:O3+0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (10, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISERR(A1:O3+0),SelectionTick(1),13)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (13, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (22, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(ISERR(A1:O3+0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(NOT(A1:O3),SelectionTick(1),4)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (2, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (3, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(NOT(A1:O3),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (5, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (9, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(NOT(A1:O3),SelectionTick(1),13)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (2, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (3, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (6, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("SUM(INDEX(NOT(A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(NOT(A1:O3+0),SelectionTick(1),4)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (2, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (3, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUM(INDEX(NOT(A1:O3+0),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (5, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (9, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(NOT(A1:O3+0),SelectionTick(1),13)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (2, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (3, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (6, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("SUM(INDEX(NOT(A1:O3+0),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (8, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (15, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (25, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("SelectionTick(TYPE(M1:O1))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(M1+{1,2,3}))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(M1:M3+{1,2}))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(M1:M3+0,{1},SelectionTick(1)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("TYPE(INDEX(M1:M3+0,SelectionTick(1),1))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(TYPE(SUM(A1:O3)))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("SelectionTick(TYPE(SelectionTick(A1:O3)))", [(2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (7, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(TYPE(SelectionTick(M1:M3)))", [(2, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (3, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (7, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(SUM(ISNUMBER(A1:O3)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("SelectionTick(SUMPRODUCT(--ISNUMBER(A1:O3)))", [(1, [Value::Number(9.0), Value::Number(9.0), Value::Number(9.0)]), (1, [Value::Number(9.0), Value::Number(9.0), Value::Number(9.0)]), (1, [Value::Number(9.0), Value::Number(9.0), Value::Number(9.0)]), (1, [Value::Number(9.0), Value::Number(9.0), Value::Number(9.0)])]),
        ("SelectionTick(TYPE(ISNUMBER(A1:O3)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (7, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (10, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(ISNUMBER(A1:O3))", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (13, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (22, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("SUMPRODUCT(INDEX(ISNUMBER(A1:O3),{1;2;3},SelectionTick(4)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (10, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("SUMPRODUCT(INDEX(ISNUMBER(A1:O3),{1;2;3},SelectionTick(13)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (7, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (13, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (22, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(IF(ISNUMBER(A1:O3),1,0),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (7, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (10, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(--ISNUMBER(A1:O3),SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (7, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (10, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("LET(x,ISNUMBER(A1:O3),SelectionTick(SUM(x)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("LET(x,ISNUMBER(A1:O3),SelectionTick(TYPE(x)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (7, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (10, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("SelectionTick(SUM(ISNUMBER(M1:M3)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("SelectionTick(TYPE(ISNUMBER(M1:M3)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("INDEX(ISNUMBER({1}),SelectionTick(1),1)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (4, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (7, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (10, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("SelectionTick(ISNUMBER({1}))", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (4, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (7, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (10, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("LET(t,ISNUMBER({1}),SelectionTick(1))", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (7, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (10, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(ISNUMBER(D1),SelectionTick(1),1)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (2, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (3, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (4, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("INDEX(ISNUMBER(D1+{0}),SelectionTick(1),1)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (4, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (7, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (10, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("INDEX(ISERROR({1}),SelectionTick(1),1)", [(1, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (4, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (7, [Value::Logical(false), Value::Logical(false), Value::Logical(false)]), (10, [Value::Logical(false), Value::Logical(false), Value::Logical(false)])]),
        ("INDEX(ISNUMBER({1;2;3}),SelectionTick(1),1)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (4, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (7, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (10, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("INDEX(ISNUMBER(D1:D3),SelectionTick(1),1)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (4, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (7, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (10, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("SUM(INDEX(ISNUMBER(D1:D3),{1;2;3},SelectionTick(1)))", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (4, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (7, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (10, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("TYPE(INDEX(ISNUMBER(D1:D3),{1;2;3},SelectionTick(1)))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (7, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)]), (10, [Value::Number(4.0), Value::Number(4.0), Value::Number(4.0)])]),
        ("SelectionTick(TYPE(ISNUMBER({1})))", [(1, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (4, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (7, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)]), (10, [Value::Number(64.0), Value::Number(64.0), Value::Number(64.0)])]),
        ("INDEX(ISNUMBER(M1+{0}),SelectionTick(1),1)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("INDEX(M1+{0},SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(M1+{0,0},SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(M1+1,SelectionTick(1),1)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(-M1,SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(M1*{1;1;1},SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(M1+{0,0},SelectionTick(1),2)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(IF(M1+{0,1},1,0),SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ISNUMBER(M1+{0,0}),SelectionTick(1),1)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("INDEX(ISNUMBER(M1+{0;0;0}),SelectionTick(1),1)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("INDEX(ISNUMBER(M1),SelectionTick(1),1)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
        ("INDEX(ABS(M1+{0,0}),SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(M1,SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(IF(1,M1),SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (2, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (3, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (6, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(ABS(M1),SelectionTick(1),1)", [(1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)]), (1, [Value::Number(0.0), Value::Number(0.0), Value::Number(0.0)])]),
        ("INDEX(M1+A1:O3,SelectionTick(1),4)", [(1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)]), (1, [Value::Number(1.0), Value::Number(1.0), Value::Number(1.0)])]),
        ("INDEX(ISNUMBER(M1+A1:O3),SelectionTick(1),4)", [(1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)]), (1, [Value::Logical(true), Value::Logical(true), Value::Logical(true)])]),
    ];
    let mut mismatches = Vec::new();
    for (formula, stages) in CASES {
        for version in [1u8, 2] {
            let count = Rc::new(RefCell::new(0u32));
            let counted = count.clone();
            set_user_function_hook(Some(Box::new(move |name, args, _| {
                if !name.eq_ignore_ascii_case("SelectionTick") { return None; }
                *counted.borrow_mut() += 1;
                Some(match args.first() {
                    Some(UserArg::Value(value)) => Arg::Value(value.clone()),
                    Some(UserArg::Block(block)) => Arg::Range(RangeData {
                        origin: BlockOrigin::Array, ..block.clone()
                    }),
                    Some(UserArg::Cells { range, values, .. }) if range.is_single() =>
                        Arg::Value(values.first().cloned().unwrap_or(Value::Blank)),
                    Some(UserArg::Cells { range, values, .. }) => Arg::Range(RangeData {
                        origin: BlockOrigin::Array,
                        width: range.width() as usize, height: range.height() as usize,
                        cells: values.clone(),
                    }),
                    _ => Arg::Value(Value::Blank),
                })
            })));
            oxicells_calc::engine::set_user_function_variant_arguments(&[("SelectionTick".to_string(), vec![0])]);
            oxicells_calc::engine::set_user_function_value_read_arguments(&[("SelectionTick".to_string(), vec![0])]);
            let mut held = Workbook::new();
            held.add_sheet("S");
            held.set_compatibility_version(version);
            // Match the archived VBA driver, including isolated UTF-16 surrogates.
            for (address, value) in [
                ("A1", Value::Utf16(vec![97, 55357, 56832, 98])),
                ("B1", Value::Number(-1.0)),
                ("C1", Value::Utf16(vec![97])),
                ("D1", Value::Number(1.0)),
                ("E1", Value::Utf16(vec![55357])),
                ("A2", Value::Utf16(vec![55357])),
                ("B2", Value::Number(0.0)),
                ("C2", Value::Utf16(vec![56832])),
                ("D2", Value::Number(2.0)),
                ("E2", Value::Utf16(vec![56832])),
                ("A3", Value::Utf16(vec![56832])),
                ("B3", Value::Number(1.0)),
                ("C3", Value::Utf16(vec![126, 63])),
                ("D3", Value::Number(3.0)),
            ] {
                held.set_value("S", address, value).unwrap();
            }
            let wanted = vec![("S".to_string(), (12, 0)), ("S".to_string(), (12, 1)), ("S".to_string(), (12, 2))];
            for (row, cell) in ["M1", "M2", "M3"].iter().enumerate() {
                held.set_array_member("S", cell, formula, (0, row as u32)).unwrap();
            }
            let mut reverse = wanted.clone(); reverse.reverse();
            held.set_calculation_order(&reverse);
            for (stage, (expected_calls, expected_values)) in stages.iter().enumerate() {
                held.set_chain_probes(stage != 0);
                oxicells_calc::engine::set_user_function_wait_for_dependencies(true);
                oxicells_calc::engine::set_user_function_full_rebuild(stage == 3);
                match stage {
                    0 => { held.recalculate_array_entry(&wanted); }
                    1 | 2 => { held.recalculate_these(&wanted); }
                    _ => { held.recalculate(); }
                }
                let values: Vec<_> = ["M1", "M2", "M3"].iter().map(|cell| held.value("S", cell)).collect();
                let calls = *count.borrow();
                if values.as_slice() != expected_values || calls != *expected_calls {
                    mismatches.push(format!("{formula}, version {version}, stage {stage}: calls {calls} != {expected_calls}, {values:?} != {expected_values:?}"));
                }
            }
        }
    }
    assert!(mismatches.is_empty(), "{} Native dynamic INDEX stage mismatches:\n{}", mismatches.len(), mismatches.join("\n"));
}
