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
            index: row, height: None, custom_height: false, style_font: None,
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
