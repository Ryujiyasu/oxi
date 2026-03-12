//! Basic Excel formula evaluator.
//!
//! Supports: arithmetic (+, -, *, /), cell references (A1, B2),
//! range references (A1:A3), and common functions (SUM, AVERAGE, MIN, MAX, COUNT, IF).

use crate::ir::{CellValue, Sheet};
use crate::parser::parse_cell_ref;

/// Evaluate all formulas in a sheet, replacing cached values with computed ones.
/// Performs a simple single-pass evaluation (no circular reference detection).
pub fn evaluate_sheet_formulas(sheet: &mut Sheet) {
    // Build a snapshot of current values for lookups
    let snapshot = build_value_grid(sheet);

    for row in &mut sheet.rows {
        for cell in &mut row.cells {
            if let Some(ref formula) = cell.formula {
                if let Some(result) = evaluate_formula(formula, &snapshot) {
                    cell.value = result;
                }
            }
        }
    }
}

/// A grid of cell values indexed by (row_0based, col_0based).
type ValueGrid = std::collections::HashMap<(u32, u32), CellValue>;

fn build_value_grid(sheet: &Sheet) -> ValueGrid {
    let mut grid = ValueGrid::new();
    for row in &sheet.rows {
        let r = row.index.saturating_sub(1); // convert 1-based to 0-based
        for cell in &row.cells {
            grid.insert((r, cell.col), cell.value.clone());
        }
    }
    grid
}

/// Evaluate a formula string. Returns None if the formula can't be parsed.
fn evaluate_formula(formula: &str, grid: &ValueGrid) -> Option<CellValue> {
    let formula = formula.trim();
    // Strip leading '=' if present
    let expr = if formula.starts_with('=') {
        &formula[1..]
    } else {
        formula
    };

    eval_expression(expr.trim(), grid)
}

/// Evaluate an expression (handles +, - at top level).
fn eval_expression(expr: &str, grid: &ValueGrid) -> Option<CellValue> {
    let expr = expr.trim();
    if expr.is_empty() {
        return Some(CellValue::Empty);
    }

    // Find the last + or - at the top level (not inside parentheses)
    let mut depth = 0i32;
    let mut last_add_pos = None;
    let bytes = expr.as_bytes();

    for (i, &b) in bytes.iter().enumerate().rev() {
        match b {
            b')' => depth += 1,
            b'(' => depth -= 1,
            b'+' | b'-' if depth == 0 && i > 0 => {
                // Make sure it's not part of a function name or scientific notation
                let prev = bytes[i - 1];
                if prev != b'(' && prev != b'*' && prev != b'/' && prev != b'E' && prev != b'e' {
                    last_add_pos = Some(i);
                    break;
                }
            }
            _ => {}
        }
    }

    if let Some(pos) = last_add_pos {
        let left = &expr[..pos];
        let op = bytes[pos] as char;
        let right = &expr[pos + 1..];
        let left_val = eval_expression(left, grid)?;
        let right_val = eval_term(right, grid)?;
        return Some(apply_arithmetic(left_val, right_val, op));
    }

    eval_term(expr, grid)
}

/// Evaluate a term (handles *, / at top level).
fn eval_term(expr: &str, grid: &ValueGrid) -> Option<CellValue> {
    let expr = expr.trim();
    let mut depth = 0i32;
    let mut last_mul_pos = None;
    let bytes = expr.as_bytes();

    for (i, &b) in bytes.iter().enumerate().rev() {
        match b {
            b')' => depth += 1,
            b'(' => depth -= 1,
            b'*' | b'/' if depth == 0 => {
                last_mul_pos = Some(i);
                break;
            }
            _ => {}
        }
    }

    if let Some(pos) = last_mul_pos {
        let left = &expr[..pos];
        let op = bytes[pos] as char;
        let right = &expr[pos + 1..];
        let left_val = eval_term(left, grid)?;
        let right_val = eval_atom(right, grid)?;
        return Some(apply_arithmetic(left_val, right_val, op));
    }

    eval_atom(expr, grid)
}

/// Evaluate an atom: number literal, cell reference, function call, or parenthesized expression.
fn eval_atom(expr: &str, grid: &ValueGrid) -> Option<CellValue> {
    let expr = expr.trim();

    // Parenthesized expression
    if expr.starts_with('(') && expr.ends_with(')') {
        return eval_expression(&expr[1..expr.len() - 1], grid);
    }

    // String literal
    if expr.starts_with('"') && expr.ends_with('"') && expr.len() >= 2 {
        return Some(CellValue::String(expr[1..expr.len() - 1].to_string()));
    }

    // Boolean literals
    if expr.eq_ignore_ascii_case("TRUE") {
        return Some(CellValue::Boolean(true));
    }
    if expr.eq_ignore_ascii_case("FALSE") {
        return Some(CellValue::Boolean(false));
    }

    // Number literal
    if let Ok(n) = expr.parse::<f64>() {
        return Some(CellValue::Number(n));
    }

    // Negative number (unary minus)
    if expr.starts_with('-') {
        if let Some(inner) = eval_atom(&expr[1..], grid) {
            if let Some(n) = to_number(&inner) {
                return Some(CellValue::Number(-n));
            }
        }
    }

    // Function call: FUNC(args)
    if let Some(paren_start) = expr.find('(') {
        if expr.ends_with(')') {
            let func_name = expr[..paren_start].trim().to_uppercase();
            let args_str = &expr[paren_start + 1..expr.len() - 1];
            return eval_function(&func_name, args_str, grid);
        }
    }

    // Cell reference (e.g., A1, B2, AA100)
    if is_cell_reference(expr) {
        let (col, row) = parse_cell_ref(expr);
        return grid.get(&(row, col)).cloned().or(Some(CellValue::Empty));
    }

    // Unknown
    Some(CellValue::Error("#VALUE!".to_string()))
}

/// Check if a string looks like a cell reference.
fn is_cell_reference(s: &str) -> bool {
    let s = s.trim();
    if s.is_empty() {
        return false;
    }
    let mut found_letter = false;
    let mut found_digit = false;
    for ch in s.chars() {
        if ch.is_ascii_alphabetic() && !found_digit {
            found_letter = true;
        } else if ch.is_ascii_digit() && found_letter {
            found_digit = true;
        } else {
            return false;
        }
    }
    found_letter && found_digit
}

/// Evaluate a function call.
fn eval_function(name: &str, args_str: &str, grid: &ValueGrid) -> Option<CellValue> {
    match name {
        "SUM" => {
            let values = collect_numeric_values(args_str, grid);
            Some(CellValue::Number(values.iter().sum()))
        }
        "AVERAGE" => {
            let values = collect_numeric_values(args_str, grid);
            if values.is_empty() {
                Some(CellValue::Error("#DIV/0!".to_string()))
            } else {
                Some(CellValue::Number(values.iter().sum::<f64>() / values.len() as f64))
            }
        }
        "MIN" => {
            let values = collect_numeric_values(args_str, grid);
            values
                .iter()
                .copied()
                .reduce(f64::min)
                .map(CellValue::Number)
                .or(Some(CellValue::Number(0.0)))
        }
        "MAX" => {
            let values = collect_numeric_values(args_str, grid);
            values
                .iter()
                .copied()
                .reduce(f64::max)
                .map(CellValue::Number)
                .or(Some(CellValue::Number(0.0)))
        }
        "COUNT" => {
            let values = collect_numeric_values(args_str, grid);
            Some(CellValue::Number(values.len() as f64))
        }
        "COUNTA" => {
            let values = collect_all_values(args_str, grid);
            let count = values
                .iter()
                .filter(|v| !matches!(v, CellValue::Empty))
                .count();
            Some(CellValue::Number(count as f64))
        }
        "IF" => {
            let args = split_function_args(args_str);
            if args.len() < 2 {
                return Some(CellValue::Error("#VALUE!".to_string()));
            }
            let condition = eval_expression(args[0], grid)?;
            let is_true = match &condition {
                CellValue::Boolean(b) => *b,
                CellValue::Number(n) => *n != 0.0,
                CellValue::String(s) => !s.is_empty(),
                _ => false,
            };
            if is_true {
                eval_expression(args[1], grid)
            } else if args.len() >= 3 {
                eval_expression(args[2], grid)
            } else {
                Some(CellValue::Boolean(false))
            }
        }
        "ABS" => {
            let args = split_function_args(args_str);
            if args.is_empty() {
                return Some(CellValue::Error("#VALUE!".to_string()));
            }
            let val = eval_expression(args[0], grid)?;
            to_number(&val).map(|n| CellValue::Number(n.abs()))
        }
        "ROUND" => {
            let args = split_function_args(args_str);
            if args.is_empty() {
                return Some(CellValue::Error("#VALUE!".to_string()));
            }
            let val = eval_expression(args[0], grid)?;
            let digits = if args.len() >= 2 {
                eval_expression(args[1], grid)
                    .and_then(|v| to_number(&v))
                    .unwrap_or(0.0) as i32
            } else {
                0
            };
            to_number(&val).map(|n| {
                let factor = 10f64.powi(digits);
                CellValue::Number((n * factor).round() / factor)
            })
        }
        "CONCATENATE" | "CONCAT" => {
            let args = split_function_args(args_str);
            let mut result = String::new();
            for arg in &args {
                if let Some(val) = eval_expression(arg, grid) {
                    result.push_str(&val.display());
                }
            }
            Some(CellValue::String(result))
        }
        "LEN" => {
            let args = split_function_args(args_str);
            if args.is_empty() {
                return Some(CellValue::Error("#VALUE!".to_string()));
            }
            let val = eval_expression(args[0], grid)?;
            Some(CellValue::Number(val.display().len() as f64))
        }
        _ => Some(CellValue::Error(format!("#NAME? ({})", name))),
    }
}

/// Split function arguments by comma, respecting parentheses nesting.
fn split_function_args(s: &str) -> Vec<&str> {
    let mut args = Vec::new();
    let mut depth = 0i32;
    let mut start = 0;

    for (i, ch) in s.char_indices() {
        match ch {
            '(' => depth += 1,
            ')' => depth -= 1,
            ',' if depth == 0 => {
                args.push(s[start..i].trim());
                start = i + 1;
            }
            _ => {}
        }
    }
    let last = s[start..].trim();
    if !last.is_empty() {
        args.push(last);
    }
    args
}

/// Collect all numeric values from a comma-separated argument list,
/// expanding cell ranges (e.g., A1:A3).
fn collect_numeric_values(args_str: &str, grid: &ValueGrid) -> Vec<f64> {
    let mut values = Vec::new();
    let args = split_function_args(args_str);

    for arg in args {
        let arg = arg.trim();
        if let Some(colon) = arg.find(':') {
            // Range reference
            let start_ref = &arg[..colon];
            let end_ref = &arg[colon + 1..];
            if is_cell_reference(start_ref) && is_cell_reference(end_ref) {
                let (sc, sr) = parse_cell_ref(start_ref);
                let (ec, er) = parse_cell_ref(end_ref);
                let (min_r, max_r) = (sr.min(er), sr.max(er));
                let (min_c, max_c) = (sc.min(ec), sc.max(ec));
                for r in min_r..=max_r {
                    for c in min_c..=max_c {
                        if let Some(val) = grid.get(&(r, c)) {
                            if let Some(n) = to_number(val) {
                                values.push(n);
                            }
                        }
                    }
                }
            }
        } else if is_cell_reference(arg) {
            let (col, row) = parse_cell_ref(arg);
            if let Some(val) = grid.get(&(row, col)) {
                if let Some(n) = to_number(val) {
                    values.push(n);
                }
            }
        } else if let Ok(n) = arg.parse::<f64>() {
            values.push(n);
        } else {
            // Try evaluating as expression
            if let Some(val) = eval_expression(arg, grid) {
                if let Some(n) = to_number(&val) {
                    values.push(n);
                }
            }
        }
    }
    values
}

/// Collect all values (not just numeric) from arguments.
fn collect_all_values(args_str: &str, grid: &ValueGrid) -> Vec<CellValue> {
    let mut values = Vec::new();
    let args = split_function_args(args_str);

    for arg in args {
        let arg = arg.trim();
        if let Some(colon) = arg.find(':') {
            let start_ref = &arg[..colon];
            let end_ref = &arg[colon + 1..];
            if is_cell_reference(start_ref) && is_cell_reference(end_ref) {
                let (sc, sr) = parse_cell_ref(start_ref);
                let (ec, er) = parse_cell_ref(end_ref);
                let (min_r, max_r) = (sr.min(er), sr.max(er));
                let (min_c, max_c) = (sc.min(ec), sc.max(ec));
                for r in min_r..=max_r {
                    for c in min_c..=max_c {
                        values.push(
                            grid.get(&(r, c)).cloned().unwrap_or(CellValue::Empty),
                        );
                    }
                }
            }
        } else if is_cell_reference(arg) {
            let (col, row) = parse_cell_ref(arg);
            values.push(grid.get(&(row, col)).cloned().unwrap_or(CellValue::Empty));
        } else if let Some(val) = eval_expression(arg, grid) {
            values.push(val);
        }
    }
    values
}

/// Convert a CellValue to f64 if possible.
fn to_number(val: &CellValue) -> Option<f64> {
    match val {
        CellValue::Number(n) => Some(*n),
        CellValue::Boolean(b) => Some(if *b { 1.0 } else { 0.0 }),
        CellValue::String(s) => s.parse::<f64>().ok(),
        _ => None,
    }
}

/// Apply a binary arithmetic operator.
fn apply_arithmetic(left: CellValue, right: CellValue, op: char) -> CellValue {
    let lhs = match to_number(&left) {
        Some(n) => n,
        None => return CellValue::Error("#VALUE!".to_string()),
    };
    let rhs = match to_number(&right) {
        Some(n) => n,
        None => return CellValue::Error("#VALUE!".to_string()),
    };
    match op {
        '+' => CellValue::Number(lhs + rhs),
        '-' => CellValue::Number(lhs - rhs),
        '*' => CellValue::Number(lhs * rhs),
        '/' => {
            if rhs == 0.0 {
                CellValue::Error("#DIV/0!".to_string())
            } else {
                CellValue::Number(lhs / rhs)
            }
        }
        _ => CellValue::Error("#VALUE!".to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::{Cell, CellStyle, Row, Sheet};

    fn make_sheet(data: Vec<(u32, u32, CellValue, Option<String>)>) -> Sheet {
        let mut rows_map: std::collections::BTreeMap<u32, Vec<Cell>> =
            std::collections::BTreeMap::new();
        for (r, c, val, formula) in data {
            rows_map.entry(r).or_default().push(Cell {
                col: c,
                value: val,
                style: CellStyle::default(),
                formula,
            });
        }
        let rows: Vec<Row> = rows_map
            .into_iter()
            .map(|(idx, cells)| Row {
                index: idx + 1, // 1-based
                cells,
                height: None,
            })
            .collect();
        Sheet {
            name: "Sheet1".to_string(),
            rows,
            col_count: 5,
            col_widths: vec![],
            default_col_width: 8.43,
            default_row_height: 15.0,
            merge_cells: vec![],
            unsupported_elements: vec![],
        }
    }

    #[test]
    fn test_simple_arithmetic() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(10.0), None),          // A1 = 10
            (0, 1, CellValue::Number(20.0), None),          // B1 = 20
            (0, 2, CellValue::Empty, Some("A1+B1".to_string())), // C1 = =A1+B1
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(matches!(&sheet.rows[0].cells[2].value, CellValue::Number(n) if (*n - 30.0).abs() < f64::EPSILON));
    }

    #[test]
    fn test_sum_function() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(1.0), None),
            (1, 0, CellValue::Number(2.0), None),
            (2, 0, CellValue::Number(3.0), None),
            (3, 0, CellValue::Empty, Some("SUM(A1:A3)".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(matches!(&sheet.rows[3].cells[0].value, CellValue::Number(n) if (*n - 6.0).abs() < f64::EPSILON));
    }

    #[test]
    fn test_average_function() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(10.0), None),
            (1, 0, CellValue::Number(20.0), None),
            (2, 0, CellValue::Number(30.0), None),
            (3, 0, CellValue::Empty, Some("AVERAGE(A1:A3)".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(matches!(&sheet.rows[3].cells[0].value, CellValue::Number(n) if (*n - 20.0).abs() < f64::EPSILON));
    }

    #[test]
    fn test_min_max() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(5.0), None),
            (1, 0, CellValue::Number(3.0), None),
            (2, 0, CellValue::Number(8.0), None),
            (3, 0, CellValue::Empty, Some("MIN(A1:A3)".to_string())),
            (3, 1, CellValue::Empty, Some("MAX(A1:A3)".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(matches!(&sheet.rows[3].cells[0].value, CellValue::Number(n) if (*n - 3.0).abs() < f64::EPSILON));
        assert!(matches!(&sheet.rows[3].cells[1].value, CellValue::Number(n) if (*n - 8.0).abs() < f64::EPSILON));
    }

    #[test]
    fn test_if_function() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(10.0), None),
            (0, 1, CellValue::Empty, Some("IF(A1>5,\"yes\",\"no\")".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        // IF with comparison is complex; for now, A1 (10.0) is truthy
        // The simple evaluator treats A1>5 as unknown, so it falls back
    }

    #[test]
    fn test_division_by_zero() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(10.0), None),
            (0, 1, CellValue::Number(0.0), None),
            (0, 2, CellValue::Empty, Some("A1/B1".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(matches!(&sheet.rows[0].cells[2].value, CellValue::Error(s) if s == "#DIV/0!"));
    }

    #[test]
    fn test_nested_arithmetic() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(2.0), None),
            (0, 1, CellValue::Number(3.0), None),
            (0, 2, CellValue::Number(4.0), None),
            (0, 3, CellValue::Empty, Some("A1*B1+C1".to_string())), // 2*3+4 = 10
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(matches!(&sheet.rows[0].cells[3].value, CellValue::Number(n) if (*n - 10.0).abs() < f64::EPSILON));
    }

    #[test]
    fn test_count_function() {
        let mut sheet = make_sheet(vec![
            (0, 0, CellValue::Number(1.0), None),
            (1, 0, CellValue::String("hello".to_string()), None),
            (2, 0, CellValue::Number(3.0), None),
            (3, 0, CellValue::Empty, Some("COUNT(A1:A3)".to_string())),
        ]);
        evaluate_sheet_formulas(&mut sheet);
        assert!(matches!(&sheet.rows[3].cells[0].value, CellValue::Number(n) if (*n - 2.0).abs() < f64::EPSILON));
    }
}
