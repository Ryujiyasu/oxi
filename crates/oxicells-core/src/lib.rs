pub mod ir;
pub mod parser;

pub use parser::parse_xlsx;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_basic_xlsx() {
        let data = include_bytes!("../../../tests/fixtures/basic_test.xlsx");
        let wb = parse_xlsx(data).expect("parse failed");
        assert_eq!(wb.sheets.len(), 1);
        assert_eq!(wb.sheets[0].name, "Sales");
        assert_eq!(wb.sheets[0].rows.len(), 4); // rows 1,2,3,5
        assert_eq!(wb.sheets[0].col_count, 5);

        // Header row
        let row1 = &wb.sheets[0].rows[0];
        assert!(matches!(&row1.cells[0].value, ir::CellValue::String(s) if s == "Product"));
        assert!(matches!(&row1.cells[1].value, ir::CellValue::String(s) if s == "Q1"));

        // Data row
        let row2 = &wb.sheets[0].rows[1];
        assert!(matches!(&row2.cells[0].value, ir::CellValue::String(s) if s == "Widget A"));
        assert!(matches!(&row2.cells[1].value, ir::CellValue::Number(n) if (*n - 1200.0).abs() < f64::EPSILON));
        assert!(matches!(&row2.cells[2].value, ir::CellValue::Number(n) if (*n - 1500.5).abs() < f64::EPSILON));

        // Row with special types (boolean, error, inline string)
        let row5 = &wb.sheets[0].rows[3];
        assert!(matches!(&row5.cells[0].value, ir::CellValue::Boolean(true)));
        assert!(matches!(&row5.cells[1].value, ir::CellValue::Error(s) if s == "#N/A"));
        assert!(matches!(&row5.cells[2].value, ir::CellValue::String(s) if s == "inline text"));
    }

    #[test]
    fn test_parse_multi_sheet_xlsx() {
        let data = include_bytes!("../../../tests/fixtures/multi_sheet.xlsx");
        let wb = parse_xlsx(data).expect("parse failed");
        assert_eq!(wb.sheets.len(), 2);
        assert_eq!(wb.sheets[0].name, "Data");
        assert_eq!(wb.sheets[1].name, "Summary");

        // Sheet 1
        assert_eq!(wb.sheets[0].rows.len(), 3);
        let alice_row = &wb.sheets[0].rows[1];
        assert!(matches!(&alice_row.cells[0].value, ir::CellValue::String(s) if s == "Alice"));
        assert!(matches!(&alice_row.cells[1].value, ir::CellValue::Number(n) if (*n - 95.0).abs() < f64::EPSILON));

        // Sheet 2
        assert_eq!(wb.sheets[1].rows.len(), 1);
        assert!(matches!(&wb.sheets[1].rows[0].cells[0].value, ir::CellValue::String(s) if s == "Total"));
        assert!(matches!(&wb.sheets[1].rows[0].cells[1].value, ir::CellValue::Number(n) if (*n - 182.0).abs() < f64::EPSILON));
    }
}
