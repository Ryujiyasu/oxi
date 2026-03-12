use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Sheet {
    pub name: String,
    pub rows: Vec<Row>,
    pub col_count: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Row {
    pub index: u32,
    pub cells: Vec<Cell>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Cell {
    pub col: u32,
    pub value: CellValue,
    pub style: CellStyle,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum CellValue {
    Empty,
    String(String),
    Number(f64),
    Boolean(bool),
    Error(String),
}

impl CellValue {
    pub fn display(&self) -> String {
        match self {
            CellValue::Empty => String::new(),
            CellValue::String(s) => s.clone(),
            CellValue::Number(n) => {
                if *n == (*n as i64) as f64 {
                    format!("{}", *n as i64)
                } else {
                    format!("{}", n)
                }
            }
            CellValue::Boolean(b) => {
                if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            CellValue::Error(e) => e.clone(),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct CellStyle {
    pub bold: bool,
    pub italic: bool,
    pub font_size: Option<f32>,
    pub font_color: Option<String>,
    pub bg_color: Option<String>,
    pub number_format: Option<String>,
}
