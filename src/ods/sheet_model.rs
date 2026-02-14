use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "data", rename_all = "lowercase")]
pub enum CellValue {
    String(String),
    Number(f64),
    Boolean(bool),
    Empty,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    pub value: CellValue,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Sheet {
    pub name: String,
    pub rows: Vec<Vec<Cell>>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Workbook {
    pub sheets: Vec<Sheet>,
}

impl Cell {
    pub fn empty() -> Self {
        Self {
            value: CellValue::Empty,
        }
    }
}

impl Sheet {
    pub fn new(name: String) -> Self {
        Self {
            name,
            rows: Vec::new(),
        }
    }

    pub fn ensure_cell_mut(&mut self, row: usize, col: usize) -> &mut Cell {
        // Grow rows first.
        while self.rows.len() <= row {
            self.rows.push(Vec::new());
        }
        // Then grow every row to keep a rectangular matrix shape.
        for row_cells in &mut self.rows {
            while row_cells.len() <= col {
                row_cells.push(Cell::empty());
            }
        }
        &mut self.rows[row][col]
    }

    pub fn get_cell(&self, row: usize, col: usize) -> Option<&Cell> {
        self.rows.get(row).and_then(|r| r.get(col))
    }

    pub fn max_cols(&self) -> usize {
        self.rows.iter().map(|r| r.len()).max().unwrap_or(0)
    }
}

impl Workbook {
    pub fn new(initial_sheet_name: String) -> Self {
        Self {
            sheets: vec![Sheet::new(initial_sheet_name)],
        }
    }

    pub fn sheet_index_by_name(&self, name: &str) -> Option<usize> {
        self.sheets.iter().position(|s| s.name == name)
    }
}
