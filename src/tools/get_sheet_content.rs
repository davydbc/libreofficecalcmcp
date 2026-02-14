use crate::common::errors::AppError;
use crate::common::fs::FsUtil;
use crate::common::json::JsonUtil;
use crate::ods::ods_file::OdsFile;
use crate::ods::sheet_model::CellValue;
use crate::tools::sheet_ref::SheetRef;
use serde::{Deserialize, Serialize};
use serde_json::Value;

#[derive(Debug, Deserialize)]
struct GetSheetContentInput {
    path: String,
    sheet: SheetRef,
    #[serde(default = "default_mode")]
    mode: String,
    #[serde(default = "default_max_rows")]
    max_rows: usize,
    #[serde(default = "default_max_cols")]
    max_cols: usize,
    #[serde(default)]
    include_empty_trailing: bool,
}

#[derive(Debug, Serialize)]
struct GetSheetContentOutput {
    sheet: String,
    rows: usize,
    cols: usize,
    data: Vec<Vec<String>>,
}

fn default_mode() -> String {
    "matrix".to_string()
}
fn default_max_rows() -> usize {
    200
}
fn default_max_cols() -> usize {
    50
}

pub fn handle(params: Value) -> Result<Value, AppError> {
    // Builds a bounded 2D matrix representation optimized for LLM consumption.
    let input: GetSheetContentInput = JsonUtil::from_value(params)?;
    if input.mode != "matrix" {
        return Err(AppError::InvalidInput(
            "only mode=matrix is supported".to_string(),
        ));
    }

    let path = FsUtil::resolve_ods_path(&input.path)?;
    if !path.exists() {
        return Err(AppError::FileNotFound(path.display().to_string()));
    }

    let workbook = OdsFile::read_workbook(&path)?;
    let (sheet_index, sheet_name) = input.sheet.resolve_in_workbook(&workbook)?;
    let sheet = &workbook.sheets[sheet_index];

    let row_limit = std::cmp::min(sheet.rows.len(), input.max_rows);
    let col_limit = std::cmp::min(sheet.max_cols(), input.max_cols);

    let mut matrix = Vec::with_capacity(row_limit);
    for r in 0..row_limit {
        let mut row_values = Vec::with_capacity(col_limit);
        for c in 0..col_limit {
            let text = sheet
                .get_cell(r, c)
                .map(|c| value_as_string(&c.value))
                .unwrap_or_default();
            row_values.push(text);
        }
        matrix.push(row_values);
    }

    let (rows, cols, data) = if input.include_empty_trailing {
        (matrix.len(), col_limit, matrix)
    } else {
        trim_trailing(matrix)
    };

    JsonUtil::to_value(GetSheetContentOutput {
        sheet: sheet_name,
        rows,
        cols,
        data,
    })
}

fn value_as_string(value: &CellValue) -> String {
    match value {
        CellValue::String(v) => v.clone(),
        CellValue::Number(v) => v.to_string(),
        CellValue::Boolean(v) => v.to_string(),
        CellValue::Empty => String::new(),
    }
}

fn trim_trailing(mut matrix: Vec<Vec<String>>) -> (usize, usize, Vec<Vec<String>>) {
    // Removes trailing empty rows/columns to reduce output noise and token usage.
    while matrix
        .last()
        .map(|r| r.iter().all(|v| v.is_empty()))
        .unwrap_or(false)
    {
        matrix.pop();
    }

    let mut max_col = 0usize;
    for row in &matrix {
        if let Some(last) = row.iter().rposition(|v| !v.is_empty()) {
            max_col = max_col.max(last + 1);
        }
    }

    for row in &mut matrix {
        row.truncate(max_col);
    }

    (matrix.len(), max_col, matrix)
}
