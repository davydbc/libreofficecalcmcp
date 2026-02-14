use crate::common::errors::AppError;
use crate::common::fs::FsUtil;
use crate::common::json::JsonUtil;
use crate::ods::cell_address::CellAddress;
use crate::ods::ods_file::OdsFile;
use crate::ods::sheet_model::CellValue;
use serde::{Deserialize, Serialize};
use serde_json::Value;

#[derive(Debug, Deserialize)]
#[serde(untagged)]
enum SheetRef {
    Name { name: String },
    Index { index: usize },
}

#[derive(Debug, Deserialize)]
struct SetRangeValuesInput {
    path: String,
    sheet: SheetRef,
    start_cell: String,
    data: Vec<Vec<String>>,
}

#[derive(Debug, Serialize)]
struct SetRangeValuesOutput {
    updated: bool,
    rows_written: usize,
    cols_written: usize,
}

pub fn handle(params: Value) -> Result<Value, AppError> {
    // Writes a matrix starting at start_cell, expanding by rows and columns.
    let input: SetRangeValuesInput = JsonUtil::from_value(params)?;
    let path = FsUtil::resolve_ods_path(&input.path)?;
    if !path.exists() {
        return Err(AppError::FileNotFound(path.display().to_string()));
    }

    let mut workbook = OdsFile::read_workbook(&path)?;
    let (sheet_index, _) = resolve_sheet(&workbook, input.sheet)?;
    let start = CellAddress::parse(&input.start_cell)?;

    let rows = input.data.len();
    let cols = input.data.iter().map(|r| r.len()).max().unwrap_or(0);

    for (r_off, row) in input.data.iter().enumerate() {
        for (c_off, value) in row.iter().enumerate() {
            // Current implementation stores range data as string values.
            workbook.sheets[sheet_index]
                .ensure_cell_mut(start.row + r_off, start.col + c_off)
                .value = CellValue::String(value.clone());
        }
    }

    OdsFile::write_workbook(&path, &workbook)?;
    JsonUtil::to_value(SetRangeValuesOutput {
        updated: true,
        rows_written: rows,
        cols_written: cols,
    })
}

fn resolve_sheet(
    workbook: &crate::ods::sheet_model::Workbook,
    reference: SheetRef,
) -> Result<(usize, String), AppError> {
    // Sheet selectors accept either {name} or {index}.
    match reference {
        SheetRef::Name { name } => workbook
            .sheet_index_by_name(&name)
            .map(|idx| (idx, workbook.sheets[idx].name.clone()))
            .ok_or(AppError::SheetNotFound(name)),
        SheetRef::Index { index } => {
            if index >= workbook.sheets.len() {
                Err(AppError::SheetNotFound(index.to_string()))
            } else {
                Ok((index, workbook.sheets[index].name.clone()))
            }
        }
    }
}
