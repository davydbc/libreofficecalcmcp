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
struct GetCellValueInput {
    path: String,
    sheet: SheetRef,
    cell: String,
}

#[derive(Debug, Serialize)]
struct GetCellValueOutput {
    sheet: String,
    cell: String,
    value: CellValue,
}

pub fn handle(params: Value) -> Result<Value, AppError> {
    // Fast path when caller needs only one cell and not the full matrix.
    let input: GetCellValueInput = JsonUtil::from_value(params)?;
    let path = FsUtil::resolve_ods_path(&input.path)?;
    if !path.exists() {
        return Err(AppError::FileNotFound(path.display().to_string()));
    }

    let workbook = OdsFile::read_workbook(&path)?;
    let (sheet_index, sheet_name) = resolve_sheet(&workbook, input.sheet)?;

    let address = CellAddress::parse(&input.cell)?;
    let value = workbook.sheets[sheet_index]
        .get_cell(address.row, address.col)
        .map(|c| c.value.clone())
        .unwrap_or(CellValue::Empty);

    JsonUtil::to_value(GetCellValueOutput {
        sheet: sheet_name,
        cell: input.cell,
        value,
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
