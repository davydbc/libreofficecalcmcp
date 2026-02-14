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
struct SetCellValueInput {
    path: String,
    sheet: SheetRef,
    cell: String,
    value: CellValue,
}

#[derive(Debug, Serialize)]
struct SetCellValueOutput {
    updated: bool,
    sheet: String,
    cell: String,
}

pub fn handle(params: Value) -> Result<Value, AppError> {
    // Updates one address and keeps untouched cells/styles as-is in XML model.
    let input: SetCellValueInput = JsonUtil::from_value(params)?;
    let path = FsUtil::resolve_ods_path(&input.path)?;
    if !path.exists() {
        return Err(AppError::FileNotFound(path.display().to_string()));
    }

    let mut workbook = OdsFile::read_workbook(&path)?;
    let (sheet_index, sheet_name) = resolve_sheet(&workbook, input.sheet)?;
    let address = CellAddress::parse(&input.cell)?;

    workbook.sheets[sheet_index]
        .ensure_cell_mut(address.row, address.col)
        .value = input.value;

    OdsFile::write_workbook(&path, &workbook)?;
    JsonUtil::to_value(SetCellValueOutput {
        updated: true,
        sheet: sheet_name,
        cell: input.cell,
    })
}

fn resolve_sheet(
    workbook: &crate::ods::sheet_model::Workbook,
    reference: SheetRef,
) -> Result<(usize, String), AppError> {
    // Shared helper to map name/index selectors into a concrete sheet index.
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
