use crate::common::errors::AppError;
use crate::common::fs::FsUtil;
use crate::common::json::JsonUtil;
use crate::ods::ods_file::OdsFile;
use serde::{Deserialize, Serialize};
use serde_json::Value;

#[derive(Debug, Deserialize)]
#[serde(untagged)]
enum SheetRef {
    Name { name: String },
    Index { index: usize },
}

#[derive(Debug, Deserialize)]
struct DuplicateSheetInput {
    path: String,
    source_sheet: SheetRef,
    new_sheet_name: String,
}

#[derive(Debug, Serialize)]
struct DuplicateSheetOutput {
    sheets: Vec<String>,
}

pub fn handle(params: Value) -> Result<Value, AppError> {
    let input: DuplicateSheetInput = JsonUtil::from_value(params)?;
    let path = FsUtil::resolve_ods_path(&input.path)?;
    if !path.exists() {
        return Err(AppError::FileNotFound(path.display().to_string()));
    }

    let mut workbook = OdsFile::read_workbook(&path)?;
    if workbook
        .sheets
        .iter()
        .any(|s| s.name == input.new_sheet_name)
    {
        return Err(AppError::SheetNameAlreadyExists(input.new_sheet_name));
    }

    let idx = match input.source_sheet {
        SheetRef::Name { name } => workbook
            .sheet_index_by_name(&name)
            .ok_or(AppError::SheetNotFound(name))?,
        SheetRef::Index { index } => {
            if index >= workbook.sheets.len() {
                return Err(AppError::SheetNotFound(index.to_string()));
            }
            index
        }
    };

    let mut copy = workbook.sheets[idx].clone();
    copy.name = input.new_sheet_name;
    workbook.sheets.insert(idx + 1, copy);

    OdsFile::write_workbook(&path, &workbook)?;
    JsonUtil::to_value(DuplicateSheetOutput {
        sheets: workbook.sheets.into_iter().map(|s| s.name).collect(),
    })
}
