use crate::common::errors::AppError;
use crate::common::fs::FsUtil;
use crate::common::json::JsonUtil;
use crate::ods::ods_file::OdsFile;
use crate::ods::sheet_model::Sheet;
use serde::{Deserialize, Serialize};
use serde_json::Value;

#[derive(Debug, Deserialize)]
struct AddSheetInput {
    path: String,
    sheet_name: String,
    #[serde(default = "default_position")]
    position: String,
}

#[derive(Debug, Serialize)]
struct AddSheetOutput {
    sheets: Vec<String>,
}

fn default_position() -> String {
    "end".to_string()
}

pub fn handle(params: Value) -> Result<Value, AppError> {
    let input: AddSheetInput = JsonUtil::from_value(params)?;
    let path = FsUtil::resolve_ods_path(&input.path)?;
    if !path.exists() {
        return Err(AppError::FileNotFound(path.display().to_string()));
    }

    let mut workbook = OdsFile::read_workbook(&path)?;
    if workbook.sheets.iter().any(|s| s.name == input.sheet_name) {
        return Err(AppError::SheetNameAlreadyExists(input.sheet_name));
    }

    let new_sheet = Sheet::new(input.sheet_name);
    if input.position.eq_ignore_ascii_case("start") {
        workbook.sheets.insert(0, new_sheet);
    } else {
        workbook.sheets.push(new_sheet);
    }

    OdsFile::write_workbook(&path, &workbook)?;
    JsonUtil::to_value(AddSheetOutput {
        sheets: workbook.sheets.into_iter().map(|s| s.name).collect(),
    })
}
