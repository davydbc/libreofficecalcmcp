use crate::common::errors::AppError;
use crate::common::fs::FsUtil;
use crate::common::json::JsonUtil;
use crate::ods::ods_file::OdsFile;
use serde::{Deserialize, Serialize};
use serde_json::Value;

#[derive(Debug, Deserialize)]
struct GetSheetsInput {
    path: String,
}

#[derive(Debug, Serialize)]
struct GetSheetsOutput {
    sheets: Vec<String>,
}

pub fn handle(params: Value) -> Result<Value, AppError> {
    // Returns sheet names preserving the same order as content.xml.
    let input: GetSheetsInput = JsonUtil::from_value(params)?;
    let path = FsUtil::resolve_ods_path(&input.path)?;
    if !path.exists() {
        return Err(AppError::FileNotFound(path.display().to_string()));
    }

    let workbook = OdsFile::read_workbook(&path)?;
    let sheets = workbook.sheets.into_iter().map(|s| s.name).collect();
    JsonUtil::to_value(GetSheetsOutput { sheets })
}
