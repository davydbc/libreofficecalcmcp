use crate::common::errors::AppError;
use crate::common::fs::FsUtil;
use crate::common::json::JsonUtil;
use crate::ods::ods_file::OdsFile;
use serde::{Deserialize, Serialize};
use serde_json::Value;

#[derive(Debug, Deserialize)]
struct CreateOdsInput {
    path: String,
    #[serde(default)]
    overwrite: bool,
    #[serde(default = "default_sheet_name")]
    initial_sheet_name: String,
}

#[derive(Debug, Serialize)]
struct CreateOdsOutput {
    path: String,
    sheets: Vec<String>,
}

fn default_sheet_name() -> String {
    "Hoja1".to_string()
}

pub fn handle(params: Value) -> Result<Value, AppError> {
    // 1) parse input, 2) validate path, 3) create ODS skeleton.
    let input: CreateOdsInput = JsonUtil::from_value(params)?;
    let path = FsUtil::resolve_ods_path(&input.path)?;

    if path.exists() && !input.overwrite {
        return Err(AppError::AlreadyExists(path.display().to_string()));
    }

    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)?;
    }

    OdsFile::create(&path, input.initial_sheet_name.clone())?;

    JsonUtil::to_value(CreateOdsOutput {
        path: path.display().to_string(),
        sheets: vec![input.initial_sheet_name],
    })
}
