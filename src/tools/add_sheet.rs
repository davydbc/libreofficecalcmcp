use crate::common::errors::AppError;
use crate::common::fs::FsUtil;
use crate::common::json::JsonUtil;
use crate::ods::content_xml::ContentXml;
use crate::ods::ods_file::OdsFile;
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
    // Adds a new sheet by patching content.xml directly to preserve Calc XML structures.
    let input: AddSheetInput = JsonUtil::from_value(params)?;
    let path = FsUtil::resolve_ods_path(&input.path)?;
    if !path.exists() {
        return Err(AppError::FileNotFound(path.display().to_string()));
    }

    let original_content = OdsFile::read_content_xml(&path)?;
    let updated_content = ContentXml::add_sheet_preserving_styles_raw(
        &original_content,
        &input.sheet_name,
        &input.position,
    )?;
    OdsFile::write_content_xml(&path, &updated_content)?;
    let sheets = ContentXml::sheet_names_from_content_raw(&updated_content)?;

    JsonUtil::to_value(AddSheetOutput {
        sheets,
    })
}
