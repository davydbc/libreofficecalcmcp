use crate::common::errors::AppError;
use crate::common::fs::FsUtil;
use crate::common::json::JsonUtil;
use crate::ods::content_xml::ContentXml;
use crate::ods::ods_file::OdsFile;
use crate::tools::sheet_ref::SheetRef;
use serde::{Deserialize, Serialize};
use serde_json::Value;

#[derive(Debug, Deserialize)]
struct RenameSheetInput {
    path: String,
    sheet: SheetRef,
    new_sheet_name: String,
}

#[derive(Debug, Serialize)]
struct RenameSheetOutput {
    sheets: Vec<String>,
}

pub fn handle(params: Value) -> Result<Value, AppError> {
    // Renames a sheet in-place preserving all original styles and workbook structure.
    let input: RenameSheetInput = JsonUtil::from_value(params)?;
    let path = FsUtil::resolve_ods_path(&input.path)?;
    if !path.exists() {
        return Err(AppError::FileNotFound(path.display().to_string()));
    }

    let original_content = OdsFile::read_content_xml(&path)?;
    let source_name = input.sheet.as_name();
    let source_index = input.sheet.as_index();
    let updated_content = ContentXml::rename_sheet_preserving_styles_raw(
        &original_content,
        source_name,
        source_index,
        &input.new_sheet_name,
    )?;
    OdsFile::write_content_xml(&path, &updated_content)?;

    let sheets = ContentXml::sheet_names_from_content_raw(&updated_content)?;
    JsonUtil::to_value(RenameSheetOutput { sheets })
}
