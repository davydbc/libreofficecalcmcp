use crate::common::errors::AppError;
use crate::common::fs::FsUtil;
use crate::common::json::JsonUtil;
use crate::ods::cell_address::CellAddress;
use crate::ods::content_xml::ContentXml;
use crate::ods::ods_file::OdsFile;
use crate::ods::sheet_model::CellValue;
use crate::tools::sheet_ref::SheetRef;
use serde::{Deserialize, Serialize};
use serde_json::Value;

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

    let original_content = OdsFile::read_content_xml(&path)?;
    let sheet_names = ContentXml::sheet_names_from_content_raw(&original_content)?;
    let (sheet_index, sheet_name) = input.sheet.resolve_in_names(&sheet_names)?;
    let address = CellAddress::parse(&input.cell)?;

    let (target_row, target_col) = ContentXml::resolve_merged_anchor_raw(
        &original_content,
        sheet_index,
        address.row,
        address.col,
    )?;

    let updated_content = ContentXml::set_cell_value_preserving_styles_raw(
        &original_content,
        sheet_index,
        target_row,
        target_col,
        &input.value,
    )?;
    OdsFile::write_content_xml(&path, &updated_content)?;
    JsonUtil::to_value(SetCellValueOutput {
        updated: true,
        sheet: sheet_name,
        cell: input.cell,
    })
}
