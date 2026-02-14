use crate::common::errors::AppError;
use crate::common::fs::FsUtil;
use crate::common::json::JsonUtil;
use crate::ods::cell_address::CellAddress;
use crate::ods::content_xml::ContentXml;
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

    let original_content = OdsFile::read_content_xml(&path)?;
    let sheet_names = ContentXml::sheet_names_from_content_raw(&original_content)?;
    let (sheet_index, sheet_name) = resolve_sheet(&sheet_names, input.sheet)?;
    let address = CellAddress::parse(&input.cell)?;

    let updated_content = ContentXml::set_cell_value_preserving_styles_raw(
        &original_content,
        sheet_index,
        address.row,
        address.col,
        &input.value,
    )?;
    OdsFile::write_content_xml(&path, &updated_content)?;
    JsonUtil::to_value(SetCellValueOutput {
        updated: true,
        sheet: sheet_name,
        cell: input.cell,
    })
}

fn resolve_sheet(sheet_names: &[String], reference: SheetRef) -> Result<(usize, String), AppError> {
    // Shared helper to map name/index selectors into a concrete sheet index.
    match reference {
        SheetRef::Name { name } => sheet_names
            .iter()
            .position(|n| n == &name)
            .map(|idx| (idx, sheet_names[idx].clone()))
            .ok_or(AppError::SheetNotFound(name)),
        SheetRef::Index { index } => {
            if index >= sheet_names.len() {
                Err(AppError::SheetNotFound(index.to_string()))
            } else {
                Ok((index, sheet_names[index].clone()))
            }
        }
    }
}
