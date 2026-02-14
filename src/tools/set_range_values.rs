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
struct SetRangeValuesInput {
    path: String,
    sheet: SheetRef,
    start_cell: String,
    data: Vec<Vec<String>>,
}

#[derive(Debug, Serialize)]
struct SetRangeValuesOutput {
    updated: bool,
    rows_written: usize,
    cols_written: usize,
}

pub fn handle(params: Value) -> Result<Value, AppError> {
    // Writes a matrix by patching content.xml in-place to preserve Calc-compatible structure.
    let input: SetRangeValuesInput = JsonUtil::from_value(params)?;
    let path = FsUtil::resolve_ods_path(&input.path)?;
    if !path.exists() {
        return Err(AppError::FileNotFound(path.display().to_string()));
    }

    let mut content_xml = OdsFile::read_content_xml(&path)?;
    let sheet_names = ContentXml::sheet_names_from_content_raw(&content_xml)?;
    let (sheet_index, _) = input.sheet.resolve_in_names(&sheet_names)?;
    let start = CellAddress::parse(&input.start_cell)?;

    let rows = input.data.len();
    let cols = input.data.iter().map(|r| r.len()).max().unwrap_or(0);

    for (r_off, row) in input.data.iter().enumerate() {
        for (c_off, value) in row.iter().enumerate() {
            let source_row = start.row + r_off;
            let source_col = start.col + c_off;
            let (target_row, target_col) = ContentXml::resolve_merged_anchor_raw(
                &content_xml,
                sheet_index,
                source_row,
                source_col,
            )?;
            content_xml = ContentXml::set_cell_value_preserving_styles_raw(
                &content_xml,
                sheet_index,
                target_row,
                target_col,
                &CellValue::String(value.clone()),
            )?;
        }
    }

    OdsFile::write_content_xml(&path, &content_xml)?;
    JsonUtil::to_value(SetRangeValuesOutput {
        updated: true,
        rows_written: rows,
        cols_written: cols,
    })
}
