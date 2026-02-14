use mcp_ods::common::errors::AppError;
use mcp_ods::common::fs::FsUtil;
use mcp_ods::common::json::JsonUtil;
use mcp_ods::ods::cell_address::CellAddress;
use mcp_ods::ods::content_xml::ContentXml;
use mcp_ods::ods::sheet_model::{CellValue, Sheet, Workbook};
use serde_json::json;

#[test]
fn parse_and_format_cell_address() {
    let a1 = CellAddress::parse("C12").expect("valid address");
    assert_eq!(a1.row, 11);
    assert_eq!(a1.col, 2);
    assert_eq!(a1.to_a1(), "C12");
}

#[test]
fn invalid_cell_address_is_rejected() {
    let err = CellAddress::parse("12A").expect_err("invalid");
    assert!(matches!(err, AppError::InvalidCellAddress(_)));
}

#[test]
fn resolve_ods_path_requires_ods_extension() {
    let err = FsUtil::resolve_ods_path("demo.xlsx").expect_err("should fail");
    assert!(matches!(err, AppError::InvalidPath(_)));
}

#[test]
fn json_roundtrip_works() {
    let value = json!({
        "type": "string",
        "data": "hola"
    });
    let parsed: CellValue = JsonUtil::from_value(value.clone()).expect("parse");
    let output = JsonUtil::to_value(parsed).expect("serialize");
    assert_eq!(output, value);
}

#[test]
fn sheet_ensure_cell_mut_expands_matrix() {
    let mut sheet = Sheet::new("Hoja1".to_string());
    sheet.ensure_cell_mut(2, 3).value = CellValue::String("x".to_string());
    assert_eq!(sheet.rows.len(), 3);
    assert_eq!(sheet.max_cols(), 4);
    assert_eq!(
        sheet.get_cell(2, 3).expect("cell").value,
        CellValue::String("x".to_string())
    );
}

#[test]
fn content_xml_render_and_parse_preserves_basic_values() {
    let mut workbook = Workbook::new("Hoja1".to_string());
    workbook.sheets[0].ensure_cell_mut(0, 0).value = CellValue::String("hola".to_string());
    workbook.sheets[0].ensure_cell_mut(0, 1).value = CellValue::Number(42.0);
    workbook.sheets[0].ensure_cell_mut(1, 0).value = CellValue::Boolean(true);

    let xml = ContentXml::render(&workbook).expect("render");
    let parsed = ContentXml::parse(&xml).expect("parse");

    assert_eq!(parsed.sheets.len(), 1);
    assert_eq!(
        parsed.sheets[0].get_cell(0, 0).expect("a1").value,
        CellValue::String("hola".to_string())
    );
    assert_eq!(
        parsed.sheets[0].get_cell(0, 1).expect("b1").value,
        CellValue::Number(42.0)
    );
    assert_eq!(
        parsed.sheets[0].get_cell(1, 0).expect("a2").value,
        CellValue::Boolean(true)
    );
}

#[test]
fn workbook_sheet_index_by_name_returns_expected_index() {
    let mut workbook = Workbook::new("Hoja1".to_string());
    workbook.sheets.push(Sheet::new("Datos".to_string()));
    assert_eq!(workbook.sheet_index_by_name("Datos"), Some(1));
    assert_eq!(workbook.sheet_index_by_name("NoExiste"), None);
}
