use mcp_ods::ods::sheet_model::{CellValue, Sheet, Workbook};

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
fn workbook_sheet_index_by_name_returns_expected_index() {
    let mut workbook = Workbook::new("Hoja1".to_string());
    workbook.sheets.push(Sheet::new("Datos".to_string()));
    assert_eq!(workbook.sheet_index_by_name("Datos"), Some(1));
    assert_eq!(workbook.sheet_index_by_name("NoExiste"), None);
}
