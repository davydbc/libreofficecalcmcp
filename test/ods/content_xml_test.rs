use mcp_ods::ods::content_xml::ContentXml;
use mcp_ods::ods::sheet_model::{CellValue, Workbook};

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
fn render_preserving_original_keeps_non_table_spreadsheet_nodes() {
    let mut workbook = Workbook::new("Hoja1".to_string());
    workbook.sheets[0].ensure_cell_mut(0, 0).value = CellValue::String("x".to_string());

    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2">
  <office:body>
    <office:spreadsheet>
      <table:calculation-settings table:case-sensitive="false"/>
      <table:table table:name="Old"/>
    </office:spreadsheet>
  </office:body>
</office:document-content>"#;

    let rendered =
        ContentXml::render_preserving_original(&workbook, original).expect("preserving render");
    assert!(rendered.contains("calculation-settings"));
    assert!(rendered.contains("table:name=\"Hoja1\""));
}
