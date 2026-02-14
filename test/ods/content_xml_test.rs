use mcp_ods::common::errors::AppError;
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

#[test]
fn set_cell_value_preserving_styles_raw_does_not_expand_row_cells_individually() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2">
  <office:automatic-styles>
    <style:style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="ce1" style:family="table-cell"/>
  </office:automatic-styles>
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Hoja1">
        <table:table-column table:default-cell-style-name="ce1"/>
        <table:table-row>
          <table:table-cell/>
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        0,
        2,
        &CellValue::String("VALOR_DADO".to_string()),
    )
    .expect("update C1");

    assert!(updated.contains("<table:table-cell table:style-name=\"Default\"/>"));
    assert!(updated.contains("office:value-type=\"string\""));
}

#[test]
fn set_cell_value_preserving_styles_raw_does_not_materialize_many_rows() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2">
  <office:automatic-styles>
    <style:style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="ce1" style:family="table-cell"/>
  </office:automatic-styles>
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Hoja1">
        <table:table-column table:default-cell-style-name="ce1"/>
        <table:table-row>
          <table:table-cell/>
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        7,
        7,
        &CellValue::String("VALOR_DADO".to_string()),
    )
    .expect("update H8");

    assert!(updated.contains("table:number-rows-repeated=\"6\""));
    assert!(updated.contains("table:number-columns-repeated=\"7\""));
}

#[test]
fn resolve_merged_anchor_raw_maps_covered_cell_to_top_left_anchor() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Hoja1">
        <table:table-row table:number-rows-repeated="4"/>
        <table:table-row>
          <table:table-cell table:number-rows-spanned="2" office:value-type="string">
            <text:p>MERGED</text:p>
          </table:table-cell>
          <table:table-cell/>
        </table:table-row>
        <table:table-row>
          <table:covered-table-cell/>
          <table:table-cell/>
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>"#;

    let anchor = ContentXml::resolve_merged_anchor_raw(original, 0, 5, 0).expect("resolve");
    assert_eq!(anchor, (4, 0));

    let regular = ContentXml::resolve_merged_anchor_raw(original, 0, 5, 1).expect("resolve");
    assert_eq!(regular, (5, 1));
}

#[test]
fn duplicate_sheet_preserving_styles_raw_supports_name_and_index() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.2">
  <office:body><office:spreadsheet>
    <table:table table:name="S1"/><table:table table:name="S2"/>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let by_name =
        ContentXml::duplicate_sheet_preserving_styles_raw(original, Some("S1"), None, "S1Copy")
            .expect("dup name");
    let names = ContentXml::sheet_names_from_content_raw(&by_name).expect("names");
    assert_eq!(names, vec!["S1", "S1Copy", "S2"]);

    let by_index =
        ContentXml::duplicate_sheet_preserving_styles_raw(original, None, Some(1), "S2Copy")
            .expect("dup index");
    let names = ContentXml::sheet_names_from_content_raw(&by_index).expect("names");
    assert_eq!(names, vec!["S1", "S2", "S2Copy"]);
}

#[test]
fn duplicate_sheet_preserving_styles_raw_validates_inputs() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.2">
  <office:body><office:spreadsheet>
    <table:table table:name="S1"/>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let err = ContentXml::duplicate_sheet_preserving_styles_raw(original, Some("S1"), None, "S1")
        .expect_err("duplicate name");
    assert!(err.to_string().contains("sheet name already exists"));

    let err = ContentXml::duplicate_sheet_preserving_styles_raw(original, Some("XX"), None, "N")
        .expect_err("missing source");
    assert!(err.to_string().contains("sheet not found"));

    let err = ContentXml::duplicate_sheet_preserving_styles_raw(original, None, Some(2), "N")
        .expect_err("index out");
    assert!(err.to_string().contains("sheet not found"));
}

#[test]
fn rename_first_sheet_name_raw_requires_table_name_attribute() {
    let invalid = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table/></office:spreadsheet></office:body></office:document-content>"#;
    let err = ContentXml::rename_first_sheet_name_raw(invalid, "New").expect_err("missing name");
    assert!(matches!(err, AppError::InvalidOdsFormat(_)));
}
