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

    assert!(!updated.contains("table:style-name=\"Default\""));
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

#[test]
fn set_cell_value_preserving_styles_raw_supports_self_closing_table() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1"/>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        1,
        1,
        &CellValue::String("X".to_string()),
    )
    .expect("set");

    assert!(updated.contains("<table:table table:name=\"Hoja1\">"));
    assert!(updated.contains("table:number-rows-repeated=\"1\""));
    assert!(updated.contains("<text:p>X</text:p>"));
}

#[test]
fn set_cell_value_preserving_styles_raw_updates_selected_sheet_index() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="S1"><table:table-row><table:table-cell/></table:table-row></table:table>
    <table:table table:name="S2"><table:table-row><table:table-cell/></table:table-row></table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        1,
        0,
        0,
        &CellValue::String("S2_ONLY".to_string()),
    )
    .expect("set");

    assert!(updated.contains("S2_ONLY"));
    // First sheet remains unchanged (only empty A1 cell).
    assert_eq!(updated.matches("S2_ONLY").count(), 1);
}

#[test]
fn set_cell_value_preserving_styles_raw_splits_repeated_empty_cell() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row>
        <table:table-cell table:number-columns-repeated="5"/>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        0,
        2,
        &CellValue::String("M".to_string()),
    )
    .expect("set");

    assert!(updated.contains("table:number-columns-repeated=\"2\""));
    assert!(updated.contains("<text:p>M</text:p>"));
}

#[test]
fn set_cell_value_preserving_styles_raw_rejects_repeated_non_empty_cell() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row>
        <table:table-cell table:number-columns-repeated="3" office:value-type="string"><text:p>v</text:p></table:table-cell>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let err = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        0,
        1,
        &CellValue::String("x".to_string()),
    )
    .expect_err("must fail");
    assert!(err
        .to_string()
        .contains("cannot safely edit repeated non-empty cell"));
}

#[test]
fn set_cell_value_preserving_styles_raw_preserves_style_attributes() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row>
        <table:table-cell table:style-name="ce5"/>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        0,
        0,
        &CellValue::String("Styled".to_string()),
    )
    .expect("set");

    assert!(updated.contains("table:style-name=\"ce5\""));
    assert!(updated.contains("<text:p>Styled</text:p>"));
}

#[test]
fn set_cell_value_preserving_styles_raw_supports_number_boolean_and_empty() {
    let base = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row><table:table-cell/></table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let number =
        ContentXml::set_cell_value_preserving_styles_raw(base, 0, 0, 0, &CellValue::Number(3.5))
            .expect("number");
    assert!(number.contains("office:value-type=\"float\""));
    assert!(number.contains("office:value=\"3.5\""));

    let boolean = ContentXml::set_cell_value_preserving_styles_raw(
        &number,
        0,
        0,
        0,
        &CellValue::Boolean(false),
    )
    .expect("boolean");
    assert!(boolean.contains("office:value-type=\"boolean\""));
    assert!(boolean.contains("office:boolean-value=\"false\""));

    let empty =
        ContentXml::set_cell_value_preserving_styles_raw(&boolean, 0, 0, 0, &CellValue::Empty)
            .expect("empty");
    assert!(empty.contains("<table:table-cell"));
    assert!(!empty.contains("office:value-type=\"boolean\""));
}

#[test]
fn set_cell_value_preserving_styles_raw_splits_repeated_row_and_updates_middle_copy() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row table:number-rows-repeated="3">
        <table:table-cell table:number-columns-repeated="2"/>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        1,
        1,
        &CellValue::String("MID".to_string()),
    )
    .expect("set");

    assert!(updated.contains("<text:p>MID</text:p>"));
    assert_eq!(updated.matches("<text:p>MID</text:p>").count(), 1);
}

#[test]
fn set_cell_value_preserving_styles_raw_returns_error_for_covered_cell_in_repeated_row() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row table:number-rows-repeated="2">
        <table:covered-table-cell/>
        <table:table-cell/>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let err = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        0,
        0,
        &CellValue::String("X".to_string()),
    )
    .expect_err("covered");
    assert!(err.to_string().contains("covered cell"));
}

#[test]
fn set_cell_value_preserving_styles_raw_writes_on_row_end_when_target_is_after_existing_cells() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row>
        <table:table-cell office:value-type="string"><text:p>A</text:p></table:table-cell>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        0,
        3,
        &CellValue::String("D1".to_string()),
    )
    .expect("set");
    assert!(updated.contains("<text:p>D1</text:p>"));
    assert!(updated.contains("table:number-columns-repeated=\"2\""));
}

#[test]
fn set_cell_value_preserving_styles_raw_writes_missing_rows_before_table_end() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row><table:table-cell/></table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        4,
        0,
        &CellValue::String("A5".to_string()),
    )
    .expect("set");

    assert!(updated.contains("table:number-rows-repeated=\"3\""));
    assert!(updated.contains("<text:p>A5</text:p>"));
}

#[test]
fn set_cell_value_preserving_styles_raw_replaces_nested_non_empty_cell_content() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row>
        <table:table-cell office:value-type="string">
          <text:p>old</text:p>
          <text:span>more</text:span>
        </table:table-cell>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        0,
        0,
        &CellValue::String("new".to_string()),
    )
    .expect("set");

    assert!(updated.contains("<text:p>new</text:p>"));
    assert!(!updated.contains("<text:span>more</text:span>"));
}

#[test]
fn set_cell_value_preserving_styles_raw_supports_start_covered_table_cell_nodes() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row>
        <table:covered-table-cell></table:covered-table-cell>
        <table:table-cell/>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        0,
        1,
        &CellValue::String("B1".to_string()),
    )
    .expect("set");
    assert!(updated.contains("<text:p>B1</text:p>"));
}

#[test]
fn set_cell_value_preserving_styles_raw_errors_when_sheet_index_does_not_exist() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1"><table:table-row><table:table-cell/></table:table-row></table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let err = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        3,
        0,
        0,
        &CellValue::String("x".to_string()),
    )
    .expect_err("missing sheet");
    assert!(err.to_string().contains("target cell could not be written"));
}

#[test]
fn add_sheet_preserving_styles_raw_inserts_at_end_without_dropping_blocks() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:automatic-styles><style:style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="ce1" style:family="table-cell"/></office:automatic-styles>
  <office:body><office:spreadsheet>
    <table:table table:name="S1"/>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::add_sheet_preserving_styles_raw(original, "S2", "end").expect("add");
    let names = ContentXml::sheet_names_from_content_raw(&updated).expect("names");
    assert_eq!(names, vec!["S1", "S2"]);
    assert!(updated.contains("<office:automatic-styles>"));
    assert!(updated.contains("style:name=\"ce1\""));
}

#[test]
fn add_sheet_preserving_styles_raw_inserts_at_start_and_escapes_name() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Base"/>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated =
        ContentXml::add_sheet_preserving_styles_raw(original, "A&B", "start").expect("add start");
    let names = ContentXml::sheet_names_from_content_raw(&updated).expect("names");
    assert_eq!(names, vec!["A&amp;B", "Base"]);
    assert!(updated.contains("<table:table table:name=\"A&amp;B\"/>"));
}

#[test]
fn add_sheet_preserving_styles_raw_rejects_duplicate_name() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="S1"/>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let err =
        ContentXml::add_sheet_preserving_styles_raw(original, "S1", "end").expect_err("duplicate");
    assert!(matches!(err, AppError::SheetNameAlreadyExists(_)));
}

#[test]
fn add_sheet_preserving_styles_raw_requires_at_least_one_table() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">
  <office:body><office:spreadsheet></office:spreadsheet></office:body>
</office:document-content>"#;

    let err =
        ContentXml::add_sheet_preserving_styles_raw(original, "S2", "end").expect_err("no table");
    assert!(err.to_string().contains("no table:table blocks found"));
}

#[test]
fn set_cell_value_preserving_styles_raw_handles_repeated_row_with_covered_cells_before_target() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row table:number-rows-repeated="2">
        <table:covered-table-cell table:number-columns-repeated="2"/>
        <table:table-cell/>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        0,
        2,
        &CellValue::String("C1".to_string()),
    )
    .expect("set");
    assert!(updated.contains("<text:p>C1</text:p>"));
}

#[test]
fn set_cell_value_preserving_styles_raw_rejects_target_inside_start_covered_cell_in_repeated_row() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row table:number-rows-repeated="2">
        <table:covered-table-cell></table:covered-table-cell>
        <table:table-cell/>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let err = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        1,
        0,
        &CellValue::String("X".to_string()),
    )
    .expect_err("covered");
    assert!(err.to_string().contains("covered cell"));
}

#[test]
fn set_cell_value_preserving_styles_raw_splits_repeated_empty_row_with_before_and_after() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row table:number-rows-repeated="5"/>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        2,
        1,
        &CellValue::String("B3".to_string()),
    )
    .expect("set");

    assert!(updated.contains("table:number-rows-repeated=\"2\""));
    assert!(updated.contains("<text:p>B3</text:p>"));
}

#[test]
fn set_cell_value_preserving_styles_raw_splits_repeated_empty_row_when_target_is_last() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row table:number-rows-repeated="3"/>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        2,
        0,
        &CellValue::String("A3".to_string()),
    )
    .expect("set");

    assert!(updated.contains("table:number-rows-repeated=\"2\""));
    assert!(updated.contains("<text:p>A3</text:p>"));
}

#[test]
fn set_cell_value_preserving_styles_raw_writes_inside_second_sheet_when_first_is_self_closing() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="S1"/>
    <table:table table:name="S2">
      <table:table-row><table:table-cell/></table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        1,
        0,
        0,
        &CellValue::String("S2_A1".to_string()),
    )
    .expect("set");

    assert!(updated.contains("S2_A1"));
    assert_eq!(updated.matches("S2_A1").count(), 1);
}

#[test]
fn set_cell_value_preserving_styles_raw_repeated_row_capture_rejects_non_empty_repeated_cell() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row table:number-rows-repeated="2">
        <table:table-cell table:number-columns-repeated="2" office:value-type="string">
          <text:p>v</text:p>
        </table:table-cell>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let err = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        1,
        1,
        &CellValue::String("X".to_string()),
    )
    .expect_err("must fail");
    assert!(err
        .to_string()
        .contains("cannot safely edit repeated non-empty cell"));
}

#[test]
fn set_cell_value_preserving_styles_raw_repeated_row_capture_rejects_target_inside_covered_range() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row table:number-rows-repeated="2">
        <table:covered-table-cell table:number-columns-repeated="2"></table:covered-table-cell>
        <table:table-cell/>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let err = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        1,
        0,
        &CellValue::String("X".to_string()),
    )
    .expect_err("must fail");
    assert!(err.to_string().contains("covered cell"));
}

#[test]
fn rename_first_sheet_name_raw_escapes_xml_sensitive_chars() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1"/>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::rename_first_sheet_name_raw(original, "A\"<>&B").expect("rename");
    assert!(updated.contains("table:name=\"A&quot;&lt;&gt;&amp;B\""));
}

#[test]
fn set_cell_value_preserving_styles_raw_repeated_row_capture_preserves_non_target_row_content() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row table:number-rows-repeated="3">
        <table:table-cell office:value-type="string">
          <text:p>base</text:p>
          <text:span>detail</text:span>
        </table:table-cell>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;

    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        1,
        1,
        &CellValue::String("B2".to_string()),
    )
    .expect("set");

    assert!(updated.contains("<text:span>detail</text:span>"));
    assert!(updated.contains("<text:p>B2</text:p>"));
}
