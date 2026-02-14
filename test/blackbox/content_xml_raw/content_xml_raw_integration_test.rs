use mcp_ods::common::errors::AppError;
use mcp_ods::common::json::JsonUtil;
use mcp_ods::ods::cell_address::CellAddress;
use mcp_ods::ods::content_xml::ContentXml;
use mcp_ods::ods::manifest::Manifest;
use mcp_ods::ods::ods_templates::OdsTemplates;
use mcp_ods::ods::sheet_model::CellValue;
use serde_json::json;

#[test]
fn content_xml_raw_supports_self_closing_table_set_cell() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet><table:table table:name="Hoja1"/></office:spreadsheet></office:body>
</office:document-content>"#;
    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        1,
        1,
        &CellValue::String("X".to_string()),
    )
    .expect("set");
    assert!(updated.contains("<text:p>X</text:p>"));
}

#[test]
fn content_xml_raw_rejects_repeated_non_empty_cell() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet><table:table table:name="Hoja1"><table:table-row>
    <table:table-cell table:number-columns-repeated="3" office:value-type="string"><text:p>v</text:p></table:table-cell>
  </table:table-row></table:table></office:spreadsheet></office:body>
</office:document-content>"#;
    let err = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        0,
        1,
        &CellValue::String("n".to_string()),
    )
    .expect_err("must fail");
    assert!(err.to_string().contains("cannot safely edit repeated non-empty cell"));
}

#[test]
fn content_xml_raw_handles_repeated_row_split() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet><table:table table:name="Hoja1">
    <table:table-row table:number-rows-repeated="3"><table:table-cell table:number-columns-repeated="2"/></table:table-row>
  </table:table></office:spreadsheet></office:body>
</office:document-content>"#;
    let updated = ContentXml::set_cell_value_preserving_styles_raw(
        original,
        0,
        1,
        1,
        &CellValue::String("MID".to_string()),
    )
    .expect("set");
    assert_eq!(updated.matches("<text:p>MID</text:p>").count(), 1);
}

#[test]
fn content_xml_raw_handles_covered_cells_and_writes_after() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet><table:table table:name="Hoja1"><table:table-row>
    <table:covered-table-cell table:number-columns-repeated="2"/>
    <table:table-cell/>
  </table:table-row></table:table></office:spreadsheet></office:body>
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
fn content_xml_raw_errors_for_covered_cell_target_inside_repeated_row_capture() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet><table:table table:name="Hoja1">
    <table:table-row table:number-rows-repeated="2">
      <table:covered-table-cell></table:covered-table-cell><table:table-cell/>
    </table:table-row>
  </table:table></office:spreadsheet></office:body>
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
fn content_xml_raw_helpers_and_templates_are_available() {
    let manifest = Manifest::minimal_manifest_xml();
    assert!(manifest.contains("manifest:file-entry"));
    assert!(manifest.contains("content.xml"));

    assert_eq!(
        OdsTemplates::mimetype(),
        "application/vnd.oasis.opendocument.spreadsheet"
    );
    assert!(!OdsTemplates::empty_calc_template().is_empty());
    assert!(OdsTemplates::meta_xml().contains("document-meta"));
    assert!(OdsTemplates::styles_xml().contains("document-styles"));
    assert!(OdsTemplates::settings_xml().contains("document-settings"));
    assert!(OdsTemplates::manifest_xml().contains("manifest:manifest"));
    let rendered = OdsTemplates::content_xml("HojaInicial".to_string()).expect("render");
    assert!(rendered.contains("HojaInicial"));
}

#[test]
fn content_xml_raw_common_helpers_cover_errors_and_json_paths() {
    let parsed = CellAddress::parse("D4").expect("parse");
    assert_eq!(parsed.row, 3);
    assert_eq!(parsed.col, 3);
    let invalid = CellAddress::parse("44");
    assert!(invalid.is_err());

    let value = json!({"type":"string","data":"ok"});
    let parsed_value: CellValue = JsonUtil::from_value(value.clone()).expect("from");
    let roundtrip = JsonUtil::to_value(parsed_value).expect("to");
    assert_eq!(roundtrip, value);
    let bad = JsonUtil::from_value::<CellValue>(json!({"type":"number","data":"bad"}));
    assert!(bad.is_err());

    let all_errors = [
        AppError::InvalidPath("x".to_string()),
        AppError::FileNotFound("x".to_string()),
        AppError::AlreadyExists("x".to_string()),
        AppError::InvalidOdsFormat("x".to_string()),
        AppError::SheetNotFound("x".to_string()),
        AppError::SheetNameAlreadyExists("x".to_string()),
        AppError::InvalidCellAddress("x".to_string()),
        AppError::XmlParseError("x".to_string()),
        AppError::ZipError("x".to_string()),
        AppError::IoError("x".to_string()),
        AppError::InvalidInput("x".to_string()),
    ];
    for err in all_errors {
        assert!(err.code() >= 1001);
    }
}

#[test]
fn content_xml_raw_duplicate_sheet_supports_name_and_index() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body><office:spreadsheet><table:table table:name="S1"/><table:table table:name="S2"/></office:spreadsheet></office:body>
</office:document-content>"#;

    let by_name =
        ContentXml::duplicate_sheet_preserving_styles_raw(original, Some("S1"), None, "S1Copy")
            .expect("dup name");
    assert_eq!(
        ContentXml::sheet_names_from_content_raw(&by_name).expect("names"),
        vec!["S1", "S1Copy", "S2"]
    );

    let by_index =
        ContentXml::duplicate_sheet_preserving_styles_raw(original, None, Some(1), "S2Copy")
            .expect("dup index");
    assert_eq!(
        ContentXml::sheet_names_from_content_raw(&by_index).expect("names"),
        vec!["S1", "S2", "S2Copy"]
    );
}

#[test]
fn content_xml_raw_duplicate_sheet_validates_error_cases() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body><office:spreadsheet><table:table table:name="S1"/></office:spreadsheet></office:body>
</office:document-content>"#;

    let dup_name = ContentXml::duplicate_sheet_preserving_styles_raw(original, Some("S1"), None, "S1")
        .expect_err("dup name");
    assert!(dup_name.to_string().contains("already exists"));
    let no_source =
        ContentXml::duplicate_sheet_preserving_styles_raw(original, Some("X"), None, "N")
            .expect_err("missing source");
    assert!(no_source.to_string().contains("sheet not found"));
}

#[test]
fn content_xml_raw_rename_and_sheet_names_helpers_work() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body><office:spreadsheet><table:table table:name="Hoja1"/></office:spreadsheet></office:body>
</office:document-content>"#;
    let renamed = ContentXml::rename_first_sheet_name_raw(original, "Inicio").expect("rename");
    assert_eq!(
        ContentXml::sheet_names_from_content_raw(&renamed).expect("names"),
        vec!["Inicio"]
    );

    let invalid = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table/></office:spreadsheet></office:body></office:document-content>"#;
    let err = ContentXml::rename_first_sheet_name_raw(invalid, "X").expect_err("missing name");
    assert!(matches!(err, AppError::InvalidOdsFormat(_)));
}

#[test]
fn content_xml_raw_merged_anchor_maps_covered_cell_to_anchor() {
    let original = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row table:number-rows-repeated="4"/>
      <table:table-row><table:table-cell table:number-rows-spanned="2"><text:p>base</text:p></table:table-cell></table:table-row>
      <table:table-row><table:covered-table-cell/></table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;
    let anchor = ContentXml::resolve_merged_anchor_raw(original, 0, 5, 0).expect("anchor");
    assert_eq!(anchor, (4, 0));
}

#[test]
fn content_xml_raw_parse_and_render_workbook_roundtrip() {
    let xml = OdsTemplates::content_xml("Hoja1".to_string()).expect("render");
    let workbook = ContentXml::parse(&xml).expect("parse");
    let rendered = ContentXml::render(&workbook).expect("render");
    let names = ContentXml::sheet_names_from_content_raw(&rendered).expect("names");
    assert_eq!(names, vec!["Hoja1"]);
}

#[test]
fn content_xml_raw_app_error_conversions_are_mapped() {
    let io_err = std::io::Error::other("io");
    assert!(matches!(AppError::from(io_err), AppError::IoError(_)));
    let zip_err = zip::result::ZipError::FileNotFound;
    assert!(matches!(AppError::from(zip_err), AppError::ZipError(_)));

    let mut reader = quick_xml::Reader::from_str("<root><a></root>");
    let quick_xml_err = loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Eof) => panic!("expected xml error"),
            Ok(_) => continue,
            Err(err) => break err,
        }
    };
    assert!(matches!(
        AppError::from(quick_xml_err),
        AppError::XmlParseError(_)
    ));
}
