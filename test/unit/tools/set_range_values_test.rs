use mcp_ods::tools::{add_sheet, create_ods, get_sheet_content, set_range_values};
use mcp_ods::ods::ods_file::OdsFile;
use serde_json::json;
use tempfile::tempdir;

#[test]
fn set_range_values_reports_written_shape() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("range_unit.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    let out = set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "start_cell": "B2",
        "data": [["1","2"],["3","4"]]
    }))
    .expect("set range");

    assert_eq!(out["rows_written"], 2);
    assert_eq!(out["cols_written"], 2);

    let content = get_sheet_content::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "mode": "matrix",
        "max_rows": 10,
        "max_cols": 10,
        "include_empty_trailing": false
    }))
    .expect("content");
    assert_eq!(content["data"][2][2], "4");
}

#[test]
fn set_range_values_rejects_invalid_start_cell() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("range_invalid_start.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    let err = set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "start_cell": "11",
        "data": [["a"]]
    }))
    .expect_err("invalid start");
    assert!(err.to_string().contains("invalid cell address"));
}

#[test]
fn set_range_values_accepts_sheet_name_reference() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("range_by_name.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    add_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "sheet_name": "S2",
        "position": "end"
    }))
    .expect("add");

    set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "name": "S2" },
        "start_cell": "A1",
        "data": [["x","y"]]
    }))
    .expect("set range");

    let content = get_sheet_content::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "name": "S2" },
        "mode": "matrix",
        "max_rows": 10,
        "max_cols": 10,
        "include_empty_trailing": false
    }))
    .expect("content");
    assert_eq!(content["data"], json!([["x", "y"]]));
}

#[test]
fn set_range_values_rejects_invalid_sheet_name() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("range_invalid_sheet.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    let err = set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "name": "NOPE" },
        "start_cell": "A1",
        "data": [["a"]]
    }))
    .expect_err("invalid sheet");
    assert!(err.to_string().contains("sheet not found"));
}

#[test]
fn set_range_values_returns_file_not_found_for_missing_file() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("missing_set_range.ods");

    let err = set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "start_cell": "A1",
        "data": [["x"]]
    }))
    .expect_err("missing file");
    assert!(err.to_string().contains("file not found"));
}

#[test]
fn set_range_values_reports_zero_columns_for_empty_data() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("range_empty_data.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    let out = set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "start_cell": "A1",
        "data": []
    }))
    .expect("set empty range");

    assert_eq!(out["rows_written"], 0);
    assert_eq!(out["cols_written"], 0);
}

#[test]
fn set_range_values_propagates_resolve_anchor_xml_errors() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("range_bad_content.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    OdsFile::write_content_xml(&path, "<broken").expect("write broken xml");

    let err = set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "start_cell": "A1",
        "data": [["x"]]
    }))
    .expect_err("must fail");
    assert!(!err.to_string().is_empty());
}

#[test]
fn set_range_values_propagates_set_cell_errors_for_repeated_non_empty_cell() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("range_repeated_non_empty.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row>
        <table:table-cell table:number-columns-repeated="2" office:value-type="string"><text:p>v</text:p></table:table-cell>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;
    OdsFile::write_content_xml(&path, xml).expect("write xml");

    let err = set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "start_cell": "A1",
        "data": [["x"]]
    }))
    .expect_err("must fail");
    assert!(err
        .to_string()
        .contains("cannot safely edit repeated non-empty cell"));
}
