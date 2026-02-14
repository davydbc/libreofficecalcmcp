use mcp_ods::tools::{create_ods, get_cell_value, set_cell_value};
use mcp_ods::ods::ods_file::OdsFile;
use serde_json::json;
use tempfile::tempdir;

#[test]
fn set_cell_value_updates_single_cell() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("set_cell_unit.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    set_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "B2",
        "value": { "type": "string", "data": "hola" }
    }))
    .expect("set");

    let out = get_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "B2"
    }))
    .expect("get");

    assert_eq!(out["value"], json!({"type":"string","data":"hola"}));
}

#[test]
fn set_cell_value_does_not_fill_previous_rows_when_target_is_d4() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("set_cell_d4_unit.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    set_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "D4",
        "value": { "type": "string", "data": "value" }
    }))
    .expect("set");

    for cell in ["D1", "D2", "D3"] {
        let out = get_cell_value::handle(json!({
            "path": path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": cell
        }))
        .expect("get");
        assert_eq!(
            out["value"],
            json!({"type":"empty"}),
            "expected {cell} to stay empty"
        );
    }

    let d4 = get_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "D4"
    }))
    .expect("get d4");
    assert_eq!(d4["value"], json!({"type":"string","data":"value"}));
}

#[test]
fn set_cell_value_returns_error_for_invalid_sheet_index() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("set_cell_bad_sheet.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    let err = set_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 4 },
        "cell": "A1",
        "value": { "type": "string", "data": "x" }
    }))
    .expect_err("sheet error");
    assert!(err.to_string().contains("sheet not found"));
}

#[test]
fn set_cell_value_supports_number_boolean_and_empty_values() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("set_cell_typed_values.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");

    set_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "name": "Hoja1" },
        "cell": "C3",
        "value": { "type": "number", "data": 123.5 }
    }))
    .expect("set number");
    let number = get_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "name": "Hoja1" },
        "cell": "C3"
    }))
    .expect("get number");
    assert_eq!(number["value"], json!({"type":"number","data":123.5}));

    set_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "name": "Hoja1" },
        "cell": "D4",
        "value": { "type": "boolean", "data": true }
    }))
    .expect("set bool");
    let boolean = get_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "name": "Hoja1" },
        "cell": "D4"
    }))
    .expect("get bool");
    assert_eq!(boolean["value"], json!({"type":"boolean","data":true}));

    set_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "name": "Hoja1" },
        "cell": "D4",
        "value": { "type": "empty" }
    }))
    .expect("set empty");
    let empty = get_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "name": "Hoja1" },
        "cell": "D4"
    }))
    .expect("get empty");
    assert_eq!(empty["value"], json!({"type":"empty"}));
}

#[test]
fn set_cell_value_returns_error_for_invalid_cell_address() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("set_cell_bad_address.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    let err = set_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "99Z",
        "value": { "type": "string", "data": "x" }
    }))
    .expect_err("address error");
    assert!(err.to_string().contains("invalid cell address"));
}

#[test]
fn set_cell_value_returns_file_not_found_for_missing_file() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("missing_set_cell.ods");

    let err = set_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "A1",
        "value": { "type": "string", "data": "x" }
    }))
    .expect_err("missing file");
    assert!(err.to_string().contains("file not found"));
}

#[test]
fn set_cell_value_propagates_resolve_anchor_xml_errors() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("set_cell_bad_content.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    OdsFile::write_content_xml(&path, "<broken").expect("write broken xml");

    let err = set_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "A1",
        "value": { "type": "string", "data": "x" }
    }))
    .expect_err("must fail");

    assert!(!err.to_string().is_empty());
}
