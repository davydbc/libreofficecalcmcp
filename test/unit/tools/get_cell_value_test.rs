use mcp_ods::tools::{create_ods, get_cell_value, set_cell_value};
use serde_json::json;
use tempfile::tempdir;

#[test]
fn get_cell_value_returns_empty_for_unset_cell() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("get_cell_unit.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    set_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "A1",
        "value": { "type": "string", "data": "hola" }
    }))
    .expect("set");

    let empty = get_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "Z9"
    }))
    .expect("get empty");

    assert_eq!(empty["value"], json!({"type":"empty"}));
}

#[test]
fn get_cell_value_rejects_invalid_address() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("get_cell_invalid_addr.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    let err = get_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "1A"
    }))
    .expect_err("invalid address");

    assert!(err.to_string().contains("invalid cell address"));
}

#[test]
fn get_cell_value_reads_number_and_boolean_types() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("get_cell_typed_values.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    set_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "A1",
        "value": { "type": "number", "data": 7.25 }
    }))
    .expect("set number");
    set_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "A2",
        "value": { "type": "boolean", "data": false }
    }))
    .expect("set bool");

    let number = get_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "A1"
    }))
    .expect("get number");
    let boolean = get_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "A2"
    }))
    .expect("get bool");

    assert_eq!(number["value"], json!({"type":"number","data":7.25}));
    assert_eq!(boolean["value"], json!({"type":"boolean","data":false}));
}
