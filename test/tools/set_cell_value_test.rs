use mcp_ods::tools::{create_ods, get_cell_value, set_cell_value};
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
