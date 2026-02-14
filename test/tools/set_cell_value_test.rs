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
