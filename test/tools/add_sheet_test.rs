use mcp_ods::tools::{add_sheet, create_ods, get_sheets};
use serde_json::json;
use tempfile::tempdir;

#[test]
fn add_sheet_appends_new_sheet() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("add_sheet_unit.ods");

    create_ods::handle(json!({
        "path": path.to_string_lossy(),
        "overwrite": true,
        "initial_sheet_name": "Hoja1"
    }))
    .expect("create");

    add_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "sheet_name": "Nueva",
        "position": "end"
    }))
    .expect("add");

    let out = get_sheets::handle(json!({ "path": path.to_string_lossy() })).expect("sheets");
    assert_eq!(out["sheets"], json!(["Hoja1", "Nueva"]));
}
