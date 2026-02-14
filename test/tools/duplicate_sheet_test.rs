use mcp_ods::tools::{create_ods, duplicate_sheet, get_sheets};
use serde_json::json;
use tempfile::tempdir;

#[test]
fn duplicate_sheet_creates_copy_after_source() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("duplicate_unit.ods");

    create_ods::handle(json!({
        "path": path.to_string_lossy(),
        "overwrite": true,
        "initial_sheet_name": "Datos"
    }))
    .expect("create");

    duplicate_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "source_sheet": { "name": "Datos" },
        "new_sheet_name": "Datos (copia)"
    }))
    .expect("duplicate");

    let out = get_sheets::handle(json!({ "path": path.to_string_lossy() })).expect("get sheets");
    assert_eq!(out["sheets"], json!(["Datos", "Datos (copia)"]));
}
