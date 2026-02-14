use mcp_ods::tools::{create_ods, get_sheets};
use serde_json::json;
use tempfile::tempdir;

#[test]
fn get_sheets_returns_created_sheet() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("get_sheets_unit.ods");

    create_ods::handle(json!({
        "path": path.to_string_lossy(),
        "overwrite": true,
        "initial_sheet_name": "Base"
    }))
    .expect("create");

    let out = get_sheets::handle(json!({ "path": path.to_string_lossy() })).expect("get_sheets");
    assert_eq!(out["sheets"], json!(["Base"]));
}

#[test]
fn get_sheets_returns_not_found_for_missing_file() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("missing_get_sheets.ods");

    let err = get_sheets::handle(json!({ "path": path.to_string_lossy() })).expect_err("missing");
    assert!(err.to_string().contains("file not found"));
}

#[test]
fn get_sheets_rejects_non_ods_path_extension() {
    let err = get_sheets::handle(json!({ "path": "demo.xlsx" })).expect_err("invalid extension");
    assert!(err.to_string().contains("expected .ods extension"));
}
