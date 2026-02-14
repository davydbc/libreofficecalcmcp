use mcp_ods::tools::{create_ods, delete_sheet, get_sheets};
use serde_json::json;
use tempfile::tempdir;

#[test]
fn delete_sheet_removes_selected_sheet() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("delete_sheet_unit.ods");

    create_ods::handle(json!({
        "path": path.to_string_lossy(),
        "overwrite": true,
        "initial_sheet_name": "S1"
    }))
    .expect("create");
    mcp_ods::tools::add_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "sheet_name": "S2",
        "position": "end"
    }))
    .expect("add");

    delete_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": {"name": "S1"}
    }))
    .expect("delete");

    let out = get_sheets::handle(json!({ "path": path.to_string_lossy() })).expect("sheets");
    assert_eq!(out["sheets"], json!(["S2"]));
}

#[test]
fn delete_sheet_rejects_deleting_last_sheet() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("delete_last_sheet_unit.ods");

    create_ods::handle(json!({
        "path": path.to_string_lossy(),
        "overwrite": true
    }))
    .expect("create");

    let err = delete_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": {"index": 0}
    }))
    .expect_err("must fail");
    assert!(err.to_string().contains("last remaining sheet"));
}

#[test]
fn delete_sheet_returns_file_not_found_for_missing_file() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("missing_delete_sheet.ods");

    let err = delete_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": {"index": 0}
    }))
    .expect_err("missing");
    assert!(err.to_string().contains("file not found"));
}
