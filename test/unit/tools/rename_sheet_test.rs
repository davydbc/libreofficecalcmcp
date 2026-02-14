use mcp_ods::tools::{create_ods, get_sheets, rename_sheet};
use serde_json::json;
use tempfile::tempdir;

#[test]
fn rename_sheet_renames_selected_sheet() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("rename_sheet_unit.ods");

    create_ods::handle(json!({
        "path": path.to_string_lossy(),
        "overwrite": true,
        "initial_sheet_name": "Original"
    }))
    .expect("create");

    rename_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": {"name": "Original"},
        "new_sheet_name": "Renombrada"
    }))
    .expect("rename");

    let out = get_sheets::handle(json!({ "path": path.to_string_lossy() })).expect("sheets");
    assert_eq!(out["sheets"], json!(["Renombrada"]));
}

#[test]
fn rename_sheet_rejects_duplicate_name() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("rename_sheet_duplicate.ods");

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

    let err = rename_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": {"name": "S2"},
        "new_sheet_name": "S1"
    }))
    .expect_err("duplicate");
    assert!(err.to_string().contains("already exists"));
}

#[test]
fn rename_sheet_returns_file_not_found_for_missing_file() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("missing_rename_sheet.ods");

    let err = rename_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": {"index": 0},
        "new_sheet_name": "X"
    }))
    .expect_err("missing");
    assert!(err.to_string().contains("file not found"));
}
