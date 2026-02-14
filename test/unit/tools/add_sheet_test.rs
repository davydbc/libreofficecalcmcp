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

#[test]
fn add_sheet_rejects_duplicate_name() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("add_sheet_duplicate.ods");

    create_ods::handle(json!({
        "path": path.to_string_lossy(),
        "overwrite": true,
        "initial_sheet_name": "Hoja1"
    }))
    .expect("create");

    let err = add_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "sheet_name": "Hoja1",
        "position": "end"
    }))
    .expect_err("duplicate");
    assert!(err.to_string().contains("already exists"));
}

#[test]
fn add_sheet_can_insert_at_start() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("add_sheet_start.ods");

    create_ods::handle(json!({
        "path": path.to_string_lossy(),
        "overwrite": true,
        "initial_sheet_name": "Hoja1"
    }))
    .expect("create");

    add_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "sheet_name": "Primera",
        "position": "start"
    }))
    .expect("add");

    let out = get_sheets::handle(json!({ "path": path.to_string_lossy() })).expect("sheets");
    assert_eq!(out["sheets"], json!(["Primera", "Hoja1"]));
}

#[test]
fn add_sheet_returns_not_found_for_missing_file() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("missing_add_sheet.ods");

    let err = add_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "sheet_name": "Nueva"
    }))
    .expect_err("missing file");
    assert!(err.to_string().contains("file not found"));
}
