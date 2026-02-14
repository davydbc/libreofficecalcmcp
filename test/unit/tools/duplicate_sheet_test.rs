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

#[test]
fn duplicate_sheet_returns_error_for_missing_source() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("duplicate_missing_source.ods");

    create_ods::handle(json!({
        "path": path.to_string_lossy(),
        "overwrite": true,
        "initial_sheet_name": "Datos"
    }))
    .expect("create");

    let err = duplicate_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "source_sheet": { "name": "NoExiste" },
        "new_sheet_name": "X"
    }))
    .expect_err("missing source");
    assert!(err.to_string().contains("sheet not found"));
}

#[test]
fn duplicate_sheet_supports_source_by_index() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("duplicate_by_index.ods");

    create_ods::handle(json!({
        "path": path.to_string_lossy(),
        "overwrite": true,
        "initial_sheet_name": "S1"
    }))
    .expect("create");

    duplicate_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "source_sheet": { "index": 0 },
        "new_sheet_name": "S1_copy"
    }))
    .expect("duplicate");

    let out = get_sheets::handle(json!({ "path": path.to_string_lossy() })).expect("get sheets");
    assert_eq!(out["sheets"], json!(["S1", "S1_copy"]));
}

#[test]
fn duplicate_sheet_returns_file_not_found_for_missing_path() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("missing_duplicate.ods");

    let err = duplicate_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "source_sheet": { "index": 0 },
        "new_sheet_name": "Copy"
    }))
    .expect_err("missing file");

    assert!(err.to_string().contains("file not found"));
}
