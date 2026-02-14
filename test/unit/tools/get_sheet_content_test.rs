use mcp_ods::tools::{create_ods, get_sheet_content, set_cell_value, set_range_values};
use serde_json::json;
use tempfile::tempdir;

#[test]
fn get_sheet_content_returns_matrix() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("content_unit.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "start_cell": "A1",
        "data": [["A","B"],["1","2"]]
    }))
    .expect("set range");

    let out = get_sheet_content::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "mode": "matrix",
        "max_rows": 10,
        "max_cols": 10,
        "include_empty_trailing": false
    }))
    .expect("content");

    assert_eq!(out["rows"], 2);
    assert_eq!(out["cols"], 2);
}

#[test]
fn get_sheet_content_rejects_unsupported_mode() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("content_invalid_mode.ods");
    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");

    let err = get_sheet_content::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "mode": "table"
    }))
    .expect_err("invalid mode");
    assert!(err.to_string().contains("mode=matrix"));
}

#[test]
fn get_sheet_content_keeps_trailing_area_when_requested() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("content_trailing.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "start_cell": "B2",
        "data": [["X"]]
    }))
    .expect("set");

    let out = get_sheet_content::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "mode": "matrix",
        "max_rows": 3,
        "max_cols": 3,
        "include_empty_trailing": true
    }))
    .expect("content");

    assert_eq!(out["rows"], 2);
    assert_eq!(out["cols"], 2);
    assert_eq!(out["data"][1][1], "X");
}

#[test]
fn get_sheet_content_stringifies_number_and_boolean_values() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("content_typed.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    set_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "A1",
        "value": { "type": "number", "data": 7.5 }
    }))
    .expect("set number");
    set_cell_value::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "cell": "A2",
        "value": { "type": "boolean", "data": true }
    }))
    .expect("set bool");

    let out = get_sheet_content::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "mode": "matrix",
        "max_rows": 5,
        "max_cols": 5,
        "include_empty_trailing": false
    }))
    .expect("content");

    assert_eq!(out["data"][0][0], "7.5");
    assert_eq!(out["data"][1][0], "true");
}

#[test]
fn get_sheet_content_uses_default_mode_and_limits() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("content_default_mode.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "start_cell": "A1",
        "data": [["v"]]
    }))
    .expect("set");

    let out = get_sheet_content::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 }
    }))
    .expect("content");
    assert_eq!(out["rows"], 1);
    assert_eq!(out["cols"], 1);
}

#[test]
fn get_sheet_content_returns_not_found_for_missing_file() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("missing_content.ods");

    let err = get_sheet_content::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 }
    }))
    .expect_err("missing file");
    assert!(err.to_string().contains("file not found"));
}
