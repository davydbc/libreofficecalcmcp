use mcp_ods::tools::{add_sheet, create_ods, get_sheet_content, set_range_values};
use serde_json::json;
use tempfile::tempdir;

#[test]
fn set_range_values_reports_written_shape() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("range_unit.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    let out = set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "start_cell": "B2",
        "data": [["1","2"],["3","4"]]
    }))
    .expect("set range");

    assert_eq!(out["rows_written"], 2);
    assert_eq!(out["cols_written"], 2);

    let content = get_sheet_content::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "mode": "matrix",
        "max_rows": 10,
        "max_cols": 10,
        "include_empty_trailing": false
    }))
    .expect("content");
    assert_eq!(content["data"][2][2], "4");
}

#[test]
fn set_range_values_rejects_invalid_start_cell() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("range_invalid_start.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    let err = set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "index": 0 },
        "start_cell": "11",
        "data": [["a"]]
    }))
    .expect_err("invalid start");
    assert!(err.to_string().contains("invalid cell address"));
}

#[test]
fn set_range_values_accepts_sheet_name_reference() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("range_by_name.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    add_sheet::handle(json!({
        "path": path.to_string_lossy(),
        "sheet_name": "S2",
        "position": "end"
    }))
    .expect("add");

    set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "name": "S2" },
        "start_cell": "A1",
        "data": [["x","y"]]
    }))
    .expect("set range");

    let content = get_sheet_content::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "name": "S2" },
        "mode": "matrix",
        "max_rows": 10,
        "max_cols": 10,
        "include_empty_trailing": false
    }))
    .expect("content");
    assert_eq!(content["data"], json!([["x", "y"]]));
}

#[test]
fn set_range_values_rejects_invalid_sheet_name() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("range_invalid_sheet.ods");

    create_ods::handle(json!({ "path": path.to_string_lossy(), "overwrite": true }))
        .expect("create");
    let err = set_range_values::handle(json!({
        "path": path.to_string_lossy(),
        "sheet": { "name": "NOPE" },
        "start_cell": "A1",
        "data": [["a"]]
    }))
    .expect_err("invalid sheet");
    assert!(err.to_string().contains("sheet not found"));
}
