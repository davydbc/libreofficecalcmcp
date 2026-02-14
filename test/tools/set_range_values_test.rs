use mcp_ods::tools::{create_ods, get_sheet_content, set_range_values};
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
