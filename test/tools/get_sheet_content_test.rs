use mcp_ods::tools::{create_ods, get_sheet_content, set_range_values};
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
