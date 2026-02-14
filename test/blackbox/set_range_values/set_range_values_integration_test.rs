use crate::common::{create_base_ods, dispatch, new_ods_path};
use serde_json::json;

#[test]
fn set_range_values_writes_matrix_from_start_cell() {
    let (_dir, file_path) = new_ods_path("range.ods");
    create_base_ods(&file_path, "Hoja1");

    let result = dispatch(
        "set_range_values",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "start_cell": "B2",
            "data": [["A", "B"], ["1", "2"]]
        }),
    )
    .expect("set range");

    assert_eq!(result["rows_written"], 2);
    assert_eq!(result["cols_written"], 2);

    let content = dispatch(
        "get_sheet_content",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "mode": "matrix",
            "max_rows": 10,
            "max_cols": 10,
            "include_empty_trailing": false
        }),
    )
    .expect("content");

    assert_eq!(content["rows"], 3);
    assert_eq!(content["cols"], 3);
    assert_eq!(content["data"][1][1], "A");
    assert_eq!(content["data"][2][2], "2");
}

#[test]
fn set_range_values_accepts_sheet_by_name() {
    let (_dir, file_path) = new_ods_path("range_by_name.ods");
    create_base_ods(&file_path, "S1");
    dispatch(
        "add_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet_name": "S2",
            "position": "end"
        }),
    )
    .expect("add");

    dispatch(
        "set_range_values",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "S2" },
            "start_cell": "A1",
            "data": [["x","y"]]
        }),
    )
    .expect("set range");

    let content = dispatch(
        "get_sheet_content",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "S2" },
            "mode": "matrix",
            "max_rows": 10,
            "max_cols": 10,
            "include_empty_trailing": false
        }),
    )
    .expect("content");

    assert_eq!(content["data"], json!([["x", "y"]]));
}

#[test]
fn set_range_values_rejects_invalid_sheet() {
    let (_dir, file_path) = new_ods_path("range_bad_sheet.ods");
    create_base_ods(&file_path, "Hoja1");

    let err = dispatch(
        "set_range_values",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 4 },
            "start_cell": "A1",
            "data": [["v"]]
        }),
    )
    .expect_err("bad sheet");
    assert!(err.to_string().contains("sheet not found"));
}
