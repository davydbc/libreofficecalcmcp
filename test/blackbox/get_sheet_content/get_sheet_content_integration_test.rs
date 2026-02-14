use crate::common::{create_base_ods, dispatch, new_ods_path};
use serde_json::json;

#[test]
fn get_sheet_content_supports_trailing_trim_and_limits() {
    let (_dir, file_path) = new_ods_path("trim.ods");
    create_base_ods(&file_path, "Hoja1");

    dispatch(
        "set_range_values",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "start_cell": "A1",
            "data": [["x", "", ""], ["", "", ""]]
        }),
    )
    .expect("set range");

    let trimmed = dispatch(
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
    .expect("trimmed");
    assert_eq!(trimmed["rows"], 1);
    assert_eq!(trimmed["cols"], 1);

    let with_trailing = dispatch(
        "get_sheet_content",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "mode": "matrix",
            "max_rows": 1,
            "max_cols": 2,
            "include_empty_trailing": true
        }),
    )
    .expect("with trailing");
    assert_eq!(with_trailing["rows"], 1);
    assert_eq!(with_trailing["cols"], 2);
}
