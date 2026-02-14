use crate::common::{create_base_ods, dispatch, new_ods_path};
use serde_json::json;

#[test]
fn delete_sheet_removes_selected_sheet_by_name() {
    let (_dir, file_path) = new_ods_path("delete_sheet.ods");
    create_base_ods(&file_path, "Hoja1");
    dispatch(
        "add_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet_name": "Hoja2",
            "position": "end"
        }),
    )
    .expect("add");

    let out = dispatch(
        "delete_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "Hoja1" }
        }),
    )
    .expect("delete");

    assert_eq!(out["sheets"], json!(["Hoja2"]));
}

#[test]
fn delete_sheet_rejects_deleting_last_sheet() {
    let (_dir, file_path) = new_ods_path("delete_last_sheet.ods");
    create_base_ods(&file_path, "Unica");

    let err = dispatch(
        "delete_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 }
        }),
    )
    .expect_err("last sheet");
    assert!(err.to_string().contains("last remaining sheet"));
}
