use crate::common::{create_base_ods, dispatch, new_ods_path};
use serde_json::json;

#[test]
fn rename_sheet_renames_target_sheet() {
    let (_dir, file_path) = new_ods_path("rename_sheet.ods");
    create_base_ods(&file_path, "Original");

    let out = dispatch(
        "rename_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "Original" },
            "new_sheet_name": "Renombrada"
        }),
    )
    .expect("rename");

    assert_eq!(out["sheets"], json!(["Renombrada"]));
}

#[test]
fn rename_sheet_rejects_duplicate_name() {
    let (_dir, file_path) = new_ods_path("rename_duplicate.ods");
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

    let err = dispatch(
        "rename_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "S2" },
            "new_sheet_name": "S1"
        }),
    )
    .expect_err("duplicate");
    assert!(err.to_string().contains("already exists"));
}
