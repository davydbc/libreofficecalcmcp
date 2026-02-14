use crate::common::{create_base_ods, dispatch, new_ods_path};
use serde_json::json;

#[test]
fn add_sheet_inserts_new_sheet_at_end() {
    let (_dir, file_path) = new_ods_path("add_sheet.ods");
    create_base_ods(&file_path, "Hoja1");

    let out = dispatch(
        "add_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet_name": "NuevaHoja",
            "position": "end"
        }),
    )
    .expect("add_sheet");

    assert_eq!(out["sheets"], json!(["Hoja1", "NuevaHoja"]));
}

#[test]
fn add_sheet_can_insert_at_start() {
    let (_dir, file_path) = new_ods_path("add_sheet_start.ods");
    create_base_ods(&file_path, "Hoja1");

    let out = dispatch(
        "add_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet_name": "Inicio",
            "position": "start"
        }),
    )
    .expect("add_sheet");

    assert_eq!(out["sheets"], json!(["Inicio", "Hoja1"]));
}

#[test]
fn add_sheet_rejects_duplicate_name() {
    let (_dir, file_path) = new_ods_path("add_sheet_duplicate.ods");
    create_base_ods(&file_path, "Hoja1");

    let err = dispatch(
        "add_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet_name": "Hoja1",
            "position": "end"
        }),
    )
    .expect_err("duplicate");
    assert!(err.to_string().contains("already exists"));
}
