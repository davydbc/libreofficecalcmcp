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
