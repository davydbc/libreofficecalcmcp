use crate::common::{create_base_ods, dispatch, new_ods_path};
use serde_json::json;

#[test]
fn get_sheets_add_sheet_and_duplicate_sheet_workflow() {
    let (_dir, file_path) = new_ods_path("sheets.ods");
    create_base_ods(&file_path, "Base");

    dispatch(
        "add_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet_name": "Datos",
            "position": "end"
        }),
    )
    .expect("add sheet");

    dispatch(
        "duplicate_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "source_sheet": { "name": "Datos" },
            "new_sheet_name": "Datos (copia)"
        }),
    )
    .expect("duplicate");

    let sheets = dispatch(
        "get_sheets",
        json!({
            "path": file_path.to_string_lossy()
        }),
    )
    .expect("get_sheets");

    assert_eq!(sheets["sheets"], json!(["Base", "Datos", "Datos (copia)"]));
}
