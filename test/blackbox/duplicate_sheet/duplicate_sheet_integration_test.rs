use crate::common::{create_base_ods, dispatch, new_ods_path};
use serde_json::json;

#[test]
fn duplicate_sheet_creates_named_copy() {
    let (_dir, file_path) = new_ods_path("duplicate.ods");
    create_base_ods(&file_path, "Datos");

    let out = dispatch(
        "duplicate_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "source_sheet": { "name": "Datos" },
            "new_sheet_name": "Datos (copia)"
        }),
    )
    .expect("duplicate_sheet");

    assert_eq!(out["sheets"], json!(["Datos", "Datos (copia)"]));
}
