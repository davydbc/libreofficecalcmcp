use crate::common::{create_base_ods, dispatch, new_ods_path};
use serde_json::json;

#[test]
fn set_cell_value_updates_and_persists_value() {
    let (_dir, file_path) = new_ods_path("set_cell.ods");
    create_base_ods(&file_path, "Hoja1");

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "Hoja1" },
            "cell": "B2",
            "value": { "type": "string", "data": "hola" }
        }),
    )
    .expect("set_cell_value");

    let value = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "Hoja1" },
            "cell": "B2"
        }),
    )
    .expect("get_cell_value");

    assert_eq!(value["value"], json!({"type":"string","data":"hola"}));
}
