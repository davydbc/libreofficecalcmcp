use crate::common::{create_base_ods, dispatch, new_ods_path};
use serde_json::json;

#[test]
fn get_cell_value_supports_string_number_boolean_and_empty() {
    let (_dir, file_path) = new_ods_path("values.ods");
    create_base_ods(&file_path, "Hoja1");

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "A1",
            "value": { "type": "string", "data": "txt" }
        }),
    )
    .expect("set string");

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "B1",
            "value": { "type": "number", "data": 3.5 }
        }),
    )
    .expect("set number");

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "C1",
            "value": { "type": "boolean", "data": true }
        }),
    )
    .expect("set bool");

    let a1 = dispatch(
        "get_cell_value",
        json!({ "path": file_path.to_string_lossy(), "sheet": { "index": 0 }, "cell": "A1" }),
    )
    .expect("a1");
    let b1 = dispatch(
        "get_cell_value",
        json!({ "path": file_path.to_string_lossy(), "sheet": { "index": 0 }, "cell": "B1" }),
    )
    .expect("b1");
    let c1 = dispatch(
        "get_cell_value",
        json!({ "path": file_path.to_string_lossy(), "sheet": { "index": 0 }, "cell": "C1" }),
    )
    .expect("c1");
    let z9 = dispatch(
        "get_cell_value",
        json!({ "path": file_path.to_string_lossy(), "sheet": { "index": 0 }, "cell": "Z9" }),
    )
    .expect("z9");

    assert_eq!(a1["value"], json!({"type":"string","data":"txt"}));
    assert_eq!(b1["value"], json!({"type":"number","data":3.5}));
    assert_eq!(c1["value"], json!({"type":"boolean","data":true}));
    assert_eq!(z9["value"], json!({"type":"empty"}));
}

#[test]
fn get_cell_value_rejects_invalid_cell_address() {
    let (_dir, file_path) = new_ods_path("invalid_cell_addr.ods");
    create_base_ods(&file_path, "Hoja1");

    let err = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "22A"
        }),
    )
    .expect_err("invalid address");
    assert!(err.to_string().contains("invalid cell address"));
}
