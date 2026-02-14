use mcp_ods::common::json::JsonUtil;
use mcp_ods::ods::sheet_model::CellValue;
use serde_json::json;

#[test]
fn json_roundtrip_works() {
    let value = json!({ "type": "string", "data": "hola" });
    let parsed: CellValue = JsonUtil::from_value(value.clone()).expect("parse");
    let output = JsonUtil::to_value(parsed).expect("serialize");
    assert_eq!(output, value);
}

#[test]
fn json_from_value_returns_invalid_input_error() {
    let value = json!({ "type": "number", "data": "not_a_number" });
    let err = JsonUtil::from_value::<CellValue>(value).expect_err("invalid");
    assert!(err.to_string().contains("invalid input"));
}
