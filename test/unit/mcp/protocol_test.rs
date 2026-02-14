use mcp_ods::mcp::protocol::{JsonRpcRequest, JsonRpcResponse};
use serde_json::json;

#[test]
fn protocol_request_deserializes_minimal_payload() {
    let payload = r#"{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"x":1}}"#;
    let req: JsonRpcRequest = serde_json::from_str(payload).expect("deserialize");
    assert_eq!(req.jsonrpc.as_deref(), Some("2.0"));
    assert_eq!(req.id, Some(json!(1)));
    assert_eq!(req.method, "initialize");
    assert_eq!(req.params, Some(json!({"x":1})));
}

#[test]
fn protocol_success_response_serializes_without_error_field() {
    let resp = JsonRpcResponse::success(Some(json!(5)), json!({"ok":true}));
    let value = serde_json::to_value(resp).expect("serialize");
    assert_eq!(value["jsonrpc"], "2.0");
    assert_eq!(value["id"], json!(5));
    assert_eq!(value["result"], json!({"ok":true}));
    assert!(value.get("error").is_none());
}

#[test]
fn protocol_failure_response_serializes_without_result_field() {
    let resp = JsonRpcResponse::failure(Some(json!("id")), 123, "boom".to_string());
    let value = serde_json::to_value(resp).expect("serialize");
    assert_eq!(value["jsonrpc"], "2.0");
    assert_eq!(value["id"], json!("id"));
    assert_eq!(value["error"]["code"], 123);
    assert_eq!(value["error"]["message"], "boom");
    assert!(value.get("result").is_none());
}
