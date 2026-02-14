use crate::common::dispatch;
use mcp_ods::mcp::dispatcher::Dispatcher;
use serde_json::json;

#[test]
fn dispatcher_returns_error_for_unknown_tool() {
    let result = dispatch("tool_que_no_existe", json!({}));
    assert!(result.is_err());
}

#[test]
fn dispatcher_returns_clear_errors_for_invalid_inputs() {
    let dir = tempfile::tempdir().expect("tempdir");
    let file_path = dir.path().join("errors.ods");

    dispatch(
        "create_ods",
        json!({ "path": file_path.to_string_lossy(), "overwrite": true }),
    )
    .expect("create");

    let bad_mode = dispatch(
        "get_sheet_content",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "mode": "records"
        }),
    )
    .expect_err("invalid mode");
    assert!(bad_mode.to_string().contains("mode=matrix"));
}

#[test]
fn dispatcher_supports_initialized_notifications() {
    let out = Dispatcher::dispatch("initialized", None).expect("initialized");
    assert!(out.is_null());

    let out = Dispatcher::dispatch("notifications/initialized", None).expect("initialized ns");
    assert!(out.is_null());
}

#[test]
fn dispatcher_tools_call_validates_missing_payload_fields() {
    let missing_params = Dispatcher::dispatch("tools/call", None).expect_err("missing params");
    assert!(missing_params.to_string().contains("missing params"));

    let missing_tool = Dispatcher::dispatch("tools/call", Some(json!({}))).expect_err("name");
    assert!(missing_tool.to_string().contains("missing tool name"));
}

#[test]
fn dispatcher_allows_direct_tool_invocation_path() {
    let dir = tempfile::tempdir().expect("tempdir");
    let file_path = dir.path().join("direct_call.ods");

    Dispatcher::dispatch(
        "create_ods",
        Some(json!({ "path": file_path.to_string_lossy(), "overwrite": true })),
    )
    .expect("create");
    let out = Dispatcher::dispatch(
        "get_sheets",
        Some(json!({ "path": file_path.to_string_lossy() })),
    )
    .expect("get_sheets");
    assert_eq!(out["sheets"], json!(["Hoja1"]));
}

#[test]
fn dispatcher_tools_call_accepts_json_encoded_sheet_ref_string() {
    let dir = tempfile::tempdir().expect("tempdir");
    let file_path = dir.path().join("sheet_ref_string.ods");

    Dispatcher::dispatch(
        "tools/call",
        Some(json!({
            "name": "create_ods",
            "arguments": {
                "path": file_path.to_string_lossy(),
                "overwrite": true
            }
        })),
    )
    .expect("create");

    let set_result = Dispatcher::dispatch(
        "tools/call",
        Some(json!({
            "name": "set_cell_value",
            "arguments": {
                "path": file_path.to_string_lossy(),
                "sheet": "{\"name\":\"Hoja1\"}",
                "cell": "A1",
                "value": { "type": "string", "data": "ok" }
            }
        })),
    )
    .expect("set");
    assert_eq!(set_result["isError"], json!(false));

    let get_result = Dispatcher::dispatch(
        "tools/call",
        Some(json!({
            "name": "get_cell_value",
            "arguments": {
                "path": file_path.to_string_lossy(),
                "sheet": "{\"index\":\"0\"}",
                "cell": "A1"
            }
        })),
    )
    .expect("get");

    assert_eq!(get_result["isError"], json!(false));
    assert_eq!(
        get_result["structuredContent"]["value"],
        json!({"type":"string","data":"ok"})
    );
}
