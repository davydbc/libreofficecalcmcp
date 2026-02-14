use mcp_ods::mcp::dispatcher::Dispatcher;
use serde_json::json;

#[test]
fn dispatch_unknown_tool_returns_error() {
    let result = Dispatcher::dispatch("tool_que_no_existe", Some(json!({})));
    assert!(result.is_err());
}

#[test]
fn dispatch_initialize_returns_server_info() {
    let result = Dispatcher::dispatch("initialize", Some(json!({}))).expect("initialize");
    assert_eq!(result["capabilities"]["tools"], json!({}));
    assert_eq!(result["serverInfo"]["name"], "libreoffice-calc-mcp");
}

#[test]
fn dispatch_tools_list_returns_registered_tools() {
    let result = Dispatcher::dispatch("tools/list", Some(json!({}))).expect("tools/list");
    let tools = result["tools"].as_array().expect("tools array");
    assert!(tools.iter().any(|t| t["name"] == "create_ods"));
    assert!(tools.iter().any(|t| t["name"] == "set_range_values"));
    assert!(tools.iter().any(|t| t["name"] == "delete_sheet"));
    assert!(tools.iter().any(|t| t["name"] == "rename_sheet"));
}
