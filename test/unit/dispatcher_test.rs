use mcp_ods::mcp::dispatcher::Dispatcher;
use serde_json::json;

#[test]
fn dispatch_unknown_tool_returns_error() {
    let result = Dispatcher::dispatch("tool_que_no_existe", Some(json!({})));
    assert!(result.is_err());
}
