use serde_json::{json, Value};
use std::io::{BufRead, BufReader, Write};
use std::process::{Command, Stdio};

#[test]
fn server_stdio_supports_initialize_and_tools_call_flow() {
    let exe = env!("CARGO_BIN_EXE_mcp-ods");
    let mut child = Command::new(exe)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .spawn()
        .expect("spawn mcp");

    let mut stdin = child.stdin.take().expect("stdin");
    writeln!(
        stdin,
        "{}",
        json!({"jsonrpc":"2.0","id":1,"method":"initialize","params":{}})
    )
    .expect("init");
    writeln!(
        stdin,
        "{}",
        json!({"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}})
    )
    .expect("list");
    drop(stdin);

    let stdout = child.stdout.take().expect("stdout");
    let lines: Vec<String> = BufReader::new(stdout)
        .lines()
        .collect::<Result<_, _>>()
        .expect("read lines");
    let status = child.wait().expect("wait");
    assert!(status.success());
    assert_eq!(lines.len(), 2);

    let init: Value = serde_json::from_str(&lines[0]).expect("json");
    assert_eq!(init["id"], 1);
    assert_eq!(init["result"]["serverInfo"]["name"], "libreoffice-calc-mcp");

    let list: Value = serde_json::from_str(&lines[1]).expect("json");
    assert_eq!(list["id"], 2);
    assert!(list["result"]["tools"].as_array().expect("tools").len() >= 8);
}

#[test]
fn server_stdio_returns_application_error_for_unknown_tool() {
    let exe = env!("CARGO_BIN_EXE_mcp-ods");
    let mut child = Command::new(exe)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .spawn()
        .expect("spawn mcp");

    let mut stdin = child.stdin.take().expect("stdin");
    writeln!(
        stdin,
        "{}",
        json!({"jsonrpc":"2.0","id":7,"method":"tools/call","params":{"name":"missing_tool","arguments":{}}})
    )
    .expect("request");
    drop(stdin);

    let stdout = child.stdout.take().expect("stdout");
    let lines: Vec<String> = BufReader::new(stdout)
        .lines()
        .collect::<Result<_, _>>()
        .expect("read lines");
    let status = child.wait().expect("wait");
    assert!(status.success());
    assert_eq!(lines.len(), 1);

    let error: Value = serde_json::from_str(&lines[0]).expect("json");
    assert_eq!(error["id"], 7);
    assert_eq!(error["error"]["code"], 1011);
}

#[test]
fn server_stdio_ignores_notifications_without_id() {
    let exe = env!("CARGO_BIN_EXE_mcp-ods");
    let mut child = Command::new(exe)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .spawn()
        .expect("spawn mcp");

    let mut stdin = child.stdin.take().expect("stdin");
    writeln!(
        stdin,
        "{}",
        json!({"jsonrpc":"2.0","method":"initialized","params":{}})
    )
    .expect("notification");
    writeln!(
        stdin,
        "{}",
        json!({"jsonrpc":"2.0","id":9,"method":"tools/list","params":{}})
    )
    .expect("request");
    drop(stdin);

    let stdout = child.stdout.take().expect("stdout");
    let lines: Vec<String> = BufReader::new(stdout)
        .lines()
        .collect::<Result<_, _>>()
        .expect("read lines");
    let status = child.wait().expect("wait");
    assert!(status.success());
    assert_eq!(lines.len(), 1);

    let list: Value = serde_json::from_str(&lines[0]).expect("json");
    assert_eq!(list["id"], 9);
}

#[test]
fn server_stdio_returns_error_when_tools_call_payload_is_invalid() {
    let exe = env!("CARGO_BIN_EXE_mcp-ods");
    let mut child = Command::new(exe)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .spawn()
        .expect("spawn mcp");

    let mut stdin = child.stdin.take().expect("stdin");
    writeln!(
        stdin,
        "{}",
        json!({"jsonrpc":"2.0","id":10,"method":"tools/call","params":{}})
    )
    .expect("request");
    drop(stdin);

    let stdout = child.stdout.take().expect("stdout");
    let lines: Vec<String> = BufReader::new(stdout)
        .lines()
        .collect::<Result<_, _>>()
        .expect("read lines");
    let status = child.wait().expect("wait");
    assert!(status.success());
    assert_eq!(lines.len(), 1);

    let error: Value = serde_json::from_str(&lines[0]).expect("json");
    assert_eq!(error["id"], 10);
    assert_eq!(error["error"]["code"], 1011);
    assert!(
        error["error"]["message"]
            .as_str()
            .expect("message")
            .contains("missing tool name")
    );
}
