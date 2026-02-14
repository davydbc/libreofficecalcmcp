use serde_json::{json, Value};
use std::io::{BufRead, BufReader, Write};
use std::process::{Command, Stdio};

#[test]
fn stdio_server_handles_initialize_and_tools_list() {
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
    .expect("write init");
    writeln!(
        stdin,
        "{}",
        json!({"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}})
    )
    .expect("write list");
    drop(stdin);

    let stdout = child.stdout.take().expect("stdout");
    let lines: Vec<String> = BufReader::new(stdout)
        .lines()
        .collect::<Result<_, _>>()
        .expect("read lines");
    let status = child.wait().expect("wait");
    assert!(status.success());
    assert!(lines.len() >= 2);

    let init: Value = serde_json::from_str(&lines[0]).expect("init json");
    assert_eq!(init["id"], 1);
    assert_eq!(init["result"]["protocolVersion"], "2024-11-05");

    let list: Value = serde_json::from_str(&lines[1]).expect("list json");
    assert_eq!(list["id"], 2);
    assert!(list["result"]["tools"].is_array());
}

#[test]
fn stdio_server_returns_parse_error_for_invalid_json() {
    let exe = env!("CARGO_BIN_EXE_mcp-ods");
    let mut child = Command::new(exe)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .spawn()
        .expect("spawn mcp");

    let mut stdin = child.stdin.take().expect("stdin");
    writeln!(stdin, "{{ invalid json").expect("write invalid");
    drop(stdin);

    let stdout = child.stdout.take().expect("stdout");
    let lines: Vec<String> = BufReader::new(stdout)
        .lines()
        .collect::<Result<_, _>>()
        .expect("read lines");
    let status = child.wait().expect("wait");
    assert!(status.success());
    assert_eq!(lines.len(), 1);

    let parse_error: Value = serde_json::from_str(&lines[0]).expect("json");
    assert_eq!(parse_error["error"]["code"], -32700);
}

#[test]
fn stdio_server_ignores_notification_without_id() {
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
    .expect("write notification");
    writeln!(
        stdin,
        "{}",
        json!({"jsonrpc":"2.0","id":3,"method":"tools/list","params":{}})
    )
    .expect("write request");
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
    assert_eq!(list["id"], 3);
}

#[test]
fn stdio_server_returns_invalid_input_for_malformed_tools_call_payload() {
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
        json!({"jsonrpc":"2.0","id":11,"method":"tools/call","params":{}})
    )
    .expect("write request");
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
    assert_eq!(error["id"], 11);
    assert_eq!(error["error"]["code"], 1011);
}
