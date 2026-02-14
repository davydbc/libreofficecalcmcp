use crate::common::dispatch;
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
