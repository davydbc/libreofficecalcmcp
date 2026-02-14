use mcp_ods::common::errors::AppError;
use mcp_ods::mcp::dispatcher::Dispatcher;
use serde_json::{json, Value};
use std::path::PathBuf;
use tempfile::TempDir;

pub fn new_ods_path(filename: &str) -> (TempDir, PathBuf) {
    let dir = tempfile::tempdir().expect("tempdir");
    let file_path = dir.path().join(filename);
    (dir, file_path)
}

pub fn dispatch(method: &str, params: Value) -> Result<Value, AppError> {
    Dispatcher::dispatch(method, Some(params))
}

pub fn create_base_ods(path: &PathBuf, sheet_name: &str) {
    dispatch(
        "create_ods",
        json!({
            "path": path.to_string_lossy(),
            "overwrite": true,
            "initial_sheet_name": sheet_name
        }),
    )
    .expect("create_ods");
}
