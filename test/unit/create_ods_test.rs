use mcp_ods::tools::create_ods;
use serde_json::json;
use tempfile::tempdir;

#[test]
fn create_ods_creates_file() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("create_unit.ods");

    let out = create_ods::handle(json!({
        "path": path.to_string_lossy(),
        "overwrite": false,
        "initial_sheet_name": "Hoja1"
    }))
    .expect("create");

    assert_eq!(out["sheets"][0], "Hoja1");
    assert!(path.exists());
}
