use crate::common::dispatch;
use serde_json::json;

#[test]
fn create_ods_overwrite_flag_is_enforced() {
    let dir = tempfile::tempdir().expect("tempdir");
    let file_path = dir.path().join("overwrite.ods");

    dispatch(
        "create_ods",
        json!({
            "path": file_path.to_string_lossy(),
            "overwrite": false
        }),
    )
    .expect("first create");

    let err = dispatch(
        "create_ods",
        json!({
            "path": file_path.to_string_lossy(),
            "overwrite": false
        }),
    )
    .expect_err("should fail");

    assert!(err.to_string().contains("already exists"));
}

#[test]
fn create_ods_rejects_non_ods_extension() {
    let dir = tempfile::tempdir().expect("tempdir");
    let file_path = dir.path().join("invalid.txt");

    let err = dispatch(
        "create_ods",
        json!({
            "path": file_path.to_string_lossy(),
            "overwrite": true
        }),
    )
    .expect_err("must fail");

    assert!(err.to_string().contains("invalid path"));
}
