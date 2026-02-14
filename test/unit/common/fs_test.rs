use mcp_ods::common::errors::AppError;
use mcp_ods::common::fs::FsUtil;

#[test]
fn resolve_ods_path_requires_ods_extension() {
    let err = FsUtil::resolve_ods_path("demo.xlsx").expect_err("should fail");
    assert!(matches!(err, AppError::InvalidPath(_)));
}

#[test]
fn resolve_ods_path_accepts_ods_extension() {
    let path = FsUtil::resolve_ods_path("demo.ods").expect("should work");
    assert_eq!(path.extension().and_then(|e| e.to_str()), Some("ods"));
}

#[test]
fn resolve_ods_path_rejects_empty_path() {
    let err = FsUtil::resolve_ods_path("   ").expect_err("should fail");
    assert!(matches!(err, AppError::InvalidPath(_)));
}
