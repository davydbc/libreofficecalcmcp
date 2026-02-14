use mcp_ods::tools::create_ods;
use serde_json::json;
use std::fs::{self, File};
use tempfile::tempdir;
use zip::ZipArchive;

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

    let file = File::open(&path).expect("open ods");
    let mut zip = ZipArchive::new(file).expect("zip");
    assert!(zip.by_name("manifest.rdf").is_ok());
    assert!(zip.by_name("Thumbnails/thumbnail.png").is_ok());
    assert!(zip.by_name("META-INF/manifest.xml").is_ok());
}

#[test]
fn create_ods_respects_overwrite_flag() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("create_no_overwrite.ods");
    fs::write(&path, "dummy").expect("seed");

    let err = create_ods::handle(json!({
        "path": path.to_string_lossy(),
        "overwrite": false
    }))
    .expect_err("must fail");
    assert!(err.to_string().contains("already exists"));
}

#[test]
fn create_ods_rejects_non_ods_extension() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("create_invalid.xlsx");

    let err = create_ods::handle(json!({
        "path": path.to_string_lossy(),
        "overwrite": true
    }))
    .expect_err("must fail");
    assert!(err.to_string().contains("expected .ods extension"));
}

#[test]
fn create_ods_creates_parent_directories_when_missing() {
    let dir = tempdir().expect("tempdir");
    let nested = dir.path().join("a").join("b").join("nested_create.ods");

    create_ods::handle(json!({
        "path": nested.to_string_lossy(),
        "overwrite": true
    }))
    .expect("create nested");

    assert!(nested.exists());
}
