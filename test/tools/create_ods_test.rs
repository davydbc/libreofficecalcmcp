use mcp_ods::tools::create_ods;
use serde_json::json;
use std::fs::File;
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
