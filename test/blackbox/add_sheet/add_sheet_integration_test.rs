use crate::common::{create_base_ods, dispatch, new_ods_path};
use serde_json::json;
use std::fs::File;
use std::io::{Read, Write};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn overwrite_content_xml(file_path: &std::path::Path, new_content_xml: &str) {
    let src = File::open(file_path).expect("open source");
    let mut zip = ZipArchive::new(src).expect("zip read");
    let mut entries: Vec<(String, bool, Vec<u8>)> = Vec::new();
    for i in 0..zip.len() {
        let mut file = zip.by_index(i).expect("entry");
        let name = file.name().to_string();
        if name.ends_with('/') {
            entries.push((name, true, Vec::new()));
            continue;
        }
        let mut content = Vec::new();
        file.read_to_end(&mut content).expect("read entry");
        entries.push((name, false, content));
    }
    drop(zip);

    let out = File::create(file_path).expect("rewrite");
    let mut writer = ZipWriter::new(out);
    let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

    for (name, is_dir, content) in entries {
        if name == "mimetype" {
            writer.start_file(name, stored).expect("mimetype");
            writer.write_all(&content).expect("write mimetype");
            continue;
        }
        if is_dir {
            let _ = writer.add_directory(name, deflated);
            continue;
        }
        writer.start_file(name.clone(), deflated).expect("entry");
        if name == "content.xml" {
            writer
                .write_all(new_content_xml.as_bytes())
                .expect("write content.xml");
        } else {
            writer.write_all(&content).expect("write entry");
        }
    }
    writer.finish().expect("finish");
}

#[test]
fn add_sheet_inserts_new_sheet_at_end() {
    let (_dir, file_path) = new_ods_path("add_sheet.ods");
    create_base_ods(&file_path, "Hoja1");

    let out = dispatch(
        "add_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet_name": "NuevaHoja",
            "position": "end"
        }),
    )
    .expect("add_sheet");

    assert_eq!(out["sheets"], json!(["Hoja1", "NuevaHoja"]));
}

#[test]
fn add_sheet_can_insert_at_start() {
    let (_dir, file_path) = new_ods_path("add_sheet_start.ods");
    create_base_ods(&file_path, "Hoja1");

    let out = dispatch(
        "add_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet_name": "Inicio",
            "position": "start"
        }),
    )
    .expect("add_sheet");

    assert_eq!(out["sheets"], json!(["Inicio", "Hoja1"]));
}

#[test]
fn add_sheet_rejects_duplicate_name() {
    let (_dir, file_path) = new_ods_path("add_sheet_duplicate.ods");
    create_base_ods(&file_path, "Hoja1");

    let err = dispatch(
        "add_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet_name": "Hoja1",
            "position": "end"
        }),
    )
    .expect_err("duplicate");
    assert!(err.to_string().contains("already exists"));
}

#[test]
fn add_sheet_returns_file_not_found_for_missing_path() {
    let dir = tempfile::tempdir().expect("tempdir");
    let file_path = dir.path().join("missing.ods");
    let err = dispatch(
        "add_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet_name": "Nueva",
            "position": "end"
        }),
    )
    .expect_err("missing");
    assert!(err.to_string().contains("file not found"));
}

#[test]
fn add_sheet_preserves_existing_automatic_styles_block() {
    let (_dir, file_path) = new_ods_path("add_sheet_preserve_blocks.ods");
    create_base_ods(&file_path, "Hoja1");

    let custom = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:automatic-styles>
    <style:style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="ceX" style:family="table-cell"/>
  </office:automatic-styles>
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1"/>
  </office:spreadsheet></office:body>
</office:document-content>"#;
    overwrite_content_xml(&file_path, custom);

    dispatch(
        "add_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet_name": "Nueva",
            "position": "end"
        }),
    )
    .expect("add");

    let file = File::open(&file_path).expect("open");
    let mut zip = ZipArchive::new(file).expect("zip");
    let mut xml = String::new();
    zip.by_name("content.xml")
        .expect("content")
        .read_to_string(&mut xml)
        .expect("read");

    assert!(xml.contains("<office:automatic-styles>"));
    assert!(xml.contains("style:name=\"ceX\""));
    assert!(xml.contains("table:name=\"Nueva\""));
}
