use mcp_ods::common::errors::AppError;
use mcp_ods::ods::ods_file::OdsFile;
use mcp_ods::ods::sheet_model::{CellValue, Workbook};
use std::fs::File;
use std::io::{Read, Write};
use tempfile::tempdir;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

#[test]
fn ods_file_create_and_read_workbook_roundtrip() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("roundtrip.ods");

    OdsFile::create(&path, "Datos".to_string()).expect("create");
    let workbook = OdsFile::read_workbook(&path).expect("read");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Datos");
}

#[test]
fn ods_file_read_content_xml_rejects_invalid_mimetype() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("bad_mimetype.ods");
    write_ods_like_zip(
        &path,
        vec![
            ("mimetype".to_string(), false, b"text/plain".to_vec()),
            (
                "content.xml".to_string(),
                false,
                br#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#.to_vec(),
            ),
        ],
    );

    let err = OdsFile::read_content_xml(&path).expect_err("invalid mimetype");
    assert!(matches!(err, AppError::InvalidOdsFormat(_)));
}

#[test]
fn ods_file_write_content_xml_updates_entry_and_keeps_other_files() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("write_content.ods");
    OdsFile::create(&path, "Hoja1".to_string()).expect("create");

    let new_content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
  <office:body><office:spreadsheet><table:table table:name="Nueva"/></office:spreadsheet></office:body>
</office:document-content>"#;
    OdsFile::write_content_xml(&path, new_content).expect("write content");

    let read_back = OdsFile::read_content_xml(&path).expect("read content");
    assert!(read_back.contains("table:name=\"Nueva\""));

    let mut zip = ZipArchive::new(File::open(&path).expect("open")).expect("zip");
    assert!(zip.by_name("styles.xml").is_ok());
}

#[test]
fn ods_file_write_content_xml_adds_default_mimetype_when_missing() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("missing_mimetype.ods");
    write_ods_like_zip(
        &path,
        vec![(
            "content.xml".to_string(),
            false,
            br#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#.to_vec(),
        )],
    );

    OdsFile::write_content_xml(
        &path,
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#,
    )
    .expect("write");

    let mut zip = ZipArchive::new(File::open(&path).expect("open")).expect("zip");
    let mut mimetype = String::new();
    zip.by_name("mimetype")
        .expect("mimetype exists")
        .read_to_string(&mut mimetype)
        .expect("read mimetype");
    assert_eq!(mimetype, "application/vnd.oasis.opendocument.spreadsheet");
}

#[test]
fn ods_file_write_workbook_persists_new_values() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("write_workbook.ods");
    OdsFile::create(&path, "Hoja1".to_string()).expect("create");

    let mut workbook = Workbook::new("Hoja1".to_string());
    workbook.sheets[0].ensure_cell_mut(0, 0).value = CellValue::String("A1".to_string());
    workbook.sheets[0].ensure_cell_mut(1, 0).value = CellValue::Boolean(true);
    OdsFile::write_workbook(&path, &workbook).expect("write workbook");

    let loaded = OdsFile::read_workbook(&path).expect("read workbook");
    assert_eq!(
        loaded.sheets[0].get_cell(0, 0).expect("a1").value,
        CellValue::String("A1".to_string())
    );
    assert_eq!(
        loaded.sheets[0].get_cell(1, 0).expect("a2").value,
        CellValue::Boolean(true)
    );
}

fn write_ods_like_zip(path: &std::path::Path, entries: Vec<(String, bool, Vec<u8>)>) {
    let out = File::create(path).expect("create zip");
    let mut writer = ZipWriter::new(out);
    let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

    for (name, is_dir, data) in entries {
        if is_dir {
            writer.add_directory(name, deflated).expect("add dir");
            continue;
        }
        if name == "mimetype" {
            writer.start_file(name, stored).expect("start mimetype");
        } else {
            writer.start_file(name, deflated).expect("start file");
        }
        writer.write_all(&data).expect("write file");
    }

    writer.finish().expect("finish");
}
