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
fn set_range_values_writes_matrix_from_start_cell() {
    let (_dir, file_path) = new_ods_path("range.ods");
    create_base_ods(&file_path, "Hoja1");

    let result = dispatch(
        "set_range_values",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "start_cell": "B2",
            "data": [["A", "B"], ["1", "2"]]
        }),
    )
    .expect("set range");

    assert_eq!(result["rows_written"], 2);
    assert_eq!(result["cols_written"], 2);

    let content = dispatch(
        "get_sheet_content",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "mode": "matrix",
            "max_rows": 10,
            "max_cols": 10,
            "include_empty_trailing": false
        }),
    )
    .expect("content");

    assert_eq!(content["rows"], 3);
    assert_eq!(content["cols"], 3);
    assert_eq!(content["data"][1][1], "A");
    assert_eq!(content["data"][2][2], "2");
}

#[test]
fn set_range_values_accepts_sheet_by_name() {
    let (_dir, file_path) = new_ods_path("range_by_name.ods");
    create_base_ods(&file_path, "S1");
    dispatch(
        "add_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet_name": "S2",
            "position": "end"
        }),
    )
    .expect("add");

    dispatch(
        "set_range_values",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "S2" },
            "start_cell": "A1",
            "data": [["x","y"]]
        }),
    )
    .expect("set range");

    let content = dispatch(
        "get_sheet_content",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "S2" },
            "mode": "matrix",
            "max_rows": 10,
            "max_cols": 10,
            "include_empty_trailing": false
        }),
    )
    .expect("content");

    assert_eq!(content["data"], json!([["x", "y"]]));
}

#[test]
fn set_range_values_rejects_invalid_sheet() {
    let (_dir, file_path) = new_ods_path("range_bad_sheet.ods");
    create_base_ods(&file_path, "Hoja1");

    let err = dispatch(
        "set_range_values",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 4 },
            "start_cell": "A1",
            "data": [["v"]]
        }),
    )
    .expect_err("bad sheet");
    assert!(err.to_string().contains("sheet not found"));
}

#[test]
fn set_range_values_rejects_invalid_start_cell() {
    let (_dir, file_path) = new_ods_path("range_bad_start.ods");
    create_base_ods(&file_path, "Hoja1");

    let err = dispatch(
        "set_range_values",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "start_cell": "11",
            "data": [["v"]]
        }),
    )
    .expect_err("bad start");
    assert!(err.to_string().contains("invalid cell address"));
}

#[test]
fn set_range_values_preserves_unrelated_existing_styles() {
    let (_dir, file_path) = new_ods_path("range_preserve_styles.ods");
    create_base_ods(&file_path, "Hoja1");

    let content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row>
        <table:table-cell table:style-name="ce1" office:value-type="string"><text:p>styled</text:p></table:table-cell>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;
    overwrite_content_xml(&file_path, content);

    dispatch(
        "set_range_values",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "start_cell": "C3",
            "data": [["x","y"]]
        }),
    )
    .expect("set range");

    let file = File::open(&file_path).expect("open");
    let mut zip = ZipArchive::new(file).expect("zip");
    let mut xml = String::new();
    zip.by_name("content.xml")
        .expect("content")
        .read_to_string(&mut xml)
        .expect("read");
    assert!(xml.contains("table:style-name=\"ce1\""));
}

#[test]
fn set_range_values_uses_merged_anchor_when_start_cell_is_covered() {
    let (_dir, file_path) = new_ods_path("range_merged_anchor.ods");
    create_base_ods(&file_path, "Hoja1");

    let content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row table:number-rows-repeated="4"/>
      <table:table-row>
        <table:table-cell table:number-rows-spanned="2" office:value-type="string"><text:p>base</text:p></table:table-cell>
      </table:table-row>
      <table:table-row>
        <table:covered-table-cell/>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;
    overwrite_content_xml(&file_path, content);

    dispatch(
        "set_range_values",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "start_cell": "A6",
            "data": [["merged"]]
        }),
    )
    .expect("set range");

    let anchor = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "A5"
        }),
    )
    .expect("anchor");
    assert_eq!(anchor["value"], json!({"type":"string","data":"merged"}));
}
