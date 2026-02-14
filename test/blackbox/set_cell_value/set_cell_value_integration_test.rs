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
fn set_cell_value_updates_and_persists_value() {
    let (_dir, file_path) = new_ods_path("set_cell.ods");
    create_base_ods(&file_path, "Hoja1");

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "Hoja1" },
            "cell": "B2",
            "value": { "type": "string", "data": "hola" }
        }),
    )
    .expect("set_cell_value");

    let value = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "Hoja1" },
            "cell": "B2"
        }),
    )
    .expect("get_cell_value");

    assert_eq!(value["value"], json!({"type":"string","data":"hola"}));
}

#[test]
fn set_cell_value_preserves_directory_entries_from_source_zip() {
    let (_dir, file_path) = new_ods_path("set_cell_nested.ods");
    create_base_ods(&file_path, "Hoja1");

    // Simulate an ODS enriched by external tools with nested file entries.
    {
        let src = File::open(&file_path).expect("open source");
        let mut zip = ZipArchive::new(src).expect("zip read");
        let mut entries: Vec<(String, Vec<u8>)> = Vec::new();
        for i in 0..zip.len() {
            let mut file = zip.by_index(i).expect("entry");
            if file.name().ends_with('/') {
                continue;
            }
            let mut content = Vec::new();
            file.read_to_end(&mut content).expect("read entry");
            entries.push((file.name().to_string(), content));
        }
        drop(zip);

        let out = File::create(&file_path).expect("rewrite");
        let mut writer = ZipWriter::new(out);
        let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
        let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        for (name, content) in entries {
            if name == "mimetype" {
                writer.start_file(name, stored).expect("mimetype");
            } else {
                writer.start_file(name, deflated).expect("entry");
            }
            writer.write_all(&content).expect("write");
        }
        writer
            .start_file("Configurations2/statusbar/state.xml", deflated)
            .expect("nested file");
        writer
            .write_all(br#"<?xml version="1.0" encoding="UTF-8"?><state/>"#)
            .expect("write nested");
        writer.finish().expect("finish");
    }

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "Hoja1" },
            "cell": "A1",
            "value": { "type": "string", "data": "ok" }
        }),
    )
    .expect("set_cell_value");

    // Output zip should preserve nested file entries from source zip.
    let file = File::open(&file_path).expect("open result");
    let mut zip = ZipArchive::new(file).expect("read result zip");
    assert!(zip.by_name("Configurations2/statusbar/state.xml").is_ok());
}

#[test]
fn set_cell_value_rejects_repeated_non_empty_cell_in_integration_flow() {
    let (_dir, file_path) = new_ods_path("set_cell_repeated_non_empty.ods");
    create_base_ods(&file_path, "Hoja1");

    let content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row>
        <table:table-cell table:number-columns-repeated="3" office:value-type="string"><text:p>v</text:p></table:table-cell>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;
    overwrite_content_xml(&file_path, content);

    let err = dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "B1",
            "value": { "type": "string", "data": "x" }
        }),
    )
    .expect_err("must fail");
    assert!(err.to_string().contains("cannot safely edit repeated non-empty cell"));
}

#[test]
fn set_cell_value_updates_second_sheet_only() {
    let (_dir, file_path) = new_ods_path("set_cell_second_sheet.ods");
    create_base_ods(&file_path, "S1");

    dispatch(
        "add_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet_name": "S2",
            "position": "end"
        }),
    )
    .expect("add sheet");

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 1 },
            "cell": "A1",
            "value": { "type": "string", "data": "en_s2" }
        }),
    )
    .expect("set");

    let s1 = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "A1"
        }),
    )
    .expect("get s1");
    let s2 = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 1 },
            "cell": "A1"
        }),
    )
    .expect("get s2");

    assert_eq!(s1["value"], json!({"type":"empty"}));
    assert_eq!(s2["value"], json!({"type":"string","data":"en_s2"}));
}

#[test]
fn set_cell_value_inside_merged_range_writes_anchor_cell() {
    let (_dir, file_path) = new_ods_path("set_cell_merged_anchor.ods");
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
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "A6",
            "value": { "type": "string", "data": "merged_value" }
        }),
    )
    .expect("set");

    let anchor = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "A5"
        }),
    )
    .expect("get anchor");
    assert_eq!(anchor["value"], json!({"type":"string","data":"merged_value"}));
}

#[test]
fn set_cell_value_supports_number_boolean_and_empty() {
    let (_dir, file_path) = new_ods_path("set_cell_typed.ods");
    create_base_ods(&file_path, "Hoja1");

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "C1",
            "value": { "type": "number", "data": 9.25 }
        }),
    )
    .expect("set number");
    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "D1",
            "value": { "type": "boolean", "data": true }
        }),
    )
    .expect("set bool");
    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "D1",
            "value": { "type": "empty" }
        }),
    )
    .expect("set empty");

    let c1 = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "C1"
        }),
    )
    .expect("get c1");
    let d1 = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "D1"
        }),
    )
    .expect("get d1");

    assert_eq!(c1["value"], json!({"type":"number","data":9.25}));
    assert_eq!(d1["value"], json!({"type":"empty"}));
}

#[test]
fn set_cell_value_rejects_invalid_cell_address() {
    let (_dir, file_path) = new_ods_path("set_cell_bad_addr.ods");
    create_base_ods(&file_path, "Hoja1");

    let err = dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "22A",
            "value": { "type": "string", "data": "x" }
        }),
    )
    .expect_err("bad address");
    assert!(err.to_string().contains("invalid cell address"));
}

#[test]
fn set_cell_value_handles_self_closing_table_content() {
    let (_dir, file_path) = new_ods_path("set_cell_self_closing_table.ods");
    create_base_ods(&file_path, "Hoja1");

    let content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1"/>
  </office:spreadsheet></office:body>
</office:document-content>"#;
    overwrite_content_xml(&file_path, content);

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "B2",
            "value": { "type": "string", "data": "X" }
        }),
    )
    .expect("set");

    let b2 = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "B2"
        }),
    )
    .expect("get");
    assert_eq!(b2["value"], json!({"type":"string","data":"X"}));
}

#[test]
fn set_cell_value_splits_repeated_row_and_only_updates_target() {
    let (_dir, file_path) = new_ods_path("set_cell_repeated_row_split.ods");
    create_base_ods(&file_path, "Hoja1");

    let content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row table:number-rows-repeated="3">
        <table:table-cell table:number-columns-repeated="2"/>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;
    overwrite_content_xml(&file_path, content);

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "B2",
            "value": { "type": "string", "data": "MID" }
        }),
    )
    .expect("set");

    let b2 = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "B2"
        }),
    )
    .expect("b2");
    let b1 = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "B1"
        }),
    )
    .expect("b1");
    let b3 = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "B3"
        }),
    )
    .expect("b3");

    assert_eq!(b2["value"], json!({"type":"string","data":"MID"}));
    assert_eq!(b1["value"], json!({"type":"empty"}));
    assert_eq!(b3["value"], json!({"type":"empty"}));
}

#[test]
fn set_cell_value_replaces_non_repeated_non_empty_cell() {
    let (_dir, file_path) = new_ods_path("set_cell_non_repeated_non_empty.ods");
    create_base_ods(&file_path, "Hoja1");

    let content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row>
        <table:table-cell office:value-type="string"><text:p>old</text:p></table:table-cell>
        <table:table-cell/>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;
    overwrite_content_xml(&file_path, content);

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "A1",
            "value": { "type": "string", "data": "new" }
        }),
    )
    .expect("set");

    let a1 = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "A1"
        }),
    )
    .expect("a1");
    assert_eq!(a1["value"], json!({"type":"string","data":"new"}));
}

#[test]
fn set_cell_value_skips_covered_cells_and_writes_after_them() {
    let (_dir, file_path) = new_ods_path("set_cell_after_covered.ods");
    create_base_ods(&file_path, "Hoja1");

    let content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row>
        <table:covered-table-cell table:number-columns-repeated="2"/>
        <table:table-cell/>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;
    overwrite_content_xml(&file_path, content);

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "C1",
            "value": { "type": "string", "data": "after_cov" }
        }),
    )
    .expect("set");

    let content = dispatch(
        "get_sheet_content",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "mode": "matrix",
            "max_rows": 5,
            "max_cols": 8,
            "include_empty_trailing": true
        }),
    )
    .expect("content");
    let matrix = content["data"].as_array().expect("rows");
    let found = matrix
        .iter()
        .filter_map(|row| row.as_array())
        .flat_map(|row| row.iter())
        .any(|v| v.as_str() == Some("after_cov"));
    assert!(found, "expected written value in matrix");
}

#[test]
fn set_cell_value_writes_far_cell_when_row_exists_but_target_col_is_after_existing_data() {
    let (_dir, file_path) = new_ods_path("set_cell_far_col.ods");
    create_base_ods(&file_path, "Hoja1");

    let content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row>
        <table:table-cell office:value-type="string"><text:p>A</text:p></table:table-cell>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;
    overwrite_content_xml(&file_path, content);

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "D1",
            "value": { "type": "string", "data": "D1" }
        }),
    )
    .expect("set");

    let d1 = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "D1"
        }),
    )
    .expect("get");
    assert_eq!(d1["value"], json!({"type":"string","data":"D1"}));
}

#[test]
fn set_cell_value_can_create_missing_rows_before_target() {
    let (_dir, file_path) = new_ods_path("set_cell_missing_rows.ods");
    create_base_ods(&file_path, "Hoja1");

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "A5",
            "value": { "type": "string", "data": "A5" }
        }),
    )
    .expect("set");

    let a5 = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "A5"
        }),
    )
    .expect("a5");
    let a4 = dispatch(
        "get_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "A4"
        }),
    )
    .expect("a4");
    assert_eq!(a5["value"], json!({"type":"string","data":"A5"}));
    assert_eq!(a4["value"], json!({"type":"empty"}));
}

#[test]
fn set_cell_value_handles_repeated_row_with_covered_cells_before_target() {
    let (_dir, file_path) = new_ods_path("set_cell_repeated_covered_before_target.ods");
    create_base_ods(&file_path, "Hoja1");

    let content = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:spreadsheet>
    <table:table table:name="Hoja1">
      <table:table-row table:number-rows-repeated="2">
        <table:covered-table-cell table:number-columns-repeated="2"/>
        <table:table-cell/>
      </table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#;
    overwrite_content_xml(&file_path, content);

    dispatch(
        "set_cell_value",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "C1",
            "value": { "type": "string", "data": "C1" }
        }),
    )
    .expect("set");

    let content = dispatch(
        "get_sheet_content",
        json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "mode": "matrix",
            "max_rows": 8,
            "max_cols": 8,
            "include_empty_trailing": true
        }),
    )
    .expect("content");
    let matrix = content["data"].as_array().expect("rows");
    let found = matrix
        .iter()
        .filter_map(|row| row.as_array())
        .flat_map(|row| row.iter())
        .any(|v| v.as_str() == Some("C1"));
    assert!(found, "expected written value in matrix");
}
