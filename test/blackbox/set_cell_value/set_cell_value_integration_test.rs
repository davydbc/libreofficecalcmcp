use crate::common::{create_base_ods, dispatch, new_ods_path};
use serde_json::json;
use std::fs::File;
use std::io::{Read, Write};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

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
fn set_cell_value_keeps_zip_valid_without_directory_entries() {
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

    // Output zip should contain only file entries (no explicit directories).
    let file = File::open(&file_path).expect("open result");
    let mut zip = ZipArchive::new(file).expect("read result zip");
    for i in 0..zip.len() {
        let name = zip.by_index(i).expect("entry").name().to_string();
        assert!(
            !name.ends_with('/'),
            "unexpected directory entry in output zip: {name}"
        );
    }
}
