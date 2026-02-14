use crate::common::{create_base_ods, dispatch, new_ods_path};
use serde_json::json;
use std::fs::File;
use std::io::{Read, Write};
use xmltree::{Element, XMLNode};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

#[test]
fn duplicate_sheet_creates_named_copy() {
    let (_dir, file_path) = new_ods_path("duplicate.ods");
    create_base_ods(&file_path, "Datos");

    let out = dispatch(
        "duplicate_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "source_sheet": { "name": "Datos" },
            "new_sheet_name": "Datos (copia)"
        }),
    )
    .expect("duplicate_sheet");

    assert_eq!(out["sheets"], json!(["Datos", "Datos (copia)"]));
}

#[test]
fn duplicate_sheet_preserves_style_attributes() {
    let (_dir, file_path) = new_ods_path("duplicate_styles.ods");
    create_base_ods(&file_path, "Original");

    // Inject a styled content.xml similar to what LibreOffice can generate.
    let styled_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2">
  <office:body>
    <office:spreadsheet>
      <table:table table:name="Original">
        <table:table-row>
          <table:table-cell table:style-name="ce1" office:value-type="string">
            <text:p>valor</text:p>
          </table:table-cell>
        </table:table-row>
      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>"#;

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
            file.read_to_end(&mut content).expect("read");
            entries.push((file.name().to_string(), content));
        }
        drop(zip);

        let out = File::create(&file_path).expect("create out");
        let mut writer = ZipWriter::new(out);
        let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
        let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        for (name, mut content) in entries {
            if name == "content.xml" {
                content = styled_xml.as_bytes().to_vec();
            }
            if name == "mimetype" {
                writer.start_file(name, stored).expect("mimetype");
            } else {
                writer.start_file(name, deflated).expect("entry");
            }
            writer.write_all(&content).expect("write");
        }
        writer.finish().expect("finish");
    }

    dispatch(
        "duplicate_sheet",
        json!({
            "path": file_path.to_string_lossy(),
            "source_sheet": { "name": "Original" },
            "new_sheet_name": "Duplicado"
        }),
    )
    .expect("duplicate");

    let file = File::open(&file_path).expect("open result");
    let mut zip = ZipArchive::new(file).expect("zip result");
    let mut xml = String::new();
    zip.by_name("content.xml")
        .expect("content.xml")
        .read_to_string(&mut xml)
        .expect("read content");

    let root = Element::parse(xml.as_bytes()).expect("parse xml");
    let body = find_child_local(&root, "body").expect("body");
    let spreadsheet = find_child_local(body, "spreadsheet").expect("spreadsheet");

    let mut table_names = Vec::new();
    let mut style_hits = 0usize;
    for child in &spreadsheet.children {
        if let XMLNode::Element(table) = child {
            if local_name(&table.name) != "table" {
                continue;
            }
            for (k, v) in &table.attributes {
                if local_name(k) == "name" {
                    table_names.push(v.clone());
                }
            }
            for row_node in &table.children {
                if let XMLNode::Element(row) = row_node {
                    if local_name(&row.name) != "table-row" {
                        continue;
                    }
                    for cell_node in &row.children {
                        if let XMLNode::Element(cell) = cell_node {
                            if local_name(&cell.name) != "table-cell" {
                                continue;
                            }
                            for (k, v) in &cell.attributes {
                                if local_name(k) == "style-name" && v == "ce1" {
                                    style_hits += 1;
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    assert!(table_names.iter().any(|n| n == "Original"));
    assert!(table_names.iter().any(|n| n == "Duplicado"));
    assert_eq!(style_hits, 2);
}

fn find_child_local<'a>(element: &'a Element, target: &str) -> Option<&'a Element> {
    for child in &element.children {
        if let XMLNode::Element(e) = child {
            if local_name(&e.name) == target {
                return Some(e);
            }
        }
    }
    None
}

fn local_name(name: &str) -> &str {
    name.rsplit(':').next().unwrap_or(name)
}
