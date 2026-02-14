use crate::common::errors::AppError;
use crate::ods::content_xml::ContentXml;
use crate::ods::ods_templates::OdsTemplates;
use crate::ods::sheet_model::Workbook;
use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

pub struct OdsFile;

impl OdsFile {
    // Creates a minimal but valid ODS zip package from templates.
    pub fn create(path: &Path, initial_sheet_name: String) -> Result<(), AppError> {
        let file = File::create(path)?;
        let mut zip = ZipWriter::new(file);

        // ODS requires "mimetype" to be first and stored (not compressed).
        let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
        zip.start_file("mimetype", stored)?;
        zip.write_all(OdsTemplates::mimetype().as_bytes())?;

        let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

        zip.start_file("content.xml", deflated)?;
        zip.write_all(OdsTemplates::content_xml(initial_sheet_name)?.as_bytes())?;

        zip.start_file("styles.xml", deflated)?;
        zip.write_all(OdsTemplates::styles_xml().as_bytes())?;

        zip.start_file("meta.xml", deflated)?;
        zip.write_all(OdsTemplates::meta_xml().as_bytes())?;

        zip.start_file("settings.xml", deflated)?;
        zip.write_all(OdsTemplates::settings_xml().as_bytes())?;

        zip.add_directory("META-INF/", deflated)?;
        zip.start_file("META-INF/manifest.xml", deflated)?;
        zip.write_all(OdsTemplates::manifest_xml().as_bytes())?;

        zip.finish()?;
        Ok(())
    }

    pub fn read_workbook(path: &Path) -> Result<Workbook, AppError> {
        let content = Self::read_content_xml(path)?;
        ContentXml::parse(&content)
    }

    pub fn read_content_xml(path: &Path) -> Result<String, AppError> {
        let file = File::open(path)?;
        let mut zip = ZipArchive::new(file)?;

        let mut mimetype = String::new();
        zip.by_name("mimetype")?.read_to_string(&mut mimetype)?;
        if mimetype.trim() != OdsTemplates::mimetype() {
            return Err(AppError::InvalidOdsFormat("invalid mimetype".to_string()));
        }

        let mut content = String::new();
        zip.by_name("content.xml")?.read_to_string(&mut content)?;
        Ok(content)
    }

    pub fn write_workbook(path: &Path, workbook: &Workbook) -> Result<(), AppError> {
        // Use deterministic renderer to avoid namespace-prefix rewrites
        // that can happen with generic XML tree serializers.
        let new_content = ContentXml::render(workbook)?;
        Self::write_content_xml(path, &new_content)
    }

    pub fn write_content_xml(path: &Path, content_xml: &str) -> Result<(), AppError> {
        // Rebuild the zip to preserve non-content entries and replace only content.xml.
        let src = File::open(path)?;
        let mut zip = ZipArchive::new(src)?;
        let mut entries: HashMap<String, Vec<u8>> = HashMap::new();

        for i in 0..zip.len() {
            let mut file = zip.by_index(i)?;
            let name = file.name().to_string();
            if name.ends_with('/') {
                continue;
            }
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes)?;
            entries.insert(name, bytes);
        }
        drop(zip);

        entries.insert("content.xml".to_string(), content_xml.as_bytes().to_vec());

        let out = File::create(path)?;
        let mut writer = ZipWriter::new(out);

        let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
        writer.start_file("mimetype", stored)?;
        let mimetype = entries
            .get("mimetype")
            .cloned()
            .unwrap_or_else(|| OdsTemplates::mimetype().as_bytes().to_vec());
        writer.write_all(&mimetype)?;

        let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        let mut names: Vec<_> = entries
            .keys()
            .filter(|n| n.as_str() != "mimetype")
            .cloned()
            .collect();
        names.sort();

        for name in names {
            if let Some(content) = entries.get(&name) {
                // ZIP files do not require explicit directory entries.
                // Writing only file entries avoids duplicate/ambiguous directories
                // when re-saving ODS files that contain many nested paths.
                writer.start_file(name, deflated)?;
                writer.write_all(content)?;
            }
        }

        writer.finish()?;
        Ok(())
    }
}
