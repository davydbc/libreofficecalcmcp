use crate::common::errors::AppError;
use crate::ods::content_xml::ContentXml;
use crate::ods::ods_templates::OdsTemplates;
use crate::ods::sheet_model::Workbook;
use std::fs::File;
use std::io::{Cursor, Read, Write};
use std::path::Path;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

pub struct OdsFile;

impl OdsFile {
    // Creates a minimal but valid ODS zip package from templates.
    pub fn create(path: &Path, initial_sheet_name: String) -> Result<(), AppError> {
        // Build new files from a LibreOffice-generated template so structure
        // matches what Calc expects by default.
        let template_bytes = OdsTemplates::empty_calc_template();
        let reader = Cursor::new(template_bytes);
        let mut template = ZipArchive::new(reader)?;

        let out_file = File::create(path)?;
        let mut writer = ZipWriter::new(out_file);
        let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
        let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

        // Keep ODS rule: first entry must be mimetype stored with no compression.
        let mut mimetype = String::new();
        template
            .by_name("mimetype")?
            .read_to_string(&mut mimetype)?;
        writer.start_file("mimetype", stored)?;
        writer.write_all(mimetype.as_bytes())?;

        let mut dir_names = Vec::new();
        let mut names = Vec::new();
        for i in 0..template.len() {
            let name = template.by_index(i)?.name().to_string();
            if name == "mimetype" {
                continue;
            }
            if name.ends_with('/') {
                dir_names.push(name);
                continue;
            }
            names.push(name);
        }
        dir_names.sort();
        names.sort();

        for dir in dir_names {
            let _ = writer.add_directory(dir, deflated);
        }

        for name in names {
            let mut entry = template.by_name(&name)?;
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes)?;

            if name == "content.xml" {
                let content = String::from_utf8(bytes)
                    .map_err(|e| AppError::InvalidOdsFormat(e.to_string()))?;
                let renamed =
                    ContentXml::rename_first_sheet_name_raw(&content, &initial_sheet_name)?;
                writer.start_file(name, deflated)?;
                writer.write_all(renamed.as_bytes())?;
            } else {
                writer.start_file(name, deflated)?;
                writer.write_all(&bytes)?;
            }
        }

        writer.finish()?;
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
        // Rebuild the zip preserving original entry order and directories.
        let src = File::open(path)?;
        let mut zip = ZipArchive::new(src)?;
        let mut entries: Vec<(String, bool, Vec<u8>)> = Vec::new();

        for i in 0..zip.len() {
            let mut file = zip.by_index(i)?;
            let name = file.name().to_string();
            if name.ends_with('/') {
                entries.push((name, true, Vec::new()));
                continue;
            }
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes)?;
            entries.push((name, false, bytes));
        }
        drop(zip);

        let out = File::create(path)?;
        let mut writer = ZipWriter::new(out);

        let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
        writer.start_file("mimetype", stored)?;
        let mimetype = entries
            .iter()
            .find(|(name, is_dir, _)| name == "mimetype" && !*is_dir)
            .map(|(_, _, data)| data.clone())
            .unwrap_or_else(|| OdsTemplates::mimetype().as_bytes().to_vec());
        writer.write_all(&mimetype)?;

        let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        for (name, is_dir, data) in entries {
            if name == "mimetype" {
                continue;
            }
            if is_dir {
                let _ = writer.add_directory(name, deflated);
                continue;
            }
            writer.start_file(name.clone(), deflated)?;
            if name == "content.xml" {
                writer.write_all(content_xml.as_bytes())?;
            } else {
                writer.write_all(&data)?;
            }
        }

        writer.finish()?;
        Ok(())
    }
}
