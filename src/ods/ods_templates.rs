use crate::common::errors::AppError;
use crate::ods::content_xml::ContentXml;
use crate::ods::manifest::Manifest;
use crate::ods::sheet_model::Workbook;

pub struct OdsTemplates;

impl OdsTemplates {
    // MIME string checked by spreadsheet apps before parsing XML.
    pub fn mimetype() -> &'static str {
        "application/vnd.oasis.opendocument.spreadsheet"
    }

    pub fn meta_xml() -> &'static str {
        r#"<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" office:version=\"1.2\">
  <office:meta />
</office:document-meta>"#
    }

    pub fn styles_xml() -> &'static str {
        r#"<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" office:version=\"1.2\">
  <office:styles />
</office:document-styles>"#
    }

    pub fn settings_xml() -> &'static str {
        r#"<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<office:document-settings xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" office:version=\"1.2\">
  <office:settings />
</office:document-settings>"#
    }

    pub fn manifest_xml() -> &'static str {
        Manifest::minimal_manifest_xml()
    }

    pub fn content_xml(initial_sheet_name: String) -> Result<String, AppError> {
        // Start every new document with one empty sheet.
        ContentXml::render(&Workbook::new(initial_sheet_name))
    }
}
