use mcp_ods::ods::ods_templates::OdsTemplates;

#[test]
fn ods_templates_expose_expected_static_assets() {
    assert!(OdsTemplates::empty_calc_template().len() > 100);
    assert_eq!(
        OdsTemplates::mimetype(),
        "application/vnd.oasis.opendocument.spreadsheet"
    );
    assert!(OdsTemplates::meta_xml().contains("document-meta"));
    assert!(OdsTemplates::styles_xml().contains("document-styles"));
    assert!(OdsTemplates::settings_xml().contains("document-settings"));
    assert!(OdsTemplates::manifest_xml().contains("manifest:file-entry"));
}

#[test]
fn ods_templates_content_xml_renders_initial_sheet() {
    let xml = OdsTemplates::content_xml("Inicial".to_string()).expect("content xml");
    assert!(xml.contains("table:name=\"Inicial\""));
    assert!(xml.contains("office:document-content"));
}
