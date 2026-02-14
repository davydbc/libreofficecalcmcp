use mcp_ods::ods::manifest::Manifest;

#[test]
fn manifest_xml_contains_required_entries() {
    let xml = Manifest::minimal_manifest_xml();
    assert!(xml.contains("manifest:manifest"));
    assert!(xml.contains("manifest:full-path=\"/\""));
    assert!(xml.contains("manifest:full-path=\"content.xml\""));
    assert!(xml.contains("manifest:full-path=\"styles.xml\""));
    assert!(xml.contains("manifest:full-path=\"meta.xml\""));
    assert!(xml.contains("manifest:full-path=\"settings.xml\""));
}
