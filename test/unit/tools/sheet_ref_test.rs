use mcp_ods::tools::sheet_ref::SheetRef;
use serde::Deserialize;
use serde_json::json;

#[derive(Debug, Deserialize)]
struct SheetInput {
    sheet: SheetRef,
}

#[test]
fn sheet_ref_parses_object_by_name() {
    let parsed: SheetInput =
        serde_json::from_value(json!({ "sheet": { "name": "Hoja1" } })).expect("parse name");
    assert_eq!(
        parsed.sheet,
        SheetRef::Name {
            name: "Hoja1".to_string()
        }
    );
}

#[test]
fn sheet_ref_parses_object_by_index_number() {
    let parsed: SheetInput =
        serde_json::from_value(json!({ "sheet": { "index": 0 } })).expect("parse index");
    assert_eq!(parsed.sheet, SheetRef::Index { index: 0 });
}

#[test]
fn sheet_ref_parses_object_by_index_string() {
    let parsed: SheetInput =
        serde_json::from_value(json!({ "sheet": { "index": "0" } })).expect("parse index string");
    assert_eq!(parsed.sheet, SheetRef::Index { index: 0 });
}

#[test]
fn sheet_ref_parses_plain_sheet_name_string() {
    let parsed: SheetInput =
        serde_json::from_value(json!({ "sheet": "Hoja1" })).expect("parse plain sheet name");
    assert_eq!(
        parsed.sheet,
        SheetRef::Name {
            name: "Hoja1".to_string()
        }
    );
}

#[test]
fn sheet_ref_parses_json_encoded_sheet_object_string() {
    let parsed: SheetInput = serde_json::from_value(json!({
        "sheet": "{\"name\":\"Hoja1\"}"
    }))
    .expect("parse json encoded sheet object");
    assert_eq!(
        parsed.sheet,
        SheetRef::Name {
            name: "Hoja1".to_string()
        }
    );
}

#[test]
fn sheet_ref_rejects_invalid_shape() {
    let err = serde_json::from_value::<SheetInput>(json!({ "sheet": { "foo": "bar" } }))
        .expect_err("invalid shape should fail");
    assert!(err.to_string().contains("either name or index"));
}

#[test]
fn sheet_ref_rejects_invalid_index_string() {
    let err = serde_json::from_value::<SheetInput>(json!({ "sheet": { "index": "abc" } }))
        .expect_err("invalid index");
    assert!(err.to_string().contains("non-negative integer"));
}

#[test]
fn sheet_ref_rejects_invalid_json_encoded_sheet_object_string() {
    let err = serde_json::from_value::<SheetInput>(json!({
        "sheet": "{\"name\":"
    }))
    .expect_err("invalid json string");
    assert!(err.to_string().contains("not a valid object"));
}
