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

#[test]
fn sheet_ref_rejects_non_string_name() {
    let err = serde_json::from_value::<SheetInput>(json!({ "sheet": { "name": 7 } }))
        .expect_err("name must fail");
    assert!(err.to_string().contains("sheet.name must be a string"));
}

#[test]
fn sheet_ref_rejects_non_integer_number_index() {
    let err = serde_json::from_value::<SheetInput>(json!({ "sheet": { "index": 1.5 } }))
        .expect_err("float index");
    assert!(err.to_string().contains("non-negative integer"));
}

#[test]
fn sheet_ref_rejects_non_object_non_string_non_number_value() {
    let err = serde_json::from_value::<SheetInput>(json!({ "sheet": true }))
        .expect_err("bool should fail");
    assert!(err.to_string().contains("object or string"));
}

#[test]
fn sheet_ref_rejects_invalid_index_value_type() {
    let err = serde_json::from_value::<SheetInput>(json!({ "sheet": { "index": [] } }))
        .expect_err("array index should fail");
    assert!(err.to_string().contains("non-negative integer"));
}

#[test]
fn sheet_ref_accepts_root_numeric_index() {
    let parsed: SheetInput = serde_json::from_value(json!({ "sheet": 2 })).expect("parse");
    assert_eq!(parsed.sheet, SheetRef::Index { index: 2 });
}

#[cfg(target_pointer_width = "32")]
#[test]
fn sheet_ref_rejects_too_large_numeric_index() {
    let err = serde_json::from_value::<SheetInput>(json!({
        "sheet": { "index": 4294967296u64 }
    }))
    .expect_err("too large");
    assert!(err.to_string().contains("too large"));
}
