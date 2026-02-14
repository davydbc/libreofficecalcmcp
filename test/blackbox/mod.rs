use mcp_ods::mcp::dispatcher::Dispatcher;
use serde_json::json;
use tempfile::tempdir;

#[test]
fn dispatch_unknown_tool_returns_error() {
    let result = Dispatcher::dispatch("tool_que_no_existe", Some(json!({})));
    assert!(result.is_err());
}

#[test]
fn create_set_and_get_sheet_content() {
    let dir = tempdir().expect("tempdir");
    let file_path = dir.path().join("demo.ods");

    let create = Dispatcher::dispatch(
        "create_ods",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "overwrite": false,
            "initial_sheet_name": "Hoja1"
        })),
    )
    .expect("create_ods");
    assert_eq!(create["sheets"][0], "Hoja1");

    Dispatcher::dispatch(
        "set_cell_value",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "Hoja1" },
            "cell": "B2",
            "value": { "type": "string", "data": "hola" }
        })),
    )
    .expect("set_cell_value");

    let content = Dispatcher::dispatch(
        "get_sheet_content",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "Hoja1" },
            "mode": "matrix",
            "max_rows": 10,
            "max_cols": 10,
            "include_empty_trailing": false
        })),
    )
    .expect("get_sheet_content");

    assert_eq!(content["rows"], 2);
    assert_eq!(content["cols"], 2);
    assert_eq!(content["data"][1][1], "hola");
}

#[test]
fn create_ods_overwrite_flag_is_enforced() {
    let dir = tempdir().expect("tempdir");
    let file_path = dir.path().join("overwrite.ods");
    let path = file_path.to_string_lossy().to_string();

    Dispatcher::dispatch(
        "create_ods",
        Some(json!({
            "path": path,
            "overwrite": false
        })),
    )
    .expect("first create");

    let err = Dispatcher::dispatch(
        "create_ods",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "overwrite": false
        })),
    )
    .expect_err("should fail");
    assert!(err.to_string().contains("already exists"));
}

#[test]
fn get_sheets_add_sheet_and_duplicate_sheet_workflow() {
    let dir = tempdir().expect("tempdir");
    let file_path = dir.path().join("sheets.ods");
    let path = file_path.to_string_lossy().to_string();

    Dispatcher::dispatch(
        "create_ods",
        Some(json!({
            "path": path,
            "overwrite": true,
            "initial_sheet_name": "Base"
        })),
    )
    .expect("create");

    Dispatcher::dispatch(
        "add_sheet",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet_name": "Datos",
            "position": "end"
        })),
    )
    .expect("add sheet");

    Dispatcher::dispatch(
        "duplicate_sheet",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "source_sheet": { "name": "Datos" },
            "new_sheet_name": "Datos (copia)"
        })),
    )
    .expect("duplicate");

    let sheets = Dispatcher::dispatch(
        "get_sheets",
        Some(json!({
            "path": file_path.to_string_lossy()
        })),
    )
    .expect("get_sheets");

    assert_eq!(sheets["sheets"], json!(["Base", "Datos", "Datos (copia)"]));
}

#[test]
fn get_cell_value_supports_string_number_boolean_and_empty() {
    let dir = tempdir().expect("tempdir");
    let file_path = dir.path().join("values.ods");

    Dispatcher::dispatch(
        "create_ods",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "overwrite": true
        })),
    )
    .expect("create");

    Dispatcher::dispatch(
        "set_cell_value",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "A1",
            "value": { "type": "string", "data": "txt" }
        })),
    )
    .expect("set string");

    Dispatcher::dispatch(
        "set_cell_value",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "B1",
            "value": { "type": "number", "data": 3.5 }
        })),
    )
    .expect("set number");

    Dispatcher::dispatch(
        "set_cell_value",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "C1",
            "value": { "type": "boolean", "data": true }
        })),
    )
    .expect("set bool");

    let a1 = Dispatcher::dispatch(
        "get_cell_value",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "A1"
        })),
    )
    .expect("get a1");
    assert_eq!(a1["value"], json!({"type":"string","data":"txt"}));

    let b1 = Dispatcher::dispatch(
        "get_cell_value",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "B1"
        })),
    )
    .expect("get b1");
    assert_eq!(b1["value"], json!({"type":"number","data":3.5}));

    let c1 = Dispatcher::dispatch(
        "get_cell_value",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "C1"
        })),
    )
    .expect("get c1");
    assert_eq!(c1["value"], json!({"type":"boolean","data":true}));

    let z9 = Dispatcher::dispatch(
        "get_cell_value",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "Z9"
        })),
    )
    .expect("get empty");
    assert_eq!(z9["value"], json!({"type":"empty"}));
}

#[test]
fn set_range_values_writes_matrix_from_start_cell() {
    let dir = tempdir().expect("tempdir");
    let file_path = dir.path().join("range.ods");

    Dispatcher::dispatch(
        "create_ods",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "overwrite": true
        })),
    )
    .expect("create");

    let result = Dispatcher::dispatch(
        "set_range_values",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "start_cell": "B2",
            "data": [
                ["A","B"],
                ["1","2"]
            ]
        })),
    )
    .expect("set range");
    assert_eq!(result["rows_written"], 2);
    assert_eq!(result["cols_written"], 2);

    let content = Dispatcher::dispatch(
        "get_sheet_content",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "mode": "matrix",
            "max_rows": 10,
            "max_cols": 10,
            "include_empty_trailing": false
        })),
    )
    .expect("content");

    assert_eq!(content["rows"], 3);
    assert_eq!(content["cols"], 3);
    assert_eq!(content["data"][1][1], "A");
    assert_eq!(content["data"][2][2], "2");
}

#[test]
fn get_sheet_content_supports_trailing_trim_and_limits() {
    let dir = tempdir().expect("tempdir");
    let file_path = dir.path().join("trim.ods");

    Dispatcher::dispatch(
        "create_ods",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "overwrite": true
        })),
    )
    .expect("create");

    Dispatcher::dispatch(
        "set_range_values",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "start_cell": "A1",
            "data": [
                ["x","",""],
                ["","",""]
            ]
        })),
    )
    .expect("set range");

    let trimmed = Dispatcher::dispatch(
        "get_sheet_content",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "mode": "matrix",
            "max_rows": 10,
            "max_cols": 10,
            "include_empty_trailing": false
        })),
    )
    .expect("trimmed");
    assert_eq!(trimmed["rows"], 1);
    assert_eq!(trimmed["cols"], 1);

    let with_trailing = Dispatcher::dispatch(
        "get_sheet_content",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "mode": "matrix",
            "max_rows": 1,
            "max_cols": 2,
            "include_empty_trailing": true
        })),
    )
    .expect("with trailing");
    assert_eq!(with_trailing["rows"], 1);
    assert_eq!(with_trailing["cols"], 2);
}

#[test]
fn tools_return_clear_errors_for_invalid_inputs() {
    let dir = tempdir().expect("tempdir");
    let file_path = dir.path().join("errors.ods");

    Dispatcher::dispatch(
        "create_ods",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "overwrite": true
        })),
    )
    .expect("create");

    let bad_mode = Dispatcher::dispatch(
        "get_sheet_content",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "mode": "records"
        })),
    )
    .expect_err("invalid mode");
    assert!(bad_mode.to_string().contains("mode=matrix"));

    let bad_sheet = Dispatcher::dispatch(
        "get_cell_value",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "name": "NoExiste" },
            "cell": "A1"
        })),
    )
    .expect_err("invalid sheet");
    assert!(bad_sheet.to_string().contains("sheet not found"));

    let bad_cell = Dispatcher::dispatch(
        "set_cell_value",
        Some(json!({
            "path": file_path.to_string_lossy(),
            "sheet": { "index": 0 },
            "cell": "1A",
            "value": { "type": "string", "data": "x" }
        })),
    )
    .expect_err("invalid cell");
    assert!(bad_cell.to_string().contains("invalid cell address"));
}
