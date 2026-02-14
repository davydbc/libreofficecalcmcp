use crate::common::errors::AppError;
use crate::tools;
use serde_json::{json, Value};

pub struct Dispatcher;

impl Dispatcher {
    pub fn dispatch(method: &str, params: Option<Value>) -> Result<Value, AppError> {
        match method {
            // MCP initialization handshake used by clients like Cline.
            "initialize" => Ok(Self::initialize_result()),
            // Notification can arrive with or without namespace depending on the client.
            "initialized" | "notifications/initialized" => Ok(Value::Null),
            "tools/list" => Ok(Self::tools_list_result()),
            "tools/call" => Self::dispatch_tools_call(params),
            // Keep direct tool-name calls for manual CLI testing and existing tests.
            _ => Self::dispatch_direct_tool(method, params.unwrap_or(Value::Null)),
        }
    }

    fn dispatch_tools_call(params: Option<Value>) -> Result<Value, AppError> {
        let payload = params.ok_or_else(|| AppError::InvalidInput("missing params".to_string()))?;
        let name = payload
            .get("name")
            .and_then(|v| v.as_str())
            .ok_or_else(|| AppError::InvalidInput("missing tool name".to_string()))?;
        let arguments = payload.get("arguments").cloned().unwrap_or(Value::Null);

        let result = Self::dispatch_direct_tool(name, arguments)?;
        let text = serde_json::to_string_pretty(&result)
            .unwrap_or_else(|_| "{\"error\":\"failed to render tool result\"}".to_string());

        Ok(json!({
            "content": [
                {
                    "type": "text",
                    "text": text
                }
            ],
            "structuredContent": result,
            "isError": false
        }))
    }

    fn dispatch_direct_tool(tool_name: &str, args: Value) -> Result<Value, AppError> {
        match tool_name {
            "create_ods" => tools::create_ods::handle(args),
            "get_sheets" => tools::get_sheets::handle(args),
            "get_sheet_content" => tools::get_sheet_content::handle(args),
            "set_cell_value" => tools::set_cell_value::handle(args),
            "duplicate_sheet" => tools::duplicate_sheet::handle(args),
            "add_sheet" => tools::add_sheet::handle(args),
            "get_cell_value" => tools::get_cell_value::handle(args),
            "set_range_values" => tools::set_range_values::handle(args),
            _ => Err(AppError::InvalidInput(format!(
                "unknown method/tool: {tool_name}"
            ))),
        }
    }

    fn initialize_result() -> Value {
        json!({
            "protocolVersion": "2024-11-05",
            "capabilities": {
                "tools": {}
            },
            "serverInfo": {
                "name": "libreoffice-calc-mcp",
                "version": env!("CARGO_PKG_VERSION")
            }
        })
    }

    fn tools_list_result() -> Value {
        json!({
            "tools": [
                {
                    "name": "create_ods",
                    "description": "Create a valid ODS file with an initial sheet.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "path": { "type": "string" },
                            "overwrite": { "type": "boolean", "default": false },
                            "initial_sheet_name": { "type": "string", "default": "Hoja1" }
                        },
                        "required": ["path"]
                    }
                },
                {
                    "name": "get_sheets",
                    "description": "Return sheet names in workbook order.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "path": { "type": "string" }
                        },
                        "required": ["path"]
                    }
                },
                {
                    "name": "get_sheet_content",
                    "description": "Return a sheet as a 2D matrix.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "path": { "type": "string" },
                            "sheet": Self::sheet_selector_schema(),
                            "mode": { "type": "string", "enum": ["matrix"] },
                            "max_rows": { "type": "integer" },
                            "max_cols": { "type": "integer" },
                            "include_empty_trailing": { "type": "boolean" }
                        },
                        "required": ["path", "sheet"]
                    }
                },
                {
                    "name": "set_cell_value",
                    "description": "Set a single cell value by A1 address.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "path": { "type": "string" },
                            "sheet": Self::sheet_selector_schema(),
                            "cell": { "type": "string" },
                            "value": {
                                "type": "object",
                                "properties": {
                                    "type": { "type": "string", "enum": ["string", "number", "boolean", "empty"] },
                                    "data": {}
                                },
                                "required": ["type"]
                            }
                        },
                        "required": ["path", "sheet", "cell", "value"]
                    }
                },
                {
                    "name": "duplicate_sheet",
                    "description": "Duplicate a sheet and insert it after source.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "path": { "type": "string" },
                            "source_sheet": Self::sheet_selector_schema(),
                            "new_sheet_name": { "type": "string" }
                        },
                        "required": ["path", "source_sheet", "new_sheet_name"]
                    }
                },
                {
                    "name": "add_sheet",
                    "description": "Add an empty sheet at start or end.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "path": { "type": "string" },
                            "sheet_name": { "type": "string" },
                            "position": { "type": "string", "enum": ["start", "end"] }
                        },
                        "required": ["path", "sheet_name"]
                    }
                },
                {
                    "name": "get_cell_value",
                    "description": "Read one typed cell value.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "path": { "type": "string" },
                            "sheet": Self::sheet_selector_schema(),
                            "cell": { "type": "string" }
                        },
                        "required": ["path", "sheet", "cell"]
                    }
                },
                {
                    "name": "set_range_values",
                    "description": "Write a matrix from a start cell.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "path": { "type": "string" },
                            "sheet": Self::sheet_selector_schema(),
                            "start_cell": { "type": "string" },
                            "data": {
                                "type": "array",
                                "items": {
                                    "type": "array",
                                    "items": { "type": "string" }
                                }
                            }
                        },
                        "required": ["path", "sheet", "start_cell", "data"]
                    }
                }
            ]
        })
    }

    fn sheet_selector_schema() -> Value {
        json!({
            "oneOf": [
                {
                    "type": "object",
                    "properties": {
                        "name": { "type": "string" }
                    },
                    "required": ["name"]
                },
                {
                    "type": "object",
                    "properties": {
                        "index": { "type": "integer", "minimum": 0 }
                    },
                    "required": ["index"]
                },
                {
                    "type": "string",
                    "description": "Sheet name or JSON string like {\"name\":\"Hoja1\"} or {\"index\":0}"
                }
            ]
        })
    }
}
