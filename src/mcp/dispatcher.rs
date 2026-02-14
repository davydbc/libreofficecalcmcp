use crate::common::errors::AppError;
use crate::tools;
use serde_json::Value;

pub struct Dispatcher;

impl Dispatcher {
    pub fn dispatch(method: &str, params: Option<Value>) -> Result<Value, AppError> {
        // Support both direct calls ("create_ods") and MCP "tools/call" envelope.
        let (tool_name, args) = if method == "tools/call" {
            let payload =
                params.ok_or_else(|| AppError::InvalidInput("missing params".to_string()))?;
            let name = payload
                .get("name")
                .and_then(|v| v.as_str())
                .ok_or_else(|| AppError::InvalidInput("missing tool name".to_string()))?;
            let arguments = payload.get("arguments").cloned().unwrap_or(Value::Null);
            (name.to_string(), arguments)
        } else {
            (method.to_string(), params.unwrap_or(Value::Null))
        };

        match tool_name.as_str() {
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
}
