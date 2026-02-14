use crate::common::errors::AppError;
use crate::mcp::dispatcher::Dispatcher;
use crate::mcp::protocol::{JsonRpcRequest, JsonRpcResponse};
use std::io::{self, BufRead, Write};
use tracing::error;

pub struct McpServer;

impl McpServer {
    pub fn run_stdio() -> Result<(), AppError> {
        let stdin = io::stdin();
        let mut stdout = io::stdout().lock();

        for line in stdin.lock().lines() {
            let line = line?;
            if line.trim().is_empty() {
                continue;
            }

            if let Some(response) = Self::handle_line(&line) {
                Self::write_response(&mut stdout, &response)?;
            }
        }

        Ok(())
    }

    fn handle_line(line: &str) -> Option<JsonRpcResponse> {
        // Protocol is line-oriented: one JSON-RPC request per stdin line.
        let request: Result<JsonRpcRequest, _> = serde_json::from_str(line);
        match request {
            Ok(req) => {
                let result = Dispatcher::dispatch(&req.method, req.params);

                // JSON-RPC notification: no id means fire-and-forget, no response body.
                if req.id.is_none() {
                    if let Err(err) = result {
                        error!("notification handling error: {err}");
                    }
                    return None;
                }

                Some(match result {
                    Ok(result) => JsonRpcResponse::success(req.id, result),
                    Err(err) => JsonRpcResponse::failure(req.id, err.code(), err.to_string()),
                })
            }
            Err(err) => Some(JsonRpcResponse::failure(
                None,
                -32700,
                format!("invalid json-rpc request: {err}"),
            )),
        }
    }

    fn write_response(stdout: &mut impl Write, response: &JsonRpcResponse) -> Result<(), AppError> {
        let output =
            serde_json::to_string(response).map_err(|e| AppError::InvalidInput(e.to_string()))?;
        if let Err(e) = writeln!(stdout, "{output}") {
            error!("failed to write response: {e}");
            return Err(AppError::IoError(e.to_string()));
        }
        stdout.flush()?;
        Ok(())
    }
}
