use mcp_ods::mcp::server::McpServer;

fn main() {
    // Logging goes to stderr so stdout stays reserved for JSON-RPC responses.
    tracing_subscriber::fmt()
        .with_env_filter(
            tracing_subscriber::EnvFilter::try_from_default_env().unwrap_or_else(|_| "info".into()),
        )
        .with_writer(std::io::stderr)
        .init();

    if let Err(err) = McpServer::run_stdio() {
        eprintln!("server error: {err}");
        std::process::exit(1);
    }
}
