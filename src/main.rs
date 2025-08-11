use anyhow::Result;
use mcp_server::{Server, ServerBuilder, ServerOptions};
use mcp_core::ToolManager;
use tracing::info;
use tracing_subscriber::{EnvFilter, fmt, prelude::*};

mod docx_tools;
mod docx_handler;
mod converter;
mod pure_converter;
mod advanced_docx;
mod security;

#[cfg(feature = "embedded-fonts")]
mod fonts;

use docx_tools::DocxToolsProvider;

#[tokio::main]
async fn main() -> Result<()> {
    tracing_subscriber::registry()
        .with(fmt::layer())
        .with(EnvFilter::from_default_env())
        .init();

    // Load security configuration from environment
    let security_config = security::SecurityConfig::from_env();
    info!("Starting DOCX MCP Server - Security: {}", security_config.get_summary());

    let docx_provider = DocxToolsProvider::new_with_security(security_config);
    
    let options = ServerOptions::default()
        .with_name("docx-mcp-server")
        .with_version("0.1.0");

    let server = ServerBuilder::new(options)
        .with_tool_provider(docx_provider)
        .build();

    server.run().await?;

    Ok(())
}