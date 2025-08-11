use anyhow::Result;
use mcp_server::{Server, ServerBuilder, ServerOptions};
use mcp_core::ToolManager;
use tracing::{info, warn};
use tracing_subscriber::{EnvFilter, fmt, prelude::*};
use clap::Parser;

mod docx_tools;
mod docx_handler;
mod converter;
mod pure_converter;
mod advanced_docx;
mod security;

#[cfg(feature = "embedded-fonts")]
mod fonts;

use docx_tools::DocxToolsProvider;
use std::process::Command;

#[tokio::main]
async fn main() -> Result<()> {
    tracing_subscriber::registry()
        .with(fmt::layer())
        .with(EnvFilter::from_default_env())
        .init();

    // Parse command line arguments (which also includes environment variables)
    let args = security::Args::parse();

    // Handle top-level subcommands that should run and exit
    if let Some(cmd) = &args.command {
        match cmd {
            security::CliCommand::Fonts { action } => {
                match action {
                    security::FontsAction::Download => {
                        info!("Downloading fonts via embedded helper...");
                        // Prefer the script if available; otherwise, fetch directly in the future
                        let script_path = "./download_fonts.sh";
                        if !std::path::Path::new(script_path).exists() {
                            warn!("download_fonts.sh not found; please run it manually or pull latest");
                            anyhow::bail!("download_fonts.sh not found");
                        }

                        let status = Command::new(script_path).status()?;
                        if !status.success() {
                            anyhow::bail!("Font download helper failed");
                        }
                        info!("Fonts downloaded successfully");
                        return Ok(());
                    }
                }
            }
        }
    }

    let security_config = security::SecurityConfig::from_args(args);
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