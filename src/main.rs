use anyhow::Result;
#[cfg(feature = "runtime-server")]
use mcp_server::Server;
use tracing::info;
use tracing_subscriber::{EnvFilter, fmt, prelude::*};
use clap::Parser;

#[cfg(feature = "runtime-server")]
mod docx_tools;
#[cfg(feature = "runtime-server")]
mod docx_handler;
#[cfg(feature = "runtime-server")]
mod converter;
#[cfg(feature = "runtime-server")]
mod pure_converter;
#[cfg(feature = "runtime-server")]
mod advanced_docx;
mod security;

#[cfg(feature = "embedded-fonts")]
mod fonts;

#[cfg(feature = "runtime-server")]
use docx_tools::DocxToolsProvider;

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
                        docx_mcp::fonts_cli::download_fonts_blocking()?;
                        info!("Fonts downloaded successfully");
                        return Ok(());
                    }
                    security::FontsAction::Verify => {
                        docx_mcp::fonts_cli::verify_fonts_blocking()?;
                        info!("Fonts verified successfully");
                        return Ok(());
                    }
                }
            }
        }
    }

    #[cfg(feature = "runtime-server")]
    {
        let security_config = security::SecurityConfig::from_args(args);
        info!("Starting DOCX MCP Server - Security: {}", security_config.get_summary());

        // TODO: Integrate with mcp-server Router here. For now, just exit successfully.
        info!("Server integration pending refactor; exiting.");
    }

    #[cfg(not(feature = "runtime-server"))]
    {
        // No runtime server compiled in; if no subcommand was used, exit with guidance
        eprintln!("Runtime server disabled. Rebuild with --features runtime-server to run the MCP server.");
    }

    Ok(())
}