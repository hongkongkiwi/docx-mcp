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
#[cfg(all(feature = "runtime-server", feature = "advanced-docx"))]
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
        use mcp_server::{Router, Server};
        use mcp_server::router::RouterService;
        use mcp_server::router::CapabilitiesBuilder;
        use mcp_spec::{prompt::Prompt, resource::Resource};
        use mcp_spec::protocol::ServerCapabilities;
        use mcp_spec::content::Content;
        use mcp_spec::tool::Tool as SpecTool;
        use serde_json::Value as JsonValue;
        use std::pin::Pin;
        use std::future::Future;
        use tokio::io::{stdin, stdout};

        let security_config = security::SecurityConfig::from_args(args);
        info!("Starting DOCX MCP Server - Security: {}", security_config.get_summary());

        #[derive(Clone)]
        struct DocxRouter(docx_tools::DocxToolsProvider);

        impl Router for DocxRouter {
            fn name(&self) -> String { "docx-mcp-server".to_string() }
            fn instructions(&self) -> String { "DOCX tools for reading and exporting".to_string() }
            fn capabilities(&self) -> ServerCapabilities {
                CapabilitiesBuilder::new().with_tools(true).build()
            }
            fn list_tools(&self) -> Vec<SpecTool> {
                // DocxToolsProvider::list_tools is async; block briefly with tokio runtime handle
                let rt = tokio::runtime::Handle::current();
                let tools = rt.block_on(self.0.list_tools());
                tools.into_iter().map(|t| SpecTool{ name: t.name, description: t.description.unwrap_or_default(), input_schema: t.input_schema }).collect()
            }
            fn call_tool(&self, tool_name: &str, arguments: JsonValue) -> Pin<Box<dyn Future<Output = Result<Vec<Content>, mcp_spec::handler::ToolError>> + Send + 'static>> {
                let provider = self.0.clone();
                let name = tool_name.to_string();
                Box::pin(async move {
                    let resp = provider.call_tool(&name, arguments).await;
                    // Convert our CallToolResponse (text JSON) to Content::text
                    let text = match resp.content.get(0) {
                        Some(mcp_core::types::ToolResponseContent::Text(t)) => t.text.clone(),
                        _ => serde_json::to_string(&resp).unwrap_or_else(|_| "{}".to_string()),
                    };
                    Ok(vec![Content::text(text)])
                })
            }
            fn list_resources(&self) -> Vec<Resource> { vec![] }
            fn read_resource(&self, _uri: &str) -> Pin<Box<dyn Future<Output = Result<String, mcp_spec::handler::ResourceError>> + Send + 'static>> {
                Box::pin(async { Ok(String::new()) })
            }
            fn list_prompts(&self) -> Vec<Prompt> { vec![] }
            fn get_prompt(&self, _prompt_name: &str) -> Pin<Box<dyn Future<Output = Result<String, mcp_spec::handler::PromptError>> + Send + 'static>> {
                Box::pin(async { Ok(String::new()) })
            }
        }

        let router = DocxRouter(DocxToolsProvider::new_with_security(security_config));
        let service = RouterService(router);
        let server = Server::new(service);
        let transport = mcp_server::ByteTransport::new(stdin(), stdout());
        server.run(transport).await?;
    }

    #[cfg(not(feature = "runtime-server"))]
    {
        // No runtime server compiled in; if no subcommand was used, exit with guidance
        eprintln!("Runtime server disabled. Rebuild with --features runtime-server to run the MCP server.");
    }

    Ok(())
}