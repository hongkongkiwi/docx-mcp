use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use std::env;
use tracing::{debug, info, warn};
use clap::{Parser, Subcommand};

/// Command line arguments for the DOCX MCP server
#[derive(Parser, Debug)]
#[command(name = "docx-mcp")]
#[command(about = "A comprehensive Model Context Protocol (MCP) server for Microsoft Word DOCX file manipulation")]
#[command(version)]
pub struct Args {
    /// Enable readonly mode - only allow viewing operations
    #[arg(long, env = "DOCX_MCP_READONLY")]
    pub readonly: bool,

    /// Comma-separated whitelist of allowed commands
    #[arg(long, env = "DOCX_MCP_WHITELIST", value_delimiter = ',')]
    pub whitelist: Option<Vec<String>>,

    /// Comma-separated blacklist of forbidden commands  
    #[arg(long, env = "DOCX_MCP_BLACKLIST", value_delimiter = ',')]
    pub blacklist: Option<Vec<String>>,

    /// Enable sandbox mode - restrict file operations to temp directory only
    #[arg(long, env = "DOCX_MCP_SANDBOX")]
    pub sandbox: bool,

    /// Disable external tools (LibreOffice, etc.)
    #[arg(long, env = "DOCX_MCP_NO_EXTERNAL_TOOLS")]
    pub no_external_tools: bool,

    /// Disable network operations
    #[arg(long, env = "DOCX_MCP_NO_NETWORK")]
    pub no_network: bool,

    /// Maximum document size in bytes
    #[arg(long, env = "DOCX_MCP_MAX_SIZE")]
    pub max_size: Option<usize>,

    /// Maximum number of open documents
    #[arg(long, env = "DOCX_MCP_MAX_DOCS")]
    pub max_docs: Option<usize>,

    /// Optional top-level subcommand (e.g., fonts download)
    #[command(subcommand)]
    pub command: Option<CliCommand>,
}

/// Security configuration for the MCP server
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SecurityConfig {
    /// If true, only allow read-only operations
    pub readonly_mode: bool,
    
    /// Whitelist of allowed commands (if set, only these commands are allowed)
    pub command_whitelist: Option<HashSet<String>>,
    
    /// Blacklist of forbidden commands (if set, these commands are blocked)
    pub command_blacklist: Option<HashSet<String>>,
    
    /// Maximum document size in bytes (default: 100MB)
    pub max_document_size: usize,
    
    /// Maximum number of open documents (default: 50)
    pub max_open_documents: usize,
    
    /// Allow external tool usage (LibreOffice, etc.)
    pub allow_external_tools: bool,
    
    /// Allow network operations (downloading templates, fonts, etc.)
    pub allow_network: bool,
    
    /// Sandbox mode - restricts file operations to temp directory only
    pub sandbox_mode: bool,
}

/// Top-level CLI subcommands
#[derive(Subcommand, Debug, Clone, Serialize, Deserialize)]
pub enum CliCommand {
    /// Font utilities
    Fonts {
        #[command(subcommand)]
        action: FontsAction,
    },
}

/// Font-related actions
#[derive(Subcommand, Debug, Clone, Serialize, Deserialize)]
pub enum FontsAction {
    /// Download open-source fonts into assets/fonts
    Download,
    /// Verify checksums of fonts in assets/fonts
    Verify,
}

impl Default for SecurityConfig {
    fn default() -> Self {
        Self {
            readonly_mode: false,
            command_whitelist: None,
            command_blacklist: None,
            max_document_size: 100 * 1024 * 1024, // 100MB
            max_open_documents: 50,
            allow_external_tools: true,
            allow_network: true,
            sandbox_mode: false,
        }
    }
}

impl SecurityConfig {
    /// Create configuration from command line arguments
    pub fn from_args(args: Args) -> Self {
        let mut config = Self::default();
        
        // Apply command line arguments
        if args.readonly {
            config.readonly_mode = true;
            info!("Running in READONLY mode - only viewing operations allowed");
        }
        
        if let Some(whitelist) = args.whitelist {
            let commands: HashSet<String> = whitelist.into_iter().collect();
            info!("Command whitelist enabled with {} commands", commands.len());
            config.command_whitelist = Some(commands);
        }
        
        if let Some(blacklist) = args.blacklist {
            let commands: HashSet<String> = blacklist.into_iter().collect();
            info!("Command blacklist enabled with {} blocked commands", commands.len());
            config.command_blacklist = Some(commands);
        }
        
        if args.sandbox {
            config.sandbox_mode = true;
            config.allow_external_tools = false;
            config.allow_network = false;
            info!("Running in SANDBOX mode - restricted file operations");
        }
        
        if args.no_external_tools {
            config.allow_external_tools = false;
            info!("External tools disabled");
        }
        
        if args.no_network {
            config.allow_network = false;
            info!("Network operations disabled");
        }
        
        if let Some(size) = args.max_size {
            config.max_document_size = size;
            info!("Max document size set to {} bytes", size);
        }
        
        if let Some(max) = args.max_docs {
            config.max_open_documents = max;
            info!("Max open documents set to {}", max);
        }
        
        config
    }

    /// Load configuration from environment variables (deprecated, use from_args instead)
    pub fn from_env() -> Self {
        let mut config = Self::default();
        
        // Check for readonly mode
        if env::var("DOCX_MCP_READONLY").unwrap_or_default() == "true" {
            config.readonly_mode = true;
            info!("Running in READONLY mode - only viewing operations allowed");
        }
        
        // Check for command whitelist
        if let Ok(whitelist) = env::var("DOCX_MCP_WHITELIST") {
            let commands: HashSet<String> = whitelist
                .split(',')
                .map(|s| s.trim().to_string())
                .collect();
            config.command_whitelist = Some(commands.clone());
            info!("Command whitelist enabled with {} commands", commands.len());
        }
        
        // Check for command blacklist
        if let Ok(blacklist) = env::var("DOCX_MCP_BLACKLIST") {
            let commands: HashSet<String> = blacklist
                .split(',')
                .map(|s| s.trim().to_string())
                .collect();
            config.command_blacklist = Some(commands.clone());
            info!("Command blacklist enabled with {} blocked commands", commands.len());
        }
        
        // Check for sandbox mode
        if env::var("DOCX_MCP_SANDBOX").unwrap_or_default() == "true" {
            config.sandbox_mode = true;
            config.allow_external_tools = false;
            config.allow_network = false;
            info!("Running in SANDBOX mode - restricted file operations");
        }
        
        // Check for external tools permission
        if env::var("DOCX_MCP_NO_EXTERNAL_TOOLS").unwrap_or_default() == "true" {
            config.allow_external_tools = false;
            info!("External tools disabled");
        }
        
        // Check for network permission
        if env::var("DOCX_MCP_NO_NETWORK").unwrap_or_default() == "true" {
            config.allow_network = false;
            info!("Network operations disabled");
        }
        
        // Max document size
        if let Ok(size) = env::var("DOCX_MCP_MAX_SIZE") {
            if let Ok(bytes) = size.parse::<usize>() {
                config.max_document_size = bytes;
                info!("Max document size set to {} bytes", bytes);
            }
        }
        
        // Max open documents
        if let Ok(max) = env::var("DOCX_MCP_MAX_DOCS") {
            if let Ok(count) = max.parse::<usize>() {
                config.max_open_documents = count;
                info!("Max open documents set to {}", count);
            }
        }
        
        config
    }
    
    /// Check if a command is allowed based on security configuration
    pub fn is_command_allowed(&self, command: &str) -> bool {
        // First check if it's a readonly command
        let readonly_commands = Self::get_readonly_commands();
        let is_readonly_command = readonly_commands.contains(command);
        
        // In readonly mode, only allow readonly commands
        if self.readonly_mode && !is_readonly_command {
            debug!("Command '{}' blocked: readonly mode", command);
            return false;
        }
        
        // Check whitelist (if set, only whitelisted commands are allowed)
        if let Some(ref whitelist) = self.command_whitelist {
            if !whitelist.contains(command) {
                debug!("Command '{}' blocked: not in whitelist", command);
                return false;
            }
        }
        
        // Check blacklist (if set, blacklisted commands are blocked)
        if let Some(ref blacklist) = self.command_blacklist {
            if blacklist.contains(command) {
                debug!("Command '{}' blocked: in blacklist", command);
                return false;
            }
        }
        
        // Additional checks for specific command categories
        if command.starts_with("convert_") && !self.allow_external_tools {
            debug!("Command '{}' blocked: external tools disabled", command);
            return false;
        }
        
        true
    }
    
    /// Get list of readonly commands
    pub fn get_readonly_commands() -> HashSet<&'static str> {
        let mut commands = HashSet::new();
        
        // Document viewing commands
        commands.insert("open_document");
        commands.insert("extract_text");
        commands.insert("get_metadata");
        commands.insert("list_documents");
        commands.insert("get_document_info");
        commands.insert("read_paragraph");
        commands.insert("read_table");
        commands.insert("read_section");
        commands.insert("search_text");
        commands.insert("get_document_structure");
        commands.insert("get_styles");
        commands.insert("get_headers_footers");
        commands.insert("get_page_count");
        commands.insert("get_word_count");
        commands.insert("get_table_of_contents");
        commands.insert("list_bookmarks");
        commands.insert("list_hyperlinks");
        commands.insert("list_comments");
        commands.insert("list_footnotes");
        commands.insert("list_endnotes");
        commands.insert("get_document_properties");
        
        // Analysis commands
        commands.insert("analyze_formatting");
        commands.insert("check_spelling");
        commands.insert("check_grammar");
        commands.insert("get_statistics");
        commands.insert("compare_documents");
        
        // Export commands (readonly as they don't modify the original)
        commands.insert("export_to_json");
        commands.insert("export_to_markdown");
        commands.insert("export_to_html");
        commands.insert("create_preview");
        
        commands
    }
    
    /// Get list of write commands (for documentation)
    pub fn get_write_commands() -> HashSet<&'static str> {
        let mut commands = HashSet::new();
        
        // Document creation/modification
        commands.insert("create_document");
        commands.insert("save_document");
        commands.insert("close_document");
        
        // Content addition
        commands.insert("add_paragraph");
        commands.insert("add_heading");
        commands.insert("add_table");
        commands.insert("add_list");
        commands.insert("add_page_break");
        commands.insert("add_section_break");
        commands.insert("add_image");
        commands.insert("add_chart");
        commands.insert("add_shape");
        commands.insert("add_hyperlink");
        commands.insert("add_bookmark");
        commands.insert("add_footnote");
        commands.insert("add_endnote");
        commands.insert("add_comment");
        commands.insert("add_watermark");
        
        // Content modification
        commands.insert("edit_paragraph");
        commands.insert("delete_paragraph");
        commands.insert("find_and_replace");
        commands.insert("update_table");
        commands.insert("update_style");
        commands.insert("set_header");
        commands.insert("set_footer");
        commands.insert("set_margins");
        commands.insert("set_page_size");
        commands.insert("apply_template");
        commands.insert("apply_style");
        commands.insert("apply_theme");
        
        // Document operations
        commands.insert("merge_documents");
        commands.insert("split_document");
        commands.insert("convert_to_pdf");
        commands.insert("convert_to_images");
        commands.insert("protect_document");
        commands.insert("unprotect_document");
        commands.insert("track_changes");
        commands.insert("accept_changes");
        commands.insert("reject_changes");
        
        commands
    }
    
    /// Check if a file path is allowed based on sandbox configuration
    pub fn is_path_allowed(&self, path: &std::path::Path) -> bool {
        if !self.sandbox_mode {
            return true;
        }
        
        // In sandbox mode, only allow operations in temp directory
        let temp_dir = std::env::temp_dir();
        if let Ok(canonical_path) = path.canonicalize() {
            if let Ok(canonical_temp) = temp_dir.canonicalize() {
                return canonical_path.starts_with(canonical_temp);
            }
        }
        
        false
    }
    
    /// Get a summary of current security settings
    pub fn get_summary(&self) -> String {
        let mut summary: Vec<String> = Vec::new();
        
        if self.readonly_mode {
            summary.push("📖 READONLY MODE".to_string());
        }
        
        if self.sandbox_mode {
            summary.push("🔒 SANDBOX MODE".to_string());
        }
        
        if let Some(ref whitelist) = self.command_whitelist {
            summary.push(format!("✅ Whitelist: {} commands", whitelist.len()));
        }
        
        if let Some(ref blacklist) = self.command_blacklist {
            summary.push(format!("🚫 Blacklist: {} commands", blacklist.len()));
        }
        
        if !self.allow_external_tools {
            summary.push("🔧 No external tools".to_string());
        }
        
        if !self.allow_network {
            summary.push("🌐 No network access".to_string());
        }
        
        if summary.is_empty() {
            "Standard mode (all features enabled)".to_string()
        } else {
            summary.join(" | ")
        }
    }
}

/// Security middleware to check commands before execution
pub struct SecurityMiddleware {
    config: SecurityConfig,
}

impl SecurityMiddleware {
    pub fn new(config: SecurityConfig) -> Self {
        Self { config }
    }
    
    /// Check if a command should be allowed to execute
    pub fn check_command(&self, command: &str, arguments: &serde_json::Value) -> Result<(), SecurityError> {
        // Check if command is allowed
        if !self.config.is_command_allowed(command) {
            return Err(SecurityError::CommandNotAllowed(command.to_string()));
        }
        
        // Check file paths in arguments if in sandbox mode
        if self.config.sandbox_mode {
            self.check_paths_in_arguments(arguments)?;
        }
        
        // Check document size limits for open/create operations
        if command == "open_document" {
            if let Some(path) = arguments.get("path").and_then(|v| v.as_str()) {
                self.check_file_size(path)?;
            }
        }
        
        Ok(())
    }
    
    fn check_paths_in_arguments(&self, arguments: &serde_json::Value) -> Result<(), SecurityError> {
        // Recursively check all string values that look like paths
        match arguments {
            serde_json::Value::String(s) => {
                if s.contains('/') || s.contains('\\') {
                    let path = std::path::Path::new(s);
                    if !self.config.is_path_allowed(path) {
                        return Err(SecurityError::PathNotAllowed(s.to_string()));
                    }
                }
            }
            serde_json::Value::Object(map) => {
                for value in map.values() {
                    self.check_paths_in_arguments(value)?;
                }
            }
            serde_json::Value::Array(arr) => {
                for value in arr {
                    self.check_paths_in_arguments(value)?;
                }
            }
            _ => {}
        }
        Ok(())
    }
    
    fn check_file_size(&self, path: &str) -> Result<(), SecurityError> {
        let file_path = std::path::Path::new(path);
        if let Ok(metadata) = std::fs::metadata(file_path) {
            if metadata.len() as usize > self.config.max_document_size {
                return Err(SecurityError::FileTooLarge {
                    size: metadata.len() as usize,
                    max_size: self.config.max_document_size,
                });
            }
        }
        Ok(())
    }
}

#[derive(Debug, thiserror::Error)]
pub enum SecurityError {
    #[error("Command not allowed: {0}")]
    CommandNotAllowed(String),
    
    #[error("Path not allowed in sandbox mode: {0}")]
    PathNotAllowed(String),
    
    #[error("File too large: {size} bytes (max: {max_size} bytes)")]
    FileTooLarge { size: usize, max_size: usize },
    
    #[error("Maximum number of open documents exceeded")]
    TooManyDocuments,
    
    #[error("Operation requires external tools which are disabled")]
    ExternalToolsDisabled,
    
    #[error("Operation requires network access which is disabled")]
    NetworkDisabled,
}