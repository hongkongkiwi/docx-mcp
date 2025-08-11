use docx_mcp::security::{SecurityConfig, SecurityMiddleware, SecurityError};
use serde_json::json;
use std::collections::HashSet;
use pretty_assertions::assert_eq;
use rstest::*;

#[test]
fn test_default_security_config() {
    let config = SecurityConfig::default();
    
    assert!(!config.readonly_mode);
    assert!(config.command_whitelist.is_none());
    assert!(config.command_blacklist.is_none());
    assert_eq!(config.max_document_size, 100 * 1024 * 1024);
    assert_eq!(config.max_open_documents, 50);
    assert!(config.allow_external_tools);
    assert!(config.allow_network);
    assert!(!config.sandbox_mode);
}

#[test]
fn test_readonly_mode_allows_only_safe_commands() {
    let config = SecurityConfig {
        readonly_mode: true,
        ..Default::default()
    };
    
    // Should allow readonly commands
    assert!(config.is_command_allowed("open_document"));
    assert!(config.is_command_allowed("extract_text"));
    assert!(config.is_command_allowed("get_metadata"));
    assert!(config.is_command_allowed("search_text"));
    assert!(config.is_command_allowed("export_to_markdown"));
    
    // Should block write commands
    assert!(!config.is_command_allowed("create_document"));
    assert!(!config.is_command_allowed("add_paragraph"));
    assert!(!config.is_command_allowed("save_document"));
    assert!(!config.is_command_allowed("find_and_replace"));
    assert!(!config.is_command_allowed("convert_to_pdf"));
}

#[test]
fn test_command_whitelist() {
    let mut whitelist = HashSet::new();
    whitelist.insert("open_document".to_string());
    whitelist.insert("extract_text".to_string());
    
    let config = SecurityConfig {
        command_whitelist: Some(whitelist),
        command_blacklist: None,
        ..Default::default()
    };
    
    // Should allow whitelisted commands
    assert!(config.is_command_allowed("open_document"));
    assert!(config.is_command_allowed("extract_text"));
    
    // Should block non-whitelisted commands
    assert!(!config.is_command_allowed("create_document"));
    assert!(!config.is_command_allowed("add_paragraph"));
    assert!(!config.is_command_allowed("get_metadata"));
}

#[test]
fn test_command_blacklist() {
    let mut blacklist = HashSet::new();
    blacklist.insert("save_document".to_string());
    blacklist.insert("convert_to_pdf".to_string());
    
    let config = SecurityConfig {
        command_whitelist: None,
        command_blacklist: Some(blacklist),
        ..Default::default()
    };
    
    // Should allow non-blacklisted commands
    assert!(config.is_command_allowed("open_document"));
    assert!(config.is_command_allowed("extract_text"));
    assert!(config.is_command_allowed("add_paragraph"));
    
    // Should block blacklisted commands
    assert!(!config.is_command_allowed("save_document"));
    assert!(!config.is_command_allowed("convert_to_pdf"));
}

#[test]
fn test_whitelist_overrides_blacklist() {
    let mut whitelist = HashSet::new();
    whitelist.insert("save_document".to_string());
    
    let mut blacklist = HashSet::new();
    blacklist.insert("save_document".to_string());
    
    let config = SecurityConfig {
        command_whitelist: Some(whitelist),
        command_blacklist: Some(blacklist),
        ..Default::default()
    };
    
    // Whitelist should take precedence
    assert!(config.is_command_allowed("save_document"));
}

#[test]
fn test_external_tools_restriction() {
    let config = SecurityConfig {
        allow_external_tools: false,
        ..Default::default()
    };
    
    // Should block conversion commands that might use external tools
    assert!(!config.is_command_allowed("convert_to_pdf"));
    assert!(!config.is_command_allowed("convert_to_images"));
    
    // Should allow other commands
    assert!(config.is_command_allowed("open_document"));
    assert!(config.is_command_allowed("add_paragraph"));
}

#[test]
fn test_security_middleware_command_check() {
    let config = SecurityConfig {
        readonly_mode: true,
        ..Default::default()
    };
    let middleware = SecurityMiddleware::new(config);
    
    let safe_args = json!({"document_id": "test"});
    
    // Should pass readonly commands
    let result = middleware.check_command("extract_text", &safe_args);
    assert!(result.is_ok());
    
    // Should fail write commands
    let result = middleware.check_command("add_paragraph", &safe_args);
    assert!(matches!(result, Err(SecurityError::CommandNotAllowed(_))));
}

#[test]
fn test_sandbox_mode_path_restrictions() {
    let config = SecurityConfig {
        sandbox_mode: true,
        ..Default::default()
    };
    let middleware = SecurityMiddleware::new(config);
    
    // Should allow temp directory paths
    let temp_args = json!({"path": "/tmp/docx-mcp/test.docx"});
    let result = middleware.check_command("open_document", &temp_args);
    assert!(result.is_ok());
    
    // Should block paths outside temp directory
    let home_args = json!({"path": "/home/user/documents/test.docx"});
    let result = middleware.check_command("open_document", &home_args);
    assert!(matches!(result, Err(SecurityError::PathNotAllowed(_))));
}

#[test]
fn test_file_size_limits() {
    use tempfile::NamedTempFile;
    use std::io::Write;
    
    let config = SecurityConfig {
        max_document_size: 100, // 100 bytes limit
        ..Default::default()
    };
    let middleware = SecurityMiddleware::new(config);
    
    // Create a test file larger than limit
    let mut temp_file = NamedTempFile::new().unwrap();
    let large_content = vec![0u8; 200]; // 200 bytes
    temp_file.write_all(&large_content).unwrap();
    temp_file.flush().unwrap();
    
    let args = json!({"path": temp_file.path().to_str().unwrap()});
    let result = middleware.check_command("open_document", &args);
    
    assert!(matches!(result, Err(SecurityError::FileTooLarge { .. })));
}

#[test]
fn test_readonly_commands_list() {
    let readonly_commands = SecurityConfig::get_readonly_commands();
    
    // Should include expected readonly commands
    assert!(readonly_commands.contains("open_document"));
    assert!(readonly_commands.contains("extract_text"));
    assert!(readonly_commands.contains("get_metadata"));
    assert!(readonly_commands.contains("search_text"));
    assert!(readonly_commands.contains("analyze_formatting"));
    
    // Should not include write commands
    assert!(!readonly_commands.contains("create_document"));
    assert!(!readonly_commands.contains("add_paragraph"));
    assert!(!readonly_commands.contains("save_document"));
}

#[test]
fn test_write_commands_list() {
    let write_commands = SecurityConfig::get_write_commands();
    
    // Should include expected write commands
    assert!(write_commands.contains("create_document"));
    assert!(write_commands.contains("add_paragraph"));
    assert!(write_commands.contains("save_document"));
    assert!(write_commands.contains("find_and_replace"));
    
    // Should not include readonly commands
    assert!(!write_commands.contains("open_document"));
    assert!(!write_commands.contains("extract_text"));
    assert!(!write_commands.contains("get_metadata"));
}

#[test]
fn test_security_summary() {
    let config = SecurityConfig {
        readonly_mode: true,
        sandbox_mode: true,
        allow_external_tools: false,
        ..Default::default()
    };
    
    let summary = config.get_summary();
    assert!(summary.contains("READONLY MODE"));
    assert!(summary.contains("SANDBOX MODE"));
    assert!(summary.contains("No external tools"));
}

#[test]
fn test_combined_security_modes() {
    let mut whitelist = HashSet::new();
    whitelist.insert("open_document".to_string());
    whitelist.insert("extract_text".to_string());
    
    let config = SecurityConfig {
        readonly_mode: true,
        sandbox_mode: true,
        command_whitelist: Some(whitelist),
        command_blacklist: None,
        allow_external_tools: false,
        allow_network: false,
        max_document_size: 1024,
        ..Default::default()
    };
    
    // Should only allow whitelisted readonly commands
    assert!(config.is_command_allowed("open_document"));
    assert!(config.is_command_allowed("extract_text"));
    
    // Should block everything else
    assert!(!config.is_command_allowed("get_metadata")); // Not in whitelist
    assert!(!config.is_command_allowed("add_paragraph")); // Not readonly
    assert!(!config.is_command_allowed("convert_to_pdf")); // External tools disabled
}

#[test]
fn test_recursive_path_argument_checking() {
    let config = SecurityConfig {
        sandbox_mode: true,
        ..Default::default()
    };
    let middleware = SecurityMiddleware::new(config);
    
    // Complex nested arguments with paths
    let nested_args = json!({
        "document_id": "test",
        "options": {
            "output_path": "/home/user/bad/path.docx",
            "settings": {
                "temp_file": "/tmp/safe/path.tmp"
            }
        },
        "files": [
            "/home/user/another/bad/path.docx",
            "/tmp/docx-mcp/safe/path.docx"
        ]
    });
    
    let result = middleware.check_command("some_command", &nested_args);
    assert!(matches!(result, Err(SecurityError::PathNotAllowed(_))));
}

#[test]
fn test_security_error_messages() {
    let error = SecurityError::CommandNotAllowed("dangerous_command".to_string());
    assert!(error.to_string().contains("dangerous_command"));
    
    let error = SecurityError::PathNotAllowed("/bad/path".to_string());
    assert!(error.to_string().contains("/bad/path"));
    
    let error = SecurityError::FileTooLarge { size: 2000, max_size: 1000 };
    assert!(error.to_string().contains("2000"));
    assert!(error.to_string().contains("1000"));
}

#[fixture]
fn readonly_config() -> SecurityConfig {
    SecurityConfig {
        readonly_mode: true,
        command_blacklist: None,
        ..Default::default()
    }
}

#[fixture] 
fn sandbox_config() -> SecurityConfig {
    SecurityConfig {
        sandbox_mode: true,
        allow_external_tools: false,
        allow_network: false,
        command_blacklist: None,
        ..Default::default()
    }
}

#[fixture]
fn restrictive_config() -> SecurityConfig {
    let mut whitelist = HashSet::new();
    whitelist.insert("open_document".to_string());
    whitelist.insert("extract_text".to_string());
    
    SecurityConfig {
        readonly_mode: true,
        sandbox_mode: true,
        command_whitelist: Some(whitelist),
        command_blacklist: None,
        max_document_size: 1024 * 1024, // 1MB
        max_open_documents: 5,
        allow_external_tools: false,
        allow_network: false,
    }
}

#[rstest]
#[case("open_document", true)]
#[case("extract_text", true)]  
#[case("get_metadata", true)]
#[case("create_document", false)]
#[case("add_paragraph", false)]
#[case("save_document", false)]
fn test_readonly_mode_commands(readonly_config: SecurityConfig, #[case] command: &str, #[case] expected: bool) {
    assert_eq!(readonly_config.is_command_allowed(command), expected);
}

#[rstest]
#[case("open_document", true)]
#[case("extract_text", true)]
#[case("add_paragraph", false)] // Not in whitelist
#[case("get_metadata", false)]  // Not in whitelist
fn test_restrictive_mode_commands(restrictive_config: SecurityConfig, #[case] command: &str, #[case] expected: bool) {
    assert_eq!(restrictive_config.is_command_allowed(command), expected);
}