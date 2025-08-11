use docx_mcp::security::{Args, SecurityConfig};
use clap::Parser;
use std::env;

fn reset_env() {
    for (k, _) in env::vars() {
        if k.starts_with("DOCX_MCP_") {
            env::remove_var(k);
        }
    }
}

#[test]
fn parses_flags_and_lists() {
    reset_env();

    let argv = [
        "docx-mcp",
        "--readonly",
        "--sandbox",
        "--no-external-tools",
        "--no-network",
        "--whitelist",
        "open_document,extract_text,get_metadata",
        "--blacklist",
        "save_document,add_paragraph",
        "--max-size",
        "1048576",
        "--max-docs",
        "10",
    ];

    let args = Args::parse_from(&argv);
    assert!(args.readonly);
    assert!(args.sandbox);
    assert!(args.no_external_tools);
    assert!(args.no_network);
    assert_eq!(args.max_size, Some(1_048_576));
    assert_eq!(args.max_docs, Some(10));

    let wl = args.whitelist.clone().unwrap();
    assert_eq!(wl, vec![
        "open_document".to_string(),
        "extract_text".to_string(),
        "get_metadata".to_string(),
    ]);

    let bl = args.blacklist.clone().unwrap();
    assert_eq!(bl, vec![
        "save_document".to_string(),
        "add_paragraph".to_string(),
    ]);

    let cfg = SecurityConfig::from_args(args);
    assert!(cfg.readonly_mode);
    assert!(cfg.sandbox_mode);
    assert!(!cfg.allow_external_tools);
    assert!(!cfg.allow_network);
    assert_eq!(cfg.max_document_size, 1_048_576);
    assert_eq!(cfg.max_open_documents, 10);

    let wlset = cfg.command_whitelist.unwrap();
    assert!(wlset.contains("open_document"));
    assert!(wlset.contains("extract_text"));
    assert!(wlset.contains("get_metadata"));

    let blset = cfg.command_blacklist.unwrap();
    assert!(blset.contains("save_document"));
    assert!(blset.contains("add_paragraph"));
}

#[test]
fn parses_from_environment() {
    reset_env();

    env::set_var("DOCX_MCP_READONLY", "true");
    env::set_var("DOCX_MCP_SANDBOX", "true");
    env::set_var("DOCX_MCP_NO_EXTERNAL_TOOLS", "true");
    env::set_var("DOCX_MCP_NO_NETWORK", "true");
    env::set_var("DOCX_MCP_WHITELIST", "open_document,extract_text");
    env::set_var("DOCX_MCP_BLACKLIST", "save_document");
    env::set_var("DOCX_MCP_MAX_SIZE", "2048");
    env::set_var("DOCX_MCP_MAX_DOCS", "7");

    let cfg = SecurityConfig::from_env();

    assert!(cfg.readonly_mode);
    assert!(cfg.sandbox_mode);
    assert!(!cfg.allow_external_tools);
    assert!(!cfg.allow_network);
    assert_eq!(cfg.max_document_size, 2048);
    assert_eq!(cfg.max_open_documents, 7);

    let wl = cfg.command_whitelist.unwrap();
    assert!(wl.contains("open_document"));
    assert!(wl.contains("extract_text"));

    let bl = cfg.command_blacklist.unwrap();
    assert!(bl.contains("save_document"));
}
