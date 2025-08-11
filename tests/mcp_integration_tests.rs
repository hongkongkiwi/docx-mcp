use docx_mcp::docx_tools::DocxToolsProvider;
use docx_mcp::security::SecurityConfig;
use mcp_core::{ToolProvider, ToolResult};
use serde_json::json;
use tempfile::TempDir;
use tokio_test;
use pretty_assertions::assert_eq;
use rstest::*;

async fn create_test_provider() -> (DocxToolsProvider, TempDir) {
    let temp_dir = TempDir::new().unwrap();
    std::env::set_var("TMPDIR", temp_dir.path());
    
    let provider = DocxToolsProvider::new();
    (provider, temp_dir)
}

async fn create_test_provider_with_security(config: SecurityConfig) -> (DocxToolsProvider, TempDir) {
    let temp_dir = TempDir::new().unwrap();
    std::env::set_var("TMPDIR", temp_dir.path());
    
    let provider = DocxToolsProvider::new_with_security(config);
    (provider, temp_dir)
}

#[tokio::test]
async fn test_list_tools_default_config() {
    let (provider, _temp_dir) = create_test_provider().await;
    
    let tools = provider.list_tools().await;
    
    // Should have all tools in default configuration
    assert!(tools.len() > 20);
    
    let tool_names: Vec<_> = tools.iter().map(|t| &t.name).collect();
    assert!(tool_names.contains(&&"create_document".to_string()));
    assert!(tool_names.contains(&&"add_paragraph".to_string()));
    assert!(tool_names.contains(&&"convert_to_pdf".to_string()));
    assert!(tool_names.contains(&&"extract_text".to_string()));
    assert!(tool_names.contains(&&"get_security_info".to_string()));
}

#[tokio::test]
async fn test_list_tools_readonly_config() {
    let config = SecurityConfig {
        readonly_mode: true,
        ..Default::default()
    };
    let (provider, _temp_dir) = create_test_provider_with_security(config).await;
    
    let tools = provider.list_tools().await;
    let tool_names: Vec<_> = tools.iter().map(|t| &t.name).collect();
    
    // Should include readonly tools
    assert!(tool_names.contains(&&"extract_text".to_string()));
    assert!(tool_names.contains(&&"get_metadata".to_string()));
    assert!(tool_names.contains(&&"search_text".to_string()));
    
    // Should not include write tools
    assert!(!tool_names.contains(&&"create_document".to_string()));
    assert!(!tool_names.contains(&&"add_paragraph".to_string()));
    assert!(!tool_names.contains(&&"save_document".to_string()));
}

#[tokio::test]
async fn test_create_document_tool() {
    let (provider, _temp_dir) = create_test_provider().await;
    
    let result = provider.call_tool("create_document", json!({})).await;
    
    match result {
        ToolResult::Success(value) => {
            assert!(value["success"].as_bool().unwrap());
            assert!(value["document_id"].is_string());
            let doc_id = value["document_id"].as_str().unwrap();
            assert!(!doc_id.is_empty());
        }
        ToolResult::Error(e) => panic!("Expected success, got error: {}", e),
    }
}

#[tokio::test]
async fn test_add_paragraph_tool() {
    let (provider, _temp_dir) = create_test_provider().await;
    
    // First create a document
    let create_result = provider.call_tool("create_document", json!({})).await;
    let doc_id = match create_result {
        ToolResult::Success(value) => value["document_id"].as_str().unwrap().to_string(),
        _ => panic!("Failed to create document"),
    };
    
    // Add paragraph
    let args = json!({
        "document_id": doc_id,
        "text": "Test paragraph content"
    });
    
    let result = provider.call_tool("add_paragraph", args).await;
    
    match result {
        ToolResult::Success(value) => {
            assert!(value["success"].as_bool().unwrap());
        }
        ToolResult::Error(e) => panic!("Expected success, got error: {}", e),
    }
    
    // Verify content was added
    let extract_args = json!({"document_id": doc_id});
    let extract_result = provider.call_tool("extract_text", extract_args).await;
    
    match extract_result {
        ToolResult::Success(value) => {
            let text = value["text"].as_str().unwrap();
            assert!(text.contains("Test paragraph content"));
        }
        ToolResult::Error(e) => panic!("Failed to extract text: {}", e),
    }
}

#[tokio::test]
async fn test_add_paragraph_with_style() {
    let (provider, _temp_dir) = create_test_provider().await;
    
    let create_result = provider.call_tool("create_document", json!({})).await;
    let doc_id = match create_result {
        ToolResult::Success(value) => value["document_id"].as_str().unwrap().to_string(),
        _ => panic!("Failed to create document"),
    };
    
    let args = json!({
        "document_id": doc_id,
        "text": "Styled paragraph",
        "style": {
            "font_size": 16,
            "bold": true,
            "color": "#FF0000",
            "alignment": "center"
        }
    });
    
    let result = provider.call_tool("add_paragraph", args).await;
    
    match result {
        ToolResult::Success(value) => {
            assert!(value["success"].as_bool().unwrap());
        }
        ToolResult::Error(e) => panic!("Expected success, got error: {}", e),
    }
}

#[tokio::test]
async fn test_add_table_tool() {
    let (provider, _temp_dir) = create_test_provider().await;
    
    let create_result = provider.call_tool("create_document", json!({})).await;
    let doc_id = match create_result {
        ToolResult::Success(value) => value["document_id"].as_str().unwrap().to_string(),
        _ => panic!("Failed to create document"),
    };
    
    let args = json!({
        "document_id": doc_id,
        "rows": [
            ["Name", "Age", "City"],
            ["Alice", "30", "New York"],
            ["Bob", "25", "Los Angeles"]
        ]
    });
    
    let result = provider.call_tool("add_table", args).await;
    
    match result {
        ToolResult::Success(value) => {
            assert!(value["success"].as_bool().unwrap());
        }
        ToolResult::Error(e) => panic!("Expected success, got error: {}", e),
    }
    
    // Verify table content
    let extract_args = json!({"document_id": doc_id});
    let extract_result = provider.call_tool("extract_text", extract_args).await;
    
    match extract_result {
        ToolResult::Success(value) => {
            let text = value["text"].as_str().unwrap();
            assert!(text.contains("Alice"));
            assert!(text.contains("New York"));
        }
        ToolResult::Error(e) => panic!("Failed to extract text: {}", e),
    }
}

#[tokio::test]
async fn test_add_heading_tool() {
    let (provider, _temp_dir) = create_test_provider().await;
    
    let create_result = provider.call_tool("create_document", json!({})).await;
    let doc_id = match create_result {
        ToolResult::Success(value) => value["document_id"].as_str().unwrap().to_string(),
        _ => panic!("Failed to create document"),
    };
    
    // Test different heading levels
    for level in 1..=6 {
        let args = json!({
            "document_id": doc_id,
            "text": format!("Heading Level {}", level),
            "level": level
        });
        
        let result = provider.call_tool("add_heading", args).await;
        
        match result {
            ToolResult::Success(value) => {
                assert!(value["success"].as_bool().unwrap());
            }
            ToolResult::Error(e) => panic!("Expected success for level {}, got error: {}", level, e),
        }
    }
}

#[tokio::test]
async fn test_add_list_tool() {
    let (provider, _temp_dir) = create_test_provider().await;
    
    let create_result = provider.call_tool("create_document", json!({})).await;
    let doc_id = match create_result {
        ToolResult::Success(value) => value["document_id"].as_str().unwrap().to_string(),
        _ => panic!("Failed to create document"),
    };
    
    // Test ordered list
    let ordered_args = json!({
        "document_id": doc_id,
        "items": ["First item", "Second item", "Third item"],
        "ordered": true
    });
    
    let result = provider.call_tool("add_list", ordered_args).await;
    assert!(matches!(result, ToolResult::Success(_)));
    
    // Test unordered list
    let unordered_args = json!({
        "document_id": doc_id,
        "items": ["Bullet one", "Bullet two", "Bullet three"],
        "ordered": false
    });
    
    let result = provider.call_tool("add_list", unordered_args).await;
    assert!(matches!(result, ToolResult::Success(_)));
}

#[tokio::test]
async fn test_get_metadata_tool() {
    let (provider, _temp_dir) = create_test_provider().await;
    
    let create_result = provider.call_tool("create_document", json!({})).await;
    let doc_id = match create_result {
        ToolResult::Success(value) => value["document_id"].as_str().unwrap().to_string(),
        _ => panic!("Failed to create document"),
    };
    
    let args = json!({"document_id": doc_id});
    let result = provider.call_tool("get_metadata", args).await;
    
    match result {
        ToolResult::Success(value) => {
            assert!(value["success"].as_bool().unwrap());
            let metadata = &value["metadata"];
            assert_eq!(metadata["id"], doc_id);
            assert!(metadata["path"].is_string());
            assert!(metadata["created_at"].is_string());
        }
        ToolResult::Error(e) => panic!("Expected success, got error: {}", e),
    }
}

#[tokio::test]
async fn test_search_text_tool() {
    let (provider, _temp_dir) = create_test_provider().await;
    
    let create_result = provider.call_tool("create_document", json!({})).await;
    let doc_id = match create_result {
        ToolResult::Success(value) => value["document_id"].as_str().unwrap().to_string(),
        _ => panic!("Failed to create document"),
    };
    
    // Add some content to search
    let add_args = json!({
        "document_id": doc_id,
        "text": "This is a test document with searchable content. The word test appears multiple times."
    });
    provider.call_tool("add_paragraph", add_args).await;
    
    // Search for text
    let search_args = json!({
        "document_id": doc_id,
        "search_term": "test",
        "case_sensitive": false
    });
    
    let result = provider.call_tool("search_text", search_args).await;
    
    match result {
        ToolResult::Success(value) => {
            assert!(value["success"].as_bool().unwrap());
            let matches = value["matches"].as_array().unwrap();
            assert!(matches.len() > 0);
            assert!(value["total_matches"].as_u64().unwrap() > 0);
        }
        ToolResult::Error(e) => panic!("Expected success, got error: {}", e),
    }
}

#[tokio::test]
async fn test_get_word_count_tool() {
    let (provider, _temp_dir) = create_test_provider().await;
    
    let create_result = provider.call_tool("create_document", json!({})).await;
    let doc_id = match create_result {
        ToolResult::Success(value) => value["document_id"].as_str().unwrap().to_string(),
        _ => panic!("Failed to create document"),
    };
    
    // Add content with known word count
    let content = "This sentence has exactly five words. This is another sentence with seven words total.";
    let add_args = json!({
        "document_id": doc_id,
        "text": content
    });
    provider.call_tool("add_paragraph", add_args).await;
    
    let args = json!({"document_id": doc_id});
    let result = provider.call_tool("get_word_count", args).await;
    
    match result {
        ToolResult::Success(value) => {
            assert!(value["success"].as_bool().unwrap());
            let stats = &value["statistics"];
            assert!(stats["words"].as_u64().unwrap() > 10);
            assert!(stats["characters"].as_u64().unwrap() > 0);
            assert!(stats["sentences"].as_u64().unwrap() >= 2);
        }
        ToolResult::Error(e) => panic!("Expected success, got error: {}", e),
    }
}

#[tokio::test]
async fn test_get_security_info_tool() {
    let config = SecurityConfig {
        readonly_mode: true,
        sandbox_mode: true,
        ..Default::default()
    };
    let (provider, _temp_dir) = create_test_provider_with_security(config).await;
    
    let result = provider.call_tool("get_security_info", json!({})).await;
    
    match result {
        ToolResult::Success(value) => {
            assert!(value["success"].as_bool().unwrap());
            let security = &value["security"];
            assert_eq!(security["readonly_mode"], true);
            assert_eq!(security["sandbox_mode"], true);
            assert!(security["summary"].is_string());
        }
        ToolResult::Error(e) => panic!("Expected success, got error: {}", e),
    }
}

#[tokio::test]
async fn test_readonly_mode_blocks_write_operations() {
    let config = SecurityConfig {
        readonly_mode: true,
        ..Default::default()
    };
    let (provider, _temp_dir) = create_test_provider_with_security(config).await;
    
    // Should fail to create document in readonly mode
    let result = provider.call_tool("create_document", json!({})).await;
    
    match result {
        ToolResult::Error(e) => {
            assert!(e.contains("Security check failed"));
            assert!(e.contains("Command not allowed"));
        }
        ToolResult::Success(_) => panic!("Expected security error, got success"),
    }
}

#[tokio::test]
async fn test_document_not_found_error() {
    let (provider, _temp_dir) = create_test_provider().await;
    
    let args = json!({"document_id": "nonexistent-doc-id"});
    let result = provider.call_tool("extract_text", args).await;
    
    match result {
        ToolResult::Success(value) => {
            assert!(!value["success"].as_bool().unwrap());
            assert!(value["error"].as_str().unwrap().contains("Document not found"));
        }
        ToolResult::Error(_) => {
            // This is also acceptable - depends on implementation
        }
    }
}

#[tokio::test]
async fn test_invalid_tool_name() {
    let (provider, _temp_dir) = create_test_provider().await;
    
    let result = provider.call_tool("nonexistent_tool", json!({})).await;
    
    match result {
        ToolResult::Success(value) => {
            assert!(!value["success"].as_bool().unwrap());
            assert!(value["error"].as_str().unwrap().contains("Unknown tool"));
        }
        ToolResult::Error(e) => {
            assert!(e.contains("Unknown tool"));
        }
    }
}

#[tokio::test]
async fn test_multiple_documents() {
    let (provider, _temp_dir) = create_test_provider().await;
    
    let mut doc_ids = Vec::new();
    
    // Create multiple documents
    for i in 0..3 {
        let result = provider.call_tool("create_document", json!({})).await;
        let doc_id = match result {
            ToolResult::Success(value) => value["document_id"].as_str().unwrap().to_string(),
            _ => panic!("Failed to create document {}", i),
        };
        
        // Add unique content to each
        let args = json!({
            "document_id": doc_id,
            "text": format!("Document {} content", i)
        });
        provider.call_tool("add_paragraph", args).await;
        
        doc_ids.push(doc_id);
    }
    
    // List documents
    let list_result = provider.call_tool("list_documents", json!({})).await;
    
    match list_result {
        ToolResult::Success(value) => {
            assert!(value["success"].as_bool().unwrap());
            let documents = value["documents"].as_array().unwrap();
            assert!(documents.len() >= 3);
        }
        ToolResult::Error(e) => panic!("Failed to list documents: {}", e),
    }
    
    // Verify each document has its unique content
    for (i, doc_id) in doc_ids.iter().enumerate() {
        let args = json!({"document_id": doc_id});
        let result = provider.call_tool("extract_text", args).await;
        
        match result {
            ToolResult::Success(value) => {
                let text = value["text"].as_str().unwrap();
                assert!(text.contains(&format!("Document {} content", i)));
            }
            ToolResult::Error(e) => panic!("Failed to extract text from document {}: {}", i, e),
        }
    }
}

#[tokio::test]
async fn test_export_to_markdown() {
    let (provider, temp_dir) = create_test_provider().await;
    
    let create_result = provider.call_tool("create_document", json!({})).await;
    let doc_id = match create_result {
        ToolResult::Success(value) => value["document_id"].as_str().unwrap().to_string(),
        _ => panic!("Failed to create document"),
    };
    
    // Add content
    provider.call_tool("add_heading", json!({
        "document_id": doc_id,
        "text": "Test Document",
        "level": 1
    })).await;
    
    provider.call_tool("add_paragraph", json!({
        "document_id": doc_id,
        "text": "This is a test paragraph."
    })).await;
    
    // Export to markdown
    let output_path = temp_dir.path().join("test_export.md");
    let args = json!({
        "document_id": doc_id,
        "output_path": output_path.to_str().unwrap()
    });
    
    let result = provider.call_tool("export_to_markdown", args).await;
    
    match result {
        ToolResult::Success(value) => {
            assert!(value["success"].as_bool().unwrap());
            assert!(output_path.exists());
            
            let content = std::fs::read_to_string(&output_path).unwrap();
            assert!(content.contains("# Test Document"));
            assert!(content.contains("test paragraph"));
        }
        ToolResult::Error(e) => panic!("Expected success, got error: {}", e),
    }
}

// Parametrized test using rstest
#[rstest]
#[case("create_document", json!({}))]
#[case("list_documents", json!({}))]
#[case("get_security_info", json!({}))]
#[tokio::test]
async fn test_tools_without_document_id(#[case] tool_name: &str, #[case] args: serde_json::Value) {
    let (provider, _temp_dir) = create_test_provider().await;
    
    let result = provider.call_tool(tool_name, args).await;
    
    // These tools should work without requiring a document_id
    match result {
        ToolResult::Success(value) => {
            assert!(value["success"].as_bool().unwrap_or(false));
        }
        ToolResult::Error(e) => panic!("Tool {} failed: {}", tool_name, e),
    }
}

#[tokio::test]
async fn test_tool_input_validation() {
    let (provider, _temp_dir) = create_test_provider().await;
    
    // Missing required arguments should fail gracefully
    let result = provider.call_tool("add_paragraph", json!({})).await;
    
    match result {
        ToolResult::Success(value) => {
            // Should fail due to missing document_id
            assert!(!value["success"].as_bool().unwrap_or(true));
        }
        ToolResult::Error(_) => {
            // This is also acceptable
        }
    }
}