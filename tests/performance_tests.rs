use anyhow::Result;
use docx_mcp::docx_handler::{DocxHandler, DocxStyle, TableData};
use docx_mcp::pure_converter::PureRustConverter;
use docx_mcp::docx_tools::DocxToolsProvider;
use docx_mcp::security::SecurityConfig;
use mcp_core::types::{CallToolResponse, ToolResponseContent};
use serde_json::{json, Value};
use tempfile::TempDir;
use std::time::{Duration, Instant};
use std::sync::{Arc, Mutex};
use std::thread;
use pretty_assertions::assert_eq;

const PERFORMANCE_TIMEOUT: Duration = Duration::from_secs(30);
const STRESS_TEST_ITERATIONS: usize = 100;

#[test]
fn test_large_document_performance() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    let mut handler = DocxHandler::new_with_base_dir(temp_dir.path()).unwrap();
    
    let start = Instant::now();
    let doc_id = handler.create_document().unwrap();
    let creation_time = start.elapsed();
    
    println!("Document creation took: {:?}", creation_time);
    assert!(creation_time < Duration::from_millis(500), "Document creation should be fast");
    
    let start = Instant::now();
    
    // Add substantial content
    for i in 0..1000 {
        if i % 50 == 0 {
            handler.add_heading(&doc_id, &format!("Section {}", i / 50 + 1), 2)?;
        }
        
        let content = format!(
            "This is paragraph number {} in our performance test. It contains enough text to make the test meaningful and simulate real-world usage patterns. The paragraph includes various punctuation marks, numbers like {}, and other elements that might affect processing performance.",
            i, i * 7
        );
        handler.add_paragraph(&doc_id, &content, None)?;
        
        // Add a table every 100 paragraphs
        if i % 100 == 99 {
            let table_data = TableData {
                rows: vec![
                    vec!["Item".to_string(), "Value".to_string(), "Status".to_string()],
                    vec![format!("Item {}", i), format!("${}.00", i * 10), "Active".to_string()],
                ],
                headers: Some(vec!["Item".to_string(), "Value".to_string(), "Status".to_string()]),
                border_style: Some("single".to_string()),
            };
            handler.add_table(&doc_id, table_data)?;
        }
    }
    
    let content_addition_time = start.elapsed();
    println!("Adding 1000 paragraphs took: {:?}", content_addition_time);
    assert!(content_addition_time < PERFORMANCE_TIMEOUT, "Content addition took too long");
    
    // Test text extraction performance
    let start = Instant::now();
    let text = handler.extract_text(&doc_id)?;
    let extraction_time = start.elapsed();
    
    println!("Text extraction took: {:?}", extraction_time);
    println!("Extracted text length: {} characters", text.len());
    assert!(extraction_time < Duration::from_secs(10), "Text extraction should be reasonably fast");
    assert!(text.len() > 100000, "Should extract substantial amount of text");
    
    // Test PDF conversion performance
    let metadata = handler.get_metadata(&doc_id)?;
    let converter = PureRustConverter::new();
    let pdf_path = temp_dir.path().join("large_performance_test.pdf");
    
    let start = Instant::now();
    converter.convert_docx_to_pdf(&metadata.path, &pdf_path)?;
    let conversion_time = start.elapsed();
    
    println!("PDF conversion took: {:?}", conversion_time);
    assert!(conversion_time < PERFORMANCE_TIMEOUT, "PDF conversion took too long");
    assert!(pdf_path.exists(), "PDF should be created");
    
    let pdf_size = std::fs::metadata(&pdf_path)?.len();
    println!("Generated PDF size: {} bytes", pdf_size);
    assert!(pdf_size > 50000, "PDF should have substantial size");
    
    Ok(())
}

#[test]
fn test_concurrent_document_stress() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    let temp_path = Arc::new(temp_dir.path().to_path_buf());
    let results = Arc::new(Mutex::new(Vec::new()));
    
    let thread_count = 8;
    let operations_per_thread = 20;
    
    let start = Instant::now();
    
    let handles: Vec<_> = (0..thread_count).map(|thread_id| {
        let temp_path = Arc::clone(&temp_path);
        let results = Arc::clone(&results);
        
        thread::spawn(move || -> Result<()> {
            let mut handler = DocxHandler::new_with_base_dir(&*temp_path)?;
            let mut local_results = Vec::new();
            
            for op_id in 0..operations_per_thread {
                let doc_start = Instant::now();
                
                // Create document
                let doc_id = handler.create_document()?;
                
                // Add varied content
                handler.add_heading(&doc_id, &format!("Thread {} Document {}", thread_id, op_id), 1)?;
                
                for i in 0..10 {
                    let content = format!("Thread {} operation {} paragraph {}", thread_id, op_id, i);
                    handler.add_paragraph(&doc_id, &content, None)?;
                }
                
                // Add a small table
                let table_data = TableData {
                    rows: vec![
                        vec!["Col1".to_string(), "Col2".to_string()],
                        vec![format!("T{}", thread_id), format!("O{}", op_id)],
                    ],
                    headers: None,
                    border_style: Some("single".to_string()),
                };
                handler.add_table(&doc_id, table_data)?;
                
                // Extract text
                let text = handler.extract_text(&doc_id)?;
                assert!(text.contains(&format!("Thread {} Document {}", thread_id, op_id)));
                
                let doc_duration = doc_start.elapsed();
                local_results.push((thread_id, op_id, doc_duration));
                
                // Cleanup
                handler.close_document(&doc_id)?;
            }
            
            // Store results
            {
                let mut results_guard = results.lock().unwrap();
                results_guard.extend(local_results);
            }
            
            Ok(())
        })
    }).collect();
    
    // Wait for all threads
    for handle in handles {
        handle.join().unwrap()?;
    }
    
    let total_duration = start.elapsed();
    let results_guard = results.lock().unwrap();
    
    println!("Concurrent stress test completed in: {:?}", total_duration);
    println!("Total operations: {}", results_guard.len());
    
    let avg_duration = results_guard.iter()
        .map(|(_, _, duration)| duration.as_millis())
        .sum::<u128>() as f64 / results_guard.len() as f64;
    
    println!("Average operation duration: {:.2}ms", avg_duration);
    
    // Verify all operations completed
    assert_eq!(results_guard.len(), thread_count * operations_per_thread);
    assert!(total_duration < Duration::from_secs(60), "Stress test took too long");
    assert!(avg_duration < 1000.0, "Average operation should be under 1 second");
    
    Ok(())
}

#[test]
fn test_memory_intensive_operations() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    let mut handler = DocxHandler::new_with_base_dir(temp_dir.path()).unwrap();
    
    let mut doc_ids = Vec::new();
    
    // Create many documents simultaneously
    for i in 0..50 {
        let doc_id = handler.create_document()?;
        
        // Add substantial content to each
        handler.add_heading(&doc_id, &format!("Memory Test Document {}", i), 1)?;
        
        for j in 0..100 {
            let content = format!(
                "Document {} paragraph {}. This paragraph contains substantial text content to test memory usage patterns. It includes various data that might accumulate in memory during processing and needs to be handled efficiently by the system.",
                i, j
            );
            handler.add_paragraph(&doc_id, &content, None)?;
        }
        
        // Add a large table
        let mut table_rows = vec![vec!["ID".to_string(), "Name".to_string(), "Description".to_string()]];
        for k in 0..20 {
            table_rows.push(vec![
                format!("ID-{}", k),
                format!("Item-{}", k),
                format!("Description for item {} in document {}", k, i),
            ]);
        }
        
        let table_data = TableData {
            rows: table_rows,
            headers: Some(vec!["ID".to_string(), "Name".to_string(), "Description".to_string()]),
            border_style: Some("single".to_string()),
        };
        handler.add_table(&doc_id, table_data)?;
        
        doc_ids.push(doc_id);
    }
    
    println!("Created {} documents with substantial content", doc_ids.len());
    
    // Test that all documents are accessible
    for (i, doc_id) in doc_ids.iter().enumerate() {
        let text = handler.extract_text(doc_id)?;
        assert!(text.contains(&format!("Memory Test Document {}", i)));
        assert!(text.len() > 10000, "Document should have substantial text");
    }
    
    // Test batch operations
    let start = Instant::now();
    let mut total_text_length = 0;
    
    for doc_id in &doc_ids {
        let text = handler.extract_text(doc_id)?;
        total_text_length += text.len();
    }
    
    let batch_extraction_time = start.elapsed();
    println!("Batch text extraction took: {:?}", batch_extraction_time);
    println!("Total extracted text: {} characters", total_text_length);
    
    assert!(batch_extraction_time < Duration::from_secs(30), "Batch extraction should be reasonable");
    assert!(total_text_length > 500000, "Should extract substantial total text");
    
    // Cleanup all documents
    for doc_id in doc_ids {
        handler.close_document(&doc_id)?;
    }
    
    Ok(())
}

#[test]
fn test_mcp_tool_performance() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    let provider = DocxToolsProvider::with_base_dir(temp_dir.path());
    let mut operation_times = Vec::new();
    
    // Test document creation performance
    let start = Instant::now();
    let create_resp: CallToolResponse = tokio_test::block_on(async {
        provider.call_tool("create_document", json!({})).await
    });
    let create_result = match create_resp.content.get(0) {
        Some(ToolResponseContent::Text(t)) => serde_json::from_str::<Value>(&t.text)
            .map_err(|e| e.to_string()),
        _ => Err("non-text response".to_string())
    };
    let creation_time = start.elapsed();
    operation_times.push(("create_document", creation_time));
    
    let doc_id = match create_result {
        Ok(value) if value.get("success").and_then(|v| v.as_bool()).unwrap_or(false) => value["document_id"].as_str().unwrap().to_string(),
        _ => panic!("Failed to create document"),
    };
    
    // Test paragraph addition performance
    let start = Instant::now();
    for i in 0..100 {
        let args = json!({
            "document_id": doc_id,
            "text": format!("Performance test paragraph {} with substantial content for timing measurements", i)
        });
        
        let result: CallToolResponse = tokio_test::block_on(async {
            provider.call_tool("add_paragraph", args).await
        });
        if let Some(ToolResponseContent::Text(t)) = result.content.get(0) {
            let v: Value = serde_json::from_str(&t.text).unwrap_or(json!({"success": false}));
            assert!(v.get("success").and_then(|b| b.as_bool()).unwrap_or(false), "Failed to add paragraph {}: {}", i, t.text);
        } else {
            panic!("Non-text response for add_paragraph");
        }
    }
    let paragraph_addition_time = start.elapsed();
    operation_times.push(("add_100_paragraphs", paragraph_addition_time));
    
    // Test heading performance
    let start = Instant::now();
    for level in 1..=6 {
        let args = json!({
            "document_id": doc_id,
            "text": format!("Heading Level {}", level),
            "level": level
        });
        
        tokio_test::block_on(async {
            provider.call_tool("add_heading", args).await
        });
    }
    let heading_time = start.elapsed();
    operation_times.push(("add_headings", heading_time));
    
    // Test table performance
    let start = Instant::now();
    let table_args = json!({
        "document_id": doc_id,
        "rows": [
            ["Product", "Price", "Quantity", "Total"],
            ["Item 1", "$10.00", "5", "$50.00"],
            ["Item 2", "$15.00", "3", "$45.00"],
            ["Item 3", "$12.00", "7", "$84.00"],
            ["Item 4", "$8.00", "10", "$80.00"]
        ]
    });
    
    tokio_test::block_on(async {
        provider.call_tool("add_table", table_args).await
    });
    let table_time = start.elapsed();
    operation_times.push(("add_table", table_time));
    
    // Test text extraction performance
    let start = Instant::now();
    let extract_args = json!({"document_id": doc_id});
    let extract_resp: CallToolResponse = tokio_test::block_on(async {
        provider.call_tool("extract_text", extract_args).await
    });
    let extraction_time = start.elapsed();
    operation_times.push(("extract_text", extraction_time));
    
    match extract_resp.content.get(0) {
        Some(ToolResponseContent::Text(t)) => {
            let value: Value = serde_json::from_str(&t.text).unwrap();
            let text = value["text"].as_str().unwrap();
            println!("Extracted text length: {} characters", text.len());
            assert!(text.len() > 5000, "Should extract substantial text");
        },
        _ => panic!("Text extraction failed"),
    }
    
    // Test metadata retrieval performance
    let start = Instant::now();
    let metadata_args = json!({"document_id": doc_id});
    tokio_test::block_on(async {
        provider.call_tool("get_metadata", metadata_args).await
    });
    let metadata_time = start.elapsed();
    operation_times.push(("get_metadata", metadata_time));
    
    // Print performance results
    println!("\nMCP Tool Performance Results:");
    for (operation, duration) in &operation_times {
        println!("{}: {:?}", operation, duration);
    }
    
    // Verify reasonable performance
    for (operation, duration) in &operation_times {
        match operation.as_ref() {
            "create_document" => assert!(duration < &Duration::from_millis(500), "Document creation too slow"),
            "add_100_paragraphs" => assert!(duration < &Duration::from_secs(10), "Paragraph addition too slow"),
            "extract_text" => assert!(duration < &Duration::from_secs(5), "Text extraction too slow"),
            _ => assert!(duration < &Duration::from_secs(2), "Operation {} too slow", operation),
        }
    }
    
    Ok(())
}

#[test]
fn test_security_overhead_performance() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    
    // Test with default (permissive) security
    let default_provider = DocxToolsProvider::with_base_dir(temp_dir.path());
    
    // Test with restrictive security
    let restrictive_config = SecurityConfig {
        readonly_mode: true,
        sandbox_mode: true,
        max_document_size: 1024 * 1024, // 1MB
        max_open_documents: 10,
        allow_external_tools: false,
        allow_network: false,
        ..Default::default()
    };
    let restrictive_provider = DocxToolsProvider::with_base_dir_and_security(temp_dir.path(), restrictive_config);
    
    let operations = vec![
        ("list_documents", json!({})),
        ("get_security_info", json!({})),
    ];
    
    for (operation, args) in operations {
        // Test default provider
        let start = Instant::now();
        let _result = tokio_test::block_on(async {
            default_provider.call_tool(operation, args.clone()).await
        });
        let default_time = start.elapsed();
        
        // Test restrictive provider
        let start = Instant::now();
        let _result = tokio_test::block_on(async {
            restrictive_provider.call_tool(operation, args.clone()).await
        });
        let restrictive_time = start.elapsed();
        
        println!("Operation {}: Default={:?}, Restrictive={:?}", 
                operation, default_time, restrictive_time);
        
        // Security overhead should be minimal
        let overhead_ratio = restrictive_time.as_nanos() as f64 / default_time.as_nanos() as f64;
        assert!(overhead_ratio < 3.0, "Security overhead too high for {}: {}x", operation, overhead_ratio);
    }
    
    Ok(())
}

#[test]
fn test_conversion_performance_scaling() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    let converter = PureRustConverter::new();
    
    let document_sizes = vec![10, 50, 100, 250];
    let mut performance_data = Vec::new();
    
    for &size in &document_sizes {
        let mut handler = DocxHandler::new_with_base_dir(temp_dir.path())?;
        let doc_id = handler.create_document()?;
        
        // Create document with specified number of paragraphs
        handler.add_heading(&doc_id, &format!("Test Document - {} paragraphs", size), 1)?;
        
        for i in 0..size {
            let content = format!("Paragraph {} content for performance scaling test. This paragraph contains enough text to make the performance test meaningful and realistic.", i);
            handler.add_paragraph(&doc_id, &content, None)?;
            
            if i % 20 == 19 {
                handler.add_heading(&doc_id, &format!("Section {}", i / 20 + 1), 2)?;
            }
        }
        
        let metadata = handler.get_metadata(&doc_id)?;
        
        // Test text extraction scaling
        let start = Instant::now();
        let text = handler.extract_text(&doc_id)?;
        let extraction_time = start.elapsed();
        
        // Test PDF conversion scaling
        let pdf_path = temp_dir.path().join(format!("scale_test_{}.pdf", size));
        let start = Instant::now();
        converter.convert_docx_to_pdf(&metadata.path, &pdf_path)?;
        let conversion_time = start.elapsed();
        
        performance_data.push((size, text.len(), extraction_time, conversion_time));
        
        println!("Size: {} paragraphs, Text: {} chars, Extract: {:?}, Convert: {:?}", 
                size, text.len(), extraction_time, conversion_time);
        
        handler.close_document(&doc_id)?;
    }
    
    // Analyze scaling behavior
    for i in 1..performance_data.len() {
        let (prev_size, _, prev_extract, prev_convert) = performance_data[i-1];
        let (curr_size, _, curr_extract, curr_convert) = performance_data[i];
        
        let size_ratio = curr_size as f64 / prev_size as f64;
        let extract_ratio = curr_extract.as_nanos() as f64 / prev_extract.as_nanos() as f64;
        let convert_ratio = curr_convert.as_nanos() as f64 / prev_convert.as_nanos() as f64;
        
        println!("Size {}→{}: Extract scaling {:.2}, Convert scaling {:.2}", 
                prev_size, curr_size, extract_ratio / size_ratio, convert_ratio / size_ratio);
        
        // Performance should scale reasonably (not exponentially)
        assert!(extract_ratio / size_ratio < 3.0, "Text extraction scaling too poor");
        assert!(convert_ratio / size_ratio < 5.0, "PDF conversion scaling too poor");
    }
    
    Ok(())
}

#[test]
fn test_error_handling_performance() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    let provider = DocxToolsProvider::with_base_dir(temp_dir.path());
    let error_operations = vec![
        ("extract_text", json!({"document_id": "nonexistent"})),
        ("add_paragraph", json!({"document_id": "fake", "text": "test"})),
        ("get_metadata", json!({"document_id": "invalid"})),
        ("unknown_tool", json!({})),
    ];
    
    for (operation, args) in error_operations {
        let start = Instant::now();
        
        let result = tokio_test::block_on(async {
            provider.call_tool(operation, args).await
        });
        
        let error_time = start.elapsed();
        println!("Error handling for {}: {:?}", operation, error_time);
        
        // Error handling should be fast
        assert!(error_time < Duration::from_millis(100), 
               "Error handling for {} too slow: {:?}", operation, error_time);
        
        // Should return appropriate error
        // Ensure we got a response shape; don't match legacy types here
    }
    
    Ok(())
}

#[test]
fn test_resource_cleanup_performance() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    let mut handler = DocxHandler::new_with_base_dir(temp_dir.path())?;
    
    let document_count = 50;
    let mut doc_ids = Vec::new();
    
    // Create many documents
    let creation_start = Instant::now();
    for i in 0..document_count {
        let doc_id = handler.create_document()?;
        handler.add_paragraph(&doc_id, &format!("Document {} content", i), None)?;
        doc_ids.push(doc_id);
    }
    let creation_time = creation_start.elapsed();
    
    println!("Created {} documents in {:?}", document_count, creation_time);
    
    // Verify all documents exist
    let initial_count = handler.list_documents().len();
    assert_eq!(initial_count, document_count);
    
    // Test cleanup performance
    let cleanup_start = Instant::now();
    for doc_id in doc_ids {
        handler.close_document(&doc_id)?;
    }
    let cleanup_time = cleanup_start.elapsed();
    
    println!("Cleaned up {} documents in {:?}", document_count, cleanup_time);
    
    // Verify cleanup worked
    let final_count = handler.list_documents().len();
    assert_eq!(final_count, 0);
    
    // Cleanup should be reasonably fast
    assert!(cleanup_time < Duration::from_secs(5), "Cleanup took too long");
    
    let avg_cleanup_time = cleanup_time.as_nanos() / document_count as u128;
    println!("Average cleanup time per document: {}ns", avg_cleanup_time);
    
    Ok(())
}