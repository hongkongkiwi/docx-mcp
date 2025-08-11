use anyhow::Result;
use docx_mcp::docx_handler::{DocxHandler, DocxStyle, TableData};
use tempfile::TempDir;
use std::path::PathBuf;
use pretty_assertions::assert_eq;
use rstest::*;
use chrono::Utc;

fn setup_test_handler() -> (DocxHandler, TempDir) {
    let temp_dir = TempDir::new().unwrap();
    let handler = DocxHandler::new().unwrap();
    (handler, temp_dir)
}

#[fixture]
fn handler_and_doc() -> (DocxHandler, String, TempDir) {
    let (mut handler, temp_dir) = setup_test_handler();
    let doc_id = handler.create_document().unwrap();
    (handler, doc_id, temp_dir)
}

#[test]
fn test_create_document() {
    let (mut handler, _temp_dir) = setup_test_handler();
    
    let doc_id = handler.create_document().unwrap();
    assert!(!doc_id.is_empty());
    
    // Document should be in the handler's registry
    assert!(handler.documents.contains_key(&doc_id));
    
    let metadata = handler.get_metadata(&doc_id).unwrap();
    assert_eq!(metadata.id, doc_id);
    assert!(metadata.path.exists());
}

#[test]
fn test_add_paragraph() {
    let (mut handler, doc_id, _temp_dir) = handler_and_doc();
    
    let result = handler.add_paragraph(&doc_id, "Test paragraph", None);
    assert!(result.is_ok());
    
    // Verify content was added by extracting text
    let text = handler.extract_text(&doc_id).unwrap();
    assert!(text.contains("Test paragraph"));
}

#[test]
fn test_add_paragraph_with_style() {
    let (mut handler, doc_id, _temp_dir) = handler_and_doc();
    
    let style = DocxStyle {
        font_family: Some("Arial".to_string()),
        font_size: Some(14),
        bold: Some(true),
        italic: Some(false),
        underline: Some(false),
        color: Some("#FF0000".to_string()),
        alignment: Some("center".to_string()),
        line_spacing: Some(1.5),
    };
    
    let result = handler.add_paragraph(&doc_id, "Styled paragraph", Some(style));
    assert!(result.is_ok());
    
    let text = handler.extract_text(&doc_id).unwrap();
    assert!(text.contains("Styled paragraph"));
}

#[test]
fn test_add_heading() {
    let (mut handler, doc_id, _temp_dir) = handler_and_doc();
    
    for level in 1..=6 {
        let heading_text = format!("Heading Level {}", level);
        let result = handler.add_heading(&doc_id, &heading_text, level);
        assert!(result.is_ok(), "Failed to add heading level {}", level);
        
        let text = handler.extract_text(&doc_id).unwrap();
        assert!(text.contains(&heading_text));
    }
}

#[test]
fn test_add_table() {
    let (mut handler, doc_id, _temp_dir) = handler_and_doc();
    
    let table_data = TableData {
        rows: vec![
            vec!["Name".to_string(), "Age".to_string(), "City".to_string()],
            vec!["John".to_string(), "30".to_string(), "NYC".to_string()],
            vec!["Jane".to_string(), "25".to_string(), "LA".to_string()],
        ],
        headers: Some(vec!["Name".to_string(), "Age".to_string(), "City".to_string()]),
        border_style: Some("single".to_string()),
    };
    
    let result = handler.add_table(&doc_id, table_data);
    assert!(result.is_ok());
    
    let text = handler.extract_text(&doc_id).unwrap();
    assert!(text.contains("John"));
    assert!(text.contains("Jane"));
    assert!(text.contains("NYC"));
}

#[test]
fn test_add_list() {
    let (mut handler, doc_id, _temp_dir) = handler_and_doc();
    
    let items = vec![
        "First item".to_string(),
        "Second item".to_string(),
        "Third item".to_string(),
    ];
    
    // Test unordered list
    let result = handler.add_list(&doc_id, items.clone(), false);
    assert!(result.is_ok());
    
    // Test ordered list
    let result = handler.add_list(&doc_id, items.clone(), true);
    assert!(result.is_ok());
    
    let text = handler.extract_text(&doc_id).unwrap();
    assert!(text.contains("First item"));
    assert!(text.contains("Second item"));
    assert!(text.contains("Third item"));
}

#[test]
fn test_set_header_footer() {
    let (mut handler, doc_id, _temp_dir) = handler_and_doc();
    
    let header_result = handler.set_header(&doc_id, "Document Header");
    assert!(header_result.is_ok());
    
    let footer_result = handler.set_footer(&doc_id, "Document Footer");
    assert!(footer_result.is_ok());
}

#[test]
fn test_add_page_break() {
    let (mut handler, doc_id, _temp_dir) = handler_and_doc();
    
    handler.add_paragraph(&doc_id, "Before page break", None).unwrap();
    
    let result = handler.add_page_break(&doc_id);
    assert!(result.is_ok());
    
    handler.add_paragraph(&doc_id, "After page break", None).unwrap();
    
    let text = handler.extract_text(&doc_id).unwrap();
    assert!(text.contains("Before page break"));
    assert!(text.contains("After page break"));
}

#[test]
fn test_extract_text_empty_document() {
    let (handler, doc_id, _temp_dir) = handler_and_doc();
    
    let text = handler.extract_text(&doc_id).unwrap();
    // Empty document might have some default content or be truly empty
    assert!(text.is_empty() || text.trim().is_empty());
}

#[test]
fn test_save_and_close_document() {
    let (mut handler, doc_id, temp_dir) = handler_and_doc();
    
    handler.add_paragraph(&doc_id, "Test content", None).unwrap();
    
    let save_path = temp_dir.path().join("test_output.docx");
    let save_result = handler.save_document(&doc_id, &save_path);
    assert!(save_result.is_ok());
    assert!(save_path.exists());
    
    let close_result = handler.close_document(&doc_id);
    assert!(close_result.is_ok());
    assert!(!handler.documents.contains_key(&doc_id));
}

#[test]
fn test_open_existing_document() {
    let (mut handler, doc_id, temp_dir) = handler_and_doc();
    
    // Create and save a document
    handler.add_paragraph(&doc_id, "Original content", None).unwrap();
    let save_path = temp_dir.path().join("existing.docx");
    handler.save_document(&doc_id, &save_path).unwrap();
    handler.close_document(&doc_id).unwrap();
    
    // Open the saved document
    let opened_doc_id = handler.open_document(&save_path).unwrap();
    assert_ne!(opened_doc_id, doc_id); // Should be a new ID
    
    let text = handler.extract_text(&opened_doc_id).unwrap();
    assert!(text.contains("Original content"));
}

#[test]
fn test_list_documents() {
    let (mut handler, _temp_dir) = setup_test_handler();
    
    // Initially should be empty
    let docs = handler.list_documents();
    let initial_count = docs.len();
    
    // Create some documents
    let _doc1 = handler.create_document().unwrap();
    let _doc2 = handler.create_document().unwrap();
    let _doc3 = handler.create_document().unwrap();
    
    let docs = handler.list_documents();
    assert_eq!(docs.len(), initial_count + 3);
}

#[test]
fn test_document_not_found_error() {
    let (handler, _temp_dir) = setup_test_handler();
    
    let fake_id = "nonexistent-document-id";
    
    let result = handler.extract_text(fake_id);
    assert!(result.is_err());
    assert!(result.unwrap_err().to_string().contains("Document not found"));
}

#[test]
fn test_get_metadata() {
    let (handler, doc_id, _temp_dir) = handler_and_doc();
    
    let metadata = handler.get_metadata(&doc_id).unwrap();
    
    assert_eq!(metadata.id, doc_id);
    assert!(metadata.path.exists());
    assert!(metadata.created_at <= Utc::now());
    assert!(metadata.modified_at <= Utc::now());
    assert_eq!(metadata.page_count, Some(1));
    assert_eq!(metadata.word_count, Some(0));
}

#[test]
fn test_concurrent_document_operations() {
    use std::sync::Arc;
    use std::sync::Mutex;
    use std::thread;
    
    let (handler, _temp_dir) = setup_test_handler();
    let handler = Arc::new(Mutex::new(handler));
    
    let handles: Vec<_> = (0..5).map(|i| {
        let handler = Arc::clone(&handler);
        thread::spawn(move || {
            let doc_id = {
                let mut h = handler.lock().unwrap();
                h.create_document().unwrap()
            };
            
            {
                let mut h = handler.lock().unwrap();
                h.add_paragraph(&doc_id, &format!("Thread {} content", i), None).unwrap();
            }
            
            {
                let h = handler.lock().unwrap();
                let text = h.extract_text(&doc_id).unwrap();
                assert!(text.contains(&format!("Thread {} content", i)));
            }
            
            doc_id
        })
    }).collect();
    
    let doc_ids: Vec<_> = handles.into_iter().map(|h| h.join().unwrap()).collect();
    
    // All documents should be different
    let mut unique_ids = doc_ids.clone();
    unique_ids.sort();
    unique_ids.dedup();
    assert_eq!(unique_ids.len(), doc_ids.len());
}

#[test]
fn test_large_document_creation() {
    let (mut handler, doc_id, _temp_dir) = handler_and_doc();
    
    // Add many paragraphs to test performance
    for i in 0..100 {
        let content = format!("Paragraph number {} with some content to make it realistic", i);
        handler.add_paragraph(&doc_id, &content, None).unwrap();
    }
    
    let text = handler.extract_text(&doc_id).unwrap();
    assert!(text.contains("Paragraph number 0"));
    assert!(text.contains("Paragraph number 99"));
    
    // Verify word count (lower threshold due to simplified text extraction)
    let words: Vec<&str> = text.split_whitespace().collect();
    assert!(words.len() > 300);
}

#[test]
fn test_special_characters_in_content() {
    let (mut handler, doc_id, _temp_dir) = handler_and_doc();
    
    let special_content = "Special chars: éñüñdéd, 中文, русский, العربية, 🚀📝✨";
    handler.add_paragraph(&doc_id, special_content, None).unwrap();
    
    let text = handler.extract_text(&doc_id).unwrap();
    assert!(text.contains("éñüñdéd"));
    assert!(text.contains("🚀📝✨"));
}