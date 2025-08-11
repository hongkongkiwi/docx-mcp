//! Test fixtures and helper data for the docx-mcp test suite

use anyhow::Result;
use docx_mcp::docx_handler::{DocxHandler, DocxStyle, TableData};
use serde_json::{json, Value};
use std::collections::HashMap;
use tempfile::TempDir;

pub mod sample_documents;
pub mod test_data;

/// Common test fixture for creating a handler with a temporary directory
pub fn create_test_handler() -> (DocxHandler, TempDir) {
    let temp_dir = TempDir::new().unwrap();
    let handler = DocxHandler::new_with_temp_dir(temp_dir.path()).unwrap();
    (handler, temp_dir)
}

/// Create a handler with a document containing basic content
pub fn create_handler_with_document() -> (DocxHandler, String, TempDir) {
    let (mut handler, temp_dir) = create_test_handler();
    let doc_id = handler.create_document().unwrap();
    (handler, doc_id, temp_dir)
}

/// Standard document styles for testing
pub struct TestStyles;

impl TestStyles {
    pub fn basic() -> DocxStyle {
        DocxStyle {
            font_family: Some("Calibri".to_string()),
            font_size: Some(11),
            bold: Some(false),
            italic: Some(false),
            underline: Some(false),
            color: Some("#000000".to_string()),
            alignment: Some("left".to_string()),
            line_spacing: Some(1.15),
        }
    }
    
    pub fn heading() -> DocxStyle {
        DocxStyle {
            font_family: Some("Calibri".to_string()),
            font_size: Some(16),
            bold: Some(true),
            italic: Some(false),
            underline: Some(false),
            color: Some("#1f4e79".to_string()),
            alignment: Some("left".to_string()),
            line_spacing: Some(1.15),
        }
    }
    
    pub fn emphasis() -> DocxStyle {
        DocxStyle {
            font_family: Some("Calibri".to_string()),
            font_size: Some(11),
            bold: Some(true),
            italic: Some(true),
            underline: Some(false),
            color: Some("#c55a11".to_string()),
            alignment: Some("left".to_string()),
            line_spacing: Some(1.15),
        }
    }
    
    pub fn centered() -> DocxStyle {
        DocxStyle {
            font_family: Some("Calibri".to_string()),
            font_size: Some(11),
            bold: Some(false),
            italic: Some(false),
            underline: Some(false),
            color: Some("#000000".to_string()),
            alignment: Some("center".to_string()),
            line_spacing: Some(1.15),
        }
    }
}

/// Standard table data for testing
pub struct TestTables;

impl TestTables {
    pub fn simple_2x2() -> TableData {
        TableData {
            rows: vec![
                vec!["Row 1 Col 1".to_string(), "Row 1 Col 2".to_string()],
                vec!["Row 2 Col 1".to_string(), "Row 2 Col 2".to_string()],
            ],
            headers: None,
            border_style: Some("single".to_string()),
        }
    }
    
    pub fn with_headers() -> TableData {
        TableData {
            rows: vec![
                vec!["Name".to_string(), "Age".to_string(), "City".to_string()],
                vec!["John".to_string(), "30".to_string(), "New York".to_string()],
                vec!["Jane".to_string(), "25".to_string(), "Los Angeles".to_string()],
                vec!["Bob".to_string(), "35".to_string(), "Chicago".to_string()],
            ],
            headers: Some(vec!["Name".to_string(), "Age".to_string(), "City".to_string()]),
            border_style: Some("single".to_string()),
        }
    }
    
    pub fn financial_data() -> TableData {
        TableData {
            rows: vec![
                vec!["Quarter".to_string(), "Revenue".to_string(), "Profit".to_string(), "Growth".to_string()],
                vec!["Q1 2024".to_string(), "$1.2M".to_string(), "$240K".to_string(), "15%".to_string()],
                vec!["Q2 2024".to_string(), "$1.4M".to_string(), "$290K".to_string(), "18%".to_string()],
                vec!["Q3 2024".to_string(), "$1.6M".to_string(), "$340K".to_string(), "22%".to_string()],
                vec!["Q4 2024".to_string(), "$1.8M".to_string(), "$380K".to_string(), "25%".to_string()],
            ],
            headers: Some(vec!["Quarter".to_string(), "Revenue".to_string(), "Profit".to_string(), "Growth".to_string()]),
            border_style: Some("single".to_string()),
        }
    }
    
    pub fn large_table(rows: usize, cols: usize) -> TableData {
        let mut table_rows = Vec::new();
        
        // Header row
        let header_row: Vec<String> = (0..cols)
            .map(|i| format!("Column {}", i + 1))
            .collect();
        table_rows.push(header_row.clone());
        
        // Data rows
        for row in 0..rows {
            let data_row: Vec<String> = (0..cols)
                .map(|col| format!("R{}C{}", row + 1, col + 1))
                .collect();
            table_rows.push(data_row);
        }
        
        TableData {
            rows: table_rows,
            headers: Some(header_row),
            border_style: Some("single".to_string()),
        }
    }
}

/// Standard list data for testing
pub struct TestLists;

impl TestLists {
    pub fn simple_bullets() -> Vec<String> {
        vec![
            "First bullet point".to_string(),
            "Second bullet point".to_string(),
            "Third bullet point".to_string(),
        ]
    }
    
    pub fn numbered_steps() -> Vec<String> {
        vec![
            "Open the application".to_string(),
            "Navigate to the settings menu".to_string(),
            "Select the desired configuration".to_string(),
            "Save your changes".to_string(),
            "Restart the application".to_string(),
        ]
    }
    
    pub fn features_list() -> Vec<String> {
        vec![
            "Advanced document editing capabilities".to_string(),
            "Real-time collaboration tools".to_string(),
            "Cloud synchronization".to_string(),
            "Version control and history tracking".to_string(),
            "Export to multiple formats (PDF, HTML, Markdown)".to_string(),
            "Template library with professional designs".to_string(),
            "Advanced formatting and styling options".to_string(),
        ]
    }
    
    pub fn technical_requirements() -> Vec<String> {
        vec![
            "Rust 1.70 or higher".to_string(),
            "Memory: 2GB RAM minimum, 4GB recommended".to_string(),
            "Storage: 500MB available space".to_string(),
            "Network: Internet connection for cloud features".to_string(),
            "OS: Windows 10, macOS 10.15, or Linux (Ubuntu 20.04+)".to_string(),
        ]
    }
    
    pub fn large_list(item_count: usize) -> Vec<String> {
        (1..=item_count)
            .map(|i| format!("List item number {} with descriptive content", i))
            .collect()
    }
}

/// Sample text content for testing
pub struct TestContent;

impl TestContent {
    pub fn lorem_ipsum() -> &'static str {
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum."
    }
    
    pub fn technical_paragraph() -> &'static str {
        "This application leverages cutting-edge Rust technology to provide high-performance document processing capabilities. The architecture is built on modern asynchronous programming patterns, ensuring efficient resource utilization and scalability. Key features include memory-safe operations, zero-cost abstractions, and excellent concurrent processing performance."
    }
    
    pub fn business_paragraph() -> &'static str {
        "Our comprehensive business solution addresses the evolving needs of modern enterprises through innovative technology and streamlined workflows. With a focus on productivity enhancement and cost reduction, this platform delivers measurable value across multiple departments and use cases. The solution integrates seamlessly with existing infrastructure while providing robust security and compliance features."
    }
    
    pub fn multilingual_content() -> Vec<(&'static str, &'static str)> {
        vec![
            ("English", "The quick brown fox jumps over the lazy dog."),
            ("Spanish", "El zorro marrón rápido salta sobre el perro perezoso."),
            ("French", "Le renard brun rapide saute par-dessus le chien paresseux."),
            ("German", "Der schnelle braune Fuchs springt über den faulen Hund."),
            ("Italian", "La volpe marrone veloce salta sopra il cane pigro."),
            ("Portuguese", "A raposa marrom rápida pula sobre o cão preguiçoso."),
            ("Japanese", "素早い茶色のキツネは怠惰な犬を飛び越える。"),
            ("Chinese", "敏捷的棕色狐狸跳过懒狗。"),
            ("Korean", "빠른 갈색 여우가 게으른 개를 뛰어넘는다."),
            ("Russian", "Быстрая коричневая лиса прыгает через ленивую собаку."),
        ]
    }
    
    pub fn special_characters() -> &'static str {
        "Special characters test: àáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ ĀāĂăĄąĆćĈĉĊċČčĎďĐđĒēĔĕĖėĘęĚě"
    }
    
    pub fn symbols_and_math() -> &'static str {
        "Mathematical symbols: ∑ ∏ ∫ √ ≤ ≥ ≠ ± ∞ ∂ ∇ Ω α β γ δ ε θ λ μ π σ φ ψ ω"
    }
    
    pub fn long_paragraph(sentence_count: usize) -> String {
        let sentences = vec![
            "This is a comprehensive test of document processing capabilities.",
            "The system handles various types of content efficiently and accurately.",
            "Performance optimization ensures smooth operation even with large documents.",
            "Advanced formatting features provide professional document appearance.",
            "Error handling mechanisms maintain system stability under all conditions.",
            "Security features protect sensitive information throughout the process.",
            "Integration capabilities allow seamless workflow with existing systems.",
            "User-friendly interfaces make complex operations simple and intuitive.",
            "Scalable architecture supports growing business requirements.",
            "Continuous improvements ensure the solution remains cutting-edge.",
        ];
        
        let mut result = String::new();
        for i in 0..sentence_count {
            let sentence = sentences[i % sentences.len()];
            result.push_str(sentence);
            if i < sentence_count - 1 {
                result.push(' ');
            }
        }
        result
    }
}

/// MCP tool call arguments for testing
pub struct TestMcpArgs;

impl TestMcpArgs {
    pub fn create_document() -> Value {
        json!({})
    }
    
    pub fn add_paragraph(doc_id: &str, text: &str, style: Option<DocxStyle>) -> Value {
        let mut args = json!({
            "document_id": doc_id,
            "text": text
        });
        
        if let Some(s) = style {
            args["style"] = json!({
                "font_family": s.font_family,
                "font_size": s.font_size,
                "bold": s.bold,
                "italic": s.italic,
                "underline": s.underline,
                "color": s.color,
                "alignment": s.alignment,
                "line_spacing": s.line_spacing
            });
        }
        
        args
    }
    
    pub fn add_heading(doc_id: &str, text: &str, level: usize) -> Value {
        json!({
            "document_id": doc_id,
            "text": text,
            "level": level
        })
    }
    
    pub fn add_table(doc_id: &str, table_data: &TableData) -> Value {
        json!({
            "document_id": doc_id,
            "rows": table_data.rows
        })
    }
    
    pub fn add_list(doc_id: &str, items: &[String], ordered: bool) -> Value {
        json!({
            "document_id": doc_id,
            "items": items,
            "ordered": ordered
        })
    }
    
    pub fn extract_text(doc_id: &str) -> Value {
        json!({
            "document_id": doc_id
        })
    }
    
    pub fn search_text(doc_id: &str, search_term: &str, case_sensitive: bool) -> Value {
        json!({
            "document_id": doc_id,
            "search_term": search_term,
            "case_sensitive": case_sensitive
        })
    }
    
    pub fn get_metadata(doc_id: &str) -> Value {
        json!({
            "document_id": doc_id
        })
    }
    
    pub fn convert_to_pdf(doc_id: &str, output_path: &str) -> Value {
        json!({
            "document_id": doc_id,
            "output_path": output_path
        })
    }
    
    pub fn save_document(doc_id: &str, output_path: &str) -> Value {
        json!({
            "document_id": doc_id,
            "output_path": output_path
        })
    }
}

/// Performance test data generators
pub struct PerformanceData;

impl PerformanceData {
    pub fn create_large_document(handler: &mut DocxHandler, paragraph_count: usize) -> Result<String> {
        let doc_id = handler.create_document()?;
        
        handler.add_heading(&doc_id, "Performance Test Document", 1)?;
        
        for i in 0..paragraph_count {
            if i % 50 == 0 && i > 0 {
                handler.add_heading(&doc_id, &format!("Section {}", i / 50), 2)?;
            }
            
            let content = format!(
                "This is paragraph {} in our performance test document. It contains substantial text content to simulate real-world usage patterns and test system performance under realistic load conditions. The paragraph includes various punctuation marks, numbers like {}, and other elements that affect processing performance.",
                i + 1, (i + 1) * 7
            );
            
            handler.add_paragraph(&doc_id, &content, None)?;
            
            // Add tables periodically
            if i % 100 == 99 {
                let table_data = TestTables::simple_2x2();
                handler.add_table(&doc_id, table_data)?;
            }
        }
        
        Ok(doc_id)
    }
    
    pub fn create_complex_document(handler: &mut DocxHandler) -> Result<String> {
        let doc_id = handler.create_document()?;
        
        // Add comprehensive content with all features
        handler.add_heading(&doc_id, "Complex Document Test", 1)?;
        
        handler.set_header(&doc_id, "Complex Document Header")?;
        handler.set_footer(&doc_id, "Complex Document Footer")?;
        
        handler.add_paragraph(&doc_id, TestContent::business_paragraph(), Some(TestStyles::basic()))?;
        
        handler.add_heading(&doc_id, "Technical Details", 2)?;
        handler.add_paragraph(&doc_id, TestContent::technical_paragraph(), None)?;
        
        let features_list = TestLists::features_list();
        handler.add_list(&doc_id, features_list, false)?;
        
        handler.add_heading(&doc_id, "Financial Overview", 2)?;
        let financial_table = TestTables::financial_data();
        handler.add_table(&doc_id, financial_table)?;
        
        handler.add_page_break(&doc_id)?;
        
        handler.add_heading(&doc_id, "Multilingual Content", 2)?;
        for (language, text) in TestContent::multilingual_content() {
            handler.add_paragraph(&doc_id, &format!("{}: {}", language, text), None)?;
        }
        
        handler.add_heading(&doc_id, "Special Characters", 2)?;
        handler.add_paragraph(&doc_id, TestContent::special_characters(), None)?;
        handler.add_paragraph(&doc_id, TestContent::symbols_and_math(), None)?;
        
        Ok(doc_id)
    }
}

/// Error testing utilities
pub struct ErrorTestCases;

impl ErrorTestCases {
    pub fn invalid_document_ids() -> Vec<&'static str> {
        vec![
            "nonexistent-123",
            "fake-document-id",
            "invalid-uuid",
            "",
            "   ",
            "null",
            "undefined",
        ]
    }
    
    pub fn invalid_mcp_calls() -> Vec<(&'static str, Value)> {
        vec![
            ("add_paragraph", json!({"text": "missing document_id"})),
            ("add_heading", json!({"document_id": "test", "level": 10})),
            ("add_table", json!({"document_id": "test", "rows": "not_an_array"})),
            ("add_list", json!({"document_id": "test", "items": 123})),
            ("search_text", json!({"document_id": "test"})), // Missing search_term
            ("convert_to_pdf", json!({"document_id": "test"})), // Missing output_path
        ]
    }
    
    pub fn security_blocked_operations() -> Vec<(&'static str, Value)> {
        vec![
            ("create_document", json!({})),
            ("add_paragraph", json!({"document_id": "test", "text": "blocked"})),
            ("save_document", json!({"document_id": "test", "output_path": "/tmp/test.docx"})),
            ("convert_to_pdf", json!({"document_id": "test", "output_path": "/tmp/test.pdf"})),
            ("find_and_replace", json!({"document_id": "test", "find_text": "a", "replace_text": "b"})),
        ]
    }
}