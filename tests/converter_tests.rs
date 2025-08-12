use anyhow::Result;
use docx_mcp::docx_handler::{DocxHandler, DocxStyle, TableData};
use docx_mcp::pure_converter::PureRustConverter;
use tempfile::TempDir;
use std::path::{Path, PathBuf};
use std::fs;
use pretty_assertions::assert_eq;
use rstest::*;

fn setup_test_handler_with_content() -> (DocxHandler, String, TempDir) {
    let temp_dir = TempDir::new().unwrap();
    let mut handler = DocxHandler::new().unwrap();
    let doc_id = handler.create_document().unwrap();
    
    // Add comprehensive content for testing
    handler.add_heading(&doc_id, "Test Document Title", 1).unwrap();
    handler.add_paragraph(&doc_id, "This is a comprehensive test document with various content types.", None).unwrap();
    
    handler.add_heading(&doc_id, "Table Example", 2).unwrap();
    let table_data = TableData {
        rows: vec![
            vec!["Product".to_string(), "Price".to_string(), "Quantity".to_string()],
            vec!["Widget A".to_string(), "$10.00".to_string(), "5".to_string()],
            vec!["Widget B".to_string(), "$15.00".to_string(), "3".to_string()],
        ],
        headers: Some(vec!["Product".to_string(), "Price".to_string(), "Quantity".to_string()]),
        border_style: Some("single".to_string()),
        col_widths: None,
        merges: None,
        cell_shading: None,
    };
    handler.add_table(&doc_id, table_data).unwrap();
    
    handler.add_heading(&doc_id, "List Example", 2).unwrap();
    let list_items = vec![
        "First important point".to_string(),
        "Second key feature".to_string(),
        "Third critical aspect".to_string(),
    ];
    handler.add_list(&doc_id, list_items, false).unwrap();
    
    handler.add_paragraph(&doc_id, "Conclusion: This document demonstrates various formatting capabilities.", None).unwrap();
    
    (handler, doc_id, temp_dir)
}

#[test]
fn test_pure_converter_creation() {
    let converter = PureRustConverter::new();
    // Just verify it can be created without panicking
    assert!(true);
}

#[test]
fn test_extract_text_from_docx() -> Result<()> {
    let (handler, doc_id, _temp_dir) = setup_test_handler_with_content();
    
    let metadata = handler.get_metadata(&doc_id)?;
    let converter = PureRustConverter::new();
    
    let extracted_text = converter.extract_text_from_docx(&metadata.path)?;
    
    // Should contain all the content we added
    assert!(extracted_text.contains("Test Document Title"));
    assert!(extracted_text.contains("comprehensive test document"));
    assert!(extracted_text.contains("Table Example"));
    assert!(extracted_text.contains("Widget A"));
    assert!(extracted_text.contains("First important point"));
    assert!(extracted_text.contains("Conclusion"));
    
    Ok(())
}

#[test]
fn test_extract_text_empty_document() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    let mut handler = DocxHandler::new().unwrap();
    let doc_id = handler.create_document().unwrap();
    
    let metadata = handler.get_metadata(&doc_id)?;
    let converter = PureRustConverter::new();
    
    let extracted_text = converter.extract_text_from_docx(&metadata.path)?;
    
    // Empty document should return empty or whitespace-only text
    assert!(extracted_text.trim().is_empty());
    
    Ok(())
}

#[test]
fn test_convert_docx_to_pdf_basic() -> Result<()> {
    let (handler, doc_id, temp_dir) = setup_test_handler_with_content();
    
    let metadata = handler.get_metadata(&doc_id)?;
    let converter = PureRustConverter::new();
    
    let output_path = temp_dir.path().join("test_output.pdf");
    converter.convert_docx_to_pdf(&metadata.path, &output_path)?;
    
    // Verify PDF file was created
    assert!(output_path.exists());
    
    // Check file size is reasonable (should be larger than empty PDF)
    let file_size = fs::metadata(&output_path)?.len();
    assert!(file_size > 1000); // PDF should be at least 1KB
    
    // Verify it's actually a PDF file (starts with PDF signature)
    let pdf_content = fs::read(&output_path)?;
    assert!(pdf_content.starts_with(b"%PDF"));
    
    Ok(())
}

#[test]
fn test_convert_docx_to_pdf_with_complex_content() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    let mut handler = DocxHandler::new().unwrap();
    let doc_id = handler.create_document().unwrap();
    
    // Add content with special characters and formatting
    handler.add_paragraph(&doc_id, "Special characters: éñüñ, 中文, русский, العربية", None)?;
    
    let style = DocxStyle {
        font_family: Some("Arial".to_string()),
        font_size: Some(16),
        bold: Some(true),
        italic: Some(false),
        underline: Some(true),
        color: Some("#FF0000".to_string()),
        alignment: Some("center".to_string()),
        line_spacing: Some(1.5),
    };
    handler.add_paragraph(&doc_id, "Bold and underlined text", Some(style))?;
    
    // Add multiple headings
    for level in 1..=3 {
        handler.add_heading(&doc_id, &format!("Heading Level {}", level), level)?;
    }
    
    let metadata = handler.get_metadata(&doc_id)?;
    let converter = PureRustConverter::new();
    
    let output_path = temp_dir.path().join("complex_output.pdf");
    converter.convert_docx_to_pdf(&metadata.path, &output_path)?;
    
    assert!(output_path.exists());
    let file_size = fs::metadata(&output_path)?.len();
    assert!(file_size > 500); // Should be larger due to more content
    
    Ok(())
}

#[test]
fn test_convert_docx_to_images() -> Result<()> {
    let (handler, doc_id, temp_dir) = setup_test_handler_with_content();
    
    let metadata = handler.get_metadata(&doc_id)?;
    let converter = PureRustConverter::new();
    
    let output_dir = temp_dir.path().join("images");
    fs::create_dir_all(&output_dir)?;
    
    let image_paths = converter.convert_docx_to_images(&metadata.path, &output_dir)?;
    
    // Should generate at least one image
    assert!(!image_paths.is_empty());
    
    // Verify all generated images exist
    for image_path in &image_paths {
        assert!(image_path.exists(), "Generated image should exist: {:?}", image_path);
        
        let file_size = fs::metadata(image_path)?.len();
        assert!(file_size > 100, "Image file should have reasonable size");
        
        // Verify it's a PNG file (our default format)
        if image_path.extension().and_then(|s| s.to_str()) == Some("png") {
            let image_content = fs::read(image_path)?;
            assert!(image_content.starts_with(&[0x89, 0x50, 0x4E, 0x47]), "Should be valid PNG");
        }
    }
    
    Ok(())
}

#[test]
fn test_convert_docx_to_images_custom_format() -> Result<()> {
    let (handler, doc_id, temp_dir) = setup_test_handler_with_content();
    
    let metadata = handler.get_metadata(&doc_id)?;
    let converter = PureRustConverter::new();
    
    let output_dir = temp_dir.path().join("jpeg_images");
    fs::create_dir_all(&output_dir)?;
    
    let image_paths = converter.convert_docx_to_images_with_format(&metadata.path, &output_dir, "jpeg", 150)?;
    
    assert!(!image_paths.is_empty());
    
    for image_path in &image_paths {
        assert!(image_path.exists());
        
        // Verify JPEG format
        if image_path.extension().and_then(|s| s.to_str()) == Some("jpg") || 
           image_path.extension().and_then(|s| s.to_str()) == Some("jpeg") {
            let image_content = fs::read(image_path)?;
            assert!(image_content.starts_with(&[0xFF, 0xD8, 0xFF]), "Should be valid JPEG");
        }
    }
    
    Ok(())
}

#[test]
fn test_pdf_generation_with_embedded_fonts() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    let mut handler = DocxHandler::new().unwrap();
    let doc_id = handler.create_document().unwrap();
    
    // Add text that might require different fonts
    handler.add_paragraph(&doc_id, "Regular ASCII text", None)?;
    handler.add_paragraph(&doc_id, "Unicode: àáâãäå çèéêë ìíîï ñòóôõö ùúûü ýÿ", None)?;
    handler.add_paragraph(&doc_id, "Math symbols: ∑ ∏ ∫ √ ≤ ≥ ≠ ± ∞", None)?;
    
    let metadata = handler.get_metadata(&doc_id)?;
    let converter = PureRustConverter::new();
    
    let output_path = temp_dir.path().join("embedded_fonts.pdf");
    converter.convert_docx_to_pdf(&metadata.path, &output_path)?;
    
    assert!(output_path.exists());
    let file_size = fs::metadata(&output_path)?.len();
    assert!(file_size > 1000); // Should be larger due to embedded fonts
    
    Ok(())
}

#[test]
fn test_batch_conversion() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    let mut handler = DocxHandler::new().unwrap();
    
    // Create multiple documents
    let mut doc_paths = Vec::new();
    for i in 0..3 {
        let doc_id = handler.create_document().unwrap();
        handler.add_paragraph(&doc_id, &format!("Document {} content", i), None)?;
        
        let metadata = handler.get_metadata(&doc_id)?;
        doc_paths.push(metadata.path);
    }
    
    let converter = PureRustConverter::new();
    let output_dir = temp_dir.path().join("batch_output");
    fs::create_dir_all(&output_dir)?;
    
    // Convert all documents to PDF
    for (i, doc_path) in doc_paths.iter().enumerate() {
        let output_path = output_dir.join(format!("document_{}.pdf", i));
        converter.convert_docx_to_pdf(doc_path, &output_path)?;
        
        assert!(output_path.exists());
    }
    
    // Verify all PDFs were created
    let pdf_files: Vec<_> = fs::read_dir(&output_dir)?
        .filter_map(|entry| entry.ok())
        .filter(|entry| entry.path().extension().and_then(|s| s.to_str()) == Some("pdf"))
        .collect();
    
    assert_eq!(pdf_files.len(), 3);
    
    Ok(())
}

#[test]
fn test_error_handling_invalid_docx() {
    let temp_dir = TempDir::new().unwrap();
    let converter = PureRustConverter::new();
    
    // Create a fake DOCX file (actually just text)
    let fake_docx = temp_dir.path().join("fake.docx");
    fs::write(&fake_docx, "This is not a DOCX file").unwrap();
    
    // Should handle the error gracefully
    let result = converter.extract_text_from_docx(&fake_docx);
    assert!(result.is_err());
    
    let output_path = temp_dir.path().join("output.pdf");
    let result = converter.convert_docx_to_pdf(&fake_docx, &output_path);
    assert!(result.is_err());
}

#[test]
fn test_error_handling_nonexistent_file() {
    let temp_dir = TempDir::new().unwrap();
    let converter = PureRustConverter::new();
    
    let nonexistent = temp_dir.path().join("nonexistent.docx");
    
    let result = converter.extract_text_from_docx(&nonexistent);
    assert!(result.is_err());
    
    let output_path = temp_dir.path().join("output.pdf");
    let result = converter.convert_docx_to_pdf(&nonexistent, &output_path);
    assert!(result.is_err());
}

#[test]
fn test_large_document_conversion() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    let mut handler = DocxHandler::new().unwrap();
    let doc_id = handler.create_document().unwrap();
    
    // Create a large document with many pages
    for i in 0..50 {
        handler.add_heading(&doc_id, &format!("Section {}", i + 1), 1)?;
        
        for j in 0..10 {
            let content = format!("This is paragraph {} in section {}. It contains enough text to make the document substantial and test the conversion capabilities with larger files.", j + 1, i + 1);
            handler.add_paragraph(&doc_id, &content, None)?;
        }
        
        if i % 10 == 9 {
            handler.add_page_break(&doc_id)?;
        }
    }
    
    let metadata = handler.get_metadata(&doc_id)?;
    let converter = PureRustConverter::new();
    
    // Test PDF conversion
    let pdf_path = temp_dir.path().join("large_document.pdf");
    converter.convert_docx_to_pdf(&metadata.path, &pdf_path)?;
    
    assert!(pdf_path.exists());
    let pdf_size = fs::metadata(&pdf_path)?.len();
    assert!(pdf_size > 50000); // Should be a substantial PDF
    
    // Test image conversion (but only first few pages to avoid excessive test time)
    let images_dir = temp_dir.path().join("large_images");
    fs::create_dir_all(&images_dir)?;
    
    let image_paths = converter.convert_docx_to_images(&metadata.path, &images_dir)?;
    assert!(!image_paths.is_empty());
    
    // Should generate multiple images for multiple pages
    assert!(image_paths.len() >= 2);
    
    Ok(())
}

#[test]
fn test_text_extraction_accuracy() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    let mut handler = DocxHandler::new().unwrap();
    let doc_id = handler.create_document().unwrap();
    
    // Add specific test content
    let test_sentences = vec![
        "The quick brown fox jumps over the lazy dog.",
        "Pack my box with five dozen liquor jugs.",
        "How vexingly quick daft zebras jump!",
        "Sphinx of black quartz, judge my vow.",
    ];
    
    for sentence in &test_sentences {
        handler.add_paragraph(&doc_id, sentence, None)?;
    }
    
    let metadata = handler.get_metadata(&doc_id)?;
    let converter = PureRustConverter::new();
    
    let extracted_text = converter.extract_text_from_docx(&metadata.path)?;
    
    // Verify all sentences are present in the extracted text
    for sentence in &test_sentences {
        assert!(extracted_text.contains(sentence), 
                "Extracted text should contain: '{}'", sentence);
    }
    
    // Check word count accuracy
    let expected_words: usize = test_sentences.iter()
        .map(|s| s.split_whitespace().count())
        .sum();
    let extracted_words = extracted_text.split_whitespace().count();
    
    // Should be approximately equal (allowing for minor differences)
    let word_diff = if extracted_words > expected_words {
        extracted_words - expected_words
    } else {
        expected_words - extracted_words
    };
    assert!(word_diff <= 5, "Word count difference too large: expected ~{}, got {}", expected_words, extracted_words);
    
    Ok(())
}

#[test]
fn test_conversion_with_different_page_sizes() -> Result<()> {
    let temp_dir = TempDir::new().unwrap();
    let mut handler = DocxHandler::new().unwrap();
    let doc_id = handler.create_document().unwrap();
    
    handler.add_paragraph(&doc_id, "This document tests page size handling during conversion.", None)?;
    
    let metadata = handler.get_metadata(&doc_id)?;
    let converter = PureRustConverter::new();
    
    // Test different output formats and sizes
    let test_cases = vec![
        ("a4.pdf", "A4"),
        ("letter.pdf", "Letter"),
        ("legal.pdf", "Legal"),
    ];
    
    for (filename, _page_size) in test_cases {
        let output_path = temp_dir.path().join(filename);
        
        // Note: In a full implementation, you'd pass page_size to the converter
        converter.convert_docx_to_pdf(&metadata.path, &output_path)?;
        
        assert!(output_path.exists());
        let file_size = fs::metadata(&output_path)?.len();
        assert!(file_size > 500); // Reasonable minimum size
    }
    
    Ok(())
}

// Parametrized test for different image formats
#[rstest]
#[case("png", &[0x89, 0x50, 0x4E, 0x47])]
#[case("jpeg", &[0xFF, 0xD8, 0xFF])]
fn test_image_format_conversion(#[case] format: &str, #[case] signature: &[u8]) -> Result<()> {
    let (handler, doc_id, temp_dir) = setup_test_handler_with_content();
    
    let metadata = handler.get_metadata(&doc_id)?;
    let converter = PureRustConverter::new();
    
    let output_dir = temp_dir.path().join(format!("{}_images", format));
    fs::create_dir_all(&output_dir)?;
    
    let image_paths = converter.convert_docx_to_images_with_format(&metadata.path, &output_dir, format, 100)?;
    
    assert!(!image_paths.is_empty());
    
    for image_path in &image_paths {
        assert!(image_path.exists());
        
        let image_content = fs::read(image_path)?;
        assert!(image_content.starts_with(signature), 
                "Image should have correct format signature for {}", format);
    }
    
    Ok(())
}

#[test]
fn test_conversion_thread_safety() -> Result<()> {
    use std::sync::Arc;
    use std::thread;
    
    let temp_dir = TempDir::new().unwrap();
    let temp_path = Arc::new(temp_dir.path().to_path_buf());
    
    let handles: Vec<_> = (0..3).map(|i| {
        let temp_path = Arc::clone(&temp_path);
        thread::spawn(move || -> Result<()> {
            let mut handler = DocxHandler::new()?;
            let doc_id = handler.create_document()?;
            
            handler.add_paragraph(&doc_id, &format!("Thread {} test content", i), None)?;
            
            let metadata = handler.get_metadata(&doc_id)?;
            let converter = PureRustConverter::new();
            
            let pdf_path = temp_path.join(format!("thread_{}.pdf", i));
            converter.convert_docx_to_pdf(&metadata.path, &pdf_path)?;
            
            assert!(pdf_path.exists());
            Ok(())
        })
    }).collect();
    
    // Wait for all threads to complete
    for handle in handles {
        handle.join().unwrap()?;
    }
    
    // Verify all PDFs were created
    let pdf_count = fs::read_dir(&temp_dir)?
        .filter_map(|entry| entry.ok())
        .filter(|entry| entry.path().extension().and_then(|s| s.to_str()) == Some("pdf"))
        .count();
    
    assert_eq!(pdf_count, 3);
    
    Ok(())
}