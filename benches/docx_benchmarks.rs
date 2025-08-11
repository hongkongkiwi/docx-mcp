use criterion::{black_box, criterion_group, criterion_main, Criterion, BenchmarkId};
use docx_mcp::docx_handler::{DocxHandler, DocxStyle, TableData};
use docx_mcp::pure_converter::PureRustConverter;
use tempfile::TempDir;
use std::time::Duration;

fn setup_handler() -> (DocxHandler, TempDir) {
    let temp_dir = TempDir::new().unwrap();
    let handler = DocxHandler::new_with_temp_dir(temp_dir.path()).unwrap();
    (handler, temp_dir)
}

fn bench_document_creation(c: &mut Criterion) {
    c.bench_function("create_document", |b| {
        b.iter_batched(
            || setup_handler(),
            |(mut handler, _temp_dir)| {
                black_box(handler.create_document().unwrap())
            },
            criterion::BatchSize::LargeInput,
        )
    });
}

fn bench_paragraph_addition(c: &mut Criterion) {
    let mut group = c.benchmark_group("add_paragraph");
    
    for paragraph_count in [1, 10, 100, 1000].iter() {
        group.bench_with_input(
            BenchmarkId::new("count", paragraph_count),
            paragraph_count,
            |b, &count| {
                b.iter_batched(
                    || {
                        let (mut handler, temp_dir) = setup_handler();
                        let doc_id = handler.create_document().unwrap();
                        (handler, doc_id, temp_dir)
                    },
                    |(mut handler, doc_id, _temp_dir)| {
                        for i in 0..count {
                            let text = format!("This is paragraph number {} with some content", i);
                            handler.add_paragraph(&doc_id, &text, None).unwrap();
                        }
                        black_box(doc_id)
                    },
                    criterion::BatchSize::LargeInput,
                )
            },
        );
    }
    group.finish();
}

fn bench_styled_paragraph_addition(c: &mut Criterion) {
    c.bench_function("add_styled_paragraph", |b| {
        b.iter_batched(
            || {
                let (mut handler, temp_dir) = setup_handler();
                let doc_id = handler.create_document().unwrap();
                let style = DocxStyle {
                    font_family: Some("Arial".to_string()),
                    font_size: Some(12),
                    bold: Some(true),
                    italic: Some(false),
                    underline: Some(false),
                    color: Some("#000000".to_string()),
                    alignment: Some("left".to_string()),
                    line_spacing: Some(1.0),
                };
                (handler, doc_id, temp_dir, style)
            },
            |(mut handler, doc_id, _temp_dir, style)| {
                black_box(handler.add_paragraph(&doc_id, "Styled paragraph", Some(style)).unwrap())
            },
            criterion::BatchSize::LargeInput,
        )
    });
}

fn bench_heading_addition(c: &mut Criterion) {
    let mut group = c.benchmark_group("add_heading");
    
    for level in 1..=6 {
        group.bench_with_input(
            BenchmarkId::new("level", level),
            &level,
            |b, &level| {
                b.iter_batched(
                    || {
                        let (mut handler, temp_dir) = setup_handler();
                        let doc_id = handler.create_document().unwrap();
                        (handler, doc_id, temp_dir)
                    },
                    |(mut handler, doc_id, _temp_dir)| {
                        black_box(handler.add_heading(&doc_id, &format!("Heading Level {}", level), level).unwrap())
                    },
                    criterion::BatchSize::LargeInput,
                )
            },
        );
    }
    group.finish();
}

fn bench_table_addition(c: &mut Criterion) {
    let mut group = c.benchmark_group("add_table");
    
    for size in [(2, 2), (5, 5), (10, 10), (20, 10)].iter() {
        group.bench_with_input(
            BenchmarkId::new("size", format!("{}x{}", size.0, size.1)),
            size,
            |b, &(rows, cols)| {
                b.iter_batched(
                    || {
                        let (mut handler, temp_dir) = setup_handler();
                        let doc_id = handler.create_document().unwrap();
                        
                        let mut table_rows = Vec::new();
                        for i in 0..rows {
                            let mut row = Vec::new();
                            for j in 0..cols {
                                row.push(format!("Cell {}x{}", i, j));
                            }
                            table_rows.push(row);
                        }
                        
                        let table_data = TableData {
                            rows: table_rows,
                            headers: None,
                            border_style: Some("single".to_string()),
                        };
                        
                        (handler, doc_id, temp_dir, table_data)
                    },
                    |(mut handler, doc_id, _temp_dir, table_data)| {
                        black_box(handler.add_table(&doc_id, table_data).unwrap())
                    },
                    criterion::BatchSize::LargeInput,
                )
            },
        );
    }
    group.finish();
}

fn bench_list_addition(c: &mut Criterion) {
    let mut group = c.benchmark_group("add_list");
    
    for item_count in [5, 20, 50, 100].iter() {
        group.bench_with_input(
            BenchmarkId::new("items", item_count),
            item_count,
            |b, &count| {
                b.iter_batched(
                    || {
                        let (mut handler, temp_dir) = setup_handler();
                        let doc_id = handler.create_document().unwrap();
                        
                        let items: Vec<String> = (0..count)
                            .map(|i| format!("List item number {}", i))
                            .collect();
                        
                        (handler, doc_id, temp_dir, items)
                    },
                    |(mut handler, doc_id, _temp_dir, items)| {
                        black_box(handler.add_list(&doc_id, items, false).unwrap())
                    },
                    criterion::BatchSize::LargeInput,
                )
            },
        );
    }
    group.finish();
}

fn bench_text_extraction(c: &mut Criterion) {
    let mut group = c.benchmark_group("extract_text");
    
    for paragraph_count in [10, 100, 500, 1000].iter() {
        group.bench_with_input(
            BenchmarkId::new("paragraphs", paragraph_count),
            paragraph_count,
            |b, &count| {
                b.iter_batched(
                    || {
                        let (mut handler, temp_dir) = setup_handler();
                        let doc_id = handler.create_document().unwrap();
                        
                        // Create document with many paragraphs
                        for i in 0..count {
                            let text = format!("This is paragraph {} with substantial content to test text extraction performance. It includes various words and punctuation to make it realistic.", i);
                            handler.add_paragraph(&doc_id, &text, None).unwrap();
                        }
                        
                        (handler, doc_id, temp_dir)
                    },
                    |(handler, doc_id, _temp_dir)| {
                        black_box(handler.extract_text(&doc_id).unwrap())
                    },
                    criterion::BatchSize::LargeInput,
                )
            },
        );
    }
    group.finish();
}

fn bench_pdf_conversion(c: &mut Criterion) {
    let mut group = c.benchmark_group("pdf_conversion");
    group.measurement_time(Duration::from_secs(30)); // Longer measurement for PDF conversion
    
    for paragraph_count in [10, 50, 200].iter() {
        group.bench_with_input(
            BenchmarkId::new("paragraphs", paragraph_count),
            paragraph_count,
            |b, &count| {
                b.iter_batched(
                    || {
                        let (mut handler, temp_dir) = setup_handler();
                        let doc_id = handler.create_document().unwrap();
                        
                        // Create substantial document content
                        handler.add_heading(&doc_id, "Performance Test Document", 1).unwrap();
                        
                        for i in 0..count {
                            if i % 20 == 0 {
                                handler.add_heading(&doc_id, &format!("Section {}", i / 20 + 1), 2).unwrap();
                            }
                            
                            let text = format!("This is paragraph {} designed to test PDF conversion performance. It contains enough text to make the conversion meaningful and test the system under realistic load conditions.", i);
                            handler.add_paragraph(&doc_id, &text, None).unwrap();
                        }
                        
                        let metadata = handler.get_metadata(&doc_id).unwrap();
                        let converter = PureRustConverter::new();
                        let output_path = temp_dir.path().join("benchmark.pdf");
                        
                        (metadata, converter, output_path, temp_dir)
                    },
                    |(metadata, converter, output_path, _temp_dir)| {
                        black_box(converter.convert_docx_to_pdf(&metadata.path, &output_path).unwrap())
                    },
                    criterion::BatchSize::LargeInput,
                )
            },
        );
    }
    group.finish();
}

fn bench_image_conversion(c: &mut Criterion) {
    let mut group = c.benchmark_group("image_conversion");
    group.measurement_time(Duration::from_secs(45)); // Even longer for image conversion
    
    for paragraph_count in [5, 20, 50].iter() {
        group.bench_with_input(
            BenchmarkId::new("paragraphs", paragraph_count),
            paragraph_count,
            |b, &count| {
                b.iter_batched(
                    || {
                        let (mut handler, temp_dir) = setup_handler();
                        let doc_id = handler.create_document().unwrap();
                        
                        handler.add_heading(&doc_id, "Image Conversion Test", 1).unwrap();
                        
                        for i in 0..count {
                            let text = format!("Paragraph {} for image conversion testing.", i);
                            handler.add_paragraph(&doc_id, &text, None).unwrap();
                        }
                        
                        let metadata = handler.get_metadata(&doc_id).unwrap();
                        let converter = PureRustConverter::new();
                        let output_dir = temp_dir.path().join("images");
                        std::fs::create_dir_all(&output_dir).unwrap();
                        
                        (metadata, converter, output_dir, temp_dir)
                    },
                    |(metadata, converter, output_dir, _temp_dir)| {
                        black_box(converter.convert_docx_to_images(&metadata.path, &output_dir).unwrap())
                    },
                    criterion::BatchSize::LargeInput,
                )
            },
        );
    }
    group.finish();
}

fn bench_concurrent_operations(c: &mut Criterion) {
    let mut group = c.benchmark_group("concurrent_operations");
    
    for thread_count in [2, 4, 8].iter() {
        group.bench_with_input(
            BenchmarkId::new("threads", thread_count),
            thread_count,
            |b, &threads| {
                b.iter_batched(
                    || {
                        let temp_dir = TempDir::new().unwrap();
                        (temp_dir, threads)
                    },
                    |(temp_dir, thread_count)| {
                        use std::sync::Arc;
                        use std::thread;
                        
                        let temp_path = Arc::new(temp_dir.path().to_path_buf());
                        
                        let handles: Vec<_> = (0..thread_count).map(|i| {
                            let temp_path = Arc::clone(&temp_path);
                            thread::spawn(move || {
                                let mut handler = DocxHandler::new_with_temp_dir(&temp_path).unwrap();
                                let doc_id = handler.create_document().unwrap();
                                
                                for j in 0..10 {
                                    let text = format!("Thread {} paragraph {}", i, j);
                                    handler.add_paragraph(&doc_id, &text, None).unwrap();
                                }
                                
                                handler.extract_text(&doc_id).unwrap()
                            })
                        }).collect();
                        
                        for handle in handles {
                            handle.join().unwrap();
                        }
                        
                        black_box(())
                    },
                    criterion::BatchSize::LargeInput,
                )
            },
        );
    }
    group.finish();
}

fn bench_memory_usage(c: &mut Criterion) {
    let mut group = c.benchmark_group("memory_usage");
    
    for doc_count in [5, 20, 50].iter() {
        group.bench_with_input(
            BenchmarkId::new("documents", doc_count),
            doc_count,
            |b, &count| {
                b.iter_batched(
                    || setup_handler(),
                    |(mut handler, _temp_dir)| {
                        let mut doc_ids = Vec::new();
                        
                        // Create multiple documents
                        for i in 0..count {
                            let doc_id = handler.create_document().unwrap();
                            
                            // Add content to each document
                            handler.add_heading(&doc_id, &format!("Document {}", i), 1).unwrap();
                            for j in 0..20 {
                                let text = format!("Content paragraph {} in document {}", j, i);
                                handler.add_paragraph(&doc_id, &text, None).unwrap();
                            }
                            
                            doc_ids.push(doc_id);
                        }
                        
                        // Extract text from all documents
                        for doc_id in &doc_ids {
                            handler.extract_text(doc_id).unwrap();
                        }
                        
                        black_box(doc_ids)
                    },
                    criterion::BatchSize::LargeInput,
                )
            },
        );
    }
    group.finish();
}

fn bench_complex_document_operations(c: &mut Criterion) {
    c.bench_function("complex_document", |b| {
        b.iter_batched(
            || setup_handler(),
            |(mut handler, _temp_dir)| {
                let doc_id = handler.create_document().unwrap();
                
                // Create a complex document with all features
                handler.add_heading(&doc_id, "Complex Document Test", 1).unwrap();
                handler.add_paragraph(&doc_id, "This is a comprehensive test document.", None).unwrap();
                
                // Add styled paragraph
                let style = DocxStyle {
                    font_size: Some(14),
                    bold: Some(true),
                    color: Some("#FF0000".to_string()),
                    alignment: Some("center".to_string()),
                    ..Default::default()
                };
                handler.add_paragraph(&doc_id, "Styled paragraph", Some(style)).unwrap();
                
                // Add table
                let table_data = TableData {
                    rows: vec![
                        vec!["Header 1".to_string(), "Header 2".to_string(), "Header 3".to_string()],
                        vec!["Row 1 Col 1".to_string(), "Row 1 Col 2".to_string(), "Row 1 Col 3".to_string()],
                        vec!["Row 2 Col 1".to_string(), "Row 2 Col 2".to_string(), "Row 2 Col 3".to_string()],
                    ],
                    headers: Some(vec!["Header 1".to_string(), "Header 2".to_string(), "Header 3".to_string()]),
                    border_style: Some("single".to_string()),
                };
                handler.add_table(&doc_id, table_data).unwrap();
                
                // Add list
                let items = vec![
                    "First item".to_string(),
                    "Second item".to_string(),
                    "Third item".to_string(),
                ];
                handler.add_list(&doc_id, items, true).unwrap();
                
                // Add page break and more content
                handler.add_page_break(&doc_id).unwrap();
                handler.add_heading(&doc_id, "Second Page", 1).unwrap();
                handler.add_paragraph(&doc_id, "Content on second page", None).unwrap();
                
                // Set header and footer
                handler.set_header(&doc_id, "Document Header").unwrap();
                handler.set_footer(&doc_id, "Document Footer").unwrap();
                
                // Extract all text
                let text = handler.extract_text(&doc_id).unwrap();
                
                black_box(text)
            },
            criterion::BatchSize::LargeInput,
        )
    });
}

criterion_group!(
    benches,
    bench_document_creation,
    bench_paragraph_addition,
    bench_styled_paragraph_addition,
    bench_heading_addition,
    bench_table_addition,
    bench_list_addition,
    bench_text_extraction,
    bench_pdf_conversion,
    bench_image_conversion,
    bench_concurrent_operations,
    bench_memory_usage,
    bench_complex_document_operations
);

criterion_main!(benches);