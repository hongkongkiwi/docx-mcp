use anyhow::{Context, Result};
use ::image::{DynamicImage, ImageFormat, Rgba, RgbaImage};
use printpdf::*;
use std::fs::{self, File};
use std::io::{BufWriter, Read};
use std::path::{Path, PathBuf};
use tempfile::NamedTempFile;
use tracing::{info};
use roxmltree;
use zip::ZipArchive;
use ::lopdf::{dictionary, Object};

pub struct PureRustConverter;

impl PureRustConverter {
    pub fn new() -> Self {
        Self
    }

    /// Extract text from DOCX using pure Rust XML parsing
    pub fn extract_text_from_docx(&self, docx_path: &Path) -> Result<String> {
        let file = File::open(docx_path)?;
        let mut archive = ZipArchive::new(file)?;
        
        // Find the main document XML
        let mut document_xml = String::new();
        
        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let name = file.name().to_string();
            
            if name == "word/document.xml" {
                let mut buf = Vec::new();
                file.read_to_end(&mut buf)?;
                document_xml = String::from_utf8_lossy(&buf).to_string();
                break;
            }
        }
        
        if document_xml.is_empty() {
            anyhow::bail!("No document.xml found in DOCX file");
        }
        
        // Parse XML and extract text with basic whitespace semantics
        let doc = roxmltree::Document::parse(&document_xml)?;
        let mut text = String::new();
        let mut last_char: Option<char> = None;

        for node in doc.descendants() {
            let name = node.tag_name().name();
            match name {
                // Paragraph boundary
                "p" => {
                    if !text.ends_with('\n') {
                        text.push('\n');
                        last_char = Some('\n');
                    }
                }
                // Text run
                "t" => {
                    if let Some(node_text) = node.text() {
                        // Preserve spaces if xml:space="preserve"
                        let preserve = node.attribute(("xml", "space")).map(|v| v == "preserve").unwrap_or(false);
                        let mut content = node_text.to_string();
                        if !preserve {
                            // Collapse internal newlines and excessive spaces
                            content = content.replace('\n', " ");
                        }
                        if !content.is_empty() {
                            // Insert a space if needed between words
                            if let Some(c) = last_char { if !c.is_whitespace() && !content.starts_with([' ', '\n', '\t']) { text.push(' '); } }
                            text.push_str(&content);
                            last_char = content.chars().rev().next();
                        }
                    }
                }
                // Line break
                "br" => {
                    text.push('\n');
                    last_char = Some('\n');
                }
                // Tab
                "tab" => {
                    text.push('\t');
                    last_char = Some('\t');
                }
                _ => {}
            }
        }

        Ok(text.trim().to_string())
    }

    /// Convert DOCX to PDF using pure Rust (no external dependencies)
    pub fn docx_to_pdf_pure(&self, docx_path: &Path, pdf_path: &Path) -> Result<()> {
        // Extract text from DOCX
        let text = self.extract_text_from_docx(docx_path)
            .with_context(|| format!("Failed to extract text from {:?}", docx_path))?;
        
        // Create PDF with extracted text
        self.create_pdf_from_text(&text, pdf_path)?;
        
        info!("Successfully converted DOCX to PDF using pure Rust");
        Ok(())
    }

    // Backward-compat wrapper names expected by tests
    pub fn convert_docx_to_pdf(&self, docx_path: &Path, pdf_path: &Path) -> Result<()> {
        self.docx_to_pdf_pure(docx_path, pdf_path)
    }

    pub fn convert_docx_to_images(&self, docx_path: &Path, output_dir: &Path) -> Result<Vec<PathBuf>> {
        self.docx_to_images_pure(docx_path, output_dir, ImageFormat::Png)
    }

    pub fn convert_docx_to_images_with_format(&self, docx_path: &Path, output_dir: &Path, format: &str, _dpi: u32) -> Result<Vec<PathBuf>> {
        let fmt = match format.to_lowercase().as_str() {
            "jpg" | "jpeg" => ImageFormat::Jpeg,
            _ => ImageFormat::Png,
        };
        self.docx_to_images_pure(docx_path, output_dir, fmt)
    }

    /// Create a PDF from text content
    pub fn create_pdf_from_text(&self, text: &str, pdf_path: &Path) -> Result<()> {
        let (doc, page1, layer1) = PdfDocument::new("Document", Mm(210.0), Mm(297.0), "Layer 1");
        let current_layer = doc.get_page(page1).get_layer(layer1);
        
        // Use embedded font or built-in font
        let font = doc.add_builtin_font(BuiltinFont::Helvetica)?;
        
        // Configure text layout
        let font_size = 11.0;
        let line_height = Mm(5.0);
        let margin_left = Mm(20.0);
        let margin_top = Mm(280.0);
        let margin_bottom = Mm(20.0);
        let page_width = Mm(210.0);
        let page_height = Mm(297.0);
        let text_width = page_width - (margin_left * 2.0);
        
        let lines: Vec<&str> = text.lines().collect();
        let mut current_page = page1;
        let mut current_layer = layer1;
        let mut y_position = margin_top;
        
        for line in lines {
            // Check if we need a new page
            if y_position < margin_bottom {
                let (new_page, new_layer) = doc.add_page(Mm(210.0), Mm(297.0), "Page layer");
                current_page = new_page;
                current_layer = new_layer;
                y_position = margin_top;
            }
            
            // Word wrap if line is too long
            let words: Vec<&str> = line.split_whitespace().collect();
            let mut current_line = String::new();
            let max_chars_per_line = 80; // Approximate
            
            for word in words {
                if current_line.len() + word.len() + 1 > max_chars_per_line {
                    // Write current line
                    if !current_line.is_empty() {
                        doc.get_page(current_page)
                            .get_layer(current_layer)
                            .use_text(&current_line, font_size, margin_left, y_position, &font);
                        y_position -= line_height;
                        current_line.clear();
                        
                        // Check for new page
                        if y_position < margin_bottom {
                            let (new_page, new_layer) = doc.add_page(Mm(210.0), Mm(297.0), "Page layer");
                            current_page = new_page;
                            current_layer = new_layer;
                            y_position = margin_top;
                        }
                    }
                }
                
                if !current_line.is_empty() {
                    current_line.push(' ');
                }
                current_line.push_str(word);
            }
            
            // Write remaining text in line
            if !current_line.is_empty() {
                doc.get_page(current_page)
                    .get_layer(current_layer)
                    .use_text(&current_line, font_size, margin_left, y_position, &font);
                y_position -= line_height;
            }
        }
        
        // Save PDF
        doc.save(&mut BufWriter::new(File::create(pdf_path)?))?;
        Ok(())
    }

    /// Convert PDF to images using pure Rust
    pub fn pdf_to_images_pure(
        &self,
        pdf_path: &Path,
        output_dir: &Path,
        format: ImageFormat,
    ) -> Result<Vec<PathBuf>> {
        // Parse PDF
        let doc = lopdf::Document::load(pdf_path)?;
        let pages = doc.get_pages();
        
        fs::create_dir_all(output_dir)?;
        let mut output_paths = Vec::new();
        
        // For each page, render to image
        for (page_num, (_page_num, _page_id)) in pages.iter().enumerate() {
            // Create a blank image for the page
            // In a real implementation, you would render the PDF content
            let img = self.render_pdf_page_to_image(&doc, page_num)?;
            
            // Save image
            let extension = match format {
                ImageFormat::Png => "png",
                ImageFormat::Jpeg => "jpg",
                _ => "png",
            };
            
            let output_path = output_dir.join(format!("page_{:03}.{}", page_num + 1, extension));
            // JPEG does not support RGBA; convert to RGB if needed
            if let ImageFormat::Jpeg = format {
                let rgb = img.to_rgb8();
                ::image::DynamicImage::ImageRgb8(rgb).save_with_format(&output_path, format)?;
            } else {
                img.save_with_format(&output_path, format)?;
            }
            output_paths.push(output_path);
        }
        
        Ok(output_paths)
    }

    /// Render a PDF page to image (simplified implementation)
    fn render_pdf_page_to_image(&self, _doc: &lopdf::Document, _page_num: usize) -> Result<DynamicImage> {
        // This is a simplified implementation
        // A full implementation would parse PDF content and render it
        
        // Create a white image as placeholder
        let width = 1240;  // A4 at 150 DPI
        let height = 1754; // A4 at 150 DPI
        
        let mut img = RgbaImage::new(width, height);
        
        // Fill with white background
        for pixel in img.pixels_mut() {
            *pixel = Rgba([255, 255, 255, 255]);
        }
        
        // Add a simple text indicator
        // In production, you would properly render PDF content
        
        Ok(DynamicImage::ImageRgba8(img))
    }

    /// Convert DOCX to images using pure Rust
    pub fn docx_to_images_pure(
        &self,
        docx_path: &Path,
        output_dir: &Path,
        format: ImageFormat,
    ) -> Result<Vec<PathBuf>> {
        // First convert to PDF
        let temp_pdf = NamedTempFile::new()?.into_temp_path();
        self.docx_to_pdf_pure(docx_path, &temp_pdf)?;
        
        // Then convert PDF to images
        self.pdf_to_images_pure(&temp_pdf, output_dir, format)
    }

    /// Create a thumbnail from an image
    pub fn create_thumbnail(
        &self,
        image_path: &Path,
        output_path: &Path,
        width: u32,
        height: u32,
    ) -> Result<()> {
        let img = ::image::open(image_path)
            .with_context(|| format!("Failed to open image {:?}", image_path))?;
        
        let thumbnail = img.thumbnail(width, height);
        thumbnail.save(output_path)
            .with_context(|| format!("Failed to save thumbnail to {:?}", output_path))?;
        
        info!("Created thumbnail {}x{} at {:?}", width, height, output_path);
        Ok(())
    }

    /// Merge multiple PDFs using pure Rust
    pub fn merge_pdfs_pure(&self, pdf_paths: &[PathBuf], output_path: &Path) -> Result<()> {
        use ::lopdf::{Document, Object};
        
        // Create a new document for merging
        let mut merged_doc = Document::with_version("1.5");
        
        // Track page tree
        let mut all_pages = Vec::new();
        
        for pdf_path in pdf_paths {
            let doc = Document::load(pdf_path)?;
            
            // Get pages from the document
            let pages = doc.get_pages();
            
            for (_page_num, page_id) in pages.iter() {
                // Clone the page object
                if let Ok(page_obj) = doc.get_object(*page_id) {
                    let new_id = merged_doc.new_object_id();
                    merged_doc.objects.insert(new_id, page_obj.clone());
                    all_pages.push(new_id);
                }
            }
        }
        
        // Build the page tree for merged document
        let pages_id = merged_doc.new_object_id();
        let pages_dict = ::lopdf::dictionary! {
            "Type" => "Pages",
            "Kids" => all_pages.iter().map(|id| Object::Reference(*id)).collect::<Vec<_>>(),
            "Count" => all_pages.len() as i32,
        };
        merged_doc.objects.insert(pages_id, Object::Dictionary(pages_dict));
        
        // Update catalog
        let catalog_id = merged_doc.new_object_id();
        let catalog = ::lopdf::dictionary! {
            "Type" => "Catalog",
            "Pages" => Object::Reference(pages_id),
        };
        merged_doc.objects.insert(catalog_id, Object::Dictionary(catalog));
        merged_doc.trailer.set("Root", Object::Reference(catalog_id));
        
        // Save the merged PDF
        merged_doc.save(output_path)?;
        
        info!("Successfully merged {} PDFs into {:?}", pdf_paths.len(), output_path);
        Ok(())
    }

    /// Split a PDF into individual pages using pure Rust
    pub fn split_pdf_pure(&self, pdf_path: &Path, output_dir: &Path) -> Result<Vec<PathBuf>> {
        use ::lopdf::Document;
        
        fs::create_dir_all(output_dir)?;
        
        let doc = Document::load(pdf_path)?;
        let pages = doc.get_pages();
        let mut output_paths = Vec::new();
        
        for (i, (_page_num, page_id)) in pages.iter().enumerate() {
            // Create a new document with just this page
            let mut single_page_doc = Document::with_version("1.5");
            
            // Clone the page
            if let Ok(page_obj) = doc.get_object(*page_id) {
                let new_page_id = single_page_doc.new_object_id();
                single_page_doc.objects.insert(new_page_id, page_obj.clone());
                
                // Create page tree
                let pages_id = single_page_doc.new_object_id();
                let pages_dict = ::lopdf::dictionary! {
                    "Type" => "Pages",
                    "Kids" => vec![Object::Reference(new_page_id)],
                    "Count" => 1,
                };
                single_page_doc.objects.insert(pages_id, Object::Dictionary(pages_dict));
                
                // Create catalog
                let catalog_id = single_page_doc.new_object_id();
                let catalog = ::lopdf::dictionary! {
                    "Type" => "Catalog",
                    "Pages" => Object::Reference(pages_id),
                };
                single_page_doc.objects.insert(catalog_id, Object::Dictionary(catalog));
                single_page_doc.trailer.set("Root", Object::Reference(catalog_id));
                
                // Save the page
                let output_path = output_dir.join(format!("page_{:03}.pdf", i + 1));
                single_page_doc.save(&output_path)?;
                output_paths.push(output_path);
            }
        }
        
        info!("Split PDF into {} pages", output_paths.len());
        Ok(output_paths)
    }

    /// Parse and render markdown to PDF
    pub fn markdown_to_pdf(&self, markdown: &str, pdf_path: &Path) -> Result<()> {
        use pulldown_cmark::{Parser, Event, Tag, TagEnd};
        
        let parser = Parser::new(markdown);
        let mut plain_text = String::new();
        let mut in_code_block = false;
        let mut list_depth = 0;
        
        for event in parser {
            match event {
                Event::Text(text) => {
                    if in_code_block {
                        plain_text.push_str("    ");
                    } else if list_depth > 0 {
                        plain_text.push_str(&"  ".repeat(list_depth));
                    }
                    plain_text.push_str(&text);
                }
                Event::Start(tag) => {
                    match tag {
                        Tag::Heading { level, .. } => {
                            plain_text.push('\n');
                            plain_text.push_str(&"#".repeat(level as usize));
                            plain_text.push(' ');
                        }
                        Tag::Paragraph => {
                            if !plain_text.is_empty() {
                                plain_text.push_str("\n\n");
                            }
                        }
                        Tag::List(_) => {
                            list_depth += 1;
                            plain_text.push('\n');
                        }
                        Tag::Item => {
                            plain_text.push_str("• ");
                        }
                        Tag::CodeBlock(_) => {
                            in_code_block = true;
                            plain_text.push_str("\n\n");
                        }
                        Tag::Emphasis => plain_text.push('*'),
                        Tag::Strong => plain_text.push_str("**"),
                        _ => {}
                    }
                }
                Event::End(tag) => {
                    match tag {
                        TagEnd::Heading(_) => plain_text.push_str("\n\n"),
                        TagEnd::Paragraph => plain_text.push('\n'),
                        TagEnd::List(_) => {
                            list_depth = list_depth.saturating_sub(1);
                            plain_text.push('\n');
                        }
                        TagEnd::Item => plain_text.push('\n'),
                        TagEnd::CodeBlock => {
                            in_code_block = false;
                            plain_text.push_str("\n\n");
                        }
                        TagEnd::Emphasis => plain_text.push('*'),
                        TagEnd::Strong => plain_text.push_str("**"),
                        _ => {}
                    }
                }
                Event::Code(code) => {
                    plain_text.push('`');
                    plain_text.push_str(&code);
                    plain_text.push('`');
                }
                Event::SoftBreak => plain_text.push(' '),
                Event::HardBreak => plain_text.push('\n'),
                _ => {}
            }
        }
        
        self.create_pdf_from_text(&plain_text, pdf_path)?;
        Ok(())
    }
}