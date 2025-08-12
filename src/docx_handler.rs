use anyhow::{Context, Result};
use docx_rs::*;
use std::fs::{self, File};
use std::path::{Path, PathBuf};
use uuid::Uuid;
use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};
use tracing::{info, warn};
use zip::{ZipArchive, ZipWriter};
use zip::write::FileOptions;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DocxMetadata {
    pub id: String,
    pub path: PathBuf,
    pub created_at: DateTime<Utc>,
    pub modified_at: DateTime<Utc>,
    pub size_bytes: u64,
    pub page_count: Option<usize>,
    pub word_count: Option<usize>,
    pub author: Option<String>,
    pub title: Option<String>,
    pub subject: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DocxStyle {
    pub font_family: Option<String>,
    pub font_size: Option<usize>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub color: Option<String>,
    pub alignment: Option<String>,
    pub line_spacing: Option<f32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableData {
    pub rows: Vec<Vec<String>>,
    pub headers: Option<Vec<String>>,
    pub border_style: Option<String>,
    pub col_widths: Option<Vec<u32>>, // approximate column widths (px)
    pub merges: Option<Vec<TableMerge>>, // best-effort merge specs
    pub cell_shading: Option<String>, // hex RGB like "EEEEEE"
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableMerge {
    pub row: usize,
    pub col: usize,
    pub row_span: usize,
    pub col_span: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ImageData {
    pub data: Vec<u8>,
    pub width: Option<u32>,
    pub height: Option<u32>,
    pub alt_text: Option<String>,
}

pub struct DocxHandler {
    temp_dir: PathBuf,
    pub documents: std::collections::HashMap<String, DocxMetadata>,
    // In-memory operations for documents created via this handler
    in_memory_ops: std::collections::HashMap<String, Vec<DocxOp>>,
}

impl DocxHandler {
    pub fn new() -> Result<Self> {
        let base = std::env::var_os("DOCX_MCP_TEMP").map(PathBuf::from).unwrap_or_else(|| std::env::temp_dir());
        let temp_dir = base.join("docx-mcp");
        fs::create_dir_all(&temp_dir)?;
        
        Ok(Self {
            temp_dir,
            documents: std::collections::HashMap::new(),
            in_memory_ops: std::collections::HashMap::new(),
        })
    }

    /// Create a handler that stores temporary documents under the provided base directory
    pub fn new_with_base_dir<P: AsRef<Path>>(base_dir: P) -> Result<Self> {
        let temp_dir = base_dir.as_ref().join("docx-mcp");
        fs::create_dir_all(&temp_dir)?;
        Ok(Self {
            temp_dir,
            documents: std::collections::HashMap::new(),
            in_memory_ops: std::collections::HashMap::new(),
        })
    }

    #[cfg(test)]
    pub fn new_with_temp_dir(temp_dir: &Path) -> Result<Self> {
        let temp_dir = temp_dir.to_path_buf();
        fs::create_dir_all(&temp_dir)?;
        
        Ok(Self {
            temp_dir,
            documents: std::collections::HashMap::new(),
            in_memory_ops: std::collections::HashMap::new(),
        })
    }

    pub fn create_document(&mut self) -> Result<String> {
        let doc_id = Uuid::new_v4().to_string();
        let doc_path = self.temp_dir.join(format!("{}.docx", doc_id));
        
        // Initialize empty document on disk
        if let Some(parent) = doc_path.parent() {
            fs::create_dir_all(parent)
                .with_context(|| format!("Failed to create parent directory for {:?}", doc_path))?;
        }
        let docx = Docx::new();
        let file = File::create(&doc_path)
            .with_context(|| format!("Failed to create DOCX file at {:?}", doc_path))?;
        docx.build().pack(file)
            .with_context(|| format!("Failed to write DOCX package at {:?}", doc_path))?;
        
        let metadata = DocxMetadata {
            id: doc_id.clone(),
            path: doc_path,
            created_at: Utc::now(),
            modified_at: Utc::now(),
            size_bytes: 0,
            page_count: Some(1),
            word_count: Some(0),
            author: None,
            title: None,
            subject: None,
        };
        
        self.documents.insert(doc_id.clone(), metadata);
        self.in_memory_ops.insert(doc_id.clone(), Vec::new());
        info!("Created new document with ID: {}", doc_id);
        
        Ok(doc_id)
    }

    pub fn open_document(&mut self, path: &Path) -> Result<String> {
        let doc_id = Uuid::new_v4().to_string();
        let doc_path = self.temp_dir.join(format!("{}.docx", doc_id));
        
        if let Some(parent) = doc_path.parent() {
            fs::create_dir_all(parent)
                .with_context(|| format!("Failed to create parent directory for {:?}", doc_path))?;
        }
        fs::copy(path, &doc_path)
            .with_context(|| format!("Failed to copy document from {:?}", path))?;
        
        let file_metadata = fs::metadata(&doc_path)?;
        
        let metadata = DocxMetadata {
            id: doc_id.clone(),
            path: doc_path,
            created_at: Utc::now(),
            modified_at: Utc::now(),
            size_bytes: file_metadata.len(),
            page_count: None,
            word_count: None,
            author: None,
            title: None,
            subject: None,
        };
        
        self.documents.insert(doc_id.clone(), metadata);
        info!("Opened document from {:?} with ID: {}", path, doc_id);
        
        Ok(doc_id)
    }

    pub fn add_paragraph(&mut self, doc_id: &str, text: &str, style: Option<DocxStyle>) -> Result<()> {
        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::Paragraph { text: text.to_string(), style });
        self.write_docx(doc_id)?;
        info!("Added paragraph to document {}", doc_id);
        Ok(())
    }

    pub fn add_heading(&mut self, doc_id: &str, text: &str, level: usize) -> Result<()> {
        let _metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        let heading_style = match level {
            1 => "Heading1",
            2 => "Heading2",
            3 => "Heading3",
            4 => "Heading4",
            5 => "Heading5",
            6 => "Heading6",
            _ => "Heading1",
        };
        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::Heading { text: text.to_string(), style: heading_style.to_string() });
        self.write_docx(doc_id)?;
        info!("Added heading level {} to document {}", level, doc_id);
        Ok(())
    }

    pub fn add_table(&mut self, doc_id: &str, table_data: TableData) -> Result<()> {
        let _metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::Table { data: table_data });
        self.write_docx(doc_id)?;
        info!("Added table to document {}", doc_id);
        Ok(())
    }

    pub fn add_list(&mut self, doc_id: &str, items: Vec<String>, ordered: bool) -> Result<()> {
        let _metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::List { items, ordered });
        self.write_docx(doc_id)?;
        info!("Added {} list to document {}", if ordered { "ordered" } else { "unordered" }, doc_id);
        Ok(())
    }

    /// Add a single list item with an explicit indent level (0-based)
    pub fn add_list_item(&mut self, doc_id: &str, text: &str, level: usize, ordered: bool) -> Result<()> {
        let _metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;

        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::ListItem { text: text.to_string(), level, ordered });
        self.write_docx(doc_id)?;
        info!("Added list item (level {}) to document {}", level, doc_id);
        Ok(())
    }

    /// Add an image to the document
    pub fn add_image(&mut self, doc_id: &str, image: ImageData) -> Result<()> {
        let _metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;

        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        let width = image.width.unwrap_or(100);
        let height = image.height.unwrap_or(100);
        ops.push(DocxOp::Image { data: image.data, width, height, alt_text: image.alt_text });
        self.write_docx(doc_id)?;
        info!("Added image to document {}", doc_id);
        Ok(())
    }

    /// Add a hyperlink to the document
    pub fn add_hyperlink(&mut self, doc_id: &str, text: &str, url: &str) -> Result<()> {
        let _metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;

        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::Hyperlink { text: text.to_string(), url: url.to_string() });
        self.write_docx(doc_id)?;
        info!("Added hyperlink to document {}", doc_id);
        Ok(())
    }

    /// Insert a section break with optional page setup (best-effort)
    pub fn add_section_break(
        &mut self,
        doc_id: &str,
        page_size: Option<&str>,
        orientation: Option<&str>,
        margins: Option<MarginsSpec>,
    ) -> Result<()> {
        let _metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;

        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::SectionBreak {
            page_size: page_size.map(|s| s.to_string()),
            orientation: orientation.map(|s| s.to_string()),
            margins,
        });
        self.write_docx(doc_id)?;
        info!("Added section break to document {}", doc_id);
        Ok(())
    }

    pub fn add_page_break(&mut self, doc_id: &str) -> Result<()> {
        let _metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::PageBreak);
        self.write_docx(doc_id)?;
        info!("Added page break to document {}", doc_id);
        Ok(())
    }

    pub fn set_header(&mut self, doc_id: &str, text: &str) -> Result<()> {
        let _metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::Header(text.to_string()));
        self.write_docx(doc_id)?;
        info!("Set header for document {}", doc_id);
        Ok(())
    }

    pub fn set_footer(&mut self, doc_id: &str, text: &str) -> Result<()> {
        let _metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::Footer(text.to_string()));
        self.write_docx(doc_id)?;
        info!("Set footer for document {}", doc_id);
        Ok(())
    }

    /// Convenience: set simple page numbering text in header or footer
    pub fn set_page_numbering(&mut self, doc_id: &str, location: &str, template: Option<&str>) -> Result<()> {
        let text = template.unwrap_or("Page {PAGE} of {PAGES}");
        match location {
            "header" => self.set_header(doc_id, text),
            "footer" => self.set_footer(doc_id, text),
            _ => anyhow::bail!("invalid location: {}", location),
        }
    }

    /// Attempt to replace placeholder page numbering text in header with Word field codes (PAGE/NUMPAGES)
    /// This is a best-effort, post-processing step that edits the zipped DOCX XML in-place by rebuilding the archive.
    pub fn embed_page_number_fields(&self, doc_id: &str) -> Result<()> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        if !metadata.path.exists() {
            anyhow::bail!("Document file missing: {:?}", metadata.path);
        }

        let src_file = std::fs::File::open(&metadata.path)?;
        let mut archive = ZipArchive::new(src_file)?;

        // Prepare buffer to write new archive
        let temp_path = metadata.path.with_extension("docx.tmp");
        let dst_file = std::fs::File::create(&temp_path)?;
        let mut writer = ZipWriter::new(dst_file);
        let options = FileOptions::default().compression_method(zip::CompressionMethod::Stored);

        let mut did_replace = false;
        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let name = file.name().to_string();

            if (name.starts_with("word/header") || name.starts_with("word/footer")) && name.ends_with(".xml") {
                let mut xml = String::new();
                use std::io::Read as _;
                file.read_to_string(&mut xml)?;

                if xml.contains("Page {PAGE} of {PAGES}") {
                    let field_runs = concat!(
                        "Page ",
                        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>",
                        "<w:r><w:instrText xml:space=\"preserve\"> PAGE </w:instrText></w:r>",
                        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>",
                        " of ",
                        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>",
                        "<w:r><w:instrText xml:space=\"preserve\"> NUMPAGES </w:instrText></w:r>",
                        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"
                    );
                    xml = xml.replace("Page {PAGE} of {PAGES}", field_runs);
                    did_replace = true;
                }

                writer.start_file(name, options)?;
                use std::io::Write as _;
                writer.write_all(xml.as_bytes())?;
            } else {
                // Copy other file entries verbatim
                writer.start_file(name, options)?;
                use std::io::Read as _;
                let mut buf = Vec::new();
                file.read_to_end(&mut buf)?;
                use std::io::Write as _;
                writer.write_all(&buf)?;
            }
        }

        writer.finish()?;

        // Replace original archive only if we changed something
        if did_replace {
            std::fs::rename(&temp_path, &metadata.path)?;
            info!("Embedded PAGE/NUMPAGES fields into header for {}", doc_id);
        } else {
            // Cleanup temp
            let _ = std::fs::remove_file(&temp_path);
            info!("No placeholder found to replace for page numbering in {}", doc_id);
        }

        Ok(())
    }

    pub fn find_and_replace(&mut self, doc_id: &str, _find_text: &str, _replace_text: &str) -> Result<usize> {
        let _metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        // Note: This is a simplified implementation
        // Real implementation would need to parse the DOCX XML structure
        // and perform replacements while preserving formatting
        
        warn!("Find and replace operation requires advanced XML manipulation");
        Ok(0)
    }

    /// Advanced find and replace over in-memory operations (LLM-friendly), preserving runs
    /// Supports regex, case sensitivity, and whole word boundaries
    pub fn find_and_replace_advanced(
        &mut self,
        doc_id: &str,
        pattern: &str,
        replacement: &str,
        case_sensitive: bool,
        whole_word: bool,
        use_regex: bool,
    ) -> Result<usize> {
        use regex::RegexBuilder;

        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id)
            .ok_or_else(|| anyhow::anyhow!("No in-memory ops for document: {}", doc_id))?;

        // Build regex
        let pattern = if use_regex { pattern.to_string() } else { regex::escape(pattern) };
        let pattern = if whole_word { format!("\\b{}\\b", pattern) } else { pattern };
        let re = RegexBuilder::new(&pattern)
            .case_insensitive(!case_sensitive)
            .build()
            .with_context(|| "Invalid regex pattern")?;

        let mut total_replacements = 0usize;

        let mut replace_text = |text: &str| -> (String, usize) {
            let mut count = 0usize;
            let result = re.replace_all(text, |_: &regex::Captures| {
                count += 1;
                replacement.to_string()
            });
            (result.into_owned(), count)
        };

        for op in ops.iter_mut() {
            match op {
                DocxOp::Paragraph { text, .. } => {
                    let (new_text, cnt) = replace_text(text);
                    if cnt > 0 { *text = new_text; total_replacements += cnt; }
                }
                DocxOp::Heading { text, .. } => {
                    let (new_text, cnt) = replace_text(text);
                    if cnt > 0 { *text = new_text; total_replacements += cnt; }
                }
                DocxOp::List { items, .. } => {
                    for item in items.iter_mut() {
                        let (new_text, cnt) = replace_text(item);
                        if cnt > 0 { *item = new_text; total_replacements += cnt; }
                    }
                }
                DocxOp::ListItem { text, .. } => {
                    let (new_text, cnt) = replace_text(text);
                    if cnt > 0 { *text = new_text; total_replacements += cnt; }
                }
                DocxOp::Table { data } => {
                    for row in data.rows.iter_mut() {
                        for cell in row.iter_mut() {
                            let (new_text, cnt) = replace_text(cell);
                            if cnt > 0 { *cell = new_text; total_replacements += cnt; }
                        }
                    }
                }
                DocxOp::Header(text) | DocxOp::Footer(text) => {
                    let (new_text, cnt) = replace_text(text);
                    if cnt > 0 { *text = new_text; total_replacements += cnt; }
                }
                DocxOp::Image { .. } | DocxOp::Hyperlink { .. } => {}
                DocxOp::PageBreak => {}
                DocxOp::SectionBreak { .. } => {}
            }
        }

        // Persist changes
        self.write_docx(doc_id)?;
        Ok(total_replacements)
    }

    /// Analyze document structure using in-memory ops (if available)
    pub fn analyze_structure(&self, doc_id: &str) -> Result<serde_json::Value> {
        let ops = match self.in_memory_ops.get(doc_id) {
            Some(ops) => ops,
            None => {
                // Fallback to text-based outline if ops not available
                let text = self.extract_text(doc_id).unwrap_or_default();
                let mut outline = Vec::new();
                for line in text.lines() {
                    let trimmed = line.trim();
                    if trimmed.is_empty() { continue; }
                    if trimmed.len() < 100 && trimmed.chars().any(|c| c.is_uppercase()) {
                        let level = if trimmed.chars().all(|c| c.is_uppercase() || c.is_whitespace()) { 1 } else { 2 };
                        outline.push(serde_json::json!({"type":"heading","text":trimmed,"level":level}));
                    }
                }
                return Ok(serde_json::json!({
                    "has_ops": false,
                    "outline": outline,
                    "lists": [],
                    "tables": [],
                    "images": [],
                    "links": [],
                    "styles": {}
                }));
            }
        };

        let mut outline = Vec::new();
        let mut lists = Vec::new();
        let mut tables = Vec::new();
        let mut images = Vec::new();
        let mut links = Vec::new();
        let mut styles_used: std::collections::HashMap<String, usize> = std::collections::HashMap::new();

        for op in ops.iter() {
            match op {
                DocxOp::Heading { text, style } => {
                    let level = style.chars().last().and_then(|c| c.to_digit(10)).map(|d| d as usize).unwrap_or(1);
                    outline.push(serde_json::json!({"text": text, "level": level}));
                }
                DocxOp::List { items, .. } => {
                    lists.push(serde_json::json!({"level": 0, "items": items}));
                }
                DocxOp::ListItem { text, level, .. } => {
                    lists.push(serde_json::json!({"level": level, "items": [text]}));
                }
                DocxOp::Table { data } => {
                    let rows = data.rows.len();
                    let cols = data.rows.first().map(|r| r.len()).unwrap_or(0);
                    tables.push(serde_json::json!({"rows": rows, "cols": cols}));
                }
                DocxOp::Image { width, height, .. } => {
                    images.push(serde_json::json!({"width": width, "height": height}));
                }
                DocxOp::Hyperlink { text, url } => {
                    links.push(serde_json::json!({"text": text, "url": url}));
                }
                DocxOp::Paragraph { style, .. } => {
                    if let Some(s) = style {
                        if s.bold == Some(true) { *styles_used.entry("bold".into()).or_default() += 1; }
                        if s.italic == Some(true) { *styles_used.entry("italic".into()).or_default() += 1; }
                        if s.underline == Some(true) { *styles_used.entry("underline".into()).or_default() += 1; }
                        if s.font_family.is_some() { *styles_used.entry("font_family".into()).or_default() += 1; }
                        if s.font_size.is_some() { *styles_used.entry("font_size".into()).or_default() += 1; }
                        if s.color.is_some() { *styles_used.entry("color".into()).or_default() += 1; }
                        if s.alignment.is_some() { *styles_used.entry("alignment".into()).or_default() += 1; }
                    }
                }
                DocxOp::Header(_) | DocxOp::Footer(_) | DocxOp::PageBreak | DocxOp::SectionBreak { .. } => {}
            }
        }

        Ok(serde_json::json!({
            "has_ops": true,
            "outline": outline,
            "lists": lists,
            "tables": tables,
            "images": images,
            "links": links,
            "styles": styles_used,
        }))
    }

    pub fn extract_text(&self, doc_id: &str) -> Result<String> {
        let _metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        // Use pure Rust text extraction
        use crate::pure_converter::PureRustConverter;
        let converter = PureRustConverter::new();
        let text = converter.extract_text_from_docx(&self.documents.get(doc_id).unwrap().path)
            .with_context(|| format!("Failed to extract text from document {}", doc_id))?;
        
        Ok(text)
    }

    pub fn get_metadata(&self, doc_id: &str) -> Result<DocxMetadata> {
        self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))
            .map(|m| m.clone())
    }

    /// Update document core properties stored in our metadata (best-effort)
    pub fn set_document_properties(
        &mut self,
        doc_id: &str,
        title: Option<String>,
        subject: Option<String>,
        author: Option<String>,
    ) -> Result<()> {
        let meta = self.documents.get_mut(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        if let Some(t) = title { meta.title = Some(t); }
        if let Some(s) = subject { meta.subject = Some(s); }
        if let Some(a) = author { meta.author = Some(a); }
        Ok(())
    }

    pub fn get_document_properties_json(&self, doc_id: &str) -> Result<serde_json::Value> {
        let meta = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        Ok(serde_json::json!({
            "title": meta.title,
            "subject": meta.subject,
            "author": meta.author,
            "created_at": meta.created_at,
            "modified_at": meta.modified_at,
        }))
    }

    /// Insert a paragraph after the first heading that matches `heading_text`
    pub fn insert_after_heading(&mut self, doc_id: &str, heading_text: &str, text: &str) -> Result<bool> {
        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        if let Some(pos) = ops.iter().position(|op| matches!(op, DocxOp::Heading { text: t, .. } if t == heading_text)) {
            ops.insert(pos + 1, DocxOp::Paragraph { text: text.to_string(), style: None });
            self.write_docx(doc_id)?;
            return Ok(true);
        }
        Ok(false)
    }

    /// Remove external hyperlinks (basic sanitizer)
    pub fn sanitize_external_links(&mut self, doc_id: &str) -> Result<usize> {
        self.ensure_modifiable(doc_id)?;
        let removed = {
            let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
            let before = ops.len();
            ops.retain(|op| match op {
                DocxOp::Hyperlink { url, .. } => {
                    let lower = url.to_lowercase();
                    !(lower.starts_with("http://") || lower.starts_with("https://"))
                }
                _ => true,
            });
            before.saturating_sub(ops.len())
        };
        self.write_docx(doc_id)?;
        Ok(removed)
    }

    /// Redact text using advanced find/replace with a block character
    pub fn redact_text(&mut self, doc_id: &str, pattern: &str, use_regex: bool, whole_word: bool, case_sensitive: bool) -> Result<usize> {
        self.find_and_replace_advanced(doc_id, pattern, "█", case_sensitive, whole_word, use_regex)
    }

    pub fn save_document(&self, doc_id: &str, output_path: &Path) -> Result<()> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        fs::copy(&metadata.path, output_path)
            .with_context(|| format!("Failed to save document to {:?}", output_path))?;
        
        info!("Saved document {} to {:?}", doc_id, output_path);
        Ok(())
    }

    pub fn close_document(&mut self, doc_id: &str) -> Result<()> {
        let metadata = self.documents.remove(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        if metadata.path.exists() {
            fs::remove_file(&metadata.path)?;
        }
        self.in_memory_ops.remove(doc_id);
        
        info!("Closed document {}", doc_id);
        Ok(())
    }

    pub fn list_documents(&self) -> Vec<DocxMetadata> {
        self.documents.values().cloned().collect()
    }

    pub fn temp_dir_path(&self) -> PathBuf {
        self.temp_dir.clone()
    }

    pub fn get_storage_info(&self) -> Result<serde_json::Value> {
        use std::time::UNIX_EPOCH;
        let mut total_bytes: u64 = 0;
        let mut file_count: u64 = 0;
        let mut oldest: Option<u64> = None;
        let mut newest: Option<u64> = None;
        if self.temp_dir.exists() {
            for entry in walkdir::WalkDir::new(&self.temp_dir).into_iter().filter_map(|e| e.ok()) {
                if entry.file_type().is_file() {
                    file_count += 1;
                    if let Ok(meta) = entry.metadata() {
                        total_bytes = total_bytes.saturating_add(meta.len());
                        if let Ok(modified) = meta.modified() {
                            if let Ok(secs) = modified.duration_since(UNIX_EPOCH) {
                                let ts = secs.as_secs();
                                oldest = Some(oldest.map_or(ts, |o| o.min(ts)));
                                newest = Some(newest.map_or(ts, |n| n.max(ts)));
                            }
                        }
                    }
                }
            }
        }
        Ok(serde_json::json!({
            "success": true,
            "storage": {
                "base_dir": self.temp_dir,
                "file_count": file_count,
                "total_bytes": total_bytes,
                "oldest_modified": oldest,
                "newest_modified": newest,
            }
        }))
    }
}

#[derive(Debug, Clone)]
enum DocxOp {
    Paragraph { text: String, style: Option<DocxStyle> },
    Heading { text: String, style: String },
    Table { data: TableData },
    List { items: Vec<String>, ordered: bool },
    ListItem { text: String, level: usize, ordered: bool },
    PageBreak,
    Header(String),
    Footer(String),
    Image { data: Vec<u8>, width: u32, height: u32, alt_text: Option<String> },
    Hyperlink { text: String, url: String },
    SectionBreak { page_size: Option<String>, orientation: Option<String>, margins: Option<MarginsSpec> },
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MarginsSpec {
    pub top: Option<f32>,
    pub bottom: Option<f32>,
    pub left: Option<f32>,
    pub right: Option<f32>,
}

impl DocxHandler {
    fn ensure_modifiable(&self, doc_id: &str) -> Result<()> {
        if !self.in_memory_ops.contains_key(doc_id) {
            anyhow::bail!("Modifications are supported only for documents created by this server (doc_id: {})", doc_id);
        }
        Ok(())
    }

    fn write_docx(&self, doc_id: &str) -> Result<()> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        let ops = self.in_memory_ops.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("No in-memory ops for document: {}", doc_id))?;

        let mut docx = Docx::new();
        let mut header_text: Option<String> = None;
        let mut footer_text: Option<String> = None;

        for op in ops {
            match op {
                DocxOp::Paragraph { text, style } => {
                    let mut run = Run::new().add_text(text);
                    if let Some(st) = style {
                        if let Some(size) = st.font_size { run = run.size(size); }
                        if st.bold == Some(true) { run = run.bold(); }
                        if st.italic == Some(true) { run = run.italic(); }
                        if st.underline == Some(true) { run = run.underline("single"); }
                        if let Some(color) = &st.color { run = run.color(color.clone()); }
                    }
                    let para = Paragraph::new().add_run(run);
                    docx = docx.add_paragraph(para);
                }
                DocxOp::Heading { text, style } => {
                    let para = Paragraph::new().add_run(Run::new().add_text(text)).style(style);
                    docx = docx.add_paragraph(para);
                }
                DocxOp::Table { data } => {
                    let col_count = data.rows.get(0).map(|r| r.len()).unwrap_or(0);
                    // Note: docx-rs Table::new takes rows, not grid. We'll add rows then (optionally) rely on defaults.
                    let mut table = Table::new(vec![]);

                    // Note: We rely on XML post-processing to inject tblGrid widths when feature enabled.

                    // Pre-compute merge coverage map (best-effort)
                    use std::collections::HashSet;
                    let mut covered: HashSet<(usize, usize)> = HashSet::new();
                    let mut topleft: HashSet<(usize, usize)> = HashSet::new();
                    if let Some(merges) = &data.merges {
                        for m in merges {
                            topleft.insert((m.row, m.col));
                            for dr in 0..m.row_span.max(1) {
                                for dc in 0..m.col_span.max(1) {
                                    covered.insert((m.row + dr, m.col + dc));
                                }
                            }
                        }
                    }

                    let has_header = data.headers.as_ref().map(|h| !h.is_empty()).unwrap_or(false);
                    for (ri, row) in data.rows.iter().enumerate() {
                        let mut cells: Vec<TableCell> = Vec::new();
                        for (ci, cell_text) in row.iter().enumerate() {
                            let pos = (ri, ci);
                            // Only render text in top-left cell of a merge region; others empty
                            let text_to_render = if covered.contains(&pos) && !topleft.contains(&pos) { "" } else { cell_text.as_str() };
                            let mut para = Paragraph::new().add_run(Run::new().add_text(text_to_render));
                            if has_header && ri == 0 {
                                // Mark first row as header style; post-processing will add style definition
                                para = para.style("TableHeader");
                            }
                            let cell = TableCell::new().add_paragraph(para);
                            cells.push(cell);
                        }
                        while cells.len() < col_count { cells.push(TableCell::new()); }
                        table = table.add_row(TableRow::new(cells));
                    }
                    docx = docx.add_table(table);
                }
                DocxOp::List { items, ordered } => {
                    // Use separate numbering ids for ordered vs unordered so we can post-process numbering.xml
                    let (abstract_id, concrete_id) = if *ordered { (10usize, 11usize) } else { (20usize, 21usize) };
                    docx = docx
                        .add_abstract_numbering(docx_rs::AbstractNumbering::new(abstract_id))
                        .add_numbering(docx_rs::Numbering::new(concrete_id, abstract_id));
                    for item in items {
                        let para = Paragraph::new()
                            .add_run(Run::new().add_text(item))
                            .numbering(NumberingId::new(concrete_id), IndentLevel::new(0));
                        docx = docx.add_paragraph(para);
                    }
                }
                DocxOp::ListItem { text, level, ordered } => {
                    let (abstract_id, concrete_id) = if *ordered { (10usize, 11usize) } else { (20usize, 21usize) };
                    docx = docx
                        .add_abstract_numbering(docx_rs::AbstractNumbering::new(abstract_id))
                        .add_numbering(docx_rs::Numbering::new(concrete_id, abstract_id));
                    let para = Paragraph::new()
                        .add_run(Run::new().add_text(text))
                        .numbering(NumberingId::new(concrete_id), IndentLevel::new(*level));
                    docx = docx.add_paragraph(para);
                }
                DocxOp::PageBreak => {
                    let para = Paragraph::new().add_run(Run::new().add_break(BreakType::Page));
                    docx = docx.add_paragraph(para);
                }
                DocxOp::Header(text) => { header_text = Some(text.clone()); }
                DocxOp::Footer(text) => { footer_text = Some(text.clone()); }
                DocxOp::Image { data, width, height, alt_text: _ } => {
                    let run = Run::new();
                    let pic = Pic::new_with_dimensions(data.clone(), *width, *height);
                    let para = Paragraph::new().add_run(run.add_image(pic));
                    docx = docx.add_paragraph(para);
                }
                DocxOp::Hyperlink { text, url } => {
                    let link = Hyperlink::new(url, HyperlinkType::External)
                        .add_run(Run::new().add_text(text).color("0000FF").underline("single"));
                    let para = Paragraph::new().add_hyperlink(link);
                    docx = docx.add_paragraph(para);
                }
                DocxOp::SectionBreak { .. } => {
                    // Best-effort: denote a section break with a page break
                    let para = Paragraph::new().add_run(Run::new().add_break(BreakType::Page));
                    docx = docx.add_paragraph(para);
                }
            }
        }

        if let Some(h) = header_text {
            let header = Header::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(h)));
            docx = docx.header(header);
        }
        if let Some(f) = footer_text {
            let footer = Footer::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(f)));
            docx = docx.footer(footer);
        }

        let file = File::create(&metadata.path)?;
        docx.build().pack(file)?;

        // Optionally post-process to inject high-fidelity XML
        #[cfg(feature = "hi-fidelity-tables")]
        {
            self.apply_table_xml_properties(&metadata.path, ops)?;
        }
        #[cfg(feature = "hi-fidelity-styles")]
        {
            self.apply_styles_xml_properties(&metadata.path)?;
        }
        #[cfg(feature = "hi-fidelity-lists")]
        {
            self.apply_numbering_xml_properties(&metadata.path, ops)?;
        }
        #[cfg(feature = "hi-fidelity-sections")]
        {
            self.apply_section_xml_properties(&metadata.path, ops)?;
        }
        Ok(())
    }
}

#[cfg(feature = "hi-fidelity-tables")]
impl DocxHandler {
    fn apply_table_xml_properties(&self, docx_path: &Path, ops: &Vec<DocxOp>) -> Result<()> {
        // Open existing archive
        let src_file = std::fs::File::open(docx_path)?;
        let mut archive = ZipArchive::new(src_file)?;

        // Read document.xml into memory
        let mut document_xml = String::new();
        {
            let mut f = archive.by_name("word/document.xml")?;
            use std::io::Read as _;
            f.read_to_string(&mut document_xml)?;
        }

        // Count tables and build a merge map per table based on ops order
        // We assume each DocxOp::Table corresponds to a <w:tbl> in order.
        let mut table_merge_specs: Vec<(Option<Vec<u32>>, Option<Vec<TableMerge>>)> = Vec::new();
        for op in ops.iter() {
            if let DocxOp::Table { data } = op {
                table_merge_specs.push((data.col_widths.clone(), data.merges.clone()));
            }
        }

        if table_merge_specs.is_empty() {
            return Ok(());
        }

        // Perform a minimal XML manipulation using string operations to inject gridSpan/vMerge
        // This is a best-effort approach and assumes simple structure generated by docx-rs.
        // Strategy:
        // - Iterate through each <w:tbl> block sequentially.
        // - Within each table, iterate rows and cells; when a merge starts at (r,c), add w:gridSpan and/or w:vMerge="restart".
        // - For cells covered by vertical continuation, set w:vMerge="continue" and remove text if present.
        // - If col_widths provided, ensure a <w:tblGrid> with <w:gridCol w:w="..."/> entries exists.

        // Split tables
        let mut output = String::new();
        let mut rest = document_xml.as_str();
        let mut tbl_index = 0usize;
        while let Some(start) = rest.find("<w:tbl") {
            let (head, after_head) = rest.split_at(start);
            output.push_str(head);
            // Find end of table
            if let Some(end) = after_head.find("</w:tbl>") {
                let (tbl_block, tail) = after_head.split_at(end + "</w:tbl>".len());
                let processed = self.process_single_table_xml(tbl_block, table_merge_specs.get(tbl_index))?;
                output.push_str(&processed);
                rest = tail;
                tbl_index += 1;
            } else {
                // Malformed; break
                output.push_str(after_head);
                rest = "";
                break;
            }
        }
        output.push_str(rest);

        if output != document_xml {
            // Rebuild archive with modified document.xml
            let temp_path = docx_path.with_extension("docx.tmp");
            let dst_file = std::fs::File::create(&temp_path)?;
            let mut writer = ZipWriter::new(dst_file);
            let options = FileOptions::default().compression_method(zip::CompressionMethod::Stored);

            for i in 0..archive.len() {
                let mut file = archive.by_index(i)?;
                let name = file.name().to_string();
                writer.start_file(name.clone(), options)?;
                use std::io::{Read as _, Write as _};
                if name == "word/document.xml" {
                    writer.write_all(output.as_bytes())?;
                } else {
                    let mut buf = Vec::new();
                    file.read_to_end(&mut buf)?;
                    writer.write_all(&buf)?;
                }
            }

            writer.finish()?;
            std::fs::rename(&temp_path, docx_path)?;
        }

        Ok(())
    }

    fn process_single_table_xml(&self, tbl_xml: &str, spec: Option<&(Option<Vec<u32>>, Option<Vec<TableMerge>>)>) -> Result<String> {
        if spec.is_none() { return Ok(tbl_xml.to_string()); }
        let (col_widths, merges_opt) = spec.unwrap();
        let mut out = tbl_xml.to_string();

        // Ensure tblGrid
        if let Some(widths) = col_widths {
            if !widths.is_empty() {
                if !out.contains("<w:tblGrid") {
                    // Insert after <w:tblPr> if present, else right after <w:tbl>
                    if let Some(pr_end) = out.find("</w:tblPr>") {
                        let insert_pos = pr_end + "</w:tblPr>".len();
                        let grid_xml = self.render_tbl_grid(widths);
                        out.insert_str(insert_pos, &grid_xml);
                    } else if let Some(tbl_start_end) = out.find(">") {
                        // after opening <w:tbl>
                        let insert_pos = tbl_start_end + 1;
                        let grid_xml = self.render_tbl_grid(widths);
                        out.insert_str(insert_pos, &grid_xml);
                    }
                } else {
                    // Replace existing grid (supports normal and self-closing forms)
                    let grid_xml = self.render_tbl_grid(widths);
                    if let Some(gstart) = out.find("<w:tblGrid") {
                        let rel = &out[gstart..];
                        if let Some(self_close) = rel.find("/>") {
                            let end_abs = gstart + self_close + 2; // include "/>"
                            out.replace_range(gstart..end_abs, &grid_xml);
                        } else if let Some(gend) = rel.find("</w:tblGrid>") {
                            let gend_abs = gstart + gend + "</w:tblGrid>".len();
                            out.replace_range(gstart..gend_abs, &grid_xml);
                        }
                    }
                }
            }
        }

        // Apply merges
        if let Some(merges) = merges_opt {
            // Tokenize rows and cells sequentially best-effort
            let mut ri = 0usize;
            let mut cursor = 0usize;
            while let Some(tr_start_off) = out[cursor..].find("<w:tr") {
                let tr_start = cursor + tr_start_off;
                if let Some(tr_end_rel) = out[tr_start..].find("</w:tr>") {
                    let tr_end = tr_start + tr_end_rel + "</w:tr>".len();
                    let mut tr_block = out[tr_start..tr_end].to_string();

                    // Walk cells
                    let mut ci = 0usize;
                    let mut tr_cursor = 0usize;
                    while let Some(tc_start_off) = tr_block[tr_cursor..].find("<w:tc") {
                        let tc_start = tr_cursor + tc_start_off;
                        if let Some(tc_end_rel) = tr_block[tc_start..].find("</w:tc>") {
                            let tc_end = tc_start + tc_end_rel + "</w:tc>".len();
                            let mut tc_block = tr_block[tc_start..tc_end].to_string();

                            // Determine merge action for this cell
                            let mut grid_span: Option<usize> = None;
                            let mut vmerge: Option<&'static str> = None; // "restart" or "continue"
                            for m in merges {
                                if m.row == ri && m.col == ci {
                                    if m.col_span > 1 { grid_span = Some(m.col_span); }
                                    if m.row_span > 1 { vmerge = Some("restart"); }
                                } else if m.col == ci && ri > m.row && ri < m.row + m.row_span && ci >= m.col && ci < m.col + m.col_span {
                                    // vertically covered cell
                                    if m.row_span > 1 { vmerge = Some("continue"); }
                                }
                            }

                            if grid_span.is_some() || vmerge.is_some() {
                                // Ensure <w:tcPr> exists
                                if let Some(pr_start) = tc_block.find("<w:tcPr>") {
                                    let insert_at = pr_start + "<w:tcPr>".len();
                                    let mut props = String::new();
                                    if let Some(span) = grid_span { props.push_str(&format!("<w:gridSpan w:val=\"{}\"/>", span)); }
                                    if let Some(vm) = vmerge { props.push_str(&format!("<w:vMerge w:val=\"{}\"/>", vm)); }
                                    tc_block.insert_str(insert_at, &props);
                                } else {
                                    // Insert tcPr after <w:tc>
                                    if let Some(tc_open_end) = tc_block.find(">") {
                                        let insert_at = tc_open_end + 1;
                                        let mut props = String::new();
                                        props.push_str("<w:tcPr>");
                                        if let Some(span) = grid_span { props.push_str(&format!("<w:gridSpan w:val=\"{}\"/>", span)); }
                                        if let Some(vm) = vmerge { props.push_str(&format!("<w:vMerge w:val=\"{}\"/>", vm)); }
                                        props.push_str("</w:tcPr>");
                                        tc_block.insert_str(insert_at, &props);
                                    }
                                }
                            }

                            // Replace back this cell
                            tr_block.replace_range(tc_start..tc_end, &tc_block);
                            tr_cursor = tc_start + tc_block.len();
                            ci += 1;
                        } else { break; }
                    }

                    // Replace back this row
                    out.replace_range(tr_start..tr_end, &tr_block);
                    cursor = tr_start + tr_block.len();
                    ri += 1;
                } else { break; }
            }
        }

        Ok(out)
    }

    fn render_tbl_grid(&self, widths: &Vec<u32>) -> String {
        let mut s = String::from("<w:tblGrid>");
        for w in widths.iter() {
            s.push_str(&format!("<w:gridCol w:w=\"{}\"/>", w));
        }
        s.push_str("</w:tblGrid>");
        s
    }
}

#[cfg(feature = "hi-fidelity-styles")]
impl DocxHandler {
    fn apply_styles_xml_properties(&self, docx_path: &Path) -> Result<()> {
        let src_file = std::fs::File::open(docx_path)?;
        let mut archive = ZipArchive::new(src_file)?;

        // Read or initialize styles.xml
        let mut styles_xml = String::new();
        let mut has_styles = false;
        if let Ok(mut f) = archive.by_name("word/styles.xml") {
            use std::io::Read as _;
            f.read_to_string(&mut styles_xml)?;
            has_styles = true;
        } else {
            styles_xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"></w:styles>");
        }

        if !styles_xml.contains("w:styleId=\"TableHeader\"") {
            let style_def = concat!(
                "<w:style w:type=\"paragraph\" w:styleId=\"TableHeader\">",
                "<w:name w:val=\"TableHeader\"/>",
                "<w:basedOn w:val=\"Normal\"/>",
                "<w:qFormat/>",
                "<w:rPr><w:b/><w:sz w:val=\"24\"/></w:rPr>",
                "<w:pPr><w:spacing w:after=\"0\"/><w:jc w:val=\"center\"/></w:pPr>",
                "</w:style>"
            );
            if let Some(pos) = styles_xml.rfind("</w:styles>") {
                styles_xml.insert_str(pos, style_def);
            }
        }

        // Repack archive with updated styles.xml
        let temp_path = docx_path.with_extension("docx.tmp");
        let dst_file = std::fs::File::create(&temp_path)?;
        let mut writer = ZipWriter::new(dst_file);
        let options = FileOptions::default().compression_method(zip::CompressionMethod::Stored);
        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let name = file.name().to_string();
            use std::io::{Read as _, Write as _};
            writer.start_file(name.clone(), options)?;
            if name == "word/styles.xml" {
                writer.write_all(styles_xml.as_bytes())?;
            } else {
                let mut buf = Vec::new();
                file.read_to_end(&mut buf)?;
                writer.write_all(&buf)?;
            }
        }

        if !has_styles {
            // If styles.xml was missing originally, ensure it is added
            writer.start_file("word/styles.xml".to_string(), options)?;
            use std::io::Write as _;
            writer.write_all(styles_xml.as_bytes())?;
        }

        writer.finish()?;
        std::fs::rename(&temp_path, docx_path)?;
        Ok(())
    }
}

#[cfg(feature = "hi-fidelity-lists")]
impl DocxHandler {
    fn apply_numbering_xml_properties(&self, docx_path: &Path, ops: &Vec<DocxOp>) -> Result<()> {
        // Determine which list types are used
        let mut need_ordered = false;
        let mut need_unordered = false;
        for op in ops.iter() {
            match op {
                DocxOp::List { ordered, .. } => { if *ordered { need_ordered = true; } else { need_unordered = true; } }
                DocxOp::ListItem { ordered, .. } => { if *ordered { need_ordered = true; } else { need_unordered = true; } }
                _ => {}
            }
        }
        if !need_ordered && !need_unordered { return Ok(()); }

        let src_file = std::fs::File::open(docx_path)?;
        let mut archive = ZipArchive::new(src_file)?;

        // Read numbering.xml
        let mut numbering_xml = String::new();
        {
            let mut f = archive.by_name("word/numbering.xml").map_err(|_| anyhow::anyhow!("numbering.xml not found; ensure lists are added before calling"))?;
            use std::io::Read as _;
            f.read_to_string(&mut numbering_xml)?;
        }

        // Ensure abstractNum for ordered (10) and unordered (20)
        if need_ordered && !numbering_xml.contains("w:abstractNumId=\"10\"") {
            let block = self.make_abstract_num_block(10, false);
            if let Some(pos) = numbering_xml.find("</w:numbering>") {
                numbering_xml.insert_str(pos, &block);
            }
        }
        if need_unordered && !numbering_xml.contains("w:abstractNumId=\"20\"") {
            let block = self.make_abstract_num_block(20, true);
            if let Some(pos) = numbering_xml.find("</w:numbering>") {
                numbering_xml.insert_str(pos, &block);
            }
        }

        // Write back
        let temp_path = docx_path.with_extension("docx.tmp");
        let dst_file = std::fs::File::create(&temp_path)?;
        let mut writer = ZipWriter::new(dst_file);
        let options = FileOptions::default().compression_method(zip::CompressionMethod::Stored);
        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let name = file.name().to_string();
            use std::io::{Read as _, Write as _};
            writer.start_file(name.clone(), options)?;
            if name == "word/numbering.xml" {
                writer.write_all(numbering_xml.as_bytes())?;
            } else {
                let mut buf = Vec::new();
                file.read_to_end(&mut buf)?;
                writer.write_all(&buf)?;
            }
        }
        writer.finish()?;
        std::fs::rename(&temp_path, docx_path)?;
        Ok(())
    }

    fn make_abstract_num_block(&self, abstract_id: usize, bullet: bool) -> String {
        let mut s = format!("<w:abstractNum w:abstractNumId=\"{}\">", abstract_id);
        for lvl in 0..9 {
            let (fmt, txt) = if bullet { ("bullet", "•") } else { ("decimal", match lvl { 0 => "%1.", 1 => "%2.", 2 => "%3.", 3 => "%4.", 4 => "%5.", 5 => "%6.", 6 => "%7.", 7 => "%8.", _ => "%9." }) };
            let lvl_text = if bullet { "•".to_string() } else { txt.to_string() };
            s.push_str(&format!(
                concat!(
                    "<w:lvl w:ilvl=\"{lvl}\">",
                    "<w:start w:val=\"1\"/>",
                    "<w:numFmt w:val=\"{fmt}\"/>",
                    "<w:lvlText w:val=\"{lvl_text}\"/>",
                    "<w:lvlJc w:val=\"left\"/>",
                    "<w:pPr><w:ind w:left=\"{left}\" w:hanging=\"{hang}\"/></w:pPr>",
                    "</w:lvl>"
                ),
                lvl=lvl,
                fmt=fmt,
                lvl_text=lvl_text,
                left=(lvl as i32 + 1) * 720,
                hang=360,
            ));
        }
        s.push_str("</w:abstractNum>");
        s
    }
}

#[cfg(feature = "hi-fidelity-sections")]
impl DocxHandler {
    fn apply_section_xml_properties(&self, docx_path: &Path, ops: &Vec<DocxOp>) -> Result<()> {
        // Use the last section break spec, if any
        let mut last_spec: Option<(Option<String>, Option<String>, Option<MarginsSpec>)> = None;
        for op in ops.iter() {
            if let DocxOp::SectionBreak { page_size, orientation, margins } = op {
                last_spec = Some((page_size.clone(), orientation.clone(), margins.clone()));
            }
        }
        if last_spec.is_none() { return Ok(()); }
        let (page_size, orientation, margins) = last_spec.unwrap();

        let (mut w, mut h) = match page_size.as_deref() {
            Some("Letter") => (12240i32, 15840i32), // 8.5x11 in
            _ => (11906i32, 16838i32), // default A4 210x297mm
        };
        if orientation.as_deref() == Some("landscape") {
            std::mem::swap(&mut w, &mut h);
        }
        let margins = margins.unwrap_or(MarginsSpec { top: Some(1.0), bottom: Some(1.0), left: Some(1.0), right: Some(1.0) });
        let to_twips = |opt: Option<f32>| -> i32 { ((opt.unwrap_or(1.0) * 1440.0).round() as i32).max(0) };
        let mt = to_twips(margins.top);
        let mb = to_twips(margins.bottom);
        let ml = to_twips(margins.left);
        let mr = to_twips(margins.right);

        let sect_pr = if orientation.as_deref() == Some("landscape") {
            format!("<w:sectPr><w:pgSz w:w=\"{}\" w:h=\"{}\" w:orient=\"landscape\"/><w:pgMar w:top=\"{}\" w:bottom=\"{}\" w:left=\"{}\" w:right=\"{}\"/></w:sectPr>", w, h, mt, mb, ml, mr)
        } else {
            format!("<w:sectPr><w:pgSz w:w=\"{}\" w:h=\"{}\"/><w:pgMar w:top=\"{}\" w:bottom=\"{}\" w:left=\"{}\" w:right=\"{}\"/></w:sectPr>", w, h, mt, mb, ml, mr)
        };

        let src_file = std::fs::File::open(docx_path)?;
        let mut archive = ZipArchive::new(src_file)?;
        let mut document_xml = String::new();
        {
            let mut f = archive.by_name("word/document.xml")?;
            use std::io::Read as _;
            f.read_to_string(&mut document_xml)?;
        }

        if let Some(pos) = document_xml.rfind("</w:body>") {
            // Replace existing sectPr if present near end
            if let Some(existing_start_rel) = document_xml[..pos].rfind("<w:sectPr") {
                let closing_rel = document_xml[existing_start_rel..].find("</w:sectPr>");
                if let Some(closing_rel) = closing_rel {
                    let closing_abs = existing_start_rel + closing_rel + "</w:sectPr>".len();
                    document_xml.replace_range(existing_start_rel..closing_abs, &sect_pr);
                } else {
                    document_xml.insert_str(pos, &sect_pr);
                }
            } else {
                document_xml.insert_str(pos, &sect_pr);
            }
        }

        // Write back
        let temp_path = docx_path.with_extension("docx.tmp");
        let dst_file = std::fs::File::create(&temp_path)?;
        let mut writer = ZipWriter::new(dst_file);
        let options = FileOptions::default().compression_method(zip::CompressionMethod::Stored);
        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let name = file.name().to_string();
            use std::io::{Read as _, Write as _};
            writer.start_file(name.clone(), options)?;
            if name == "word/document.xml" {
                writer.write_all(document_xml.as_bytes())?;
            } else {
                let mut buf = Vec::new();
                file.read_to_end(&mut buf)?;
                writer.write_all(&buf)?;
            }
        }
        writer.finish()?;
        std::fs::rename(&temp_path, docx_path)?;
        Ok(())
    }
}