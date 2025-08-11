use anyhow::{Context, Result};
use docx_rs::*;
use std::fs::{self, File};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use tempfile::NamedTempFile;
use uuid::Uuid;
use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};
use tracing::{debug, info, warn};

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
        let metadata = self.documents.get(doc_id)
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
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::Table { data: table_data });
        self.write_docx(doc_id)?;
        info!("Added table to document {}", doc_id);
        Ok(())
    }

    pub fn add_list(&mut self, doc_id: &str, items: Vec<String>, ordered: bool) -> Result<()> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::List { items, ordered });
        self.write_docx(doc_id)?;
        info!("Added {} list to document {}", if ordered { "ordered" } else { "unordered" }, doc_id);
        Ok(())
    }

    pub fn add_page_break(&mut self, doc_id: &str) -> Result<()> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::PageBreak);
        self.write_docx(doc_id)?;
        info!("Added page break to document {}", doc_id);
        Ok(())
    }

    pub fn set_header(&mut self, doc_id: &str, text: &str) -> Result<()> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::Header(text.to_string()));
        self.write_docx(doc_id)?;
        info!("Set header for document {}", doc_id);
        Ok(())
    }

    pub fn set_footer(&mut self, doc_id: &str, text: &str) -> Result<()> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        self.ensure_modifiable(doc_id)?;
        let ops = self.in_memory_ops.get_mut(doc_id).unwrap();
        ops.push(DocxOp::Footer(text.to_string()));
        self.write_docx(doc_id)?;
        info!("Set footer for document {}", doc_id);
        Ok(())
    }

    pub fn find_and_replace(&mut self, doc_id: &str, find_text: &str, replace_text: &str) -> Result<usize> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        // Note: This is a simplified implementation
        // Real implementation would need to parse the DOCX XML structure
        // and perform replacements while preserving formatting
        
        warn!("Find and replace operation requires advanced XML manipulation");
        Ok(0)
    }

    pub fn extract_text(&self, doc_id: &str) -> Result<String> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        // Use pure Rust text extraction
        use crate::pure_converter::PureRustConverter;
        let converter = PureRustConverter::new();
        let text = converter.extract_text_from_docx(&metadata.path)
            .with_context(|| format!("Failed to extract text from document {}", doc_id))?;
        
        Ok(text)
    }

    pub fn get_metadata(&self, doc_id: &str) -> Result<DocxMetadata> {
        self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))
            .map(|m| m.clone())
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
}

#[derive(Debug, Clone)]
enum DocxOp {
    Paragraph { text: String, style: Option<DocxStyle> },
    Heading { text: String, style: String },
    Table { data: TableData },
    List { items: Vec<String>, ordered: bool },
    PageBreak,
    Header(String),
    Footer(String),
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
                    // Build rows
                    let mut table = Table::new(vec![]);
                    for row in &data.rows {
                        let mut cells: Vec<TableCell> = Vec::new();
                        for cell_text in row {
                            let cell = TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(cell_text)));
                            cells.push(cell);
                        }
                        while cells.len() < col_count { cells.push(TableCell::new()); }
                        table = table.add_row(TableRow::new(cells));
                    }
                    docx = docx.add_table(table);
                }
                DocxOp::List { items, ordered } => {
                    // Ensure minimal numbering definitions exist: abstract (0) and concrete (1)
                    let abstract_id = 0usize;
                    let concrete_id = 1usize;
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
                DocxOp::PageBreak => {
                    let para = Paragraph::new().add_run(Run::new().add_break(BreakType::Page));
                    docx = docx.add_paragraph(para);
                }
                DocxOp::Header(text) => { header_text = Some(text.clone()); }
                DocxOp::Footer(text) => { footer_text = Some(text.clone()); }
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
        Ok(())
    }
}