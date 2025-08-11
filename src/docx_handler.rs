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
}

impl DocxHandler {
    pub fn new() -> Result<Self> {
        let temp_dir = std::env::temp_dir().join("docx-mcp");
        fs::create_dir_all(&temp_dir)?;
        
        Ok(Self {
            temp_dir,
            documents: std::collections::HashMap::new(),
        })
    }

    #[cfg(test)]
    pub fn new_with_temp_dir(temp_dir: &Path) -> Result<Self> {
        let temp_dir = temp_dir.to_path_buf();
        fs::create_dir_all(&temp_dir)?;
        
        Ok(Self {
            temp_dir,
            documents: std::collections::HashMap::new(),
        })
    }

    pub fn create_document(&mut self) -> Result<String> {
        let doc_id = Uuid::new_v4().to_string();
        let doc_path = self.temp_dir.join(format!("{}.docx", doc_id));
        
        let docx = Docx::new();
        let file = File::create(&doc_path)?;
        docx.build().pack(file)?;
        
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
        info!("Created new document with ID: {}", doc_id);
        
        Ok(doc_id)
    }

    pub fn open_document(&mut self, path: &Path) -> Result<String> {
        let doc_id = Uuid::new_v4().to_string();
        let doc_path = self.temp_dir.join(format!("{}.docx", doc_id));
        
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
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        let mut file = File::open(&metadata.path)?;
        let mut buffer = Vec::new();
        file.read_to_end(&mut buffer)?;
        
        let mut docx = Docx::from_reader(&buffer[..])?;
        
        let mut paragraph = Paragraph::new().add_run(Run::new().add_text(text));
        
        if let Some(style) = style {
            let mut run = Run::new().add_text(text);
            
            if let Some(size) = style.font_size {
                run = run.size(size);
            }
            if style.bold == Some(true) {
                run = run.bold();
            }
            if style.italic == Some(true) {
                run = run.italic();
            }
            if style.underline == Some(true) {
                run = run.underline("single");
            }
            if let Some(color) = style.color {
                run = run.color(color);
            }
            
            paragraph = Paragraph::new().add_run(run);
            
            if let Some(alignment) = style.alignment {
                paragraph = match alignment.as_str() {
                    "left" => paragraph.align(AlignmentType::Left),
                    "center" => paragraph.align(AlignmentType::Center),
                    "right" => paragraph.align(AlignmentType::Right),
                    "justify" => paragraph.align(AlignmentType::Justified),
                    _ => paragraph,
                };
            }
        }
        
        docx = docx.add_paragraph(paragraph);
        
        let file = File::create(&metadata.path)?;
        docx.build().pack(file)?;
        
        info!("Added paragraph to document {}", doc_id);
        Ok(())
    }

    pub fn add_heading(&mut self, doc_id: &str, text: &str, level: usize) -> Result<()> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        let mut file = File::open(&metadata.path)?;
        let mut buffer = Vec::new();
        file.read_to_end(&mut buffer)?;
        
        let mut docx = Docx::from_reader(&buffer[..])?;
        
        let heading_style = match level {
            1 => "Heading1",
            2 => "Heading2",
            3 => "Heading3",
            4 => "Heading4",
            5 => "Heading5",
            6 => "Heading6",
            _ => "Heading1",
        };
        
        let paragraph = Paragraph::new()
            .add_run(Run::new().add_text(text))
            .style(heading_style);
        
        docx = docx.add_paragraph(paragraph);
        
        let file = File::create(&metadata.path)?;
        docx.build().pack(file)?;
        
        info!("Added heading level {} to document {}", level, doc_id);
        Ok(())
    }

    pub fn add_table(&mut self, doc_id: &str, table_data: TableData) -> Result<()> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        let mut file = File::open(&metadata.path)?;
        let mut buffer = Vec::new();
        file.read_to_end(&mut buffer)?;
        
        let mut docx = Docx::from_reader(&buffer[..])?;
        
        let col_count = table_data.rows.get(0).map(|r| r.len()).unwrap_or(0);
        let mut table = Table::new(vec![TableCell::new(); col_count]);
        
        for row_data in table_data.rows {
            let mut cells = Vec::new();
            for cell_text in row_data {
                let cell = TableCell::new()
                    .add_paragraph(Paragraph::new().add_run(Run::new().add_text(cell_text)));
                cells.push(cell);
            }
            
            while cells.len() < col_count {
                cells.push(TableCell::new());
            }
            
            table = table.add_row(TableRow::new(cells));
        }
        
        docx = docx.add_table(table);
        
        let file = File::create(&metadata.path)?;
        docx.build().pack(file)?;
        
        info!("Added table to document {}", doc_id);
        Ok(())
    }

    pub fn add_list(&mut self, doc_id: &str, items: Vec<String>, ordered: bool) -> Result<()> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        let mut file = File::open(&metadata.path)?;
        let mut buffer = Vec::new();
        file.read_to_end(&mut buffer)?;
        
        let mut docx = Docx::from_reader(&buffer[..])?;
        
        let numbering_id = if ordered { 1 } else { 2 };
        
        for item in items {
            let paragraph = Paragraph::new()
                .add_run(Run::new().add_text(item))
                .numbering(NumberingId::new(numbering_id), IndentLevel::new(0));
            
            docx = docx.add_paragraph(paragraph);
        }
        
        let file = File::create(&metadata.path)?;
        docx.build().pack(file)?;
        
        info!("Added {} list to document {}", if ordered { "ordered" } else { "unordered" }, doc_id);
        Ok(())
    }

    pub fn add_page_break(&mut self, doc_id: &str) -> Result<()> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        let mut file = File::open(&metadata.path)?;
        let mut buffer = Vec::new();
        file.read_to_end(&mut buffer)?;
        
        let mut docx = Docx::from_reader(&buffer[..])?;
        
        let paragraph = Paragraph::new().add_run(Run::new().add_break(BreakType::Page));
        docx = docx.add_paragraph(paragraph);
        
        let file = File::create(&metadata.path)?;
        docx.build().pack(file)?;
        
        info!("Added page break to document {}", doc_id);
        Ok(())
    }

    pub fn set_header(&mut self, doc_id: &str, text: &str) -> Result<()> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        let mut file = File::open(&metadata.path)?;
        let mut buffer = Vec::new();
        file.read_to_end(&mut buffer)?;
        
        let mut docx = Docx::from_reader(&buffer[..])?;
        
        let header = Header::new().add_paragraph(
            Paragraph::new().add_run(Run::new().add_text(text))
        );
        
        docx = docx.header(header);
        
        let file = File::create(&metadata.path)?;
        docx.build().pack(file)?;
        
        info!("Set header for document {}", doc_id);
        Ok(())
    }

    pub fn set_footer(&mut self, doc_id: &str, text: &str) -> Result<()> {
        let metadata = self.documents.get(doc_id)
            .ok_or_else(|| anyhow::anyhow!("Document not found: {}", doc_id))?;
        
        let mut file = File::open(&metadata.path)?;
        let mut buffer = Vec::new();
        file.read_to_end(&mut buffer)?;
        
        let mut docx = Docx::from_reader(&buffer[..])?;
        
        let footer = Footer::new().add_paragraph(
            Paragraph::new().add_run(Run::new().add_text(text))
        );
        
        docx = docx.footer(footer);
        
        let file = File::create(&metadata.path)?;
        docx.build().pack(file)?;
        
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
        
        info!("Closed document {}", doc_id);
        Ok(())
    }

    pub fn list_documents(&self) -> Vec<DocxMetadata> {
        self.documents.values().cloned().collect()
    }
}