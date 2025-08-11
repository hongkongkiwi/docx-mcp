use anyhow::{Context, Result};
use docx_rs::*;
use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::Path;
use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};
use uuid::Uuid;
use base64;

/// Advanced DOCX manipulation features
pub struct AdvancedDocxHandler;

impl AdvancedDocxHandler {
    pub fn new() -> Self {
        Self
    }

    /// Create a document with professional template
    pub fn create_from_template(&self, template_type: DocumentTemplate) -> Result<Docx> {
        let mut docx = Docx::new();
        
        match template_type {
            DocumentTemplate::BusinessLetter => {
                docx = self.apply_business_letter_template(docx)?;
            }
            DocumentTemplate::Resume => {
                docx = self.apply_resume_template(docx)?;
            }
            DocumentTemplate::Report => {
                docx = self.apply_report_template(docx)?;
            }
            DocumentTemplate::Invoice => {
                docx = self.apply_invoice_template(docx)?;
            }
            DocumentTemplate::Contract => {
                docx = self.apply_contract_template(docx)?;
            }
            DocumentTemplate::Memo => {
                docx = self.apply_memo_template(docx)?;
            }
            DocumentTemplate::Newsletter => {
                docx = self.apply_newsletter_template(docx)?;
            }
        }
        
        Ok(docx)
    }

    /// Add a table of contents
    pub fn add_table_of_contents(&self, docx: Docx) -> Result<Docx> {
        // Basic TOC insertion (heading text paragraph + placeholder)
        let mut docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("Table of Contents").bold().size(28))
                .style("TOCHeading")
        );
        
        // Add instruction text
        let instruction = Paragraph::new()
            .add_run(
                Run::new()
                    .add_text("Right-click and select 'Update Field' to refresh the table of contents")
                    .italic()
                    .size(20)
                    .color("808080")
            );
        
        docx = docx.add_paragraph(instruction);
        docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_break(BreakType::Page)));
        
        Ok(docx)
    }

    /// Add an image to the document
    pub fn add_image(
        &self, 
        docx: Docx, 
        _image_data: &[u8], 
        width_px: u32, 
        height_px: u32,
        alt_text: Option<&str>
    ) -> Result<Docx> {
        // Try to attach a Drawing to the Run via RunChild using the public add_pic shortcut
        let pic = Pic::new_with_dimensions(_image_data.to_vec(), width_px, height_px);
        let paragraph = Paragraph::new().add_run({
            let run = Run::new();
            run.add_image(pic)
        });
        Ok(docx.add_paragraph(paragraph))
    }

    /// Add a chart to the document
    pub fn add_chart(&self, docx: Docx, chart_type: ChartType, data: ChartData) -> Result<Docx> {
        // Charts in DOCX are complex and usually require embedding Excel data
        // For now, we'll create a table representation
        let mut table = Table::new(vec![]);
        
        // Add headers
        let mut header_cells = vec![TableCell::new().add_paragraph(
            Paragraph::new().add_run(Run::new().add_text("Category").bold())
        )];
        
        for series in &data.series {
            header_cells.push(
                TableCell::new().add_paragraph(
                    Paragraph::new().add_run(Run::new().add_text(&series.name).bold())
                )
            );
        }
        table = table.add_row(TableRow::new(header_cells));
        
        // Add data rows
        for (i, category) in data.categories.iter().enumerate() {
            let mut row_cells = vec![TableCell::new().add_paragraph(
                Paragraph::new().add_run(Run::new().add_text(category))
            )];
            
            for series in &data.series {
                if let Some(value) = series.values.get(i) {
                    row_cells.push(
                        TableCell::new().add_paragraph(
                            Paragraph::new().add_run(Run::new().add_text(&value.to_string()))
                        )
                    );
                }
            }
            table = table.add_row(TableRow::new(row_cells));
        }
        
        // Add title for the chart
        let title = Paragraph::new()
            .add_run(Run::new().add_text(&format!("{:?}: {}", chart_type, data.title)).bold())
            .align(AlignmentType::Center);
        
        Ok(docx.add_paragraph(title).add_table(table))
    }

    /// Add a hyperlink
    pub fn add_hyperlink(&self, docx: Docx, text: &str, url: &str) -> Result<Docx> {
        let hyperlink = Hyperlink::new(url, HyperlinkType::External)
            .add_run(Run::new().add_text(text).color("0000FF").underline("single"));
        
        let paragraph = Paragraph::new().add_hyperlink(hyperlink);
        
        Ok(docx.add_paragraph(paragraph))
    }

    /// Add a bookmark
    pub fn add_bookmark(&self, docx: Docx, bookmark_name: &str, text: &str) -> Result<Docx> {
        // Bookmark IDs in 0.4 are usize; fallback to plain paragraph with text
        let paragraph = Paragraph::new().add_run(Run::new().add_text(text));
        
        Ok(docx.add_paragraph(paragraph))
    }

    /// Add a cross-reference
    pub fn add_cross_reference(&self, docx: Docx, bookmark_name: &str, display_text: &str) -> Result<Docx> {
        // Cross-references in DOCX use field codes
        // Complex field support is limited in current docx-rs; fallback to plain hyperlink
        // Fallback: hyperlink not wired; emit text with target in brackets
        let paragraph = Paragraph::new().add_run(Run::new().add_text(format!("{} ({})", display_text, bookmark_name)));
        
        Ok(docx.add_paragraph(paragraph))
    }

    /// Add document properties and metadata
    pub fn set_document_properties(&self, docx: Docx, _properties: DocumentProperties) -> Result<Docx> {
        // Metadata setters not exposed; return unchanged
        Ok(docx)
    }

    /// Add a custom styled section
    pub fn add_section(&self, docx: Docx, section_config: SectionConfig) -> Result<Docx> {
        // Basic section properties (defaults). Page size/columns APIs differ; using defaults.
        Ok(docx)
    }

    /// Add a watermark
    pub fn add_watermark(&self, docx: Docx, text: &str, style: WatermarkStyle) -> Result<Docx> {
        let watermark = match style {
            WatermarkStyle::Diagonal => {
                Run::new()
                    .add_text(text)
                    .size(144) // Large size
                    .color("C0C0C0") // Light gray
                    .bold()
            }
            WatermarkStyle::Horizontal => {
                Run::new()
                    .add_text(text)
                    .size(100)
                    .color("E0E0E0")
            }
        };
        
        // Watermarks are typically added to headers
        let header = Header::new().add_paragraph(
            Paragraph::new()
                .add_run(watermark)
                .align(AlignmentType::Center)
        );
        
        Ok(docx.header(header))
    }

    /// Add footnote
    pub fn add_footnote(&self, docx: Docx, reference_text: &str, footnote_text: &str) -> Result<Docx> {
        let footnote_id = Uuid::new_v4().to_string();
        
        // docx-rs footnote APIs are in flux; append note text inline as fallback
        let paragraph = Paragraph::new()
            .add_run(Run::new().add_text(reference_text))
            .add_run(Run::new().add_text(format!(" [{}]", footnote_text)));
        Ok(docx.add_paragraph(paragraph))
    }

    /// Add endnote
    pub fn add_endnote(&self, docx: Docx, reference_text: &str, endnote_text: &str) -> Result<Docx> {
        let endnote_id = Uuid::new_v4().to_string();
        
        // Fallback inline rendering for endnotes
        let paragraph = Paragraph::new()
            .add_run(Run::new().add_text(reference_text))
            .add_run(Run::new().add_text(format!(" [{}]", endnote_text)));
        Ok(docx.add_paragraph(paragraph))
    }

    /// Add custom styles
    pub fn add_custom_style(&self, docx: Docx, _style: CustomStyle) -> Result<Docx> {
        // Style builder APIs differ; skip custom styles for now
        Ok(docx)
    }

    /// Mail merge functionality
    pub fn prepare_mail_merge_template(&self, docx: Docx, fields: Vec<String>) -> Result<Docx> {
        let mut docx = docx;
        
        for field in fields {
            let paragraph = Paragraph::new()
                .add_run(Run::new().add_text(format!("«{}»", field)));
            
            docx = docx.add_paragraph(paragraph);
        }
        
        Ok(docx)
    }

    /// Add comments (annotations)
    pub fn add_comment(&self, docx: Docx, text: &str, comment: &str, author: &str) -> Result<Docx> {
        let comment_id = Uuid::new_v4().to_string();
        let date = Utc::now();
        
        // Fallback: inline annotation style rendering (no true comment element)
        let paragraph = Paragraph::new()
            .add_run(Run::new().add_text(text))
            .add_run(Run::new().add_text(format!("  [Comment by {}: {}]", author, comment)));
        Ok(docx.add_paragraph(paragraph))
    }

    // Template helper methods
    
    fn apply_business_letter_template(&self, mut docx: Docx) -> Result<Docx> {
        // Add sender info placeholder
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[Your Name]"))
                .add_run(Run::new().add_break(BreakType::TextWrapping))
                .add_run(Run::new().add_text("[Your Address]"))
                .add_run(Run::new().add_break(BreakType::TextWrapping))
                .add_run(Run::new().add_text("[City, State ZIP]"))
                .add_run(Run::new().add_break(BreakType::TextWrapping))
                .add_run(Run::new().add_text("[Your Email]"))
                .add_run(Run::new().add_break(BreakType::TextWrapping))
                .add_run(Run::new().add_text("[Your Phone]"))
        );
        
        docx = docx.add_paragraph(Paragraph::new());
        
        // Date
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[Date]"))
        );
        
        docx = docx.add_paragraph(Paragraph::new());
        
        // Recipient info
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[Recipient Name]"))
                .add_run(Run::new().add_break(BreakType::TextWrapping))
                .add_run(Run::new().add_text("[Title]"))
                .add_run(Run::new().add_break(BreakType::TextWrapping))
                .add_run(Run::new().add_text("[Company]"))
                .add_run(Run::new().add_break(BreakType::TextWrapping))
                .add_run(Run::new().add_text("[Address]"))
                .add_run(Run::new().add_break(BreakType::TextWrapping))
                .add_run(Run::new().add_text("[City, State ZIP]"))
        );
        
        docx = docx.add_paragraph(Paragraph::new());
        
        // Salutation
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("Dear [Recipient Name]:"))
        );
        
        docx = docx.add_paragraph(Paragraph::new());
        
        // Body placeholder
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[Letter body paragraph 1]"))
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[Letter body paragraph 2]"))
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[Letter body paragraph 3]"))
        );
        
        docx = docx.add_paragraph(Paragraph::new());
        
        // Closing
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("Sincerely,"))
        );
        
        docx = docx.add_paragraph(Paragraph::new());
        docx = docx.add_paragraph(Paragraph::new());
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[Your Name]"))
        );
        
        Ok(docx)
    }
    
    fn apply_resume_template(&self, mut docx: Docx) -> Result<Docx> {
        // Name header
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[YOUR NAME]").size(32).bold())
                .align(AlignmentType::Center)
        );
        
        // Contact info
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[Email] | [Phone] | [LinkedIn] | [Location]").size(22))
                .align(AlignmentType::Center)
        );
        
        docx = docx.add_paragraph(Paragraph::new().add_run(Run::new().add_text("").size(12)));
        
        // Professional Summary
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("PROFESSIONAL SUMMARY").size(24).bold())
                .style("Heading2")
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[2-3 lines summarizing your experience and key skills]"))
        );
        
        // Experience
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("EXPERIENCE").size(24).bold())
                .style("Heading2")
        );
        
        // Education
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("EDUCATION").size(24).bold())
                .style("Heading2")
        );
        
        // Skills
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("SKILLS").size(24).bold())
                .style("Heading2")
        );
        
        Ok(docx)
    }
    
    fn apply_report_template(&self, mut docx: Docx) -> Result<Docx> {
        // Title page
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text(""))
                .add_run(Run::new().add_break(BreakType::TextWrapping))
                .add_run(Run::new().add_break(BreakType::TextWrapping))
                .add_run(Run::new().add_break(BreakType::TextWrapping))
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[REPORT TITLE]").size(36).bold())
                .align(AlignmentType::Center)
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[Subtitle or Description]").size(24))
                .align(AlignmentType::Center)
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_break(BreakType::TextWrapping))
                .add_run(Run::new().add_break(BreakType::TextWrapping))
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("Prepared by:").size(20))
                .align(AlignmentType::Center)
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[Author Name]").size(20))
                .align(AlignmentType::Center)
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[Date]").size(20))
                .align(AlignmentType::Center)
        );
        
        // Page break
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_break(BreakType::Page))
        );
        
        // Table of Contents placeholder
        docx = self.add_table_of_contents(docx)?;
        
        // Executive Summary
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("Executive Summary").size(28).bold())
                .style("Heading1")
        );
        
        Ok(docx)
    }
    
    fn apply_invoice_template(&self, mut docx: Docx) -> Result<Docx> {
        // Company header
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[COMPANY NAME]").size(32).bold())
                .align(AlignmentType::Right)
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("INVOICE").size(28).bold())
                .align(AlignmentType::Right)
        );
        
        // Invoice details table
        let mut invoice_info = Table::new(vec![])
        .add_row(TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("Invoice #:"))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("[INV-0001]"))),
        ]))
        .add_row(TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("Date:"))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("[Date]"))),
        ]))
        .add_row(TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("Due Date:"))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text("[Due Date]"))),
        ]));
        
        docx = docx.add_table(invoice_info);
        
        Ok(docx)
    }
    
    fn apply_contract_template(&self, mut docx: Docx) -> Result<Docx> {
        // Contract title
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[CONTRACT TYPE] AGREEMENT").size(28).bold())
                .align(AlignmentType::Center)
        );
        
        docx = docx.add_paragraph(Paragraph::new());
        
        // Parties
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("This Agreement is entered into as of [Date] between:"))
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[Party 1 Name], a [Entity Type] (\"Party 1\")"))
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("and"))
                .align(AlignmentType::Center)
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[Party 2 Name], a [Entity Type] (\"Party 2\")"))
        );
        
        Ok(docx)
    }
    
    fn apply_memo_template(&self, mut docx: Docx) -> Result<Docx> {
        // Memo header
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("MEMORANDUM").size(24).bold())
                .align(AlignmentType::Center)
        );
        
        docx = docx.add_paragraph(Paragraph::new());
        
        // Memo fields
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("TO: ").bold())
                .add_run(Run::new().add_text("[Recipient(s)]"))
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("FROM: ").bold())
                .add_run(Run::new().add_text("[Sender]"))
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("DATE: ").bold())
                .add_run(Run::new().add_text("[Date]"))
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("SUBJECT: ").bold())
                .add_run(Run::new().add_text("[Subject]"))
        );
        
        // Divider line
        let mut divider = Paragraph::new();
        for _ in 0..70 { divider = divider.add_run(Run::new().add_text("_")); }
        docx = docx.add_paragraph(divider);
        
        Ok(docx)
    }
    
    fn apply_newsletter_template(&self, mut docx: Docx) -> Result<Docx> {
        // Newsletter header
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[NEWSLETTER TITLE]").size(36).bold())
                .align(AlignmentType::Center)
        );
        
        docx = docx.add_paragraph(
            Paragraph::new()
                .add_run(Run::new().add_text("[Issue #] | [Date]").size(18))
                .align(AlignmentType::Center)
        );
        
        // Two-column layout requires section APIs; skip for now
        
        Ok(docx)
    }
}

// Supporting types

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum DocumentTemplate {
    BusinessLetter,
    Resume,
    Report,
    Invoice,
    Contract,
    Memo,
    Newsletter,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DocumentProperties {
    pub title: String,
    pub subject: String,
    pub author: String,
    pub keywords: Vec<String>,
    pub description: String,
    pub company: Option<String>,
    pub manager: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SectionConfig {
    pub page_size: PageSize,
    pub landscape: bool,
    pub margins: Margins,
    pub columns: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum PageSize {
    A4,
    Letter,
    Legal,
    A3,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Margins {
    pub top: f32,
    pub bottom: f32,
    pub left: f32,
    pub right: f32,
    pub header: f32,
    pub footer: f32,
}

impl Default for Margins {
    fn default() -> Self {
        Self {
            top: 25.4,      // 1 inch in mm
            bottom: 25.4,
            left: 25.4,
            right: 25.4,
            header: 12.7,   // 0.5 inch
            footer: 12.7,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum ChartType {
    Bar,
    Column,
    Line,
    Pie,
    Area,
    Scatter,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartData {
    pub title: String,
    pub categories: Vec<String>,
    pub series: Vec<ChartSeries>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartSeries {
    pub name: String,
    pub values: Vec<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum WatermarkStyle {
    Diagonal,
    Horizontal,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomStyle {
    pub id: String,
    pub name: String,
    pub based_on: Option<String>,
    pub font: Option<String>,
    pub size: Option<usize>,
    pub bold: bool,
    pub italic: bool,
    pub color: Option<String>,
    pub spacing: Option<StyleSpacing>,
    pub indent: Option<StyleIndent>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StyleSpacing {
    pub before: i32,
    pub after: i32,
    pub line: f32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct StyleIndent {
    pub left: i32,
    pub right: i32,
    pub first_line: i32,
}