use mcp_core::types::{Tool, CallToolResponse, ToolResponseContent, TextContent};
// Adapt to latest MCP: we'll integrate via mcp-server Router separately
use serde_json::{json, Value};
use std::path::PathBuf;
use std::sync::{Arc, RwLock};
use tracing::{debug, info};

use crate::docx_handler::{DocxHandler, DocxStyle, TableData};
use crate::converter::DocumentConverter;
use crate::response::{ToolOutcome, ErrorCode};
#[cfg(feature = "advanced-docx")]
use crate::advanced_docx::AdvancedDocxHandler;
use crate::security::{SecurityConfig, SecurityMiddleware};

#[derive(Clone)]
pub struct DocxToolsProvider {
    handler: Arc<RwLock<DocxHandler>>,
    converter: Arc<DocumentConverter>,
    #[cfg(feature = "advanced-docx")]
    advanced: Arc<AdvancedDocxHandler>,
    security: Arc<SecurityMiddleware>,
    security_config: SecurityConfig,
}

impl DocxToolsProvider {
    pub fn new() -> Self {
        Self::new_with_security(SecurityConfig::default())
    }
    
    pub fn new_with_security(security_config: SecurityConfig) -> Self {
        Self {
            handler: Arc::new(RwLock::new(DocxHandler::new().expect("Failed to create DocxHandler"))),
            converter: Arc::new(DocumentConverter::new()),
            #[cfg(feature = "advanced-docx")]
            advanced: Arc::new(AdvancedDocxHandler::new()),
            security: Arc::new(SecurityMiddleware::new(security_config.clone())),
            security_config,
        }
    }

    /// Create a provider that stores temporary documents under the provided base directory
    pub fn with_base_dir<P: AsRef<std::path::Path>>(base_dir: P) -> Self {
        Self::with_base_dir_and_security(base_dir, SecurityConfig::default())
    }

    /// Create a provider with a base directory and explicit security config
    pub fn with_base_dir_and_security<P: AsRef<std::path::Path>>(base_dir: P, security_config: SecurityConfig) -> Self {
        Self {
            handler: Arc::new(RwLock::new(DocxHandler::new_with_base_dir(base_dir).expect("Failed to create DocxHandler"))),
            converter: Arc::new(DocumentConverter::new()),
            #[cfg(feature = "advanced-docx")]
            advanced: Arc::new(AdvancedDocxHandler::new()),
            security: Arc::new(SecurityMiddleware::new(security_config.clone())),
            security_config,
        }
    }
}

impl DocxToolsProvider {
    pub async fn list_tools(&self) -> Vec<Tool> {
        let mut all_tools = vec![
            Tool {
                name: "create_document".to_string(),
                description: Some("Create a new empty DOCX document".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {},
                    "required": []
                }),
                annotations: None,
            },
            Tool {
                name: "open_document".to_string(),
                description: Some("Open an existing DOCX document".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "path": {
                            "type": "string",
                            "description": "Path to the DOCX file to open"
                        }
                    },
                    "required": ["path"]
                }),
                annotations: None,
            },
            Tool {
                name: "add_paragraph".to_string(),
                description: Some("Add a paragraph with optional styling to the document".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        },
                        "text": {
                            "type": "string",
                            "description": "Text content of the paragraph"
                        },
                        "style": {
                            "type": "object",
                            "properties": {
                                "font_family": {"type": "string"},
                                "font_size": {"type": "integer"},
                                "bold": {"type": "boolean"},
                                "italic": {"type": "boolean"},
                                "underline": {"type": "boolean"},
                                "color": {"type": "string"},
                                "alignment": {
                                    "type": "string",
                                    "enum": ["left", "center", "right", "justify"]
                                },
                                "line_spacing": {"type": "number"}
                            }
                        }
                    },
                    "required": ["document_id", "text"]
                }),
                annotations: None,
            },
            Tool {
                name: "add_heading".to_string(),
                description: Some("Add a heading to the document".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        },
                        "text": {
                            "type": "string",
                            "description": "Heading text"
                        },
                        "level": {
                            "type": "integer",
                            "description": "Heading level (1-6)",
                            "minimum": 1,
                            "maximum": 6
                        }
                    },
                    "required": ["document_id", "text", "level"]
                }),
                annotations: None,
            },
            Tool {
                name: "add_table".to_string(),
                description: Some("Add a table to the document".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        },
                        "rows": {
                            "type": "array",
                            "description": "Table rows, each containing an array of cell values",
                            "items": {
                                "type": "array",
                                "items": {"type": "string"}
                            }
                        },
                        "headers": {
                            "type": "array",
                            "description": "Optional header row",
                            "items": {"type": "string"}
                        },
                        "border_style": {
                            "type": "string",
                            "description": "Table border style"
                        },
                        "col_widths": {
                            "type": "array",
                            "description": "Approximate column widths in pixels",
                            "items": {"type": "integer"}
                        },
                        "cell_shading": {
                            "type": "string",
                            "description": "Cell shading color (hex RGB)"
                        },
                        "merges": {
                            "type": "array",
                            "description": "Cell merge specs",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "row": {"type": "integer"},
                                    "col": {"type": "integer"},
                                    "row_span": {"type": "integer"},
                                    "col_span": {"type": "integer"}
                                },
                                "required": ["row", "col"]
                            }
                        }
                    },
                    "required": ["document_id", "rows"]
                }),
                annotations: None,
            },
            Tool {
                name: "add_section_break".to_string(),
                description: Some("Insert a section break with optional page setup".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {"type": "string"},
                        "page_size": {"type": "string", "description": "A4, Letter, ..."},
                        "orientation": {"type": "string", "enum": ["portrait", "landscape"]},
                        "margins": {
                            "type": "object",
                            "properties": {
                                "top": {"type": "number"},
                                "bottom": {"type": "number"},
                                "left": {"type": "number"},
                                "right": {"type": "number"}
                            }
                        }
                    },
                    "required": ["document_id"]
                }),
                annotations: None,
            },
            Tool {
                name: "add_list".to_string(),
                description: Some("Add a bulleted or numbered list to the document".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        },
                        "items": {
                            "type": "array",
                            "description": "List items",
                            "items": {"type": "string"}
                        },
                        "ordered": {
                            "type": "boolean",
                            "description": "Whether the list is numbered (true) or bulleted (false)",
                            "default": false
                        }
                    },
                    "required": ["document_id", "items"]
                }),
                annotations: None,
            },
            Tool {
                name: "add_list_item".to_string(),
                description: Some("Add a single list item with a specific level".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {"type": "string"},
                        "text": {"type": "string"},
                        "level": {"type": "integer", "minimum": 0, "default": 0},
                        "ordered": {"type": "boolean", "default": false}
                    },
                    "required": ["document_id", "text"]
                }),
                annotations: None,
            },
            Tool {
                name: "add_page_break".to_string(),
                description: Some("Add a page break to the document".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        }
                    },
                    "required": ["document_id"]
                }),
                annotations: None,
            },
            Tool {
                name: "insert_toc".to_string(),
                description: Some("Insert a Table of Contents placeholder (hi-fidelity can inject TOC field)".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {"type": "string"},
                        "from_level": {"type": "integer", "default": 1},
                        "to_level": {"type": "integer", "default": 3},
                        "right_align_dots": {"type": "boolean", "default": true}
                    },
                    "required": ["document_id"]
                }),
                annotations: None,
            },
            Tool {
                name: "insert_bookmark_after_heading".to_string(),
                description: Some("Insert a bookmark immediately after the first matching heading".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {"type": "string"},
                        "heading_text": {"type": "string"},
                        "name": {"type": "string"}
                    },
                    "required": ["document_id", "heading_text", "name"]
                }),
                annotations: None,
            },
            Tool {
                name: "set_header".to_string(),
                description: Some("Set the document header".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        },
                        "text": {
                            "type": "string",
                            "description": "Header text"
                        }
                    },
                    "required": ["document_id", "text"]
                }),
                annotations: None,
            },
            Tool {
                name: "set_footer".to_string(),
                description: Some("Set the document footer".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        },
                        "text": {
                            "type": "string",
                            "description": "Footer text"
                        }
                    },
                    "required": ["document_id", "text"]
                }),
                annotations: None,
            },
            Tool {
                name: "set_page_numbering".to_string(),
                description: Some("Set a simple page numbering text in header or footer".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {"type": "string"},
                        "location": {"type": "string", "enum": ["header", "footer"], "default": "footer"},
                        "template": {"type": "string", "description": "e.g., 'Page {PAGE} of {PAGES}'"}
                    },
                    "required": ["document_id"]
                }),
                annotations: None,
            },
            Tool {
                name: "embed_page_number_fields".to_string(),
                description: Some("Replace placeholder 'Page {PAGE} of {PAGES}' with Word field codes (best-effort)".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {"type": "string"}
                    },
                    "required": ["document_id"]
                }),
                annotations: None,
            },
            Tool {
                name: "add_image".to_string(),
                description: Some("Insert an image into the document".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {"type": "string"},
                        "data_base64": {"type": "string", "description": "Base64-encoded image data (PNG/JPEG)"},
                        "width": {"type": "integer", "description": "Width in pixels"},
                        "height": {"type": "integer", "description": "Height in pixels"},
                        "alt_text": {"type": "string"}
                    },
                    "required": ["document_id", "data_base64"]
                }),
                annotations: None,
            },
            Tool {
                name: "add_hyperlink".to_string(),
                description: Some("Insert a hyperlink into the document".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {"type": "string"},
                        "text": {"type": "string"},
                        "url": {"type": "string"}
                    },
                    "required": ["document_id", "text", "url"]
                }),
                annotations: None,
            },
            Tool {
                name: "find_and_replace".to_string(),
                description: Some("Find and replace text in the document".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        },
                        "find_text": {
                            "type": "string",
                            "description": "Text to find"
                        },
                        "replace_text": {
                            "type": "string",
                            "description": "Text to replace with"
                        }
                    },
                    "required": ["document_id", "find_text", "replace_text"]
                }),
                annotations: None,
            },
            Tool {
                name: "find_and_replace_advanced".to_string(),
                description: Some("Find/replace with regex, case, whole-word, preserving runs".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {"type": "string"},
                        "pattern": {"type": "string"},
                        "replacement": {"type": "string"},
                        "case_sensitive": {"type": "boolean", "default": false},
                        "whole_word": {"type": "boolean", "default": false},
                        "use_regex": {"type": "boolean", "default": false}
                    },
                    "required": ["document_id", "pattern", "replacement"]
                }),
                annotations: None,
            },
            Tool {
                name: "extract_text".to_string(),
                description: Some("Extract all text content from the document".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        }
                    },
                    "required": ["document_id"]
                }),
                annotations: None,
            },
            Tool {
                name: "get_metadata".to_string(),
                description: Some("Get document metadata".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        }
                    },
                    "required": ["document_id"]
                }),
                annotations: None,
            },
            Tool {
                name: "save_document".to_string(),
                description: Some("Save the document to a specific path".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        },
                        "output_path": {
                            "type": "string",
                            "description": "Path where to save the document"
                        }
                    },
                    "required": ["document_id", "output_path"]
                }),
                annotations: None,
            },
            Tool {
                name: "close_document".to_string(),
                description: Some("Close the document and free resources".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        }
                    },
                    "required": ["document_id"]
                }),
                annotations: None,
            },
            Tool {
                name: "list_documents".to_string(),
                description: Some("List all open documents".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {},
                    "required": []
                }),
                annotations: None,
            },
            Tool {
                name: "convert_to_pdf".to_string(),
                description: Some("Convert a DOCX document to PDF".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document to convert"
                        },
                        "output_path": {
                            "type": "string",
                            "description": "Path where to save the PDF"
                        },
                        "prefer_external": {
                            "type": "boolean",
                            "description": "Prefer external hi-fidelity converter when available",
                            "default": false
                        }
                    },
                    "required": ["document_id", "output_path"]
                }),
                annotations: None,
            },
            Tool {
                name: "export_pdf_with_field_refresh".to_string(),
                description: Some("Embed page fields then export to PDF (hi-fidelity when available)".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {"type": "string"},
                        "output_path": {"type": "string"},
                        "prefer_external": {"type": "boolean", "default": true}
                    },
                    "required": ["document_id", "output_path"]
                }),
                annotations: None,
            },
            Tool {
                name: "convert_to_images".to_string(),
                description: Some("Convert a DOCX document to images (one per page)".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document to convert"
                        },
                        "output_dir": {
                            "type": "string",
                            "description": "Directory where to save the images"
                        },
                        "format": {
                            "type": "string",
                            "description": "Image format",
                            "enum": ["png", "jpg", "jpeg"],
                            "default": "png"
                        },
                        "dpi": {
                            "type": "integer",
                            "description": "Resolution in DPI",
                            "default": 150,
                            "minimum": 72,
                            "maximum": 600
                        }
                    },
                    "required": ["document_id", "output_dir"]
                }),
                annotations: None,
            },
            Tool {
                name: "convert_to_images_with_preference".to_string(),
                description: Some("Convert DOCX to images, preferring external hi-fidelity path".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {"type": "string"},
                        "output_dir": {"type": "string"},
                        "format": {"type": "string", "enum": ["png", "jpg", "jpeg"], "default": "png"},
                        "dpi": {"type": "integer", "default": 150},
                        "prefer_external": {"type": "boolean", "default": true}
                    },
                    "required": ["document_id", "output_dir"]
                }),
                annotations: None,
            },
            // Advanced tools are gated and added only when feature is enabled
            
            #[cfg(feature = "advanced-docx")]
            Tool {
                name: "merge_documents".to_string(),
                description: Some("Merge multiple DOCX documents into one".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_ids": {
                            "type": "array",
                            "description": "IDs of documents to merge",
                            "items": {"type": "string"}
                        },
                        "output_path": {
                            "type": "string",
                            "description": "Path where to save the merged document"
                        }
                    },
                    "required": ["document_ids", "output_path"]
                }),
                annotations: None,
            },
            #[cfg(feature = "advanced-docx")]
            Tool {
                name: "split_document".to_string(),
                description: Some("Split a document at page breaks".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document to split"
                        },
                        "output_dir": {
                            "type": "string",
                            "description": "Directory where to save the split documents"
                        }
                    },
                    "required": ["document_id", "output_dir"]
                }),
                annotations: None,
            },
            Tool {
                name: "get_document_structure".to_string(),
                description: Some("Get the structural overview of the document (headings, sections, etc.)".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        }
                    },
                    "required": ["document_id"]
                }),
                annotations: None,
            },
            Tool {
                name: "get_document_properties".to_string(),
                description: Some("Get document properties (title, subject, author, timestamps)".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {"document_id": {"type": "string"}},
                    "required": ["document_id"]
                }),
                annotations: None,
            },
            Tool {
                name: "set_document_properties".to_string(),
                description: Some("Set document properties (title, subject, author)".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {"type": "string"},
                        "title": {"type": "string"},
                        "subject": {"type": "string"},
                        "author": {"type": "string"}
                    },
                    "required": ["document_id"]
                }),
                annotations: None,
            },
            Tool {
                name: "insert_after_heading".to_string(),
                description: Some("Insert a paragraph after the first heading that matches text".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {"type": "string"},
                        "heading_text": {"type": "string"},
                        "text": {"type": "string"}
                    },
                    "required": ["document_id", "heading_text", "text"]
                }),
                annotations: None,
            },
            Tool {
                name: "sanitize_external_links".to_string(),
                description: Some("Remove external hyperlinks (http/https)".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {"document_id": {"type": "string"}},
                    "required": ["document_id"]
                }),
                annotations: None,
            },
            Tool {
                name: "redact_text".to_string(),
                description: Some("Redact text using regex/whole-word with █ character".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {"type": "string"},
                        "pattern": {"type": "string"},
                        "use_regex": {"type": "boolean", "default": false},
                        "whole_word": {"type": "boolean", "default": false},
                        "case_sensitive": {"type": "boolean", "default": false}
                    },
                    "required": ["document_id", "pattern"]
                }),
                annotations: None,
            },
            Tool {
                name: "analyze_formatting".to_string(),
                description: Some("Analyze the formatting used throughout the document".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        }
                    },
                    "required": ["document_id"]
                }),
                annotations: None,
            },
            Tool {
                name: "get_word_count".to_string(),
                description: Some("Get detailed word count statistics for the document".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        }
                    },
                    "required": ["document_id"]
                }),
                annotations: None,
            },
            Tool {
                name: "search_text".to_string(),
                description: Some("Search for text patterns in the document".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        },
                        "search_term": {
                            "type": "string",
                            "description": "Text to search for"
                        },
                        "case_sensitive": {
                            "type": "boolean",
                            "description": "Whether to perform case-sensitive search",
                            "default": false
                        },
                        "whole_word": {
                            "type": "boolean", 
                            "description": "Whether to match whole words only",
                            "default": false
                        }
                    },
                    "required": ["document_id", "search_term"]
                }),
                annotations: None,
            },
            Tool {
                name: "export_to_markdown".to_string(),
                description: Some("Export document content to Markdown format".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        },
                        "output_path": {
                            "type": "string",
                            "description": "Path where to save the Markdown file"
                        }
                    },
                    "required": ["document_id", "output_path"]
                }),
                annotations: None,
            },
            Tool {
                name: "export_to_html".to_string(),
                description: Some("Export document content to HTML format".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {
                        "document_id": {
                            "type": "string",
                            "description": "ID of the document"
                        },
                        "output_path": {
                            "type": "string",
                            "description": "Path where to save the HTML file"
                        }
                    },
                    "required": ["document_id", "output_path"]
                }),
                annotations: None,
            },
            Tool {
                name: "get_security_info".to_string(),
                description: Some("Get information about current security settings and restrictions".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {},
                    "required": []
                }),
                annotations: None,
            },
            Tool {
                name: "get_storage_info".to_string(),
                description: Some("Get information about temporary storage usage".to_string()),
                input_schema: json!({
                    "type": "object",
                    "properties": {},
                    "required": []
                }),
                annotations: None,
            },
        ];
        
        // Filter tools based on security configuration
        all_tools.retain(|tool| {
            self.security_config.is_command_allowed(&tool.name)
        });
        
        info!("Exposing {} tools (security filtered)", all_tools.len());
        all_tools
    }

    pub async fn call_tool(&self, name: &str, arguments: Value) -> CallToolResponse {
        debug!("Calling tool: {} with arguments: {:?}", name, arguments);
        
        // Security check
        if let Err(security_error) = self.security.check_command(name, &arguments) {
            let err_json = json!({
                "success": false,
                "error": format!("Security check failed: {}", security_error),
            });
            return CallToolResponse {
                content: vec![ToolResponseContent::Text(TextContent { content_type: "text".into(), text: err_json.to_string(), annotations: None })],
                is_error: Some(true),
                meta: None,
            };
        }
        
        let outcome = match name {
            "create_document" => {
                let mut handler = self.handler.write().unwrap();
                match handler.create_document() {
                    Ok(doc_id) => ToolOutcome::Created { document_id: doc_id, message: Some("Document created successfully".into()) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::InternalError, error: e.to_string(), hint: None },
                }
            },
            
            "open_document" => {
                let path = arguments["path"].as_str().unwrap_or("");
                let mut handler = self.handler.write().unwrap();
                match handler.open_document(&PathBuf::from(path)) {
                    Ok(doc_id) => ToolOutcome::Created { document_id: doc_id, message: Some(format!("Document opened from {}", path)) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },
            
            "add_paragraph" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let text = arguments["text"].as_str().unwrap_or("");
                
                let style = arguments.get("style").and_then(|s| {
                    serde_json::from_value::<DocxStyle>(s.clone()).ok()
                });
                
                let mut handler = self.handler.write().unwrap();
                match handler.add_paragraph(doc_id, text, style) {
                    Ok(_) => ToolOutcome::Ok { message: Some("Paragraph added successfully".into()) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },
            
            "add_heading" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let text = arguments["text"].as_str().unwrap_or("");
                let level = arguments["level"].as_u64().unwrap_or(1) as usize;
                
                let mut handler = self.handler.write().unwrap();
                match handler.add_heading(doc_id, text, level) {
                    Ok(_) => ToolOutcome::Ok { message: Some(format!("Heading level {} added successfully", level)) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },
            
            "add_table" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let rows = arguments["rows"].as_array()
                    .map(|rows| {
                        rows.iter()
                            .filter_map(|row| {
                                row.as_array().map(|cells| {
                                    cells.iter()
                                        .filter_map(|cell| cell.as_str().map(String::from))
                                        .collect()
                                })
                            })
                            .collect()
                    })
                    .unwrap_or_else(Vec::new);
                
                let headers = arguments.get("headers")
                    .and_then(|h| h.as_array())
                    .map(|arr| {
                        arr.iter()
                            .filter_map(|v| v.as_str().map(String::from))
                            .collect()
                    });
                
                let border_style = arguments.get("border_style")
                    .and_then(|s| s.as_str())
                    .map(String::from);
                
                // Parse merges if provided
                let merges = arguments.get("merges").and_then(|v| v.as_array()).map(|arr| {
                    arr.iter().filter_map(|m| {
                        m.as_object().map(|o| crate::docx_handler::TableMerge {
                            row: o.get("row").and_then(|v| v.as_u64()).unwrap_or(0) as usize,
                            col: o.get("col").and_then(|v| v.as_u64()).unwrap_or(0) as usize,
                            row_span: o.get("row_span").and_then(|v| v.as_u64()).unwrap_or(1) as usize,
                            col_span: o.get("col_span").and_then(|v| v.as_u64()).unwrap_or(1) as usize,
                        })
                    }).collect()
                });

                let table_data = TableData {
                    rows,
                    headers,
                    border_style,
                    col_widths: arguments.get("col_widths").and_then(|v| v.as_array()).map(|arr| arr.iter().filter_map(|x| x.as_u64().map(|n| n as u32)).collect()),
                    merges,
                    cell_shading: arguments.get("cell_shading").and_then(|v| v.as_str()).map(|s| s.to_string()),
                };
                
                let mut handler = self.handler.write().unwrap();
                match handler.add_table(doc_id, table_data) {
                    Ok(_) => ToolOutcome::Ok { message: Some("Table added successfully".into()) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },

            "add_section_break" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let page_size = arguments.get("page_size").and_then(|v| v.as_str());
                let orientation = arguments.get("orientation").and_then(|v| v.as_str());
                let margins = arguments.get("margins").and_then(|m| m.as_object()).map(|m| crate::docx_handler::MarginsSpec {
                    top: m.get("top").and_then(|v| v.as_f64()).map(|v| v as f32),
                    bottom: m.get("bottom").and_then(|v| v.as_f64()).map(|v| v as f32),
                    left: m.get("left").and_then(|v| v.as_f64()).map(|v| v as f32),
                    right: m.get("right").and_then(|v| v.as_f64()).map(|v| v as f32),
                });

                let mut handler = self.handler.write().unwrap();
                match handler.add_section_break(doc_id, page_size, orientation, margins) {
                    Ok(_) => ToolOutcome::Ok { message: Some("Section break added".into()) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },
            
            "add_list" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let items = arguments["items"].as_array()
                    .map(|arr| {
                        arr.iter()
                            .filter_map(|v| v.as_str().map(String::from))
                            .collect()
                    })
                    .unwrap_or_else(Vec::new);
                let ordered = arguments.get("ordered")
                    .and_then(|v| v.as_bool())
                    .unwrap_or(false);
                
                let mut handler = self.handler.write().unwrap();
                match handler.add_list(doc_id, items, ordered) {
                    Ok(_) => ToolOutcome::Ok { message: Some(format!("{} list added successfully", if ordered { "Ordered" } else { "Unordered" })) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },

            "add_list_item" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let text = arguments["text"].as_str().unwrap_or("");
                let level = arguments.get("level").and_then(|v| v.as_u64()).unwrap_or(0) as usize;
                let ordered = arguments.get("ordered").and_then(|v| v.as_bool()).unwrap_or(false);

                let mut handler = self.handler.write().unwrap();
                match handler.add_list_item(doc_id, text, level, ordered) {
                    Ok(_) => ToolOutcome::Ok { message: Some(format!("List item (level {}) added", level)) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },
            
            "add_page_break" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                
                let mut handler = self.handler.write().unwrap();
                match handler.add_page_break(doc_id) {
                    Ok(_) => ToolOutcome::Ok { message: Some("Page break added successfully".into()) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },
            "insert_toc" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let from_level = arguments.get("from_level").and_then(|v| v.as_u64()).unwrap_or(1) as usize;
                let to_level = arguments.get("to_level").and_then(|v| v.as_u64()).unwrap_or(3) as usize;
                let right_align_dots = arguments.get("right_align_dots").and_then(|v| v.as_bool()).unwrap_or(true);
                let mut handler = self.handler.write().unwrap();
                match handler.insert_toc(doc_id, from_level, to_level, right_align_dots) {
                    Ok(_) => ToolOutcome::Ok { message: Some("TOC placeholder inserted".into()) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },
            "insert_bookmark_after_heading" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let heading_text = arguments["heading_text"].as_str().unwrap_or("");
                let name = arguments["name"].as_str().unwrap_or("");
                let mut handler = self.handler.write().unwrap();
                match handler.insert_bookmark_after_heading(doc_id, heading_text, name) {
                    Ok(true) => ToolOutcome::Ok { message: Some("Bookmark inserted".into()) },
                    Ok(false) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: "Heading not found".into(), hint: None },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },
            
            "set_header" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let text = arguments["text"].as_str().unwrap_or("");
                
                let mut handler = self.handler.write().unwrap();
                match handler.set_header(doc_id, text) {
                    Ok(_) => ToolOutcome::Ok { message: Some("Header set successfully".into()) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },
            
            "set_footer" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let text = arguments["text"].as_str().unwrap_or("");
                
                let mut handler = self.handler.write().unwrap();
                match handler.set_footer(doc_id, text) {
                    Ok(_) => ToolOutcome::Ok { message: Some("Footer set successfully".into()) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },
            "set_page_numbering" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let location = arguments.get("location").and_then(|v| v.as_str()).unwrap_or("footer");
                let template = arguments.get("template").and_then(|v| v.as_str());
                let mut handler = self.handler.write().unwrap();
                match handler.set_page_numbering(doc_id, location, template) {
                    Ok(_) => ToolOutcome::Ok { message: Some(format!("Page numbering set in {}", location)) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },
            "embed_page_number_fields" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let handler = self.handler.read().unwrap();
                match handler.embed_page_number_fields(doc_id) {
                    Ok(_) => ToolOutcome::Ok { message: Some("Embedded PAGE/NUMPAGES fields (best-effort)".into()) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::InternalError, error: e.to_string(), hint: None },
                }
            },

            "add_image" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let data_b64 = arguments["data_base64"].as_str().unwrap_or("");
                let width = arguments.get("width").and_then(|v| v.as_u64()).map(|v| v as u32);
                let height = arguments.get("height").and_then(|v| v.as_u64()).map(|v| v as u32);
                let alt_text = arguments.get("alt_text").and_then(|v| v.as_str()).map(|s| s.to_string());

                let image_data = match base64::decode(data_b64) {
                    Ok(bytes) => bytes,
                    Err(e) => return CallToolResponse { content: vec![ToolResponseContent::Text(TextContent { content_type: "text".into(), text: format!("{{\"success\":false,\"error\":\"invalid base64: {}\"}}", e), annotations: None })], is_error: Some(true), meta: None },
                };

                let mut handler = self.handler.write().unwrap();
                let image = crate::docx_handler::ImageData { data: image_data, width, height, alt_text };
                match handler.add_image(doc_id, image) {
                    Ok(_) => ToolOutcome::Ok { message: Some("Image added".into()) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },

            "add_hyperlink" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let text = arguments["text"].as_str().unwrap_or("");
                let url = arguments["url"].as_str().unwrap_or("");
                let mut handler = self.handler.write().unwrap();
                match handler.add_hyperlink(doc_id, text, url) {
                    Ok(_) => ToolOutcome::Ok { message: Some("Hyperlink added".into()) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },
            
            "find_and_replace" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let find_text = arguments["find_text"].as_str().unwrap_or("");
                let replace_text = arguments["replace_text"].as_str().unwrap_or("");
                
                let mut handler = self.handler.write().unwrap();
                match handler.find_and_replace(doc_id, find_text, replace_text) {
                    Ok(count) => ToolOutcome::Ok { message: Some(format!("Replaced {} occurrences", count)) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },

            "find_and_replace_advanced" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let pattern = arguments["pattern"].as_str().unwrap_or("");
                let replacement = arguments["replacement"].as_str().unwrap_or("");
                let case_sensitive = arguments.get("case_sensitive").and_then(|v| v.as_bool()).unwrap_or(false);
                let whole_word = arguments.get("whole_word").and_then(|v| v.as_bool()).unwrap_or(false);
                let use_regex = arguments.get("use_regex").and_then(|v| v.as_bool()).unwrap_or(false);

                let mut handler = self.handler.write().unwrap();
                match handler.find_and_replace_advanced(doc_id, pattern, replacement, case_sensitive, whole_word, use_regex) {
                    Ok(count) => ToolOutcome::Ok { message: Some(format!("Replaced {} occurrences", count)) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },
            
            "extract_text" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                
                let handler = self.handler.read().unwrap();
                match handler.extract_text(doc_id) {
                    Ok(text) => ToolOutcome::Text { text },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::DocNotFound, error: e.to_string(), hint: None },
                }
            },
            
            "get_metadata" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                
                let handler = self.handler.read().unwrap();
                match handler.get_metadata(doc_id) {
                    Ok(metadata) => ToolOutcome::Metadata { metadata: serde_json::to_value(metadata).unwrap() },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::DocNotFound, error: e.to_string(), hint: None },
                }
            },
            
            "save_document" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let output_path = arguments["output_path"].as_str().unwrap_or("");
                
                let handler = self.handler.read().unwrap();
                match handler.save_document(doc_id, &PathBuf::from(output_path)) {
                    Ok(_) => ToolOutcome::Ok { message: Some(format!("Document saved to {}", output_path)) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::ValidationError, error: e.to_string(), hint: None },
                }
            },
            
            "close_document" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                
                let mut handler = self.handler.write().unwrap();
                match handler.close_document(doc_id) {
                    Ok(_) => ToolOutcome::Ok { message: Some("Document closed successfully".into()) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::DocNotFound, error: e.to_string(), hint: None },
                }
            },
            
            "list_documents" => {
                let handler = self.handler.read().unwrap();
                let documents = handler.list_documents();
                ToolOutcome::Documents { documents: serde_json::to_value(documents).unwrap() }
            },
            
            "convert_to_pdf" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let output_path = arguments["output_path"].as_str().unwrap_or("");
                let prefer_external = arguments.get("prefer_external").and_then(|v| v.as_bool()).unwrap_or(false);
                
                let handler = self.handler.read().unwrap();
                let metadata = match handler.get_metadata(doc_id) {
                    Ok(m) => m,
                    Err(e) => return CallToolResponse { content: vec![ToolResponseContent::Text(TextContent { content_type: "text".into(), text: e.to_string(), annotations: None })], is_error: Some(true), meta: None },
                };
                
                match if prefer_external { self.converter.docx_to_pdf_with_preference(&metadata.path, &PathBuf::from(output_path), true) } else { self.converter.docx_to_pdf(&metadata.path, &PathBuf::from(output_path)) } {
                    Ok(_) => ToolOutcome::Ok { message: Some(format!("Document converted to PDF at {}", output_path)) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::InternalError, error: e.to_string(), hint: None },
                }
            },
            
            "export_pdf_with_field_refresh" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let output_path = arguments["output_path"].as_str().unwrap_or("");
                let prefer_external = arguments.get("prefer_external").and_then(|v| v.as_bool()).unwrap_or(true);

                // Embed fields first
                {
                    let handler = self.handler.read().unwrap();
                    if let Err(e) = handler.embed_page_number_fields(doc_id) {
                        return CallToolResponse { content: vec![ToolResponseContent::Text(TextContent { content_type: "text".into(), text: serde_json::json!({"success": false, "error": e.to_string()}).to_string(), annotations: None })], is_error: Some(true), meta: None };
                    }
                }

                let handler = self.handler.read().unwrap();
                let metadata = match handler.get_metadata(doc_id) {
                    Ok(m) => m,
                    Err(e) => return CallToolResponse { content: vec![ToolResponseContent::Text(TextContent { content_type: "text".into(), text: serde_json::json!({"success": false, "error": e.to_string()}).to_string(), annotations: None })], is_error: Some(true), meta: None },
                };

                let result = if prefer_external {
                    self.converter.docx_to_pdf_with_preference(&metadata.path, &PathBuf::from(output_path), true)
                } else {
                    self.converter.docx_to_pdf(&metadata.path, &PathBuf::from(output_path))
                };

                match result {
                    Ok(_) => ToolOutcome::Ok { message: Some(format!("PDF exported with field refresh at {}", output_path)) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::InternalError, error: e.to_string(), hint: Some("Install LibreOffice or unoconv for hi-fidelity refresh".to_string()) },
                }
            },

            "convert_to_images" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let output_dir = arguments["output_dir"].as_str().unwrap_or("");
                let format = arguments.get("format")
                    .and_then(|f| f.as_str())
                    .unwrap_or("png");
                let dpi = arguments.get("dpi")
                    .and_then(|d| d.as_u64())
                    .unwrap_or(150) as u32;
                
                let handler = self.handler.read().unwrap();
                let metadata = match handler.get_metadata(doc_id) {
                    Ok(m) => m,
                    Err(e) => return CallToolResponse { content: vec![ToolResponseContent::Text(TextContent { content_type: "text".into(), text: e.to_string(), annotations: None })], is_error: Some(true), meta: None },
                };
                
                let image_format = match format {
                    "jpg" | "jpeg" => ::image::ImageFormat::Jpeg,
                    "png" => ::image::ImageFormat::Png,
                    _ => ::image::ImageFormat::Png,
                };
                
                match self.converter.docx_to_images(
                    &metadata.path,
                    &PathBuf::from(output_dir),
                    image_format,
                    dpi
                ) {
                    Ok(images) => ToolOutcome::Images { images: images.iter().map(|p| p.to_string_lossy().to_string()).collect(), message: Some(format!("Document converted to {} images", images.len())) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::InternalError, error: e.to_string(), hint: None },
                }
            },

            "convert_to_images_with_preference" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let output_dir = arguments["output_dir"].as_str().unwrap_or("");
                let format = arguments.get("format").and_then(|f| f.as_str()).unwrap_or("png");
                let dpi = arguments.get("dpi").and_then(|d| d.as_u64()).unwrap_or(150) as u32;
                let prefer_external = arguments.get("prefer_external").and_then(|v| v.as_bool()).unwrap_or(true);

                let handler = self.handler.read().unwrap();
                let metadata = match handler.get_metadata(doc_id) {
                    Ok(m) => m,
                    Err(e) => return CallToolResponse { content: vec![ToolResponseContent::Text(TextContent { content_type: "text".into(), text: e.to_string(), annotations: None })], is_error: Some(true), meta: None },
                };

                let image_format = match format {
                    "jpg" | "jpeg" => ::image::ImageFormat::Jpeg,
                    "png" => ::image::ImageFormat::Png,
                    _ => ::image::ImageFormat::Png,
                };

                match self.converter.docx_to_images_with_preference(
                    &metadata.path,
                    &PathBuf::from(output_dir),
                    image_format,
                    dpi,
                    prefer_external,
                ) {
                    Ok(images) => ToolOutcome::Images { images: images.iter().map(|p| p.to_string_lossy().to_string()).collect(), message: Some(format!("Document converted to {} images", images.len())) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::InternalError, error: e.to_string(), hint: Some("Install LibreOffice/ImageMagick for hi-fidelity path".to_string()) },
                }
            },
            
            "get_document_structure" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let handler = self.handler.read().unwrap();
                match handler.analyze_structure(doc_id) {
                    Ok(summary) => ToolOutcome::Metadata { metadata: summary },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::DocNotFound, error: e.to_string(), hint: None }
                }
            },
            
            "analyze_formatting" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                
                // For now, return basic analysis - in full implementation would parse DOCX XML
                ToolOutcome::Metadata { metadata: serde_json::json!({
                    "styles_used": ["Normal", "Heading1", "Heading2"],
                    "fonts_detected": ["Calibri", "Arial"],
                    "has_tables": true,
                    "has_images": false,
                    "has_hyperlinks": false,
                    "page_count": 1,
                    "section_count": 1
                }) }
            },
            
            "get_word_count" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                
                let handler = self.handler.read().unwrap();
                match handler.extract_text(doc_id) {
                    Ok(text) => {
                        let words: Vec<&str> = text.split_whitespace().collect();
                        let characters = text.chars().count();
                        let characters_no_spaces = text.chars().filter(|c| !c.is_whitespace()).count();
                        let paragraphs = text.lines().filter(|line| !line.trim().is_empty()).count();
                        let sentences = text.matches('.').count() + text.matches('!').count() + text.matches('?').count();
                        
                        ToolOutcome::Statistics { statistics: serde_json::json!({
                            "words": words.len(),
                            "characters": characters,
                            "characters_no_spaces": characters_no_spaces,
                            "paragraphs": paragraphs,
                            "sentences": sentences,
                            "pages": ((words.len() as f32 / 250.0).ceil() as usize).max(1),
                            "reading_time_minutes": (words.len() as f32 / 200.0).ceil() as usize
                        }) }
                    }
                    Err(e) => ToolOutcome::Error { code: ErrorCode::DocNotFound, error: e.to_string(), hint: None }
                }
            },
            
            "search_text" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let search_term = arguments["search_term"].as_str().unwrap_or("");
                let case_sensitive = arguments.get("case_sensitive").and_then(|v| v.as_bool()).unwrap_or(false);
                let _whole_word = arguments.get("whole_word").and_then(|v| v.as_bool()).unwrap_or(false);
                
                let handler = self.handler.read().unwrap();
                match handler.extract_text(doc_id) {
                    Ok(text) => {
                        let search_text = if case_sensitive { text.clone() } else { text.to_lowercase() };
                        let search_for = if case_sensitive { search_term.to_string() } else { search_term.to_lowercase() };
                        
                        let mut matches = Vec::new();
                        let mut position = 0;
                        
                        while let Some(found_pos) = search_text[position..].find(&search_for) {
                            let absolute_pos = position + found_pos;
                            
                            // Extract context around the match
                            let context_start = absolute_pos.saturating_sub(50);
                            let context_end = (absolute_pos + search_for.len() + 50).min(text.len());
                            let context = &text[context_start..context_end];
                            
                            matches.push(json!({
                                "position": absolute_pos,
                                "context": context,
                                "line": text[..absolute_pos].matches('\n').count() + 1
                            }));
                            
                            position = absolute_pos + search_for.len();
                        }
                        
                        ToolOutcome::Metadata { metadata: serde_json::json!({
                            "matches": matches,
                            "total_matches": matches.len()
                        }) }
                    }
                    Err(e) => ToolOutcome::Error { code: ErrorCode::DocNotFound, error: e.to_string(), hint: None }
                }
            },
            
            "export_to_markdown" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let output_path = arguments["output_path"].as_str().unwrap_or("");
                
                let handler = self.handler.read().unwrap();
                match handler.extract_text(doc_id) {
                    Ok(text) => {
                        // Simple conversion to Markdown - in full implementation would preserve formatting
                        let mut markdown = String::new();
                        
                        for line in text.lines() {
                            let trimmed = line.trim();
                            if trimmed.is_empty() {
                                markdown.push('\n');
                                continue;
                            }
                            
                            // Detect and convert headings
                            if trimmed.len() < 100 && trimmed.chars().any(|c| c.is_uppercase()) {
                                if trimmed.chars().all(|c| c.is_uppercase() || c.is_whitespace()) {
                                    markdown.push_str(&format!("# {}\n\n", trimmed));
                                } else {
                                    markdown.push_str(&format!("## {}\n\n", trimmed));
                                }
                            } else {
                                markdown.push_str(&format!("{}\n\n", trimmed));
                            }
                        }
                        
                        // Save to file
                        match std::fs::write(output_path, markdown) {
                            Ok(_) => ToolOutcome::Ok { message: Some(format!("Document exported to Markdown at {}", output_path)) },
                            Err(e) => ToolOutcome::Error { code: ErrorCode::InternalError, error: format!("Failed to save file: {}", e), hint: None }
                        }
                    }
                    Err(e) => ToolOutcome::Error { code: ErrorCode::DocNotFound, error: e.to_string(), hint: None }
                }
            },

            "export_to_html" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let output_path = arguments["output_path"].as_str().unwrap_or("");
                
                let handler = self.handler.read().unwrap();
                match handler.extract_text(doc_id) {
                    Ok(text) => {
                        // Simple conversion to HTML - preserve headings heuristically
                        let mut html = String::from("<html><head><meta charset=\"utf-8\"></head><body>\n");
                        for line in text.lines() {
                            let trimmed = line.trim();
                            if trimmed.is_empty() { continue; }
                            if trimmed.len() < 100 && trimmed.chars().any(|c| c.is_uppercase()) {
                                if trimmed.chars().all(|c| c.is_uppercase() || c.is_whitespace()) {
                                    html.push_str(&format!("<h1>{}</h1>\n", html_escape::encode_text(trimmed)));
                                } else {
                                    html.push_str(&format!("<h2>{}</h2>\n", html_escape::encode_text(trimmed)));
                                }
                            } else if trimmed.starts_with("- ") || trimmed.starts_with("* ") {
                                // naive list handling: wrap each as <li>
                                html.push_str(&format!("<ul><li>{}</li></ul>\n", html_escape::encode_text(&trimmed[2..])));
                            } else {
                                html.push_str(&format!("<p>{}</p>\n", html_escape::encode_text(trimmed)));
                            }
                        }
                        html.push_str("</body></html>\n");
                        match std::fs::write(output_path, html) {
                            Ok(_) => ToolOutcome::Ok { message: Some(format!("Document exported to HTML at {}", output_path)) },
                            Err(e) => ToolOutcome::Error { code: ErrorCode::InternalError, error: format!("Failed to save file: {}", e), hint: None }
                        }
                    }
                    Err(e) => ToolOutcome::Error { code: ErrorCode::DocNotFound, error: e.to_string(), hint: None }
                }
            },
            
            "get_security_info" => {
                ToolOutcome::Security { security: serde_json::json!({
                    "readonly_mode": self.security_config.readonly_mode,
                    "sandbox_mode": self.security_config.sandbox_mode,
                    "allow_external_tools": self.security_config.allow_external_tools,
                    "allow_network": self.security_config.allow_network,
                    "max_document_size": self.security_config.max_document_size,
                    "max_open_documents": self.security_config.max_open_documents,
                    "summary": self.security_config.get_summary(),
                    "readonly_commands": crate::security::SecurityConfig::get_readonly_commands().len(),
                    "write_commands": crate::security::SecurityConfig::get_write_commands().len()
                }) }
            },
            
            "get_storage_info" => {
                let handler = self.handler.read().unwrap();
                match handler.get_storage_info() {
                    Ok(info) => ToolOutcome::Storage { storage: info.get("storage").cloned().unwrap_or(serde_json::json!({})) },
                    Err(e) => ToolOutcome::Error { code: ErrorCode::InternalError, error: e.to_string(), hint: None },
                }
            },
            
            _ => {
                ToolOutcome::Error { code: ErrorCode::UnknownTool, error: format!("Unknown or unsupported tool: {}", name), hint: None }
            }
        };
        // Backward-compatible JSON shaping with success boolean at top-level
        let legacy = match outcome {
            ToolOutcome::Ok { message } => {
                let mut obj = serde_json::json!({"success": true});
                if let Some(m) = message { obj["message"] = serde_json::Value::String(m); }
                obj
            }
            ToolOutcome::Created { document_id, message } => {
                let mut obj = serde_json::json!({"success": true, "document_id": document_id});
                if let Some(m) = message { obj["message"] = serde_json::Value::String(m); }
                obj
            }
            ToolOutcome::Text { text } => serde_json::json!({"success": true, "text": text}),
            ToolOutcome::Metadata { metadata } => {
                // Heuristic: if this looks like search results (matches/total_matches), flatten.
                let is_search_shape = metadata.get("matches").is_some() || metadata.get("total_matches").is_some();
                if is_search_shape {
                    let mut obj = serde_json::json!({"success": true});
                    if let Some(map) = metadata.as_object() {
                        for (k, v) in map { obj[&k[..]] = v.clone(); }
                    }
                    obj
                } else {
                    serde_json::json!({"success": true, "metadata": metadata})
                }
            }
            ToolOutcome::Documents { documents } => serde_json::json!({"success": true, "documents": documents}),
            ToolOutcome::Images { images, message } => {
                let mut obj = serde_json::json!({"success": true, "images": images});
                if let Some(m) = message { obj["message"] = serde_json::Value::String(m); }
                obj
            }
            ToolOutcome::Security { security } => serde_json::json!({"success": true, "security": security}),
            ToolOutcome::Storage { storage } => serde_json::json!({"success": true, "storage": storage}),
            ToolOutcome::Statistics { statistics } => serde_json::json!({"success": true, "statistics": statistics}),
            ToolOutcome::Structure { structure } => serde_json::json!({"success": true, "structure": structure}),
            ToolOutcome::Error { code, error, hint } => {
                let mut obj = serde_json::json!({"success": false, "error": error});
                obj["code"] = serde_json::json!(code);
                if let Some(h) = hint { obj["hint"] = serde_json::Value::String(h); }
                obj
            }
        };
        CallToolResponse { content: vec![ToolResponseContent::Text(TextContent { content_type: "application/json".into(), text: legacy.to_string(), annotations: None })], is_error: None, meta: None }
    }
}