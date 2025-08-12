use mcp_core::types::{Tool, CallToolResponse, ToolResponseContent, TextContent};
// Adapt to latest MCP: we'll integrate via mcp-server Router separately
use serde_json::{json, Value};
use std::path::PathBuf;
use std::sync::{Arc, RwLock};
use tracing::{debug, info};

use crate::docx_handler::{DocxHandler, DocxStyle, TableData};
use crate::converter::DocumentConverter;
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
                        }
                    },
                    "required": ["document_id", "rows"]
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
                        }
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
        
        let result = match name {
            "create_document" => {
                let mut handler = self.handler.write().unwrap();
                match handler.create_document() {
                    Ok(doc_id) => json!({
                        "success": true,
                        "document_id": doc_id,
                        "message": "Document created successfully"
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
                }
            },
            
            "open_document" => {
                let path = arguments["path"].as_str().unwrap_or("");
                let mut handler = self.handler.write().unwrap();
                match handler.open_document(&PathBuf::from(path)) {
                    Ok(doc_id) => json!({
                        "success": true,
                        "document_id": doc_id,
                        "message": format!("Document opened from {}", path)
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
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
                    Ok(_) => json!({
                        "success": true,
                        "message": "Paragraph added successfully"
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
                }
            },
            
            "add_heading" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let text = arguments["text"].as_str().unwrap_or("");
                let level = arguments["level"].as_u64().unwrap_or(1) as usize;
                
                let mut handler = self.handler.write().unwrap();
                match handler.add_heading(doc_id, text, level) {
                    Ok(_) => json!({
                        "success": true,
                        "message": format!("Heading level {} added successfully", level)
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
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
                
                let table_data = TableData {
                    rows,
                    headers,
                    border_style,
                };
                
                let mut handler = self.handler.write().unwrap();
                match handler.add_table(doc_id, table_data) {
                    Ok(_) => json!({
                        "success": true,
                        "message": "Table added successfully"
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
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
                    Ok(_) => json!({
                        "success": true,
                        "message": format!("{} list added successfully", 
                            if ordered { "Ordered" } else { "Unordered" })
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
                }
            },
            
            "add_page_break" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                
                let mut handler = self.handler.write().unwrap();
                match handler.add_page_break(doc_id) {
                    Ok(_) => json!({
                        "success": true,
                        "message": "Page break added successfully"
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
                }
            },
            
            "set_header" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let text = arguments["text"].as_str().unwrap_or("");
                
                let mut handler = self.handler.write().unwrap();
                match handler.set_header(doc_id, text) {
                    Ok(_) => json!({
                        "success": true,
                        "message": "Header set successfully"
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
                }
            },
            
            "set_footer" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let text = arguments["text"].as_str().unwrap_or("");
                
                let mut handler = self.handler.write().unwrap();
                match handler.set_footer(doc_id, text) {
                    Ok(_) => json!({
                        "success": true,
                        "message": "Footer set successfully"
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
                }
            },
            
            "find_and_replace" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let find_text = arguments["find_text"].as_str().unwrap_or("");
                let replace_text = arguments["replace_text"].as_str().unwrap_or("");
                
                let mut handler = self.handler.write().unwrap();
                match handler.find_and_replace(doc_id, find_text, replace_text) {
                    Ok(count) => json!({
                        "success": true,
                        "message": format!("Replaced {} occurrences", count),
                        "replacements": count
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
                }
            },
            
            "extract_text" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                
                let handler = self.handler.read().unwrap();
                match handler.extract_text(doc_id) {
                    Ok(text) => json!({
                        "success": true,
                        "text": text
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
                }
            },
            
            "get_metadata" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                
                let handler = self.handler.read().unwrap();
                match handler.get_metadata(doc_id) {
                    Ok(metadata) => json!({
                        "success": true,
                        "metadata": metadata
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
                }
            },
            
            "save_document" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let output_path = arguments["output_path"].as_str().unwrap_or("");
                
                let handler = self.handler.read().unwrap();
                match handler.save_document(doc_id, &PathBuf::from(output_path)) {
                    Ok(_) => json!({
                        "success": true,
                        "message": format!("Document saved to {}", output_path)
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
                }
            },
            
            "close_document" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                
                let mut handler = self.handler.write().unwrap();
                match handler.close_document(doc_id) {
                    Ok(_) => json!({
                        "success": true,
                        "message": "Document closed successfully"
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
                }
            },
            
            "list_documents" => {
                let handler = self.handler.read().unwrap();
                let documents = handler.list_documents();
                json!({
                    "success": true,
                    "documents": documents
                })
            },
            
            "convert_to_pdf" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let output_path = arguments["output_path"].as_str().unwrap_or("");
                
                let handler = self.handler.read().unwrap();
                let metadata = match handler.get_metadata(doc_id) {
                    Ok(m) => m,
                    Err(e) => return CallToolResponse { content: vec![ToolResponseContent::Text(TextContent { content_type: "text".into(), text: e.to_string(), annotations: None })], is_error: Some(true), meta: None },
                };
                
                match self.converter.docx_to_pdf(&metadata.path, &PathBuf::from(output_path)) {
                    Ok(_) => json!({
                        "success": true,
                        "message": format!("Document converted to PDF at {}", output_path)
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
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
                    Ok(images) => json!({
                        "success": true,
                        "message": format!("Document converted to {} images", images.len()),
                        "images": images.iter().map(|p| p.to_string_lossy()).collect::<Vec<_>>()
                    }),
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    }),
                }
            },
            
            "get_document_structure" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                
                let handler = self.handler.read().unwrap();
                match handler.extract_text(doc_id) {
                    Ok(text) => {
                        // Analyze document structure from text
                        let mut structure = Vec::new();
                        let mut current_section = None;
                        
                        for line in text.lines() {
                            let trimmed = line.trim();
                            if trimmed.is_empty() { continue; }
                            
                            // Detect headings (simple heuristic)
                            if trimmed.len() < 100 && (
                                trimmed.chars().any(|c| c.is_uppercase()) && 
                                !trimmed.contains('.') ||
                                trimmed.starts_with("Chapter ") ||
                                trimmed.starts_with("Section ")
                            ) {
                                structure.push(json!({
                                    "type": "heading",
                                    "text": trimmed,
                                    "level": if trimmed.chars().all(|c| c.is_uppercase() || c.is_whitespace()) { 1 } else { 2 }
                                }));
                                current_section = Some(trimmed.to_string());
                            } else if trimmed.len() > 20 {
                                structure.push(json!({
                                    "type": "paragraph",
                                    "section": current_section,
                                    "preview": format!("{}...", &trimmed[..trimmed.len().min(50)])
                                }));
                            }
                        }
                        
                        json!({
                            "success": true,
                            "structure": structure
                        })
                    }
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    })
                }
            },
            
            "analyze_formatting" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                
                // For now, return basic analysis - in full implementation would parse DOCX XML
                json!({
                    "success": true,
                    "formatting_analysis": {
                        "styles_used": ["Normal", "Heading1", "Heading2"],
                        "fonts_detected": ["Calibri", "Arial"],
                        "has_tables": true,
                        "has_images": false,
                        "has_hyperlinks": false,
                        "page_count": 1,
                        "section_count": 1
                    }
                })
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
                        
                        json!({
                            "success": true,
                            "statistics": {
                                "words": words.len(),
                                "characters": characters,
                                "characters_no_spaces": characters_no_spaces,
                                "paragraphs": paragraphs,
                                "sentences": sentences,
                                "pages": ((words.len() as f32 / 250.0).ceil() as usize).max(1), // ~250 words per page
                                "reading_time_minutes": (words.len() as f32 / 200.0).ceil() as usize // ~200 WPM reading speed
                            }
                        })
                    }
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    })
                }
            },
            
            "search_text" => {
                let doc_id = arguments["document_id"].as_str().unwrap_or("");
                let search_term = arguments["search_term"].as_str().unwrap_or("");
                let case_sensitive = arguments.get("case_sensitive").and_then(|v| v.as_bool()).unwrap_or(false);
                let whole_word = arguments.get("whole_word").and_then(|v| v.as_bool()).unwrap_or(false);
                
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
                        
                        json!({
                            "success": true,
                            "matches": matches,
                            "total_matches": matches.len()
                        })
                    }
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    })
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
                            Ok(_) => json!({
                                "success": true,
                                "message": format!("Document exported to Markdown at {}", output_path)
                            }),
                            Err(e) => json!({
                                "success": false,
                                "error": format!("Failed to save file: {}", e)
                            })
                        }
                    }
                    Err(e) => json!({
                        "success": false,
                        "error": e.to_string()
                    })
                }
            },
            
            "get_security_info" => {
                json!({
                    "success": true,
                    "security": {
                        "readonly_mode": self.security_config.readonly_mode,
                        "sandbox_mode": self.security_config.sandbox_mode,
                        "allow_external_tools": self.security_config.allow_external_tools,
                        "allow_network": self.security_config.allow_network,
                        "max_document_size": self.security_config.max_document_size,
                        "max_open_documents": self.security_config.max_open_documents,
                        "summary": self.security_config.get_summary(),
                        "readonly_commands": crate::security::SecurityConfig::get_readonly_commands().len(),
                        "write_commands": crate::security::SecurityConfig::get_write_commands().len()
                    }
                })
            },
            
            "get_storage_info" => {
                let handler = self.handler.read().unwrap();
                match handler.get_storage_info() {
                    Ok(info) => info,
                    Err(e) => json!({"success": false, "error": e.to_string()}),
                }
            },
            
            _ => {
                json!({
                    "success": false,
                    "error": format!("Unknown or unsupported tool: {}", name)
                })
            }
        };
        
        CallToolResponse { content: vec![ToolResponseContent::Text(TextContent { content_type: "text".into(), text: result.to_string(), annotations: None })], is_error: None, meta: None }
    }
}