# DOCX MCP Server

A comprehensive Model Context Protocol (MCP) server for Microsoft Word DOCX file manipulation, built with Rust. This server provides AI systems with powerful tools to create, edit, convert, and manage Word documents programmatically.

## 📖 Table of Contents

- [Quick Start](#-quick-start)
- [AI Tool Integration](#-ai-tool-integration)
  - [Claude Desktop](#claude-desktop)
  - [Cursor](#cursor)
  - [Windsurf](#windsurf-codeium)
  - [Continue.dev](#continuedev)
  - [VS Code](#vs-code-with-mcp-extension)
- [Features](#-features)
- [Real-World Usage Examples](#-real-world-usage-examples-with-ai-assistants)
- [Prerequisites](#-prerequisites)
- [Installation](#-installation)
- [Common Use Cases](#-common-use-cases)
- [Available Tools](#available-tools)
- [Example Workflows](#example-workflows)
- [Architecture](#architecture)
- [Development](#development)
- [Troubleshooting](#-troubleshooting)
- [Examples Directory](#-examples-directory)
- [Contributing](#contributing)
- [License](#license)

## 🚀 Quick Start

```bash
# Clone the repository
git clone https://github.com/yourusername/docx-mcp.git
cd docx-mcp

# Download embedded fonts for standalone operation (optional but recommended)
./download_fonts.sh

# Build the server (creates a fully standalone binary)
./build.sh

# The server is now ready - no external dependencies required!
```

### 🎯 Standalone Operation

This MCP server is designed to work **completely standalone** without requiring LibreOffice, unoconv, or any external tools:

- ✅ **Pure Rust DOCX parsing** - No external libraries needed
- ✅ **Built-in PDF generation** - Creates PDFs without LibreOffice
- ✅ **Embedded fonts** - Professional typography included in the binary
- ✅ **Native image processing** - PNG/JPG generation without ImageMagick
- ✅ **Zero external dependencies** - Single binary deployment

The server will automatically use external tools if available for enhanced quality, but they are **completely optional**.

## 🔒 Security Features

The server includes comprehensive security features for enterprise and restricted environments:

### Readonly Mode
```bash
# Enable readonly mode - only allows document viewing and analysis
export DOCX_MCP_READONLY=true
./target/release/docx-mcp
```

In readonly mode, only these operations are allowed:
- Open and view documents
- Extract text and analyze structure
- Export to other formats (Markdown, PDF)
- Search and word count analysis
- Get document metadata and statistics

### Command Filtering
```bash
# Whitelist specific commands only
export DOCX_MCP_WHITELIST="open_document,extract_text,get_metadata,export_to_markdown"

# Or blacklist dangerous commands
export DOCX_MCP_BLACKLIST="save_document,convert_to_pdf,merge_documents"
```

### Sandbox Mode
```bash
# Restrict all file operations to temp directory only
export DOCX_MCP_SANDBOX=true
./target/release/docx-mcp
```

### Resource Limits
```bash
# Set maximum document size (100MB default)
export DOCX_MCP_MAX_SIZE=52428800  # 50MB

# Set maximum number of open documents
export DOCX_MCP_MAX_DOCS=20

# Disable external tools
export DOCX_MCP_NO_EXTERNAL_TOOLS=true

# Disable network operations
export DOCX_MCP_NO_NETWORK=true
```

## 🤖 AI Tool Integration

### Claude Desktop

Add to your Claude Desktop configuration file:

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`  
**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "docx": {
      "command": "/absolute/path/to/docx-mcp/target/release/docx-mcp",
      "args": [],
      "env": {
        "RUST_LOG": "info"
      }
    }
  }
}
```

After adding, restart Claude Desktop. You can then ask Claude to:
- "Create a new Word document with our Q4 report"
- "Convert this DOCX file to PDF"
- "Extract all text from my Word documents"
- "Add a table with sales data to the document"

### Cursor

Add to your Cursor settings (`~/.cursor/config.json` or through Settings UI):

```json
{
  "mcp": {
    "servers": {
      "docx": {
        "command": "/absolute/path/to/docx-mcp/target/release/docx-mcp",
        "args": [],
        "env": {
          "RUST_LOG": "info"
        }
      }
    }
  }
}
```

### Windsurf (Codeium)

Add to your Windsurf configuration (`~/.windsurf/config.json`):

```json
{
  "mcp": {
    "servers": {
      "docx": {
        "command": "/absolute/path/to/docx-mcp/target/release/docx-mcp",
        "args": [],
        "env": {
          "RUST_LOG": "info"
        }
      }
    }
  }
}
```

### Continue.dev

Add to your Continue configuration (`~/.continue/config.json`):

```json
{
  "models": [
    {
      "title": "Your Model",
      "provider": "your-provider",
      "mcp_servers": {
        "docx": {
          "command": "/absolute/path/to/docx-mcp/target/release/docx-mcp",
          "args": []
        }
      }
    }
  ]
}
```

### VS Code with MCP Extension

If using the MCP extension for VS Code, add to your workspace settings (`.vscode/settings.json`):

```json
{
  "mcp.servers": {
    "docx": {
      "command": "/absolute/path/to/docx-mcp/target/release/docx-mcp",
      "args": [],
      "env": {
        "RUST_LOG": "info"
      }
    }
  }
}
```

## 📚 Features

### Document Operations
- **Create & Open**: Create new documents or open existing DOCX files
- **Text Manipulation**: Add paragraphs, headings, lists with full styling support
- **Tables**: Create and format tables with custom layouts
- **Page Layout**: Add page breaks, set headers/footers
- **Find & Replace**: Search and replace text throughout documents
- **Text Extraction**: Extract plain text content from documents

### Conversion Capabilities
- **DOCX to PDF**: Convert Word documents to PDF format
  - Uses LibreOffice/unoconv for high-fidelity conversion
  - Fallback to basic PDF generation if external tools unavailable
- **DOCX to Images**: Convert document pages to PNG/JPG images
  - Configurable DPI for quality control
  - Support for multiple image formats
- **PDF Operations**: Split, merge, and manipulate PDF files

### Advanced Features
- **Document Metadata**: Track creation time, size, author, etc.
- **Styling Support**: Font family, size, bold, italic, underline, colors, alignment
- **Multiple Documents**: Handle multiple documents simultaneously
- **Temp File Management**: Automatic cleanup of temporary files

### Professional Templates
- **Business Letters**: Professional correspondence with proper formatting
- **Resumes**: Modern resume layouts with sections for experience, education, skills
- **Reports**: Technical and business reports with table of contents
- **Invoices**: Professional invoice templates with itemized billing
- **Contracts**: Legal document templates with signature blocks
- **Memos**: Corporate memorandum format
- **Newsletters**: Multi-column layouts for publications

### Advanced Document Features
- **Table of Contents**: Automatic TOC generation with heading links
- **Images & Charts**: Embed images and create data visualizations
- **Hyperlinks & Bookmarks**: Internal and external linking with navigation
- **Footnotes & Endnotes**: Academic and professional citation support
- **Comments & Track Changes**: Collaboration features for document review
- **Watermarks**: Confidential, draft, and custom watermarks
- **Mail Merge**: Automated personalized document generation
- **Custom Styles**: Create and apply consistent formatting themes

### Analysis & Review Tools
- **Document Structure Analysis**: Outline view of headings and sections
- **Formatting Analysis**: Detect fonts, styles, and formatting inconsistencies
- **Advanced Search**: Pattern matching with context and positioning
- **Word Count Statistics**: Detailed metrics including reading time
- **Export Options**: Convert to Markdown, HTML, and other formats

## 💬 Real-World Usage Examples with AI Assistants

### With Claude Desktop

Once configured, you can have natural conversations with Claude:

```
You: "Create a professional invoice template for my consulting business"

Claude will:
1. Create a new DOCX document
2. Add your company header
3. Insert a table for line items
4. Add payment terms and footer
5. Save it as invoice_template.docx
```

```
You: "Convert all the Word documents in my reports folder to PDF"

Claude will:
1. List all DOCX files
2. Open each document
3. Convert to PDF with the same name
4. Report completion status
```

### With Cursor/Windsurf

While coding, you can generate documentation:

```
You: "Generate API documentation from these TypeScript interfaces and save as Word"

The AI will:
1. Parse your code
2. Create a formatted DOCX with:
   - Title and table of contents
   - Endpoint descriptions
   - Request/response examples
   - Error codes table
3. Convert to PDF for distribution
```

### Automation Examples

```python
# Ask your AI: "Create a script to generate monthly reports"
# The AI can use the DOCX server to:

async def generate_monthly_report(month, year):
    # Create document
    doc = await mcp.call("create_document")
    
    # Add dynamic content
    await mcp.call("add_heading", {
        "document_id": doc.id,
        "text": f"Monthly Report - {month} {year}",
        "level": 1
    })
    
    # Add data from your database
    sales_data = fetch_sales_data(month, year)
    await mcp.call("add_table", {
        "document_id": doc.id,
        "rows": format_sales_table(sales_data)
    })
    
    # Convert to PDF and email
    await mcp.call("convert_to_pdf", {
        "document_id": doc.id,
        "output_path": f"reports/{year}_{month}_report.pdf"
    })
```

## 📋 Prerequisites

### Required
- Rust 1.70+ and Cargo (for building from source)
- MCP-compatible AI client (Claude Desktop, Cursor, Windsurf, etc.)

### Completely Optional (for enhanced features)

The server works standalone, but can optionally use these tools if available:
- **LibreOffice** (recommended): For high-quality DOCX to PDF conversion
  ```bash
  # macOS
  brew install libreoffice
  
  # Ubuntu/Debian
  sudo apt-get install libreoffice
  
  # Windows
  # Download from https://www.libreoffice.org/
  ```

- **PDF to Image Tools** (any one of these):
  - pdftoppm (part of poppler-utils)
  - ImageMagick
  - Ghostscript

  ```bash
  # macOS
  brew install poppler imagemagick ghostscript
  
  # Ubuntu/Debian
  sudo apt-get install poppler-utils imagemagick ghostscript
  ```

## 🔧 Installation

### Method 1: Build from Source

```bash
# Clone the repository
git clone https://github.com/yourusername/docx-mcp.git
cd docx-mcp

# Build the server (uses the build script)
./build.sh

# Or manually with cargo
cargo build --release

# Optional: Enable Chrome-based PDF conversion
cargo build --release --features chrome-pdf
```

### Method 2: Download Pre-built Binary (Coming Soon)

```bash
# Download the latest release
curl -L https://github.com/yourusername/docx-mcp/releases/latest/download/docx-mcp-linux-x64 -o docx-mcp
chmod +x docx-mcp
```

### Verify Installation

```bash
# Test the server
./target/release/docx-mcp --version

# Check for optional dependencies
./build.sh
```

## 🎯 Common Use Cases

### 1. Document Automation
- Generate contracts, invoices, and reports
- Mail merge operations
- Batch document processing
- Template-based document creation

### 2. Data Export
- Export database reports to Word/PDF
- Create formatted documentation from JSON/CSV
- Generate test reports with charts and tables

### 3. Document Conversion Pipeline
- DOCX → PDF for archival
- DOCX → Images for previews
- Batch conversion of legacy documents

### 4. Content Management
- Extract text for indexing
- Find and replace across multiple documents
- Document metadata management

### 5. Integration Scenarios
- CI/CD documentation generation
- API documentation from code
- Automated report generation from monitoring tools

## Available Tools

### Document Management

#### `create_document`
Creates a new empty DOCX document.
```json
{
  "tool": "create_document",
  "arguments": {}
}
```

#### `open_document`
Opens an existing DOCX file.
```json
{
  "tool": "open_document",
  "arguments": {
    "path": "/path/to/document.docx"
  }
}
```

#### `save_document`
Saves the document to a specified path.
```json
{
  "tool": "save_document",
  "arguments": {
    "document_id": "doc_123",
    "output_path": "/path/to/output.docx"
  }
}
```

### Content Addition

#### `add_paragraph`
Adds a styled paragraph to the document.
```json
{
  "tool": "add_paragraph",
  "arguments": {
    "document_id": "doc_123",
    "text": "This is a paragraph",
    "style": {
      "font_size": 12,
      "bold": true,
      "color": "#FF0000",
      "alignment": "center"
    }
  }
}
```

#### `add_heading`
Adds a heading (levels 1-6).
```json
{
  "tool": "add_heading",
  "arguments": {
    "document_id": "doc_123",
    "text": "Chapter 1",
    "level": 1
  }
}
```

#### `add_table`
Creates a table with specified data.
```json
{
  "tool": "add_table",
  "arguments": {
    "document_id": "doc_123",
    "rows": [
      ["Name", "Age", "City"],
      ["Alice", "30", "New York"],
      ["Bob", "25", "Los Angeles"]
    ],
    "headers": ["Name", "Age", "City"]
  }
}
```

#### `add_list`
Adds a bulleted or numbered list.
```json
{
  "tool": "add_list",
  "arguments": {
    "document_id": "doc_123",
    "items": ["First item", "Second item", "Third item"],
    "ordered": true
  }
}
```

### Document Conversion

#### `convert_to_pdf`
Converts the document to PDF format.
```json
{
  "tool": "convert_to_pdf",
  "arguments": {
    "document_id": "doc_123",
    "output_path": "/path/to/output.pdf"
  }
}
```

#### `convert_to_images`
Converts document pages to images.
```json
{
  "tool": "convert_to_images",
  "arguments": {
    "document_id": "doc_123",
    "output_dir": "/path/to/images/",
    "format": "png",
    "dpi": 300
  }
}
```

### Text Operations

#### `extract_text`
Extracts all text content from the document.
```json
{
  "tool": "extract_text",
  "arguments": {
    "document_id": "doc_123"
  }
}
```

#### `find_and_replace`
Finds and replaces text in the document.
```json
{
  "tool": "find_and_replace",
  "arguments": {
    "document_id": "doc_123",
    "find_text": "old text",
    "replace_text": "new text"
  }
}
```

## Example Workflows

### Creating a Report
```javascript
// 1. Create a new document
const doc = await mcp.call("create_document", {});

// 2. Add title
await mcp.call("add_heading", {
  document_id: doc.document_id,
  text: "Annual Report 2024",
  level: 1
});

// 3. Add executive summary
await mcp.call("add_paragraph", {
  document_id: doc.document_id,
  text: "This report provides a comprehensive overview...",
  style: { font_size: 12, alignment: "justify" }
});

// 4. Add data table
await mcp.call("add_table", {
  document_id: doc.document_id,
  rows: [
    ["Quarter", "Revenue", "Growth"],
    ["Q1", "$1.2M", "15%"],
    ["Q2", "$1.5M", "25%"]
  ]
});

// 5. Convert to PDF
await mcp.call("convert_to_pdf", {
  document_id: doc.document_id,
  output_path: "./annual_report_2024.pdf"
});
```

### Batch Processing Documents
```javascript
// Open and convert multiple documents
const documents = ["doc1.docx", "doc2.docx", "doc3.docx"];

for (const docPath of documents) {
  const doc = await mcp.call("open_document", { path: docPath });
  
  // Extract text for analysis
  const text = await mcp.call("extract_text", { 
    document_id: doc.document_id 
  });
  
  // Convert to PDF
  await mcp.call("convert_to_pdf", {
    document_id: doc.document_id,
    output_path: docPath.replace(".docx", ".pdf")
  });
  
  // Generate thumbnails
  await mcp.call("convert_to_images", {
    document_id: doc.document_id,
    output_dir: "./thumbnails/",
    format: "jpg",
    dpi: 72
  });
  
  await mcp.call("close_document", { document_id: doc.document_id });
}
```

## Architecture

The server is built with a modular architecture:

- **`main.rs`**: MCP server setup and initialization
- **`docx_handler.rs`**: Core DOCX manipulation logic
- **`converter.rs`**: PDF and image conversion functionality
- **`docx_tools.rs`**: MCP tool definitions and handlers

## Development

### Building from Source
```bash
cargo build
```

### Running Tests
```bash
cargo test
```

### Debug Mode
```bash
RUST_LOG=debug cargo run
```

## 🐛 Troubleshooting

### AI Tool Specific Issues

#### Claude Desktop Not Recognizing the Server
1. Ensure the path in config is absolute, not relative
2. Restart Claude Desktop after config changes
3. Check logs: `tail -f ~/Library/Logs/Claude/mcp.log` (macOS)
4. Verify the binary is executable: `chmod +x /path/to/docx-mcp`

#### Cursor/Windsurf Connection Issues
1. Check the MCP server is running: `ps aux | grep docx-mcp`
2. Verify port availability: `lsof -i :3000`
3. Try reloading the window: `Cmd/Ctrl + R`
4. Check developer console for errors: `Cmd/Ctrl + Shift + I`

#### "Tool not found" Errors
1. Ensure the server is properly configured in your AI tool
2. Check the server is running with: `RUST_LOG=debug /path/to/docx-mcp`
3. Verify tool names match exactly (case-sensitive)

### Conversion Issues

#### LibreOffice Not Found
```bash
# Check if installed
which libreoffice

# Install if missing
# macOS
brew install libreoffice

# Ubuntu/Debian
sudo apt-get install libreoffice

# Fedora
sudo dnf install libreoffice
```

#### PDF to Image Conversion Fails
```bash
# Install at least one converter
# Option 1: pdftoppm (fastest)
sudo apt-get install poppler-utils  # Linux
brew install poppler                 # macOS

# Option 2: ImageMagick
sudo apt-get install imagemagick     # Linux
brew install imagemagick              # macOS

# Option 3: Ghostscript
sudo apt-get install ghostscript     # Linux
brew install ghostscript              # macOS
```

### Permission Errors
```bash
# Check temp directory permissions
ls -la /tmp/docx-mcp/

# Fix permissions if needed
mkdir -p /tmp/docx-mcp
chmod 755 /tmp/docx-mcp

# For system-wide installation
sudo chown $USER:$USER /tmp/docx-mcp
```

### Memory Issues with Large Documents
```bash
# Increase Rust stack size if needed
export RUST_MIN_STACK=8388608  # 8MB
./target/release/docx-mcp
```

### Debugging Tips
```bash
# Run with verbose logging
RUST_LOG=trace ./target/release/docx-mcp

# Test with the example client
python3 example/test_client.py

# Check MCP communication
RUST_LOG=mcp_server=debug ./target/release/docx-mcp
```

## 📁 Examples Directory

The `example/` directory contains comprehensive examples and templates:

### Files Included

- **`test_client.py`** - Python client to test all MCP server functions
- **`claude_examples.md`** - Real-world examples for Claude Desktop users
- **`config_examples.json`** - Configuration templates for all supported AI tools
- **`automation_example.py`** - Advanced automation workflows including:
  - Monthly report generation
  - Mail merge operations
  - Document processing pipelines
  - Contract generation

### Running Examples

```bash
# Test the server functionality
python3 example/test_client.py

# Run automation examples
python3 example/automation_example.py

# View Claude Desktop usage examples
cat example/claude_examples.md
```

### Example Categories

1. **Basic Operations**: Create, edit, save documents
2. **Formatting**: Styles, tables, lists, headers/footers
3. **Conversion**: DOCX to PDF, DOCX to images
4. **Automation**: Batch processing, mail merge, report generation
5. **Integration**: Working with CSV data, template processing

## 🤝 Contributing

We welcome contributions! Here's how you can help:

### Areas for Contribution

- Additional document manipulation features
- Support for more conversion formats
- Performance optimizations
- Documentation improvements
- Bug fixes and testing

### How to Contribute

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Development Setup

```bash
# Clone your fork
git clone https://github.com/yourusername/docx-mcp.git
cd docx-mcp

# Install development dependencies
cargo install cargo-watch cargo-expand

# Run tests
cargo test

# Run with watch mode for development
cargo watch -x run
```

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with the official [MCP Rust SDK](https://github.com/modelcontextprotocol/rust-sdk)
- Uses [docx-rs](https://github.com/bokuweb/docx-rs) for DOCX manipulation
- PDF generation with [printpdf](https://github.com/fschutt/printpdf)
- Image processing with [image-rs](https://github.com/image-rs/image)