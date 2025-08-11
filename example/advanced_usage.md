# Advanced DOCX MCP Server Usage Examples

This document demonstrates the advanced capabilities of the DOCX MCP server with real-world examples.

## Professional Document Templates

### Creating a Business Report

```javascript
// Ask your AI: "Create a professional quarterly report with our sales data"

// 1. Create from report template
const doc = await mcp.call("create_from_template", {
  template: "Report"
});

// 2. Set document properties
await mcp.call("set_document_properties", {
  document_id: doc.document_id,
  properties: {
    title: "Q3 2024 Sales Report",
    subject: "Quarterly Business Review",
    author: "Sales Team",
    company: "TechCorp Inc",
    keywords: ["sales", "quarterly", "2024", "revenue"]
  }
});

// 3. Add custom sections with advanced formatting
await mcp.call("add_section", {
  document_id: doc.document_id,
  section_config: {
    page_size: "Letter",
    landscape: false,
    margins: {
      top: 25.4,
      bottom: 25.4, 
      left: 31.8,
      right: 25.4
    },
    columns: 1
  }
});

// 4. Add charts and data visualization
await mcp.call("add_chart", {
  document_id: doc.document_id,
  chart_type: "Column",
  data: {
    title: "Quarterly Revenue Growth",
    categories: ["Q1", "Q2", "Q3"],
    series: [{
      name: "Revenue ($M)",
      values: [1.2, 1.5, 1.8]
    }]
  }
});
```

### Advanced Mail Merge Campaign

```javascript
// Ask your AI: "Create personalized letters for our client list with custom fields"

// 1. Create template with merge fields
const template = await mcp.call("create_from_template", {
  template: "BusinessLetter"
});

await mcp.call("prepare_mail_merge_template", {
  document_id: template.document_id,
  fields: ["ClientName", "Company", "LastOrderDate", "AccountManager", "SpecialOffer"]
});

// 2. Process each recipient
const recipients = [
  {
    ClientName: "John Smith", 
    Company: "ABC Corp",
    LastOrderDate: "2024-02-15",
    AccountManager: "Sarah Johnson",
    SpecialOffer: "20% off next order"
  }
  // ... more recipients
];

for (const recipient of recipients) {
  // Create personalized document
  const personalDoc = await mcp.call("merge_template", {
    template_id: template.document_id,
    data: recipient
  });
  
  // Add watermark for draft review
  await mcp.call("add_watermark", {
    document_id: personalDoc.document_id,
    text: "CONFIDENTIAL",
    style: "Diagonal"
  });
}
```

## Document Analysis & Review

### Comprehensive Document Analysis

```javascript
// Ask your AI: "Analyze this contract for structure, formatting, and key terms"

const doc = await mcp.call("open_document", {
  path: "./contracts/service_agreement.docx"
});

// 1. Get document structure
const structure = await mcp.call("get_document_structure", {
  document_id: doc.document_id
});

// 2. Analyze formatting consistency
const formatting = await mcp.call("analyze_formatting", {
  document_id: doc.document_id
});

// 3. Get detailed statistics
const stats = await mcp.call("get_word_count", {
  document_id: doc.document_id
});

// 4. Search for key legal terms
const terms = ["liability", "indemnification", "termination", "confidential"];
for (const term of terms) {
  const results = await mcp.call("search_text", {
    document_id: doc.document_id,
    search_term: term,
    case_sensitive: false,
    whole_word: true
  });
  
  console.log(`Found "${term}" ${results.total_matches} times`);
}

// 5. Export analysis to Markdown
await mcp.call("export_to_markdown", {
  document_id: doc.document_id,
  output_path: "./analysis/contract_analysis.md"
});
```

### Collaborative Review Process

```javascript
// Ask your AI: "Set up this document for review with comments and track changes"

// 1. Enable track changes
await mcp.call("enable_track_changes", {
  document_id: doc.document_id,
  author: "Legal Review Team"
});

// 2. Add review comments
await mcp.call("add_comment", {
  document_id: doc.document_id,
  text: "Payment terms in section 3.2",
  comment: "Consider reducing payment terms from 60 to 30 days",
  author: "Finance Team"
});

// 3. Add footnotes for clarification
await mcp.call("add_footnote", {
  document_id: doc.document_id,
  reference_text: "governing law",
  footnote_text: "This clause should specify the state jurisdiction for legal disputes"
});

// 4. Create bookmarks for easy navigation
await mcp.call("add_bookmark", {
  document_id: doc.document_id,
  bookmark_name: "payment_terms",
  text: "3.2 Payment Terms"
});

// 5. Add cross-references
await mcp.call("add_cross_reference", {
  document_id: doc.document_id,
  bookmark_name: "payment_terms",
  display_text: "See Payment Terms section"
});
```

## Security & Compliance Examples

### Readonly Document Review

```bash
# Start server in readonly mode for document review only
export DOCX_MCP_READONLY=true
./target/release/docx-mcp
```

```javascript
// In readonly mode, these operations are available:
const doc = await mcp.call("open_document", {
  path: "./confidential/annual_report.docx"
});

// ✅ Allowed: Extract and analyze content
const text = await mcp.call("extract_text", {
  document_id: doc.document_id
});

const structure = await mcp.call("get_document_structure", {
  document_id: doc.document_id
});

// ✅ Allowed: Export for analysis
await mcp.call("export_to_markdown", {
  document_id: doc.document_id,
  output_path: "./analysis/report_content.md"
});

// ❌ Blocked: Any modification attempts
// These would return security errors:
// - add_paragraph
// - save_document  
// - find_and_replace
```

### Sandboxed Environment

```bash
# Run in sandbox mode - restricts file operations to temp directory
export DOCX_MCP_SANDBOX=true
export DOCX_MCP_NO_EXTERNAL_TOOLS=true
./target/release/docx-mcp
```

```javascript
// All file operations restricted to temporary directory
// Perfect for untrusted document processing

const doc = await mcp.call("create_document", {});

// ✅ Allowed: Operations in temp directory
await mcp.call("save_document", {
  document_id: doc.document_id,
  output_path: "/tmp/docx-mcp/safe_output.docx"
});

// ❌ Blocked: Operations outside temp directory
// This would return a security error:
await mcp.call("save_document", {
  document_id: doc.document_id,
  output_path: "/home/user/documents/output.docx" // BLOCKED
});
```

## Advanced Automation Workflows

### Automated Report Generation Pipeline

```javascript
// Ask your AI: "Create an automated monthly report generation system"

class ReportGenerator {
  async generateMonthlyReport(month, year, data) {
    // 1. Create from template
    const doc = await mcp.call("create_from_template", {
      template: "Report"
    });
    
    // 2. Set up custom styles
    await mcp.call("add_custom_style", {
      document_id: doc.document_id,
      style: {
        id: "CompanyHeading",
        name: "Company Heading",
        font: "Arial",
        size: 18,
        bold: true,
        color: "#2E86C1",
        spacing: {
          before: 12,
          after: 6,
          line: 1.15
        }
      }
    });
    
    // 3. Add dynamic content with bookmarks
    await mcp.call("add_bookmark", {
      document_id: doc.document_id,
      bookmark_name: "executive_summary",
      text: "Executive Summary"
    });
    
    // 4. Insert data charts
    for (const metric of data.metrics) {
      await mcp.call("add_chart", {
        document_id: doc.document_id,
        chart_type: metric.type,
        data: {
          title: metric.title,
          categories: metric.categories,
          series: metric.series
        }
      });
    }
    
    // 5. Add table of contents
    await mcp.call("add_table_of_contents", {
      document_id: doc.document_id
    });
    
    // 6. Apply watermark
    await mcp.call("add_watermark", {
      document_id: doc.document_id,
      text: "INTERNAL USE ONLY",
      style: "Horizontal"
    });
    
    // 7. Generate multiple formats
    const filename = `monthly_report_${year}_${month}`;
    
    // Save DOCX
    await mcp.call("save_document", {
      document_id: doc.document_id,
      output_path: `./reports/${filename}.docx`
    });
    
    // Convert to PDF
    await mcp.call("convert_to_pdf", {
      document_id: doc.document_id,
      output_path: `./reports/${filename}.pdf`
    });
    
    // Generate preview images
    await mcp.call("convert_to_images", {
      document_id: doc.document_id,
      output_dir: `./reports/previews/`,
      format: "png",
      dpi: 150
    });
    
    return {
      docx: `./reports/${filename}.docx`,
      pdf: `./reports/${filename}.pdf`,
      preview: `./reports/previews/`
    };
  }
}
```

### Document Quality Assurance

```javascript
// Ask your AI: "Create a document QA system that checks formatting and compliance"

class DocumentQA {
  async auditDocument(documentPath) {
    const doc = await mcp.call("open_document", {
      path: documentPath
    });
    
    const audit = {
      document: documentPath,
      timestamp: new Date().toISOString(),
      issues: [],
      recommendations: []
    };
    
    // 1. Check document structure
    const structure = await mcp.call("get_document_structure", {
      document_id: doc.document_id
    });
    
    if (structure.structure.filter(s => s.type === "heading").length < 2) {
      audit.issues.push("Document lacks proper heading structure");
    }
    
    // 2. Analyze formatting consistency
    const formatting = await mcp.call("analyze_formatting", {
      document_id: doc.document_id
    });
    
    if (formatting.formatting_analysis.fonts_detected.length > 3) {
      audit.issues.push("Too many fonts used - limit to 2-3 for consistency");
    }
    
    // 3. Check for required content
    const requiredTerms = ["confidential", "copyright", "contact"];
    for (const term of requiredTerms) {
      const search = await mcp.call("search_text", {
        document_id: doc.document_id,
        search_term: term,
        case_sensitive: false
      });
      
      if (search.total_matches === 0) {
        audit.recommendations.push(`Consider adding ${term} information`);
      }
    }
    
    // 4. Check document statistics
    const stats = await mcp.call("get_word_count", {
      document_id: doc.document_id
    });
    
    if (stats.statistics.words < 500) {
      audit.issues.push("Document may be too short for professional standards");
    }
    
    // 5. Generate audit report
    const auditDoc = await mcp.call("create_document", {});
    
    await mcp.call("add_heading", {
      document_id: auditDoc.document_id,
      text: "Document Quality Audit Report",
      level: 1
    });
    
    await mcp.call("add_paragraph", {
      document_id: auditDoc.document_id,
      text: `Audit completed for: ${documentPath}`
    });
    
    // Add issues table
    const issuesData = audit.issues.map(issue => ["Issue", issue]);
    await mcp.call("add_table", {
      document_id: auditDoc.document_id,
      rows: [["Type", "Description"], ...issuesData]
    });
    
    await mcp.call("save_document", {
      document_id: auditDoc.document_id,
      output_path: `./qa/audit_${Date.now()}.docx`
    });
    
    return audit;
  }
}
```

## Security Configuration Examples

### Enterprise Security Setup

```bash
#!/bin/bash
# Enterprise security configuration script

# Readonly mode for document review workstations
export DOCX_MCP_READONLY=true

# Whitelist only analysis and export commands
export DOCX_MCP_WHITELIST="open_document,extract_text,get_metadata,get_document_structure,analyze_formatting,get_word_count,search_text,export_to_markdown,export_to_html,list_documents,get_security_info"

# Sandbox mode for processing untrusted documents
export DOCX_MCP_SANDBOX=true

# Resource limits
export DOCX_MCP_MAX_SIZE=10485760  # 10MB max file size
export DOCX_MCP_MAX_DOCS=5         # Max 5 open documents

# Disable external tools and network
export DOCX_MCP_NO_EXTERNAL_TOOLS=true
export DOCX_MCP_NO_NETWORK=true

echo "🔒 Starting DOCX MCP Server in Enterprise Security Mode"
./target/release/docx-mcp
```

### Development Environment Setup

```bash
#!/bin/bash
# Development environment with full features

# Allow all operations but with reasonable limits
export DOCX_MCP_MAX_SIZE=104857600  # 100MB max file size
export DOCX_MCP_MAX_DOCS=25         # Max 25 open documents

# Enable all features
unset DOCX_MCP_READONLY
unset DOCX_MCP_SANDBOX
unset DOCX_MCP_WHITELIST
unset DOCX_MCP_BLACKLIST

echo "🚀 Starting DOCX MCP Server in Development Mode"
./target/release/docx-mcp
```

These examples demonstrate the full power and flexibility of the DOCX MCP server for professional document workflows, from simple document creation to complex enterprise automation systems.