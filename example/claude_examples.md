# Claude Desktop Examples

These are real examples you can use with Claude Desktop once the DOCX MCP server is configured.

## Basic Document Creation

```
You: Create a new Word document with a professional letterhead for "Acme Corp" and save it as letterhead.docx
```

Claude will create a document with:
- Company name as heading
- Address and contact details
- Professional formatting
- Save to the specified file

## Invoice Generation

```
You: Generate an invoice for client "TechStart Inc" with these items:
- 10 hours consulting at $150/hour
- 1 software license at $500
- Add 10% tax
Save as invoice_2024_001.docx and convert to PDF
```

## Batch Processing

```
You: I have 5 DOCX files in the ./reports folder. Please:
1. Add page numbers to each
2. Set the header to "Confidential - Internal Use Only"
3. Convert all to PDF
4. Create a summary document listing all reports
```

## Data-Driven Documents

```
You: Create a sales report from this data:
Q1: $1.2M (15% growth)
Q2: $1.5M (25% growth)
Q3: $1.3M (8% growth)
Q4: $1.8M (38% growth)

Include:
- Executive summary
- Quarterly breakdown table
- Year-over-year comparison
- Recommendations section
Convert to PDF when done
```

## Template Operations

```
You: Open template.docx and replace these placeholders:
- {{CLIENT_NAME}} with "John Smith"
- {{DATE}} with today's date
- {{PROJECT}} with "Website Redesign"
- {{AMOUNT}} with "$5,000"
Save as contract_john_smith.docx
```

## Document Merging

```
You: Merge these documents in order:
1. cover_page.docx
2. executive_summary.docx
3. main_report.docx
4. appendix.docx

Add page numbers and a table of contents, then save as final_report.docx
```

## Text Extraction and Analysis

```
You: Extract all text from the documents in ./legal folder and:
1. Find all mentions of "liability"
2. Create a summary document with each mention and its context
3. Add a table showing which document contains which terms
```

## Report Formatting

```
You: Format this markdown content as a professional Word document:

# Project Status Report
## Overview
Project is on track...
## Milestones
- [x] Phase 1 complete
- [ ] Phase 2 in progress
## Budget
Current spend: $45,000 of $100,000

Add proper styling, convert checkboxes to a status table, and export as PDF.
```

## Document Comparison

```
You: Open contract_v1.docx and contract_v2.docx, then:
1. Extract text from both
2. Create a new document highlighting the differences
3. Add a summary table of all changes
4. Save as contract_comparison.docx
```

## Automated Documentation

```
You: Create API documentation from this OpenAPI spec file (api.yaml):
1. Generate a Word document with proper formatting
2. Include endpoint descriptions in a table
3. Add request/response examples
4. Create a PDF version for distribution
```

## Meeting Minutes Template

```
You: Create a meeting minutes template with:
- Company header
- Date, time, attendees fields
- Agenda items section
- Action items table with owner and due date columns
- Next meeting section
Save as meeting_template.docx
```

## Bulk Conversion

```
You: Convert all Word documents in my Downloads folder to:
1. PDF files in ./pdfs folder
2. PNG images (first page only) in ./thumbnails folder
3. Create an index.docx with links to all documents
```

## Complex Formatting

```
You: Create a technical specification document with:
1. Title page with document version and date
2. Table of contents (auto-generated)
3. Multiple heading levels
4. Code blocks with syntax highlighting effect
5. Diagrams placeholder sections
6. Numbered requirements list
7. Glossary table at the end
8. Footer with page numbers
```

## Mail Merge Simulation

```
You: I have a CSV with client data (clients.csv). For each client:
1. Create a personalized letter using template.docx
2. Replace all placeholders with client data
3. Save as PDF with client name in filename
4. Create a summary document listing all generated files
```