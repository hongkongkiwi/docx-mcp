#!/usr/bin/env python3
"""
Advanced automation example using the DOCX MCP Server.
This demonstrates how to build document automation workflows.
"""

import json
import asyncio
import csv
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any

# This would normally be your MCP client library
# For demonstration, we're showing the structure
class MCPClient:
    """Mock MCP Client for demonstration"""
    
    async def call(self, tool_name: str, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """Call an MCP tool"""
        # In reality, this would communicate with the MCP server
        print(f"Calling {tool_name} with {arguments}")
        return {"success": True, "result": {}}

# Initialize client
mcp = MCPClient()

# === Example 1: Generate Monthly Reports ===
async def generate_monthly_report(month: int, year: int, data: Dict[str, Any]):
    """Generate a comprehensive monthly report"""
    
    # Create new document
    doc = await mcp.call("create_document", {})
    doc_id = doc["result"]["document_id"]
    
    # Add title page
    await mcp.call("add_heading", {
        "document_id": doc_id,
        "text": f"{data['company_name']}",
        "level": 1
    })
    
    await mcp.call("add_heading", {
        "document_id": doc_id,
        "text": f"Monthly Report - {datetime(year, month, 1).strftime('%B %Y')}",
        "level": 2
    })
    
    await mcp.call("add_page_break", {"document_id": doc_id})
    
    # Executive Summary
    await mcp.call("add_heading", {
        "document_id": doc_id,
        "text": "Executive Summary",
        "level": 1
    })
    
    await mcp.call("add_paragraph", {
        "document_id": doc_id,
        "text": data["executive_summary"],
        "style": {
            "font_size": 12,
            "alignment": "justify"
        }
    })
    
    # Key Metrics Table
    await mcp.call("add_heading", {
        "document_id": doc_id,
        "text": "Key Performance Indicators",
        "level": 2
    })
    
    await mcp.call("add_table", {
        "document_id": doc_id,
        "rows": [
            ["Metric", "Target", "Actual", "Variance"],
            ["Revenue", f"${data['targets']['revenue']:,}", f"${data['actuals']['revenue']:,}", 
             f"{((data['actuals']['revenue'] / data['targets']['revenue'] - 1) * 100):.1f}%"],
            ["New Customers", str(data['targets']['customers']), str(data['actuals']['customers']),
             f"{data['actuals']['customers'] - data['targets']['customers']:+d}"],
            ["Satisfaction Score", f"{data['targets']['satisfaction']}%", f"{data['actuals']['satisfaction']}%",
             f"{data['actuals']['satisfaction'] - data['targets']['satisfaction']:+.1f}%"]
        ]
    })
    
    # Department Reports
    for dept in data['departments']:
        await mcp.call("add_heading", {
            "document_id": doc_id,
            "text": f"{dept['name']} Department",
            "level": 2
        })
        
        await mcp.call("add_paragraph", {
            "document_id": doc_id,
            "text": dept['summary']
        })
        
        if dept.get('achievements'):
            await mcp.call("add_list", {
                "document_id": doc_id,
                "items": dept['achievements'],
                "ordered": False
            })
    
    # Action Items
    await mcp.call("add_heading", {
        "document_id": doc_id,
        "text": "Action Items for Next Month",
        "level": 2
    })
    
    await mcp.call("add_table", {
        "document_id": doc_id,
        "rows": [
            ["Action", "Owner", "Due Date", "Priority"],
            *[[item['action'], item['owner'], item['due_date'], item['priority']] 
              for item in data['action_items']]
        ]
    })
    
    # Add footer
    await mcp.call("set_footer", {
        "document_id": doc_id,
        "text": f"Confidential - {data['company_name']} - Page"
    })
    
    # Save as DOCX
    filename = f"monthly_report_{year}_{month:02d}.docx"
    await mcp.call("save_document", {
        "document_id": doc_id,
        "output_path": f"./reports/{filename}"
    })
    
    # Convert to PDF
    await mcp.call("convert_to_pdf", {
        "document_id": doc_id,
        "output_path": f"./reports/{filename.replace('.docx', '.pdf')}"
    })
    
    # Generate thumbnail
    await mcp.call("convert_to_images", {
        "document_id": doc_id,
        "output_dir": "./reports/thumbnails/",
        "format": "png",
        "dpi": 72
    })
    
    await mcp.call("close_document", {"document_id": doc_id})
    
    return filename

# === Example 2: Mail Merge ===
async def mail_merge(template_path: str, csv_path: str, output_dir: str):
    """Perform mail merge with CSV data"""
    
    # Read CSV data
    with open(csv_path, 'r') as f:
        reader = csv.DictReader(f)
        recipients = list(reader)
    
    generated_files = []
    
    for recipient in recipients:
        # Open template
        template = await mcp.call("open_document", {"path": template_path})
        doc_id = template["result"]["document_id"]
        
        # Extract template text
        text_result = await mcp.call("extract_text", {"document_id": doc_id})
        text = text_result["result"]["text"]
        
        # Replace placeholders
        for field, value in recipient.items():
            placeholder = f"{{{{{field}}}}}"
            if placeholder in text:
                await mcp.call("find_and_replace", {
                    "document_id": doc_id,
                    "find_text": placeholder,
                    "replace_text": value
                })
        
        # Save personalized document
        output_filename = f"{recipient.get('name', 'document').replace(' ', '_')}.docx"
        output_path = f"{output_dir}/{output_filename}"
        
        await mcp.call("save_document", {
            "document_id": doc_id,
            "output_path": output_path
        })
        
        # Convert to PDF
        pdf_path = output_path.replace('.docx', '.pdf')
        await mcp.call("convert_to_pdf", {
            "document_id": doc_id,
            "output_path": pdf_path
        })
        
        generated_files.append({
            "recipient": recipient['name'],
            "docx": output_path,
            "pdf": pdf_path
        })
        
        await mcp.call("close_document", {"document_id": doc_id})
    
    # Create summary document
    summary = await mcp.call("create_document", {})
    summary_id = summary["result"]["document_id"]
    
    await mcp.call("add_heading", {
        "document_id": summary_id,
        "text": "Mail Merge Summary",
        "level": 1
    })
    
    await mcp.call("add_paragraph", {
        "document_id": summary_id,
        "text": f"Generated {len(generated_files)} documents on {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    })
    
    # Add summary table
    rows = [["Recipient", "DOCX File", "PDF File"]]
    for file_info in generated_files:
        rows.append([
            file_info['recipient'],
            file_info['docx'],
            file_info['pdf']
        ])
    
    await mcp.call("add_table", {
        "document_id": summary_id,
        "rows": rows
    })
    
    await mcp.call("save_document", {
        "document_id": summary_id,
        "output_path": f"{output_dir}/merge_summary.docx"
    })
    
    await mcp.call("close_document", {"document_id": summary_id})
    
    return generated_files

# === Example 3: Document Pipeline ===
async def document_processing_pipeline(input_dir: str):
    """Process multiple documents through a pipeline"""
    
    input_path = Path(input_dir)
    docx_files = list(input_path.glob("*.docx"))
    
    results = []
    
    for docx_file in docx_files:
        print(f"Processing {docx_file.name}...")
        
        # Open document
        doc = await mcp.call("open_document", {"path": str(docx_file)})
        doc_id = doc["result"]["document_id"]
        
        # Add watermark (header)
        await mcp.call("set_header", {
            "document_id": doc_id,
            "text": "DRAFT - CONFIDENTIAL"
        })
        
        # Add footer with date
        await mcp.call("set_footer", {
            "document_id": doc_id,
            "text": f"Processed on {datetime.now().strftime('%Y-%m-%d')}"
        })
        
        # Extract text for indexing
        text_result = await mcp.call("extract_text", {"document_id": doc_id})
        text = text_result["result"]["text"]
        word_count = len(text.split())
        
        # Save modified document
        output_docx = f"./processed/{docx_file.stem}_processed.docx"
        await mcp.call("save_document", {
            "document_id": doc_id,
            "output_path": output_docx
        })
        
        # Convert to PDF
        output_pdf = output_docx.replace('.docx', '.pdf')
        await mcp.call("convert_to_pdf", {
            "document_id": doc_id,
            "output_path": output_pdf
        })
        
        # Generate thumbnail
        await mcp.call("convert_to_images", {
            "document_id": doc_id,
            "output_dir": "./processed/thumbnails/",
            "format": "jpg",
            "dpi": 96
        })
        
        results.append({
            "original": docx_file.name,
            "word_count": word_count,
            "docx": output_docx,
            "pdf": output_pdf
        })
        
        await mcp.call("close_document", {"document_id": doc_id})
    
    # Create index document
    index = await mcp.call("create_document", {})
    index_id = index["result"]["document_id"]
    
    await mcp.call("add_heading", {
        "document_id": index_id,
        "text": "Document Processing Report",
        "level": 1
    })
    
    await mcp.call("add_paragraph", {
        "document_id": index_id,
        "text": f"Processed {len(results)} documents"
    })
    
    # Statistics table
    rows = [["Original File", "Word Count", "Output DOCX", "Output PDF"]]
    for result in results:
        rows.append([
            result['original'],
            str(result['word_count']),
            result['docx'],
            result['pdf']
        ])
    
    await mcp.call("add_table", {
        "document_id": index_id,
        "rows": rows
    })
    
    await mcp.call("save_document", {
        "document_id": index_id,
        "output_path": "./processed/index.docx"
    })
    
    await mcp.call("close_document", {"document_id": index_id})
    
    return results

# === Example 4: Contract Generator ===
async def generate_contract(contract_type: str, parties: Dict[str, Any], terms: Dict[str, Any]):
    """Generate a legal contract based on type and terms"""
    
    doc = await mcp.call("create_document", {})
    doc_id = doc["result"]["document_id"]
    
    # Title
    await mcp.call("add_heading", {
        "document_id": doc_id,
        "text": f"{contract_type.upper()} AGREEMENT",
        "level": 1
    })
    
    # Date and parties
    await mcp.call("add_paragraph", {
        "document_id": doc_id,
        "text": f"This Agreement is entered into as of {terms['date']} between:"
    })
    
    await mcp.call("add_list", {
        "document_id": doc_id,
        "items": [
            f"{parties['party1']['name']}, a {parties['party1']['type']} (\"Party 1\")",
            f"{parties['party2']['name']}, a {parties['party2']['type']} (\"Party 2\")"
        ],
        "ordered": False
    })
    
    # Terms sections
    section_num = 1
    for section_title, section_content in terms['sections'].items():
        await mcp.call("add_heading", {
            "document_id": doc_id,
            "text": f"{section_num}. {section_title}",
            "level": 2
        })
        
        if isinstance(section_content, list):
            await mcp.call("add_list", {
                "document_id": doc_id,
                "items": section_content,
                "ordered": True
            })
        else:
            await mcp.call("add_paragraph", {
                "document_id": doc_id,
                "text": section_content
            })
        
        section_num += 1
    
    # Signature block
    await mcp.call("add_page_break", {"document_id": doc_id})
    await mcp.call("add_heading", {
        "document_id": doc_id,
        "text": "SIGNATURES",
        "level": 2
    })
    
    signature_table = [
        ["Party 1:", "", "Party 2:", ""],
        ["", "", "", ""],
        ["_" * 30, "", "_" * 30, ""],
        ["Name:", parties['party1']['signatory'], "Name:", parties['party2']['signatory']],
        ["Title:", parties['party1']['title'], "Title:", parties['party2']['title']],
        ["Date:", "_" * 20, "Date:", "_" * 20]
    ]
    
    await mcp.call("add_table", {
        "document_id": doc_id,
        "rows": signature_table
    })
    
    # Save and convert
    filename = f"{contract_type.lower().replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}"
    await mcp.call("save_document", {
        "document_id": doc_id,
        "output_path": f"./contracts/{filename}.docx"
    })
    
    await mcp.call("convert_to_pdf", {
        "document_id": doc_id,
        "output_path": f"./contracts/{filename}.pdf"
    })
    
    await mcp.call("close_document", {"document_id": doc_id})
    
    return filename

# === Main execution ===
async def main():
    """Run example automations"""
    
    print("Document Automation Examples")
    print("=" * 40)
    
    # Example data for monthly report
    report_data = {
        "company_name": "TechCorp Industries",
        "executive_summary": "This month showed strong growth across all departments...",
        "targets": {"revenue": 1000000, "customers": 50, "satisfaction": 85},
        "actuals": {"revenue": 1150000, "customers": 62, "satisfaction": 88.5},
        "departments": [
            {
                "name": "Sales",
                "summary": "Sales exceeded targets by 15%",
                "achievements": ["Closed 3 enterprise deals", "Expanded into new market"]
            },
            {
                "name": "Engineering",
                "summary": "Delivered 2 major features on schedule",
                "achievements": ["Reduced bug count by 30%", "Improved performance by 25%"]
            }
        ],
        "action_items": [
            {"action": "Hire 2 senior developers", "owner": "HR", "due_date": "2024-02-15", "priority": "High"},
            {"action": "Launch marketing campaign", "owner": "Marketing", "due_date": "2024-02-01", "priority": "Medium"}
        ]
    }
    
    # Generate monthly report
    print("\n1. Generating monthly report...")
    report_file = await generate_monthly_report(1, 2024, report_data)
    print(f"   ✓ Generated: {report_file}")
    
    # Contract generation
    print("\n2. Generating service agreement...")
    contract_file = await generate_contract(
        "Service Agreement",
        {
            "party1": {"name": "ABC Corp", "type": "corporation", "signatory": "John Smith", "title": "CEO"},
            "party2": {"name": "XYZ Ltd", "type": "limited company", "signatory": "Jane Doe", "title": "Director"}
        },
        {
            "date": "January 15, 2024",
            "sections": {
                "Scope of Services": "Party 2 agrees to provide consulting services...",
                "Payment Terms": ["Monthly fee of $10,000", "Payment due within 30 days", "Late fee of 1.5% per month"],
                "Term and Termination": "This agreement shall commence on the date first written above...",
                "Confidentiality": "Both parties agree to maintain strict confidentiality..."
            }
        }
    )
    print(f"   ✓ Generated: {contract_file}")
    
    print("\n✅ All automation examples completed!")

if __name__ == "__main__":
    # Create necessary directories
    for dir_path in ["./reports", "./reports/thumbnails", "./contracts", "./processed", "./processed/thumbnails"]:
        Path(dir_path).mkdir(parents=True, exist_ok=True)
    
    # Run examples
    asyncio.run(main())