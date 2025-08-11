#!/usr/bin/env python3
"""
Example client to test the DOCX MCP Server.
This demonstrates how to interact with the server using JSON-RPC.
"""

import json
import sys
import asyncio
import websockets

async def call_tool(websocket, tool_name, arguments):
    """Call a tool on the MCP server"""
    request = {
        "jsonrpc": "2.0",
        "method": "tools/call",
        "params": {
            "name": tool_name,
            "arguments": arguments
        },
        "id": 1
    }
    
    await websocket.send(json.dumps(request))
    response = await websocket.recv()
    return json.loads(response)

async def main():
    # Connect to the MCP server (adjust the URI as needed)
    uri = "ws://localhost:3000"  # Default MCP server port
    
    async with websockets.connect(uri) as websocket:
        print("Connected to DOCX MCP Server")
        
        # Example 1: Create a new document
        print("\n1. Creating new document...")
        result = await call_tool(websocket, "create_document", {})
        doc_id = result["result"]["document_id"]
        print(f"   Document created with ID: {doc_id}")
        
        # Example 2: Add a heading
        print("\n2. Adding heading...")
        result = await call_tool(websocket, "add_heading", {
            "document_id": doc_id,
            "text": "Sample Document",
            "level": 1
        })
        print("   Heading added")
        
        # Example 3: Add a paragraph with styling
        print("\n3. Adding styled paragraph...")
        result = await call_tool(websocket, "add_paragraph", {
            "document_id": doc_id,
            "text": "This is a sample paragraph with custom styling.",
            "style": {
                "font_size": 14,
                "bold": True,
                "color": "#0066CC",
                "alignment": "center"
            }
        })
        print("   Styled paragraph added")
        
        # Example 4: Add a table
        print("\n4. Adding table...")
        result = await call_tool(websocket, "add_table", {
            "document_id": doc_id,
            "rows": [
                ["Product", "Price", "Quantity"],
                ["Widget A", "$10.99", "100"],
                ["Widget B", "$15.99", "75"],
                ["Widget C", "$8.99", "150"]
            ]
        })
        print("   Table added")
        
        # Example 5: Add a numbered list
        print("\n5. Adding numbered list...")
        result = await call_tool(websocket, "add_list", {
            "document_id": doc_id,
            "items": [
                "First item in the list",
                "Second item with more text",
                "Third and final item"
            ],
            "ordered": True
        })
        print("   Numbered list added")
        
        # Example 6: Set header and footer
        print("\n6. Setting header and footer...")
        result = await call_tool(websocket, "set_header", {
            "document_id": doc_id,
            "text": "Sample Document Header"
        })
        result = await call_tool(websocket, "set_footer", {
            "document_id": doc_id,
            "text": "Page 1 | Confidential"
        })
        print("   Header and footer set")
        
        # Example 7: Save the document
        print("\n7. Saving document...")
        result = await call_tool(websocket, "save_document", {
            "document_id": doc_id,
            "output_path": "./sample_output.docx"
        })
        print("   Document saved to sample_output.docx")
        
        # Example 8: Convert to PDF
        print("\n8. Converting to PDF...")
        result = await call_tool(websocket, "convert_to_pdf", {
            "document_id": doc_id,
            "output_path": "./sample_output.pdf"
        })
        if result["result"]["success"]:
            print("   Document converted to PDF")
        else:
            print(f"   PDF conversion failed: {result['result'].get('error', 'Unknown error')}")
        
        # Example 9: Convert to images
        print("\n9. Converting to images...")
        result = await call_tool(websocket, "convert_to_images", {
            "document_id": doc_id,
            "output_dir": "./images/",
            "format": "png",
            "dpi": 150
        })
        if result["result"]["success"]:
            print(f"   Document converted to images: {result['result']['images']}")
        else:
            print(f"   Image conversion failed: {result['result'].get('error', 'Unknown error')}")
        
        # Example 10: Extract text
        print("\n10. Extracting text...")
        result = await call_tool(websocket, "extract_text", {
            "document_id": doc_id
        })
        text = result["result"]["text"]
        print(f"   Extracted text (first 100 chars): {text[:100]}...")
        
        # Example 11: Get metadata
        print("\n11. Getting metadata...")
        result = await call_tool(websocket, "get_metadata", {
            "document_id": doc_id
        })
        metadata = result["result"]["metadata"]
        print(f"   Document metadata: {json.dumps(metadata, indent=2)}")
        
        # Example 12: Close the document
        print("\n12. Closing document...")
        result = await call_tool(websocket, "close_document", {
            "document_id": doc_id
        })
        print("   Document closed")
        
        print("\n✅ All tests completed successfully!")

if __name__ == "__main__":
    asyncio.run(main())