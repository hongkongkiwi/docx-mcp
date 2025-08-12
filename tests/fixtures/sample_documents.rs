//! Sample document templates and content for testing

use anyhow::Result;
use docx_mcp::docx_handler::{DocxHandler, DocxStyle, TableData};
use super::{TestStyles, TestTables, TestLists, TestContent};

/// Creates a business letter document for testing
pub fn create_business_letter(handler: &mut DocxHandler) -> Result<String> {
    let doc_id = handler.create_document()?;
    
    // Header
    handler.set_header(&doc_id, "ACME Corporation | 123 Business St, City, State 12345")?;
    
    // Date
    handler.add_paragraph(&doc_id, "December 15, 2024", Some(TestStyles::basic()))?;
    handler.add_paragraph(&doc_id, "", None)?; // Empty line
    
    // Recipient
    handler.add_paragraph(&doc_id, "Ms. Jane Smith", Some(TestStyles::basic()))?;
    handler.add_paragraph(&doc_id, "Director of Operations", Some(TestStyles::basic()))?;
    handler.add_paragraph(&doc_id, "XYZ Company", Some(TestStyles::basic()))?;
    handler.add_paragraph(&doc_id, "456 Corporate Ave", Some(TestStyles::basic()))?;
    handler.add_paragraph(&doc_id, "Business City, State 67890", Some(TestStyles::basic()))?;
    handler.add_paragraph(&doc_id, "", None)?; // Empty line
    
    // Subject
    handler.add_paragraph(&doc_id, "RE: Partnership Proposal", Some(TestStyles::emphasis()))?;
    handler.add_paragraph(&doc_id, "", None)?; // Empty line
    
    // Salutation
    handler.add_paragraph(&doc_id, "Dear Ms. Smith,", Some(TestStyles::basic()))?;
    handler.add_paragraph(&doc_id, "", None)?; // Empty line
    
    // Body paragraphs
    handler.add_paragraph(&doc_id, 
        "I am writing to propose a strategic partnership between ACME Corporation and XYZ Company that would benefit both organizations significantly. Our companies share similar values and complementary strengths that could create substantial value for our respective customers.",
        Some(TestStyles::basic()))?;
    
    handler.add_paragraph(&doc_id, 
        "ACME Corporation has been a leader in technology solutions for over 15 years, with a strong track record of innovation and customer satisfaction. We believe that combining our technical expertise with your operational excellence would create a powerful synergy in the marketplace.",
        Some(TestStyles::basic()))?;
    
    handler.add_paragraph(&doc_id, 
        "The proposed partnership would include joint product development, shared marketing initiatives, and coordinated customer support efforts. We estimate this collaboration could increase revenue for both companies by 25% within the first year.",
        Some(TestStyles::basic()))?;
    
    handler.add_paragraph(&doc_id, 
        "I would welcome the opportunity to discuss this proposal in more detail at your convenience. Please let me know when you might be available for a meeting or conference call.",
        Some(TestStyles::basic()))?;
    
    handler.add_paragraph(&doc_id, "", None)?; // Empty line
    
    // Closing
    handler.add_paragraph(&doc_id, "Sincerely,", Some(TestStyles::basic()))?;
    handler.add_paragraph(&doc_id, "", None)?; // Space for signature
    handler.add_paragraph(&doc_id, "", None)?; // Space for signature
    handler.add_paragraph(&doc_id, "John Doe", Some(TestStyles::basic()))?;
    handler.add_paragraph(&doc_id, "Chief Executive Officer", Some(TestStyles::basic()))?;
    handler.add_paragraph(&doc_id, "ACME Corporation", Some(TestStyles::basic()))?;
    
    // Footer
    handler.set_footer(&doc_id, "ACME Corporation - Confidential and Proprietary")?;
    
    Ok(doc_id)
}

/// Creates a technical report document for testing
pub fn create_technical_report(handler: &mut DocxHandler) -> Result<String> {
    let doc_id = handler.create_document()?;
    
    // Title page
    handler.add_paragraph(&doc_id, "", None)?; // Empty line for spacing
    handler.add_paragraph(&doc_id, "", None)?;
    handler.add_paragraph(&doc_id, "", None)?;
    
    handler.add_heading(&doc_id, "System Performance Analysis Report", 1)?;
    handler.add_paragraph(&doc_id, "", None)?;
    handler.add_paragraph(&doc_id, "Quarterly Assessment - Q4 2024", Some(TestStyles::centered()))?;
    handler.add_paragraph(&doc_id, "", None)?;
    handler.add_paragraph(&doc_id, "Prepared by: Technical Team", Some(TestStyles::centered()))?;
    handler.add_paragraph(&doc_id, "Date: December 15, 2024", Some(TestStyles::centered()))?;
    
    handler.add_page_break(&doc_id)?;
    
    // Executive Summary
    handler.add_heading(&doc_id, "Executive Summary", 1)?;
    handler.add_paragraph(&doc_id, 
        "This report provides a comprehensive analysis of system performance metrics for Q4 2024. Key findings include significant improvements in response times, enhanced security measures, and successful implementation of new monitoring capabilities.",
        Some(TestStyles::basic()))?;
    
    let summary_points = vec![
        "Average response time improved by 35%".to_string(),
        "System uptime achieved 99.97%".to_string(),
        "Security incidents reduced by 60%".to_string(),
        "User satisfaction increased to 94%".to_string(),
    ];
    handler.add_list(&doc_id, summary_points, false)?;
    
    // Performance Metrics
    handler.add_heading(&doc_id, "Performance Metrics", 1)?;
    
    handler.add_heading(&doc_id, "Response Time Analysis", 2)?;
    handler.add_paragraph(&doc_id, 
        "Response time measurements were collected continuously throughout Q4 2024. The data shows consistent improvement across all service endpoints.",
        Some(TestStyles::basic()))?;
    
    let response_time_data = TableData {
        rows: vec![
            vec!["Service".to_string(), "Q3 2024 (ms)".to_string(), "Q4 2024 (ms)".to_string(), "Improvement".to_string()],
            vec!["Authentication".to_string(), "245".to_string(), "158".to_string(), "35.5%".to_string()],
            vec!["Database Query".to_string(), "892".to_string(), "623".to_string(), "30.2%".to_string()],
            vec!["File Processing".to_string(), "1,240".to_string(), "789".to_string(), "36.4%".to_string()],
            vec!["Report Generation".to_string(), "3,450".to_string(), "2,180".to_string(), "36.8%".to_string()],
        ],
        headers: Some(vec!["Service".to_string(), "Q3 2024 (ms)".to_string(), "Q4 2024 (ms)".to_string(), "Improvement".to_string()]),
        border_style: Some("single".to_string()),
        col_widths: None,
        merges: None,
        cell_shading: None,
    };
    handler.add_table(&doc_id, response_time_data)?;
    
    handler.add_heading(&doc_id, "System Reliability", 2)?;
    handler.add_paragraph(&doc_id, 
        "System reliability metrics demonstrate exceptional stability and availability throughout the quarter.",
        Some(TestStyles::basic()))?;
    
    let reliability_data = TableData {
        rows: vec![
            vec!["Metric".to_string(), "Target".to_string(), "Actual".to_string(), "Status".to_string()],
            vec!["Uptime".to_string(), "99.9%".to_string(), "99.97%".to_string(), "✓ Exceeded".to_string()],
            vec!["MTBF (hours)".to_string(), "720".to_string(), "892".to_string(), "✓ Exceeded".to_string()],
            vec!["Recovery Time (min)".to_string(), "15".to_string(), "8.5".to_string(), "✓ Exceeded".to_string()],
        ],
        headers: Some(vec!["Metric".to_string(), "Target".to_string(), "Actual".to_string(), "Status".to_string()]),
        border_style: Some("single".to_string()),
        col_widths: None,
        merges: None,
        cell_shading: None,
    };
    handler.add_table(&doc_id, reliability_data)?;
    
    // Security Analysis
    handler.add_heading(&doc_id, "Security Analysis", 1)?;
    handler.add_paragraph(&doc_id, 
        "Security monitoring and incident response capabilities were significantly enhanced during Q4 2024.",
        Some(TestStyles::basic()))?;
    
    let security_improvements = vec![
        "Implemented advanced threat detection algorithms".to_string(),
        "Enhanced encryption protocols for data transmission".to_string(),
        "Deployed automated incident response systems".to_string(),
        "Conducted comprehensive security audits".to_string(),
        "Updated access control mechanisms".to_string(),
    ];
    handler.add_list(&doc_id, security_improvements, true)?;
    
    // Recommendations
    handler.add_heading(&doc_id, "Recommendations", 1)?;
    handler.add_paragraph(&doc_id, 
        "Based on the analysis conducted, the following recommendations are proposed for Q1 2025:",
        Some(TestStyles::basic()))?;
    
    let recommendations = vec![
        "Continue performance optimization initiatives".to_string(),
        "Expand monitoring coverage to include new services".to_string(),
        "Implement predictive analytics for proactive maintenance".to_string(),
        "Enhance disaster recovery procedures".to_string(),
        "Invest in additional security training for staff".to_string(),
    ];
    handler.add_list(&doc_id, recommendations, true)?;
    
    // Footer
    handler.set_footer(&doc_id, "Technical Report Q4 2024 - Confidential")?;
    
    Ok(doc_id)
}

/// Creates a meeting minutes document for testing
pub fn create_meeting_minutes(handler: &mut DocxHandler) -> Result<String> {
    let doc_id = handler.create_document()?;
    
    // Header
    handler.add_heading(&doc_id, "Project Steering Committee Meeting Minutes", 1)?;
    handler.add_paragraph(&doc_id, "", None)?;
    
    // Meeting details
    let meeting_details = TableData {
        rows: vec![
            vec!["Date:".to_string(), "December 15, 2024".to_string()],
            vec!["Time:".to_string(), "2:00 PM - 3:30 PM PST".to_string()],
            vec!["Location:".to_string(), "Conference Room A / Virtual".to_string()],
            vec!["Chair:".to_string(), "Sarah Johnson".to_string()],
            vec!["Secretary:".to_string(), "Mike Chen".to_string()],
        ],
        headers: None,
        border_style: Some("single".to_string()),
        col_widths: None,
        merges: None,
        cell_shading: None,
    };
    handler.add_table(&doc_id, meeting_details)?;
    
    // Attendees
    handler.add_heading(&doc_id, "Attendees", 2)?;
    let attendees = vec![
        "Sarah Johnson (Chair) - Project Director".to_string(),
        "Mike Chen (Secretary) - Technical Lead".to_string(),
        "Lisa Wang - Product Manager".to_string(),
        "David Rodriguez - Engineering Manager".to_string(),
        "Jennifer Kim - QA Manager".to_string(),
        "Alex Thompson - DevOps Lead".to_string(),
    ];
    handler.add_list(&doc_id, attendees, false)?;
    
    // Agenda Items
    handler.add_heading(&doc_id, "Agenda Items Discussed", 2)?;
    
    handler.add_heading(&doc_id, "1. Project Status Update", 3)?;
    handler.add_paragraph(&doc_id, 
        "Mike Chen presented the current project status, highlighting that development is 85% complete and on schedule for the January 31st deadline.",
        Some(TestStyles::basic()))?;
    
    let status_highlights = vec![
        "Core functionality implementation: 100% complete".to_string(),
        "User interface development: 90% complete".to_string(),
        "Testing and QA: 70% complete".to_string(),
        "Documentation: 60% complete".to_string(),
    ];
    handler.add_list(&doc_id, status_highlights, false)?;
    
    handler.add_heading(&doc_id, "2. Budget Review", 3)?;
    handler.add_paragraph(&doc_id, 
        "Lisa Wang reported that the project is currently 5% under budget with strong cost controls in place.",
        Some(TestStyles::basic()))?;
    
    let budget_data = TableData {
        rows: vec![
            vec!["Category".to_string(), "Budgeted".to_string(), "Actual".to_string(), "Remaining".to_string()],
            vec!["Development".to_string(), "$180,000".to_string(), "$168,000".to_string(), "$12,000".to_string()],
            vec!["Testing".to_string(), "$45,000".to_string(), "$38,000".to_string(), "$7,000".to_string()],
            vec!["Infrastructure".to_string(), "$30,000".to_string(), "$28,000".to_string(), "$2,000".to_string()],
            vec!["Total".to_string(), "$255,000".to_string(), "$234,000".to_string(), "$21,000".to_string()],
        ],
        headers: Some(vec!["Category".to_string(), "Budgeted".to_string(), "Actual".to_string(), "Remaining".to_string()]),
        border_style: Some("single".to_string()),
        col_widths: None,
        merges: None,
        cell_shading: None,
    };
    handler.add_table(&doc_id, budget_data)?;
    
    handler.add_heading(&doc_id, "3. Risk Assessment", 3)?;
    handler.add_paragraph(&doc_id, 
        "David Rodriguez presented the updated risk register with mitigation strategies for identified risks.",
        Some(TestStyles::basic()))?;
    
    let risks = vec![
        "Third-party API integration delays - Medium risk, mitigation plan in place".to_string(),
        "Resource availability during holidays - Low risk, backup resources identified".to_string(),
        "Performance requirements validation - Medium risk, load testing scheduled".to_string(),
    ];
    handler.add_list(&doc_id, risks, false)?;
    
    // Action Items
    handler.add_heading(&doc_id, "Action Items", 2)?;
    
    let action_items_data = TableData {
        rows: vec![
            vec!["Action Item".to_string(), "Owner".to_string(), "Due Date".to_string(), "Status".to_string()],
            vec!["Complete load testing scenarios".to_string(), "Jennifer Kim".to_string(), "Dec 22, 2024".to_string(), "In Progress".to_string()],
            vec!["Finalize API integration testing".to_string(), "Mike Chen".to_string(), "Dec 20, 2024".to_string(), "Not Started".to_string()],
            vec!["Update project documentation".to_string(), "Lisa Wang".to_string(), "Jan 10, 2025".to_string(), "Not Started".to_string()],
            vec!["Prepare deployment checklist".to_string(), "Alex Thompson".to_string(), "Jan 15, 2025".to_string(), "Not Started".to_string()],
        ],
        headers: Some(vec!["Action Item".to_string(), "Owner".to_string(), "Due Date".to_string(), "Status".to_string()]),
        border_style: Some("single".to_string()),
        col_widths: None,
        merges: None,
        cell_shading: None,
    };
    handler.add_table(&doc_id, action_items_data)?;
    
    // Next Meeting
    handler.add_heading(&doc_id, "Next Meeting", 2)?;
    handler.add_paragraph(&doc_id, 
        "The next steering committee meeting is scheduled for January 5, 2025, at 2:00 PM PST in Conference Room A.",
        Some(TestStyles::basic()))?;
    
    // Footer
    handler.set_footer(&doc_id, "Project Steering Committee - Meeting Minutes")?;
    
    Ok(doc_id)
}

/// Creates a product specification document for testing
pub fn create_product_spec(handler: &mut DocxHandler) -> Result<String> {
    let doc_id = handler.create_document()?;
    
    // Title page
    handler.add_paragraph(&doc_id, "", None)?;
    handler.add_paragraph(&doc_id, "", None)?;
    handler.add_heading(&doc_id, "Product Requirements Specification", 1)?;
    handler.add_paragraph(&doc_id, "", None)?;
    handler.add_paragraph(&doc_id, "Document Management System v2.0", Some(TestStyles::centered()))?;
    handler.add_paragraph(&doc_id, "", None)?;
    handler.add_paragraph(&doc_id, "Version 1.0", Some(TestStyles::centered()))?;
    handler.add_paragraph(&doc_id, "December 15, 2024", Some(TestStyles::centered()))?;
    
    handler.add_page_break(&doc_id)?;
    
    // Table of Contents (simplified)
    handler.add_heading(&doc_id, "Table of Contents", 1)?;
    let toc_items = vec![
        "1. Introduction".to_string(),
        "2. System Overview".to_string(),
        "3. Functional Requirements".to_string(),
        "4. Non-Functional Requirements".to_string(),
        "5. User Interface Requirements".to_string(),
        "6. System Architecture".to_string(),
        "7. Security Requirements".to_string(),
    ];
    handler.add_list(&doc_id, toc_items, true)?;
    
    // Introduction
    handler.add_heading(&doc_id, "1. Introduction", 1)?;
    
    handler.add_heading(&doc_id, "1.1 Purpose", 2)?;
    handler.add_paragraph(&doc_id, 
        "This document specifies the requirements for the Document Management System version 2.0. The system is designed to provide comprehensive document storage, retrieval, and collaboration capabilities for enterprise users.",
        Some(TestStyles::basic()))?;
    
    handler.add_heading(&doc_id, "1.2 Scope", 2)?;
    handler.add_paragraph(&doc_id, 
        "The Document Management System will support multiple file formats, version control, user collaboration, and advanced search capabilities. The system will be deployed as a web-based application with mobile support.",
        Some(TestStyles::basic()))?;
    
    // System Overview
    handler.add_heading(&doc_id, "2. System Overview", 1)?;
    handler.add_paragraph(&doc_id, 
        "The Document Management System consists of several integrated components working together to provide a seamless document management experience.",
        Some(TestStyles::basic()))?;
    
    let system_components = vec![
        "Document Storage Engine".to_string(),
        "Version Control System".to_string(),
        "Search and Indexing Service".to_string(),
        "User Authentication and Authorization".to_string(),
        "Collaboration Tools".to_string(),
        "Reporting and Analytics".to_string(),
    ];
    handler.add_list(&doc_id, system_components, false)?;
    
    // Functional Requirements
    handler.add_heading(&doc_id, "3. Functional Requirements", 1)?;
    
    handler.add_heading(&doc_id, "3.1 Document Upload and Storage", 2)?;
    let upload_requirements = vec![
        "FR-001: System shall support upload of files up to 100MB in size".to_string(),
        "FR-002: System shall support common file formats (PDF, DOCX, XLSX, PPTX, TXT)".to_string(),
        "FR-003: System shall automatically generate file metadata upon upload".to_string(),
        "FR-004: System shall provide drag-and-drop upload functionality".to_string(),
    ];
    handler.add_list(&doc_id, upload_requirements, false)?;
    
    handler.add_heading(&doc_id, "3.2 Search and Retrieval", 2)?;
    let search_requirements = vec![
        "FR-005: System shall provide full-text search capabilities".to_string(),
        "FR-006: System shall support advanced search with multiple criteria".to_string(),
        "FR-007: System shall provide search result ranking and relevance scoring".to_string(),
        "FR-008: System shall support search within specific document types".to_string(),
    ];
    handler.add_list(&doc_id, search_requirements, false)?;
    
    // Non-Functional Requirements
    handler.add_heading(&doc_id, "4. Non-Functional Requirements", 1)?;
    
    let nfr_data = TableData {
        rows: vec![
            vec!["Requirement".to_string(), "Specification".to_string(), "Priority".to_string()],
            vec!["Performance".to_string(), "Page load time < 3 seconds".to_string(), "High".to_string()],
            vec!["Scalability".to_string(), "Support 1000+ concurrent users".to_string(), "High".to_string()],
            vec!["Availability".to_string(), "99.9% uptime".to_string(), "High".to_string()],
            vec!["Security".to_string(), "Role-based access control".to_string(), "Critical".to_string()],
            vec!["Usability".to_string(), "Intuitive interface, minimal training".to_string(), "Medium".to_string()],
        ],
        headers: Some(vec!["Requirement".to_string(), "Specification".to_string(), "Priority".to_string()]),
        border_style: Some("single".to_string()),
        col_widths: None,
        merges: None,
        cell_shading: None,
    };
    handler.add_table(&doc_id, nfr_data)?;
    
    // Security Requirements
    handler.add_heading(&doc_id, "7. Security Requirements", 1)?;
    handler.add_paragraph(&doc_id, 
        "Security is paramount for the Document Management System. The following security measures must be implemented:",
        Some(TestStyles::basic()))?;
    
    let security_requirements = vec![
        "SEC-001: All data transmission must use HTTPS/TLS 1.3".to_string(),
        "SEC-002: User passwords must meet complexity requirements".to_string(),
        "SEC-003: System must support multi-factor authentication".to_string(),
        "SEC-004: All user actions must be logged for audit purposes".to_string(),
        "SEC-005: Document access must be controlled by user permissions".to_string(),
        "SEC-006: System must support data encryption at rest".to_string(),
    ];
    handler.add_list(&doc_id, security_requirements, true)?;
    
    // Footer
    handler.set_footer(&doc_id, "Product Requirements Specification v1.0 - Confidential")?;
    
    Ok(doc_id)
}

/// Creates a test document with international content
pub fn create_multilingual_document(handler: &mut DocxHandler) -> Result<String> {
    let doc_id = handler.create_document()?;
    
    handler.add_heading(&doc_id, "Multilingual Content Test Document", 1)?;
    handler.add_paragraph(&doc_id, 
        "This document contains text in multiple languages to test internationalization and Unicode support.",
        Some(TestStyles::basic()))?;
    
    for (language, text) in TestContent::multilingual_content() {
        handler.add_heading(&doc_id, language, 2)?;
        handler.add_paragraph(&doc_id, text, Some(TestStyles::basic()))?;
        handler.add_paragraph(&doc_id, "", None)?; // Empty line
    }
    
    handler.add_heading(&doc_id, "Special Characters and Symbols", 2)?;
    handler.add_paragraph(&doc_id, TestContent::special_characters(), Some(TestStyles::basic()))?;
    handler.add_paragraph(&doc_id, TestContent::symbols_and_math(), Some(TestStyles::basic()))?;
    
    // Currency symbols
    handler.add_paragraph(&doc_id, "Currency symbols: $ € £ ¥ ₹ ₽ ₩ ₪ ₫ ₱", Some(TestStyles::basic()))?;
    
    // Emoji (if supported)
    handler.add_paragraph(&doc_id, "Emoji test: 📄 📝 💼 🔒 🌍 ✅ ❌ ⚠️", Some(TestStyles::basic()))?;
    
    Ok(doc_id)
}

/// Creates a document with complex formatting for testing
pub fn create_formatted_document(handler: &mut DocxHandler) -> Result<String> {
    let doc_id = handler.create_document()?;
    
    handler.add_heading(&doc_id, "Formatting Test Document", 1)?;
    
    // Different paragraph styles
    handler.add_paragraph(&doc_id, "This paragraph uses the default style.", Some(TestStyles::basic()))?;
    handler.add_paragraph(&doc_id, "This paragraph uses bold formatting.", Some(DocxStyle {
        bold: Some(true),
        ..TestStyles::basic()
    }))?;
    handler.add_paragraph(&doc_id, "This paragraph uses italic formatting.", Some(DocxStyle {
        italic: Some(true),
        ..TestStyles::basic()
    }))?;
    handler.add_paragraph(&doc_id, "This paragraph is centered.", Some(TestStyles::centered()))?;
    handler.add_paragraph(&doc_id, "This paragraph uses emphasis styling.", Some(TestStyles::emphasis()))?;
    
    // Different font sizes
    handler.add_heading(&doc_id, "Font Size Tests", 2)?;
    for size in [8, 10, 12, 14, 16, 18, 24] {
        let style = DocxStyle {
            font_size: Some(size),
            ..TestStyles::basic()
        };
        handler.add_paragraph(&doc_id, &format!("This text is {} point size.", size), Some(style))?;
    }
    
    // Color tests
    handler.add_heading(&doc_id, "Color Tests", 2)?;
    let colors = vec![
        ("#000000", "Black"),
        ("#FF0000", "Red"),
        ("#00FF00", "Green"),
        ("#0000FF", "Blue"),
        ("#FF00FF", "Magenta"),
        ("#00FFFF", "Cyan"),
        ("#800080", "Purple"),
    ];
    
    for (color_code, color_name) in colors {
        let style = DocxStyle {
            color: Some(color_code.to_string()),
            ..TestStyles::basic()
        };
        handler.add_paragraph(&doc_id, &format!("This text is in {}", color_name), Some(style))?;
    }
    
    // Alignment tests
    handler.add_heading(&doc_id, "Alignment Tests", 2)?;
    let alignments = vec![
        ("left", "Left aligned text"),
        ("center", "Center aligned text"),
        ("right", "Right aligned text"),
        ("justify", "Justified text that should span the full width of the line when there is enough content to make it meaningful"),
    ];
    
    for (alignment, text) in alignments {
        let style = DocxStyle {
            alignment: Some(alignment.to_string()),
            ..TestStyles::basic()
        };
        handler.add_paragraph(&doc_id, text, Some(style))?;
    }
    
    // Complex table with formatting
    handler.add_heading(&doc_id, "Formatted Table", 2)?;
    let formatted_table = TableData {
        rows: vec![
            vec!["Item".to_string(), "Price".to_string(), "Discount".to_string(), "Final Price".to_string()],
            vec!["Widget A".to_string(), "$100.00".to_string(), "10%".to_string(), "$90.00".to_string()],
            vec!["Widget B".to_string(), "$150.00".to_string(), "15%".to_string(), "$127.50".to_string()],
            vec!["Widget C".to_string(), "$200.00".to_string(), "20%".to_string(), "$160.00".to_string()],
            vec!["Total".to_string(), "$450.00".to_string(), "".to_string(), "$377.50".to_string()],
        ],
        headers: Some(vec!["Item".to_string(), "Price".to_string(), "Discount".to_string(), "Final Price".to_string()]),
        border_style: Some("single".to_string()),
        col_widths: None,
        merges: None,
        cell_shading: None,
    };
    handler.add_table(&doc_id, formatted_table)?;
    
    Ok(doc_id)
}