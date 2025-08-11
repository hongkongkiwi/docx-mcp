//! Test data generators and utilities

use serde_json::{json, Value};
use std::collections::HashMap;

/// Generates test data for various document types and scenarios
pub struct TestDataGenerator;

impl TestDataGenerator {
    /// Generate test paragraphs with varying complexity
    pub fn generate_paragraphs(count: usize, complexity: ParagraphComplexity) -> Vec<String> {
        let base_sentences = match complexity {
            ParagraphComplexity::Simple => vec![
                "This is a simple sentence.",
                "Another basic statement follows.",
                "The text remains straightforward.",
                "No complex structures here.",
                "Plain language is used throughout.",
            ],
            ParagraphComplexity::Medium => vec![
                "This sentence demonstrates moderate complexity with additional clauses and descriptive elements.",
                "Furthermore, the content includes various punctuation marks, numbers like 123, and technical terms.",
                "The writing style incorporates both simple and compound sentence structures for variety.",
                "Additionally, references to specific dates (December 15, 2024) and percentages (85%) are included.",
                "These paragraphs simulate realistic document content found in business communications.",
            ],
            ParagraphComplexity::Complex => vec![
                "This comprehensive sentence exemplifies sophisticated linguistic structures, incorporating multiple subordinate clauses, technical terminology, and complex syntactical arrangements that challenge both human readers and automated processing systems.",
                "Moreover, the content integrates diverse elements including numerical data (such as 42.7% improvement rates), temporal references (spanning Q3 2024 through Q1 2025), geographical locations (Silicon Valley, New York, London), and industry-specific jargon that reflects real-world document complexity.",
                "The methodology employed in generating these test paragraphs considers various factors: readability indices, sentence length distribution, vocabulary diversity, and the inclusion of special characters (e.g., àáâãäå, €£¥, ∑∏∫) to ensure comprehensive testing coverage.",
                "Consequently, these multi-faceted paragraphs serve as effective benchmarks for evaluating system performance under realistic conditions, while simultaneously providing sufficient content variation to identify potential edge cases and optimization opportunities.",
            ],
        };
        
        (0..count)
            .map(|i| {
                let sentence_count = match complexity {
                    ParagraphComplexity::Simple => 2 + (i % 3),
                    ParagraphComplexity::Medium => 3 + (i % 4),
                    ParagraphComplexity::Complex => 2 + (i % 3),
                };
                
                let mut paragraph = String::new();
                for j in 0..sentence_count {
                    let sentence = &base_sentences[j % base_sentences.len()];
                    if j > 0 {
                        paragraph.push(' ');
                    }
                    paragraph.push_str(sentence);
                }
                
                paragraph
            })
            .collect()
    }
    
    /// Generate table data with specified dimensions and content type
    pub fn generate_table_data(rows: usize, cols: usize, content_type: TableContentType) -> Vec<Vec<String>> {
        let mut table_data = Vec::new();
        
        // Generate header row
        let headers: Vec<String> = (0..cols)
            .map(|i| match content_type {
                TableContentType::Generic => format!("Column {}", i + 1),
                TableContentType::Financial => match i {
                    0 => "Period".to_string(),
                    1 => "Revenue".to_string(),
                    2 => "Expenses".to_string(),
                    3 => "Profit".to_string(),
                    _ => format!("Metric {}", i + 1),
                },
                TableContentType::Personnel => match i {
                    0 => "Name".to_string(),
                    1 => "Department".to_string(),
                    2 => "Role".to_string(),
                    3 => "Start Date".to_string(),
                    _ => format!("Field {}", i + 1),
                },
                TableContentType::Technical => match i {
                    0 => "Component".to_string(),
                    1 => "Version".to_string(),
                    2 => "Status".to_string(),
                    3 => "Last Updated".to_string(),
                    _ => format!("Attribute {}", i + 1),
                },
            })
            .collect();
        
        table_data.push(headers);
        
        // Generate data rows
        for row in 0..rows {
            let row_data: Vec<String> = (0..cols)
                .map(|col| match content_type {
                    TableContentType::Generic => format!("R{}C{}", row + 1, col + 1),
                    TableContentType::Financial => match col {
                        0 => format!("Q{} 2024", (row % 4) + 1),
                        1 => format!("${:.1}M", 100.0 + row as f64 * 12.5),
                        2 => format!("${:.1}M", 70.0 + row as f64 * 8.2),
                        3 => format!("${:.1}M", 30.0 + row as f64 * 4.3),
                        _ => format!("{:.1}%", 15.0 + row as f64 * 2.1),
                    },
                    TableContentType::Personnel => match col {
                        0 => format!("Employee {}", row + 1),
                        1 => ["Engineering", "Sales", "Marketing", "Operations"][(row % 4)].to_string(),
                        2 => ["Manager", "Developer", "Analyst", "Specialist"][(row % 4)].to_string(),
                        3 => format!("2024-{:02}-{:02}", ((row % 12) + 1), ((row % 28) + 1)),
                        _ => format!("Data {}", row + 1),
                    },
                    TableContentType::Technical => match col {
                        0 => format!("Component-{}", row + 1),
                        1 => format!("v{}.{}.{}", (row % 3) + 1, (row % 5), (row % 10)),
                        2 => ["Active", "Pending", "Deprecated", "Testing"][(row % 4)].to_string(),
                        3 => format!("2024-12-{:02}", ((row % 28) + 1)),
                        _ => format!("Value {}", row + 1),
                    },
                })
                .collect();
            
            table_data.push(row_data);
        }
        
        table_data
    }
    
    /// Generate list items with specified count and category
    pub fn generate_list_items(count: usize, category: ListCategory) -> Vec<String> {
        let base_items = match category {
            ListCategory::Tasks => vec![
                "Complete project documentation",
                "Review code changes and pull requests",
                "Update system configuration files",
                "Run comprehensive test suite",
                "Deploy to staging environment",
                "Conduct security audit",
                "Optimize database performance",
                "Update user interface components",
                "Implement new feature requirements",
                "Fix reported bugs and issues",
            ],
            ListCategory::Features => vec![
                "Advanced search and filtering capabilities",
                "Real-time collaboration tools",
                "Automated backup and recovery",
                "Multi-language support",
                "Mobile-responsive design",
                "Integration with third-party services",
                "Customizable dashboard and reports",
                "Role-based access control",
                "API for external integrations",
                "Advanced analytics and insights",
            ],
            ListCategory::Requirements => vec![
                "System must support 1000+ concurrent users",
                "Response time must be under 200ms for 95% of requests",
                "Uptime must exceed 99.9% availability",
                "Data must be encrypted both in transit and at rest",
                "User interface must be accessible (WCAG 2.1 AA)",
                "System must support multi-factor authentication",
                "Backup processes must complete within 2 hours",
                "Security patches must be applied within 24 hours",
                "System must scale horizontally to handle peak loads",
                "Audit logs must be maintained for minimum 7 years",
            ],
            ListCategory::Benefits => vec![
                "Increased operational efficiency by 35%",
                "Reduced manual processing time by 60%",
                "Improved data accuracy and consistency",
                "Enhanced security and compliance posture",
                "Better user experience and satisfaction",
                "Lower total cost of ownership",
                "Faster time-to-market for new features",
                "Improved scalability and performance",
                "Better decision-making through analytics",
                "Reduced maintenance and support costs",
            ],
        };
        
        (0..count)
            .map(|i| {
                let base_item = &base_items[i % base_items.len()];
                if count > base_items.len() {
                    format!("{} (item {})", base_item, i + 1)
                } else {
                    base_item.clone()
                }
            })
            .collect()
    }
    
    /// Generate realistic business data for testing
    pub fn generate_business_data() -> BusinessDataSet {
        BusinessDataSet {
            companies: vec![
                "Acme Corporation".to_string(),
                "Global Tech Solutions".to_string(),
                "Innovation Partners LLC".to_string(),
                "Digital Dynamics Inc".to_string(),
                "Future Systems Ltd".to_string(),
            ],
            departments: vec![
                "Engineering".to_string(),
                "Sales & Marketing".to_string(),
                "Human Resources".to_string(),
                "Operations".to_string(),
                "Finance & Accounting".to_string(),
                "Research & Development".to_string(),
            ],
            positions: vec![
                "Software Engineer".to_string(),
                "Product Manager".to_string(),
                "Sales Representative".to_string(),
                "Data Analyst".to_string(),
                "Project Manager".to_string(),
                "UX Designer".to_string(),
            ],
            locations: vec![
                "San Francisco, CA".to_string(),
                "New York, NY".to_string(),
                "Austin, TX".to_string(),
                "Seattle, WA".to_string(),
                "Boston, MA".to_string(),
                "Chicago, IL".to_string(),
            ],
        }
    }
    
    /// Generate MCP tool call test data
    pub fn generate_mcp_test_calls() -> Vec<McpTestCall> {
        vec![
            McpTestCall {
                tool_name: "create_document".to_string(),
                args: json!({}),
                expected_success: true,
                expected_result_keys: vec!["success".to_string(), "document_id".to_string()],
            },
            McpTestCall {
                tool_name: "add_paragraph".to_string(),
                args: json!({
                    "document_id": "test-doc-id",
                    "text": "Test paragraph content"
                }),
                expected_success: true,
                expected_result_keys: vec!["success".to_string()],
            },
            McpTestCall {
                tool_name: "add_heading".to_string(),
                args: json!({
                    "document_id": "test-doc-id",
                    "text": "Test Heading",
                    "level": 1
                }),
                expected_success: true,
                expected_result_keys: vec!["success".to_string()],
            },
            McpTestCall {
                tool_name: "extract_text".to_string(),
                args: json!({
                    "document_id": "test-doc-id"
                }),
                expected_success: true,
                expected_result_keys: vec!["success".to_string(), "text".to_string()],
            },
            McpTestCall {
                tool_name: "get_metadata".to_string(),
                args: json!({
                    "document_id": "test-doc-id"
                }),
                expected_success: true,
                expected_result_keys: vec!["success".to_string(), "metadata".to_string()],
            },
        ]
    }
    
    /// Generate performance test scenarios
    pub fn generate_performance_scenarios() -> Vec<PerformanceScenario> {
        vec![
            PerformanceScenario {
                name: "Small Document".to_string(),
                paragraph_count: 10,
                table_count: 1,
                list_count: 2,
                expected_max_time_ms: 1000,
            },
            PerformanceScenario {
                name: "Medium Document".to_string(),
                paragraph_count: 100,
                table_count: 5,
                list_count: 10,
                expected_max_time_ms: 5000,
            },
            PerformanceScenario {
                name: "Large Document".to_string(),
                paragraph_count: 500,
                table_count: 20,
                list_count: 30,
                expected_max_time_ms: 15000,
            },
            PerformanceScenario {
                name: "Extra Large Document".to_string(),
                paragraph_count: 1000,
                table_count: 50,
                list_count: 50,
                expected_max_time_ms: 30000,
            },
        ]
    }
}

/// Complexity levels for generated paragraphs
#[derive(Debug, Clone)]
pub enum ParagraphComplexity {
    Simple,
    Medium,
    Complex,
}

/// Content types for generated tables
#[derive(Debug, Clone)]
pub enum TableContentType {
    Generic,
    Financial,
    Personnel,
    Technical,
}

/// Categories for generated lists
#[derive(Debug, Clone)]
pub enum ListCategory {
    Tasks,
    Features,
    Requirements,
    Benefits,
}

/// Business data set for realistic testing
#[derive(Debug, Clone)]
pub struct BusinessDataSet {
    pub companies: Vec<String>,
    pub departments: Vec<String>,
    pub positions: Vec<String>,
    pub locations: Vec<String>,
}

/// MCP tool call test data
#[derive(Debug, Clone)]
pub struct McpTestCall {
    pub tool_name: String,
    pub args: Value,
    pub expected_success: bool,
    pub expected_result_keys: Vec<String>,
}

/// Performance test scenario data
#[derive(Debug, Clone)]
pub struct PerformanceScenario {
    pub name: String,
    pub paragraph_count: usize,
    pub table_count: usize,
    pub list_count: usize,
    pub expected_max_time_ms: u64,
}

/// Utility functions for test data validation
pub struct TestDataValidator;

impl TestDataValidator {
    /// Validate that text contains expected content
    pub fn validate_text_content(text: &str, expected_keywords: &[&str]) -> bool {
        expected_keywords.iter().all(|keyword| text.contains(keyword))
    }
    
    /// Validate table structure
    pub fn validate_table_structure(rows: &[Vec<String>], expected_cols: usize) -> bool {
        !rows.is_empty() && rows.iter().all(|row| row.len() == expected_cols)
    }
    
    /// Validate MCP response structure
    pub fn validate_mcp_response(response: &Value, expected_keys: &[String]) -> bool {
        expected_keys.iter().all(|key| response.get(key).is_some())
    }
    
    /// Generate hash for test data consistency checking
    pub fn generate_content_hash(content: &str) -> u64 {
        use std::collections::hash_map::DefaultHasher;
        use std::hash::{Hash, Hasher};
        
        let mut hasher = DefaultHasher::new();
        content.hash(&mut hasher);
        hasher.finish()
    }
}