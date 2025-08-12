use serde::{Serialize, Deserialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum ToolOutcome {
    Ok { message: Option<String> },
    Created { document_id: String, message: Option<String> },
    Text { text: String },
    Metadata { metadata: serde_json::Value },
    Documents { documents: serde_json::Value },
    Images { images: Vec<String>, message: Option<String> },
    Security { security: serde_json::Value },
    Storage { storage: serde_json::Value },
    Statistics { statistics: serde_json::Value },
    Structure { structure: serde_json::Value },
    Error { code: ErrorCode, error: String, hint: Option<String> },
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "SCREAMING_SNAKE_CASE")]
pub enum ErrorCode {
    DocNotFound,
    ValidationError,
    SecurityDenied,
    LimitExceeded,
    UnknownTool,
    InternalError,
}

impl ToolOutcome {
    pub fn success(&self) -> bool {
        !matches!(self, ToolOutcome::Error { .. })
    }

    pub fn into_json(self) -> serde_json::Value {
        serde_json::to_value(self).unwrap_or_else(|e| serde_json::json!({
            "type": "error",
            "code": ErrorCode::InternalError,
            "error": format!("serialization failed: {}", e),
        }))
    }
}
