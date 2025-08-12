pub mod security;
pub mod fonts_cli;
pub mod response;

// Expose primary modules for tests and external use
pub mod docx_tools;
pub mod docx_handler;
pub mod pure_converter;
pub mod converter;
#[cfg(feature = "advanced-docx")]
pub mod advanced_docx;

pub use security::{Args, SecurityConfig, SecurityMiddleware, SecurityError};
