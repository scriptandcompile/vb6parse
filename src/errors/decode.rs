//! Errors related to source file decoding and parsing.
//!
//! This module contains error types for issues that occur during:
//! - Source file reading and Windows-1252 decoding
//! - File format validation

/// Errors related to source file parsing and decoding.
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum SourceFileErrorKind {
    /// Indicates that the source file is malformed.
    #[error("Unable to parse source file: {message}")]
    MalformedSource {
        /// The error message describing the issue.
        message: String,
    },
}
