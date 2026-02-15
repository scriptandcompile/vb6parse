//! Source file decoding errors.

/// Errors that can occur during source file decoding (Windows-1252 encoding, etc.).
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum SourceFileError {
    /// Source file is malformed and cannot be decoded.
    #[error("Source file is malformed")]
    Malformed {
        /// Description of the malformation.
        message: String,
    },
}
