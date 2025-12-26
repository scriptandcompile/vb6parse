//! Errors related to VB6 form resource file (FRX) parsing.
//!
//! This module contains error types for issues that occur during:
//! - Resource file (.frx) reading and parsing
//! - Binary data extraction from resource files
//! - Resource offset and size validation

/// Errors related to resource file parsing.
#[derive(thiserror::Error, Debug)]
pub enum ResourceErrorKind {
    /// I/O error while reading the resource file
    #[error("Failed to read resource file: {0}")]
    IoError(#[from] std::io::Error),

    /// Requested offset is beyond the end of the file
    #[error("Offset {offset} is out of bounds for file of length {file_length}")]
    OffsetOutOfBounds {
        /// The offset that is out of bounds
        offset: usize,
        /// The length of the file
        file_length: usize,
    },

    /// Invalid or corrupted data at the specified offset
    #[error("Invalid data at offset {offset}: {details}")]
    InvalidData {
        /// The offset where the invalid data was found
        offset: usize,
        /// Details about the invalid data
        details: String,
    },

    /// Failed to read header bytes at the specified offset
    #[error("Failed to read header at offset {offset}: {reason}")]
    HeaderReadError {
        /// The offset where the read error occurred
        offset: usize,
        /// The reason for the read error
        reason: String,
    },

    /// Record size fields don't match expected values
    #[error("Record size mismatch at offset {offset}: expected {expected}, got {actual}")]
    SizeMismatch {
        /// The offset where the size mismatch occurred
        offset: usize,
        /// The expected size
        expected: usize,
        /// The actual size found
        actual: usize,
    },

    /// Buffer slice conversion failed (e.g., `try_into` for `[u8; N]`)
    #[error("Failed to convert buffer slice at offset {offset} to fixed-size array")]
    BufferConversionError {
        /// The offset where the conversion error occurred
        offset: usize,
    },

    /// Detected corruption in list items structure
    #[error("Corrupted list items at offset {offset}: {details}")]
    CorruptedListItems {
        /// The offset where the corruption was detected
        offset: usize,
        /// Details about the corruption
        details: String,
    },
}
