//! Errors related to VB6 code tokenization and parsing.
//!
//! This module contains error types for issues that occur during:
//! - Tokenization of VB6 source code
//! - Basic syntax validation during parsing

/// Errors related to code parsing and tokenization.
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum CodeErrorKind {
    /// Indicates that a variable name exceeds the maximum allowed length in VB6.
    #[error("Variable names in VB6 have a maximum length of 255 characters.")]
    VariableNameTooLong,

    /// Indicates that an unknown token was encountered during parsing.
    #[error("Unknown token '{token}' found.")]
    UnknownToken {
        /// The unknown token that was encountered.
        token: String,
    },

    /// Indicates that an unexpected end of the code stream was encountered.
    #[error("Unexpected end of code stream.")]
    UnexpectedEndOfStream,
}
