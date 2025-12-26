//! Errors related to VB6 module file (.bas) parsing.
//!
//! This module contains error types for issues that occur during:
//! - Module file header parsing (`Attribute VB_Name`)
//! - Module-specific attribute validation
//! - Module code body parsing

use crate::errors::{CodeErrorKind, ErrorDetails};

/// Errors related to module file parsing.
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum ModuleErrorKind {
    /// Indicates that the 'Attribute' keyword is missing from the module file header.
    #[error("The 'Attribute' keyword is missing from the module file header.")]
    AttributeKeywordMissing,

    /// Indicates that there is missing whitespace between the 'Attribute' keyword and the '`VB_Name`' attribute.
    #[error("The 'Attribute' keyword and the 'VB_Name' attribute must be separated by at least one ASCII whitespace character.")]
    MissingWhitespaceInHeader,

    /// Indicates that the '`VB_Name`' attribute is missing from the module file header.
    #[error("The 'VB_Name' attribute is missing from the module file header.")]
    VBNameAttributeMissing,

    /// Indicates that the '`VB_Name`' attribute is missing the equal symbol.
    #[error("The 'VB_Name' attribute is missing the equal symbol from the module file header.")]
    EqualMissing,

    /// Indicates that the '`VB_Name`' attribute value is unquoted.
    #[error("The 'VB_Name' attribute is unquoted.")]
    VBNameAttributeValueUnquoted,

    /// Indicates that there was an error parsing VB6 tokens.
    #[error("There was an error parsing the VB6 tokens.")]
    ModuleTokenError {
        /// The underlying code error that occurred.
        code_error: CodeErrorKind,
    },
}

impl<'a> From<ErrorDetails<'a, CodeErrorKind>> for ErrorDetails<'a, ModuleErrorKind> {
    fn from(value: ErrorDetails<'a, CodeErrorKind>) -> Self {
        ErrorDetails {
            source_content: value.source_content,
            source_name: value.source_name,
            error_offset: value.error_offset,
            line_start: value.line_start,
            line_end: value.line_end,
            kind: ModuleErrorKind::ModuleTokenError {
                code_error: value.kind,
            },
        }
    }
}
