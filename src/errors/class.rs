//! Errors related to VB6 class file (.cls) parsing.
//!
//! This module contains error types for issues that occur during:
//! - Class file header parsing (VERSION, BEGIN, CLASS keywords)
//! - Class-specific attribute validation
//! - Class code body parsing

use crate::errors::{CodeErrorKind, ErrorDetails};

/// Errors related to class file parsing.
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum ClassErrorKind<'a> {
    /// Indicates that the 'VERSION' keyword is missing from the class file header.
    #[error("The 'VERSION' keyword is missing from the class file header.")]
    VersionKeywordMissing,

    /// Indicates that the 'BEGIN' keyword is missing from the class file header.
    #[error("The 'BEGIN' keyword is missing from the class file header.")]
    BeginKeywordMissing,

    /// Indicates that the 'Class' keyword is missing from the class file header.
    #[error("The 'Class' keyword is missing from the class file header.")]
    ClassKeywordMissing,

    /// Indicates that there is missing whitespace between the 'VERSION' keyword and the major version number.
    #[error(
        "After the 'VERSION' keyword there should be a space before the major version number."
    )]
    WhitespaceMissingBetweenVersionAndMajorVersionNumber,

    /// Indicates that the 'VERSION' keyword is not fully uppercase.
    #[error("The 'VERSION' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    VersionKeywordNotFullyUppercase {
        /// The text of the 'VERSION' keyword as found in the source.
        version_text: &'a str,
    },

    /// Indicates that the 'CLASS' keyword is not fully uppercase.
    #[error("The 'CLASS' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    ClassKeywordNotFullyUppercase {
        /// The text of the 'CLASS' keyword as found in the source.
        class_text: &'a str,
    },

    /// Indicates that the 'BEGIN' keyword is not fully uppercase.
    #[error("The 'BEGIN' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    BeginKeywordNotFullyUppercase {
        /// The text of the 'BEGIN' keyword as found in the source.
        begin_text: &'a str,
    },

    /// Indicates that the 'END' keyword is not fully uppercase.
    #[error(
        "The 'END' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE."
    )]
    EndKeywordNotFullyUppercase {
        /// The text of the 'END' keyword as found in the source.
        end_text: &'a str,
    },

    /// Indicates that the 'BEGIN' keyword should be on its own line.
    #[error("The 'BEGIN' keyword should stand alone on its own line.")]
    BeginKeywordShouldBeStandAlone,

    /// Indicates that the 'END' keyword should be on its own line.
    #[error("The 'END' keyword should stand alone on its own line.")]
    EndKeywordShouldBeStandAlone,

    /// Indicates that the major version number could not be parsed.
    #[error("Unable to parse the major version number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToParseMajorVersionNumber,

    /// Indicates that the major version text could not be converted to a number.
    #[error("Unable to convert the major version text to a number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToConvertMajorVersionNumber,

    /// Indicates that the minor version number could not be parsed.
    #[error("Unable to parse the minor version number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToParseMinorVersionNumber,

    /// Indicates that the minor version text could not be converted to a number.
    #[error("Unable to convert the minor version text to a number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToConvertMinorVersionNumber,

    /// Indicates that the period divider between major and minor version digits is missing.
    #[error("The '.' divider between major and minor version digits is missing.")]
    MissingPeriodDividerBetweenMajorAndMinorVersion,

    /// Indicates that there is missing whitespace between minor version digits and 'CLASS' keyword.
    #[error("Missing whitespace between minor version digits and 'CLASS' keyword. This may not be compliant with Microsoft's VB6 IDE.")]
    MissingWhitespaceAfterMinorVersion,

    /// Indicates that there is incorrect whitespace between minor version digits and 'CLASS' keyword.
    #[error("Between the minor version digits and the 'CLASS' keyword should be a single ASCII space. This may not be compliant with Microsoft's VB6 IDE.")]
    IncorrectWhitespaceAfterMinorVersion,

    /// Indicates that whitespace was used to divide between major and minor version numbers.
    #[error("Whitespace was used to divide between major and minor version information. This may not be compliant with Microsoft's VB6 IDE.")]
    WhitespaceDividerBetweenMajorAndMinorVersionNumbers,

    /// Indicates that there was an error parsing VB6 tokens.
    #[error("There was an error parsing the VB6 tokens.")]
    ClassTokenError {
        /// The underlying code error that occurred.
        code_error: CodeErrorKind,
    },

    /// Indicates that there was an error parsing the CST.
    #[error("CST parsing error: {0}")]
    CSTError(String),
}

impl<'a> From<ErrorDetails<'a, CodeErrorKind>> for ErrorDetails<'a, ClassErrorKind<'a>> {
    fn from(value: ErrorDetails<'a, CodeErrorKind>) -> Self {
        ErrorDetails {
            source_content: value.source_content,
            source_name: value.source_name,
            error_offset: value.error_offset,
            line_start: value.line_start,
            line_end: value.line_end,
            kind: ClassErrorKind::ClassTokenError {
                code_error: value.kind,
            },
            severity: value.severity,
            labels: value.labels,
            notes: value.notes,
        }
    }
}
