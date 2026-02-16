//! Class file (.cls) parsing errors.

/// Errors that can occur during class file parsing.
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum ClassError {
    /// The 'VERSION' keyword is missing from the class file header.
    #[error("The 'VERSION' keyword is missing from the class file header.")]
    VersionKeywordMissing,

    /// The 'BEGIN' keyword is missing from the class file header.
    #[error("The 'BEGIN' keyword is missing from the class file header.")]
    BeginKeywordMissing,

    /// The 'Class' keyword is missing from the class file header.
    #[error("The 'Class' keyword is missing from the class file header.")]
    KeywordMissing,

    /// Missing whitespace between 'VERSION' keyword and major version number.
    #[error(
        "After the 'VERSION' keyword there should be a space before the major version number."
    )]
    WhitespaceMissingBetweenVersionAndMajorVersionNumber,

    /// The 'VERSION' keyword is not fully uppercase.
    #[error("The 'VERSION' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    VersionKeywordNotFullyUppercase {
        /// The text of the 'VERSION' keyword as found in the source.
        version_text: String,
    },

    /// The 'CLASS' keyword is not fully uppercase.
    #[error("The 'CLASS' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    KeywordNotFullyUppercase {
        /// The text of the 'CLASS' keyword as found in the source.
        class_text: String,
    },

    /// The 'BEGIN' keyword is not fully uppercase.
    #[error("The 'BEGIN' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE.")]
    BeginKeywordNotFullyUppercase {
        /// The text of the 'BEGIN' keyword as found in the source.
        begin_text: String,
    },

    /// The 'END' keyword is not fully uppercase.
    #[error(
        "The 'END' keyword should be in uppercase to be fully compatible with Microsoft's VB6 IDE."
    )]
    EndKeywordNotFullyUppercase {
        /// The text of the 'END' keyword as found in the source.
        end_text: String,
    },

    /// The 'BEGIN' keyword should be on its own line.
    #[error("The 'BEGIN' keyword should stand alone on its own line.")]
    BeginKeywordShouldBeStandAlone,

    /// The 'END' keyword should be on its own line.
    #[error("The 'END' keyword should stand alone on its own line.")]
    EndKeywordShouldBeStandAlone,

    /// Unable to parse the major version number.
    #[error("Unable to parse the major version number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToParseMajorVersionNumber,

    /// Unable to convert the major version text to a number.
    #[error("Unable to convert the major version text to a number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToConvertMajorVersionNumber,

    /// Unable to parse the minor version number.
    #[error("Unable to parse the minor version number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToParseMinorVersionNumber,

    /// Unable to convert the minor version text to a number.
    #[error("Unable to convert the minor version text to a number. Following the 'VERSION' keyword should be a major version number, a '.', and a minor version number.")]
    UnableToConvertMinorVersionNumber,

    /// The period divider between major and minor version digits is missing.
    #[error("The '.' divider between major and minor version digits is missing.")]
    MissingPeriodDividerBetweenMajorAndMinorVersion,

    /// Missing whitespace between minor version digits and 'CLASS' keyword.
    #[error("Missing whitespace between minor version digits and 'CLASS' keyword. This may not be compliant with Microsoft's VB6 IDE.")]
    MissingWhitespaceAfterMinorVersion,

    /// Incorrect whitespace between minor version digits and 'CLASS' keyword.
    #[error("Between the minor version digits and the 'CLASS' keyword should be a single ASCII space. This may not be compliant with Microsoft's VB6 IDE.")]
    IncorrectWhitespaceAfterMinorVersion,

    /// Whitespace was used to divide between major and minor version numbers.
    #[error("Whitespace was used to divide between major and minor version information. This may not be compliant with Microsoft's VB6 IDE.")]
    WhitespaceDividerBetweenMajorAndMinorVersionNumbers,

    /// CST parsing error in class file.
    #[error("CST parsing error: {message}")]
    CSTError {
        /// Error message from CST parsing.
        message: String,
    },
}
