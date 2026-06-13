//! Module file (.bas) parsing errors.

/// Errors that can occur during module file parsing.
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum ModuleError {
    /// The 'Attribute' keyword is missing from the module file header.
    #[error("The 'Attribute' keyword is missing from the module file header.")]
    AttributeKeywordMissing,

    /// Missing whitespace in module header.
    #[error("The 'Attribute' keyword and the 'VB_Name' attribute must be separated by at least one ASCII whitespace character.")]
    MissingWhitespaceInHeader,

    /// The `VB_Name` attribute is missing from the module file header.
    #[error("The 'VB_Name' attribute is missing from the module file header.")]
    VBNameAttributeMissing,

    /// The `VB_Name` attribute is missing the equal symbol.
    #[error("The 'VB_Name' attribute is missing the equal symbol from the module file header.")]
    EqualMissing,

    /// The `VB_Name` attribute value is unquoted.
    #[error("The 'VB_Name' attribute is unquoted.")]
    VBNameAttributeValueUnquoted,

    /// Control flow nesting exceeds maximum depth
    #[error("Control flow nesting depth ({depth}) exceeds maximum allowed depth ({max_depth})")]
    NestingTooDeep {
        /// The actual nesting depth encountered
        depth: usize,
        /// The maximum allowed nesting depth
        max_depth: usize,
    },

    /// Invalid semicolon separator in a non-print procedure call.
    #[error("Invalid ';' separator in procedure call arguments. Semicolon separators are only valid for print-like calls (e.g., Debug.Print, Printer.Print).")]
    InvalidSemicolonSeparatorInProcedureCall,
}
