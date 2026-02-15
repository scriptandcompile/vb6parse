//! Lexer and tokenization errors.

/// Errors that can occur during lexical analysis and tokenization.
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum LexerError {
    /// Variable names in VB6 have a maximum length of 255 characters.
    #[error("Variable names in VB6 have a maximum length of 255 characters.")]
    VariableNameTooLong,

    /// Unknown token encountered during parsing.
    #[error("Unknown token '{token}' found.")]
    UnknownToken {
        /// The unknown token that was encountered.
        token: String,
    },

    /// Unexpected end of the code stream.
    #[error("Unexpected end of code stream.")]
    UnexpectedEndOfStream,
}
