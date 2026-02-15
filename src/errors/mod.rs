//! Module containing the error types used in the VB6 parser.
//!
//! This module is organized by the layer or file type that produces the error:
//! - [`resource`] - Form resource file (FRX) parsing errors
//! - [`class`] - Class file (CLS) specific errors
//! - [`module`] - Module file (BAS) specific errors
//! - [`project`] - Project file (VBP) specific errors
//! - [`form`] - Form file (FRM) specific errors
//! - [`lexer`] - Lexing errors
//! - [`source`] - Source file decoding errors
//!
//! The [`ErrorDetails`] type is the central error container that wraps any of these
//! error kinds along with source location information for diagnostic reporting.

use ariadne::{Label, Report, ReportKind, Source};
use core::convert::From;
use std::error::Error;
use std::fmt::{Debug, Display};

// Layer-specific error kind modules
pub mod class;
pub mod form;
pub mod lexer;
pub mod module;
pub mod project;
pub mod resource;
pub mod source;

// Re-export all error types for convenience
pub use class::ClassError;
pub use form::FormError;
pub use lexer::LexerError;
pub use module::ModuleError;
pub use project::ProjectError;
pub use resource::ResourceError;
pub use source::SourceFileError;

/// Hierarchical error kind enum that wraps layer-specific error types.
///
/// This enum organizes parsing errors by the layer that produces them,
/// providing better organization and clearer error categorization.
/// Each variant wraps a layer-specific error enum:
///
/// - `Lexer` - Tokenization and lexical analysis errors  
/// - `Class` - Class file (.cls) specific parsing errors
/// - `Module` - Module file (.bas) specific parsing errors
/// - `Form` - Form file (.frm) validation and parsing errors
/// - `Project` - Project file (.vbp) parsing errors
/// - `Resource` - Resource file (.frx) binary data errors
/// - `SourceFile` - File encoding and decoding errors
///
/// # Ergonomic Conversions
///
/// All layer-specific error types implement `From` conversion to `ErrorKind`,
/// allowing automatic conversion. The error generation methods on [`SourceStream`]
/// and [`ErrorDetails::basic()`] accept any type that implements `Into<ErrorKind>`,
/// so you can pass layer-specific errors directly without wrapping:
///
/// ```rust
/// use vb6parse::io::SourceStream;
/// use vb6parse::errors::{ErrorKind, LexerError, ModuleError};
///
/// let stream = SourceStream::new("test.bas", "Dim x");
///
/// // Old way - manual wrapping:
/// let error1 = stream.generate_error(ErrorKind::Lexer(
///     LexerError::UnknownToken { token: "???".to_string() }
/// ));
///
/// // New way - automatic conversion:
/// let error2 = stream.generate_error(
///     LexerError::UnknownToken { token: "???".to_string() }
/// );
///
/// // Works with all layer-specific error types:
/// let error3 = stream.generate_error(
///     ModuleError::AttributeKeywordMissing
/// );
/// ```
///
/// This also works with [`ErrorDetails::basic()`] and all three `generate_error*`
/// methods on [`SourceStream`].
///
/// [`SourceStream`]: crate::io::SourceStream
#[derive(thiserror::Error, Debug, Clone, PartialEq, Eq)]
pub enum ErrorKind {
    /// Lexer and tokenization errors.
    #[error(transparent)]
    Lexer(#[from] LexerError),

    /// Class file parsing errors.
    #[error(transparent)]
    Class(#[from] ClassError),

    /// Module file parsing errors.
    #[error(transparent)]
    Module(#[from] ModuleError),

    /// Form file parsing and validation errors.
    #[error(transparent)]
    Form(#[from] FormError),

    /// Project file parsing errors.
    #[error(transparent)]
    Project(#[from] ProjectError),

    /// Resource file parsing errors.
    #[error(transparent)]
    Resource(#[from] ResourceError),

    /// Source file decoding errors.
    #[error(transparent)]
    SourceFile(#[from] SourceFileError),
}

/// Represents the severity level of a parsing diagnostic.
///
/// This enum is used to distinguish between different types of issues
/// encountered during parsing, from informational notes to fatal errors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Default)]
pub enum Severity {
    /// Informational message, not a problem.
    Note,
    /// Potential issue that should be addressed but doesn't prevent usage.
    Warning,
    /// Fatal error that prevents successful parsing or usage.
    #[default]
    Error,
}

impl Display for Severity {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Severity::Note => write!(f, "note"),
            Severity::Warning => write!(f, "warning"),
            Severity::Error => write!(f, "error"),
        }
    }
}

/// Represents a span of source code, typically associated with an error or diagnostic.
///
/// A span identifies a region in the source code by offset, line numbers, and length.
/// This is used to highlight the exact location of errors in diagnostic messages.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Span {
    /// The byte offset into the source content where this span starts.
    pub offset: u32,
    /// The starting line number (1-based).
    pub line_start: u32,
    /// The ending line number (1-based).
    pub line_end: u32,
    /// The length of this span in bytes.
    pub length: u32,
}

impl Span {
    /// Creates a new span.
    #[must_use]
    pub fn new(offset: u32, line_start: u32, line_end: u32, length: u32) -> Self {
        Self {
            offset,
            line_start,
            line_end,
            length,
        }
    }

    /// Creates a zero-length span at offset 0.
    #[must_use]
    pub fn zero() -> Self {
        Self {
            offset: 0,
            line_start: 0,
            line_end: 0,
            length: 0,
        }
    }

    /// Creates a span of length 1 at the given offset and line.
    #[must_use]
    pub fn at(offset: u32, line: u32) -> Self {
        Self {
            offset,
            line_start: line,
            line_end: line,
            length: 1,
        }
    }
}

/// Represents a labeled span in a multi-span diagnostic.
///
/// Labels are used to annotate multiple locations in the source code
/// within a single error message, providing context for complex errors.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DiagnosticLabel {
    /// The span this label refers to.
    pub span: Span,
    /// The message to display for this label.
    pub message: String,
}

impl DiagnosticLabel {
    /// Creates a new label.
    pub fn new(span: Span, message: impl Into<String>) -> Self {
        Self {
            span,
            message: message.into(),
        }
    }
}

/// Contains detailed information about an error that occurred during parsing.
///
/// This struct contains the source name, source content, error offset,
/// line start and end positions, and the kind of error. All errors now use
/// the unified [`ErrorKind`] type.
///
/// Example usage:
/// ```rust
/// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity, LexerError};
///
/// let error_details = ErrorDetails {
///     source_name: "example.cls".to_string().into_boxed_str(),
///     source_content: "Some VB6 code here...",
///     error_offset: 10,
///     line_start: 1,
///     line_end: 1,
///     kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
///     severity: Severity::Error,
///     labels: vec![],
///     notes: vec![],
/// };
/// error_details.print();
/// ```
#[derive(Debug, Clone)]
pub struct ErrorDetails<'a> {
    /// The name of the source file where the error occurred.
    pub source_name: Box<str>,
    /// The content of the source file where the error occurred.
    pub source_content: &'a str,
    /// The offset in the source content where the error occurred.
    ///
    /// Note: This is a u32 to reflect VB6's 32-bit addressing limitations.
    pub error_offset: u32,
    /// The starting line number of the error.
    ///
    /// Note: This is a u32 to reflect VB6's 32-bit addressing limitations.
    pub line_start: u32,
    /// The ending line number of the error.
    ///
    /// Note: This is a u32 to reflect VB6's 32-bit addressing limitations.
    pub line_end: u32,
    /// The kind of error that occurred.
    /// Boxed to reduce the size of `Result<T, ErrorDetails>` on the stack.
    pub kind: Box<ErrorKind>,
    /// The severity of this diagnostic (`Error`, `Warning`, or `Note`).
    pub severity: Severity,
    /// Additional labeled spans for multi-span diagnostics.
    /// This allows annotating multiple locations in the source code
    /// within a single error message.
    pub labels: Vec<DiagnosticLabel>,
    /// Additional notes to provide context for this diagnostic.
    /// These are displayed after the main error message.
    pub notes: Vec<String>,
}

impl<'a> ErrorDetails<'a> {
    /// Creates a basic `ErrorDetails` with no labels or notes.
    ///
    /// This is a convenience constructor for the common case where
    /// only the basic error information is needed.
    ///
    /// Accepts any error type that can be converted to `ErrorKind`, including
    /// layer-specific errors like `LexerError`, `ModuleError`, `ProjectError`, etc.
    #[must_use]
    pub fn basic<E>(
        source_name: Box<str>,
        source_content: &'a str,
        error_offset: u32,
        line_start: u32,
        line_end: u32,
        kind: E,
        severity: Severity,
    ) -> ErrorDetails<'a>
    where
        E: Into<ErrorKind>,
    {
        ErrorDetails {
            source_name,
            source_content,
            error_offset,
            line_start,
            line_end,
            kind: Box::new(kind.into()),
            severity,
            labels: Vec::new(),
            notes: Vec::new(),
        }
    }

    /// Adds a labeled span to this error.
    #[must_use]
    pub fn with_label(mut self, label: DiagnosticLabel) -> Self {
        self.labels.push(label);
        self
    }

    /// Adds multiple labeled spans to this error.
    #[must_use]
    pub fn with_labels(mut self, labels: Vec<DiagnosticLabel>) -> Self {
        self.labels.extend(labels);
        self
    }

    /// Adds a note to this error.
    #[must_use]
    pub fn with_note(mut self, note: impl Into<String>) -> Self {
        self.notes.push(note.into());
        self
    }

    /// Adds multiple notes to this error.
    #[must_use]
    pub fn with_notes(mut self, notes: Vec<String>) -> Self {
        self.notes.extend(notes);
        self
    }
}

impl Display for ErrorDetails<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "ErrorDetails {{ source_name: {}, error_offset: {}, line_start: {}, line_end: {}, kind: {:?} }}",
            self.source_name,
            self.error_offset,
            self.line_start,
            self.line_end,
            self.kind,
        )
    }
}

impl ErrorDetails<'_> {
    /// Print the `ErrorDetails` using ariadne for formatting.
    ///
    /// Example usage:
    /// ```rust
    /// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity, LexerError};
    ///
    /// let error_details = ErrorDetails {
    ///     source_name: "example.cls".to_string().into_boxed_str(),
    ///     source_content: "Some VB6 code here...",
    ///     error_offset: 10,
    ///     line_start: 1,
    ///     line_end: 1,
    ///     kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
    ///     severity: Severity::Error,
    ///     labels: vec![],
    ///     notes: vec![],
    /// };
    /// error_details.print();
    /// ```
    pub fn print(&self) {
        let cache = (
            self.source_name.to_string(),
            Source::from(self.source_content),
        );

        let mut report = Report::build(
            ReportKind::Error,
            (
                self.source_name.to_string(),
                self.line_start as usize..=self.line_end as usize,
            ),
        )
        .with_message(self.kind.to_string())
        .with_label(
            Label::new((
                self.source_name.to_string(),
                self.error_offset as usize..=self.error_offset as usize,
            ))
            .with_message("error here"),
        );

        // Add additional labeled spans
        for label in &self.labels {
            report = report.with_label(
                Label::new((
                    self.source_name.to_string(),
                    label.span.offset as usize
                        ..=(label.span.offset + label.span.length.max(1) - 1) as usize,
                ))
                .with_message(&label.message),
            );
        }

        // Add notes
        for note in &self.notes {
            report = report.with_note(note);
        }

        let result = report.finish().print(cache);

        if let Some(e) = result.err() {
            eprint!("Error attempting to build ErrorDetails print message {e:?}");
        }
    }

    /// Eprint the `ErrorDetails` using ariadne for formatting.
    ///
    /// Example usage:
    /// ```rust
    /// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity, LexerError};
    ///
    /// let error_details = ErrorDetails {
    ///     source_name: "example.cls".to_string().into_boxed_str(),
    ///     source_content: "Some VB6 code here...",
    ///     error_offset: 10,
    ///     line_start: 1,
    ///     line_end: 1,
    ///     kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken {
    ///         token: "???".to_string(),
    ///     })),
    ///     severity: Severity::Error,
    ///     labels: vec![],
    ///     notes: vec![],
    /// };
    /// error_details.eprint();
    /// ```
    pub fn eprint(&self) {
        let cache = (
            self.source_name.to_string(),
            Source::from(self.source_content),
        );

        let mut report = Report::build(
            ReportKind::Error,
            (
                self.source_name.to_string(),
                self.line_start as usize..=self.line_end as usize,
            ),
        )
        .with_message(format!("{:?}", self.kind))
        .with_label(
            Label::new((
                self.source_name.to_string(),
                self.error_offset as usize..=self.error_offset as usize,
            ))
            .with_message("error here"),
        );

        // Add additional labeled spans
        for label in &self.labels {
            report = report.with_label(
                Label::new((
                    self.source_name.to_string(),
                    label.span.offset as usize
                        ..=(label.span.offset + label.span.length.max(1) - 1) as usize,
                ))
                .with_message(&label.message),
            );
        }

        // Add notes
        for note in &self.notes {
            report = report.with_note(note);
        }

        let result = report.finish().eprint(cache);

        if let Some(e) = result.err() {
            eprint!("Error attempting to build ErrorDetails eprint message {e:?}");
        }
    }

    /// Convert the `ErrorDetails` into a string using ariadne for formatting
    ///
    /// # Errors
    /// This function will return an error if there is an issue converting the
    /// formatted report into a UTF-8 string.
    pub fn print_to_string(&self) -> Result<String, Box<dyn Error>> {
        let cache = (
            self.source_name.to_string(),
            Source::from(self.source_content),
        );

        let mut buf = Vec::new();

        let _ = Report::build(
            ReportKind::Error,
            (
                self.source_name.to_string(),
                self.line_start as usize..=self.line_end as usize,
            ),
        )
        .with_message(self.kind.to_string())
        .with_label(
            Label::new((
                self.source_name.to_string(),
                self.error_offset as usize..=self.error_offset as usize,
            ))
            .with_message("error here"),
        )
        .finish()
        .write(cache, &mut buf);

        let text = String::from_utf8(buf.clone())?;

        Ok(text)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::io::SourceStream;

    #[test]
    fn test_automatic_error_conversion() {
        // Test that layer-specific errors can be automatically converted to ErrorKind
        // and used with generate_error methods

        let stream = SourceStream::new("test.bas", "Dim x As Integer");

        // Test with LexerError
        let _error1 = stream.generate_error(LexerError::UnknownToken {
            token: "???".to_string(),
        });

        // Test with ModuleError
        let _error2 = stream.generate_error(ModuleError::AttributeKeywordMissing);

        // Test with ClassError
        let _error3 = stream.generate_error(ClassError::VersionKeywordMissing);

        // Test with ProjectError
        let _error4 = stream.generate_error(ProjectError::UnterminatedSectionHeader);

        // Test that ErrorDetails::basic also accepts layer-specific errors
        let _error5 = ErrorDetails::basic(
            "test.bas".to_string().into_boxed_str(),
            "Dim x",
            0,
            0,
            5,
            FormError::VersionKeywordMissing,
            Severity::Error,
        );

        // All of the above should compile without needing manual ErrorKind wrapping
    }

    #[test]
    fn test_error_kind_conversion() {
        // Test that Into<ErrorKind> works for all layer-specific error types
        let lexer_err: ErrorKind = LexerError::UnknownToken {
            token: "test".to_string(),
        }
        .into();
        assert!(matches!(lexer_err, ErrorKind::Lexer(_)));

        let module_err: ErrorKind = ModuleError::AttributeKeywordMissing.into();
        assert!(matches!(module_err, ErrorKind::Module(_)));

        let class_err: ErrorKind = ClassError::VersionKeywordMissing.into();
        assert!(matches!(class_err, ErrorKind::Class(_)));

        let project_err: ErrorKind = ProjectError::UnterminatedSectionHeader.into();
        assert!(matches!(project_err, ErrorKind::Project(_)));

        let form_err: ErrorKind = FormError::VersionKeywordMissing.into();
        assert!(matches!(form_err, ErrorKind::Form(_)));

        let resource_err: ErrorKind = ResourceError::OffsetOutOfBounds {
            offset: 0,
            file_length: 10,
        }
        .into();
        assert!(matches!(resource_err, ErrorKind::Resource(_)));

        let source_err: ErrorKind = SourceFileError::Malformed {
            message: "test".to_string(),
        }
        .into();
        assert!(matches!(source_err, ErrorKind::SourceFile(_)));
    }
}
