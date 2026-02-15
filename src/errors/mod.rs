//! Module containing the error types used in the VB6 parser.
//!
//! This module is organized by the layer or file type that produces the error:
//! - [`decode`] - Source file decoding errors (Windows-1252, malformed files)
//! - [`tokenize`] - Tokenization and basic code parsing errors
//! - [`resource`] - Form resource file (FRX) parsing errors
//! - [`class`] - Class file (CLS) specific errors
//! - [`module`] - Module file (BAS) specific errors
//! - [`project`] - Project file (VBP) specific errors
//! - [`form`] - Form file (FRM) specific errors
//! - [`property`] - Property value validation errors
//!
//! The [`ErrorDetails`] type is the central error container that wraps any of these
//! error kinds along with source location information for diagnostic reporting.

use ariadne::{Label, Report, ReportKind, Source};
use core::convert::From;
use std::error::Error;
use std::fmt::{Debug, Display};

// Submodules
pub mod class;
pub mod decode;
pub mod form;
pub mod module;
pub mod project;
pub mod property;
pub mod resource;
pub mod tokenize;

// Re-export error kinds for convenience
pub use class::ClassErrorKind;
pub use decode::SourceFileErrorKind;
pub use form::FormErrorKind;
pub use module::ModuleErrorKind;
pub use project::ProjectErrorKind;
pub use property::PropertyError;
pub use resource::ResourceErrorKind;
pub use tokenize::CodeErrorKind;

/// Represents the severity level of a parsing diagnostic.
///
/// This enum is used to distinguish between different types of issues
/// encountered during parsing, from informational notes to fatal errors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum Severity {
    /// Informational message, not a problem.
    Note,
    /// Potential issue that should be addressed but doesn't prevent usage.
    Warning,
    /// Fatal error that prevents successful parsing or usage.
    Error,
}

impl Default for Severity {
    fn default() -> Self {
        Severity::Error
    }
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

/// Contains detailed information about an error that occurred during parsing.
/// This struct contains the source name, source content, error offset,
/// line start and end positions, and the kind of error.
///
/// Generic parameter `T` represents the type of error kind.
/// It must implement the `ToString` trait to allow for error message formatting.
///
/// Example usage:
/// ```rust
/// use vb6parse::errors::ErrorDetails;
/// use vb6parse::errors::CodeErrorKind;
///
/// let error_details = ErrorDetails {
///     source_name: "example.cls".to_string().into_boxed_str(),
///     source_content: "Some VB6 code here...",
///     error_offset: 10,
///     line_start: 1,
///     line_end: 1,
///     kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
/// };
/// error_details.print();
/// ```
#[derive(Debug, Clone)]
pub struct ErrorDetails<'a, T>
where
    T: ToString + Debug,
{
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
    pub kind: T,
    /// The severity of this diagnostic (Error, Warning, or Note).
    pub severity: Severity,
}

impl<T> Display for ErrorDetails<'_, T>
where
    T: ToString + Debug,
{
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

impl<T> ErrorDetails<'_, T>
where
    T: ToString + Debug,
{
    /// Print the `ErrorDetails` using ariadne for formatting
    ///
    /// Example usage:
    /// ```rust
    /// use vb6parse::errors::ErrorDetails;
    /// use vb6parse::errors::CodeErrorKind;
    ///
    /// let error_details = ErrorDetails {
    /// source_name: "example.cls".to_string().into_boxed_str(),
    ///   source_content: "Some VB6 code here...",
    ///   error_offset: 10,
    ///   line_start: 1,
    ///   line_end: 1,
    ///   kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
    /// };
    /// error_details.print();
    /// ```
    pub fn print(&self) {
        let cache = (
            self.source_name.to_string(),
            Source::from(self.source_content),
        );

        let report = Report::build(
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
        .print(cache);

        if let Some(e) = report.err() {
            eprint!("Error attempting to build ErrorDetails print message {e:?}");
        }
    }

    /// Eprint the `ErrorDetails` using ariadne for formatting
    ///
    /// Example usage:
    /// ```rust
    /// use vb6parse::errors::ErrorDetails;
    /// use vb6parse::errors::CodeErrorKind;
    /// let error_details = ErrorDetails {
    ///     source_name: "example.cls".to_string().into_boxed_str(),
    ///     source_content: "Some VB6 code here...",
    ///     error_offset: 10,
    ///     line_start: 1,
    ///     line_end: 1,
    ///     kind: CodeErrorKind::UnknownToken {
    ///         token: "???".to_string(),
    ///     },
    /// };
    /// error_details.eprint();
    /// ```
    pub fn eprint(&self) {
        let cache = (
            self.source_name.to_string(),
            Source::from(self.source_content),
        );

        let report = Report::build(
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
        )
        .finish()
        .eprint(cache);

        if let Some(e) = report.err() {
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
