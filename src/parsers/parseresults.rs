//! Module for parsing results and error handling in the VB6 parser.
//! This module defines the `ParseResult` structure, which encapsulates the outcome of parsing operations,
//! including successful results and any errors encountered during parsing.
//!
//! The `ParseResult` structure is generic over the type of the successful result and the type of error details,
//! allowing it to be used flexibly across different parsing scenarios within the VB6 parser.
//!

use core::convert::From;
use std::fmt::Display;
use std::iter::IntoIterator;

use crate::errors::{ErrorDetails, Severity};
use crate::lexer::TokenStream;

/// Collection of parsing diagnostics, separating errors from warnings.
///
/// This struct provides a more nuanced view of parsing issues by distinguishing
/// between fatal errors and warnings that don't prevent usage.
///
/// # Examples
/// ```rust
/// use vb6parse::parsers::parseresults::Diagnostics;
/// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity, LexerError};
///
/// let error = ErrorDetails {
///     source_name: "test.bas".to_string().into_boxed_str(),
///     source_content: "Some source code",
///     error_offset: 5,
///     line_start: 0,
///     line_end: 10,
///     kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
///     severity: Severity::Error,
///     labels: vec![],
///     notes: vec![],
/// };
///
/// let diagnostics = Diagnostics {
///     errors: vec![error],
///     warnings: vec![],
/// };
///
/// assert!(diagnostics.has_errors());
/// assert!(!diagnostics.has_warnings());
/// ```
#[derive(Debug, Clone, Default)]
pub struct Diagnostics<'a> {
    /// Fatal errors that prevent successful parsing or usage.
    pub errors: Vec<ErrorDetails<'a>>,
    /// Warnings that should be addressed but don't prevent usage.
    pub warnings: Vec<ErrorDetails<'a>>,
}

impl<'a> Diagnostics<'a> {
    /// Creates a new empty Diagnostics instance.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Creates Diagnostics from a list of `ErrorDetails`, categorizing by severity.
    #[must_use]
    pub fn from_details(details: Vec<ErrorDetails<'a>>) -> Self {
        let mut diagnostics = Self::new();
        for detail in details {
            match detail.severity {
                Severity::Error => diagnostics.errors.push(detail),
                Severity::Warning | Severity::Note => diagnostics.warnings.push(detail),
            }
        }
        diagnostics
    }

    /// Returns true if there are any errors.
    #[must_use]
    pub fn has_errors(&self) -> bool {
        !self.errors.is_empty()
    }

    /// Returns true if there are any warnings.
    #[must_use]
    pub fn has_warnings(&self) -> bool {
        !self.warnings.is_empty()
    }

    /// Returns true if there are any diagnostics (errors or warnings).
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.errors.is_empty() && self.warnings.is_empty()
    }

    /// Returns the total number of diagnostics.
    #[must_use]
    pub fn len(&self) -> usize {
        self.errors.len() + self.warnings.len()
    }

    /// Returns an iterator over all diagnostics (errors first, then warnings).
    pub fn iter(&self) -> impl Iterator<Item = &ErrorDetails<'a>> {
        self.errors.iter().chain(self.warnings.iter())
    }

    /// Returns an iterator over all errors.
    pub fn errors_iter(&self) -> impl Iterator<Item = &ErrorDetails<'a>> {
        self.errors.iter()
    }

    /// Returns an iterator over all warnings.
    pub fn warnings_iter(&self) -> impl Iterator<Item = &ErrorDetails<'a>> {
        self.warnings.iter()
    }

    /// Adds an error to the diagnostics.
    pub fn push_error(&mut self, error: ErrorDetails<'a>) {
        self.errors.push(error);
    }

    /// Adds a warning to the diagnostics.
    pub fn push_warning(&mut self, warning: ErrorDetails<'a>) {
        self.warnings.push(warning);
    }

    /// Adds a diagnostic, automatically categorizing by severity.
    pub fn push(&mut self, detail: ErrorDetails<'a>) {
        match detail.severity {
            Severity::Error => self.errors.push(detail),
            Severity::Warning | Severity::Note => self.warnings.push(detail),
        }
    }

    /// Merges another Diagnostics into this one.
    pub fn merge(&mut self, other: Diagnostics<'a>) {
        self.errors.extend(other.errors);
        self.warnings.extend(other.warnings);
    }
}

/// Result of a parsing operation, containing an optional result and a list of failures encountered during parsing.
/// The result is `Some` if parsing was successful, and `None` if it failed completely.
/// Failures are collected in a vector, allowing for partial successes with warnings.
///
/// # Type Parameters
/// * `'a`: Lifetime parameter for error details.
/// * `T`: The type of the successful parse result.
///
/// `ParseResult` is used across the parsing module to encapsulate the outcome of parsing operations,
/// providing both the parsed data (if any) and any errors or warnings that occurred. `ParseResult` is
/// used instead of `Result` to allow for partial successes where some data may be parsed correctly
/// while still reporting errors. This is particularly useful in scenarios where complete failure is not necessary
/// to halt processing, and where users may want to see all issues in a single pass.
///
/// # Examples
/// ```rust
/// use vb6parse::parsers::parseresults::ParseResult;
/// use vb6parse::errors::ErrorDetails;
///
/// let success_result: ParseResult<&str> = ParseResult::new(
///     Some("Parsed Successfully"),
///     vec![],
/// );
/// assert!(success_result.has_result());
/// let failure_result: ParseResult<&str> = ParseResult::new(
///     None,
///     vec![],
/// );
/// assert!(!failure_result.has_result());
/// ```
#[derive(Debug, Clone)]
pub struct ParseResult<'a, T> {
    /// The successful parse result, if any.
    result: Option<T>,
    /// A list of failures encountered during parsing.
    failures: Vec<ErrorDetails<'a>>,
}

impl<'a, T> From<ParseResult<'a, T>> for (Option<T>, Vec<ErrorDetails<'a>>) {
    fn from(pr: ParseResult<'a, T>) -> Self {
        (pr.result, pr.failures)
    }
}

impl<T> Display for ParseResult<'_, T> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match &self.result {
            Some(_) => write!(
                f,
                "ParseResult: Has result, Failures = {}",
                self.failures.len()
            ),
            None => write!(
                f,
                "ParseResult: No Result, Failures = {}",
                self.failures.len()
            ),
        }
    }
}

impl<'a, T> ParseResult<'a, T> {
    /// Checks if the parse result contains a successful result.
    ///
    /// # Returns
    /// * `true` if the result is `Some`, indicating a successful parse.
    /// * `false` if the result is `None`, indicating a failed parse.
    ///
    /// # Examples
    /// ```rust
    /// use vb6parse::parsers::parseresults::ParseResult;
    ///
    /// let success_result: ParseResult<&str> = ParseResult::new(
    ///     Some("Parsed Successfully"),
    ///     vec![],
    /// );
    /// assert!(success_result.has_result());
    ///
    /// let failure_result: ParseResult<&str> = ParseResult::new(
    ///     None,
    ///     vec![],
    /// );
    /// assert!(!failure_result.has_result());
    /// ```
    #[inline]
    pub const fn has_result(&self) -> bool {
        self.result.is_some()
    }

    /// Creates a new `ParseResult` instance.
    ///
    /// # Arguments
    ///
    /// * `result`: An optional successful parse result of type `T`.
    /// * `failures`: A vector of `ErrorDetails` representing failures encountered during parsing.
    ///
    /// # Returns
    ///
    /// * A new `ParseResult` instance containing the provided result and failures.
    ///
    pub fn new(result: Option<T>, failures: Vec<ErrorDetails<'a>>) -> Self {
        Self { result, failures }
    }

    /// Returns an iterator over the failures in the parse result.
    ///
    /// # Returns
    ///
    /// * An iterator over references to `ErrorDetails` instances representing the failures.
    ///
    /// # Examples
    /// ```rust
    ///
    /// use vb6parse::parsers::parseresults::ParseResult;
    /// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity, LexerError};
    ///
    /// let failure = ErrorDetails {
    ///     source_name: "test.bas".to_string().into_boxed_str(),
    ///     source_content: "Some source code",
    ///     error_offset: 5,
    ///     line_start: 0,
    ///     line_end: 10,
    ///     severity: Severity::Error,
    ///     labels: vec![],
    ///     notes: vec![],
    ///     kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
    /// };
    /// let parse_result: ParseResult<'_, &str> = ParseResult::new(
    ///     None,
    ///     vec![failure],
    /// );
    /// for error in parse_result.failures() {
    ///     error.print();
    /// }
    /// ```
    #[inline]
    pub fn failures(&self) -> impl Iterator<Item = &ErrorDetails<'a>> {
        self.failures.iter()
    }

    /// Consumes the parse result and returns an iterator over the failures.
    ///
    /// # Returns
    ///
    /// * An iterator over `ErrorDetails` instances representing the failures.
    ///
    /// # Examples
    /// ```rust
    ///
    /// use vb6parse::parsers::parseresults::ParseResult;
    /// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity, LexerError};
    ///
    /// let failure = ErrorDetails {
    ///     source_name: "test.bas".to_string().into_boxed_str(),
    ///     source_content: "Some source code",
    ///     error_offset: 5,
    ///     line_start: 0,
    ///     line_end: 10,
    ///     labels: vec![],
    ///     notes: vec![],
    ///     severity: Severity::Error,
    ///     kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
    /// };
    /// let parse_result: ParseResult<'_, &str> = ParseResult::new(
    ///     None,
    ///     vec![failure],
    /// );
    /// for error in parse_result.into_failures() {
    ///     error.print();
    /// }
    /// ```
    #[inline]
    pub fn into_failures(self) -> impl Iterator<Item = ErrorDetails<'a>> {
        self.failures.into_iter()
    }

    /// Checks if the parse result contains any failures.
    ///
    /// # Returns
    /// * `true` if there are one or more failures in the parse result.
    /// * `false` if there are no failures.
    ///
    /// # Examples
    /// ```rust
    ///
    /// use vb6parse::parsers::parseresults::ParseResult;
    /// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity, LexerError};
    ///
    /// let failure = ErrorDetails {
    ///     source_name: "test.bas".to_string().into_boxed_str(),
    ///     source_content: "Some source code",
    ///     error_offset: 5,
    ///     line_start: 0,
    ///     line_end: 10,
    ///     labels: vec![],
    ///     notes: vec![],
    ///     severity: Severity::Error,
    ///     kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
    /// };
    /// let failure_result: ParseResult<'_, &str> = ParseResult::new(
    ///     None,
    ///     vec![failure],
    /// );
    /// assert!(failure_result.has_failures());
    ///
    /// let success_result: ParseResult<'_, &str> = ParseResult::new(
    ///     Some("Parsed Successfully"),
    ///     vec![],
    /// );
    /// assert!(!success_result.has_failures());
    /// ```
    #[inline]
    pub const fn has_failures(&self) -> bool {
        !matches!(self.failures.len(), 0)
    }

    /// Adds a failure to the parse result's list of failures.
    ///
    /// # Arguments
    /// * `failure`: An `ErrorDetails` instance representing the failure to be added.
    ///
    /// # Examples
    /// ```rust
    ///
    /// use vb6parse::parsers::parseresults::ParseResult;
    /// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity, LexerError};
    ///
    /// let mut parse_result = ParseResult::new(
    ///     Some("Parsed Successfully"),
    ///     vec![],
    /// );
    /// let failure = ErrorDetails {
    ///     source_name: "test.bas".to_string().into_boxed_str(),
    ///     source_content: "Some source code",
    ///     error_offset: 5,
    ///     line_start: 0,
    ///     line_end: 10,
    ///     severity: Severity::Error,
    ///     notes: vec![],
    ///     labels: vec![],
    ///     kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
    /// };
    /// parse_result.push_failure(failure);
    /// assert!(parse_result.has_failures());
    /// ```
    #[inline]
    pub fn push_failure(&mut self, failure: ErrorDetails<'a>) {
        self.failures.push(failure);
    }

    /// Appends multiple failures to the parse result's list of failures.
    ///
    /// # Arguments
    /// * `failures`: A mutable reference to a vector of `ErrorDetails` instances representing the failures to be added.
    ///
    /// # Examples
    /// ```rust
    ///
    /// use vb6parse::parsers::parseresults::ParseResult;
    /// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity, LexerError};
    ///
    /// let mut parse_result = ParseResult::new(
    ///     Some("Parsed Successfully"),
    ///     vec![],
    /// );
    /// let mut failures = vec![
    ///     ErrorDetails {
    ///         source_name: "test.bas".to_string().into_boxed_str(),
    ///         source_content: "Some source code",
    ///         error_offset: 5,
    ///         line_start: 0,
    ///         line_end: 10,
    ///         labels: vec![],
    ///         notes: vec![],
    ///         severity: Severity::Error,
    ///         kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
    ///     },
    ///     ErrorDetails {
    ///         source_name: "test.bas".to_string().into_boxed_str(),
    ///         source_content: "Some more source code",
    ///         error_offset: 15,
    ///         line_start: 1,
    ///         line_end: 11,
    ///         labels: vec![],
    ///         notes: vec![],
    ///         severity: Severity::Error,
    ///         kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
    ///     },
    /// ];
    /// parse_result.append_failures(&mut failures);
    /// assert!(parse_result.has_failures());
    /// ```
    #[inline]
    pub fn append_failures(&mut self, failures: &mut Vec<ErrorDetails<'a>>) {
        self.failures.append(failures);
    }

    /// Unwraps the parse result, returning the successful result if it exists.
    ///
    /// # Panics
    /// Panics if the parse result does not contain a successful result (`None`).
    ///
    /// # Returns
    /// * The successful parse result of type `T`.
    ///
    /// # Examples
    /// ```rust
    /// use vb6parse::parsers::parseresults::ParseResult;
    ///
    /// let parse_result: ParseResult<&str> = ParseResult::new(
    ///     Some("Parsed Successfully"),
    ///     vec![],
    /// );
    ///
    /// let (result, failures) = parse_result.unpack();
    ///
    ///
    /// assert_eq!(result, Some("Parsed Successfully"));
    /// assert_eq!(failures.len(), 0);
    /// ```
    #[inline]
    pub fn unwrap(self) -> T {
        self.result
            .expect("Attempted to unwrap a ParseResult that did not have a result.")
    }

    /// Unpacks the parse result into its components.
    ///
    /// # Returns
    ///
    /// * A tuple containing:
    ///
    ///  - An `Option<T>` representing the successful parse result, if any.
    ///  - A `Vec<ErrorDetails<'a>>` containing the failures encountered during parsing.
    ///
    pub fn unpack(self) -> (Option<T>, Vec<ErrorDetails<'a>>) {
        (self.result, self.failures)
    }

    /// Unpacks the parse result into its components with errors and warnings separated.
    ///
    /// This is the Phase 2 method that separates errors from warnings.
    ///
    /// # Returns
    ///
    /// * A tuple containing:
    ///  - An `Option<T>` representing the successful parse result, if any.
    ///  - A `Vec<ErrorDetails<'a>>` containing the errors (`Severity::Error`).
    ///  - A `Vec<ErrorDetails<'a>>` containing the warnings (`Severity::Warning and Severity::Note`).
    ///
    pub fn unpack_with_severity(self) -> (Option<T>, Vec<ErrorDetails<'a>>, Vec<ErrorDetails<'a>>) {
        let diagnostics = Diagnostics::from_details(self.failures);
        (self.result, diagnostics.errors, diagnostics.warnings)
    }

    /// Returns the diagnostics (errors and warnings) from the parse result.
    ///
    /// This method categorizes all failures by severity into a Diagnostics struct.
    ///
    /// # Returns
    ///
    /// * A `Diagnostics` struct containing errors and warnings.
    ///
    pub fn diagnostics(&self) -> Diagnostics<'a> {
        Diagnostics::from_details(self.failures.clone())
    }

    /// Returns only the errors from the parse result (excludes warnings).
    ///
    /// # Returns
    ///
    /// * A vector of `ErrorDetails` containing only errors (`Severity::Error`).
    ///
    pub fn errors(&self) -> Vec<&ErrorDetails<'a>> {
        self.failures
            .iter()
            .filter(|f| f.severity == Severity::Error)
            .collect()
    }

    /// Returns only the warnings from the parse result (excludes errors).
    ///
    /// # Returns
    ///
    /// * A vector of `ErrorDetails` containing only warnings (`Severity::Warning` and `Severity::Note`).
    ///
    pub fn warnings(&self) -> Vec<&ErrorDetails<'a>> {
        self.failures
            .iter()
            .filter(|f| matches!(f.severity, Severity::Warning | Severity::Note))
            .collect()
    }

    /// Returns an iterator over all diagnostics, with errors first followed by warnings.
    ///
    /// # Returns
    ///
    /// * An iterator over references to `ErrorDetails` instances.
    ///
    pub fn all_diagnostics(&self) -> impl Iterator<Item = &ErrorDetails<'a>> {
        self.errors().into_iter().chain(self.warnings())
    }

    /// Unwraps the parse result, returning the successful result if it exists.
    /// If there are any failures, it prints them and panics.
    ///
    /// # Panics
    /// Panics if there are any failures in the parse result.
    ///
    /// # Returns
    /// * The successful parse result of type `T`.
    ///
    /// # Examples
    /// ```rust
    /// use vb6parse::parsers::parseresults::ParseResult;
    /// use vb6parse::errors::ErrorKind;
    ///
    /// let parse_result: ParseResult<&str> = ParseResult::new(
    ///     Some("Parsed Successfully"),
    ///     vec![],
    /// );
    /// let result = parse_result.unwrap_or_fail();
    /// assert_eq!(result, "Parsed Successfully");
    /// ```
    #[inline]
    pub fn unwrap_or_fail(self) -> T {
        if self.has_failures() {
            for failure in &self.failures {
                failure.eprint();
            }
            panic!(
                "Parsing had {} failure(s). See errors above.",
                self.failures.len()
            );
        }
        self.result
            .expect("Attempted to unwrap a ParseResult that did not have a result.")
    }

    /// Converts the parse result into a standard `Result` type.
    ///
    /// If there are any failures, it returns them as an `Err`. If there is a successful result
    /// and no failures, it returns the result as `Ok`.
    ///
    /// # Returns
    /// * `Ok(T)` if there is a successful result and no failures.
    /// * `Err(Vec<ErrorDetails<'a>>)` if there are any failures.
    ///
    /// # Errors
    ///
    /// * Returns a vector of `ErrorDetails` if there are any failures in the parse result.
    /// * If there are no failures but the result is `None`, it returns an empty vector of failures.
    ///
    ///
    /// # Examples
    /// ```rust
    ///
    /// use vb6parse::parsers::parseresults::ParseResult;
    /// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity, LexerError};
    ///
    /// let failure = ErrorDetails {
    ///     source_name: "test.bas".to_string().into_boxed_str(),
    ///     source_content: "Some source code",
    ///     error_offset: 5,
    ///     line_start: 0,
    ///     line_end: 10,
    ///     severity: Severity::Error,
    ///     labels: vec![],
    ///     notes: vec![],
    ///     kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
    /// };
    /// let parse_result: ParseResult<'_, &str> = ParseResult::new(
    ///     Some("Parsed Successfully"),
    ///     vec![failure],
    /// );
    /// match parse_result.ok_or_errors() {
    ///     Ok(result) => println!("Parsed result: {}", result),
    ///     Err(errors) => {
    ///         for error in errors {
    ///             error.print();
    ///         }
    ///     }
    /// }
    /// ```
    pub fn ok_or_errors(self) -> Result<T, Vec<ErrorDetails<'a>>> {
        if self.has_failures() {
            Err(self.failures)
        } else {
            self.result.ok_or(self.failures)
        }
    }

    /// Maps a function over the successful result, preserving failures.
    ///
    /// This combinator allows transforming the parsed value while keeping
    /// all diagnostic information intact.
    ///
    /// # Arguments
    ///
    /// * `f`: A function that transforms the successful result from type `T` to type `U`.
    ///
    /// # Returns
    ///
    /// * A new `ParseResult<U>` with the transformed result and the same failures.
    ///
    /// # Examples
    /// ```rust
    /// use vb6parse::parsers::parseresults::ParseResult;
    /// use vb6parse::errors::ErrorKind;
    ///
    /// let parse_result: ParseResult<i32> = ParseResult::new(
    ///     Some(42),
    ///     vec![],
    /// );
    ///
    /// let mapped = parse_result.map(|x| x * 2);
    /// assert_eq!(mapped.unwrap(), 84);
    /// ```
    pub fn map<U, F>(self, f: F) -> ParseResult<'a, U>
    where
        F: FnOnce(T) -> U,
    {
        ParseResult {
            result: self.result.map(f),
            failures: self.failures,
        }
    }

    /// Chains another parsing operation, combining failures.
    ///
    /// This combinator allows chaining multiple parsing operations while
    /// accumulating all diagnostic information.
    ///
    /// # Arguments
    ///
    /// * `f`: A function that takes the successful result and returns another `ParseResult`.
    ///
    /// # Returns
    ///
    /// * A new `ParseResult<U>` with the result from the chained operation
    ///   and all failures from both operations combined.
    ///
    /// # Examples
    /// ```rust
    /// use vb6parse::parsers::parseresults::ParseResult;
    /// use vb6parse::errors::ErrorKind;
    ///
    /// let parse_result: ParseResult<i32> = ParseResult::new(
    ///     Some(42),
    ///     vec![],
    /// );
    ///
    /// let chained = parse_result.and_then(|x| {
    ///     ParseResult::new(Some(x + 1), vec![])
    /// });
    /// assert_eq!(chained.unwrap(), 43);
    /// ```
    pub fn and_then<U, F>(self, f: F) -> ParseResult<'a, U>
    where
        F: FnOnce(T) -> ParseResult<'a, U>,
    {
        match self.result {
            Some(value) => {
                let mut next_result = f(value);
                next_result.failures.extend(self.failures);
                next_result
            }
            None => ParseResult {
                result: None,
                failures: self.failures,
            },
        }
    }

    /// Converts the `ParseResult` into a standard `Result`.
    ///
    /// This method provides a way to convert a `ParseResult` into Rust's
    /// standard `Result` type for interoperability.
    ///
    /// # Returns
    ///
    /// * `Ok(T)` if there is a successful result.
    /// * `Err(Diagnostics)` if there is no result or there are failures.
    ///
    /// # Examples
    /// ```rust
    /// use vb6parse::parsers::parseresults::ParseResult;
    /// use vb6parse::errors::ErrorKind;
    ///
    /// let parse_result: ParseResult<i32> = ParseResult::new(
    ///     Some(42),
    ///     vec![],
    /// );
    ///
    /// match parse_result.into_result() {
    ///     Ok(value) => assert_eq!(value, 42),
    ///     Err(_diagnostics) => panic!("Should not error"),
    /// }
    /// ```
    pub fn into_result(self) -> Result<T, Diagnostics<'a>> {
        if self.has_failures() || self.result.is_none() {
            Err(Diagnostics::from_details(self.failures))
        } else {
            Ok(self.result.unwrap())
        }
    }
}

impl<'a, T> From<(T, ErrorDetails<'a>)> for ParseResult<'a, T> {
    /// Converts a tuple of a successful parse result and a single failure into a `ParseResult`.
    ///
    /// # Arguments
    /// * `parse_pair`: A tuple containing the successful parse result of type `T` and an `ErrorDetails` instance.
    ///
    /// # Returns
    /// * A `ParseResult` instance containing the successful result and a vector with the single failure.
    ///
    /// # Examples
    /// ```rust
    ///
    /// use vb6parse::parsers::parseresults::ParseResult;
    /// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity , LexerError};
    ///
    /// let failure = ErrorDetails {
    ///     source_name: "test.bas".to_string().into_boxed_str(),
    ///     source_content: "Some source code",
    ///     error_offset: 5,
    ///     line_start: 0,
    ///     line_end: 10,
    ///     severity: Severity::Error,
    ///     labels: vec![],
    ///     notes: vec![],
    ///     kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
    /// };
    ///
    /// let parse_pair = ("Parsed Successfully", failure);
    /// let parse_result = ParseResult::from(parse_pair);
    ///
    /// assert!(parse_result.has_result());
    /// assert!(parse_result.has_failures());
    /// ```
    fn from(parse_pair: (T, ErrorDetails<'a>)) -> ParseResult<'a, T> {
        ParseResult {
            result: Some(parse_pair.0),
            failures: vec![parse_pair.1],
        }
    }
}

impl<'a, I, T> From<(I, Vec<ErrorDetails<'a>>)> for ParseResult<'a, Vec<T>>
where
    I: IntoIterator<Item = T>,
{
    /// Converts a tuple of an iterable collection and a vector of failures into a `ParseResult`.
    ///
    /// # Arguments
    /// * `parse_pair`: A tuple containing an iterable collection of type `I` and a vector of `ErrorDetails`.
    ///
    /// # Returns
    /// * A `ParseResult` instance containing the collected results and the provided failures.
    ///   If the collection is empty, the result will be `None`.
    ///
    /// # Examples
    /// ```rust
    ///
    /// use vb6parse::parsers::parseresults::ParseResult;
    /// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity, LexerError};
    ///
    /// let failures = vec![
    ///     ErrorDetails {
    ///         source_name: "test.bas".to_string().into_boxed_str(),
    ///         source_content: "Some source code",
    ///         error_offset: 5,
    ///         line_start: 0,
    ///         line_end: 10,
    ///         labels: vec![],
    ///         severity: Severity::Error,
    ///         notes: vec![],
    ///         kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
    ///     },
    ///     ErrorDetails {
    ///         source_name: "test.bas".to_string().into_boxed_str(),
    ///         source_content: "Some source code",
    ///         error_offset: 15,
    ///         line_start: 0,
    ///         line_end: 10,
    ///         labels: vec![],
    ///         severity: Severity::Error,
    ///         notes: vec![],
    ///         kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
    ///     },
    /// ];
    ///
    /// let parse_pair = (vec!["Item1", "Item2"], failures);
    /// let parse_result = ParseResult::from(parse_pair);
    ///
    /// assert!(parse_result.has_result());
    /// assert!(parse_result.has_failures());
    /// ```
    fn from(parse_pair: (I, Vec<ErrorDetails<'a>>)) -> ParseResult<'a, Vec<T>> {
        let collection: Vec<T> = parse_pair.0.into_iter().collect();
        if collection.is_empty() {
            return ParseResult {
                result: None,
                failures: parse_pair.1,
            };
        }

        ParseResult {
            result: Some(collection),
            failures: parse_pair.1,
        }
    }
}

impl<'a> From<(TokenStream<'a>, Vec<ErrorDetails<'a>>)> for ParseResult<'a, TokenStream<'a>> {
    /// Converts a tuple of a `TokenStream` and a vector of failures into a `ParseResult`.
    ///
    /// # Arguments
    /// * `parse_pair`: A tuple containing a `TokenStream` and a vector of `ErrorDetails`.
    ///
    /// # Returns
    /// * A `ParseResult` instance containing the `TokenStream` and the provided failures.
    ///
    /// # Examples
    /// ```rust
    ///
    /// use vb6parse::parsers::parseresults::ParseResult;
    /// use vb6parse::errors::{ErrorDetails, ErrorKind, Severity, LexerError};
    /// use vb6parse::lexer::TokenStream;
    ///
    /// let failures = vec![
    ///     ErrorDetails {
    ///         source_name: "test.bas".to_string().into_boxed_str(),
    ///         source_content: "Some source code",
    ///         error_offset: 5,
    ///         line_start: 0,
    ///         line_end: 10,
    ///         labels: vec![],
    ///         severity: Severity::Error,
    ///         notes: vec![],
    ///         kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
    ///     },
    ///     ErrorDetails {
    ///         source_name: "test.bas".to_string().into_boxed_str(),
    ///         source_content: "Some source code",
    ///         error_offset: 15,
    ///         line_start: 0,
    ///         line_end: 10,
    ///         labels: vec![],
    ///         severity: Severity::Error,
    ///         notes: vec![],
    ///         kind: Box::new(ErrorKind::Lexer(LexerError::UnknownToken { token: "???".to_string() })),
    ///     },
    /// ];
    ///
    /// let token_stream = TokenStream::new("test.bas".to_string(), vec![]);
    /// let parse_pair = (token_stream, failures);
    /// let parse_result: ParseResult<TokenStream> = ParseResult::from(parse_pair);
    ///
    /// assert!(parse_result.has_result());
    /// assert!(parse_result.has_failures());
    /// ```
    fn from(
        parse_pair: (TokenStream<'a>, Vec<ErrorDetails<'a>>),
    ) -> ParseResult<'a, TokenStream<'a>> {
        ParseResult {
            result: Some(parse_pair.0),
            failures: parse_pair.1,
        }
    }
}
