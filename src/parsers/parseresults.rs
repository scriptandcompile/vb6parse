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

use crate::errors::ErrorDetails;
use crate::tokenstream::TokenStream;

/// Result of a parsing operation, containing an optional result and a list of failures encountered during parsing.
/// The result is `Some` if parsing was successful, and `None` if it failed completely.
/// Failures are collected in a vector, allowing for partial successes with warnings.
///
/// # Type Parameters
/// * `'a`: Lifetime parameter for error details.
/// * `T`: The type of the successful parse result.
/// * `E`: The type of the error details.
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
/// let success_result: ParseResult<&str, &str> = ParseResult::new(
///     Some("Parsed Successfully"),
///     vec![],
/// );
/// assert!(success_result.has_result());
/// let failure_result: ParseResult<&str, &str> = ParseResult::new(
///     None,
///     vec![],
/// );
/// assert!(!failure_result.has_result());
/// ```
#[derive(Debug, Clone)]
pub struct ParseResult<'a, T, E>
where
    E: ToString + std::fmt::Debug,
{
    /// The successful parse result, if any.
    result: Option<T>,
    /// A list of failures encountered during parsing.
    failures: Vec<ErrorDetails<'a, E>>,
}

impl<'a, T, E> From<ParseResult<'a, T, E>> for (Option<T>, Vec<ErrorDetails<'a, E>>)
where
    E: ToString + std::fmt::Debug,
{
    fn from(pr: ParseResult<'a, T, E>) -> Self {
        (pr.result, pr.failures)
    }
}

impl<T, E> Display for ParseResult<'_, T, E>
where
    E: ToString + std::fmt::Debug,
{
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

impl<'a, T, E> ParseResult<'a, T, E>
where
    E: ToString + std::fmt::Debug,
{
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
    /// let success_result: ParseResult<&str, &str> = ParseResult::new(
    ///     Some("Parsed Successfully"),
    ///     vec![],
    /// );
    /// assert!(success_result.has_result());
    ///
    /// let failure_result: ParseResult<&str, &str> = ParseResult::new(
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
    pub fn new(result: Option<T>, failures: Vec<ErrorDetails<'a, E>>) -> Self {
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
    /// use vb6parse::errors::{ErrorDetails, CodeErrorKind};
    ///
    /// let failure = ErrorDetails {
    ///     source_name: "test.bas".to_string().into_boxed_str(),
    ///     source_content: "Some source code",
    ///     error_offset: 5,
    ///     line_start: 0,
    ///     line_end: 10,
    ///     kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
    /// };
    /// let parse_result: ParseResult<'_, &str, CodeErrorKind> = ParseResult::new(
    ///     None,
    ///     vec![failure],
    /// );
    /// for error in parse_result.failures() {
    ///     error.print();
    /// }
    /// ```
    #[inline]
    pub fn failures(&self) -> impl Iterator<Item = &ErrorDetails<'a, E>> {
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
    /// use vb6parse::errors::{ErrorDetails, CodeErrorKind};
    ///
    /// let failure = ErrorDetails {
    ///     source_name: "test.bas".to_string().into_boxed_str(),
    ///     source_content: "Some source code",
    ///     error_offset: 5,
    ///     line_start: 0,
    ///     line_end: 10,
    ///     kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
    /// };
    /// let parse_result: ParseResult<'_, &str, CodeErrorKind> = ParseResult::new(
    ///     None,
    ///     vec![failure],
    /// );
    /// for error in parse_result.into_failures() {
    ///     error.print();
    /// }
    /// ```
    #[inline]
    pub fn into_failures(self) -> impl Iterator<Item = ErrorDetails<'a, E>> {
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
    /// use vb6parse::errors::{ErrorDetails, CodeErrorKind};
    ///
    /// let failure = ErrorDetails {
    ///     source_name: "test.bas".to_string().into_boxed_str(),
    ///     source_content: "Some source code",
    ///     error_offset: 5,
    ///     line_start: 0,
    ///     line_end: 10,
    ///     kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
    /// };
    /// let failure_result: ParseResult<'_, &str, CodeErrorKind> = ParseResult::new(
    ///     None,
    ///     vec![failure],
    /// );
    /// assert!(failure_result.has_failures());
    ///
    /// let success_result: ParseResult<'_, &str, CodeErrorKind> = ParseResult::new(
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
    /// use vb6parse::errors::{ErrorDetails, CodeErrorKind};
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
    ///     kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
    /// };
    /// parse_result.push_failure(failure);
    /// assert!(parse_result.has_failures());
    /// ```
    #[inline]
    pub fn push_failure(&mut self, failure: ErrorDetails<'a, E>) {
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
    /// use vb6parse::errors::{ErrorDetails, CodeErrorKind};
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
    ///         kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
    ///     },
    ///     ErrorDetails {
    ///         source_name: "test.bas".to_string().into_boxed_str(),
    ///         source_content: "Some more source code",
    ///         error_offset: 15,
    ///         line_start: 1,
    ///         line_end: 11,
    ///         kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
    ///     },
    /// ];
    /// parse_result.append_failures(&mut failures);
    /// assert!(parse_result.has_failures());
    /// ```
    #[inline]
    pub fn append_failures(&mut self, failures: &mut Vec<ErrorDetails<'a, E>>) {
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
    /// let parse_result: ParseResult<&str, &str> = ParseResult::new(
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
    ///  - A `Vec<ErrorDetails<'a, E>>` containing the failures encountered during parsing.
    ///
    pub fn unpack(self) -> (Option<T>, Vec<ErrorDetails<'a, E>>) {
        (self.result, self.failures)
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
    /// use vb6parse::errors::CodeErrorKind;
    ///
    /// let parse_result: ParseResult<&str, &CodeErrorKind> = ParseResult::new(
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
    /// * `Err(Vec<ErrorDetails<'a, E>>)` if there are any failures.
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
    /// use vb6parse::errors::{ErrorDetails, CodeErrorKind};
    ///
    /// let failure = ErrorDetails {
    ///     source_name: "test.bas".to_string().into_boxed_str(),
    ///     source_content: "Some source code",
    ///     error_offset: 5,
    ///     line_start: 0,
    ///     line_end: 10,
    ///     kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
    /// };
    /// let parse_result: ParseResult<'_, &str, CodeErrorKind> = ParseResult::new(
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
    pub fn ok_or_errors(self) -> Result<T, Vec<ErrorDetails<'a, E>>> {
        if self.has_failures() {
            Err(self.failures)
        } else {
            self.result.ok_or(self.failures)
        }
    }
}

impl<'a, T, E> From<(T, ErrorDetails<'a, E>)> for ParseResult<'a, T, E>
where
    E: ToString + std::fmt::Debug,
{
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
    /// use vb6parse::errors::{ErrorDetails, CodeErrorKind };
    ///
    /// let failure = ErrorDetails {
    ///     source_name: "test.bas".to_string().into_boxed_str(),
    ///     source_content: "Some source code",
    ///     error_offset: 5,
    ///     line_start: 0,
    ///     line_end: 10,
    ///     kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
    /// };
    ///
    /// let parse_pair = ("Parsed Successfully", failure);
    /// let parse_result = ParseResult::from(parse_pair);
    ///
    /// assert!(parse_result.has_result());
    /// assert!(parse_result.has_failures());
    /// ```
    fn from(parse_pair: (T, ErrorDetails<'a, E>)) -> ParseResult<'a, T, E> {
        ParseResult {
            result: Some(parse_pair.0),
            failures: vec![parse_pair.1],
        }
    }
}

impl<'a, I, T, E> From<(I, Vec<ErrorDetails<'a, E>>)> for ParseResult<'a, Vec<T>, E>
where
    I: IntoIterator<Item = T>,
    E: ToString + std::fmt::Debug,
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
    /// use vb6parse::errors::{ErrorDetails, CodeErrorKind};
    ///
    /// let failures = vec![
    ///     ErrorDetails {
    ///         source_name: "test.bas".to_string().into_boxed_str(),
    ///         source_content: "Some source code",
    ///         error_offset: 5,
    ///         line_start: 0,
    ///         line_end: 10,
    ///         kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
    ///     },
    ///     ErrorDetails {
    ///         source_name: "test.bas".to_string().into_boxed_str(),
    ///         source_content: "Some source code",
    ///         error_offset: 15,
    ///         line_start: 0,
    ///         line_end: 10,
    ///         kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
    ///     },
    /// ];
    ///
    /// let parse_pair = (vec!["Item1", "Item2"], failures);
    /// let parse_result = ParseResult::from(parse_pair);
    ///
    /// assert!(parse_result.has_result());
    /// assert!(parse_result.has_failures());
    /// ```
    fn from(parse_pair: (I, Vec<ErrorDetails<'a, E>>)) -> ParseResult<'a, Vec<T>, E> {
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

impl<'a, E> From<(TokenStream<'a>, Vec<ErrorDetails<'a, E>>)>
    for ParseResult<'a, TokenStream<'a>, E>
where
    E: ToString + std::fmt::Debug,
{
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
    /// use vb6parse::errors::{ErrorDetails, CodeErrorKind};
    /// use vb6parse::tokenstream::TokenStream;
    ///
    /// let failures = vec![
    ///     ErrorDetails {
    ///         source_name: "test.bas".to_string().into_boxed_str(),
    ///         source_content: "Some source code",
    ///         error_offset: 5,
    ///         line_start: 0,
    ///         line_end: 10,
    ///         kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
    ///     },
    ///     ErrorDetails {
    ///         source_name: "test.bas".to_string().into_boxed_str(),
    ///         source_content: "Some source code",
    ///         error_offset: 15,
    ///         line_start: 0,
    ///         line_end: 10,
    ///         kind: CodeErrorKind::UnknownToken { token: "???".to_string() },
    ///     },
    /// ];
    ///
    /// let token_stream = TokenStream::new("test.bas".to_string(), vec![]);
    /// let parse_pair = (token_stream, failures);
    /// let parse_result: ParseResult<TokenStream, CodeErrorKind> = ParseResult::from(parse_pair);
    ///
    /// assert!(parse_result.has_result());
    /// assert!(parse_result.has_failures());
    /// ```
    fn from(
        parse_pair: (TokenStream<'a>, Vec<ErrorDetails<'a, E>>),
    ) -> ParseResult<'a, TokenStream<'a>, E> {
        ParseResult {
            result: Some(parse_pair.0),
            failures: parse_pair.1,
        }
    }
}
