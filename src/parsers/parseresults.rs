use core::convert::From;
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
#[derive(Debug, Clone)]
pub struct ParseResult<'a, T, E> {
    pub result: Option<T>,
    pub failures: Vec<ErrorDetails<'a, E>>,
}

impl<'a, T, E> ParseResult<'a, T, E> {
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
    /// let success_result: ParseResult<&str, &str> = ParseResult {
    ///     result: Some("Parsed Successfully"),
    ///     failures: vec![],
    /// };
    /// assert!(success_result.has_result());
    ///
    /// let failure_result: ParseResult<&str, &str> = ParseResult {
    ///     result: None,
    ///     failures: vec![],
    /// };
    /// assert!(!failure_result.has_result());
    /// ```
    #[inline]
    pub const fn has_result(&self) -> bool {
        self.result.is_some()
    }

    /// Checks if the parse result contains any failures.
    ///
    /// # Returns
    /// * `true` if there are one or more failures in the parse result.
    /// * `false` if there are no failures.
    #[inline]
    pub const fn has_failures(&self) -> bool {
        !matches!(self.failures.len(), 0)
    }

    /// Adds a failure to the parse result's list of failures.
    ///
    /// # Arguments
    /// * `failure`: An `ErrorDetails` instance representing the failure to be added.
    #[inline]
    pub fn push_failure(&mut self, failure: ErrorDetails<'a, E>) {
        self.failures.push(failure);
    }

    /// Appends multiple failures to the parse result's list of failures.
    ///
    /// # Arguments
    /// * `failures`: A mutable reference to a vector of `ErrorDetails` instances representing the failures to be added.
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
    #[inline]
    pub fn unwrap(self) -> T {
        self.result
            .expect("Attempted to unwrap a ParseResult that did not have a result.")
    }
}

impl<'a, T, E> From<(T, ErrorDetails<'a, E>)> for ParseResult<'a, T, E> {
    /// Converts a tuple of a successful parse result and a single failure into a `ParseResult`.
    ///
    /// # Arguments
    /// * `parse_pair`: A tuple containing the successful parse result of type `T` and an `ErrorDetails` instance.
    ///
    /// # Returns
    /// * A `ParseResult` instance containing the successful result and a vector with the single failure.
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
{
    /// Converts a tuple of an iterable collection and a vector of failures into a `ParseResult`.
    ///
    /// # Arguments
    /// * `parse_pair`: A tuple containing an iterable collection of type `I` and a vector of `ErrorDetails`.
    ///
    /// # Returns
    /// * A `ParseResult` instance containing the collected results and the provided failures.
    ///   If the collection is empty, the result will be `None`.
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
{
    /// Converts a tuple of a `TokenStream` and a vector of failures into a `ParseResult`.
    ///
    /// # Arguments
    /// * `parse_pair`: A tuple containing a `TokenStream` and a vector of `ErrorDetails`.
    ///
    /// # Returns
    /// * A `ParseResult` instance containing the `TokenStream` and the provided failures.
    fn from(
        parse_pair: (TokenStream<'a>, Vec<ErrorDetails<'a, E>>),
    ) -> ParseResult<'a, TokenStream<'a>, E> {
        ParseResult {
            result: Some(parse_pair.0),
            failures: parse_pair.1,
        }
    }
}
