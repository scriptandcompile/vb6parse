//! Module defining the `TokenStream` structure for managing a stream of tokens
//! with positional tracking.
//! 
//! This module provides the `TokenStream` struct which holds a vector of tokens
//! along with the source file name and current position within the token stream.
//! It includes methods for navigating through the tokens, such as advancing,
//! backtracking, and checking the current token.
//! 
//! # Example
//! ```rust
//! use vb6parse::language::Token;
//! use vb6parse::tokenstream::TokenStream;
//! 
//! let tokens = vec![("Dim", Token::DimKeyword), (" ", Token::Whitespace), ("x", Token::Identifier)];
//! let mut stream = TokenStream::new("test.bas".to_string(), tokens);
//! 
//! assert_eq!(stream.current(), Some(&("Dim", Token::DimKeyword)));
//! stream.advance();
//! assert_eq!(stream.current(), Some(&(" ", Token::Whitespace)));
//! stream.backtrack();
//! assert_eq!(stream.current(), Some(&("Dim", Token::DimKeyword)));
//! ```
//! # See Also
//! - [`tokenize`](crate::tokenize::tokenize) for tokenizing source code into tokens.
//! - [`SourceStream`](crate::sourcestream::SourceStream) for low-level character stream handling.
//! - [`Token`] for the definition of tokens used in the stream.

use crate::language::Token;

/// A stream of tokens with positional tracking.
///
/// This structure holds a vector of tokens along with the source file name
/// and current position within the token stream for parsing and error reporting.
#[derive(Debug, PartialEq, Clone, Eq, serde::Serialize)]
pub struct TokenStream<'a> {
    /// The name of the source file these tokens came from
    pub source_file: String,
    /// The vector of tokens with their text content
    pub tokens: Vec<(&'a str, Token)>,
    /// Current position/offset in the token stream
    pub offset: usize,
}

impl<'a> TokenStream<'a> {
    /// Creates a new `TokenStream` with the given source file name and tokens
    #[must_use]
    pub fn new(source_file: String, tokens: Vec<(&'a str, Token)>) -> Self {
        Self {
            source_file,
            tokens,
            offset: 0,
        }
    }

    /// Returns the current token without advancing the position
    #[must_use]
    pub fn current(&self) -> Option<&(&'a str, Token)> {
        self.tokens.get(self.offset)
    }

    /// Returns the current token and advances the position
    pub fn next(&mut self) -> Option<(&'a str, Token)> {
        if self.offset < self.tokens.len() {
            let token = self.tokens[self.offset];
            self.offset += 1;
            Some(token)
        } else {
            None
        }
    }

    /// Advances the position by one without returning the token
    pub fn advance(&mut self) {
        if self.offset < self.tokens.len() {
            self.offset += 1;
        }
    }

    /// Moves the position back by one
    pub fn backtrack(&mut self) {
        if self.offset > 0 {
            self.offset -= 1;
        }
    }

    /// Returns true if we've reached the end of the token stream
    #[must_use]
    pub fn is_at_end(&self) -> bool {
        self.offset >= self.tokens.len()
    }

    /// Returns the number of tokens in the stream
    #[must_use]
    pub fn len(&self) -> usize {
        self.tokens.len()
    }

    /// Returns true if the token stream is empty
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.tokens.is_empty()
    }

    /// Resets the offset to the beginning of the stream
    pub fn reset(&mut self) {
        self.offset = 0;
    }
}

impl<'a> std::ops::Index<usize> for TokenStream<'a> {
    type Output = (&'a str, Token);

    fn index(&self, index: usize) -> &Self::Output {
        &self.tokens[index]
    }
}

impl<'a> IntoIterator for TokenStream<'a> {
    type Item = (&'a str, Token);
    type IntoIter = std::vec::IntoIter<(&'a str, Token)>;

    fn into_iter(self) -> Self::IntoIter {
        self.tokens.into_iter()
    }
}
