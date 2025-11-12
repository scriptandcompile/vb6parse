use crate::language::VB6Token;

/// A stream of tokens with positional tracking.
///
/// This structure holds a vector of tokens along with the source file name
/// and current position within the token stream for parsing and error reporting.
#[derive(Debug, PartialEq, Clone, Eq, serde::Serialize)]
pub struct TokenStream<'a> {
    /// The name of the source file these tokens came from
    pub source_file: String,
    /// The vector of tokens with their text content
    pub tokens: Vec<(&'a str, VB6Token)>,
    /// Current position/offset in the token stream
    pub offset: usize,
}

impl<'a> TokenStream<'a> {
    /// Creates a new TokenStream with the given source file name and tokens
    pub fn new(source_file: String, tokens: Vec<(&'a str, VB6Token)>) -> Self {
        Self {
            source_file,
            tokens,
            offset: 0,
        }
    }

    /// Returns the current token without advancing the position
    pub fn current(&self) -> Option<&(&'a str, VB6Token)> {
        self.tokens.get(self.offset)
    }

    /// Returns the current token and advances the position
    pub fn next(&mut self) -> Option<(&'a str, VB6Token)> {
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
    pub fn is_at_end(&self) -> bool {
        self.offset >= self.tokens.len()
    }

    /// Returns the number of tokens in the stream
    pub fn len(&self) -> usize {
        self.tokens.len()
    }

    /// Returns true if the token stream is empty
    pub fn is_empty(&self) -> bool {
        self.tokens.is_empty()
    }

    /// Resets the offset to the beginning of the stream
    pub fn reset(&mut self) {
        self.offset = 0;
    }
}

impl<'a> std::ops::Index<usize> for TokenStream<'a> {
    type Output = (&'a str, VB6Token);

    fn index(&self, index: usize) -> &Self::Output {
        &self.tokens[index]
    }
}

impl<'a> IntoIterator for TokenStream<'a> {
    type Item = (&'a str, VB6Token);
    type IntoIter = std::vec::IntoIter<(&'a str, VB6Token)>;

    fn into_iter(self) -> Self::IntoIter {
        self.tokens.into_iter()
    }
}