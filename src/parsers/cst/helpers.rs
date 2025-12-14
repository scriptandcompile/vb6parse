//! Helper and utility methods for the CST parser.
//!
//! This module contains various helper methods used throughout the parsing process,
//! including token peeking, consumption, and logical line detection.

use super::Parser;
use crate::language::Token;
use crate::parsers::SyntaxKind;
use std::num::NonZeroUsize;

impl Parser<'_> {
    /// Check if we've reached the end of the token stream.
    pub(super) fn is_at_end(&self) -> bool {
        self.pos >= self.tokens.len()
    }

    /// Get the current token without advancing the position.
    pub(super) fn current_token(&self) -> Option<&Token> {
        self.tokens.get(self.pos).map(|(_, token)| token)
    }

    /// Check if the current token matches the given token type.
    pub(super) fn at_token(&self, token: Token) -> bool {
        self.current_token() == Some(&token)
    }

    /// Peek ahead to get the next keyword (non-whitespace token).
    pub(super) fn peek_next_keyword(&self) -> Option<Token> {
        self.peek_next_count_keywords(NonZeroUsize::new(1).unwrap())
            .next()
    }

    /// Check if the current token is an identifier.
    pub(super) fn is_identifier(&self) -> bool {
        matches!(self.current_token(), Some(Token::Identifier))
    }

    /// Check if the current token is a keyword.
    pub(super) fn at_keyword(&self) -> bool {
        match self.current_token() {
            Some(token) => token.is_keyword(),
            None => false,
        }
    }

    /// Check if the current token is a number (any numeric literal type).
    pub(super) fn is_number(&self) -> bool {
        matches!(
            self.current_token(),
            Some(
                Token::IntegerLiteral
                    | Token::LongLiteral
                    | Token::SingleLiteral
                    | Token::DoubleLiteral
                    | Token::DecimalLiteral
            )
        )
    }

    /// Peek ahead and get the next `count` non-whitespace keywords from the current position.
    ///
    /// # Arguments
    /// * `count` - Number of keywords to peek ahead (must be non-zero)
    ///
    /// # Returns
    /// An iterator over the next `count` keywords (non-whitespace tokens)
    pub(super) fn peek_next_count_keywords(
        &self,
        count: NonZeroUsize,
    ) -> impl Iterator<Item = Token> + '_ {
        self.tokens
            .iter()
            .skip(self.pos + 1)
            .filter(|(_, token)| *token != Token::Whitespace)
            .take(count.get())
            .map(|(_, token)| *token)
    }

    /// Peek ahead and get the next `count` tokens (including whitespace) from the current position.
    pub(super) fn peek_next_count_tokens(
        &self,
        count: NonZeroUsize,
    ) -> impl Iterator<Item = Token> + '_ {
        self.tokens
            .iter()
            .skip(self.pos + 1)
            .take(count.get())
            .map(|(_, token)| *token)
    }

    /// Peek ahead to get the next token (including whitespace).
    pub(super) fn peek_next_token(&self) -> Option<Token> {
        self.peek_next_count_tokens(NonZeroUsize::new(1).unwrap())
            .next()
    }

    /// Consume the current token and advance to the next position.
    pub(super) fn consume_token(&mut self) {
        if let Some((text, token)) = self.tokens.get(self.pos) {
            let kind = SyntaxKind::from(*token);
            self.builder.token(kind.to_raw(), text);
            self.pos += 1;
        }
    }

    /// Check if the current token is a keyword or identifier followed by `DollarSign`.
    /// This pattern represents functions like `Error$`, `Mid$`, `Len$`, `UCase$`, `LCase$`.
    pub(super) fn at_keyword_dollar(&self) -> bool {
        // Check for specific keywords that have $ variants
        let is_dollar_keyword = matches!(
            self.current_token(),
            Some(
                Token::ErrorKeyword
                    | Token::LenKeyword
                    | Token::MidKeyword
                    | Token::MidBKeyword
                    | Token::DateKeyword
                    | Token::StringKeyword
            )
        );

        if is_dollar_keyword {
            if let Some(Token::DollarSign) = self.peek_next_token() {
                return true;
            }
        }

        // Check for Identifier (like "UCase", "LCase", "Left", etc.) + DollarSign
        if self.at_token(Token::Identifier) {
            if let Some(Token::DollarSign) = self.peek_next_token() {
                // Only merge if it's one of the known dollar functions
                if let Some((text, _)) = self.tokens.get(self.pos) {
                    let text_upper = text.to_uppercase();
                    if matches!(
                        text_upper.as_str(),
                        "CHR"
                            | "CHRB"
                            | "CHRW"
                            | "COMMAND"
                            | "CURDIR"
                            | "DATE"
                            | "ENVIRON"
                            | "ERROR"
                            | "FORMAT"
                            | "HEX"
                            | "LCASE"
                            | "LEFT"
                            | "LEFTB"
                            | "LTRIM"
                            | "MID"
                            | "MIDB"
                            | "OCT"
                            | "RIGHT"
                            | "RIGHTB"
                            | "RTRIM"
                            | "SPACE"
                            | "STR"
                            | "TIME"
                            | "TRIM"
                            | "UCASE"
                    ) {
                        return true;
                    }
                }
            }
        }

        false
    }

    /// Consume keyword/identifier + `DollarSign` as a merged Identifier token.
    /// This merges tokens like `Error` + `$`, `Len` + `$`, `Mid` + `$`, etc. into single identifiers.
    pub(super) fn consume_keyword_dollar_as_identifier(&mut self) {
        if self.at_keyword_dollar() {
            // Get the text of both tokens
            let first_text = self.tokens.get(self.pos).map_or("", |(text, _)| *text);
            let dollar_text = self.tokens.get(self.pos + 1).map_or("", |(text, _)| *text);

            // Create a combined text for the identifier
            let combined_text = format!("{first_text}{dollar_text}");

            // Add as a single Identifier token
            self.builder
                .token(SyntaxKind::Identifier.to_raw(), &combined_text);

            // Skip both tokens
            self.pos += 2;
        }
    }

    /// Consume the current token as an Identifier, regardless of whether it's actually a keyword.
    /// This is used when keywords appear in identifier positions (e.g., variable names, property names).
    ///
    /// Special cases:
    /// - If the current token is `ErrorKeyword` followed by `DollarSign`, they are merged into "Error$"
    /// - If the current token is an Identifier (like `Len`, `Mid`, `UCase`, `LCase`) followed by `DollarSign`,
    ///   they are merged into a single identifier (e.g., `Len$`, `Mid$`, `UCase$`, `LCase$`)
    pub(super) fn consume_token_as_identifier(&mut self) {
        // Check for keyword/identifier + $ special cases
        if self.at_keyword_dollar() {
            self.consume_keyword_dollar_as_identifier();
            return;
        }

        if let Some((text, _)) = self.tokens.get(self.pos) {
            self.builder.token(SyntaxKind::Identifier.to_raw(), text);
            self.pos += 1;
        }
    }

    /// Consume all whitespace tokens at the current position.
    /// Also consumes line continuations (underscore followed by newline).
    pub(super) fn consume_whitespace(&mut self) {
        loop {
            if self.at_token(Token::Whitespace) {
                self.consume_token();
            } else if self.at_token(Token::Underscore) {
                // Check for line continuation: Underscore [Whitespace] Newline
                let mut lookahead = 1;
                let mut is_continuation = false;

                // Skip whitespace after underscore
                while let Some((_, token)) = self.tokens.get(self.pos + lookahead) {
                    if *token == Token::Whitespace {
                        lookahead += 1;
                    } else if *token == Token::Newline {
                        is_continuation = true;
                        break;
                    } else {
                        break;
                    }
                }

                if is_continuation {
                    // Consume Underscore
                    self.consume_token();
                    // Consume whitespace and Newline
                    while !self.at_token(Token::Newline) {
                        self.consume_token();
                    }
                    self.consume_token(); // Consume Newline
                } else {
                    break;
                }
            } else {
                break;
            }
        }
    }

    /// Consume the current token as an Unknown token.
    pub(super) fn consume_token_as_unknown(&mut self) {
        if let Some((text, _)) = self.tokens.get(self.pos) {
            self.builder.token(SyntaxKind::Unknown.to_raw(), text);
            self.pos += 1;
        }
    }

    /// Consume tokens until reaching the specified token or the end of input.
    /// The specified token is NOT consumed.
    ///
    /// Handles line continuations when consuming until a newline.
    /// Special handling: keyword/identifier followed by `DollarSign` is merged into a single Identifier.
    ///
    /// # Arguments
    /// * `target` - The token to stop at (will not be consumed)
    pub(super) fn consume_until(&mut self, target: Token) {
        while !self.is_at_end() && !self.at_token(target) {
            // Check for keyword/identifier + $ pattern and merge it
            if self.at_keyword_dollar() {
                self.consume_keyword_dollar_as_identifier();
            } else {
                self.consume_token();
            }
        }

        // If we're looking for a newline and we found one, check for line continuation
        // In VB6, underscore followed by whitespace and newline means "continue on next line"
        if target == Token::Newline && self.at_token(Token::Newline) {
            // Look back to see if there was an underscore before this newline
            // We need to check if the last non-whitespace token was an underscore
            let mut check_pos = self.pos;
            while check_pos > 0 {
                check_pos -= 1;
                if let Some((_, token)) = self.tokens.get(check_pos) {
                    match token {
                        Token::Whitespace => continue, // Skip whitespace
                        Token::Underscore => {
                            // Found line continuation! Consume the newline and keep going
                            self.consume_token(); // Consume the newline
                                                  // Continue consuming until we find a newline without continuation
                            self.consume_until(target);
                            return;
                        }
                        _ => break, // Not a continuation
                    }
                }
                break;
            }
        }
    }

    /// Consume tokens until reaching the specified token, then consume that token as well.
    ///
    /// This is a convenience method that combines `consume_until` with consuming the target token.
    /// Handles line continuations when consuming until a newline.
    ///
    /// # Arguments
    /// * `target` - The token to stop at and consume
    pub(super) fn consume_until_after(&mut self, target: Token) {
        self.consume_until(target);
        if self.at_token(target) {
            self.consume_token();
        }
    }
}
