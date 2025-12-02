//! Helper and utility methods for the CST parser.
//!
//! This module contains various helper methods used throughout the parsing process,
//! including token peeking, consumption, and logical line detection.

use super::Parser;
use crate::language::VB6Token;
use crate::parsers::SyntaxKind;
use std::num::NonZeroUsize;

impl<'a> Parser<'a> {
    /// Check if we've reached the end of the token stream.
    pub(super) fn is_at_end(&self) -> bool {
        self.pos >= self.tokens.len()
    }

    /// Get the current token without advancing the position.
    pub(super) fn current_token(&self) -> Option<&VB6Token> {
        self.tokens.get(self.pos).map(|(_, token)| token)
    }

    /// Check if the current token matches the given token type.
    pub(super) fn at_token(&self, token: VB6Token) -> bool {
        self.current_token() == Some(&token)
    }

    /// Peek ahead to get the next keyword (non-whitespace token).
    pub(super) fn peek_next_keyword(&self) -> Option<VB6Token> {
        self.peek_next_count_keywords(NonZeroUsize::new(1).unwrap())
            .next()
    }

    /// Check if the current token is an identifier.
    pub(super) fn is_identifier(&self) -> bool {
        matches!(self.current_token(), Some(VB6Token::Identifier))
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
                VB6Token::IntegerLiteral
                    | VB6Token::LongLiteral
                    | VB6Token::SingleLiteral
                    | VB6Token::DoubleLiteral
                    | VB6Token::DecimalLiteral
            )
        )
    }

    /// Check if we're at the end of a logical line (newline that's NOT a line continuation)
    /// In VB6, `_` followed by zero or more whitespaces and then a newline means "continue on next line"
    ///
    /// This function primarily looks forward without backtracking, except when positioned
    /// directly at a newline (which requires checking backward for a preceding underscore):
    /// - Scans forward through whitespace  
    /// - If underscore found, checks forward for newline (no backtracking)
    /// - If newline found directly, checks backward for underscore (minimal backtracking)
    /// - Otherwise returns false (not at line end)
    pub(super) fn is_at_logical_line_end(&self) -> bool {
        let mut check_pos = self.pos;

        // Skip forward over any whitespace
        while let Some((_, token)) = self.tokens.get(check_pos) {
            match token {
                VB6Token::Whitespace => {
                    check_pos += 1;
                }
                VB6Token::Underscore => {
                    // Found underscore - check forward if it's followed by whitespace* + newline
                    let mut after_underscore = check_pos + 1;

                    // Skip whitespace after underscore
                    while let Some((_, ws_token)) = self.tokens.get(after_underscore) {
                        if *ws_token == VB6Token::Whitespace {
                            after_underscore += 1;
                        } else {
                            break;
                        }
                    }

                    // Check if we found a newline after the underscore (and optional whitespace)
                    if let Some((_, next_token)) = self.tokens.get(after_underscore) {
                        if *next_token == VB6Token::Newline {
                            // This is a line continuation (underscore + whitespace* + newline)
                            return false;
                        }
                    }

                    // Underscore not followed by newline - not a line continuation
                    return false;
                }
                VB6Token::Newline => {
                    // Found newline directly - must check backward for preceding underscore
                    // This is the only case requiring backtracking, when already positioned at/near newline
                    let mut back_pos = self.pos;

                    // Skip backward over whitespace
                    while back_pos > 0 {
                        back_pos -= 1;
                        if let Some((_, back_token)) = self.tokens.get(back_pos) {
                            match back_token {
                                VB6Token::Whitespace => {}
                                VB6Token::Underscore => return false, // Line continuation
                                _ => return true,                     // Logical line end
                            }
                        }
                    }

                    // Newline at start of file
                    return true;
                }
                _ => {
                    // Hit a non-whitespace, non-underscore, non-newline token
                    // Not at end of line
                    return false;
                }
            }
        }

        // End of file
        false
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
    ) -> impl Iterator<Item = VB6Token> + '_ {
        self.tokens
            .iter()
            .skip(self.pos + 1)
            .filter(|(_, token)| *token != VB6Token::Whitespace)
            .take(count.get())
            .map(|(_, token)| *token)
    }

    /// Peek ahead and get the next `count` tokens (including whitespace) from the current position.
    pub(super) fn peek_next_count_tokens(
        &self,
        count: NonZeroUsize,
    ) -> impl Iterator<Item = VB6Token> + '_ {
        self.tokens
            .iter()
            .skip(self.pos + 1)
            .take(count.get())
            .map(|(_, token)| *token)
    }

    /// Peek ahead to get the next token (including whitespace).
    pub(super) fn peek_next_token(&self) -> Option<VB6Token> {
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

    /// Check if the current token is a keyword or identifier followed by DollarSign.
    /// This pattern represents functions like Error$, Mid$, Len$, UCase$, LCase$.
    pub(super) fn at_keyword_dollar(&self) -> bool {
        // Check for specific keywords that have $ variants
        let is_dollar_keyword = matches!(
            self.current_token(),
            Some(
                VB6Token::ErrorKeyword
                    | VB6Token::LenKeyword
                    | VB6Token::MidKeyword
                    | VB6Token::StringKeyword
            )
        );

        if is_dollar_keyword {
            if let Some(VB6Token::DollarSign) = self.peek_next_token() {
                return true;
            }
        }

        // Check for Identifier (like "UCase", "LCase", "Left", etc.) + DollarSign
        if self.at_token(VB6Token::Identifier) {
            if let Some(VB6Token::DollarSign) = self.peek_next_token() {
                // Only merge if it's one of the known dollar functions
                if let Some((text, _)) = self.tokens.get(self.pos) {
                    let text_upper = text.to_uppercase();
                    if matches!(
                        text_upper.as_str(),
                        | "CHR"
                            | "CHRB"
                            | "CHRW"
                            | "COMMAND"
                            | "CURDIR"
                            | "DATE"
                            | "ENVIRON"
                            | "FORMAT"
                            | "HEX"
                            | "LCASE"
                            | "LEFT"
                            | "LEFTB"
                            | "LTRIM"
                            | "MIDB"
                            | "OCT"
                            | "RIGHT"
                            | "RIGHTB"
                            | "RTRIM"
                            | "SPACE"
                            | "STR"
                            | "STRING"
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

    /// Consume keyword/identifier + DollarSign as a merged Identifier token.
    /// This merges tokens like Error + $, Len + $, Mid + $, etc. into single identifiers.
    pub(super) fn consume_keyword_dollar_as_identifier(&mut self) {
        if self.at_keyword_dollar() {
            // Get the text of both tokens
            let first_text = self
                .tokens
                .get(self.pos)
                .map(|(text, _)| *text)
                .unwrap_or("");
            let dollar_text = self
                .tokens
                .get(self.pos + 1)
                .map(|(text, _)| *text)
                .unwrap_or("");

            // Create a combined text for the identifier
            let combined_text = format!("{}{}", first_text, dollar_text);

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
    /// - If the current token is ErrorKeyword followed by DollarSign, they are merged into "Error$"
    /// - If the current token is an Identifier (like Len, Mid, UCase, LCase) followed by DollarSign,
    ///   they are merged into a single identifier (e.g., "Len$", "Mid$", "UCase$", "LCase$")
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
    pub(super) fn consume_whitespace(&mut self) {
        while self.at_token(VB6Token::Whitespace) {
            self.consume_token();
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
    /// Special handling: keyword/identifier followed by DollarSign is merged into a single Identifier.
    ///
    /// # Arguments
    /// * `target` - The token to stop at (will not be consumed)
    pub(super) fn consume_until(&mut self, target: VB6Token) {
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
        if target == VB6Token::Newline && self.at_token(VB6Token::Newline) {
            // Look back to see if there was an underscore before this newline
            // We need to check if the last non-whitespace token was an underscore
            let mut check_pos = self.pos;
            while check_pos > 0 {
                check_pos -= 1;
                if let Some((_, token)) = self.tokens.get(check_pos) {
                    match token {
                        VB6Token::Whitespace => continue, // Skip whitespace
                        VB6Token::Underscore => {
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
    pub(super) fn consume_until_after(&mut self, target: VB6Token) {
        self.consume_until(target);
        if self.at_token(target) {
            self.consume_token();
        }
    }
}
