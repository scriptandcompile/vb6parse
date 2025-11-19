//! Expression parsing for VB6 CST.
//!
//! This module handles parsing of various VB6 expressions:
//! - Conditional expressions (binary and unary)
//!
//! Note: ElseIf and Else clauses are in the if_controlflow module.
//! Note: Assignment statements are in the assignment module.

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse a conditional expression.
    ///
    /// This handles both:
    /// - Binary conditionals: `a = b`, `x > 5`, `name <> ""`
    /// - Unary conditionals: `Not condition`, `Not IsEmpty(x)`
    ///
    /// The conditional is parsed until "Then" or newline is encountered.
    pub(super) fn parse_conditional(&mut self) {
        // Skip any leading whitespace
        self.consume_whitespace();

        // Check if this is a unary conditional starting with "Not"
        if self.at_token(VB6Token::NotKeyword) {
            self.builder
                .start_node(SyntaxKind::UnaryConditional.to_raw());

            // Consume "Not" keyword
            self.consume_token();

            // Consume any whitespace after "Not"
            self.consume_whitespace();

            // Consume the rest of the conditional expression until "Then" or newline
            while !self.is_at_end()
                && !self.at_token(VB6Token::ThenKeyword)
                && !self.is_at_logical_line_end()
            {
                self.consume_token();
            }

            self.builder.finish_node(); // UnaryConditional
        } else {
            // Binary conditional - parse left side, operator, right side
            self.builder
                .start_node(SyntaxKind::BinaryConditional.to_raw());

            // Consume tokens until we hit a comparison operator
            while !self.is_at_end()
                && !self.at_token(VB6Token::ThenKeyword)
                && !self.is_at_logical_line_end()
            {
                // Check if we've hit a comparison operator
                if self.is_comparison_operator() {
                    // Consume the operator
                    self.consume_token();

                    // Consume any whitespace after the operator
                    self.consume_whitespace();

                    // Now consume the right side until "Then" or newline
                    while !self.is_at_end()
                        && !self.at_token(VB6Token::ThenKeyword)
                        && !self.is_at_logical_line_end()
                    {
                        self.consume_token();
                    }
                    break;
                }

                self.consume_token();
            }

            // If we didn't find an operator, we still consumed everything until "Then"
            // This handles cases like function calls that return boolean values

            self.builder.finish_node(); // BinaryConditional
        }
    }

    /// Check if the current token is a comparison operator
    pub(super) fn is_comparison_operator(&self) -> bool {
        matches!(
            self.current_token(),
            Some(VB6Token::EqualityOperator)
                | Some(VB6Token::LessThanOperator)
                | Some(VB6Token::GreaterThanOperator)
        )
    }
}
