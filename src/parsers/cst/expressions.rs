//! Expression parsing for VB6 CST.
//!
//! This module handles parsing of various VB6 expressions:
//! - Conditional expressions (binary and unary)
//! - Assignment statements
//! - ElseIf and Else clauses

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse an ElseIf clause: ElseIf condition Then ...
    pub(super) fn parse_elseif_clause(&mut self) {
        self.builder.start_node(SyntaxKind::ElseIfClause.to_raw());

        // Consume "ElseIf" keyword
        self.consume_token();

        // Parse the conditional expression
        self.parse_conditional();

        // Consume "Then" if present
        if self.at_token(VB6Token::ThenKeyword) {
            self.consume_token();
        }

        // Consume any whitespace after Then
        self.consume_whitespace();

        // Consume the newline after Then
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse body until "End If", "Else", or another "ElseIf"
        self.parse_code_block(|parser| {
            parser.at_token(VB6Token::ElseIfKeyword)
                || parser.at_token(VB6Token::ElseKeyword)
                || (parser.at_token(VB6Token::EndKeyword)
                    && parser.peek_next_keyword() == Some(VB6Token::IfKeyword))
        });

        self.builder.finish_node(); // ElseIfClause
    }

    /// Parse an Else clause: Else ...
    pub(super) fn parse_else_clause(&mut self) {
        self.builder.start_node(SyntaxKind::ElseClause.to_raw());

        // Consume "Else" keyword
        self.consume_token();

        // Consume any whitespace after Else
        self.consume_whitespace();

        // Consume the newline after Else
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse body until "End If"
        self.parse_code_block(|parser| {
            parser.at_token(VB6Token::EndKeyword)
                && parser.peek_next_keyword() == Some(VB6Token::IfKeyword)
        });

        self.builder.finish_node(); // ElseClause
    }

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

    /// Parse an assignment statement.
    ///
    /// VB6 assignment statement syntax:
    /// - variableName = expression
    /// - object.property = expression
    /// - array(index) = expression
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/assignment-operator)
    pub(super) fn parse_assignment_statement(&mut self) {

        // Assignments can appear in both header and body, so we do not modify parsing_header here.

        self.builder.start_node(SyntaxKind::AssignmentStatement.to_raw());

        // Consume everything until newline or colon (for inline If statements)
        // This includes: variable/property, "=", expression
        while !self.is_at_end() 
            && !self.at_token(VB6Token::Newline) 
            && !self.at_token(VB6Token::ColonOperator) {
            self.consume_token();
        }

        // Consume the newline if present (but not colon - that's handled by caller)
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // AssignmentStatement
    }

    /// Check if the current position is at the start of an assignment statement.
    /// This looks ahead to see if there's an `=` operator (not part of a comparison).
    pub(super) fn is_at_assignment(&self) -> bool {
        // Look ahead through the tokens to find an = operator before a newline
        // We need to skip: identifiers, periods, parentheses, array indices, etc.
        // Note: In VB6, keywords can be used as property/member names (e.g., obj.Property = value)
        let mut last_was_period = false;
        
        for (_text, token) in self.tokens.iter().skip(self.pos) {
            match token {
                VB6Token::Newline | VB6Token::EndOfLineComment | VB6Token::RemComment => {
                    // Reached end of line without finding assignment
                    return false;
                }
                VB6Token::EqualityOperator => {
                    // Found an = operator - this is likely an assignment
                    return true;
                }
                VB6Token::PeriodOperator => {
                    last_was_period = true;
                    continue;
                }
                // Skip tokens that could appear in the left-hand side of an assignment
                VB6Token::Whitespace => {
                    continue;
                }
                VB6Token::Identifier
                | VB6Token::LeftParentheses
                | VB6Token::RightParentheses
                | VB6Token::Number
                | VB6Token::Comma => {
                    last_was_period = false;
                    continue;
                }
                // After a period, keywords can be property names, so skip them
                _ if last_was_period => {
                    last_was_period = false;
                    continue;
                }
                // If we hit a keyword or other operator (not after period), it's not an assignment
                _ => {
                    return false;
                }
            }
        }
        false
    }
}
