//! Statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 statements:
//! - Object manipulation (Call, Set, With)
//! - Array operations (ReDim)
//!
//! Note: Control flow statements (If, Do, For, Select Case, GoTo, Exit, Label)
//! are in the controlflow module.
//! Built-in system statements (AppActivate, Beep, ChDir, ChDrive) are in the
//! built_in_statements module.

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse a Call statement:
    ///
    /// \[ Call \] name \[ argumentlist \]
    ///
    /// The Call statement syntax has these parts:
    ///
    /// | Part        | Optional / Required | Description |
    /// |-------------|---------------------|-------------|
    /// | Call        | Optional            | Indicates that a procedure is being called. The Call keyword is optional; if omitted, the procedure name is used directly. |
    /// | name        | Required            | Name of the procedure to be called; follows standard variable naming conventions. |
    /// | argumentlist| Optional            | List of arguments to be passed to the procedure. Arguments are enclosed in parentheses and separated by commas. |
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/call-statement)
    pub(super) fn parse_call_statement(&mut self) {
        // if we are now parsing a call statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::CallStatement.to_raw());

        // Consume "Call" keyword
        self.consume_token();

        // Consume everything until newline (preserving all tokens)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // CallStatement
    }

    /// Parse a Set statement.
    ///
    /// VB6 Set statement syntax:
    /// - Set objectVar = [New] objectExpression
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/set-statement)
    pub(super) fn parse_set_statement(&mut self) {
        // if we are now parsing a set statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::SetStatement.to_raw());

        // Consume "Set" keyword
        self.consume_token();

        // Consume everything until newline
        // This includes: variable, "=", [New], object expression
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // SetStatement
    }

    /// Parse a With statement.
    ///
    /// VB6 With statement syntax:
    /// - With object
    ///     .Property1 = value1
    ///     .Property2 = value2
    ///   End With
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/with-statement)
    pub(super) fn parse_with_statement(&mut self) {
        // if we are now parsing a with statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::WithStatement.to_raw());

        // Consume "With" keyword
        self.consume_token();

        // Consume everything until newline (the object expression)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse the body until "End With"
        self.parse_code_block(|parser| {
            parser.at_token(VB6Token::EndKeyword)
                && parser.peek_next_keyword() == Some(VB6Token::WithKeyword)
        });

        // Consume "End With" and trailing tokens
        if self.at_token(VB6Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "With"
            self.consume_whitespace();

            // Consume "With"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until(VB6Token::Newline);
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // WithStatement
    }

    /// Parse a ReDim statement.
    ///
    /// VB6 ReDim statement syntax:
    /// - ReDim [Preserve] varname(subscripts) [As type] [, varname(subscripts) [As type]] ...
    ///
    /// Used at procedure level to reallocate storage space for dynamic array variables.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/redim-statement)
    pub(super) fn parse_redim_statement(&mut self) {
        // if we are now parsing a ReDim statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ReDimStatement.to_raw());

        // Consume "ReDim" keyword
        self.consume_token();

        // Consume everything until newline (Preserve, variable declarations, etc.)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // ReDimStatement
    }

    /// Check if the current token is a statement keyword that parse_statement can handle.
    pub(super) fn is_statement_keyword(&self) -> bool {
        matches!(
            self.current_token(),
            Some(VB6Token::CallKeyword)
                | Some(VB6Token::SetKeyword)
                | Some(VB6Token::WithKeyword)
                | Some(VB6Token::ReDimKeyword)
        )
    }

    /// This is a centralized statement dispatcher that handles all VB6 statement types
    /// defined in this module (object manipulation and array operations).
    pub(super) fn parse_statement(&mut self) {
        match self.current_token() {
            Some(VB6Token::CallKeyword) => {
                self.parse_call_statement();
            }
            Some(VB6Token::SetKeyword) => {
                self.parse_set_statement();
            }
            Some(VB6Token::WithKeyword) => {
                self.parse_with_statement();
            }
            Some(VB6Token::ReDimKeyword) => {
                self.parse_redim_statement();
            }
            _ => {}
        }
    }
}
