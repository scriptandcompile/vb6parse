//! Statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 statements:
//! - Object manipulation (Call, Set, With)
//! - Array operations (ReDim)
//! - System commands (AppActivate, Beep, ChDir, ChDrive)
//!
//! Note: Control flow statements (If, Do, For, Select Case, GoTo, Exit, Label)
//! are in the controlflow module.

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

    /// Parse an AppActivate statement.
    ///
    /// VB6 AppActivate statement syntax:
    /// - AppActivate title[, wait]
    ///
    /// Activates an application window.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/appactivate-statement)
    pub(super) fn parse_appactivate_statement(&mut self) {
        // if we are now parsing an AppActivate statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::AppActivateStatement.to_raw());

        // Consume "AppActivate" keyword
        self.consume_token();

        // Consume everything until newline (title and optional wait parameter)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // AppActivateStatement
    }

    /// Parse a Beep statement.
    ///
    /// VB6 Beep statement syntax:
    /// - Beep
    ///
    /// Sounds a tone through the computer's speaker.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/beep-statement)
    pub(super) fn parse_beep_statement(&mut self) {
        // if we are now parsing a beep statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::BeepStatement.to_raw());

        // Consume "Beep" keyword
        self.consume_token();

        // Consume any whitespace and comments until newline
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // BeepStatement
    }

    /// Parse a ChDir statement.
    ///
    /// VB6 ChDir statement syntax:
    /// - ChDir path
    ///
    /// Changes the current directory or folder.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chdir-statement)
    pub(super) fn parse_chdir_statement(&mut self) {
        // if we are now parsing a ChDir statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ChDirStatement.to_raw());

        // Consume "ChDir" keyword
        self.consume_token();

        // Consume everything until newline (the path parameter)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // ChDirStatement
    }

    /// Parse a ChDrive statement.
    ///
    /// VB6 ChDrive statement syntax:
    /// - ChDrive drive
    ///
    /// Changes the current drive.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chdrive-statement)
    pub(super) fn parse_chdrive_statement(&mut self) {
        // if we are now parsing a ChDrive statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ChDriveStatement.to_raw());

        // Consume "ChDrive" keyword
        self.consume_token();

        // Consume everything until newline (the drive parameter)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // ChDriveStatement
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
                | Some(VB6Token::AppActivateKeyword)
                | Some(VB6Token::BeepKeyword)
                | Some(VB6Token::ChDirKeyword)
                | Some(VB6Token::ChDriveKeyword)
                | Some(VB6Token::ReDimKeyword)
        )
    }

    /// This is a centralized statement dispatcher that handles all VB6 statement types
    /// defined in this module (non-control-flow statements).
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
            Some(VB6Token::AppActivateKeyword) => {
                self.parse_appactivate_statement();
            }
            Some(VB6Token::BeepKeyword) => {
                self.parse_beep_statement();
            }
            Some(VB6Token::ChDirKeyword) => {
                self.parse_chdir_statement();
            }
            Some(VB6Token::ChDriveKeyword) => {
                self.parse_chdrive_statement();
            }
            Some(VB6Token::ReDimKeyword) => {
                self.parse_redim_statement();
            }
            _ => {},
        }
    }
}
