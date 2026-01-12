//! Object manipulation statements for VB6 CST.
//!
//! This module handles parsing of VB6 statements that manipulate objects:
//! - `Call` - Explicitly call a procedure or method
//! - `RaiseEvent` - Fire a custom event
//! - `Set` - Assign object references
//! - `With` - Reference an object multiple times without repeating its name
//!
//! All parsers in this module construct concrete syntax tree nodes using the `Parser` type.
//!
//! # Module Organization
//!
//! The object manipulation statement parsers are organized into focused submodules:
//! - [`call`] - `Call` statements and procedure calls
//! - [`events`] - `RaiseEvent` statements for custom events
//! - [`set`] - `Set` statements for object reference assignment
//! - [`with_block`] - `With...End With` blocks for simplified object access
//!
//! Each submodule implements parser methods as `Parser` extensions and includes
//! comprehensive tests with snapshot-based verification.

pub(crate) mod call;
pub(crate) mod events;
pub(crate) mod set;
pub(crate) mod with_block;

use crate::language::Token;
use crate::parsers::cst::Parser;

impl Parser<'_> {
    /// Check if the current token is a statement keyword that `parse_statement` can handle.
    /// Checks both current position and next non-whitespace token.
    pub(crate) fn is_statement_keyword(&self) -> bool {
        let token = if self.at_token(Token::Whitespace) {
            self.peek_next_keyword()
        } else {
            self.current_token().copied()
        };

        matches!(
            token,
            Some(
                Token::CallKeyword
                    | Token::RaiseEventKeyword
                    | Token::SetKeyword
                    | Token::WithKeyword
            )
        )
    }

    /// This is a centralized statement dispatcher that handles all VB6 statement types
    /// defined in this module (object manipulation statements).
    pub(crate) fn parse_statement(&mut self) {
        let token = if self.at_token(Token::Whitespace) {
            self.peek_next_keyword()
        } else {
            self.current_token().copied()
        };

        match token {
            Some(Token::CallKeyword) => {
                self.parse_call_statement();
            }
            Some(Token::RaiseEventKeyword) => {
                self.parse_raiseevent_statement();
            }
            Some(Token::SetKeyword) => {
                self.parse_set_statement();
            }
            Some(Token::WithKeyword) => {
                self.parse_with_statement();
            }
            _ => {}
        }
    }
}
