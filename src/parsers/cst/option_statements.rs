//! Option statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Option statements:
//! - Option Explicit - Require explicit variable declarations

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse an Option statement: Option Explicit On/Off
    pub(super) fn parse_option_statement(&mut self) {
        self.builder
            .start_node(SyntaxKind::OptionStatement.to_raw());

        // Consume "Option" keyword
        self.consume_token();

        // Consume everything until newline (preserving all tokens)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // OptionStatement
    }
}

#[cfg(test)]
mod test {
    use crate::parsers::{ConcreteSyntaxTree, SyntaxKind};

    #[test]
    fn parse_option_explicit_on() {
        let source = "Option Explicit On\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Option Explicit On\n");
    }

    #[test]
    fn parse_option_explicit_off() {
        let source = "Option Explicit Off\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Option Explicit Off\n");
    }

    #[test]
    fn parse_option_explicit() {
        let source = "Option Explicit\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);
        assert_eq!(cst.text(), "Option Explicit\n");
    }
}
