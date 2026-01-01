//! Attribute statement parsing for VB6 CST.
//!
//! This module handles parsing of Attribute statements like:
//! - `Attribute VB_Name = "modTest"`
//! - `Attribute VB_GlobalNameSpace = False`

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse an Attribute statement: `Attribute VB_Name = "value"`
    pub(super) fn parse_attribute_statement(&mut self) {
        self.builder
            .start_node(SyntaxKind::AttributeStatement.to_raw());

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume "Attribute" keyword
        self.consume_token();

        // Consume everything until newline (preserving all tokens)
        self.consume_until_after(Token::Newline);

        self.builder.finish_node(); // AttributeStatement
    }
}

#[cfg(test)]
mod test {
    use crate::parsers::{ConcreteSyntaxTree, SyntaxKind};

    #[test]
    fn parse_attribute_statement() {
        let source = "Attribute VB_Name = \"modTest\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1); // AttributeStatement
        assert_eq!(cst.text(), "Attribute VB_Name = \"modTest\"\n");

        // Use navigation methods
        assert!(cst.contains_kind(SyntaxKind::AttributeStatement));
        let attr_statements: Vec<_> = cst
            .children_by_kind(SyntaxKind::AttributeStatement)
            .collect();
        assert_eq!(attr_statements.len(), 1);
        assert_eq!(
            attr_statements[0].text(),
            "Attribute VB_Name = \"modTest\"\n"
        );
    }
}
