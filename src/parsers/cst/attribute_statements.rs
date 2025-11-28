//! Attribute statement parsing for VB6 CST.
//!
//! This module handles parsing of Attribute statements like:
//! - `Attribute VB_Name = "modTest"`
//! - `Attribute VB_GlobalNameSpace = False`

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse an Attribute statement: `Attribute VB_Name = "value"`
    pub(super) fn parse_attribute_statement(&mut self) {
        self.builder
            .start_node(SyntaxKind::AttributeStatement.to_raw());

        // Consume "Attribute" keyword
        self.consume_token();

        // Consume everything until newline (preserving all tokens)
        self.consume_until_after(VB6Token::Newline);

        self.builder.finish_node(); // AttributeStatement
    }
}

#[cfg(test)]
mod test {
    use crate::parsers::{ConcreteSyntaxTree, SyntaxKind};

    #[test]
    fn parse_attribute_statement() {
        let source = "Attribute VB_Name = \"modTest\"\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1); // AttributeStatement
        assert_eq!(cst.text(), "Attribute VB_Name = \"modTest\"\n");

        // Use navigation methods
        assert!(cst.contains_kind(SyntaxKind::AttributeStatement));
        let attr_statements = cst.find_children_by_kind(SyntaxKind::AttributeStatement);
        assert_eq!(attr_statements.len(), 1);
        assert_eq!(attr_statements[0].text, "Attribute VB_Name = \"modTest\"\n");
    }
}
