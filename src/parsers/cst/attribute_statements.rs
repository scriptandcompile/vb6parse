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
mod tests {
    use crate::*;

    #[test]
    fn parse_attribute_statement() {
        let source = "Attribute VB_Name = \"modTest\"\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../snapshots/parsers/cst/attribute_statements");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
