//! Option statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Option statements:
//! - Option Explicit - Require explicit variable declarations
//! - Option Base - Set default lower bound for array subscripts

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse an Option statement: Option Explicit On/Off or Option Base 0/1
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

    /// Parse an Option Base statement.
    ///
    /// Sets the default lower bound for array subscripts.
    ///
    /// # Syntax
    ///
    /// | Clause | Description |
    /// |--------|-------------|
    /// | `Option Base 0` | Sets the default lower bound to 0 (default) |
    /// | `Option Base 1` | Sets the default lower bound to 1 |
    ///
    /// # Remarks
    ///
    /// The `Option Base` statement is used to set the default lower bound for array subscripts
    /// in a module. By default, VB6 uses 0 as the lower bound. Using `Option Base 1` changes
    /// this to 1 for all arrays that don't explicitly specify bounds.
    ///
    /// - Must be used at module level (before any procedures)
    /// - Only values 0 and 1 are allowed
    /// - Affects only arrays declared without explicit lower bounds
    /// - Does not affect arrays declared with explicit bounds (e.g., `Dim arr(5 To 10)`)
    ///
    /// # Examples
    ///
    /// ```vb6
    /// Option Base 1
    ///
    /// Sub Example()
    ///     Dim arr(10)  ' Lower bound is 1, upper bound is 10
    ///     arr(1) = "First element"
    /// End Sub
    /// ```
    ///
    /// # References
    ///
    /// [Microsoft Documentation](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-base-statement)
    pub(super) fn parse_option_base_statement(&mut self) {
        self.parse_option_statement();
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

    #[test]
    fn parse_option_base_0() {
        let source = "Option Base 0\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("OptionKeyword"));
        assert!(debug.contains("BaseKeyword"));
        assert_eq!(cst.text(), "Option Base 0\n");
    }

    #[test]
    fn parse_option_base_1() {
        let source = "Option Base 1\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("OptionKeyword"));
        assert!(debug.contains("BaseKeyword"));
        assert_eq!(cst.text(), "Option Base 1\n");
    }

    #[test]
    fn option_base_at_module_level() {
        let source = "Option Base 1\n\nSub Test()\nEnd Sub\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
        assert!(debug.contains("SubStatement"));
    }

    #[test]
    fn option_base_with_whitespace() {
        let source = "Option  Base  1\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
        assert_eq!(cst.text(), "Option  Base  1\n");
    }

    #[test]
    fn option_base_with_comment() {
        let source = "Option Base 0 ' Set default array base\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn option_base_preserves_whitespace() {
        let source = "Option Base 1  \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "Option Base 1  \n");
    }

    #[test]
    fn multiple_option_statements() {
        let source = "Option Explicit\nOption Base 1\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert_eq!(cst.child_count(), 2);
        assert!(debug.contains("ExplicitKeyword"));
        assert!(debug.contains("BaseKeyword"));
    }

    #[test]
    fn option_base_case_insensitive() {
        let source = "OPTION BASE 1\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
    }

    #[test]
    fn option_base_with_line_continuation() {
        let source = "Option _\nBase 1\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
    }

    #[test]
    fn option_base_before_declarations() {
        let source = "Option Base 1\nDim arr(10) As Integer\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
        assert!(debug.contains("DimStatement"));
    }

    #[test]
    fn option_base_in_module() {
        let source = r#"Attribute VB_Name = "Module1"
Option Base 1
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("AttributeStatement"));
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
    }

    #[test]
    fn option_base_0_default() {
        let source = "Option Base 0\nDim x(5) As Integer\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("OptionStatement"));
        assert!(debug.contains("BaseKeyword"));
    }
}
