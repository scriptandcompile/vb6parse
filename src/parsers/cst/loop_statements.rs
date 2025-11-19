//! Do/Loop statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Do loop statements:
//! - Do While...Loop
//! - Do Until...Loop
//! - Do...Loop While
//! - Do...Loop Until
//! - Do...Loop (infinite loop)

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Parse a Do...Loop statement.
    ///
    /// VB6 supports several forms of Do loops:
    /// - Do While condition...Loop
    /// - Do Until condition...Loop
    /// - Do...Loop While condition
    /// - Do...Loop Until condition
    /// - Do...Loop (infinite loop, requires Exit Do)
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/doloop-statement)
    pub(super) fn parse_do_statement(&mut self) {
        // if we are now parsing a do statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::DoStatement.to_raw());

        // Consume "Do" keyword
        self.consume_token();

        // Consume whitespace after Do
        self.consume_whitespace();

        // Check if we have While or Until after Do
        let has_top_condition =
            self.at_token(VB6Token::WhileKeyword) || self.at_token(VB6Token::UntilKeyword);

        if has_top_condition {
            // Consume While or Until
            self.consume_token();

            // Parse condition - consume everything until newline
            self.parse_conditional();
        }

        // Consume newline after Do line
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        // Parse the loop body until "Loop"
        self.parse_code_block(|parser| parser.at_token(VB6Token::LoopKeyword));

        // Consume "Loop" keyword
        if self.at_token(VB6Token::LoopKeyword) {
            self.consume_token();

            // Consume whitespace after Loop
            self.consume_whitespace();

            // Check if we have While or Until after Loop
            if self.at_token(VB6Token::WhileKeyword) || self.at_token(VB6Token::UntilKeyword) {
                // Consume While or Until
                self.consume_token();

                // Parse condition - consume everything until newline
                self.parse_conditional();
            }

            // Consume newline after Loop
            if self.at_token(VB6Token::Newline) {
                self.consume_token();
            }
        }

        self.builder.finish_node(); // DoStatement
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn do_while_loop() {
        let source = r#"
Sub Test()
    Do While x < 10
        x = x + 1
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DoStatement"));
        assert!(debug.contains("WhileKeyword"));
        assert!(debug.contains("LoopKeyword"));
    }

    #[test]
    fn do_until_loop() {
        let source = r#"
Sub Test()
    Do Until x >= 10
        x = x + 1
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DoStatement"));
        assert!(debug.contains("UntilKeyword"));
        assert!(debug.contains("LoopKeyword"));
    }

    #[test]
    fn do_loop_while() {
        let source = r#"
Sub Test()
    Do
        x = x + 1
    Loop While x < 10
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DoStatement"));
        assert!(debug.contains("WhileKeyword"));
        assert!(debug.contains("LoopKeyword"));
    }

    #[test]
    fn do_loop_until() {
        let source = r#"
Sub Test()
    Do
        x = x + 1
    Loop Until x >= 10
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DoStatement"));
        assert!(debug.contains("UntilKeyword"));
        assert!(debug.contains("LoopKeyword"));
    }

    #[test]
    fn do_loop_infinite() {
        let source = r#"
Sub Test()
    Do
        If x > 10 Then Exit Do
        x = x + 1
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DoStatement"));
        assert!(debug.contains("LoopKeyword"));
    }

    #[test]
    fn nested_do_loops() {
        let source = r#"
Sub Test()
    Do While i < 10
        Do While j < 5
            j = j + 1
        Loop
        i = i + 1
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        // Should have two DoStatements
        let do_count = debug.matches("DoStatement").count();
        assert_eq!(do_count, 2);
    }

    #[test]
    fn do_while_with_complex_condition() {
        let source = r#"
Sub Test()
    Do While x < 10 And y > 0
        x = x + 1
        y = y - 1
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DoStatement"));
        assert!(debug.contains("WhileKeyword"));
        assert!(debug.contains("LoopKeyword"));
    }

    #[test]
    fn do_loop_preserves_whitespace() {
        let source = r#"
Sub Test()
    Do  While  x < 10
        x = x + 1
    Loop  While  y > 0
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DoStatement"));
        // Check that whitespace is preserved
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn function_with_do_loop_ending_at_end_function() {
        let source = r#"Function Test()
Do
Loop
End Function
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        // Check if "End" appears as Unknown in the tree
        let tree_str = cst.debug_tree();
        assert!(
            !tree_str.contains("Unknown"),
            "Should not have any Unknown tokens\n{}",
            tree_str
        );
    }

    #[test]
    fn function_with_do_until_loop() {
        let source = r#"Function Test()
Do Until x = ""
  y = z
Loop
End Function
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        // Check if "End" appears as Unknown in the tree
        let tree_str = cst.debug_tree();
        assert!(
            !tree_str.contains("Unknown"),
            "Should not have any Unknown tokens\n{}",
            tree_str
        );
    }
}
