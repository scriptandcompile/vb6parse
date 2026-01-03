//! For/Next and For Each/Next statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 For loop statements:
//! - For...Next loops with counter variables
//! - For Each...In...Next loops for collections
//! - Step clauses
//! - Nested loops

use crate::language::Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse a For...Next statement.
    ///
    /// VB6 For...Next loop syntax:
    /// - For counter = start To end [Step step]...Next [counter]
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/fornext-statement)
    pub(super) fn parse_for_statement(&mut self) {
        // if we are now parsing a for statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ForStatement.to_raw());

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume "For" keyword
        self.consume_token();

        // Parse counter variable (lvalue)
        self.parse_lvalue();

        self.consume_whitespace();

        // Consume "="
        if self.at_token(Token::EqualityOperator) {
            self.consume_token();
        }

        self.consume_whitespace();

        // Parse start value
        self.parse_expression();

        self.consume_whitespace();

        // Consume "To" keyword if present
        if self.at_token(Token::ToKeyword) {
            self.consume_token();

            self.consume_whitespace();

            // Parse end value
            self.parse_expression();

            self.consume_whitespace();

            // Consume "Step" keyword if present
            if self.at_token(Token::StepKeyword) {
                self.consume_token();

                self.consume_whitespace();

                // Parse step value
                self.parse_expression();
            }
        }

        // Consume newline after For line
        self.consume_until_after(Token::Newline);

        // Parse the loop body until "Next"
        self.parse_statement_list(|parser| parser.at_token(Token::NextKeyword));

        // Consume "Next" keyword
        if self.at_token(Token::NextKeyword) {
            self.consume_token();

            // Consume everything until newline (optional counter variable)
            self.consume_until_after(Token::Newline);
        }

        self.builder.finish_node(); // ForStatement
    }

    /// Parse a For Each...Next statement.
    ///
    /// VB6 For Each...Next loop syntax:
    /// - For Each element In collection...Next [element]
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/for-eachnext-statement)
    pub(super) fn parse_for_each_statement(&mut self) {
        // if we are now parsing a for each statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::ForEachStatement.to_raw());

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume "For" keyword
        self.consume_token();

        // Consume whitespace
        self.consume_whitespace();

        // Consume "Each" keyword
        if self.at_token(Token::EachKeyword) {
            self.consume_token();
        }

        // Consume everything until "In" or newline
        // This includes: element variable name and whitespace
        while !self.is_at_end()
            && !self.at_token(Token::InKeyword)
            && !self.at_token(Token::Newline)
        {
            self.consume_token();
        }

        // Consume "In" keyword if present
        if self.at_token(Token::InKeyword) {
            self.consume_token();

            // Consume everything until newline (the collection)
            self.consume_until(Token::Newline);
        }

        // Consume newline after For Each line
        self.consume_until_after(Token::Newline);

        // Parse the loop body until "Next"
        self.parse_statement_list(|parser| parser.at_token(Token::NextKeyword));

        // Consume "Next" keyword
        if self.at_token(Token::NextKeyword) {
            self.consume_token();

            // Consume everything until newline (optional element variable)
            self.consume_until_after(Token::Newline);
        }

        self.builder.finish_node(); // ForEachStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn simple_for_loop() {
        let source = r"
Sub TestSub()
    For i = 1 To 10
        Debug.Print i
    Next i
End Sub
";

        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ForStatement"));
        assert!(debug.contains("ForKeyword"));
        assert!(debug.contains("ToKeyword"));
        assert!(debug.contains("NextKeyword"));
    }

    #[test]
    fn for_loop_with_step() {
        let source = r"
Sub TestSub()
    For i = 1 To 100 Step 5
        Debug.Print i
    Next i
End Sub
";

        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ForStatement"));
        assert!(debug.contains("StepKeyword"));
    }

    #[test]
    fn for_loop_with_negative_step() {
        let source = r"
Sub TestSub()
    For i = 10 To 1 Step -1
        Debug.Print i
    Next i
End Sub
";

        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ForStatement"));
        assert!(debug.contains("StepKeyword"));
    }

    #[test]
    fn for_loop_without_counter_after_next() {
        let source = r"
Sub TestSub()
    For i = 1 To 10
        Debug.Print i
    Next
End Sub
";

        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ForStatement"));
        assert!(debug.contains("NextKeyword"));
    }

    #[test]
    fn nested_for_loops() {
        let source = r"
Sub TestSub()
    For i = 1 To 5
        For j = 1 To 5
            Debug.Print i * j
        Next j
    Next i
End Sub
";

        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        // Count occurrences of ForStatement - should have 2
        let for_count = debug.matches("ForStatement").count();
        assert_eq!(for_count, 2);
    }

    #[test]
    fn for_loop_with_function_calls() {
        let source = r"
Sub TestSub()
    For i = GetStart() To GetEnd() Step GetStep()
        Debug.Print i
    Next i
End Sub
";

        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ForStatement"));
        assert!(debug.contains("ToKeyword"));
        assert!(debug.contains("StepKeyword"));
    }

    #[test]
    fn for_loop_preserves_whitespace() {
        let source = r"
Sub TestSub()
    For   i   =   1   To   10   Step   2
        Debug.Print i
    Next   i
End Sub
";

        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ForStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn multiple_for_loops_in_sequence() {
        let source = r#"
Sub TestSub()
    For i = 1 To 5
        Debug.Print "First: " & i
    Next i
    
    For j = 10 To 20 Step 2
        Debug.Print "Second: " & j
    Next j
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        // Count occurrences of ForStatement - should have 2
        let for_count = debug.matches("ForStatement").count();
        assert_eq!(for_count, 2);
    }

    #[test]
    fn for_each_loop_simple() {
        let source = r"
Sub TestSub()
    For Each item In collection
        Debug.Print item
    Next item
End Sub
";

        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ForEachStatement"));
        assert!(debug.contains("ForKeyword"));
        assert!(debug.contains("EachKeyword"));
        assert!(debug.contains("InKeyword"));
        assert!(debug.contains("NextKeyword"));
    }

    #[test]
    fn for_each_loop_without_variable_after_next() {
        let source = r"
Sub TestSub()
    For Each element In myArray
        Debug.Print element
    Next
End Sub
";

        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ForEachStatement"));
        assert!(debug.contains("EachKeyword"));
        assert!(debug.contains("InKeyword"));
    }

    #[test]
    fn nested_for_and_for_each() {
        let source = r"
Sub TestSub()
    For i = 1 To 10
        For Each item In items(i)
            Debug.Print item
        Next item
    Next i
End Sub
";

        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ForStatement"));
        assert!(debug.contains("ForEachStatement"));
        // Should have 1 of each type
        let for_count = debug.matches("ForStatement").count();
        let for_each_count = debug.matches("ForEachStatement").count();
        assert_eq!(for_count, 1);
        assert_eq!(for_each_count, 1);
    }
}
