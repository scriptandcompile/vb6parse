//! Select Case statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Select Case statements:
//! - Select Case statements with multiple Case clauses
//! - Case Else clauses
//! - Case expressions (values, ranges, Is comparisons)

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl Parser<'_> {
    /// Parse a Select Case statement.
    ///
    /// Syntax:
    ///   Select Case testexpression
    ///     Case expression1
    ///       statements1
    ///     Case expression2
    ///       statements2
    ///     Case Else
    ///       statementsElse
    ///   End Select
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/select-case-statement)
    pub(super) fn parse_select_case_statement(&mut self) {
        // if we are now parsing a select case statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::SelectCaseStatement.to_raw());

        // Consume "Select" keyword
        self.consume_token();

        // Consume any whitespace between "Select" and "Case"
        self.consume_whitespace();

        // Consume "Case" keyword
        if self.at_token(VB6Token::CaseKeyword) {
            self.consume_token();
        }

        self.consume_whitespace();

        // Parse the test expression
        self.parse_expression();

        // Consume newline
        self.consume_until_after(VB6Token::Newline);

        // Parse Case clauses until "End Select"
        while !self.is_at_end() {
            // Check for "End Select"
            if self.at_token(VB6Token::EndKeyword)
                && self.peek_next_keyword() == Some(VB6Token::SelectKeyword)
            {
                break;
            }

            // Check for "Case" keyword
            if self.at_token(VB6Token::CaseKeyword) {
                // Check if this is "Case Else"
                let is_case_else = self.peek_next_keyword() == Some(VB6Token::ElseKeyword);

                if is_case_else {
                    // Parse Case Else clause
                    self.builder.start_node(SyntaxKind::CaseElseClause.to_raw());

                    // Consume "Case"
                    self.consume_token();

                    // Consume any whitespace between "Case" and "Else"
                    self.consume_whitespace();

                    // Consume "Else"
                    if self.at_token(VB6Token::ElseKeyword) {
                        self.consume_token();
                    }

                    // Consume until newline
                    self.consume_until_after(VB6Token::Newline);

                    // Parse statements in Case Else until next Case or End Select
                    self.parse_code_block(|parser| {
                        (parser.at_token(VB6Token::CaseKeyword))
                            || (parser.at_token(VB6Token::EndKeyword)
                                && parser.peek_next_keyword() == Some(VB6Token::SelectKeyword))
                    });

                    self.builder.finish_node(); // CaseElseClause
                } else {
                    // Parse regular Case clause
                    self.builder.start_node(SyntaxKind::CaseClause.to_raw());

                    // Consume "Case"
                    self.consume_token();

                    // Consume the case expression(s) until newline
                    self.consume_until_after(VB6Token::Newline);

                    // Parse statements in Case until next Case or End Select
                    self.parse_code_block(|parser| {
                        (parser.at_token(VB6Token::CaseKeyword))
                            || (parser.at_token(VB6Token::EndKeyword)
                                && parser.peek_next_keyword() == Some(VB6Token::SelectKeyword))
                    });

                    self.builder.finish_node(); // CaseClause
                }
            } else {
                // Consume whitespace, newlines, and comments
                self.consume_token();
            }
        }

        // Consume "End Select" and trailing tokens
        if self.at_token(VB6Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Select"
            self.consume_whitespace();

            // Consume "Select"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until_after(VB6Token::Newline);
        }

        self.builder.finish_node(); // SelectCaseStatement
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn select_case_simple() {
        let source = r#"
Sub Test()
    Select Case x
        Case 1
            Debug.Print "One"
        Case 2
            Debug.Print "Two"
        Case 3
            Debug.Print "Three"
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SelectCaseStatement"));
        assert!(debug.contains("SelectKeyword"));
        assert!(debug.contains("CaseClause"));
    }

    #[test]
    fn select_case_with_case_else() {
        let source = r#"
Sub Test()
    Select Case value
        Case 1
            result = "one"
        Case 2
            result = "two"
        Case Else
            result = "other"
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SelectCaseStatement"));
        assert!(debug.contains("CaseClause"));
        assert!(debug.contains("CaseElseClause"));
    }

    #[test]
    fn select_case_multiple_values() {
        let source = r#"
Sub Test()
    Select Case dayOfWeek
        Case 1, 7
            Debug.Print "Weekend"
        Case 2, 3, 4, 5, 6
            Debug.Print "Weekday"
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SelectCaseStatement"));
        assert!(debug.contains("CaseClause"));
    }

    #[test]
    fn select_case_with_is() {
        let source = r#"
Sub Test()
    Select Case score
        Case Is >= 90
            grade = "A"
        Case Is >= 80
            grade = "B"
        Case Is >= 70
            grade = "C"
        Case Else
            grade = "F"
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SelectCaseStatement"));
        assert!(debug.contains("IsKeyword"));
        assert!(debug.contains("CaseElseClause"));
    }

    #[test]
    fn select_case_with_to() {
        let source = r#"
Sub Test()
    Select Case temperature
        Case 0 To 32
            status = "Freezing"
        Case 33 To 65
            status = "Cold"
        Case 66 To 85
            status = "Comfortable"
        Case 86 To 100
            status = "Hot"
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SelectCaseStatement"));
        assert!(debug.contains("ToKeyword"));
    }

    #[test]
    fn select_case_string_comparison() {
        let source = r#"
Sub Test()
    Select Case userInput
        Case "yes", "y", "YES"
            DoSomething
        Case "no", "n", "NO"
            DoSomethingElse
        Case Else
            ShowError
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SelectCaseStatement"));
        assert!(debug.contains("StringLiteral"));
    }

    #[test]
    fn select_case_nested() {
        let source = r#"
Sub Test()
    Select Case x
        Case 1
            Select Case y
                Case 10
                    result = 11
                Case 20
                    result = 21
            End Select
        Case 2
            result = 2
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("SelectCaseStatement").count();
        assert_eq!(count, 2, "Expected 2 Select Case statements (nested)");
    }

    #[test]
    fn select_case_with_loops() {
        let source = r#"
Sub Test()
    Select Case operation
        Case "add"
            For i = 1 To 10
                total = total + i
            Next i
        Case "multiply"
            For i = 1 To 10
                total = total * i
            Next i
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SelectCaseStatement"));
        assert!(debug.contains("ForStatement"));
    }

    #[test]
    fn select_case_with_if() {
        let source = r#"
Sub Test()
    Select Case category
        Case 1
            If value > 100 Then
                status = "high"
            Else
                status = "low"
            End If
        Case 2
            result = "category2"
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SelectCaseStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn select_case_empty_case() {
        let source = r#"
Sub Test()
    Select Case x
        Case 1
        Case 2
            DoSomething
        Case 3
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SelectCaseStatement"));
        assert!(debug.contains("CaseClause"));
    }

    #[test]
    fn select_case_module_level() {
        let source = r#"
Public Sub ModuleLevelTest()
    Select Case globalVar
        Case 1
            result = "One"
        Case 2
            result = "Two"
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SelectCaseStatement"));
    }

    #[test]
    fn select_case_with_function_call() {
        let source = r#"
Sub Test()
    Select Case GetValue()
        Case 1
            result = "one"
        Case 2
            result = "two"
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SelectCaseStatement"));
        assert!(debug.contains("GetValue"));
    }

    #[test]
    fn select_case_case_is_relational() {
        let source = r#"
Sub Test()
    Select Case age
        Case Is < 13
            category = "child"
        Case Is < 20
            category = "teen"
        Case Is < 65
            category = "adult"
        Case Else
            category = "senior"
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SelectCaseStatement"));
        assert!(debug.contains("IsKeyword"));
    }

    #[test]
    fn select_case_mixed_expressions() {
        let source = r#"
Sub Test()
    Select Case x
        Case 1 To 5, 10, 15 To 20
            result = "range"
        Case Is > 100
            result = "large"
        Case Else
            result = "other"
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SelectCaseStatement"));
        assert!(debug.contains("ToKeyword"));
        assert!(debug.contains("IsKeyword"));
    }

    #[test]
    fn select_case_preserves_whitespace() {
        let source = r#"
Sub Test()
    Select Case x
        Case 1
            Debug.Print "test"
    End Select
End Sub
"#;

        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SelectCaseStatement"));
        assert!(debug.contains("Whitespace"));
        assert!(debug.contains("Newline"));
    }
}
