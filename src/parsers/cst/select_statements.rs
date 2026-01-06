//! Select Case statement parsing for VB6 CST.
//!
//! This module handles parsing of VB6 Select Case statements:
//! - Select Case statements with multiple Case clauses
//! - Case Else clauses
//! - Case expressions (values, ranges, Is comparisons)

use crate::language::Token;
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

        // Consume any leading whitespace
        self.consume_whitespace();

        // Consume "Select" keyword
        self.consume_token();

        // Consume any whitespace between "Select" and "Case"
        self.consume_whitespace();

        // Consume "Case" keyword
        if self.at_token(Token::CaseKeyword) {
            self.consume_token();
        }

        self.consume_whitespace();

        // Parse the test expression
        self.parse_expression();

        // Consume newline
        self.consume_until_after(Token::Newline);

        // Parse Case clauses until "End Select"
        while !self.is_at_end() {
            // Check for "End Select"
            if self.at_token(Token::EndKeyword)
                && self.peek_next_keyword() == Some(Token::SelectKeyword)
            {
                break;
            }

            // Check for "Case" keyword
            if self.at_token(Token::CaseKeyword) {
                // Check if this is "Case Else"
                let is_case_else = self.peek_next_keyword() == Some(Token::ElseKeyword);

                if is_case_else {
                    // Parse Case Else clause
                    self.builder.start_node(SyntaxKind::CaseElseClause.to_raw());

                    // Consume "Case"
                    self.consume_token();

                    // Consume any whitespace between "Case" and "Else"
                    self.consume_whitespace();

                    // Consume "Else"
                    if self.at_token(Token::ElseKeyword) {
                        self.consume_token();
                    }

                    // Consume until newline
                    self.consume_until_after(Token::Newline);

                    // Parse statements in Case Else until next Case or End Select
                    self.parse_statement_list(|parser| {
                        (parser.at_token(Token::CaseKeyword))
                            || (parser.at_token(Token::EndKeyword)
                                && parser.peek_next_keyword() == Some(Token::SelectKeyword))
                    });

                    self.builder.finish_node(); // CaseElseClause
                } else {
                    // Parse regular Case clause
                    self.builder.start_node(SyntaxKind::CaseClause.to_raw());

                    // Consume "Case"
                    self.consume_token();

                    // Consume the case expression(s) until newline
                    self.consume_until_after(Token::Newline);

                    // Parse statements in Case until next Case or End Select
                    self.parse_statement_list(|parser| {
                        (parser.at_token(Token::CaseKeyword))
                            || (parser.at_token(Token::EndKeyword)
                                && parser.peek_next_keyword() == Some(Token::SelectKeyword))
                    });

                    self.builder.finish_node(); // CaseClause
                }
            } else {
                // Consume whitespace, newlines, and comments
                self.consume_token();
            }
        }

        // Consume "End Select" and trailing tokens
        if self.at_token(Token::EndKeyword) {
            // Consume "End"
            self.consume_token();

            // Consume any whitespace between "End" and "Select"
            self.consume_whitespace();

            // Consume "Select"
            self.consume_token();

            // Consume until newline (including it)
            self.consume_until_after(Token::Newline);
        }

        self.builder.finish_node(); // SelectCaseStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
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

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"One\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("2"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Two\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("3"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Three\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("value"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"one\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("2"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"two\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseElseClause {
                            CaseKeyword,
                            Whitespace,
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"other\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("dayOfWeek"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Comma,
                            Whitespace,
                            IntegerLiteral ("7"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Weekend\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("2"),
                            Comma,
                            Whitespace,
                            IntegerLiteral ("3"),
                            Comma,
                            Whitespace,
                            IntegerLiteral ("4"),
                            Comma,
                            Whitespace,
                            IntegerLiteral ("5"),
                            Comma,
                            Whitespace,
                            IntegerLiteral ("6"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Weekday\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("score"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IsKeyword,
                            Whitespace,
                            GreaterThanOrEqualOperator,
                            Whitespace,
                            IntegerLiteral ("90"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("grade"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"A\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IsKeyword,
                            Whitespace,
                            GreaterThanOrEqualOperator,
                            Whitespace,
                            IntegerLiteral ("80"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("grade"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"B\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IsKeyword,
                            Whitespace,
                            GreaterThanOrEqualOperator,
                            Whitespace,
                            IntegerLiteral ("70"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("grade"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"C\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseElseClause {
                            CaseKeyword,
                            Whitespace,
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("grade"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"F\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("temperature"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("0"),
                            Whitespace,
                            ToKeyword,
                            Whitespace,
                            IntegerLiteral ("32"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("status"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"Freezing\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("33"),
                            Whitespace,
                            ToKeyword,
                            Whitespace,
                            IntegerLiteral ("65"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("status"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"Cold\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("66"),
                            Whitespace,
                            ToKeyword,
                            Whitespace,
                            IntegerLiteral ("85"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("status"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"Comfortable\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("86"),
                            Whitespace,
                            ToKeyword,
                            Whitespace,
                            IntegerLiteral ("100"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("status"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"Hot\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("userInput"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\"yes\""),
                            Comma,
                            Whitespace,
                            StringLiteral ("\"y\""),
                            Comma,
                            Whitespace,
                            StringLiteral ("\"YES\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("DoSomething"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\"no\""),
                            Comma,
                            Whitespace,
                            StringLiteral ("\"n\""),
                            Comma,
                            Whitespace,
                            StringLiteral ("\"NO\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("DoSomethingElse"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseElseClause {
                            CaseKeyword,
                            Whitespace,
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("ShowError"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn select_case_nested() {
        let source = r"
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
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                SelectCaseStatement {
                                    Whitespace,
                                    SelectKeyword,
                                    Whitespace,
                                    CaseKeyword,
                                    Whitespace,
                                    IdentifierExpression {
                                        Identifier ("y"),
                                    },
                                    Newline,
                                    Whitespace,
                                    CaseClause {
                                        CaseKeyword,
                                        Whitespace,
                                        IntegerLiteral ("10"),
                                        Newline,
                                        StatementList {
                                            Whitespace,
                                            AssignmentStatement {
                                                IdentifierExpression {
                                                    Identifier ("result"),
                                                },
                                                Whitespace,
                                                EqualityOperator,
                                                Whitespace,
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("11"),
                                                },
                                                Newline,
                                            },
                                            Whitespace,
                                        },
                                    },
                                    CaseClause {
                                        CaseKeyword,
                                        Whitespace,
                                        IntegerLiteral ("20"),
                                        Newline,
                                        StatementList {
                                            Whitespace,
                                            AssignmentStatement {
                                                IdentifierExpression {
                                                    Identifier ("result"),
                                                },
                                                Whitespace,
                                                EqualityOperator,
                                                Whitespace,
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("21"),
                                                },
                                                Newline,
                                            },
                                            Whitespace,
                                        },
                                    },
                                    EndKeyword,
                                    Whitespace,
                                    SelectKeyword,
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("2"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("2"),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("operation"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\"add\""),
                            Newline,
                            StatementList {
                                ForStatement {
                                    Whitespace,
                                    ForKeyword,
                                    Whitespace,
                                    IdentifierExpression {
                                        Identifier ("i"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                    Whitespace,
                                    ToKeyword,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("10"),
                                    },
                                    Newline,
                                    StatementList {
                                        Whitespace,
                                        AssignmentStatement {
                                            IdentifierExpression {
                                                Identifier ("total"),
                                            },
                                            Whitespace,
                                            EqualityOperator,
                                            Whitespace,
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    Identifier ("total"),
                                                },
                                                Whitespace,
                                                AdditionOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("i"),
                                                },
                                            },
                                            Newline,
                                        },
                                        Whitespace,
                                    },
                                    NextKeyword,
                                    Whitespace,
                                    Identifier ("i"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\"multiply\""),
                            Newline,
                            StatementList {
                                ForStatement {
                                    Whitespace,
                                    ForKeyword,
                                    Whitespace,
                                    IdentifierExpression {
                                        Identifier ("i"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                    Whitespace,
                                    ToKeyword,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("10"),
                                    },
                                    Newline,
                                    StatementList {
                                        Whitespace,
                                        AssignmentStatement {
                                            IdentifierExpression {
                                                Identifier ("total"),
                                            },
                                            Whitespace,
                                            EqualityOperator,
                                            Whitespace,
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    Identifier ("total"),
                                                },
                                                Whitespace,
                                                MultiplicationOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("i"),
                                                },
                                            },
                                            Newline,
                                        },
                                        Whitespace,
                                    },
                                    NextKeyword,
                                    Whitespace,
                                    Identifier ("i"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("category"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                IfStatement {
                                    Whitespace,
                                    IfKeyword,
                                    Whitespace,
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("value"),
                                        },
                                        Whitespace,
                                        GreaterThanOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("100"),
                                        },
                                    },
                                    Whitespace,
                                    ThenKeyword,
                                    Newline,
                                    StatementList {
                                        Whitespace,
                                        AssignmentStatement {
                                            IdentifierExpression {
                                                Identifier ("status"),
                                            },
                                            Whitespace,
                                            EqualityOperator,
                                            Whitespace,
                                            StringLiteralExpression {
                                                StringLiteral ("\"high\""),
                                            },
                                            Newline,
                                        },
                                        Whitespace,
                                    },
                                    ElseClause {
                                        ElseKeyword,
                                        Newline,
                                        StatementList {
                                            Whitespace,
                                            AssignmentStatement {
                                                IdentifierExpression {
                                                    Identifier ("status"),
                                                },
                                                Whitespace,
                                                EqualityOperator,
                                                Whitespace,
                                                StringLiteralExpression {
                                                    StringLiteral ("\"low\""),
                                                },
                                                Newline,
                                            },
                                            Whitespace,
                                        },
                                    },
                                    EndKeyword,
                                    Whitespace,
                                    IfKeyword,
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("2"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"category2\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn select_case_empty_case() {
        let source = r"
Sub Test()
    Select Case x
        Case 1
        Case 2
            DoSomething
        Case 3
    End Select
End Sub
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("2"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("DoSomething"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("3"),
                            Newline,
                            StatementList {
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                PublicKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("ModuleLevelTest"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("globalVar"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"One\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("2"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"Two\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        CallExpression {
                            Identifier ("GetValue"),
                            LeftParenthesis,
                            ArgumentList,
                            RightParenthesis,
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"one\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("2"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"two\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("age"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IsKeyword,
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            IntegerLiteral ("13"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("category"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"child\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IsKeyword,
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            IntegerLiteral ("20"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("category"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"teen\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IsKeyword,
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            IntegerLiteral ("65"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("category"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"adult\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseElseClause {
                            CaseKeyword,
                            Whitespace,
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("category"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"senior\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Whitespace,
                            ToKeyword,
                            Whitespace,
                            IntegerLiteral ("5"),
                            Comma,
                            Whitespace,
                            IntegerLiteral ("10"),
                            Comma,
                            Whitespace,
                            IntegerLiteral ("15"),
                            Whitespace,
                            ToKeyword,
                            Whitespace,
                            IntegerLiteral ("20"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"range\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IsKeyword,
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            IntegerLiteral ("100"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"large\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseElseClause {
                            CaseKeyword,
                            Whitespace,
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"other\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
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

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"test\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }
}
