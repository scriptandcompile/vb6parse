use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 Date statement syntax:
    // - Date = dateexpression
    //
    // Sets the current system date.
    //
    // dateexpression: Required. Any expression that can represent a date.
    //
    // Note: The Date statement is used to set the date. To retrieve the current date,
    // use the Date function.
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/date-statement)
    pub(crate) fn parse_date_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::DateStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn date_simple() {
        let source = r"
Sub Test()
    Date = #1/1/2024#
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
                    DateStatement {
                        Whitespace,
                        DateKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        DateLiteral ("#1/1/2024#"),
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
    fn date_with_variable() {
        let source = r"
Sub Test()
    Date = newDate
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
                    DateStatement {
                        Whitespace,
                        DateKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("newDate"),
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
    fn date_with_function_call() {
        let source = r"
Sub Test()
    Date = DateSerial(2024, 1, 1)
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
                    DateStatement {
                        Whitespace,
                        DateKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("DateSerial"),
                        LeftParenthesis,
                        IntegerLiteral ("2024"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("1"),
                        RightParenthesis,
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
    fn date_with_string_expression() {
        let source = r#"
Sub Test()
    Date = "January 1, 2024"
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
                    DateStatement {
                        Whitespace,
                        DateKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"January 1, 2024\""),
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
    fn date_with_expression() {
        let source = r"
Sub Test()
    Date = Now() + 7
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
                    DateStatement {
                        Whitespace,
                        DateKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("Now"),
                        LeftParenthesis,
                        RightParenthesis,
                        Whitespace,
                        AdditionOperator,
                        Whitespace,
                        IntegerLiteral ("7"),
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
    fn date_preserves_whitespace() {
        let source = r"
Sub Test()
    Date   =   #1/1/2024#
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
                    DateStatement {
                        Whitespace,
                        DateKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        DateLiteral ("#1/1/2024#"),
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
    fn multiple_date_statements() {
        let source = r"
Sub Test()
    Date = #1/1/2024#
    Date = #2/1/2024#
    Date = #3/1/2024#
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
                    DateStatement {
                        Whitespace,
                        DateKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        DateLiteral ("#1/1/2024#"),
                        Newline,
                    },
                    DateStatement {
                        Whitespace,
                        DateKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        DateLiteral ("#2/1/2024#"),
                        Newline,
                    },
                    DateStatement {
                        Whitespace,
                        DateKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        DateLiteral ("#3/1/2024#"),
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
    fn date_in_if_statement() {
        let source = r"
Sub Test()
    If resetDate Then
        Date = #1/1/2024#
    End If
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
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("resetDate"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            DateStatement {
                                Whitespace,
                                DateKeyword,
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                DateLiteral ("#1/1/2024#"),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
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
    fn date_inline_if() {
        let source = r"
Sub Test()
    If resetDate Then Date = #1/1/2024#
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
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("resetDate"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        DateStatement {
                            DateKeyword,
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            DateLiteral ("#1/1/2024#"),
                            Newline,
                        },
                        EndKeyword,
                        Whitespace,
                        SubKeyword,
                        Newline,
                    },
                },
            },
        ]);
    }

    #[test]
    fn date_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Date = userDate
    If Err Then MsgBox "Invalid date"
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
                    OnErrorStatement {
                        Whitespace,
                        OnKeyword,
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        ResumeKeyword,
                        Whitespace,
                        NextKeyword,
                        Newline,
                    },
                    DateStatement {
                        Whitespace,
                        DateKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("userDate"),
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("Err"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Invalid date\""),
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

    fn date_at_module_level() {
        let source = r"
Date = #1/1/2024#
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DateStatement {
                DateKeyword,
                Whitespace,
                EqualityOperator,
                Whitespace,
                DateLiteral ("#1/1/2024#"),
                Newline,
            },
        ]);
    }

    #[test]
    fn date_with_dateadd() {
        let source = r#"
Sub Test()
    Date = DateAdd("d", 30, Date)
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
                    DateStatement {
                        Whitespace,
                        DateKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        Identifier ("DateAdd"),
                        LeftParenthesis,
                        StringLiteral ("\"d\""),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("30"),
                        Comma,
                        Whitespace,
                        DateKeyword,
                        RightParenthesis,
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
