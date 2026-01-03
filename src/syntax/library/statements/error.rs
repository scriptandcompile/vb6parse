use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 Error statement syntax:
    // - Error errornumber
    //
    // Generates a run-time error; can be used instead of the Err.Raise method.
    //
    // The Error statement syntax has this part:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | errornumber   | Required. Any valid error number. |
    //
    // Remarks:
    // - The Error statement is supported for backward compatibility.
    // - In new code, use the Err object's Raise method to generate run-time errors.
    // - If errornumber is defined, the Error statement calls the error handler after the properties
    //   of the Err object are assigned the following default values:
    //   * Err.Number: The value specified as the argument to the Error statement
    //   * Err.Source: The name of the current Visual Basic project
    //   * Err.Description: String expression corresponding to the return value of the Error function
    //     for the specified Number, if this string exists
    //
    // Examples:
    // ```vb
    // Error 11  ' Generate "Division by zero" error
    // Error 53  ' Generate "File not found" error
    // Error vbObjectError + 1000  ' Generate custom error
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/error-statement)
    pub(crate) fn parse_error_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::ErrorStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // Error statement tests

    #[test]
    fn error_simple() {
        let source = r"
Sub Test()
    Error 11
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
                    ErrorStatement {
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        IntegerLiteral ("11"),
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
    fn error_at_module_level() {
        let source = "Error 53\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            ErrorStatement {
                ErrorKeyword,
                Whitespace,
                IntegerLiteral ("53"),
                Newline,
            },
        ]);
        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn error_with_literal() {
        let source = r"
Sub Test()
    Error 9
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
                    ErrorStatement {
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        IntegerLiteral ("9"),
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
    fn error_with_expression() {
        let source = r"
Sub Test()
    Error vbObjectError + 1000
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
                    ErrorStatement {
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        Identifier ("vbObjectError"),
                        Whitespace,
                        AdditionOperator,
                        Whitespace,
                        IntegerLiteral ("1000"),
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
    fn error_with_variable() {
        let source = r"
Sub Test()
    Error errorNumber
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
                    ErrorStatement {
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        Identifier ("errorNumber"),
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
    fn error_preserves_whitespace() {
        let source = "    Error    11    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            ErrorStatement {
                ErrorKeyword,
                Whitespace,
                IntegerLiteral ("11"),
                Whitespace,
                Newline,
            },
        ]);
        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn error_with_comment() {
        let source = r"
Sub Test()
    Error 11 ' Division by zero
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
                    ErrorStatement {
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        IntegerLiteral ("11"),
                        Whitespace,
                        EndOfLineComment,
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
    fn error_in_if_statement() {
        let source = r"
Sub Test()
    If shouldFail Then
        Error 5
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
                            Identifier ("shouldFail"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            ErrorStatement {
                                Whitespace,
                                ErrorKeyword,
                                Whitespace,
                                IntegerLiteral ("5"),
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
    fn error_inline_if() {
        let source = r"
Sub Test()
    If invalidData Then Error 13
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
                            Identifier ("invalidData"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        ErrorStatement {
                            ErrorKeyword,
                            Whitespace,
                            IntegerLiteral ("13"),
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
    fn error_in_select_case() {
        let source = r"
Sub Test()
    Select Case errorType
        Case 1
            Error 11
        Case 2
            Error 13
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
                            Identifier ("errorType"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                ErrorStatement {
                                    Whitespace,
                                    ErrorKeyword,
                                    Whitespace,
                                    IntegerLiteral ("11"),
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
                                ErrorStatement {
                                    Whitespace,
                                    ErrorKeyword,
                                    Whitespace,
                                    IntegerLiteral ("13"),
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
    fn error_with_error_handler() {
        let source = r"
Sub Test()
    On Error GoTo ErrorHandler
    DoSomething
    Exit Sub
ErrorHandler:
    Error 1000
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
                    OnErrorStatement {
                        Whitespace,
                        OnKeyword,
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        GotoKeyword,
                        Whitespace,
                        Identifier ("ErrorHandler"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("DoSomething"),
                        Newline,
                    },
                    ExitStatement {
                        Whitespace,
                        ExitKeyword,
                        Whitespace,
                        SubKeyword,
                        Newline,
                    },
                    LabelStatement {
                        Identifier ("ErrorHandler"),
                        ColonOperator,
                        Newline,
                    },
                    ErrorStatement {
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        IntegerLiteral ("1000"),
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
    fn error_custom_number() {
        let source = r"
Sub Test()
    Error 32000
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
                    ErrorStatement {
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        IntegerLiteral ("32000"),
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
    fn multiple_error_statements() {
        let source = r"
Sub Test()
    Error 1
    DoSomething
    Error 2
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
                    ErrorStatement {
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        IntegerLiteral ("1"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("DoSomething"),
                        Newline,
                    },
                    ErrorStatement {
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        IntegerLiteral ("2"),
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
