use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 Beep statement syntax:
    // - Beep
    //
    // Emits a standard system beep sound.
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/beep-statement)
    pub(crate) fn parse_beep_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::BeepStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn beep_simple() {
        let source = r"
Sub Test()
    Beep
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
                    BeepStatement {
                        Whitespace,
                        BeepKeyword,
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
    fn beep_in_if_statement() {
        let source = r"
Sub Test()
    If condition Then
        Beep
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
                            Identifier ("condition"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            BeepStatement {
                                Whitespace,
                                BeepKeyword,
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
    fn beep_inline_if() {
        let source = r"
Sub Test()
    If error Then Beep
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
                            ErrorKeyword,
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        BeepStatement {
                            BeepKeyword,
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
    fn multiple_beep_statements() {
        let source = r"
Sub Test()
    Beep
    Beep
    Beep
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
                    BeepStatement {
                        Whitespace,
                        BeepKeyword,
                        Newline,
                    },
                    BeepStatement {
                        Whitespace,
                        BeepKeyword,
                        Newline,
                    },
                    BeepStatement {
                        Whitespace,
                        BeepKeyword,
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
    fn beep_with_comment() {
        let source = r"
Sub Test()
    Beep ' Alert user
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
                    BeepStatement {
                        Whitespace,
                        BeepKeyword,
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
    fn beep_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 3
        Beep
    Next i
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
                            IntegerLiteral ("3"),
                        },
                        Newline,
                        StatementList {
                            BeepStatement {
                                Whitespace,
                                BeepKeyword,
                                Newline,
                            },
                            Whitespace,
                        },
                        NextKeyword,
                        Whitespace,
                        Identifier ("i"),
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
    fn beep_in_select_case() {
        let source = r"
Sub Test()
    Select Case value
        Case 1
            Beep
        Case Else
            Beep
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
                                BeepStatement {
                                    Whitespace,
                                    BeepKeyword,
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
                                BeepStatement {
                                    Whitespace,
                                    BeepKeyword,
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
    fn beep_preserves_whitespace() {
        let source = r"
Sub Test()
    Beep    ' Extra spaces
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
                    BeepStatement {
                        Whitespace,
                        BeepKeyword,
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
    fn beep_at_module_level() {
        let source = r"
Beep
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(
            cst,
            [
                Newline,
                BeepStatement {
                    BeepKeyword,
                    Newline,
                },
            ]
        );
    }

    #[test]
    fn beep_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Beep
    If Err Then MsgBox "Error occurred"
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
                    BeepStatement {
                        Whitespace,
                        BeepKeyword,
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
                        StringLiteral ("\"Error occurred\""),
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
