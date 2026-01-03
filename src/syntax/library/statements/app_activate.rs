use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 AppActivate statement syntax:
    // - AppActivate title[, wait]
    //
    // Activates an application window.
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/appactivate-statement)
    pub(crate) fn parse_app_activate_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::AppActivateStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn appactivate_simple() {
        let source = r#"
Sub Test()
    AppActivate "MyApp"
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
                    AppActivateStatement {
                        Whitespace,
                        AppActivateKeyword,
                        Whitespace,
                        StringLiteral ("\"MyApp\""),
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
    fn appactivate_with_variable() {
        let source = r"
Sub Test()
    AppActivate lstTopWin.Text
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
                    AppActivateStatement {
                        Whitespace,
                        AppActivateKeyword,
                        Whitespace,
                        Identifier ("lstTopWin"),
                        PeriodOperator,
                        TextKeyword,
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
    fn appactivate_with_wait_parameter() {
        let source = r#"
Sub Test()
    AppActivate "Calculator", True
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
                    AppActivateStatement {
                        Whitespace,
                        AppActivateKeyword,
                        Whitespace,
                        StringLiteral ("\"Calculator\""),
                        Comma,
                        Whitespace,
                        TrueKeyword,
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
    fn appactivate_with_title_variable() {
        let source = r"
Sub Test()
    AppActivate sTitle
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
                    AppActivateStatement {
                        Whitespace,
                        AppActivateKeyword,
                        Whitespace,
                        Identifier ("sTitle"),
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
    fn appactivate_preserves_whitespace() {
        let source = r#"
Sub Test()
    AppActivate   "MyApp"  ,  False
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
                    AppActivateStatement {
                        Whitespace,
                        AppActivateKeyword,
                        Whitespace,
                        StringLiteral ("\"MyApp\""),
                        Whitespace,
                        Comma,
                        Whitespace,
                        FalseKeyword,
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
    fn multiple_appactivate_statements() {
        let source = r#"
Sub Test()
    AppActivate "App1"
    AppActivate "App2"
    AppActivate windowTitle
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
                    AppActivateStatement {
                        Whitespace,
                        AppActivateKeyword,
                        Whitespace,
                        StringLiteral ("\"App1\""),
                        Newline,
                    },
                    AppActivateStatement {
                        Whitespace,
                        AppActivateKeyword,
                        Whitespace,
                        StringLiteral ("\"App2\""),
                        Newline,
                    },
                    AppActivateStatement {
                        Whitespace,
                        AppActivateKeyword,
                        Whitespace,
                        Identifier ("windowTitle"),
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
    fn appactivate_in_if_statement() {
        let source = r#"
Sub Test()
    If condition Then
        AppActivate "MyApp"
    End If
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
                            AppActivateStatement {
                                Whitespace,
                                AppActivateKeyword,
                                Whitespace,
                                StringLiteral ("\"MyApp\""),
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
    fn appactivate_inline_if() {
        let source = r"
Sub Test()
    If windowExists Then AppActivate windowTitle
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
                            Identifier ("windowExists"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        AppActivateStatement {
                            AppActivateKeyword,
                            Whitespace,
                            Identifier ("windowTitle"),
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
    fn appactivate_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    AppActivate lstTopWin.Text
    If Err Then MsgBox "AppActivate error: " & Err
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
                    AppActivateStatement {
                        Whitespace,
                        AppActivateKeyword,
                        Whitespace,
                        Identifier ("lstTopWin"),
                        PeriodOperator,
                        TextKeyword,
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
                        StringLiteral ("\"AppActivate error: \""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Err"),
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
    fn appactivate_at_module_level() {
        let source = r#"
AppActivate "MyApp"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AppActivateStatement {
                AppActivateKeyword,
                Whitespace,
                StringLiteral ("\"MyApp\""),
                Newline,
            },
        ]);
    }
}
