use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 Kill statement syntax:
    // - Kill pathname
    //
    // Deletes files from a disk.
    //
    // The Kill statement syntax has this part:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | pathname      | Required. String expression that specifies one or more file names to be deleted. May include directory or folder, and drive. |
    //
    // Remarks:
    // - Kill supports the use of multiple-character (*) and single-character (?) wildcards to specify multiple files.
    // - An error occurs if you try to use Kill to delete an open file.
    // - To remove a directory or folder, use the RmDir statement.
    //
    // Examples:
    // ```vb
    // Kill "C:\DATA.TXT"
    // Kill "C:\*.TXT"           ' Delete all .txt files
    // Kill "C:\TEST?.TXT"       ' Delete TEST1.TXT, TESTA.TXT, etc.
    // Kill App.Path & "\temp.dat"
    // Kill myFileName
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/kill-statement)
    pub(crate) fn parse_kill_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::KillStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn kill_simple() {
        let source = r#"
Sub Test()
    Kill "C:\DATA.TXT"
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
                    KillStatement {
                        Whitespace,
                        KillKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\DATA.TXT\""),
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
    fn kill_module_level() {
        let source = "Kill \"temp.dat\"\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            KillStatement {
                KillKeyword,
                Whitespace,
                StringLiteral ("\"temp.dat\""),
                Newline,
            },
        ]);
    }

    #[test]
    fn kill_with_wildcard() {
        let source = r#"
Sub Test()
    Kill "C:\*.TXT"
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
                    KillStatement {
                        Whitespace,
                        KillKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\*.TXT\""),
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
    fn kill_with_single_wildcard() {
        let source = r#"
Sub Test()
    Kill "C:\TEST?.TXT"
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
                    KillStatement {
                        Whitespace,
                        KillKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\TEST?.TXT\""),
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
    fn kill_with_variable() {
        let source = r"
Sub Test()
    Kill myFileName
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
                    KillStatement {
                        Whitespace,
                        KillKeyword,
                        Whitespace,
                        Identifier ("myFileName"),
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
    fn kill_with_app_path() {
        let source = r#"
Sub Test()
    Kill App.Path & "\temp.dat"
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
                    KillStatement {
                        Whitespace,
                        KillKeyword,
                        Whitespace,
                        Identifier ("App"),
                        PeriodOperator,
                        Identifier ("Path"),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"\\temp.dat\""),
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
    fn kill_preserves_whitespace() {
        let source = "    Kill    \"C:\\\\file.txt\"    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            KillStatement {
                KillKeyword,
                Whitespace,
                StringLiteral ("\"C:\\\\file.txt\""),
                Whitespace,
                Newline,
            },
        ]);
    }

    #[test]
    fn kill_with_comment() {
        let source = r#"
Sub Test()
    Kill "temp.txt" ' Delete temporary file
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
                    KillStatement {
                        Whitespace,
                        KillKeyword,
                        Whitespace,
                        StringLiteral ("\"temp.txt\""),
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
    fn kill_in_if_statement() {
        let source = r"
Sub Test()
    If fileExists Then
        Kill fileName
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
                            Identifier ("fileExists"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            KillStatement {
                                Whitespace,
                                KillKeyword,
                                Whitespace,
                                Identifier ("fileName"),
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
    fn kill_inline_if() {
        let source = r#"
Sub Test()
    If fileExists Then Kill "temp.dat"
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
                            Identifier ("fileExists"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        KillStatement {
                            KillKeyword,
                            Whitespace,
                            StringLiteral ("\"temp.dat\""),
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
    fn kill_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Kill "temp.txt"
    If Err.Number <> 0 Then
        MsgBox "Could not delete file"
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
                    KillStatement {
                        Whitespace,
                        KillKeyword,
                        Whitespace,
                        StringLiteral ("\"temp.txt\""),
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            MemberAccessExpression {
                                Identifier ("Err"),
                                PeriodOperator,
                                Identifier ("Number"),
                            },
                            Whitespace,
                            InequalityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("MsgBox"),
                                Whitespace,
                                StringLiteral ("\"Could not delete file\""),
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
    fn multiple_kill_statements() {
        let source = r#"
Sub Test()
    Kill "temp1.txt"
    Kill "temp2.txt"
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
                    KillStatement {
                        Whitespace,
                        KillKeyword,
                        Whitespace,
                        StringLiteral ("\"temp1.txt\""),
                        Newline,
                    },
                    KillStatement {
                        Whitespace,
                        KillKeyword,
                        Whitespace,
                        StringLiteral ("\"temp2.txt\""),
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
    fn kill_with_dir_function() {
        let source = r#"
Sub Test()
    fileName = Dir("C:\*.tmp")
    Do While fileName <> ""
        Kill "C:\" & fileName
        fileName = Dir
    Loop
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("fileName"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Dir"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"C:\\*.tmp\""),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("fileName"),
                            },
                            Whitespace,
                            InequalityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"\""),
                            },
                        },
                        Newline,
                        StatementList {
                            KillStatement {
                                Whitespace,
                                KillKeyword,
                                Whitespace,
                                StringLiteral ("\"C:\\\" & fileName"),
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("fileName"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("Dir"),
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        LoopKeyword,
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
