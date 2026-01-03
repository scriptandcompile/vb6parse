use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 ChDir statement syntax:
    // - ChDir path
    //
    // Changes the current directory or folder.
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chdir-statement)
    pub(crate) fn parse_ch_dir_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::ChDirStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn chdir_simple_string_literal() {
        let source = r#"
Sub Test()
    ChDir "C:\Windows"
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
                    ChDirStatement {
                        Whitespace,
                        ChDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Windows\""),
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
    fn chdir_with_variable() {
        let source = r"
Sub Test()
    ChDir myPath
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
                    ChDirStatement {
                        Whitespace,
                        ChDirKeyword,
                        Whitespace,
                        Identifier ("myPath"),
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
    fn chdir_with_app_path() {
        let source = r"
Sub Test()
    ChDir App.Path
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
                    ChDirStatement {
                        Whitespace,
                        ChDirKeyword,
                        Whitespace,
                        Identifier ("App"),
                        PeriodOperator,
                        Identifier ("Path"),
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
    fn chdir_with_expression() {
        let source = r#"
Sub Test()
    ChDir GetPath() & "\subdir"
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
                    ChDirStatement {
                        Whitespace,
                        ChDirKeyword,
                        Whitespace,
                        Identifier ("GetPath"),
                        LeftParenthesis,
                        RightParenthesis,
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"\\subdir\""),
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
    fn chdir_in_if_statement() {
        let source = r"
Sub Test()
    If dirExists Then ChDir newPath
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
                            Identifier ("dirExists"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        ChDirStatement {
                            ChDirKeyword,
                            Whitespace,
                            Identifier ("newPath"),
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
    fn chdir_at_module_level() {
        let source = r#"
ChDir "C:\Temp"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            ChDirStatement {
                ChDirKeyword,
                Whitespace,
                StringLiteral ("\"C:\\Temp\""),
                Newline,
            },
        ]);
    }

    #[test]
    fn chdir_with_comment() {
        let source = r"
Sub Test()
    ChDir basePath ' Change to base directory
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
                    ChDirStatement {
                        Whitespace,
                        ChDirKeyword,
                        Whitespace,
                        Identifier ("basePath"),
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
    fn chdir_multiple_in_sequence() {
        let source = r#"
Sub Test()
    ChDir "C:\Windows"
    ChDir "C:\Temp"
    ChDir originalPath
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
                    ChDirStatement {
                        Whitespace,
                        ChDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Windows\""),
                        Newline,
                    },
                    ChDirStatement {
                        Whitespace,
                        ChDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Temp\""),
                        Newline,
                    },
                    ChDirStatement {
                        Whitespace,
                        ChDirKeyword,
                        Whitespace,
                        Identifier ("originalPath"),
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
    fn chdir_in_multiline_if() {
        let source = r"
Sub Test()
    If pathValid Then
        ChDir newPath
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
                            Identifier ("pathValid"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            ChDirStatement {
                                Whitespace,
                                ChDirKeyword,
                                Whitespace,
                                Identifier ("newPath"),
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
    fn chdir_with_parentheses() {
        let source = r"
Sub Test()
    ChDir (basePath)
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
                    ChDirStatement {
                        Whitespace,
                        ChDirKeyword,
                        Whitespace,
                        LeftParenthesis,
                        Identifier ("basePath"),
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
    fn chdir_with_parentheses_without_space() {
        let source = r"
Sub Test()
    ChDir(basePath)
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
                    ChDirStatement {
                        Whitespace,
                        ChDirKeyword,
                        LeftParenthesis,
                        Identifier ("basePath"),
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
