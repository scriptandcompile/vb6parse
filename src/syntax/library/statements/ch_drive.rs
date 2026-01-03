use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 ChDrive statement syntax:
    // - ChDrive drive
    //
    // Changes the current drive.
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chdrive-statement)
    pub(crate) fn parse_ch_drive_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::ChDriveStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn chdrive_simple_string_literal() {
        let source = r#"
Sub Test()
    ChDrive "C:"
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
                    ChDriveStatement {
                        Whitespace,
                        ChDriveKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\""),
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
    fn chdrive_with_variable() {
        let source = r"
Sub Test()
    ChDrive myDrive
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
                    ChDriveStatement {
                        Whitespace,
                        ChDriveKeyword,
                        Whitespace,
                        Identifier ("myDrive"),
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
    fn chdrive_with_app_path() {
        let source = r"
Sub Test()
    ChDrive App.Path
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
                    ChDriveStatement {
                        Whitespace,
                        ChDriveKeyword,
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
    fn chdrive_with_left_function() {
        let source = r"
Sub Test()
    ChDrive Left(sInitDir, 1)
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
                    ChDriveStatement {
                        Whitespace,
                        ChDriveKeyword,
                        Whitespace,
                        Identifier ("Left"),
                        LeftParenthesis,
                        Identifier ("sInitDir"),
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
    fn chdrive_in_if_statement() {
        let source = r"
Sub Test()
    If driveValid Then ChDrive newDrive
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
                            Identifier ("driveValid"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        ChDriveStatement {
                            ChDriveKeyword,
                            Whitespace,
                            Identifier ("newDrive"),
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
    fn chdrive_at_module_level() {
        let source = r#"
ChDrive "D:"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            ChDriveStatement {
                ChDriveKeyword,
                Whitespace,
                StringLiteral ("\"D:\""),
                Newline,
            },
        ]);
    }

    #[test]
    fn chdrive_with_comment() {
        let source = r"
Sub Test()
    ChDrive driveLetter ' Change to specified drive
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
                    ChDriveStatement {
                        Whitespace,
                        ChDriveKeyword,
                        Whitespace,
                        Identifier ("driveLetter"),
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
    fn chdrive_multiple_in_sequence() {
        let source = r#"
Sub Test()
    ChDrive "C:"
    ChDrive "D:"
    ChDrive originalDrive
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
                    ChDriveStatement {
                        Whitespace,
                        ChDriveKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\""),
                        Newline,
                    },
                    ChDriveStatement {
                        Whitespace,
                        ChDriveKeyword,
                        Whitespace,
                        StringLiteral ("\"D:\""),
                        Newline,
                    },
                    ChDriveStatement {
                        Whitespace,
                        ChDriveKeyword,
                        Whitespace,
                        Identifier ("originalDrive"),
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
    fn chdrive_in_multiline_if() {
        let source = r"
Sub Test()
    If driveExists Then
        ChDrive targetDrive
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
                            Identifier ("driveExists"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            ChDriveStatement {
                                Whitespace,
                                ChDriveKeyword,
                                Whitespace,
                                Identifier ("targetDrive"),
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
    fn chdrive_with_parentheses() {
        let source = r"
Sub Test()
    ChDrive (Left$(sInitDir, 1))
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
                    ChDriveStatement {
                        Whitespace,
                        ChDriveKeyword,
                        Whitespace,
                        LeftParenthesis,
                        Identifier ("Left$"),
                        LeftParenthesis,
                        Identifier ("sInitDir"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("1"),
                        RightParenthesis,
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
    fn chdrive_with_expression() {
        let source = r"
Sub Test()
    ChDrive Left(theZtmPath, 1)
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
                    ChDriveStatement {
                        Whitespace,
                        ChDriveKeyword,
                        Whitespace,
                        Identifier ("Left"),
                        LeftParenthesis,
                        Identifier ("theZtmPath"),
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
    fn chdrive_and_chdir_together() {
        let source = r#"
Sub Test()
    ChDrive "C:"
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
                    ChDriveStatement {
                        Whitespace,
                        ChDriveKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\""),
                        Newline,
                    },
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
}
