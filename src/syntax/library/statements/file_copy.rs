use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 FileCopy statement syntax:
    // - FileCopy source, destination
    //
    // Copies a file.
    //
    // The FileCopy statement syntax has these named arguments:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | source        | Required. String expression that specifies a file name. May include directory or folder, and drive. |
    // | destination   | Required. String expression that specifies a file name. May include directory or folder, and drive. |
    //
    // Remarks:
    // - If you try to use the FileCopy statement on a currently open file, an error occurs.
    // - FileCopy can copy files between directories/folders and between drives.
    // - Both source and destination can include path information (drive and directory/folder).
    // - If destination specifies a directory/folder that doesn't exist, FileCopy creates it.
    //
    // Examples:
    // ```vb
    // FileCopy "C:\SOURCE.TXT", "C:\DEST.TXT"
    // FileCopy oldFile, newFile
    // FileCopy App.Path & "\data.dat", "C:\Backup\data.dat"
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filecopy-statement)
    pub(crate) fn parse_file_copy_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::FileCopyStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // FileCopy statement tests

    #[test]
    fn filecopy_simple() {
        let source = r#"
Sub Test()
    FileCopy "source.txt", "dest.txt"
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
                    FileCopyStatement {
                        Whitespace,
                        FileCopyKeyword,
                        Whitespace,
                        StringLiteral ("\"source.txt\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"dest.txt\""),
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
    fn filecopy_at_module_level() {
        let source = "FileCopy \"old.dat\", \"new.dat\"\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            FileCopyStatement {
                FileCopyKeyword,
                Whitespace,
                StringLiteral ("\"old.dat\""),
                Comma,
                Whitespace,
                StringLiteral ("\"new.dat\""),
                Newline,
            },
        ]);
    }

    #[test]
    fn filecopy_with_paths() {
        let source = r#"
Sub Test()
    FileCopy "C:\\SOURCE.TXT", "C:\\DEST.TXT"
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
                    FileCopyStatement {
                        Whitespace,
                        FileCopyKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\\\SOURCE.TXT\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"C:\\\\DEST.TXT\""),
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
    fn filecopy_with_variables() {
        let source = r"
Sub Test()
    FileCopy oldFile, newFile
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
                    FileCopyStatement {
                        Whitespace,
                        FileCopyKeyword,
                        Whitespace,
                        Identifier ("oldFile"),
                        Comma,
                        Whitespace,
                        Identifier ("newFile"),
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
    fn filecopy_with_expressions() {
        let source = r#"
Sub Test()
    FileCopy App.Path & "\\data.dat", "C:\\Backup\\data.dat"
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
                    FileCopyStatement {
                        Whitespace,
                        FileCopyKeyword,
                        Whitespace,
                        Identifier ("App"),
                        PeriodOperator,
                        Identifier ("Path"),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"\\\\data.dat\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"C:\\\\Backup\\\\data.dat\""),
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
    fn filecopy_preserves_whitespace() {
        let source = "    FileCopy    \"a.txt\"  ,  \"b.txt\"    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            FileCopyStatement {
                FileCopyKeyword,
                Whitespace,
                StringLiteral ("\"a.txt\""),
                Whitespace,
                Comma,
                Whitespace,
                StringLiteral ("\"b.txt\""),
                Whitespace,
                Newline,
            },
        ]);
    }

    #[test]
    fn filecopy_with_comment() {
        let source = r#"
Sub Test()
    FileCopy "source.dat", "dest.dat" ' Backup file
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
                    FileCopyStatement {
                        Whitespace,
                        FileCopyKeyword,
                        Whitespace,
                        StringLiteral ("\"source.dat\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"dest.dat\""),
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
    fn filecopy_in_if_statement() {
        let source = r#"
Sub Test()
    If needBackup Then
        FileCopy "data.txt", "backup.txt"
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
                            Identifier ("needBackup"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            FileCopyStatement {
                                Whitespace,
                                FileCopyKeyword,
                                Whitespace,
                                StringLiteral ("\"data.txt\""),
                                Comma,
                                Whitespace,
                                StringLiteral ("\"backup.txt\""),
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
    fn filecopy_inline_if() {
        let source = r#"
Sub Test()
    If createBackup Then FileCopy "file.txt", "file.bak"
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
                            Identifier ("createBackup"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        FileCopyStatement {
                            FileCopyKeyword,
                            Whitespace,
                            StringLiteral ("\"file.txt\""),
                            Comma,
                            Whitespace,
                            StringLiteral ("\"file.bak\""),
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
    fn filecopy_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    FileCopy "source.dat", "dest.dat"
    If Err.Number <> 0 Then
        MsgBox "Error copying file"
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
                    FileCopyStatement {
                        Whitespace,
                        FileCopyKeyword,
                        Whitespace,
                        StringLiteral ("\"source.dat\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"dest.dat\""),
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
                                StringLiteral ("\"Error copying file\""),
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
    fn filecopy_in_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        FileCopy "file" & i & ".txt", "backup" & i & ".txt"
    Next i
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
                            FileCopyStatement {
                                Whitespace,
                                FileCopyKeyword,
                                Whitespace,
                                StringLiteral ("\"file\""),
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                Identifier ("i"),
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                StringLiteral ("\".txt\""),
                                Comma,
                                Whitespace,
                                StringLiteral ("\"backup\""),
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                Identifier ("i"),
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                StringLiteral ("\".txt\""),
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
    fn multiple_filecopy_statements() {
        let source = r#"
Sub Test()
    FileCopy "file1.txt", "backup1.txt"
    FileCopy "file2.txt", "backup2.txt"
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
                    FileCopyStatement {
                        Whitespace,
                        FileCopyKeyword,
                        Whitespace,
                        StringLiteral ("\"file1.txt\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"backup1.txt\""),
                        Newline,
                    },
                    FileCopyStatement {
                        Whitespace,
                        FileCopyKeyword,
                        Whitespace,
                        StringLiteral ("\"file2.txt\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"backup2.txt\""),
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
    fn filecopy_network_paths() {
        let source = r#"
Sub Test()
    FileCopy "\\\\server\\share\\file.dat", "C:\\local\\file.dat"
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
                    FileCopyStatement {
                        Whitespace,
                        FileCopyKeyword,
                        Whitespace,
                        StringLiteral ("\"\\\\\\\\server\\\\share\\\\file.dat\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"C:\\\\local\\\\file.dat\""),
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
