//! # `RmDir` Statement
//!
//! Removes an empty directory or folder.
//!
//! ## Syntax
//!
//! ```vb
//! RmDir path
//! ```
//!
//! - `path`: Required. String expression that identifies the directory or folder to be removed. May include drive.
//!   If no drive is specified, `RmDir` removes the directory or folder on the current drive.
//!
//! ## Remarks
//!
//! - An error occurs if you try to use `RmDir` on a directory containing files. Use the Kill statement to delete all files before attempting to remove a directory.
//! - An error also occurs if you try to remove a directory that doesn't exist.
//! - The directory must be empty (contain no files or subdirectories) before it can be removed.
//! - The `path` argument can include absolute or relative paths.
//! - On Windows systems, both forward slashes (/) and backslashes (\) can be used as path separators.
//! - The directory name can include the drive letter.
//! - You cannot remove the current directory. You must change to a parent or different directory first.
//! - UNC paths are supported on network drives.
//! - To remove a directory tree, you must remove all subdirectories first (working from innermost to outermost).
//!
//! ## Examples
//!
//! ```vb
//! ' Remove a directory in the current directory
//! RmDir "OldFolder"
//!
//! ' Remove a directory with full path
//! RmDir "C:\Temp\TempFiles"
//!
//! ' Remove a directory on another drive
//! RmDir "D:\Data\Archive"
//!
//! ' Remove nested directories (must remove innermost first)
//! RmDir "C:\Temp\Logs\Archive"
//! RmDir "C:\Temp\Logs"
//! RmDir "C:\Temp"
//!
//! ' Remove directory on network drive
//! RmDir "\\Server\Share\OldFolder"
//!
//! ' Safe removal with error handling
//! On Error Resume Next
//! RmDir "C:\Temp\ToDelete"
//! If Err.Number <> 0 Then
//!     MsgBox "Could not remove directory"
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Common Errors
//!
//! - **Error 75**: Path/File access error - directory contains files or subdirectories
//! - **Error 76**: Path not found - directory doesn't exist
//! - **Error 5**: Invalid procedure call - trying to remove current directory
//!
//! ## Reference
//!
//! [RmDir Statement - Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/rmdir-statement)

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses an `RmDir` statement.
    pub(crate) fn parse_rmdir_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::RmDirStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // RmDir statement tests

    #[test]
    fn rmdir_simple() {
        let source = r#"
Sub Test()
    RmDir "OldFolder"
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"OldFolder\""),
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
    fn rmdir_at_module_level() {
        let source = r#"RmDir "C:\Temp""#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            RmDirStatement {
                RmDirKeyword,
                Whitespace,
                StringLiteral ("\"C:\\Temp\""),
            },
        ]);
    }

    #[test]
    fn rmdir_with_full_path() {
        let source = r#"
Sub Test()
    RmDir "C:\Temp\TempFiles"
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Temp\\TempFiles\""),
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
    fn rmdir_with_variable() {
        let source = r#"
Sub Test()
    Dim folderPath As String
    folderPath = "C:\TempData"
    RmDir folderPath
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
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("folderPath"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("folderPath"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"C:\\TempData\""),
                        },
                        Newline,
                    },
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        Identifier ("folderPath"),
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
    fn rmdir_with_drive_letter() {
        let source = r#"
Sub Test()
    RmDir "D:\Archive"
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"D:\\Archive\""),
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
    fn rmdir_relative_path() {
        let source = r#"
Sub Test()
    RmDir "SubFolder"
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"SubFolder\""),
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
    fn rmdir_parent_directory() {
        let source = r#"
Sub Test()
    RmDir "..\TempFolder"
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"..\\TempFolder\""),
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
    fn rmdir_unc_path() {
        let source = r#"
Sub Test()
    RmDir "\\Server\Share\TempFolder"
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"\\\\Server\\Share\\TempFolder\""),
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
    fn rmdir_in_if_statement() {
        let source = r#"
Sub Test()
    If folderExists Then
        RmDir "C:\Temp"
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
                            Identifier ("folderExists"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            RmDirStatement {
                                Whitespace,
                                RmDirKeyword,
                                Whitespace,
                                StringLiteral ("\"C:\\Temp\""),
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
    fn rmdir_inline_if() {
        let source = r"
Sub Test()
    If shouldDelete Then RmDir tempPath
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
                            Identifier ("shouldDelete"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        RmDirStatement {
                            RmDirKeyword,
                            Whitespace,
                            Identifier ("tempPath"),
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
    fn rmdir_with_comment() {
        let source = r#"
Sub Test()
    RmDir "C:\Temp" ' Remove temporary directory
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Temp\""),
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
    fn rmdir_preserves_whitespace() {
        let source = r#"    RmDir    "C:\Temp"    "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            RmDirStatement {
                RmDirKeyword,
                Whitespace,
                StringLiteral ("\"C:\\Temp\""),
                Whitespace,
            },
        ]);
    }

    #[test]
    fn rmdir_with_error_handling() {
        let source = r#"
Sub CleanupDirectories()
    On Error Resume Next
    RmDir "C:\Temp\Session1"
    If Err.Number <> 0 Then
        MsgBox "Could not remove directory"
    End If
    On Error GoTo 0
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("CleanupDirectories"),
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Temp\\Session1\""),
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
                                StringLiteral ("\"Could not remove directory\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                    OnErrorStatement {
                        Whitespace,
                        OnKeyword,
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        GotoKeyword,
                        Whitespace,
                        IntegerLiteral ("0"),
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
    fn rmdir_multiple_directories() {
        let source = r#"
Sub Test()
    RmDir "C:\Temp\Folder1"
    RmDir "C:\Temp\Folder2"
    RmDir "C:\Temp\Folder3"
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Temp\\Folder1\""),
                        Newline,
                    },
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Temp\\Folder2\""),
                        Newline,
                    },
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Temp\\Folder3\""),
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
    fn rmdir_in_loop() {
        let source = r#"
Sub RemoveTemporaryFolders()
    Dim i As Integer
    For i = 1 To 10
        RmDir "C:\Temp\Session" & i
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
                Identifier ("RemoveTemporaryFolders()"),
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
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
                            RmDirStatement {
                                Whitespace,
                                RmDirKeyword,
                                Whitespace,
                                StringLiteral ("\"C:\\Temp\\Session\""),
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                Identifier ("i"),
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
    fn rmdir_nested_removal() {
        let source = r#"
Sub RemoveNestedDirectories()
    ' Remove innermost first
    RmDir "C:\Project\Temp\Cache\Old"
    RmDir "C:\Project\Temp\Cache"
    RmDir "C:\Project\Temp"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("RemoveNestedDirectories()"),
                Newline,
                StatementList {
                    Whitespace,
                    EndOfLineComment,
                    Newline,
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Project\\Temp\\Cache\\Old\""),
                        Newline,
                    },
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Project\\Temp\\Cache\""),
                        Newline,
                    },
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Project\\Temp\""),
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
    fn rmdir_with_concatenation() {
        let source = r#"
Sub Test()
    Dim basePath As String
    basePath = "C:\Temp\"
    RmDir basePath & "OldData"
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
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("basePath"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("basePath"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"C:\\Temp\\\""),
                        },
                        Newline,
                    },
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        Identifier ("basePath"),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"OldData\""),
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
    fn rmdir_after_kill() {
        let source = r#"
Sub RemoveDirectoryWithFiles()
    On Error Resume Next
    Kill "C:\Temp\Session\*.*"
    RmDir "C:\Temp\Session"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("RemoveDirectoryWithFiles()"),
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
                        StringLiteral ("\"C:\\Temp\\Session\\*.*\""),
                        Newline,
                    },
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Temp\\Session\""),
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
    fn rmdir_with_function_result() {
        let source = r"
Sub Test()
    RmDir GetTempPath()
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        Identifier ("GetTempPath"),
                        LeftParenthesis,
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
    fn rmdir_in_select_case() {
        let source = r#"
Sub Test()
    Select Case action
        Case "cleanup"
            RmDir "C:\Temp"
        Case "archive"
            RmDir "C:\Archive\Old"
        Case Else
            RmDir defaultPath
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
                            Identifier ("action"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\"cleanup\""),
                            Newline,
                            StatementList {
                                RmDirStatement {
                                    Whitespace,
                                    RmDirKeyword,
                                    Whitespace,
                                    StringLiteral ("\"C:\\Temp\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\"archive\""),
                            Newline,
                            StatementList {
                                RmDirStatement {
                                    Whitespace,
                                    RmDirKeyword,
                                    Whitespace,
                                    StringLiteral ("\"C:\\Archive\\Old\""),
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
                                RmDirStatement {
                                    Whitespace,
                                    RmDirKeyword,
                                    Whitespace,
                                    Identifier ("defaultPath"),
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
    fn rmdir_with_do_loop() {
        let source = r"
Sub Test()
    Dim dirExists As Boolean
    dirExists = True
    Do While dirExists
        On Error Resume Next
        RmDir tempDir
        dirExists = (Err.Number = 0)
    Loop
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("dirExists"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        BooleanKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("dirExists"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BooleanLiteralExpression {
                            TrueKeyword,
                        },
                        Newline,
                    },
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("dirExists"),
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
                            RmDirStatement {
                                Whitespace,
                                RmDirKeyword,
                                Whitespace,
                                Identifier ("tempDir"),
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("dirExists"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                ParenthesizedExpression {
                                    LeftParenthesis,
                                    BinaryExpression {
                                        MemberAccessExpression {
                                            Identifier ("Err"),
                                            PeriodOperator,
                                            Identifier ("Number"),
                                        },
                                        Whitespace,
                                        EqualityOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("0"),
                                        },
                                    },
                                    RightParenthesis,
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

    #[test]
    fn rmdir_with_chdir() {
        let source = r#"
Sub Test()
    ChDir "C:\"
    RmDir "C:\OldTemp"
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
                        StringLiteral ("\"C:\\\""),
                        Newline,
                    },
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\OldTemp\""),
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
    fn rmdir_conditional_removal() {
        let source = r#"
Sub Test()
    If Dir("C:\Temp", vbDirectory) <> "" Then
        RmDir "C:\Temp"
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
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Dir"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        StringLiteralExpression {
                                            StringLiteral ("\"C:\\Temp\""),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("vbDirectory"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            InequalityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"\""),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            RmDirStatement {
                                Whitespace,
                                RmDirKeyword,
                                Whitespace,
                                StringLiteral ("\"C:\\Temp\""),
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
    fn rmdir_with_line_continuation() {
        let source = r#"
Sub Test()
    RmDir _
        "C:\VeryLongPathName\SubFolder\TempData"
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        Underscore,
                        Newline,
                        Whitespace,
                        StringLiteral ("\"C:\\VeryLongPathName\\SubFolder\\TempData\""),
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
    fn rmdir_in_class_module() {
        let source = r"
Private Sub Class_Terminate()
    On Error Resume Next
    RmDir m_tempDirectory
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.cls", source).unwrap();
    }

    #[test]
    fn rmdir_with_app_path() {
        let source = r#"
Sub Test()
    RmDir App.Path & "\TempCache"
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        Identifier ("App"),
                        PeriodOperator,
                        Identifier ("Path"),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"\\TempCache\""),
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
    fn rmdir_recursive_cleanup() {
        let source = r#"
Sub RecursiveRemove()
    On Error Resume Next
    Kill "C:\Temp\Data\*.*"
    RmDir "C:\Temp\Data"
    Kill "C:\Temp\*.*"
    RmDir "C:\Temp"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("RecursiveRemove"),
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
                        StringLiteral ("\"C:\\Temp\\Data\\*.*\""),
                        Newline,
                    },
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Temp\\Data\""),
                        Newline,
                    },
                    KillStatement {
                        Whitespace,
                        KillKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Temp\\*.*\""),
                        Newline,
                    },
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Temp\""),
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
    fn rmdir_with_environ() {
        let source = r#"
Sub Test()
    RmDir Environ("TEMP") & "\MyApp"
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        Identifier ("Environ"),
                        LeftParenthesis,
                        StringLiteral ("\"TEMP\""),
                        RightParenthesis,
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"\\MyApp\""),
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
    fn rmdir_after_mkdir() {
        let source = r#"
Sub Test()
    MkDir "C:\TempWork"
    ' Do some work
    RmDir "C:\TempWork"
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
                    MkDirStatement {
                        Whitespace,
                        MkDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\TempWork\""),
                        Newline,
                    },
                    Whitespace,
                    EndOfLineComment,
                    Newline,
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\TempWork\""),
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
    fn rmdir_with_explicit_error_check() {
        let source = r"
Function RemoveDirectory(path As String) As Boolean
    On Error GoTo ErrorHandler
    RmDir path
    RemoveDirectory = True
    Exit Function
ErrorHandler:
    RemoveDirectory = False
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("RemoveDirectory(path As String) As Boolean"),
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        Identifier ("path"),
                        Newline,
                    },
                    Whitespace,
                    RemComment,
                    Newline,
                    ExitStatement {
                        Whitespace,
                        ExitKeyword,
                        Whitespace,
                        FunctionKeyword,
                        Newline,
                    },
                    LabelStatement {
                        Identifier ("ErrorHandler"),
                        ColonOperator,
                        Newline,
                    },
                    Whitespace,
                    RemComment,
                    Newline,
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn rmdir_forward_slash_path() {
        let source = r#"
Sub Test()
    RmDir "C:/Temp/Session"
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
                    RmDirStatement {
                        Whitespace,
                        RmDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:/Temp/Session\""),
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
