//! # `MkDir` Statement
//!
//! Creates a new directory or folder.
//!
//! ## Syntax
//!
//! ```vb
//! MkDir path
//! ```
//!
//! - `path`: Required. String expression that identifies the directory or folder to be created. May include drive.
//!   If no drive is specified, `MkDir` creates the new directory or folder on the current drive.
//!
//! ## Remarks
//!
//! - An error occurs if you try to create a directory or folder that already exists
//! - The `path` argument can include absolute or relative paths
//! - You can use `MkDir` to create nested directories by creating parent directories first
//! - On Windows systems, both forward slashes (/) and backslashes (\) can be used as path separators
//! - The directory name can include the drive letter
//! - UNC paths are supported on network drives
//!
//! ## Examples
//!
//! ```vb
//! ' Create a directory in the current directory
//! MkDir "MyNewFolder"
//!
//! ' Create a directory with full path
//! MkDir "C:\Program Files\MyApp"
//!
//! ' Create a directory on another drive
//! MkDir "D:\Data\Reports"
//!
//! ' Create nested directories (parent must exist first)
//! MkDir "C:\Temp"
//! MkDir "C:\Temp\Logs"
//! MkDir "C:\Temp\Logs\Archive"
//!
//! ' Create directory on network drive
//! MkDir "\\Server\Share\NewFolder"
//! ```
//!
//! ## Reference
//!
//! [MkDir Statement - Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/mkdir-statement)

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses a `MkDir` statement.
    pub(crate) fn parse_mkdir_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::MkDirStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // MkDir statement tests

    #[test]
    fn mkdir_simple() {
        let source = r#"
Sub Test()
    MkDir "NewFolder"
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
                        StringLiteral ("\"NewFolder\""),
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
    fn mkdir_at_module_level() {
        let source = r#"MkDir "C:\Temp""#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            MkDirStatement {
                MkDirKeyword,
                Whitespace,
                StringLiteral ("\"C:\\Temp\""),
            },
        ]);
        let debug = cst.debug_tree();
        assert!(debug.contains("MkDirStatement"));
    }

    #[test]
    fn mkdir_with_full_path() {
        let source = r#"
Sub Test()
    MkDir "C:\Program Files\MyApp"
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
                        StringLiteral ("\"C:\\Program Files\\MyApp\""),
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
    fn mkdir_with_variable() {
        let source = r"
Sub Test()
    MkDir folderPath
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
                    MkDirStatement {
                        Whitespace,
                        MkDirKeyword,
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
    fn mkdir_preserves_whitespace() {
        let source = "    MkDir    \"MyFolder\"    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            MkDirStatement {
                MkDirKeyword,
                Whitespace,
                StringLiteral ("\"MyFolder\""),
                Whitespace,
                Newline,
            },
        ]);
        let debug = cst.debug_tree();
        assert!(debug.contains("MkDirStatement"));
    }

    #[test]
    fn mkdir_with_comment() {
        let source = r#"
Sub Test()
    MkDir "Logs" ' Create logs directory
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
                        StringLiteral ("\"Logs\""),
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
    fn mkdir_in_if_statement() {
        let source = r#"
Sub Test()
    If Not dirExists Then
        MkDir "C:\Data"
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
                        UnaryExpression {
                            NotKeyword,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("dirExists"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            MkDirStatement {
                                Whitespace,
                                MkDirKeyword,
                                Whitespace,
                                StringLiteral ("\"C:\\Data\""),
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
    fn mkdir_inline_if() {
        let source = r#"
Sub Test()
    If needsDir Then MkDir "Output"
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
                            Identifier ("needsDir"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        MkDirStatement {
                            MkDirKeyword,
                            Whitespace,
                            StringLiteral ("\"Output\""),
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
    fn multiple_mkdir_statements() {
        let source = r#"
Sub CreateDirs()
    MkDir "C:\Temp"
    MkDir "C:\Temp\Logs"
    MkDir "C:\Temp\Data"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("CreateDirs"),
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
                        StringLiteral ("\"C:\\Temp\""),
                        Newline,
                    },
                    MkDirStatement {
                        Whitespace,
                        MkDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Temp\\Logs\""),
                        Newline,
                    },
                    MkDirStatement {
                        Whitespace,
                        MkDirKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Temp\\Data\""),
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
    fn mkdir_with_concatenation() {
        let source = r#"
Sub Test()
    MkDir basePath & "\Subfolder"
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
                        Identifier ("basePath"),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"\\Subfolder\""),
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
    fn mkdir_unc_path() {
        let source = r#"
Sub Test()
    MkDir "\\Server\Share\NewFolder"
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
                        StringLiteral ("\"\\\\Server\\Share\\NewFolder\""),
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
    fn mkdir_with_function_call() {
        let source = r#"
Sub Test()
    MkDir App.Path & "\Data"
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
                        Identifier ("App"),
                        PeriodOperator,
                        Identifier ("Path"),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"\\Data\""),
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
