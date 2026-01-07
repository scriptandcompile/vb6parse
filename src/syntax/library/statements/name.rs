//! # Name Statement
//!
//! Renames a disk file, directory, or folder.
//!
//! ## Syntax
//!
//! ```vb
//! Name oldpathname As newpathname
//! ```
//!
//! - `oldpathname`: Required. String expression that specifies the existing file name and location. May include directory or folder, and drive.
//! - `newpathname`: Required. String expression that specifies the new file name and location. May include directory or folder, and drive.
//!   Cannot specify a different drive from the one specified in `oldpathname`.
//!
//! ## Remarks
//!
//! - The `Name` statement renames a file and moves it to a different directory or folder, if necessary
//! - `Name` can move a file across directories or folders, but both `oldpathname` and `newpathname` must be on the same drive
//! - Using `Name` on an open file produces an error. You must close an open file before renaming it
//! - `Name` arguments can include relative or absolute paths
//! - The `Name` statement can also rename directories or folders
//! - If `newpathname` already exists, an error occurs
//! - Wildcard characters (* and ?) are not allowed in either `oldpathname` or `newpathname`
//!
//! ## Examples
//!
//! ```vb
//! ' Rename a file
//! Name "OLDFILE.TXT" As "NEWFILE.TXT"
//!
//! ' Move and rename a file
//! Name "C:\Data\Report.doc" As "C:\Archive\OldReport.doc"
//!
//! ' Rename a directory
//! Name "C:\OldFolder" As "C:\NewFolder"
//!
//! ' Move file to different directory (same drive)
//! Name "C:\Temp\Test.dat" As "C:\Data\Test.dat"
//!
//! ' Using variables
//! Dim oldName As String, newName As String
//! oldName = "File1.txt"
//! newName = "File2.txt"
//! Name oldName As newName
//! ```
//!
//! ## Reference
//!
//! [Name Statement - Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/name-statement)

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses a Name statement.
    pub(crate) fn parse_name_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::NameStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // Name statement tests

    #[test]
    fn name_simple() {
        let source = r#"
Sub Test()
    Name "OLDFILE.TXT" As "NEWFILE.TXT"
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
                    NameStatement {
                        Whitespace,
                        NameKeyword,
                        Whitespace,
                        StringLiteral ("\"OLDFILE.TXT\""),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringLiteral ("\"NEWFILE.TXT\""),
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
    fn name_at_module_level() {
        let source = r#"Name "old.txt" As "new.txt""#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            NameStatement {
                NameKeyword,
                Whitespace,
                StringLiteral ("\"old.txt\""),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringLiteral ("\"new.txt\""),
            },
        ]);
    }

    #[test]
    fn name_with_full_paths() {
        let source = r#"
Sub Test()
    Name "C:\Data\Report.doc" As "C:\Archive\OldReport.doc"
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
                    NameStatement {
                        Whitespace,
                        NameKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Data\\Report.doc\""),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Archive\\OldReport.doc\""),
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
    fn name_with_variables() {
        let source = r"
Sub Test()
    Name oldName As newName
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
                    NameStatement {
                        Whitespace,
                        NameKeyword,
                        Whitespace,
                        Identifier ("oldName"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Identifier ("newName"),
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
    fn name_preserves_whitespace() {
        let source = "    Name    \"old.txt\"    As    \"new.txt\"    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            NameStatement {
                NameKeyword,
                Whitespace,
                StringLiteral ("\"old.txt\""),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringLiteral ("\"new.txt\""),
                Whitespace,
                Newline,
            },
        ]);
    }

    #[test]
    fn name_with_comment() {
        let source = r#"
Sub Test()
    Name "temp.dat" As "backup.dat" ' Rename temp file
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
                    NameStatement {
                        Whitespace,
                        NameKeyword,
                        Whitespace,
                        StringLiteral ("\"temp.dat\""),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringLiteral ("\"backup.dat\""),
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
    fn name_in_if_statement() {
        let source = r#"
Sub Test()
    If fileExists Then
        Name "old.log" As "archive.log"
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
                            Identifier ("fileExists"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            NameStatement {
                                Whitespace,
                                NameKeyword,
                                Whitespace,
                                StringLiteral ("\"old.log\""),
                                Whitespace,
                                AsKeyword,
                                Whitespace,
                                StringLiteral ("\"archive.log\""),
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
    fn name_inline_if() {
        let source = r"
Sub Test()
    If needsRename Then Name oldFile As newFile
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
                            Identifier ("needsRename"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        NameStatement {
                            NameKeyword,
                            Whitespace,
                            Identifier ("oldFile"),
                            Whitespace,
                            AsKeyword,
                            Whitespace,
                            Identifier ("newFile"),
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
    fn multiple_name_statements() {
        let source = r#"
Sub RenameFiles()
    Name "File1.txt" As "Backup1.txt"
    Name "File2.txt" As "Backup2.txt"
    Name "File3.txt" As "Backup3.txt"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("RenameFiles"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    NameStatement {
                        Whitespace,
                        NameKeyword,
                        Whitespace,
                        StringLiteral ("\"File1.txt\""),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringLiteral ("\"Backup1.txt\""),
                        Newline,
                    },
                    NameStatement {
                        Whitespace,
                        NameKeyword,
                        Whitespace,
                        StringLiteral ("\"File2.txt\""),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringLiteral ("\"Backup2.txt\""),
                        Newline,
                    },
                    NameStatement {
                        Whitespace,
                        NameKeyword,
                        Whitespace,
                        StringLiteral ("\"File3.txt\""),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringLiteral ("\"Backup3.txt\""),
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
    fn name_rename_directory() {
        let source = r#"
Sub Test()
    Name "C:\OldFolder" As "C:\NewFolder"
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
                    NameStatement {
                        Whitespace,
                        NameKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\OldFolder\""),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\NewFolder\""),
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
    fn name_with_concatenation() {
        let source = r#"
Sub Test()
    Name basePath & "old.dat" As basePath & "new.dat"
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
                    NameStatement {
                        Whitespace,
                        NameKeyword,
                        Whitespace,
                        Identifier ("basePath"),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"old.dat\""),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Identifier ("basePath"),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"new.dat\""),
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
    fn name_move_file() {
        let source = r#"
Sub Test()
    Name "C:\Temp\Test.dat" As "C:\Data\Test.dat"
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
                    NameStatement {
                        Whitespace,
                        NameKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Temp\\Test.dat\""),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringLiteral ("\"C:\\Data\\Test.dat\""),
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
