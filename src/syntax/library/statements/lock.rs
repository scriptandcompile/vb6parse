use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 Lock statement syntax:
    // - Lock [#]filenumber[, recordrange]
    //
    // Controls access to all or part of an open file.
    //
    // The Lock statement syntax has these parts:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | filenumber    | Required. Any valid file number. |
    // | recordrange   | Optional. Range of records to lock. Can be: record, start To end, or omitted for entire file. |
    //
    // Remarks:
    // - Lock and Unlock are used in environments where multiple processes might need access to the same file.
    // - Lock and Unlock statements are always used in pairs.
    // - The Lock statement locks all or part of a file opened using the Open statement.
    // - The first record or byte in a file is at position 1, the second at position 2, and so on.
    // - If you specify just one record number, only that record is locked.
    // - If you specify a range, all records in that range are locked.
    // - For files opened in Binary, Input, or Output mode, Lock always locks the entire file,
    //   regardless of the recordrange argument.
    // - For files opened in Random mode, Lock locks the specified record or range of records.
    // - Locked portions of a file can't be accessed by other processes until unlocked with Unlock.
    // - Use Unlock to remove the lock from a portion of a file.
    //
    // Examples:
    // ```vb
    // Lock #1
    // Lock #1, 5
    // Lock #1, 10 To 20
    // Lock fileNum, recordNum
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/lock-statement)
    pub(crate) fn parse_lock_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::LockStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // Lock statement tests

    #[test]
    fn lock_simple() {
        let source = r"
Sub Test()
    Lock #1
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
                    LockStatement {
                        Whitespace,
                        LockKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
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
    fn lock_at_module_level() {
        let source = "Lock #1\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            LockStatement {
                LockKeyword,
                Whitespace,
                Octothorpe,
                IntegerLiteral ("1"),
                Newline,
            },
        ]);
        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
    }

    #[test]
    fn lock_entire_file() {
        let source = r"
Sub Test()
    Lock #1
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
                    LockStatement {
                        Whitespace,
                        LockKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
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
    fn lock_single_record() {
        let source = r"
Sub Test()
    Lock #1, 5
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
                    LockStatement {
                        Whitespace,
                        LockKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("5"),
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
    fn lock_record_range() {
        let source = r"
Sub Test()
    Lock #1, 10 To 20
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
                    LockStatement {
                        Whitespace,
                        LockKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("10"),
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        IntegerLiteral ("20"),
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
    fn lock_with_variable() {
        let source = r"
Sub Test()
    Lock fileNum, recordNum
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
                    LockStatement {
                        Whitespace,
                        LockKeyword,
                        Whitespace,
                        Identifier ("fileNum"),
                        Comma,
                        Whitespace,
                        Identifier ("recordNum"),
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
    fn lock_preserves_whitespace() {
        let source = "    Lock    #1  ,  5    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            LockStatement {
                LockKeyword,
                Whitespace,
                Octothorpe,
                IntegerLiteral ("1"),
                Whitespace,
                Comma,
                Whitespace,
                IntegerLiteral ("5"),
                Whitespace,
                Newline,
            },
        ]);
        let debug = cst.debug_tree();
        assert!(debug.contains("LockStatement"));
    }

    #[test]
    fn lock_with_comment() {
        let source = r"
Sub Test()
    Lock #1, 5 ' Lock record 5
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
                    LockStatement {
                        Whitespace,
                        LockKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("5"),
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
    fn lock_in_if_statement() {
        let source = r"
Sub Test()
    If needsLock Then
        Lock #1, currentRecord
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
                            Identifier ("needsLock"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            LockStatement {
                                Whitespace,
                                LockKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                Identifier ("currentRecord"),
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
    fn lock_inline_if() {
        let source = r"
Sub Test()
    If multiUser Then Lock #1
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
                            Identifier ("multiUser"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        LockStatement {
                            LockKeyword,
                            Whitespace,
                            Octothorpe,
                            IntegerLiteral ("1"),
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
    fn lock_unlock_pair() {
        let source = r"
Sub Test()
    Lock #1, 5
    ' Do work
    Unlock #1, 5
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
                    LockStatement {
                        Whitespace,
                        LockKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("5"),
                        Newline,
                    },
                    Whitespace,
                    EndOfLineComment,
                    Newline,
                    UnlockStatement {
                        Whitespace,
                        UnlockKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("5"),
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
    fn lock_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Lock #1, recordNum
    If Err.Number <> 0 Then
        MsgBox "Could not lock record"
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
                    LockStatement {
                        Whitespace,
                        LockKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("recordNum"),
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
                                StringLiteral ("\"Could not lock record\""),
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
    fn lock_for_shared_file_access() {
        let source = r#"
Sub Test()
    Open "data.dat" For Random As #1 Shared
    Lock #1, 10
    ' Write to record 10
    Put #1, 10, myData
    Unlock #1, 10
    Close #1
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
                    OpenStatement {
                        Whitespace,
                        OpenKeyword,
                        Whitespace,
                        StringLiteral ("\"data.dat\""),
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        RandomKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Whitespace,
                        Identifier ("Shared"),
                        Newline,
                    },
                    LockStatement {
                        Whitespace,
                        LockKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("10"),
                        Newline,
                    },
                    Whitespace,
                    EndOfLineComment,
                    Newline,
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("10"),
                        Comma,
                        Whitespace,
                        Identifier ("myData"),
                        Newline,
                    },
                    UnlockStatement {
                        Whitespace,
                        UnlockKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("10"),
                        Newline,
                    },
                    CloseStatement {
                        Whitespace,
                        CloseKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
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
