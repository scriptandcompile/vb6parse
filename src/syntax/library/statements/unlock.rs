use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 Unlock statement syntax:
    // - Unlock [#]filenumber[, recordrange]
    //
    // Removes access restrictions on all or part of an open file.
    //
    // The Unlock statement syntax has these parts:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | filenumber    | Required. Any valid file number. |
    // | recordrange   | Optional. Range of records to unlock. Can be: record, start To end, or omitted for entire file. |
    //
    // Remarks:
    // - Unlock is used to remove locks placed on a file with the Lock statement.
    // - The Unlock statement allows other processes to access the unlocked portions of the file.
    // - The arguments to Unlock must exactly match those used with the corresponding Lock statement.
    // - The first record or byte in a file is at position 1, the second at position 2, and so on.
    // - If you specify just one record number, only that record is unlocked.
    // - If you specify a range, all records in that range are unlocked.
    // - For files opened in Binary, Input, or Output mode, Unlock always unlocks the entire file,
    //   regardless of the recordrange argument.
    // - For files opened in Random mode, Unlock unlocks the specified record or range of records.
    // - Each Lock statement must have a corresponding Unlock statement with the same file number
    //   and record range.
    //
    // Examples:
    // ```vb
    // Unlock #1
    // Unlock #1, 5
    // Unlock #1, 10 To 20
    // Unlock fileNum, recordNum
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/unlock-statement)
    pub(crate) fn parse_unlock_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::UnlockStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // Unlock statement tests

    #[test]
    fn unlock_simple() {
        let source = r"
Sub Test()
    Unlock #1
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
                    UnlockStatement {
                        Whitespace,
                        UnlockKeyword,
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
    fn unlock_at_module_level() {
        let source = "Unlock #1\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            UnlockStatement {
                UnlockKeyword,
                Whitespace,
                Octothorpe,
                IntegerLiteral ("1"),
                Newline,
            },
        ]);
    }

    #[test]
    fn unlock_entire_file() {
        let source = r"
Sub Test()
    Unlock #1
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
                    UnlockStatement {
                        Whitespace,
                        UnlockKeyword,
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
    fn unlock_single_record() {
        let source = r"
Sub Test()
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
    fn unlock_record_range() {
        let source = r"
Sub Test()
    Unlock #1, 10 To 20
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
                    UnlockStatement {
                        Whitespace,
                        UnlockKeyword,
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
    fn unlock_with_variable() {
        let source = r"
Sub Test()
    Unlock fileNum, recordNum
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
                    UnlockStatement {
                        Whitespace,
                        UnlockKeyword,
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
    fn unlock_preserves_whitespace() {
        let source = "    Unlock    #1  ,  5    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            UnlockStatement {
                UnlockKeyword,
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
    }

    #[test]
    fn unlock_with_comment() {
        let source = r"
Sub Test()
    Unlock #1, 5 ' Release record lock
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
                    UnlockStatement {
                        Whitespace,
                        UnlockKeyword,
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
    fn unlock_in_if_statement() {
        let source = r"
Sub Test()
    If isDone Then
        Unlock #1, currentRecord
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
                            Identifier ("isDone"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            UnlockStatement {
                                Whitespace,
                                UnlockKeyword,
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
    fn unlock_inline_if() {
        let source = r"
Sub Test()
    If finished Then Unlock #1
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
                            Identifier ("finished"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        UnlockStatement {
                            UnlockKeyword,
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
    fn lock_unlock_matching_pair() {
        let source = r"
Sub Test()
    Lock #1, 5
    myData.Value = 100
    Put #1, 5, myData
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
                    AssignmentStatement {
                        MemberAccessExpression {
                            Identifier ("myData"),
                            PeriodOperator,
                            Identifier ("Value"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("100"),
                        },
                        Newline,
                    },
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("5"),
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
    fn unlock_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Unlock #1, recordNum
    If Err.Number <> 0 Then
        MsgBox "Could not unlock record"
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
                    UnlockStatement {
                        Whitespace,
                        UnlockKeyword,
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
                                StringLiteral ("\"Could not unlock record\""),
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
    fn unlock_in_finally_block() {
        let source = r#"
Sub Test()
    Lock #1, recordNum
    On Error GoTo ErrorHandler
    ' Do work
    Put #1, recordNum, myData
    Unlock #1, recordNum
    Exit Sub
ErrorHandler:
    Unlock #1, recordNum
    MsgBox "Error occurred"
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
                        Identifier ("recordNum"),
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
                        Identifier ("recordNum"),
                        Newline,
                    },
                    ExitStatement {
                        Whitespace,
                        ExitKeyword,
                        Whitespace,
                        SubKeyword,
                        Newline,
                    },
                    LabelStatement {
                        Identifier ("ErrorHandler"),
                        ColonOperator,
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
                        Identifier ("recordNum"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
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
