use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    /// Parse a Put statement.
    ///
    /// VB6 Put statement syntax:
    /// - Put [#]filenumber, [recnumber], varname
    ///
    /// Writes data from a variable to a disk file.
    ///
    /// The Put statement syntax has these parts:
    ///
    /// | Part          | Description |
    /// |---------------|-------------|
    /// | filenumber    | Required. Any valid file number. |
    /// | recnumber     | Optional. Variant (Long). Record number (Random mode files) or byte number (Binary mode files) at which writing begins. |
    /// | varname       | Required. Valid variable name containing data to be written to disk. |
    ///
    /// Remarks:
    /// - Put is used with files opened in Binary or Random mode.
    /// - For files opened in Random mode, the record length specified in the Open statement determines the number of bytes written.
    /// - For files opened in Binary mode, Put writes any number of bytes.
    /// - The first record or byte in a file is at position 1, the second at position 2, and so on.
    /// - If you omit recnumber, the next record or byte following the last Put or Get statement (or pointed to by the last Seek function) is written.
    /// - You must include delimiting commas, for example: Put #1, , myVariable
    /// - For files opened in Random mode, the following rules apply:
    ///   * If the length of the data being written is less than the length specified in the Len clause, subsequent records on disk are aligned on record-length boundaries.
    ///   * The space between the end of one record and the beginning of the next is padded with the existing file contents.
    ///   * If the variable being written is a variable-length string, Put writes a 2-byte descriptor containing the string length and then writes the string data.
    /// - For files opened in Binary mode, all the Random rules apply, except:
    ///   * The Len clause in the Open statement has no effect.
    ///   * Put writes the data contiguously, with no padding between records.
    /// - Put statements usually mirror Get statements. That is, data written with Put is typically read with Get.
    ///
    /// Examples:
    /// ```vb
    /// Put #1, , myRecord
    /// Put #1, recordNumber, customerData
    /// Put fileNum, , buffer
    /// Put #1, filePosition, userData
    /// ```
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/put-statement)
    pub(crate) fn parse_put_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::PutStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // Put statement tests

    #[test]
    fn put_simple() {
        let source = r"
Sub Test()
    Put #1, , myRecord
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Comma,
                        Whitespace,
                        Identifier ("myRecord"),
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
    fn put_at_module_level() {
        let source = "Put #1, , myData\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            PutStatement {
                PutKeyword,
                Whitespace,
                Octothorpe,
                IntegerLiteral ("1"),
                Comma,
                Whitespace,
                Comma,
                Whitespace,
                Identifier ("myData"),
                Newline,
            },
        ]);
        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_with_record_number() {
        let source = r"
Sub Test()
    Put #1, recordNumber, customerData
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("recordNumber"),
                        Comma,
                        Whitespace,
                        Identifier ("customerData"),
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
    fn put_with_file_variable() {
        let source = r"
Sub Test()
    Put fileNum, , buffer
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Identifier ("fileNum"),
                        Comma,
                        Whitespace,
                        Comma,
                        Whitespace,
                        Identifier ("buffer"),
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
    fn put_with_hash_symbol() {
        let source = r"
Sub Test()
    Put #fileNumber, position, data
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        Identifier ("fileNumber"),
                        Comma,
                        Whitespace,
                        Identifier ("position"),
                        Comma,
                        Whitespace,
                        Identifier ("data"),
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
    fn put_preserves_whitespace() {
        let source = "    Put    #1  ,  ,  myVar    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            PutStatement {
                PutKeyword,
                Whitespace,
                Octothorpe,
                IntegerLiteral ("1"),
                Whitespace,
                Comma,
                Whitespace,
                Comma,
                Whitespace,
                Identifier ("myVar"),
                Whitespace,
                Newline,
            },
        ]);
        let debug = cst.debug_tree();
        assert!(debug.contains("PutStatement"));
    }

    #[test]
    fn put_with_comment() {
        let source = r"
Sub Test()
    Put #1, , myRecord ' Write next record
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Comma,
                        Whitespace,
                        Identifier ("myRecord"),
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
    fn put_in_if_statement() {
        let source = r"
Sub Test()
    If dataReady Then
        Put #1, , myData
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
                            Identifier ("dataReady"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            PutStatement {
                                Whitespace,
                                PutKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                Comma,
                                Whitespace,
                                Identifier ("myData"),
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
    fn put_multiple_in_sequence() {
        let source = r"
Sub Test()
    Put #1, , record1
    Put #1, , record2
    Put #1, , record3
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Comma,
                        Whitespace,
                        Identifier ("record1"),
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
                        Comma,
                        Whitespace,
                        Identifier ("record2"),
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
                        Comma,
                        Whitespace,
                        Identifier ("record3"),
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
    fn put_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 10
        Put #1, , records(i)
    Next i
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
                            PutStatement {
                                Whitespace,
                                PutKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                Comma,
                                Whitespace,
                                Identifier ("records"),
                                LeftParenthesis,
                                Identifier ("i"),
                                RightParenthesis,
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
    fn put_with_udt() {
        let source = r"
Sub Test()
    Put #1, , employee
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Comma,
                        Whitespace,
                        Identifier ("employee"),
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
    fn put_binary_data() {
        let source = r"
Sub Test()
    Put #1, bytePosition, buffer()
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("bytePosition"),
                        Comma,
                        Whitespace,
                        Identifier ("buffer"),
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
    fn put_with_seek_position() {
        let source = r"
Sub Test()
    Put #1, Seek(1), myData
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        SeekKeyword,
                        LeftParenthesis,
                        IntegerLiteral ("1"),
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        Identifier ("myData"),
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
    fn put_inline_if() {
        let source = r"
Sub Test()
    If writeFlag Then Put #1, , record
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
                            Identifier ("writeFlag"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        PutStatement {
                            PutKeyword,
                            Whitespace,
                            Octothorpe,
                            IntegerLiteral ("1"),
                            Comma,
                            Whitespace,
                            Comma,
                            Whitespace,
                            Identifier ("record"),
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
    fn put_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Put #1, , myRecord
    If Err.Number <> 0 Then
        MsgBox "Error writing record"
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Comma,
                        Whitespace,
                        Identifier ("myRecord"),
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
                                StringLiteral ("\"Error writing record\""),
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
    fn put_after_get() {
        let source = r"
Sub Test()
    Get #1, recordNum, myRecord
    ' Modify the record
    Put #1, recordNum, myRecord
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
                    GetStatement {
                        Whitespace,
                        GetKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("recordNum"),
                        Comma,
                        Whitespace,
                        Identifier ("myRecord"),
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
                        Identifier ("myRecord"),
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
    fn put_with_explicit_position() {
        let source = r"
Sub Test()
    Put #1, 100, headerData
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("100"),
                        Comma,
                        Whitespace,
                        Identifier ("headerData"),
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
    fn put_with_calculated_position() {
        let source = r"
Sub Test()
    Put #1, (recordNum - 1) * recordLength + 1, myData
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        LeftParenthesis,
                        Identifier ("recordNum"),
                        Whitespace,
                        SubtractionOperator,
                        Whitespace,
                        IntegerLiteral ("1"),
                        RightParenthesis,
                        Whitespace,
                        MultiplicationOperator,
                        Whitespace,
                        Identifier ("recordLength"),
                        Whitespace,
                        AdditionOperator,
                        Whitespace,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("myData"),
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
    fn put_array_element() {
        let source = r"
Sub Test()
    Put #1, , dataArray(index)
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Comma,
                        Whitespace,
                        Identifier ("dataArray"),
                        LeftParenthesis,
                        Identifier ("index"),
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
    fn put_object_property() {
        let source = r"
Sub Test()
    Put #1, , myObject.Data
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Comma,
                        Whitespace,
                        Identifier ("myObject"),
                        PeriodOperator,
                        Identifier ("Data"),
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
    fn put_with_multiline_if() {
        let source = r"
Sub Test()
    If needsWrite Then
        Put #1, recordPos, recordData
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
                            Identifier ("needsWrite"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            PutStatement {
                                Whitespace,
                                PutKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                Identifier ("recordPos"),
                                Comma,
                                Whitespace,
                                Identifier ("recordData"),
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
    fn put_in_select_case() {
        let source = r"
Sub Test()
    Select Case recordType
        Case 1
            Put #1, , type1Record
        Case 2
            Put #1, , type2Record
    End Select
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
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("recordType"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                PutStatement {
                                    Whitespace,
                                    PutKeyword,
                                    Whitespace,
                                    Octothorpe,
                                    IntegerLiteral ("1"),
                                    Comma,
                                    Whitespace,
                                    Comma,
                                    Whitespace,
                                    Identifier ("type1Record"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("2"),
                            Newline,
                            StatementList {
                                PutStatement {
                                    Whitespace,
                                    PutKeyword,
                                    Whitespace,
                                    Octothorpe,
                                    IntegerLiteral ("1"),
                                    Comma,
                                    Whitespace,
                                    Comma,
                                    Whitespace,
                                    Identifier ("type2Record"),
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
    fn put_string_variable() {
        let source = r"
Sub Test()
    Put #1, , userName
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Comma,
                        Whitespace,
                        Identifier ("userName"),
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
    fn put_numeric_literal_position() {
        let source = r"
Sub Test()
    Put #1, 1, headerRecord
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
                    PutStatement {
                        Whitespace,
                        PutKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("headerRecord"),
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
    fn put_with_do_loop() {
        let source = r"
Sub Test()
    Do While Not EOF(1)
        Get #1, , inRecord
        Put #2, , outRecord
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
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        UnaryExpression {
                            NotKeyword,
                            Whitespace,
                            CallExpression {
                                Identifier ("EOF"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Newline,
                        StatementList {
                            GetStatement {
                                Whitespace,
                                GetKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                Comma,
                                Whitespace,
                                Identifier ("inRecord"),
                                Newline,
                            },
                            PutStatement {
                                Whitespace,
                                PutKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("2"),
                                Comma,
                                Whitespace,
                                Comma,
                                Whitespace,
                                Identifier ("outRecord"),
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
    fn put_random_access_file() {
        let source = r#"
Sub Test()
    Open "data.dat" For Random As #1 Len = Len(myRecord)
    Put #1, recordNumber, myRecord
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
                        LenKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        LenKeyword,
                        LeftParenthesis,
                        Identifier ("myRecord"),
                        RightParenthesis,
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
                        Identifier ("recordNumber"),
                        Comma,
                        Whitespace,
                        Identifier ("myRecord"),
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

    #[test]
    fn put_binary_file() {
        let source = r#"
Sub Test()
    Open "binary.bin" For Binary As #1
    Put #1, , byteArray()
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
                        StringLiteral ("\"binary.bin\""),
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        BinaryKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
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
                        Comma,
                        Whitespace,
                        Identifier ("byteArray"),
                        LeftParenthesis,
                        RightParenthesis,
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
