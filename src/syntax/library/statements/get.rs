use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 Get statement syntax:
    // - Get [#]filenumber, [recnumber], varname
    //
    // Reads data from an open disk file into a variable.
    //
    // The Get statement syntax has these parts:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | filenumber    | Required. Any valid file number. |
    // | recnumber     | Optional. Variant (Long). Record number (Random mode files) or byte number (Binary mode files) at which reading begins. |
    // | varname       | Required. Valid variable name into which data is read. |
    //
    // Remarks:
    // - Get is used with files opened in Binary or Random mode.
    // - For files opened in Random mode, the record length specified in the Open statement determines the number of bytes read.
    // - For files opened in Binary mode, Get reads any number of bytes.
    // - The first record or byte in a file is at position 1, the second at position 2, and so on.
    // - If you omit recnumber, the next record or byte following the last Get or Put statement (or pointed to by the last Seek function) is read.
    // - You must include delimiting commas, for example: Get #1, , myVariable
    // - For files opened in Random mode, the following rules apply:
    //   * If the length of the data being read is less than the length specified in the Len clause, subsequent records on disk are aligned on record-length boundaries.
    //   * The space between the end of one record and the beginning of the next is padded with existing file contents.
    //   * If the variable being read is a variable-length string, Get reads a 2-byte descriptor containing the string length and then reads the string data.
    // - For files opened in Binary mode, all the Random rules apply, except:
    //   * The Len clause in the Open statement has no effect.
    //   * Get reads the data contiguously, with no padding between records.
    //
    // Examples:
    // ```vb
    // Get #1, , myRecord
    // Get #1, recordNumber, customerData
    // Get fileNum, , buffer
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/get-statement)
    pub(crate) fn parse_get_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::GetStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // Get statement tests

    #[test]
    fn get_simple() {
        let source = r"
Sub Test()
    Get #1, , myRecord
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
    fn get_at_module_level() {
        let source = "Get #1, , myData\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            GetStatement {
                GetKeyword,
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
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn get_with_record_number() {
        let source = r"
Sub Test()
    Get #1, recordNumber, customerData
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
    fn get_with_file_variable() {
        let source = r"
Sub Test()
    Get fileNum, , buffer
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
    fn get_with_hash_symbol() {
        let source = r"
Sub Test()
    Get #fileNumber, position, data
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
    fn get_preserves_whitespace() {
        let source = "    Get    #1  ,  ,  myVar    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            GetStatement {
                GetKeyword,
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
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn get_with_comment() {
        let source = r"
Sub Test()
    Get #1, , myRecord ' Read next record
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
    fn get_in_if_statement() {
        let source = r"
Sub Test()
    If Not EOF(1) Then
        Get #1, , myData
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
                        Whitespace,
                        ThenKeyword,
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
    fn get_inline_if() {
        let source = r"
Sub Test()
    If hasData Then Get #1, , buffer
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
                            Identifier ("hasData"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        GetStatement {
                            GetKeyword,
                            Whitespace,
                            Octothorpe,
                            IntegerLiteral ("1"),
                            Comma,
                            Whitespace,
                            Comma,
                            Whitespace,
                            Identifier ("buffer"),
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
    fn get_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Get #1, , myRecord
    If Err.Number <> 0 Then
        MsgBox "Error reading file"
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
                                StringLiteral ("\"Error reading file\""),
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
    fn get_in_loop() {
        let source = r"
Sub Test()
    Do While Not EOF(1)
        Get #1, , myRecord
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
                                Identifier ("myRecord"),
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
    fn multiple_get_statements() {
        let source = r"
Sub Test()
    Get #1, , record1
    Get #1, , record2
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
                        Comma,
                        Whitespace,
                        Identifier ("record1"),
                        Newline,
                    },
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
                        Identifier ("record2"),
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
    fn get_binary_mode() {
        let source = r"
Sub Test()
    Dim buffer As String * 512
    Get #1, , buffer
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
                        Identifier ("buffer"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Whitespace,
                        MultiplicationOperator,
                        Whitespace,
                        IntegerLiteral ("512"),
                        Newline,
                    },
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
}
