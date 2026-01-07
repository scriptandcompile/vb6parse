use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 Input statement syntax:
    // - Input #filenumber, varlist
    //
    // Reads data from an open sequential file and assigns the data to variables.
    //
    // The Input # statement syntax has these parts:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | filenumber    | Required. Any valid file number. |
    // | varlist       | Required. Comma-delimited list of variables that are assigned values read from the file. Variables can't be arrays or object variables. However, variables that describe an element of an array or user-defined type may be used. |
    //
    // Remarks:
    // - Data read with Input # is usually written to a file with Write #.
    // - Use this statement only with files opened in Input or Binary mode.
    // - The Input # statement reads data items from a sequential file and assigns them to variables.
    // - Data items in the file must appear in the same order as the variables in varlist and be separated by commas.
    // - If the data item to be read is a quoted string, Input # strips the quotation marks.
    // - Input # is typically used to read data that was written to a file using the Write # statement.
    // - For files opened for Binary access, Input # reads all the bytes it needs to complete the varlist.
    // - If end of file is reached before all variables are filled, an error occurs.
    //
    // Examples:
    // ```vb
    // Input #1, name, age
    // Input #fileNum, x, y, z
    // Input #1, firstName, lastName, address
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/input-statement)
    pub(crate) fn parse_input_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::InputStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // Input statement tests

    #[test]
    fn input_simple() {
        let source = r"
Sub Test()
    Input #1, name, age
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
                    InputStatement {
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        NameKeyword,
                        Comma,
                        Whitespace,
                        Identifier ("age"),
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
    fn input_at_module_level() {
        let source = "Input #1, myData\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            InputStatement {
                InputKeyword,
                Whitespace,
                Octothorpe,
                IntegerLiteral ("1"),
                Comma,
                Whitespace,
                Identifier ("myData"),
                Newline,
            },
        ]);
    }

    #[test]
    fn input_multiple_variables() {
        let source = r"
Sub Test()
    Input #1, firstName, lastName, age, address
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
                    InputStatement {
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("firstName"),
                        Comma,
                        Whitespace,
                        Identifier ("lastName"),
                        Comma,
                        Whitespace,
                        Identifier ("age"),
                        Comma,
                        Whitespace,
                        Identifier ("address"),
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
    fn input_with_file_variable() {
        let source = r"
Sub Test()
    Input #fileNum, x, y, z
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
                    InputStatement {
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        Identifier ("fileNum"),
                        Comma,
                        Whitespace,
                        Identifier ("x"),
                        Comma,
                        Whitespace,
                        Identifier ("y"),
                        Comma,
                        Whitespace,
                        Identifier ("z"),
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
    fn input_preserves_whitespace() {
        let source = "    Input    #1  ,  name  ,  age    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            InputStatement {
                InputKeyword,
                Whitespace,
                Octothorpe,
                IntegerLiteral ("1"),
                Whitespace,
                Comma,
                Whitespace,
                NameKeyword,
                Whitespace,
                Comma,
                Whitespace,
                Identifier ("age"),
                Whitespace,
                Newline,
            },
        ]);
    }

    #[test]
    fn input_with_comment() {
        let source = r"
Sub Test()
    Input #1, name, age ' Read person data
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
                    InputStatement {
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        NameKeyword,
                        Comma,
                        Whitespace,
                        Identifier ("age"),
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
    fn input_in_if_statement() {
        let source = r"
Sub Test()
    If Not EOF(1) Then
        Input #1, myData
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
                            InputStatement {
                                Whitespace,
                                InputKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
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
    fn input_inline_if() {
        let source = r"
Sub Test()
    If hasData Then Input #1, buffer
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
                        InputStatement {
                            InputKeyword,
                            Whitespace,
                            Octothorpe,
                            IntegerLiteral ("1"),
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
    fn input_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Input #1, name, age
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
                    InputStatement {
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        NameKeyword,
                        Comma,
                        Whitespace,
                        Identifier ("age"),
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
    fn input_in_loop() {
        let source = r"
Sub Test()
    Do While Not EOF(1)
        Input #1, myRecord
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
                            InputStatement {
                                Whitespace,
                                InputKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
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
    fn multiple_input_statements() {
        let source = r"
Sub Test()
    Input #1, header
    Input #1, data
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
                    InputStatement {
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("header"),
                        Newline,
                    },
                    InputStatement {
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
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
    fn input_sequential_file() {
        let source = r#"
Sub Test()
    Open "data.txt" For Input As #1
    Input #1, name, age, city
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
                        StringLiteral ("\"data.txt\""),
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Newline,
                    },
                    InputStatement {
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        NameKeyword,
                        Comma,
                        Whitespace,
                        Identifier ("age"),
                        Comma,
                        Whitespace,
                        Identifier ("city"),
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
