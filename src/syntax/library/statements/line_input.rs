use crate::language::Token;
use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    // VB6 Line Input statement syntax:
    // - Line Input #filenumber, varname
    //
    // Reads a single line from an open sequential file and assigns it to a String variable.
    //
    // The Line Input # statement syntax has these parts:
    //
    // | Part          | Description |
    // |---------------|-------------|
    // | filenumber    | Required. Any valid file number. |
    // | varname       | Required. Valid String or Variant variable name. |
    //
    // Remarks:
    // - Data read with Line Input # is usually written to a file with Print #.
    // - The Line Input # statement reads from a file one character at a time until it encounters
    //   a carriage return (Chr(13)) or carriage return–linefeed (Chr(13) + Chr(10)) sequence.
    // - Carriage return–linefeed sequences are skipped rather than appended to the character string.
    // - Line Input # is useful for reading text files that have been created in a text editor or
    //   with the Print # statement.
    // - Unlike Input #, Line Input # doesn't parse the data as it's read – you get the entire line as-is.
    // - If end of file is reached before reading a complete line, an error occurs.
    //
    // Examples:
    // ```vb
    // Line Input #1, textLine
    // Line Input #fileNum, dataBuffer
    // Line Input #1, myString
    // ```
    //
    // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/line-input-statement)
    pub(crate) fn parse_line_input_statement(&mut self) {
        // if we are now parsing a Line Input statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::LineInputStatement.to_raw());

        // Consume "Line" keyword
        self.consume_token();

        // Consume "Input" keyword (should be next)
        if self.at_token(Token::InputKeyword) {
            self.consume_token();
        }

        // Consume everything until newline
        // This includes: "#", filenumber, ",", varname
        while !self.is_at_end() && !self.at_token(Token::Newline) {
            self.consume_token();
        }

        // Consume the newline
        if self.at_token(Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // LineInputStatement
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*; // Line Input statement tests

    #[test]
    fn line_input_simple() {
        let source = r"
Sub Test()
    Line Input #1, textLine
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
                    LineInputStatement {
                        Whitespace,
                        LineKeyword,
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("textLine"),
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
    fn line_input_at_module_level() {
        let source = "Line Input #1, myLine\n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            LineInputStatement {
                LineKeyword,
                Whitespace,
                InputKeyword,
                Whitespace,
                Octothorpe,
                IntegerLiteral ("1"),
                Comma,
                Whitespace,
                Identifier ("myLine"),
                Newline,
            },
        ]);
        let debug = cst.debug_tree();
        assert!(debug.contains("LineInputStatement"));
    }

    #[test]
    fn line_input_with_file_variable() {
        let source = r"
Sub Test()
    Line Input #fileNum, buffer
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
                    LineInputStatement {
                        Whitespace,
                        LineKeyword,
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        Identifier ("fileNum"),
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
    fn line_input_preserves_whitespace() {
        let source = "    Line    Input    #1  ,  myString    \n";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Whitespace,
            LineInputStatement {
                LineKeyword,
                Whitespace,
                InputKeyword,
                Whitespace,
                Octothorpe,
                IntegerLiteral ("1"),
                Whitespace,
                Comma,
                Whitespace,
                Identifier ("myString"),
                Whitespace,
                Newline,
            },
        ]);
        let debug = cst.debug_tree();
        assert!(debug.contains("LineInputStatement"));
    }

    #[test]
    fn line_input_with_comment() {
        let source = r"
Sub Test()
    Line Input #1, textLine ' Read one line
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
                    LineInputStatement {
                        Whitespace,
                        LineKeyword,
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("textLine"),
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
    fn line_input_in_if_statement() {
        let source = r"
Sub Test()
    If Not EOF(1) Then
        Line Input #1, currentLine
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
                            LineInputStatement {
                                Whitespace,
                                LineKeyword,
                                Whitespace,
                                InputKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                Identifier ("currentLine"),
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
    fn line_input_inline_if() {
        let source = r"
Sub Test()
    If hasData Then Line Input #1, nextLine
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
                        LineInputStatement {
                            LineKeyword,
                            Whitespace,
                            InputKeyword,
                            Whitespace,
                            Octothorpe,
                            IntegerLiteral ("1"),
                            Comma,
                            Whitespace,
                            Identifier ("nextLine"),
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
    fn line_input_in_loop() {
        let source = r"
Sub Test()
    Do While Not EOF(1)
        Line Input #1, textLine
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
                            LineInputStatement {
                                Whitespace,
                                LineKeyword,
                                Whitespace,
                                InputKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                Identifier ("textLine"),
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
    fn multiple_line_input_statements() {
        let source = r"
Sub Test()
    Line Input #1, line1
    Line Input #1, line2
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
                    LineInputStatement {
                        Whitespace,
                        LineKeyword,
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("line1"),
                        Newline,
                    },
                    LineInputStatement {
                        Whitespace,
                        LineKeyword,
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("line2"),
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
    fn line_input_reading_text_file() {
        let source = r#"
Sub Test()
    Open "data.txt" For Input As #1
    Line Input #1, headerLine
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
                    LineInputStatement {
                        Whitespace,
                        LineKeyword,
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("headerLine"),
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
    fn line_input_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Line Input #1, textLine
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
                    LineInputStatement {
                        Whitespace,
                        LineKeyword,
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("textLine"),
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
    fn line_input_vs_input() {
        let source = r"
Sub Test()
    Line Input #1, wholeLine
    Input #1, parsedData
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
                    LineInputStatement {
                        Whitespace,
                        LineKeyword,
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("wholeLine"),
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
                        Identifier ("parsedData"),
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
    fn line_input_string_variable() {
        let source = r"
Sub Test()
    Dim myText As String
    Line Input #1, myText
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
                        Identifier ("myText"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    LineInputStatement {
                        Whitespace,
                        LineKeyword,
                        Whitespace,
                        InputKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("myText"),
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
