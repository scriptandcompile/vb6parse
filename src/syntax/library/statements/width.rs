//! # Width Statement
//!
//! Assigns an output line width to a file opened using the Open statement.
//!
//! ## Syntax
//!
//! ```vb
//! Width #filenumber, width
//! ```
//!
//! ## Parts
//!
//! - **filenumber**: Required. Any valid file number.
//! - **width**: Required. Numeric expression in the range 0–255, inclusive, that indicates how
//!   many characters appear on a line before a new line is started. If width equals 0, there is
//!   no limit to the length of a line. The default value for width is 0.
//!
//! ## Remarks
//!
//! - **Output Formatting**: The Width # statement is used with the Print # or Write # statements
//!   to control output formatting to files.
//! - **Line Length Control**: For files opened for sequential output, if the width of a line of
//!   output exceeds the value specified for width, a new line is automatically started.
//! - **No Effect on Input**: The Width # statement has no effect on files opened for input or
//!   binary access.
//! - **Zero Width**: Setting width to 0 means there is no line length limit, allowing continuous
//!   output without automatic line breaks.
//! - **Maximum Width**: The maximum width value is 255 characters.
//!
//! ## Examples
//!
//! ### Basic Width Setting
//!
//! ```vb
//! Open "output.txt" For Output As #1
//! Width #1, 80
//! Print #1, "This output will wrap at 80 characters"
//! Close #1
//! ```
//!
//! ### Set Unlimited Width
//!
//! ```vb
//! Open "data.csv" For Output As #2
//! Width #2, 0  ' No line length limit
//! Print #2, LongDataString
//! Close #2
//! ```
//!
//! ### Width with Multiple Files
//!
//! ```vb
//! Open "narrow.txt" For Output As #1
//! Open "wide.txt" For Output As #2
//! Width #1, 40
//! Width #2, 120
//! ```
//!
//! ### Dynamic Width Setting
//!
//! ```vb
//! Dim lineWidth As Integer
//! lineWidth = 80
//! Open "report.txt" For Output As #1
//! Width #1, lineWidth
//! ```
//!
//! ### Width for Formatted Output
//!
//! ```vb
//! Open "report.txt" For Output As #1
//! Width #1, 80
//! Print #1, Tab(10); "Header"
//! Print #1, Tab(10); String$(50, "-")
//! Close #1
//! ```
//!
//! ## Common Patterns
//!
//! ### Report Generation with Fixed Width
//!
//! ```vb
//! Sub GenerateReport()
//!     Open "report.txt" For Output As #1
//!     Width #1, 80
//!     
//!     Print #1, "Annual Sales Report"
//!     Print #1, String$(80, "=")
//!     ' ... report content ...
//!     
//!     Close #1
//! End Sub
//! ```
//!
//! ### CSV Export (No Width Limit)
//!
//! ```vb
//! Sub ExportCSV()
//!     Open "export.csv" For Output As #1
//!     Width #1, 0  ' Allow unlimited line length
//!     
//!     For i = 1 To RecordCount
//!         Print #1, BuildCSVLine(i)
//!     Next i
//!     
//!     Close #1
//! End Sub
//! ```
//!
//! ### Console-Style Output
//!
//! ```vb
//! Open "console.log" For Output As #1
//! Width #1, 80  ' Standard console width
//! Print #1, "System Log - "; Now()
//! Close #1
//! ```

use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    /// Parse a Width # statement.
    ///
    /// The Width # statement assigns an output line width to a file opened using the Open statement.
    ///
    /// Syntax:
    /// ```vb
    /// Width #filenumber, width
    /// ```
    ///
    /// Example:
    /// ```vb
    /// Width #1, 80
    /// ```
    pub(crate) fn parse_width_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::WidthStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn width_simple() {
        let source = r"
Sub Test()
    Width #1, 80
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80"),
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
    fn width_at_module_level() {
        let source = r"
Width #1, 80
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            WidthStatement {
                WidthKeyword,
                Whitespace,
                Octothorpe,
                IntegerLiteral ("1"),
                Comma,
                Whitespace,
                IntegerLiteral ("80"),
                Newline,
            },
        ]);
    }

    #[test]
    fn width_with_zero() {
        let source = r"
Sub Test()
    Width #1, 0
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
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
    fn width_with_variable() {
        let source = r"
Sub Test()
    Width #1, lineWidth
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("lineWidth"),
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
    fn width_with_expression() {
        let source = r"
Sub Test()
    Width #1, maxWidth * 2
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("maxWidth"),
                        Whitespace,
                        MultiplicationOperator,
                        Whitespace,
                        IntegerLiteral ("2"),
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
    fn width_with_file_number_variable() {
        let source = r"
Sub Test()
    Width #fileNum, 80
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        Identifier ("fileNum"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80"),
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
    fn width_max_value() {
        let source = r"
Sub Test()
    Width #1, 255
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("255"),
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
    fn width_with_comment() {
        let source = r"
Sub Test()
    Width #1, 80 ' Set standard console width
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80"),
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
    fn width_multiple_files() {
        let source = r"
Sub Test()
    Width #1, 80
    Width #2, 120
    Width #3, 0
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80"),
                        Newline,
                    },
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("2"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("120"),
                        Newline,
                    },
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("3"),
                        Comma,
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
    fn width_with_spaces() {
        let source = r"
Sub Test()
    Width  #1 ,  80
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Whitespace,
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80"),
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
    fn width_in_if_statement() {
        let source = r"
Sub Test()
    If openSuccess Then
        Width #1, 80
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
                            Identifier ("openSuccess"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            WidthStatement {
                                Whitespace,
                                WidthKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                IntegerLiteral ("80"),
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
    fn width_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 10
        Width #i, 80
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
                            WidthStatement {
                                Whitespace,
                                WidthKeyword,
                                Whitespace,
                                Octothorpe,
                                Identifier ("i"),
                                Comma,
                                Whitespace,
                                IntegerLiteral ("80"),
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
    fn width_after_open() {
        let source = r#"
Sub Test()
    Open "file.txt" For Output As #1
    Width #1, 80
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
                        StringLiteral ("\"file.txt\""),
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        OutputKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Newline,
                    },
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80"),
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
    fn width_before_print() {
        let source = r#"
Sub Test()
    Width #1, 80
    Print #1, "Output"
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80"),
                        Newline,
                    },
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Output\""),
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
    fn width_with_function_call() {
        let source = r"
Sub Test()
    Width #1, GetLineWidth()
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("GetLineWidth"),
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
    fn width_with_constant() {
        let source = r"
Sub Test()
    Width #1, MAX_WIDTH
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("MAX_WIDTH"),
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
    fn width_in_with_block() {
        let source = r"
Sub Test()
    With FileConfig
        Width #1, .LineWidth
    End With
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
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("FileConfig"),
                        Newline,
                        StatementList {
                            WidthStatement {
                                Whitespace,
                                WidthKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                PeriodOperator,
                                Identifier ("LineWidth"),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
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
    fn width_sequential_calls() {
        let source = r#"
Sub Test()
    Open "file1.txt" For Output As #1
    Width #1, 80
    Open "file2.txt" For Output As #2
    Width #2, 120
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
                        StringLiteral ("\"file1.txt\""),
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        OutputKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Newline,
                    },
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80"),
                        Newline,
                    },
                    OpenStatement {
                        Whitespace,
                        OpenKeyword,
                        Whitespace,
                        StringLiteral ("\"file2.txt\""),
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        OutputKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("2"),
                        Newline,
                    },
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("2"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("120"),
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
    fn width_with_parenthesized_file_number() {
        let source = r"
Sub Test()
    Width #(fileNum), 80
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        LeftParenthesis,
                        Identifier ("fileNum"),
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80"),
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
    fn width_with_calculated_width() {
        let source = r"
Sub Test()
    Width #1, screenWidth - marginLeft - marginRight
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("screenWidth"),
                        Whitespace,
                        SubtractionOperator,
                        Whitespace,
                        Identifier ("marginLeft"),
                        Whitespace,
                        SubtractionOperator,
                        Whitespace,
                        Identifier ("marginRight"),
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
    fn width_in_error_handler() {
        let source = r"
Sub Test()
    On Error Resume Next
    Width #1, 80
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80"),
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
    fn width_with_type_suffix() {
        let source = r"
Sub Test()
    Width #1, 80%
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80%"),
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
    fn width_with_line_continuation() {
        let source = r"
Sub Test()
    Width #1, _
        80
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Underscore,
                        Newline,
                        Whitespace,
                        IntegerLiteral ("80"),
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
    fn width_case_insensitive() {
        let source = r"
Sub Test()
    WIDTH #1, 80
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80"),
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
    fn width_standard_values() {
        let source = r"
Sub Test()
    Width #1, 40
    Width #2, 80
    Width #3, 132
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("40"),
                        Newline,
                    },
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("2"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80"),
                        Newline,
                    },
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("3"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("132"),
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
    fn width_with_file_freefile() {
        let source = r"
Sub Test()
    Dim fn As Integer
    fn = FreeFile
    Width #fn, 80
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
                        Identifier ("fn"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("fn"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("FreeFile"),
                        },
                        Newline,
                    },
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        Identifier ("fn"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80"),
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
    fn width_in_select_case() {
        let source = r"
Sub Test()
    Select Case outputType
        Case 1
            Width #1, 80
        Case 2
            Width #1, 132
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
                            Identifier ("outputType"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                WidthStatement {
                                    Whitespace,
                                    WidthKeyword,
                                    Whitespace,
                                    Octothorpe,
                                    IntegerLiteral ("1"),
                                    Comma,
                                    Whitespace,
                                    IntegerLiteral ("80"),
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
                                WidthStatement {
                                    Whitespace,
                                    WidthKeyword,
                                    Whitespace,
                                    Octothorpe,
                                    IntegerLiteral ("1"),
                                    Comma,
                                    Whitespace,
                                    IntegerLiteral ("132"),
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
    fn width_preserves_formatting() {
        let source = r"
Sub Test()
    Width    #1   ,    80
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
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Whitespace,
                        Comma,
                        Whitespace,
                        IntegerLiteral ("80"),
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
    fn width_in_nested_control_structures() {
        let source = r"
Sub Test()
    If fileOpen Then
        For i = 1 To 10
            Width #i, 80
        Next i
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
                            Identifier ("fileOpen"),
                        },
                        Whitespace,
                        ThenKeyword,
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
                                    WidthStatement {
                                        Whitespace,
                                        WidthKeyword,
                                        Whitespace,
                                        Octothorpe,
                                        Identifier ("i"),
                                        Comma,
                                        Whitespace,
                                        IntegerLiteral ("80"),
                                        Newline,
                                    },
                                    Whitespace,
                                },
                                NextKeyword,
                                Whitespace,
                                Identifier ("i"),
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
}
