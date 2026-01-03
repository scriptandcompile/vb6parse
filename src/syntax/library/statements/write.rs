//! # Write Statement
//!
//! Writes data to a sequential file.
//!
//! ## Syntax
//!
//! ```vb
//! Write #filenumber, [outputlist]
//! ```
//!
//! ## Parts
//!
//! - **filenumber**: Required. Any valid file number.
//! - **outputlist**: Optional. One or more comma-delimited numeric expressions or string expressions
//!   to write to a file.
//!
//! ## Remarks
//!
//! - **Data Formatting**: Data written with Write # is usually read from a file with Input #.
//! - **Delimiters**: The Write # statement inserts commas between items and quotation marks around
//!   strings as they are written to the file. You don't have to put explicit delimiters in the list.
//! - **Universal Data**: Write # writes data in a universal format that can be read by Input # regardless
//!   of the locale settings.
//! - **Numeric Data**: Numeric data is written with a period (.) as the decimal separator.
//! - **Boolean Values**: Boolean data is written as #TRUE# or #FALSE#.
//! - **Date Values**: Date data is written using the universal date format: #yyyy-mm-dd hh:mm:ss#
//! - **Empty Values**: If outputlist data is Empty, nothing is written. However, if outputlist data is
//!   Null, #NULL# is written.
//! - **Error Data**: Error values are written as #ERROR errorcode#. The number sign (#) ensures the keyword
//!   is not confused with a variable name.
//! - **Comparison with Print #**: Unlike Print #, Write # inserts commas between items and quotes around
//!   strings automatically.
//!
//! ## Examples
//!
//! ### Write Simple Data
//!
//! ```vb
//! Open "test.txt" For Output As #1
//! Write #1, "Hello", 42, True
//! Close #1
//! ' File contents: "Hello",42,#TRUE#
//! ```
//!
//! ### Write Multiple Lines
//!
//! ```vb
//! Open "data.txt" For Output As #1
//! For i = 1 To 10
//!     Write #1, i, i * i, i * i * i
//! Next i
//! Close #1
//! ```
//!
//! ### Write Mixed Data Types
//!
//! ```vb
//! Open "record.txt" For Output As #1
//! Write #1, "John Doe", 30, #1/1/1995#, True
//! Close #1
//! ```
//!
//! ### Write Without Data (New Line)
//!
//! ```vb
//! Open "output.txt" For Output As #1
//! Write #1, "First line"
//! Write #1
//! Write #1, "Third line"
//! Close #1
//! ```
//!
//! ### Write Null and Empty
//!
//! ```vb
//! Open "test.txt" For Output As #1
//! Write #1, Null, Empty, "data"
//! Close #1
//! ' File contents: #NULL#,,"data"
//! ```
//!
//! ### Write Error Values
//!
//! ```vb
//! Open "errors.txt" For Output As #1
//! Write #1, CVErr(2007)
//! Close #1
//! ' File contents: #ERROR 2007#
//! ```
//!
//! ## Common Patterns
//!
//! ### Export Data to CSV-like Format
//!
//! ```vb
//! Sub ExportData()
//!     Open "export.txt" For Output As #1
//!     
//!     ' Write header
//!     Write #1, "Name", "Age", "City"
//!     
//!     ' Write data rows
//!     For i = 0 To UBound(employees)
//!         Write #1, employees(i).Name, employees(i).Age, employees(i).City
//!     Next i
//!     
//!     Close #1
//! End Sub
//! ```
//!
//! ### Write Database Records
//!
//! ```vb
//! Sub SaveRecords()
//!     Open "records.dat" For Output As #1
//!     
//!     Do Until rs.EOF
//!         Write #1, rs!ID, rs!Name, rs!Date, rs!Active
//!         rs.MoveNext
//!     Loop
//!     
//!     Close #1
//! End Sub
//! ```
//!
//! ### Write Configuration Data
//!
//! ```vb
//! Sub SaveConfig()
//!     Open "config.dat" For Output As #1
//!     Write #1, appName, version, lastRun, isRegistered
//!     Close #1
//! End Sub
//! ```
//!
//! ### Write Array Data
//!
//! ```vb
//! Sub WriteArray()
//!     Open "array.dat" For Output As #1
//!     
//!     For i = LBound(data) To UBound(data)
//!         Write #1, data(i)
//!     Next i
//!     
//!     Close #1
//! End Sub
//! ```
//!
//! ### Append Data to Existing File
//!
//! ```vb
//! Sub AppendRecord()
//!     Open "log.txt" For Append As #1
//!     Write #1, Now(), userName, action, details
//!     Close #1
//! End Sub
//! ```

use crate::parsers::SyntaxKind;

use crate::parsers::cst::Parser;

impl Parser<'_> {
    /// Parse a Write # statement.
    ///
    /// The Write # statement writes data to a sequential file with automatic
    /// formatting: commas between items and quotation marks around strings.
    ///
    /// Syntax:
    /// ```vb
    /// Write #filenumber, [outputlist]
    /// ```
    ///
    /// Example:
    /// ```vb
    /// Write #1, "Hello", 42, True
    /// ```
    pub(crate) fn parse_write_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::WriteStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn write_simple() {
        let source = r#"
Sub Test()
    Write #1, "Hello"
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Hello\""),
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
    fn write_at_module_level() {
        let source = r#"
Write #1, "data"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            WriteStatement {
                WriteKeyword,
                Whitespace,
                Octothorpe,
                IntegerLiteral ("1"),
                Comma,
                Whitespace,
                StringLiteral ("\"data\""),
                Newline,
            },
        ]);
    }

    #[test]
    fn write_multiple_values() {
        let source = r#"
Sub Test()
    Write #1, "Name", 42, True
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Name\""),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("42"),
                        Comma,
                        Whitespace,
                        TrueKeyword,
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
    fn write_no_data() {
        let source = r"
Sub Test()
    Write #1
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
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
    fn write_with_variables() {
        let source = r"
Sub Test()
    Write #1, name, age, city
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
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
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn write_with_expressions() {
        let source = r"
Sub Test()
    Write #1, x + y, total * 2
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("x"),
                        Whitespace,
                        AdditionOperator,
                        Whitespace,
                        Identifier ("y"),
                        Comma,
                        Whitespace,
                        Identifier ("total"),
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
    fn write_with_file_number_variable() {
        let source = r"
Sub Test()
    Write #fileNum, data
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        Identifier ("fileNum"),
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
    fn write_with_comment() {
        let source = r"
Sub Test()
    Write #1, data ' Write data to file
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("data"),
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
    fn write_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 10
        Write #1, i, i * i
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
                            WriteStatement {
                                Whitespace,
                                WriteKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                Identifier ("i"),
                                Comma,
                                Whitespace,
                                Identifier ("i"),
                                Whitespace,
                                MultiplicationOperator,
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
    fn write_with_string_literal() {
        let source = r#"
Sub Test()
    Write #1, "Hello, World!", "Data"
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Hello, World!\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Data\""),
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
    fn write_with_numeric_literals() {
        let source = r"
Sub Test()
    Write #1, 42, 3.14, -100
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("42"),
                        Comma,
                        Whitespace,
                        SingleLiteral,
                        Comma,
                        Whitespace,
                        SubtractionOperator,
                        IntegerLiteral ("100"),
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
    fn write_with_boolean() {
        let source = r"
Sub Test()
    Write #1, True, False
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        TrueKeyword,
                        Comma,
                        Whitespace,
                        FalseKeyword,
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
    fn write_with_date() {
        let source = r"
Sub Test()
    Write #1, #1/1/2025#
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        DateLiteral ("#1, #"),
                        IntegerLiteral ("1"),
                        DivisionOperator,
                        IntegerLiteral ("1"),
                        DivisionOperator,
                        DoubleLiteral,
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
    fn write_with_null() {
        let source = r"
Sub Test()
    Write #1, Null, Empty
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        NullKeyword,
                        Comma,
                        Whitespace,
                        EmptyKeyword,
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
    fn write_with_object_property() {
        let source = r"
Sub Test()
    Write #1, obj.Name, obj.Value
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("obj"),
                        PeriodOperator,
                        NameKeyword,
                        Comma,
                        Whitespace,
                        Identifier ("obj"),
                        PeriodOperator,
                        Identifier ("Value"),
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
    fn write_with_array_access() {
        let source = r"
Sub Test()
    Write #1, arr(i), arr(j)
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("arr"),
                        LeftParenthesis,
                        Identifier ("i"),
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        Identifier ("arr"),
                        LeftParenthesis,
                        Identifier ("j"),
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
    fn write_with_function_call() {
        let source = r"
Sub Test()
    Write #1, GetValue(), ProcessData()
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("GetValue"),
                        LeftParenthesis,
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        Identifier ("ProcessData"),
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
    fn write_multiple_statements() {
        let source = r#"
Sub Test()
    Write #1, "Line 1"
    Write #1, "Line 2"
    Write #1, "Line 3"
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Line 1\""),
                        Newline,
                    },
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Line 2\""),
                        Newline,
                    },
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Line 3\""),
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
    fn write_in_if_statement() {
        let source = r"
Sub Test()
    If condition Then
        Write #1, data
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
                            Identifier ("condition"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            WriteStatement {
                                Whitespace,
                                WriteKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                Identifier ("data"),
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
    fn write_in_do_loop() {
        let source = r"
Sub Test()
    Do Until EOF(1)
        Write #2, currentRecord
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
                        UntilKeyword,
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
                        Newline,
                        StatementList {
                            WriteStatement {
                                Whitespace,
                                WriteKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("2"),
                                Comma,
                                Whitespace,
                                Identifier ("currentRecord"),
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
    fn write_with_recordset() {
        let source = r"
Sub Test()
    Write #1, rs!Name, rs!Age
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("rs"),
                        ExclamationMark,
                        NameKeyword,
                        Comma,
                        Whitespace,
                        Identifier ("rs"),
                        ExclamationMark,
                        Identifier ("Age"),
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
    fn write_preserves_whitespace() {
        let source = r"
Sub Test()
    Write  #1 ,  data1 ,  data2
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Whitespace,
                        Comma,
                        Whitespace,
                        Identifier ("data1"),
                        Whitespace,
                        Comma,
                        Whitespace,
                        Identifier ("data2"),
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
    fn write_with_line_continuation() {
        let source = r"
Sub Test()
    Write #1, _
        field1, _
        field2
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Underscore,
                        Newline,
                        Whitespace,
                        Identifier ("field1"),
                        Comma,
                        Whitespace,
                        Underscore,
                        Newline,
                        Whitespace,
                        Identifier ("field2"),
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
    fn write_in_select_case() {
        let source = r#"
Sub Test()
    Select Case recordType
        Case 1
            Write #1, "Type A", data
        Case 2
            Write #1, "Type B", data
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
                                WriteStatement {
                                    Whitespace,
                                    WriteKeyword,
                                    Whitespace,
                                    Octothorpe,
                                    IntegerLiteral ("1"),
                                    Comma,
                                    Whitespace,
                                    StringLiteral ("\"Type A\""),
                                    Comma,
                                    Whitespace,
                                    Identifier ("data"),
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
                                WriteStatement {
                                    Whitespace,
                                    WriteKeyword,
                                    Whitespace,
                                    Octothorpe,
                                    IntegerLiteral ("1"),
                                    Comma,
                                    Whitespace,
                                    StringLiteral ("\"Type B\""),
                                    Comma,
                                    Whitespace,
                                    Identifier ("data"),
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
    fn write_with_now_function() {
        let source = r"
Sub Test()
    Write #1, Now(), data
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("Now"),
                        LeftParenthesis,
                        RightParenthesis,
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
    fn write_in_with_block() {
        let source = r"
Sub Test()
    With record
        Write #1, .Name, .Value
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
                        Identifier ("record"),
                        Newline,
                        StatementList {
                            WriteStatement {
                                Whitespace,
                                WriteKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                PeriodOperator,
                                NameKeyword,
                                Comma,
                                Whitespace,
                                PeriodOperator,
                                Identifier ("Value"),
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
    fn write_case_insensitive() {
        let source = r"
Sub Test()
    WRITE #1, data
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
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
    fn write_in_error_handler() {
        let source = r"
Sub Test()
    On Error Resume Next
    Write #1, errorData
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("errorData"),
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
    fn write_with_freefile() {
        let source = r"
Sub Test()
    Dim fn As Integer
    fn = FreeFile
    Write #fn, data
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
                    WriteStatement {
                        Whitespace,
                        WriteKeyword,
                        Whitespace,
                        Octothorpe,
                        Identifier ("fn"),
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
    fn write_sequential_values() {
        let source = r"
Sub Test()
    For i = 1 To 100
        Write #1, i, i * 2, i ^ 2
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
                            IntegerLiteral ("100"),
                        },
                        Newline,
                        StatementList {
                            WriteStatement {
                                Whitespace,
                                WriteKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                Identifier ("i"),
                                Comma,
                                Whitespace,
                                Identifier ("i"),
                                Whitespace,
                                MultiplicationOperator,
                                Whitespace,
                                IntegerLiteral ("2"),
                                Comma,
                                Whitespace,
                                Identifier ("i"),
                                Whitespace,
                                ExponentiationOperator,
                                Whitespace,
                                IntegerLiteral ("2"),
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
}
