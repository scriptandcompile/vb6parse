//! # `LenB` Function
//!
//! The `LenB` function returns a `Long` containing the number of bytes used to represent a string in memory.
//! This function operates on byte count rather than character count, which is important when working with
//! ANSI strings, DBCS (Double-Byte Character Set), or when you need to know the actual memory footprint of a string.
//!
//! ## Syntax
//!
//! ```vb
//! LenB(string | varname)
//! ```
//!
//! ## Parameters
//!
//! - `string` - Any valid `String` expression.
//! - `varname` - Any valid variable name. If `varname` contains `Null`, `Null` is returned.
//!
//! ## Return Value
//!
//! Returns a `Long` specifying the number of bytes required to store the string or variable in memory.
//!
//! ## Behavior
//!
//! - For ANSI strings (single-byte character sets), `LenB` returns the same value as `Len`.
//! - For Unicode strings (VB6 default), `LenB` returns twice the value of `Len` because each Unicode character requires 2 bytes.
//! - For DBCS strings, the byte count depends on whether characters are single-byte or double-byte.
//! - If the argument is `Null`, `LenB` returns `Null`.
//! - When used with user-defined types, `LenB` returns the total byte size of the type.
//!
//! ## Difference from Len
//!
//! The `LenB` function returns the byte count, while the `Len` function returns the character count.
//! For single-byte character sets, they are identical. For Unicode (VB6's default string type),
//! `LenB` will return twice the value of `Len`.
//!
//! ## Examples
//!
//! ```vb
//! ' Get byte length of a string
//! Dim size As Long
//! size = LenB("Hello")  ' Returns 10 (5 characters * 2 bytes each in Unicode)
//!
//! ' Compare with character length
//! Dim charLen As Long
//! Dim byteLen As Long
//! charLen = Len("Test")   ' Returns 4
//! byteLen = LenB("Test")  ' Returns 8 (Unicode)
//!
//! ' Check memory size
//! Dim buffer As String
//! buffer = Space$(100)
//! Dim bufferSize As Long
//! bufferSize = LenB(buffer)  ' Returns 200 bytes
//! ```

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::parsers::cst::ConcreteSyntaxTree;
    #[test]
    fn lenb_simple() {
        let source = r#"
Sub Test()
    Dim size As Long
    size = LenB("Hello")
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("size"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("size"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LenB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Hello\""),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
    fn lenb_with_variable() {
        let source = r"
Sub Test()
    Dim text As String
    Dim size As Long
    size = LenB(text)
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
                        TextKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("size"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("size"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LenB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        TextKeyword,
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
    fn lenb_comparison() {
        let source = r"
Sub Test()
    Dim text As String
    If LenB(text) > 100 Then
        ' Do something
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        TextKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("LenB"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            TextKeyword,
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("100"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            EndOfLineComment,
                            Newline,
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
    fn lenb_buffer_size() {
        let source = r"
Sub Test()
    Dim buffer As String
    Dim bufferSize As Long
    buffer = Space$(256)
    bufferSize = LenB(buffer)
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
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("bufferSize"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("buffer"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Space$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("256"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("bufferSize"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LenB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("buffer"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
    fn lenb_in_expression() {
        let source = r"
Sub Test()
    Dim s As String
    Dim total As Long
    total = LenB(s) * 2 + 10
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
                        Identifier ("s"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("total"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("total"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("LenB"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("s"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                MultiplicationOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("2"),
                                },
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
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
    fn lenb_array_element() {
        let source = r"
Sub Test()
    Dim arr(10) As String
    Dim size As Long
    size = LenB(arr(5))
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
                        Identifier ("arr"),
                        LeftParenthesis,
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("size"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("size"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LenB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("arr"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("5"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
    fn lenb_loop_condition() {
        let source = r"
Sub Test()
    Dim data As String
    Dim i As Long
    For i = 1 To LenB(data) Step 2
        ' Process bytes
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("data"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
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
                        CallExpression {
                            Identifier ("LenB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("data"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        StepKeyword,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("2"),
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            EndOfLineComment,
                            Newline,
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
    fn lenb_validation() {
        let source = r"
Sub Test()
    Dim packet As String
    If LenB(packet) < 64 Then
        Exit Sub
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("packet"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("LenB"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("packet"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("64"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            ExitStatement {
                                Whitespace,
                                ExitKeyword,
                                Whitespace,
                                SubKeyword,
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
    fn lenb_with_midb() {
        let source = r"
Sub Test()
    Dim data As String
    Dim chunk As String
    Dim i As Long
    For i = 1 To LenB(data) Step 100
        chunk = MidB(data, i, 100)
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("data"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("chunk"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
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
                        CallExpression {
                            Identifier ("LenB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("data"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        StepKeyword,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("100"),
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("chunk"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    MidBKeyword,
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("data"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("i"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("100"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
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
    fn lenb_empty_check() {
        let source = r"
Sub Test()
    Dim s As String
    Dim isEmpty As Boolean
    isEmpty = (LenB(s) = 0)
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
                        Identifier ("s"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("isEmpty"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        BooleanKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("isEmpty"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("LenB"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("s"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("0"),
                                },
                            },
                            RightParenthesis,
                        },
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
    fn lenb_function_parameter() {
        let source = r"
Function GetByteSize(text As String) As Long
    GetByteSize = LenB(text)
End Function
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetByteSize"),
                ParameterList {
                    LeftParenthesis,
                },
                TextKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetByteSize"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LenB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        TextKeyword,
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn lenb_nested_function() {
        let source = r#"
Sub Test()
    Dim size As Long
    size = LenB(UCase$("hello"))
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("size"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("size"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LenB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("UCase$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                StringLiteralExpression {
                                                    StringLiteral ("\"hello\""),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
    fn lenb_select_case() {
        let source = r"
Sub Test()
    Dim data As String
    Select Case LenB(data)
        Case 0
            ' Empty
        Case Is > 1000
            ' Large
        Case Else
            ' Normal
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("data"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        CallExpression {
                            Identifier ("LenB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("data"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("0"),
                            Newline,
                            StatementList {
                                Whitespace,
                                EndOfLineComment,
                                Newline,
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IsKeyword,
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            IntegerLiteral ("1000"),
                            Newline,
                            StatementList {
                                Whitespace,
                                EndOfLineComment,
                                Newline,
                                Whitespace,
                            },
                        },
                        CaseElseClause {
                            CaseKeyword,
                            Whitespace,
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                EndOfLineComment,
                                Newline,
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
    fn lenb_binary_data() {
        let source = r"
Sub Test()
    Dim binaryData As String
    Dim headerSize As Long
    Dim payloadSize As Long
    headerSize = 16
    payloadSize = LenB(binaryData) - headerSize
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
                        Identifier ("binaryData"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("headerSize"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("payloadSize"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("headerSize"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("16"),
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("payloadSize"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("LenB"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("binaryData"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            SubtractionOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("headerSize"),
                            },
                        },
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
    fn lenb_redim_array() {
        let source = r"
Sub Test()
    Dim data As String
    Dim bytes() As Byte
    ReDim bytes(LenB(data) - 1)
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
                        Identifier ("data"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("bytes"),
                        LeftParenthesis,
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        ByteKeyword,
                        Newline,
                    },
                    ReDimStatement {
                        Whitespace,
                        ReDimKeyword,
                        Whitespace,
                        Identifier ("bytes"),
                        LeftParenthesis,
                        Identifier ("LenB"),
                        LeftParenthesis,
                        Identifier ("data"),
                        RightParenthesis,
                        Whitespace,
                        SubtractionOperator,
                        Whitespace,
                        IntegerLiteral ("1"),
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
    fn lenb_type_field() {
        let source = r"
Type BinaryRecord
    Data As String
    ID As String
End Type

Sub Test()
    Dim rec As BinaryRecord
    Dim size As Long
    size = LenB(rec.Data)
End Sub
";

        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            TypeStatement {
                TypeKeyword,
                Whitespace,
                Identifier ("BinaryRecord"),
                Newline,
                Whitespace,
                Identifier ("Data"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                Whitespace,
                Identifier ("ID"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                EndKeyword,
                Whitespace,
                TypeKeyword,
                Newline,
            },
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
                        Identifier ("rec"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Identifier ("BinaryRecord"),
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("size"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("size"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LenB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    MemberAccessExpression {
                                        Identifier ("rec"),
                                        PeriodOperator,
                                        Identifier ("Data"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
    fn lenb_concatenation_result() {
        let source = r"
Sub Test()
    Dim s1 As String
    Dim s2 As String
    Dim totalSize As Long
    totalSize = LenB(s1 & s2)
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
                        Identifier ("s1"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("s2"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("totalSize"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("totalSize"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LenB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("s1"),
                                        },
                                        Whitespace,
                                        Ampersand,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("s2"),
                                        },
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
    fn lenb_multiple_calls() {
        let source = r"
Sub Test()
    Dim a As String
    Dim b As String
    Dim result As Long
    result = LenB(a) + LenB(b)
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
                        Identifier ("a"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("b"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("result"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("LenB"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("a"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("LenB"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("b"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
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
    fn lenb_print_statement() {
        let source = r"
Sub Test()
    Dim text As String
    Debug.Print LenB(text)
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
                        TextKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("LenB"),
                        LeftParenthesis,
                        TextKeyword,
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
    fn lenb_division() {
        let source = r"
Sub Test()
    Dim unicodeStr As String
    Dim charCount As Long
    charCount = LenB(unicodeStr) / 2
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
                        Identifier ("unicodeStr"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("charCount"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("charCount"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("LenB"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("unicodeStr"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            DivisionOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("2"),
                            },
                        },
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
