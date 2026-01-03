//! # `Abs` Function
//!
//! Returns the absolute value of a number.
//!
//! ## Syntax
//!
//! ```vb
//! Abs(number)
//! ```
//!
//! ## Parts
//!
//! - **number**: Required. Any valid numeric expression. If number contains Null, Null is returned;
//!   if it is an uninitialized variable, zero is returned.
//!
//! ## Return Value
//!
//! The return type is the same as the input type, except:
//! - If number is a Variant containing Null, returns Null
//! - If number is an uninitialized Variant, returns 0
//! - The absolute value is always non-negative (>= 0)
//!
//! ## Remarks
//!
//! - **Absolute Value**: The absolute value of a number is its unsigned magnitude. For example,
//!   Abs(-1) and Abs(1) both return 1.
//! - **Type Preservation**: The return type matches the input type. If you pass an Integer, you
//!   get an Integer back. If you pass a Double, you get a Double back.
//! - **Null Handling**: If the argument is Null, the function returns Null.
//! - **Overflow**: For the most negative value of Integer (-32768) or Long (-2147483648), Abs
//!   will cause an overflow error because the positive equivalent is outside the valid range.
//! - **Performance**: Abs is highly optimized and very fast for numeric operations.
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! Dim result As Integer
//! result = Abs(-50)
//! ' result = 50
//! ```
//!
//! ### With Positive Numbers
//!
//! ```vb
//! Dim value As Integer
//! value = Abs(25)
//! ' value = 25 (unchanged)
//! ```
//!
//! ### With Floating Point
//!
//! ```vb
//! Dim distance As Double
//! distance = Abs(-12.75)
//! ' distance = 12.75
//! ```
//!
//! ### With Zero
//!
//! ```vb
//! Dim zero As Integer
//! zero = Abs(0)
//! ' zero = 0
//! ```
//!
//! ### With Expressions
//!
//! ```vb
//! Dim x As Integer, y As Integer
//! x = 10
//! y = 20
//! Dim difference As Integer
//! difference = Abs(x - y)
//! ' difference = 10
//! ```
//!
//! ### Calculating Distance
//!
//! ```vb
//! Function Distance(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
//!     Distance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
//! End Function
//!
//! ' Often used with Abs for 1D distance:
//! Dim dist As Double
//! dist = Abs(x2 - x1)
//! ```
//!
//! ### With Currency
//!
//! ```vb
//! Dim amount As Currency
//! amount = Abs(-1234.56@)
//! ' amount = 1234.56
//! ```
//!
//! ### With Variants
//!
//! ```vb
//! Dim v As Variant
//! v = -42
//! Dim result As Variant
//! result = Abs(v)
//! ' result = 42
//! ```
//!
//! ## Common Patterns
//!
//! ### Ensuring Positive Values
//!
//! ```vb
//! Sub ProcessValue(ByVal input As Integer)
//!     Dim positiveInput As Integer
//!     positiveInput = Abs(input)
//!     ' Always work with positive values
//!     DoSomething positiveInput
//! End Sub
//! ```
//!
//! ### Calculating Difference
//!
//! ```vb
//! Function GetDifference(a As Long, b As Long) As Long
//!     GetDifference = Abs(a - b)
//! End Function
//! ```
//!
//! ### Data Validation
//!
//! ```vb
//! Function IsWithinTolerance(actual As Double, expected As Double, tolerance As Double) As Boolean
//!     IsWithinTolerance = (Abs(actual - expected) <= tolerance)
//! End Function
//! ```
//!
//! ### Financial Calculations
//!
//! ```vb
//! Function CalculateVariance(actual As Currency, budget As Currency) As Currency
//!     CalculateVariance = Abs(actual - budget)
//! End Function
//! ```
//!
//! ### Array Processing
//!
//! ```vb
//! Sub MakeArrayPositive(arr() As Integer)
//!     Dim i As Integer
//!     For i = LBound(arr) To UBound(arr)
//!         arr(i) = Abs(arr(i))
//!     Next i
//! End Sub
//! ```
//!
//! ### Comparison Logic
//!
//! ```vb
//! Function MaxAbsValue(a As Double, b As Double) As Double
//!     If Abs(a) > Abs(b) Then
//!         MaxAbsValue = Abs(a)
//!     Else
//!         MaxAbsValue = Abs(b)
//!     End If
//! End Function
//! ```
//!
//! ### Coordinate Systems
//!
//! ```vb
//! Function ManhattanDistance(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer) As Integer
//!     ManhattanDistance = Abs(x2 - x1) + Abs(y2 - y1)
//! End Function
//! ```
//!
//! ## Related Functions
//!
//! - `Sgn`: Returns the sign of a number (-1, 0, or 1)
//! - `Fix`: Returns the integer portion of a number (truncates toward zero)
//! - `Int`: Returns the integer portion of a number (rounds down)
//! - `Round`: Rounds a number to a specified number of decimal places
//!
//! ## Type Compatibility
//!
//! | Input Type | Return Type | Notes |
//! |------------|-------------|-------|
//! | Byte | Byte | Always positive already |
//! | Integer | Integer | Can overflow at -32768 |
//! | Long | Long | Can overflow at -2147483648 |
//! | Single | Single | Preserves precision |
//! | Double | Double | Preserves precision |
//! | Currency | Currency | Preserves 4 decimal places |
//! | Variant (numeric) | Variant | Type preserved |
//! | Variant (Null) | Null | Returns Null |
//!
//! ## Performance Notes
//!
//! - `Abs` is a very fast intrinsic function
//! - No function call overhead in compiled code
//! - Optimized to CPU instructions where possible
//! - Prefer `Abs` over manual `If`/`Then` checks for performance
//!
//! `Abs` is parsed as a regular function call (`CallExpression`)
//! This module serves as documentation and reference for the Abs function

#[cfg(test)]
mod test {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn abs_simple_negative() {
        let source = r"
Sub Test()
    x = Abs(-50)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("50"),
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
    fn abs_simple_positive() {
        let source = r"
Sub Test()
    x = Abs(25)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("25"),
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
    fn abs_with_zero() {
        let source = r"
Sub Test()
    x = Abs(0)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("0"),
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
    fn abs_with_variable() {
        let source = r"
Sub Test()
    result = Abs(value)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("value"),
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
    fn abs_with_expression() {
        let source = r"
Sub Test()
    diff = Abs(x - y)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("diff"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("x"),
                                        },
                                        Whitespace,
                                        SubtractionOperator,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("y"),
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
    fn abs_floating_point() {
        let source = r"
Sub Test()
    distance = Abs(-12.75)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("distance"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            SingleLiteral,
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
    fn abs_in_assignment() {
        let source = r"
Sub Test()
    Dim x As Integer
    x = Abs(-100)
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
                        Identifier ("x"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("100"),
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
    fn abs_in_if_condition() {
        let source = r"
Sub Test()
    If Abs(value) > 100 Then
        ProcessLargeValue
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
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Abs"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("value"),
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
                            CallStatement {
                                Identifier ("ProcessLargeValue"),
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
    fn abs_in_comparison() {
        let source = r"
Sub Test()
    If Abs(x - y) < tolerance Then
        DoSomething
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
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Abs"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("x"),
                                            },
                                            Whitespace,
                                            SubtractionOperator,
                                            Whitespace,
                                            IdentifierExpression {
                                                Identifier ("y"),
                                            },
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("tolerance"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("DoSomething"),
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
    fn abs_nested_call() {
        let source = r"
Sub Test()
    result = Abs(GetValue())
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("GetValue"),
                                        LeftParenthesis,
                                        ArgumentList,
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
    fn abs_multiple_calls() {
        let source = r"
Sub Test()
    a = Abs(-10)
    b = Abs(-20)
    c = Abs(-30)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("a"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("10"),
                                        },
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
                            Identifier ("b"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("20"),
                                        },
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
                            Identifier ("c"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("30"),
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
    fn abs_in_function_return() {
        let source = r"
Function GetDistance() As Double
    GetDistance = Abs(x2 - x1)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetDistance"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetDistance"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("x2"),
                                        },
                                        Whitespace,
                                        SubtractionOperator,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("x1"),
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
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn abs_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 10
        arr(i) = Abs(arr(i))
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
                            Whitespace,
                            AssignmentStatement {
                                CallExpression {
                                    Identifier ("arr"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("i"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Abs"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            CallExpression {
                                                Identifier ("arr"),
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        IdentifierExpression {
                                                            Identifier ("i"),
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
    fn abs_with_array_element() {
        let source = r"
Sub Test()
    result = Abs(values(index))
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("values"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("index"),
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
    fn abs_with_property_access() {
        let source = r"
Sub Test()
    total = Abs(obj.Value)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("total"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    MemberAccessExpression {
                                        Identifier ("obj"),
                                        PeriodOperator,
                                        Identifier ("Value"),
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
    fn abs_case_insensitive() {
        let source = r"
Sub Test()
    x = ABS(-50)
    y = abs(-25)
    z = AbS(-10)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("ABS"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("50"),
                                        },
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
                            Identifier ("y"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("25"),
                                        },
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
                            Identifier ("z"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AbS"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("10"),
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
    fn abs_in_print() {
        let source = r"
Sub Test()
    Debug.Print Abs(-42)
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
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("Abs"),
                        LeftParenthesis,
                        SubtractionOperator,
                        IntegerLiteral ("42"),
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
    fn abs_in_select_case() {
        let source = r"
Sub Test()
    Select Case Abs(value)
        Case Is > 100
            ProcessLarge
        Case Else
            ProcessSmall
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
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("value"),
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
                            IsKeyword,
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            IntegerLiteral ("100"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("ProcessLarge"),
                                    Newline,
                                },
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
                                CallStatement {
                                    Identifier ("ProcessSmall"),
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
    fn abs_with_parenthesized_expression() {
        let source = r"
Sub Test()
    result = Abs((x + y) * 2)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        ParenthesizedExpression {
                                            LeftParenthesis,
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    Identifier ("x"),
                                                },
                                                Whitespace,
                                                AdditionOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("y"),
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
    fn abs_in_do_loop() {
        let source = r"
Sub Test()
    Do While Abs(delta) > 0.001
        Adjust
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
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Abs"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("delta"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                SingleLiteral,
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Adjust"),
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
    fn abs_with_type_suffix() {
        let source = r"
Sub Test()
    x = Abs(-100%)
    y = Abs(-200&)
    z = Abs(-3.14#)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("100%"),
                                        },
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
                            Identifier ("y"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            LongLiteral,
                                        },
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
                            Identifier ("z"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            DoubleLiteral,
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
    fn abs_in_while_loop() {
        let source = r"
Sub Test()
    While Abs(current - target) > threshold
        Step
    Wend
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
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Abs"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("current"),
                                            },
                                            Whitespace,
                                            SubtractionOperator,
                                            Whitespace,
                                            IdentifierExpression {
                                                Identifier ("target"),
                                            },
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("threshold"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            Unknown,
                            Newline,
                            Whitespace,
                        },
                        WendKeyword,
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
    fn abs_with_binary_operators() {
        let source = r"
Sub Test()
    result = Abs(a + b - c * d / e)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("a"),
                                            },
                                            Whitespace,
                                            AdditionOperator,
                                            Whitespace,
                                            IdentifierExpression {
                                                Identifier ("b"),
                                            },
                                        },
                                        Whitespace,
                                        SubtractionOperator,
                                        Whitespace,
                                        BinaryExpression {
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    Identifier ("c"),
                                                },
                                                Whitespace,
                                                MultiplicationOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("d"),
                                                },
                                            },
                                            Whitespace,
                                            DivisionOperator,
                                            Whitespace,
                                            IdentifierExpression {
                                                Identifier ("e"),
                                            },
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
    fn abs_in_with_block() {
        let source = r"
Sub Test()
    With myObject
        .Value = Abs(.Delta)
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
                        Identifier ("myObject"),
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("Value"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Abs"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    PeriodOperator,
                                                },
                                            },
                                        },
                                    },
                                },
                            },
                            CallStatement {
                                Identifier ("Delta"),
                                RightParenthesis,
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
    fn abs_currency_literal() {
        let source = r"
Sub Test()
    amount = Abs(-1234.56@)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("amount"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            DecimalLiteral,
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
    fn abs_in_function_parameter() {
        let source = r"
Sub Test()
    Call ProcessValue(Abs(-50))
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
                    CallStatement {
                        Whitespace,
                        CallKeyword,
                        Whitespace,
                        Identifier ("ProcessValue"),
                        LeftParenthesis,
                        Identifier ("Abs"),
                        LeftParenthesis,
                        SubtractionOperator,
                        IntegerLiteral ("50"),
                        RightParenthesis,
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
    fn abs_chained_operations() {
        let source = r"
Sub Test()
    result = Abs(x) + Abs(y) - Abs(z)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("Abs"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("x"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                AdditionOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Abs"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("y"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                            Whitespace,
                            SubtractionOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("Abs"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("z"),
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
    fn abs_at_module_level() {
        let source = r"
Const MAX_VALUE = Abs(-1000)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                ConstKeyword,
                Whitespace,
                Identifier ("MAX_VALUE"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Abs"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            UnaryExpression {
                                SubtractionOperator,
                                NumericLiteralExpression {
                                    IntegerLiteral ("1000"),
                                },
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn abs_with_unary_minus() {
        let source = r"
Sub Test()
    x = Abs(-x)
    y = Abs(-(a + b))
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        IdentifierExpression {
                                            Identifier ("x"),
                                        },
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
                            Identifier ("y"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        ParenthesizedExpression {
                                            LeftParenthesis,
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    Identifier ("a"),
                                                },
                                                Whitespace,
                                                AdditionOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("b"),
                                                },
                                            },
                                            RightParenthesis,
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
    fn abs_preserves_whitespace() {
        let source = r"
Sub Test()
    x = Abs  (  -50  )
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Abs"),
                            Whitespace,
                            LeftParenthesis,
                            ArgumentList {
                                Whitespace,
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("50"),
                                        },
                                    },
                                },
                                Whitespace,
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
}
