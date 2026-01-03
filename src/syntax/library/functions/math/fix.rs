//! # Fix Function
//!
//! Returns the integer portion of a number.
//!
//! ## Syntax
//!
//! ```vb
//! Fix(number)
//! ```
//!
//! ## Parameters
//!
//! - `number` (Required): Any valid numeric expression. If number contains Null, Null is returned
//!
//! ## Return Value
//!
//! Returns the integer portion of a number:
//! - For positive numbers: Returns the largest integer less than or equal to number
//! - For negative numbers: Returns the first negative integer greater than or equal to number
//! - If number is Null: Returns Null
//! - Return type matches the input type (Integer, Long, Single, Double, Currency, Decimal)
//!
//! ## Remarks
//!
//! The Fix function truncates toward zero:
//!
//! - Removes the fractional part of a number
//! - Always truncates toward zero (removes decimal without rounding)
//! - For positive numbers, behaves like `Int` (same result)
//! - For negative numbers, different from `Int` (`Int` rounds down, `Fix` truncates)
//! - `Fix`(-8.4) returns -8, `Int`(-8.4) returns -9
//! - `Fix`(8.4) returns 8, `Int`(8.4) returns 8
//! - Does not round to nearest integer (use `Round` for rounding)
//! - The return type preserves the input numeric type
//! - More intuitive for most developers (truncation toward zero)
//! - Commonly used when you want to discard fractional parts
//! - For financial calculations, consider using `Round` or `CCur` instead
//!
//! ## Typical Uses
//!
//! 1. **Truncate Decimals**: Remove fractional part without rounding
//! 2. **Integer Conversion**: Convert floating-point to integer values
//! 3. **Financial Calculations**: Remove cents from currency values
//! 4. **Data Normalization**: Ensure whole number values
//! 5. **Display Formatting**: Show only whole number portion
//! 6. **Loop Bounds**: Convert float bounds to integers
//! 7. **Array Indices**: Ensure valid integer indices
//! 8. **Coordinate Processing**: Truncate pixel coordinates
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Truncate positive number
//! Dim result As Integer
//! result = Fix(8.7)
//! Debug.Print result  ' Prints: 8
//!
//! ' Example 2: Truncate negative number
//! Dim result As Integer
//! result = Fix(-8.7)
//! Debug.Print result  ' Prints: -8 (truncates toward zero, not down)
//!
//! ' Example 3: Remove cents from currency
//! Dim price As Currency
//! Dim dollars As Currency
//! price = 19.99
//! dollars = Fix(price)
//! Debug.Print dollars  ' Prints: 19
//!
//! ' Example 4: Ensure integer for array index
//! Dim index As Long
//! Dim ratio As Double
//! ratio = 4.7
//! index = Fix(ratio)
//! value = items(index)
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Truncate toward zero
//! Function Truncate(value As Double) As Long
//!     Truncate = Fix(value)
//! End Function
//!
//! ' Pattern 2: Get whole dollars from currency
//! Function GetWholeDollars(amount As Currency) As Long
//!     GetWholeDollars = Fix(amount)
//! End Function
//!
//! ' Pattern 3: Get cents from currency
//! Function GetCents(amount As Currency) As Long
//!     Dim wholeDollars As Currency
//!     wholeDollars = Fix(amount)
//!     GetCents = Fix((amount - wholeDollars) * 100)
//! End Function
//!
//! ' Pattern 4: Split number into whole and fractional parts
//! Sub SplitNumber(value As Double, ByRef wholePart As Long, ByRef fractionalPart As Double)
//!     wholePart = Fix(value)
//!     fractionalPart = value - wholePart
//! End Sub
//!
//! ' Pattern 5: Ensure positive truncation
//! Function TruncatePositive(value As Double) As Long
//!     ' Fix truncates toward zero
//!     ' For negative values, this gives different result than Int
//!     TruncatePositive = Fix(Abs(value)) * Sgn(value)
//! End Function
//!
//! ' Pattern 6: Convert to integer without rounding
//! Function ToIntegerNoRound(value As Double) As Long
//!     ToIntegerNoRound = Fix(value)
//! End Function
//!
//! ' Pattern 7: Remove decimal places for display
//! Function FormatWholeNumber(value As Double) As String
//!     FormatWholeNumber = CStr(Fix(value))
//! End Function
//!
//! ' Pattern 8: Calculate whole units
//! Function GetWholeUnits(quantity As Double) As Long
//!     GetWholeUnits = Fix(quantity)
//! End Function
//!
//! ' Pattern 9: Truncate time to hours
//! Function GetWholeHours(timeValue As Double) As Long
//!     Dim hours As Double
//!     hours = timeValue * 24  ' Convert days to hours
//!     GetWholeHours = Fix(hours)
//! End Function
//!
//! ' Pattern 10: Floor for positive, ceiling for negative
//! Function TruncateTowardZero(value As Double) As Long
//!     ' Fix already does this
//!     TruncateTowardZero = Fix(value)
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Currency formatter class
//! Public Class CurrencyFormatter
//!     Public Function FormatAsDollarsAndCents(amount As Currency) As String
//!         Dim dollars As Long
//!         Dim cents As Long
//!         
//!         dollars = Fix(amount)
//!         cents = Fix(Abs((amount - dollars) * 100))
//!         
//!         FormatAsDollarsAndCents = "$" & dollars & "." & _
//!                                   Format$(cents, "00")
//!     End Function
//!     
//!     Public Function GetDollarPart(amount As Currency) As Long
//!         GetDollarPart = Fix(amount)
//!     End Function
//!     
//!     Public Function GetCentPart(amount As Currency) As Long
//!         Dim dollars As Currency
//!         dollars = Fix(amount)
//!         GetCentPart = Fix(Abs((amount - dollars) * 100))
//!     End Function
//!     
//!     Public Function RoundToDollars(amount As Currency) As Currency
//!         RoundToDollars = Fix(amount)
//!     End Function
//! End Class
//!
//! ' Example 2: Number splitter utility
//! Public Class NumberSplitter
//!     Private m_wholePart As Long
//!     Private m_fractionalPart As Double
//!     
//!     Public Sub Split(value As Double)
//!         m_wholePart = Fix(value)
//!         m_fractionalPart = value - m_wholePart
//!     End Sub
//!     
//!     Public Property Get WholePart() As Long
//!         WholePart = m_wholePart
//!     End Property
//!     
//!     Public Property Get FractionalPart() As Double
//!         FractionalPart = m_fractionalPart
//!     End Property
//!     
//!     Public Property Get HasFraction() As Boolean
//!         HasFraction = (m_fractionalPart <> 0)
//!     End Property
//!     
//!     Public Function Reconstruct() As Double
//!         Reconstruct = m_wholePart + m_fractionalPart
//!     End Function
//! End Class
//!
//! ' Example 3: Data truncator for normalization
//! Public Class DataTruncator
//!     Public Function TruncateArray(values() As Double) As Long()
//!         Dim result() As Long
//!         Dim i As Long
//!         
//!         ReDim result(LBound(values) To UBound(values))
//!         
//!         For i = LBound(values) To UBound(values)
//!             result(i) = Fix(values(i))
//!         Next i
//!         
//!         TruncateArray = result
//!     End Function
//!     
//!     Public Function TruncateToInteger(value As Double) As Long
//!         TruncateToInteger = Fix(value)
//!     End Function
//!     
//!     Public Function TruncateCollection(values As Collection) As Collection
//!         Dim result As New Collection
//!         Dim value As Variant
//!         
//!         For Each value In values
//!             If IsNumeric(value) Then
//!                 result.Add Fix(CDbl(value))
//!             Else
//!                 result.Add value
//!             End If
//!         Next value
//!         
//!         Set TruncateCollection = result
//!     End Function
//! End Class
//!
//! ' Example 4: Coordinate truncator
//! Public Class CoordinateTruncator
//!     Public Sub TruncatePoint(x As Double, y As Double, _
//!                             ByRef truncX As Long, ByRef truncY As Long)
//!         truncX = Fix(x)
//!         truncY = Fix(y)
//!     End Sub
//!     
//!     Public Function TruncateRectangle(left As Double, top As Double, _
//!                                       right As Double, bottom As Double) As Variant
//!         Dim coords(0 To 3) As Long
//!         
//!         coords(0) = Fix(left)
//!         coords(1) = Fix(top)
//!         coords(2) = Fix(right)
//!         coords(3) = Fix(bottom)
//!         
//!         TruncateRectangle = coords
//!     End Function
//!     
//!     Public Function GetPixelCoordinate(value As Double) As Long
//!         GetPixelCoordinate = Fix(value)
//!     End Function
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! The Fix function can raise errors or return Null:
//!
//! - **Type Mismatch (Error 13)**: If number is not a numeric expression
//! - **Invalid use of Null (Error 94)**: If number is Null and result is assigned to non-Variant
//! - **Overflow (Error 6)**: If result exceeds the range of the target data type
//!
//! ```vb
//! On Error GoTo ErrorHandler
//! Dim result As Long
//! Dim value As Double
//!
//! value = -12.75
//! result = Fix(value)
//!
//! Debug.Print "Truncated value: " & result  ' Prints: -12
//! Exit Sub
//!
//! ErrorHandler:
//!     MsgBox "Error in Fix: " & Err.Description, vbCritical
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: Fix is a very fast built-in function
//! - **Type Preservation**: Return type matches input type
//! - **No Rounding**: Faster than Round (simple truncation)
//! - **Alternative**: For floor operation, use Int (rounds down)
//! - **Currency**: More intuitive than Int for currency truncation
//!
//! ## Best Practices
//!
//! 1. **Understand Difference**: Know that Fix truncates toward zero, Int rounds down
//! 2. **Negative Numbers**: Be aware Fix(-8.7) = -8, Int(-8.7) = -9
//! 3. **Currency Operations**: Fix is more intuitive for removing cents
//! 4. **Type Awareness**: Be aware of return type matching input type
//! 5. **Null Handling**: Use Variant if input might be Null
//! 6. **No Rounding**: Use Round if you need rounding, not truncation
//! 7. **Documentation**: Comment when Fix vs Int choice matters
//!
//! ## Comparison with Other Functions
//!
//! | Function | Behavior with -8.7 | Behavior with 8.7 | Description |
//! |----------|-------------------|-------------------|-------------|
//! | `Fix` | -8 | 8 | Truncates toward zero |
//! | `Int` | -9 | 8 | Rounds down (floor) |
//! | `Round` | -9 | 9 | Rounds to nearest |
//! | `CLng` | -9 | 9 | Converts to `Long` with rounding |
//! | `CInt` | -9 | 9 | Converts to `Integer` with rounding |
//! | `\` | N/A | N/A | `Integer` division operator |
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Consistent behavior across platforms
//! - Return type matches input numeric type
//! - Truncates toward zero (like C/C++/Java integer truncation)
//! - More intuitive than Int for most developers from other languages
//!
//! ## Limitations
//!
//! - Does not round to nearest (use `Round` for that)
//! - Behavior with negative numbers differs from `Int`
//! - Return type depends on input type (can cause overflow)
//! - Cannot specify decimal places (always removes all decimals)
//! - No control over rounding direction (always toward zero)
//!
//! ## Related Functions
//!
//! - `Int`: Returns `Integer` portion, rounding down (floor)
//! - `Round`: Rounds to nearest integer or specified decimal places
//! - `CInt`: Converts to `Integer` with rounding
//! - `CLng`: Converts to `Long` with rounding
//! - `Abs`: Absolute value (often used with `Fix`)
//! - `Sgn`: Sign of number (often used with `Fix`)

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn fix_basic() {
        let source = r"
Sub Test()
    result = Fix(8.7)
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
                            Identifier ("Fix"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        SingleLiteral,
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
    fn fix_negative() {
        let source = r"
Sub Test()
    result = Fix(-8.7)
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
                            Identifier ("Fix"),
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
    fn fix_currency() {
        let source = r"
Sub Test()
    dollars = Fix(price)
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
                            Identifier ("dollars"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Fix"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("price"),
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
    fn fix_if_statement() {
        let source = r#"
Sub Test()
    If Fix(value) > 10 Then
        Debug.Print "Greater than 10"
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
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Fix"),
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
                                IntegerLiteral ("10"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                StringLiteral ("\"Greater than 10\""),
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
    fn fix_function_return() {
        let source = r"
Function Truncate(value As Double) As Long
    Truncate = Fix(value)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("Truncate"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("value"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                LongKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("Truncate"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Fix"),
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
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn fix_comparison() {
        let source = r#"
Sub Test()
    If Fix(amount) = expectedAmount Then
        MsgBox "Amount matches"
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
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Fix"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("amount"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("expectedAmount"),
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
                                StringLiteral ("\"Amount matches\""),
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
    fn fix_select_case() {
        let source = r#"
Sub Test()
    Select Case Fix(score)
        Case 100
            grade = "A+"
        Case 90 To 99
            grade = "A"
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
                        CallExpression {
                            Identifier ("Fix"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("score"),
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
                            IntegerLiteral ("100"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("grade"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"A+\""),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("90"),
                            Whitespace,
                            ToKeyword,
                            Whitespace,
                            IntegerLiteral ("99"),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("grade"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"A\""),
                                    },
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
    fn fix_for_loop() {
        let source = r"
Sub Test()
    Dim i As Long
    For i = 1 To Fix(maxValue)
        Debug.Print i
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
                            Identifier ("Fix"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("maxValue"),
                                    },
                                },
                            },
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
    fn fix_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Whole part: " & Fix(value)
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
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        StringLiteral ("\"Whole part: \""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Fix"),
                        LeftParenthesis,
                        Identifier ("value"),
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
    fn fix_array_assignment() {
        let source = r"
Sub Test()
    wholeNumbers(i) = Fix(decimals(i))
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
                        CallExpression {
                            Identifier ("wholeNumbers"),
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
                            Identifier ("Fix"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("decimals"),
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
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn fix_property_assignment() {
        let source = r"
Sub Test()
    obj.WholePart = Fix(obj.DecimalValue)
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
                        MemberAccessExpression {
                            Identifier ("obj"),
                            PeriodOperator,
                            Identifier ("WholePart"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Fix"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    MemberAccessExpression {
                                        Identifier ("obj"),
                                        PeriodOperator,
                                        Identifier ("DecimalValue"),
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
    fn fix_in_class() {
        let source = r"
Private Sub Class_Initialize()
    m_wholePart = Fix(m_value)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                PrivateKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("Class_Initialize"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("m_wholePart"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Fix"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("m_value"),
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
    fn fix_with_statement() {
        let source = r"
Sub Test()
    With splitter
        .WholePart = Fix(.Value)
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
                        Identifier ("splitter"),
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("WholePart"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Fix"),
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
                                Identifier ("Value"),
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
    fn fix_function_argument() {
        let source = r"
Sub Test()
    Call ProcessInteger(Fix(value))
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
                        Identifier ("ProcessInteger"),
                        LeftParenthesis,
                        Identifier ("Fix"),
                        LeftParenthesis,
                        Identifier ("value"),
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
    fn fix_concatenation() {
        let source = r#"
Sub Test()
    message = "Truncated: " & Fix(number)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("message"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"Truncated: \""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Fix"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("number"),
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
    fn fix_math_expression() {
        let source = r"
Sub Test()
    wholePart = Fix(value)
    fractionalPart = value - wholePart
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
                            Identifier ("wholePart"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Fix"),
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("fractionalPart"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("value"),
                            },
                            Whitespace,
                            SubtractionOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("wholePart"),
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
    fn fix_iif() {
        let source = r#"
Sub Test()
    result = IIf(Fix(value) > 0, "Positive", "Zero or Negative")
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("IIf"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        CallExpression {
                                            Identifier ("Fix"),
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
                                            IntegerLiteral ("0"),
                                        },
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Positive\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Zero or Negative\""),
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
    fn fix_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Whole value: " & Fix(amount)
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
                    CallStatement {
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Whole value: \""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Fix"),
                        LeftParenthesis,
                        Identifier ("amount"),
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
    fn fix_collection_add() {
        let source = r"
Sub Test()
    wholeNumbers.Add Fix(values(i))
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
                        Identifier ("wholeNumbers"),
                        PeriodOperator,
                        Identifier ("Add"),
                        Whitespace,
                        Identifier ("Fix"),
                        LeftParenthesis,
                        Identifier ("values"),
                        LeftParenthesis,
                        Identifier ("i"),
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
    fn fix_boolean_expression() {
        let source = r"
Sub Test()
    isValid = Fix(value) >= minValue And Fix(value) <= maxValue
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
                            Identifier ("isValid"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("Fix"),
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
                                GreaterThanOrEqualOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("minValue"),
                                },
                            },
                            Whitespace,
                            AndKeyword,
                            Whitespace,
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("Fix"),
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
                                LessThanOrEqualOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("maxValue"),
                                },
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
    fn fix_nested_call() {
        let source = r"
Sub Test()
    result = CStr(Fix(value))
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
                            Identifier ("CStr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Fix"),
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
    fn fix_do_loop() {
        let source = r"
Sub Test()
    Do While Fix(counter) < limit
        counter = counter + 0.5
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
                                Identifier ("Fix"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("counter"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("limit"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("counter"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("counter"),
                                    },
                                    Whitespace,
                                    AdditionOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        SingleLiteral,
                                    },
                                },
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
    fn fix_abs() {
        let source = r"
Sub Test()
    cents = Fix(Abs((amount - dollars) * 100))
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
                            Identifier ("cents"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Fix"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
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
                                                                Identifier ("amount"),
                                                            },
                                                            Whitespace,
                                                            SubtractionOperator,
                                                            Whitespace,
                                                            IdentifierExpression {
                                                                Identifier ("dollars"),
                                                            },
                                                        },
                                                        RightParenthesis,
                                                    },
                                                    Whitespace,
                                                    MultiplicationOperator,
                                                    Whitespace,
                                                    NumericLiteralExpression {
                                                        IntegerLiteral ("100"),
                                                    },
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
    fn fix_currency_split() {
        let source = r"
Sub Test()
    dollars = Fix(amount)
    cents = Fix((amount - dollars) * 100)
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
                            Identifier ("dollars"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Fix"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("amount"),
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
                            Identifier ("cents"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Fix"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        ParenthesizedExpression {
                                            LeftParenthesis,
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    Identifier ("amount"),
                                                },
                                                Whitespace,
                                                SubtractionOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("dollars"),
                                                },
                                            },
                                            RightParenthesis,
                                        },
                                        Whitespace,
                                        MultiplicationOperator,
                                        Whitespace,
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
    fn fix_coordinate() {
        let source = r"
Sub Test()
    pixelX = Fix(coordinateX)
    pixelY = Fix(coordinateY)
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
                            Identifier ("pixelX"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Fix"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("coordinateX"),
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
                            Identifier ("pixelY"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Fix"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("coordinateY"),
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
    fn fix_array_index() {
        let source = r"
Sub Test()
    index = Fix(ratio * arraySize)
    value = items(index)
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
                            Identifier ("index"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Fix"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("ratio"),
                                        },
                                        Whitespace,
                                        MultiplicationOperator,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("arraySize"),
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
                            Identifier ("value"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("items"),
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
    fn fix_parentheses() {
        let source = r"
Sub Test()
    value = (Fix(number))
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
                            Identifier ("value"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            CallExpression {
                                Identifier ("Fix"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("number"),
                                        },
                                    },
                                },
                                RightParenthesis,
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
