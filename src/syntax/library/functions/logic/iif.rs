//! # `IIf` Function
//!
//! Returns one of two parts, depending on the evaluation of an expression.
//!
//! ## Syntax
//!
//! ```vb
//! IIf(expr, truepart, falsepart)
//! ```
//!
//! ## Parameters
//!
//! - `expr` (Required): Expression you want to evaluate
//! - `truepart` (Required): Value or expression returned if `expr` is True
//! - `falsepart` (Required): Value or expression returned if `expr` is False
//!
//! ## Return Value
//!
//! Returns `truepart` if `expr` evaluates to True; otherwise returns `falsepart`. The return type
//! is `Variant` and depends on the types of `truepart` and `falsepart`.
//!
//! ## Remarks
//!
//! The `IIf` function provides inline conditional evaluation:
//!
//! - Always evaluates BOTH `truepart` and `falsepart` regardless of the condition result
//! - This can cause side effects if either part contains function calls or property accesses
//! - Returns `Variant` type, which may require explicit type conversion
//! - Can nest `IIf` calls for multiple conditions (though readability suffers)
//! - If `expr` is Null, the function returns Null
//! - Unlike `If...Then...Else` statements, `IIf` is an expression that returns a value
//! - Useful for inline assignments, but beware of evaluation side effects
//! - Consider using `If...Then...Else` for complex logic or when side effects matter
//!
//! ## Typical Uses
//!
//! 1. **Inline Conditionals**: Simple conditional value assignment in one line
//! 2. **String Formatting**: Choose between different string representations
//! 3. **Calculated Fields**: Conditional calculations in expressions
//! 4. **Default Values**: Provide fallback values for empty or null data
//! 5. **Display Logic**: Choose display text based on conditions
//! 6. **Data Validation**: Return appropriate values based on validation results
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Simple conditional assignment
//! Dim result As String
//! result = IIf(age >= 18, "Adult", "Minor")
//!
//! ' Example 2: Numeric comparison
//! Dim status As String
//! status = IIf(score >= 60, "Pass", "Fail")
//!
//! ' Example 3: Null handling
//! Dim display As String
//! display = IIf(IsNull(value), "N/A", CStr(value))
//!
//! ' Example 4: Sign determination
//! Dim sign As String
//! sign = IIf(number >= 0, "+", "-")
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Choose singular or plural
//! Function FormatCount(count As Long, singular As String, plural As String) As String
//!     FormatCount = count & " " & IIf(count = 1, singular, plural)
//! End Function
//! ' Usage: FormatCount(5, "item", "items") => "5 items"
//!
//! ' Pattern 2: Min/Max selection
//! Function Min(a As Double, b As Double) As Double
//!     Min = IIf(a < b, a, b)
//! End Function
//!
//! Function Max(a As Double, b As Double) As Double
//!     Max = IIf(a > b, a, b)
//! End Function
//!
//! ' Pattern 3: Safe division
//! Function SafeDivide(numerator As Double, denominator As Double) As Variant
//!     SafeDivide = IIf(denominator <> 0, numerator / denominator, Null)
//! End Function
//!
//! ' Pattern 4: Empty string default
//! Function GetDisplayName(name As String) As String
//!     GetDisplayName = IIf(Len(Trim$(name)) > 0, name, "(unnamed)")
//! End Function
//!
//! ' Pattern 5: Range clamping
//! Function Clamp(value As Long, minVal As Long, maxVal As Long) As Long
//!     Clamp = IIf(value < minVal, minVal, IIf(value > maxVal, maxVal, value))
//! End Function
//!
//! ' Pattern 6: Boolean to integer
//! Function BoolToInt(value As Boolean) As Integer
//!     BoolToInt = IIf(value, 1, 0)
//! End Function
//!
//! ' Pattern 7: Sign function
//! Function Sign(value As Double) As Integer
//!     Sign = IIf(value > 0, 1, IIf(value < 0, -1, 0))
//! End Function
//!
//! ' Pattern 8: Null coalescing
//! Function Coalesce(value As Variant, defaultValue As Variant) As Variant
//!     Coalesce = IIf(IsNull(value) Or IsEmpty(value), defaultValue, value)
//! End Function
//!
//! ' Pattern 9: Conditional formatting
//! Function FormatBalance(balance As Currency) As String
//!     FormatBalance = IIf(balance < 0, _
//!                         "(" & Format$(Abs(balance), "Currency") & ")", _
//!                         Format$(balance, "Currency"))
//! End Function
//!
//! ' Pattern 10: Toggle value
//! Function Toggle(current As Boolean) As Boolean
//!     Toggle = IIf(current, False, True)
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Grade calculator with nested IIf
//! Function GetGrade(score As Double) As String
//!     GetGrade = IIf(score >= 90, "A", _
//!                IIf(score >= 80, "B", _
//!                IIf(score >= 70, "C", _
//!                IIf(score >= 60, "D", "F"))))
//! End Function
//!
//! ' Example 2: Complex string builder
//! Function BuildMessage(userName As String, isAdmin As Boolean, messageCount As Long) As String
//!     BuildMessage = "Welcome " & IIf(Len(userName) > 0, userName, "Guest") & _
//!                    IIf(isAdmin, " (Admin)", "") & _
//!                    IIf(messageCount > 0, " - You have " & messageCount & " message" & _
//!                        IIf(messageCount = 1, "", "s"), "")
//! End Function
//!
//! ' Example 3: Data validation with IIf
//! Function ValidateAndFormat(input As String, Optional maxLen As Long = 50) As String
//!     Dim cleaned As String
//!     cleaned = Trim$(input)
//!     
//!     ValidateAndFormat = IIf(Len(cleaned) = 0, "", _
//!                         IIf(Len(cleaned) > maxLen, _
//!                             Left$(cleaned, maxLen) & "...", _
//!                             cleaned))
//! End Function
//!
//! ' Example 4: Status indicator with color codes
//! Function GetStatusDisplay(value As Double, threshold As Double) As String
//!     Dim status As String
//!     Dim color As String
//!     
//!     status = IIf(value >= threshold, "OK", "WARNING")
//!     color = IIf(value >= threshold, "Green", "Red")
//!     
//!     GetStatusDisplay = "[" & color & "] " & status & " (" & value & ")"
//! End Function
//!
//! ' Example 5: Conditional object creation (DANGEROUS - both parts evaluate!)
//! ' WARNING: This pattern has side effects!
//! Function GetConnection(useLocal As Boolean) As Object
//!     ' BOTH CreateLocalConnection AND CreateRemoteConnection will execute!
//!     ' Use If...Then...Else instead for object creation
//!     Set GetConnection = IIf(useLocal, CreateLocalConnection(), CreateRemoteConnection())
//! End Function
//!
//! ' Example 6: Safe property access
//! Function GetPropertyValue(obj As Object, propertyName As String, defaultValue As Variant) As Variant
//!     On Error Resume Next
//!     Dim value As Variant
//!     value = CallByName(obj, propertyName, VbGet)
//!     
//!     If Err.Number = 0 Then
//!         GetPropertyValue = IIf(IsNull(value), defaultValue, value)
//!     Else
//!         GetPropertyValue = defaultValue
//!     End If
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! The `IIf` function itself rarely raises errors, but be aware of:
//!
//! - **Type Mismatch (Error 13)**: Can occur if the result type doesn't match the receiving variable
//! - **Evaluation Errors**: Both `truepart` and `falsepart` are always evaluated, so errors in either will occur
//! - **Null Propagation**: If `expr` is Null, `IIf` returns Null
//! - **Division by Zero**: Can occur if either part contains division and is evaluated
//!
//! ```vb
//! ' WRONG - Both divisions execute regardless of condition!
//! result = IIf(denominator <> 0, numerator / denominator, numerator / 1)
//! ' If denominator is 0, division by zero error still occurs in first part
//!
//! ' CORRECT - Use If...Then...Else for conditional execution
//! If denominator <> 0 Then
//!     result = numerator / denominator
//! Else
//!     result = numerator / 1
//! End If
//! ```
//!
//! ## Performance Considerations
//!
//! - **Both Branches Evaluate**: `IIf` always evaluates both `truepart` and `falsepart`
//! - **Function Call Overhead**: `IIf` has function call overhead vs. `If...Then...Else`
//! - **Variant Boxing**: Results are `Variant` type, which may require type conversion
//! - **Nested Performance**: Deeply nested `IIf` calls can be slow and hard to read
//! - **Use `If...Then...Else` When**: Either branch has expensive operations or side effects
//!
//! ## Best Practices
//!
//! 1. **Avoid Side Effects**: Don't use `IIf` when either part has side effects (function calls, object creation, I/O)
//! 2. **Keep It Simple**: Use `IIf` for simple value selection only
//! 3. **Limit Nesting**: Avoid deeply nested `IIf` calls (use `Select Case` or `If...Then...Else` instead)
//! 4. **Type Safety**: Be aware of `Variant` return type and convert explicitly if needed
//! 5. **Readability**: If `IIf` makes code harder to read, use `If...Then...Else`
//! 6. **Document Expectations**: When using `IIf`, document that both branches evaluate
//!
//! ## When NOT to Use `IIf`
//!
//! ```vb
//! ' DON'T: Object creation (both execute!)
//! Set obj = IIf(condition, New ClassA, New ClassB)
//!
//! ' DON'T: Functions with side effects (both execute!)
//! result = IIf(condition, LogAndReturn("A"), LogAndReturn("B"))
//!
//! ' DON'T: Error-prone operations (both execute!)
//! value = IIf(x <> 0, 100 / x, 0)  ' Division by zero still occurs!
//!
//! ' DON'T: Complex nested logic (hard to read)
//! result = IIf(a, IIf(b, IIf(c, 1, 2), IIf(d, 3, 4)), IIf(e, 5, 6))
//!
//! ' DO: Use If...Then...Else instead
//! If condition Then
//!     Set obj = New ClassA
//! Else
//!     Set obj = New ClassB
//! End If
//! ```
//!
//! ## Comparison with Other Approaches
//!
//! | Approach | Evaluates Both | Return Type | Use Case |
//! |----------|---------------|-------------|----------|
//! | `IIf` | Yes | `Variant` | Simple inline value selection |
//! | `If...Then...Else` | No | Any | Conditional execution, side effects |
//! | `Select Case` | No | Any | Multiple conditions |
//! | `Choose` | Yes | `Variant` | Index-based selection |
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Consistent behavior across Windows platforms
//! - VBA also includes `IIf` with identical behavior
//! - Always returns `Variant` type
//! - Evaluation of both branches is by design, not a bug
//!
//! ## Limitations
//!
//! - Cannot short-circuit evaluation (both parts always execute)
//! - Returns Variant type (requires explicit conversion for strong typing)
//! - Not suitable for conditional execution (use `If...Then...Else`)
//! - Nested `IIf` calls quickly become unreadable
//! - Cannot handle multiple conditions as cleanly as `Select Case`
//! - May have performance overhead compared to `If...Then...Else`
//!
//! ## Related Functions
//!
//! - `If...Then...Else`: Statement for conditional execution with short-circuit evaluation
//! - `Choose`: Returns value from list based on numeric index (also evaluates all parts)
//! - `Switch`: Returns first value whose expression is True (evaluates sequentially)
//! - `Select Case`: Multi-condition statement with short-circuit evaluation

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn iif_basic() {
        let source = r#"
Sub Test()
    result = IIf(x > 0, "Positive", "Negative")
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
                                        IdentifierExpression {
                                            Identifier ("x"),
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
                                        StringLiteral ("\"Negative\""),
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
    fn iif_in_function() {
        let source = r#"
Function GetStatus(value As Integer) As String
    GetStatus = IIf(value >= 0, "OK", "Error")
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetStatus"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("value"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetStatus"),
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
                                        IdentifierExpression {
                                            Identifier ("value"),
                                        },
                                        Whitespace,
                                        GreaterThanOrEqualOperator,
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
                                        StringLiteral ("\"OK\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Error\""),
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
    fn iif_nested() {
        let source = r#"
Sub Test()
    grade = IIf(score >= 90, "A", IIf(score >= 80, "B", "C"))
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
                            Identifier ("grade"),
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
                                        IdentifierExpression {
                                            Identifier ("score"),
                                        },
                                        Whitespace,
                                        GreaterThanOrEqualOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("90"),
                                        },
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"A\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    CallExpression {
                                        Identifier ("IIf"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                BinaryExpression {
                                                    IdentifierExpression {
                                                        Identifier ("score"),
                                                    },
                                                    Whitespace,
                                                    GreaterThanOrEqualOperator,
                                                    Whitespace,
                                                    NumericLiteralExpression {
                                                        IntegerLiteral ("80"),
                                                    },
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                StringLiteralExpression {
                                                    StringLiteral ("\"B\""),
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                StringLiteralExpression {
                                                    StringLiteral ("\"C\""),
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
    fn iif_numeric() {
        let source = r"
Sub Test()
    value = IIf(a > b, a, b)
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
                        CallExpression {
                            Identifier ("IIf"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("a"),
                                        },
                                        Whitespace,
                                        GreaterThanOperator,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("b"),
                                        },
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("a"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("b"),
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
    fn iif_with_function_calls() {
        let source = r#"
Sub Test()
    msg = IIf(IsNull(value), "Empty", CStr(value))
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
                            Identifier ("msg"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("IIf"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("IsNull"),
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
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Empty\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    CallExpression {
                                        Identifier ("CStr"),
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
    fn iif_in_assignment() {
        let source = r#"
Sub Test()
    Dim status As String
    status = IIf(count = 1, "item", "items")
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
                        Identifier ("status"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("status"),
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
                                        IdentifierExpression {
                                            Identifier ("count"),
                                        },
                                        Whitespace,
                                        EqualityOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"item\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"items\""),
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
    fn iif_in_msgbox() {
        let source = r#"
Sub Test()
    MsgBox IIf(isValid, "Valid", "Invalid")
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
                        Identifier ("IIf"),
                        LeftParenthesis,
                        Identifier ("isValid"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Valid\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Invalid\""),
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
    fn iif_with_concatenation() {
        let source = r#"
Sub Test()
    text = "Count: " & IIf(n = 1, "one", "many")
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
                            TextKeyword,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"Count: \""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("IIf"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("n"),
                                            },
                                            Whitespace,
                                            EqualityOperator,
                                            Whitespace,
                                            NumericLiteralExpression {
                                                IntegerLiteral ("1"),
                                            },
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        StringLiteralExpression {
                                            StringLiteral ("\"one\""),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        StringLiteralExpression {
                                            StringLiteral ("\"many\""),
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
    fn iif_boolean_expression() {
        let source = r"
Sub Test()
    isEnabled = IIf(value > 0 And value < 100, True, False)
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
                            Identifier ("isEnabled"),
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
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("value"),
                                            },
                                            Whitespace,
                                            GreaterThanOperator,
                                            Whitespace,
                                            NumericLiteralExpression {
                                                IntegerLiteral ("0"),
                                            },
                                        },
                                        Whitespace,
                                        AndKeyword,
                                        Whitespace,
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("value"),
                                            },
                                            Whitespace,
                                            LessThanOperator,
                                            Whitespace,
                                            NumericLiteralExpression {
                                                IntegerLiteral ("100"),
                                            },
                                        },
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    BooleanLiteralExpression {
                                        TrueKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    BooleanLiteralExpression {
                                        FalseKeyword,
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
    fn iif_in_if_statement() {
        let source = r#"
Sub Test()
    If IIf(x > y, x, y) > 10 Then
        Debug.Print "Large"
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
                                Identifier ("IIf"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("x"),
                                            },
                                            Whitespace,
                                            GreaterThanOperator,
                                            Whitespace,
                                            IdentifierExpression {
                                                Identifier ("y"),
                                            },
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("x"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("y"),
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
                                StringLiteral ("\"Large\""),
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
    fn iif_in_for_loop() {
        let source = r"
Sub Test()
    Dim i As Integer
    For i = 1 To IIf(useMax, 100, 10)
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
                        IntegerKeyword,
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
                            Identifier ("IIf"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("useMax"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("100"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("10"),
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
    fn iif_in_select_case() {
        let source = r#"
Sub Test()
    Select Case IIf(value < 0, "negative", "positive")
        Case "negative"
            Debug.Print "Less than zero"
        Case "positive"
            Debug.Print "Greater than or equal to zero"
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
                            Identifier ("IIf"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("value"),
                                        },
                                        Whitespace,
                                        LessThanOperator,
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
                                        StringLiteral ("\"negative\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"positive\""),
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
                            StringLiteral ("\"negative\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Less than zero\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\"positive\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Greater than or equal to zero\""),
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
    fn iif_in_do_loop() {
        let source = r"
Sub Test()
    Do While IIf(count > 0, True, False)
        count = count - 1
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
                        CallExpression {
                            Identifier ("IIf"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("count"),
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
                                    BooleanLiteralExpression {
                                        TrueKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    BooleanLiteralExpression {
                                        FalseKeyword,
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("count"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("count"),
                                    },
                                    Whitespace,
                                    SubtractionOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
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
    fn iif_array_assignment() {
        let source = r"
Sub Test()
    arr(0) = IIf(flag, 1, 0)
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
                            Identifier ("arr"),
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
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("IIf"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("flag"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                                Comma,
                                Whitespace,
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
    fn iif_property_assignment() {
        let source = r"
Sub Test()
    obj.Value = IIf(enabled, 100, 0)
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
                            Identifier ("Value"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("IIf"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("enabled"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("100"),
                                    },
                                },
                                Comma,
                                Whitespace,
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
    fn iif_with_parentheses() {
        let source = r"
Sub Test()
    result = (IIf(x > 0, x, -x))
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
                        ParenthesizedExpression {
                            LeftParenthesis,
                            CallExpression {
                                Identifier ("IIf"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("x"),
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
                                        IdentifierExpression {
                                            Identifier ("x"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
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
    fn iif_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print IIf(success, "Success", "Failure")
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
                        Identifier ("IIf"),
                        LeftParenthesis,
                        Identifier ("success"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Success\""),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Failure\""),
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
    fn iif_function_argument() {
        let source = r"
Sub Test()
    Call ProcessValue(IIf(isActive, currentValue, 0))
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
                        Identifier ("IIf"),
                        LeftParenthesis,
                        Identifier ("isActive"),
                        Comma,
                        Whitespace,
                        Identifier ("currentValue"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("0"),
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
    fn iif_return_value() {
        let source = r"
Function GetMax(a As Integer, b As Integer) As Integer
    GetMax = IIf(a > b, a, b)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetMax"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("a"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("b"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetMax"),
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
                                        IdentifierExpression {
                                            Identifier ("a"),
                                        },
                                        Whitespace,
                                        GreaterThanOperator,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("b"),
                                        },
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("a"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("b"),
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
    fn iif_with_strings() {
        let source = r#"
Sub Test()
    greeting = "Hello " & IIf(isMorning, "Good morning", "Good evening")
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
                            Identifier ("greeting"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"Hello \""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("IIf"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("isMorning"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        StringLiteralExpression {
                                            StringLiteral ("\"Good morning\""),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        StringLiteralExpression {
                                            StringLiteral ("\"Good evening\""),
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
    fn iif_multiple_in_expression() {
        let source = r"
Sub Test()
    total = IIf(a > 0, a, 0) + IIf(b > 0, b, 0)
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
                        BinaryExpression {
                            CallExpression {
                                Identifier ("IIf"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("a"),
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
                                        IdentifierExpression {
                                            Identifier ("a"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("0"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("IIf"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("b"),
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
                                        IdentifierExpression {
                                            Identifier ("b"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("0"),
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
    fn iif_class_member() {
        let source = r"
Private Sub Class_Initialize()
    m_value = IIf(IsNull(initialValue), 0, initialValue)
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
                            Identifier ("m_value"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("IIf"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("IsNull"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("initialValue"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("0"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("initialValue"),
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
    fn iif_with_not() {
        let source = r#"
Sub Test()
    result = IIf(Not isEmpty, data, "")
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
                                    UnaryExpression {
                                        NotKeyword,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("isEmpty"),
                                        },
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("data"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"\""),
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
    fn iif_collection_add() {
        let source = r#"
Sub Test()
    col.Add IIf(useKey, item, Empty), IIf(useKey, key, "")
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
                        Identifier ("col"),
                        PeriodOperator,
                        Identifier ("Add"),
                        Whitespace,
                        Identifier ("IIf"),
                        LeftParenthesis,
                        Identifier ("useKey"),
                        Comma,
                        Whitespace,
                        Identifier ("item"),
                        Comma,
                        Whitespace,
                        EmptyKeyword,
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        Identifier ("IIf"),
                        LeftParenthesis,
                        Identifier ("useKey"),
                        Comma,
                        Whitespace,
                        Identifier ("key"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"\""),
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
    fn iif_with_comparison() {
        let source = r"
Sub Test()
    isValid = (IIf(value <> 0, value, 1) > threshold)
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
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("IIf"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    Identifier ("value"),
                                                },
                                                Whitespace,
                                                InequalityOperator,
                                                Whitespace,
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("0"),
                                                },
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("value"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("1"),
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
    fn iif_ternary_style() {
        let source = r"
Function Sign(n As Double) As Integer
    Sign = IIf(n > 0, 1, IIf(n < 0, -1, 0))
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("Sign"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("n"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("Sign"),
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
                                        IdentifierExpression {
                                            Identifier ("n"),
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
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    CallExpression {
                                        Identifier ("IIf"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                BinaryExpression {
                                                    IdentifierExpression {
                                                        Identifier ("n"),
                                                    },
                                                    Whitespace,
                                                    LessThanOperator,
                                                    Whitespace,
                                                    NumericLiteralExpression {
                                                        IntegerLiteral ("0"),
                                                    },
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                UnaryExpression {
                                                    SubtractionOperator,
                                                    NumericLiteralExpression {
                                                        IntegerLiteral ("1"),
                                                    },
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("0"),
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
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn iif_with_error_handling() {
        let source = r"
Sub Test()
    On Error Resume Next
    value = IIf(Err.Number = 0, result, defaultValue)
    On Error GoTo 0
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("value"),
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
                                        MemberAccessExpression {
                                            Identifier ("Err"),
                                            PeriodOperator,
                                            Identifier ("Number"),
                                        },
                                        Whitespace,
                                        EqualityOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("0"),
                                        },
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("defaultValue"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
}
