//! # `Space$` Function
//!
//! The `Space$` function in Visual Basic 6 returns a string consisting of the specified number
//! of space characters (ASCII 32). The dollar sign (`$`) suffix indicates that this function
//! always returns a `String` type, never a `Variant`.
//!
//! ## Syntax
//!
//! ```vb6
//! Space$(number)
//! ```
//!
//! ## Parameters
//!
//! - `number` - Required. Long integer indicating the number of spaces to include in the string.
//!   Must be between 0 and approximately 2 billion (2,147,483,647).
//!
//! ## Return Value
//!
//! Returns a `String` containing `number` space characters.
//!
//! ## Behavior and Characteristics
//!
//! ### Number Handling
//!
//! - If `number` = 0: Returns an empty string ("")
//! - If `number` < 0: Generates a runtime error (Invalid procedure call or argument)
//! - If `number` is a fractional value: Rounds to the nearest integer
//! - Maximum value is limited by available memory
//!
//! ### Type Differences: `Space$` vs `Space`
//!
//! - `Space$`: Always returns `String` type (never `Variant`)
//! - `Space`: Returns `Variant` containing a string
//! - Use `Space$` when you need guaranteed `String` return type
//! - Use `Space` when working with `Variant` variables
//!
//! ## Common Usage Patterns
//!
//! ### 1. Padding Strings
//!
//! ```vb6
//! Function PadRight(text As String, width As Integer) As String
//!     If Len(text) >= width Then
//!         PadRight = text
//!     Else
//!         PadRight = text & Space$(width - Len(text))
//!     End If
//! End Function
//!
//! Dim padded As String
//! padded = PadRight("Hello", 10)  ' "Hello     "
//! ```
//!
//! ### 2. Creating Fixed-Width Fields
//!
//! ```vb6
//! Function FormatField(value As String, fieldWidth As Integer) As String
//!     Dim temp As String
//!     temp = value & Space$(fieldWidth)
//!     FormatField = Left$(temp, fieldWidth)
//! End Function
//!
//! Dim field As String
//! field = FormatField("Name", 20)  ' "Name                "
//! ```
//!
//! ### 3. Indentation
//!
//! ```vb6
//! Function IndentText(text As String, level As Integer) As String
//!     IndentText = Space$(level * 4) & text
//! End Function
//!
//! Debug.Print IndentText("Nested Item", 2)  ' "        Nested Item"
//! ```
//!
//! ### 4. Column Alignment in Reports
//!
//! ```vb6
//! Sub PrintReport()
//!     Dim col1 As String, col2 As String, col3 As String
//!     col1 = "Name"
//!     col2 = "Age"
//!     col3 = "City"
//!     Debug.Print col1 & Space$(15) & col2 & Space$(10) & col3
//! End Sub
//! ```
//!
//! ### 5. Creating Separator Lines
//!
//! ```vb6
//! Function CreateSeparator(width As Integer, char As String) As String
//!     ' Create base with spaces then replace
//!     CreateSeparator = String$(width, char)
//! End Function
//!
//! ' Or use spaces for visual separation
//! Debug.Print "Header" & Space$(10) & "Value"
//! ```
//!
//! ### 6. Text Centering
//!
//! ```vb6
//! Function CenterText(text As String, width As Integer) As String
//!     Dim padding As Integer
//!     If Len(text) >= width Then
//!         CenterText = Left$(text, width)
//!     Else
//!         padding = (width - Len(text)) \ 2
//!         CenterText = Space$(padding) & text & Space$(width - Len(text) - padding)
//!     End If
//! End Function
//!
//! Dim centered As String
//! centered = CenterText("Title", 20)
//! ```
//!
//! ### 7. Creating Empty Buffers
//!
//! ```vb6
//! Function CreateBuffer(size As Integer) As String
//!     CreateBuffer = Space$(size)
//! End Function
//!
//! Dim buffer As String
//! buffer = CreateBuffer(100)  ' 100-character buffer
//! ```
//!
//! ### 8. Formatting Tables
//!
//! ```vb6
//! Sub PrintTableRow(col1 As String, col2 As String, col3 As String)
//!     Dim row As String
//!     row = Left$(col1 & Space$(20), 20) & _
//!           Left$(col2 & Space$(15), 15) & _
//!           Left$(col3 & Space$(10), 10)
//!     Debug.Print row
//! End Sub
//! ```
//!
//! ### 9. Creating Blank Lines
//!
//! ```vb6
//! Sub AddVerticalSpace(lines As Integer)
//!     Dim i As Integer
//!     For i = 1 To lines
//!         Debug.Print Space$(0)  ' Or just Debug.Print
//!     Next i
//! End Sub
//! ```
//!
//! ### 10. Formatting Currency Values
//!
//! ```vb6
//! Function FormatAmount(amount As Currency) As String
//!     Dim amountStr As String
//!     amountStr = Format$(amount, "#,##0.00")
//!     FormatAmount = Space$(15 - Len(amountStr)) & amountStr
//! End Function
//!
//! Debug.Print FormatAmount(1234.56)  ' Right-aligned in 15 chars
//! ```
//!
//! ## Related Functions
//!
//! - `Space()` - Returns a `Variant` containing the specified number of spaces
//! - `String$()` - Returns a string of repeating characters (more flexible than `Space$`)
//! - `Left$()` - Returns a specified number of characters from the left side
//! - `Right$()` - Returns a specified number of characters from the right side
//! - `Len()` - Returns the length of a string
//! - `LTrim$()` - Removes leading spaces from a string
//! - `RTrim$()` - Removes trailing spaces from a string
//! - `Trim$()` - Removes both leading and trailing spaces
//! - `Spc()` - Used in `Print` statements to insert spaces
//!
//! ## Best Practices
//!
//! ### When to Use `Space$` vs `String$`
//!
//! ```vb6
//! ' Use Space$ for creating spaces specifically
//! Dim spaces As String
//! spaces = Space$(10)  ' Clear intent
//!
//! ' Use String$ for other repeated characters
//! Dim dashes As String
//! dashes = String$(10, "-")  ' More flexible
//!
//! ' Note: Space$(10) is equivalent to String$(10, " ")
//! ```
//!
//! ### Prefer Constants for Fixed Padding
//!
//! ```vb6
//! ' Less efficient: creating spaces repeatedly
//! For i = 1 To 1000
//!     Debug.Print "Item" & Space$(10) & values(i)
//! Next i
//!
//! ' More efficient: create once
//! Const PADDING As String = "          "  ' 10 spaces
//! For i = 1 To 1000
//!     Debug.Print "Item" & PADDING & values(i)
//! Next i
//! ```
//!
//! ### Validate Negative Values
//!
//! ```vb6
//! Function SafeSpace(count As Integer) As String
//!     If count < 0 Then
//!         SafeSpace = ""
//!     Else
//!         SafeSpace = Space$(count)
//!     End If
//! End Function
//! ```
//!
//! ### Combine with `Left$` or `Right$` for Fixed Width
//!
//! ```vb6
//! ' Ensure exact width regardless of input length
//! Function FixedWidth(text As String, width As Integer) As String
//!     Dim temp As String
//!     temp = text & Space$(width)
//!     FixedWidth = Left$(temp, width)
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - `Space$` is optimized for creating space-filled strings
//! - Very efficient for small to moderate numbers of spaces (< 1000)
//! - For large numbers, consider memory implications
//! - Reuse space strings when possible instead of recreating
//!
//! ```vb6
//! ' Inefficient: creating space string in loop
//! For i = 1 To 10000
//!     output = output & Space$(5) & data(i)
//! Next i
//!
//! ' More efficient: create once
//! Dim spacer As String
//! spacer = Space$(5)
//! For i = 1 To 10000
//!     output = output & spacer & data(i)
//! Next i
//!
//! ' Even better: use array and Join
//! Dim parts() As String
//! ReDim parts(1 To 10000)
//! For i = 1 To 10000
//!     parts(i) = spacer & data(i)
//! Next i
//! output = Join(parts, "")
//! ```
//!
//! ## Common Pitfalls
//!
//! ### 1. Negative Values Cause Errors
//!
//! ```vb6
//! ' Runtime error: Invalid procedure call or argument
//! result = Space$(-5)  ' ERROR!
//!
//! ' Always validate
//! If count >= 0 Then
//!     result = Space$(count)
//! Else
//!     result = ""
//! End If
//! ```
//!
//! ### 2. Confusion with `Spc()` Function
//!
//! ```vb6
//! ' Space$ returns a string
//! Dim s As String
//! s = Space$(5)  ' Returns "     " (5 spaces)
//!
//! ' Spc is used in Print statements only
//! Debug.Print "A"; Spc(5); "B"  ' Prints "A     B"
//!
//! ' Cannot assign Spc to variable
//! s = Spc(5)  ' ERROR!
//! ```
//!
//! ### 3. Memory Issues with Large Values
//!
//! ```vb6
//! ' Be careful with very large values
//! Dim huge As String
//! huge = Space$(10000000)  ' 10 million spaces = ~20 MB
//!
//! ' Consider if you really need that many spaces
//! ' Often there are better alternatives
//! ```
//!
//! ### 4. Not Accounting for Existing Length
//!
//! ```vb6
//! ' Wrong: may create string longer than desired width
//! result = text & Space$(width)
//!
//! ' Correct: ensure exact width
//! temp = text & Space$(width)
//! result = Left$(temp, width)
//! ```
//!
//! ### 5. Using for Non-Space Padding
//!
//! ```vb6
//! ' Wrong: Space$ only creates spaces
//! underline = Space$(20)  ' Trying to create underline
//! ' This creates "                    ", not "____________________"
//!
//! ' Correct: use String$ for other characters
//! underline = String$(20, "_")
//! ```
//!
//! ### 6. Floating Point Rounding
//!
//! ```vb6
//! Debug.Print Space$(5.4)   ' Creates 5 spaces (rounds down)
//! Debug.Print Space$(5.6)   ' Creates 6 spaces (rounds up)
//! Debug.Print Space$(5.5)   ' Creates 6 spaces (banker's rounding)
//!
//! ' Be explicit with integer conversion if needed
//! Dim count As Integer
//! count = Int(5.7)  ' Force truncation
//! result = Space$(count)
//! ```
//!
//! ## Practical Examples
//!
//! ### Creating Fixed-Width Reports
//!
//! ```vb6
//! Sub PrintInvoice()
//!     Dim header As String
//!     Dim separator As String
//!     
//!     ' Create header with aligned columns
//!     header = Left$("Item" & Space$(30), 30) & _
//!              Left$("Qty" & Space$(10), 10) & _
//!              Left$("Price" & Space$(15), 15)
//!     
//!     separator = String$(55, "-")
//!     
//!     Debug.Print header
//!     Debug.Print separator
//!     Debug.Print Left$("Widget" & Space$(30), 30) & _
//!                 Right$(Space$(10) & "5", 10) & _
//!                 Right$(Space$(15) & "$10.00", 15)
//! End Sub
//! ```
//!
//! ### Building Hierarchical Output
//!
//! ```vb6
//! Sub PrintTree(text As String, level As Integer)
//!     Debug.Print Space$(level * 2) & "- " & text
//! End Sub
//!
//! PrintTree("Root", 0)
//! PrintTree("Child 1", 1)
//! PrintTree("Grandchild", 2)
//! PrintTree("Child 2", 1)
//! ```
//!
//! ### Formatting Data for Fixed-Width Files
//!
//! ```vb6
//! Function FormatRecord(name As String, age As Integer, city As String) As String
//!     Dim record As String
//!     record = Left$(name & Space$(25), 25) & _
//!              Right$(Space$(3) & CStr(age), 3) & _
//!              Left$(city & Space$(20), 20)
//!     FormatRecord = record
//! End Function
//!
//! ' Writes: "John Doe              25New York            "
//! Print #1, FormatRecord("John Doe", 25, "New York")
//! ```
//!
//! ## Limitations
//!
//! - Can only create space characters (ASCII 32), not other whitespace
//! - Negative values cause runtime errors
//! - Very large values can cause out-of-memory errors
//! - Cannot be used directly in `Print` statements like `Spc()`
//! - Floating-point parameters are rounded (may be unexpected)
//! - Maximum string length limited by VB6 string constraints (~2 GB)

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn space_dollar_simple() {
        let source = r"
Sub Main()
    result = Space$(10)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
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
                            Identifier ("Space$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("10"),
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
    fn space_dollar_assignment() {
        let source = r"
Sub Main()
    Dim padding As String
    padding = Space$(5)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
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
                        Identifier ("padding"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("padding"),
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
                                        IntegerLiteral ("5"),
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
    fn space_dollar_variable() {
        let source = r"
Sub Main()
    Dim count As Integer
    Dim spaces As String
    count = 20
    spaces = Space$(count)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
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
                        Identifier ("count"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("spaces"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("count"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("20"),
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("spaces"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Space$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("count"),
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
    fn space_dollar_padding() {
        let source = r"
Function PadRight(text As String, width As Integer) As String
    If Len(text) >= width Then
        PadRight = text
    Else
        PadRight = text & Space$(width - Len(text))
    End If
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("PadRight"),
                ParameterList {
                    LeftParenthesis,
                },
                TextKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Comma,
                Whitespace,
                WidthKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                LenKeyword,
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
                            GreaterThanOrEqualOperator,
                            Whitespace,
                            IdentifierExpression {
                                WidthKeyword,
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("PadRight"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                IdentifierExpression {
                                    TextKeyword,
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        ElseClause {
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("PadRight"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    BinaryExpression {
                                        IdentifierExpression {
                                            TextKeyword,
                                        },
                                        Whitespace,
                                        Ampersand,
                                        Whitespace,
                                        CallExpression {
                                            Identifier ("Space$"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    BinaryExpression {
                                                        IdentifierExpression {
                                                            WidthKeyword,
                                                        },
                                                        Whitespace,
                                                        SubtractionOperator,
                                                        Whitespace,
                                                        CallExpression {
                                                            LenKeyword,
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
                                                    },
                                                },
                                            },
                                            RightParenthesis,
                                        },
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
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
    fn space_dollar_fixed_width() {
        let source = r"
Function FormatField(value As String, fieldWidth As Integer) As String
    Dim temp As String
    temp = value & Space$(fieldWidth)
    FormatField = Left$(temp, fieldWidth)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("FormatField"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("value"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("fieldWidth"),
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
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("temp"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("temp"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("value"),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Space$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("fieldWidth"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("FormatField"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Left$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("temp"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("fieldWidth"),
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
    fn space_dollar_indentation() {
        let source = r"
Function IndentText(text As String, level As Integer) As String
    IndentText = Space$(level * 4) & text
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("IndentText"),
                ParameterList {
                    LeftParenthesis,
                },
                TextKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Comma,
                Whitespace,
                Identifier ("level"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("IndentText"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Space$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("level"),
                                            },
                                            Whitespace,
                                            MultiplicationOperator,
                                            Whitespace,
                                            NumericLiteralExpression {
                                                IntegerLiteral ("4"),
                                            },
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            IdentifierExpression {
                                TextKeyword,
                            },
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
    fn space_dollar_in_condition() {
        let source = r#"
Sub Main()
    If Len(Space$(count)) > 0 Then
        Debug.Print "Has spaces"
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
                Identifier ("Main"),
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
                                LenKeyword,
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        CallExpression {
                                            Identifier ("Space$"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    IdentifierExpression {
                                                        Identifier ("count"),
                                                    },
                                                },
                                            },
                                            RightParenthesis,
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
                                StringLiteral ("\"Has spaces\""),
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
    fn space_dollar_centering() {
        let source = r"
Function CenterText(text As String, width As Integer) As String
    Dim padding As Integer
    If Len(text) >= width Then
        CenterText = Left$(text, width)
    Else
        padding = (width - Len(text)) \ 2
        CenterText = Space$(padding) & text
    End If
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CenterText"),
                ParameterList {
                    LeftParenthesis,
                },
                TextKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Comma,
                Whitespace,
                WidthKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("padding"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                LenKeyword,
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
                            GreaterThanOrEqualOperator,
                            Whitespace,
                            IdentifierExpression {
                                WidthKeyword,
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("CenterText"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Left$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                TextKeyword,
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            IdentifierExpression {
                                                WidthKeyword,
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        ElseClause {
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("padding"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    BinaryExpression {
                                        ParenthesizedExpression {
                                            LeftParenthesis,
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    WidthKeyword,
                                                },
                                                Whitespace,
                                                SubtractionOperator,
                                                Whitespace,
                                                CallExpression {
                                                    LenKeyword,
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
                                            },
                                            RightParenthesis,
                                        },
                                        Whitespace,
                                        BackwardSlashOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("2"),
                                        },
                                    },
                                    Newline,
                                },
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("CenterText"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    BinaryExpression {
                                        CallExpression {
                                            Identifier ("Space$"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    IdentifierExpression {
                                                        Identifier ("padding"),
                                                    },
                                                },
                                            },
                                            RightParenthesis,
                                        },
                                        Whitespace,
                                        Ampersand,
                                        Whitespace,
                                        IdentifierExpression {
                                            TextKeyword,
                                        },
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
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
    fn space_dollar_buffer_creation() {
        let source = r"
Function CreateBuffer(size As Integer) As String
    CreateBuffer = Space$(size)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CreateBuffer"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("size"),
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
                            Identifier ("CreateBuffer"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Space$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("size"),
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
    fn space_dollar_multiple_uses() {
        let source = r#"
Sub PrintReport()
    Debug.Print "Name" & Space$(15) & "Age" & Space$(10) & "City"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("PrintReport"),
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
                        StringLiteral ("\"Name\""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Space$"),
                        LeftParenthesis,
                        IntegerLiteral ("15"),
                        RightParenthesis,
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"Age\""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Space$"),
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"City\""),
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
    fn space_dollar_select_case() {
        let source = r#"
Sub Main()
    Select Case Len(Space$(width))
        Case 0
            Debug.Print "Empty"
        Case Else
            Debug.Print "Has spaces"
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
                Identifier ("Main"),
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
                            LenKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Space$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    WidthKeyword,
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
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("0"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Empty\""),
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
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Has spaces\""),
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
    fn space_dollar_concatenation() {
        let source = r#"
Sub Main()
    Dim output As String
    output = "Value:" & Space$(5) & valueStr
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
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
                        OutputKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            OutputKeyword,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                StringLiteralExpression {
                                    StringLiteral ("\"Value:\""),
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Space$"),
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
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("valueStr"),
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
    fn space_dollar_table_formatting() {
        let source = r"
Sub PrintTableRow(col1 As String, col2 As String, col3 As String)
    Dim row As String
    row = Left$(col1 & Space$(20), 20) & Left$(col2 & Space$(15), 15)
    Debug.Print row
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("PrintTableRow"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("col1"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("col2"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("col3"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("row"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("row"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Left$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("col1"),
                                            },
                                            Whitespace,
                                            Ampersand,
                                            Whitespace,
                                            CallExpression {
                                                Identifier ("Space$"),
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        NumericLiteralExpression {
                                                            IntegerLiteral ("20"),
                                                        },
                                                    },
                                                },
                                                RightParenthesis,
                                            },
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("20"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Left$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("col2"),
                                            },
                                            Whitespace,
                                            Ampersand,
                                            Whitespace,
                                            CallExpression {
                                                Identifier ("Space$"),
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        NumericLiteralExpression {
                                                            IntegerLiteral ("15"),
                                                        },
                                                    },
                                                },
                                                RightParenthesis,
                                            },
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("15"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("row"),
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
    fn space_dollar_zero_spaces() {
        let source = r"
Sub Main()
    Dim empty As String
    empty = Space$(0)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
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
                        EmptyKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        LiteralExpression {
                            EmptyKeyword,
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
    fn space_dollar_expression_arg() {
        let source = r"
Sub Main()
    Dim result As String
    result = Space$(maxWidth - Len(text))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
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
                        Identifier ("result"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
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
                        CallExpression {
                            Identifier ("Space$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("maxWidth"),
                                        },
                                        Whitespace,
                                        SubtractionOperator,
                                        Whitespace,
                                        CallExpression {
                                            LenKeyword,
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
    fn space_dollar_format_amount() {
        let source = r##"
Function FormatAmount(amount As Currency) As String
    Dim amountStr As String
    amountStr = Format$(amount, "#,##0.00")
    FormatAmount = Space$(15 - Len(amountStr)) & amountStr
End Function
"##;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("FormatAmount"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("amount"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    CurrencyKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("amountStr"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("amountStr"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Format$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("amount"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"#,##0.00\""),
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
                            Identifier ("FormatAmount"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Space$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("15"),
                                            },
                                            Whitespace,
                                            SubtractionOperator,
                                            Whitespace,
                                            CallExpression {
                                                LenKeyword,
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        IdentifierExpression {
                                                            Identifier ("amountStr"),
                                                        },
                                                    },
                                                },
                                                RightParenthesis,
                                            },
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("amountStr"),
                            },
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
    fn space_dollar_hierarchical_output() {
        let source = r#"
Sub PrintTree(text As String, level As Integer)
    Debug.Print Space$(level * 2) & "- " & text
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("PrintTree"),
                ParameterList {
                    LeftParenthesis,
                },
                TextKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Comma,
                Whitespace,
                Identifier ("level"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                RightParenthesis,
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("Space$"),
                        LeftParenthesis,
                        Identifier ("level"),
                        Whitespace,
                        MultiplicationOperator,
                        Whitespace,
                        IntegerLiteral ("2"),
                        RightParenthesis,
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"- \""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        TextKeyword,
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
    fn space_dollar_loop_processing() {
        let source = r#"
Sub CreatePaddedList()
    Dim i As Integer
    For i = 1 To 10
        Debug.Print "Item" & Space$(5) & CStr(i)
    Next i
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("CreatePaddedList"),
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
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                StringLiteral ("\"Item\""),
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                Identifier ("Space$"),
                                LeftParenthesis,
                                IntegerLiteral ("5"),
                                RightParenthesis,
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                Identifier ("CStr"),
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
    fn space_dollar_in_function() {
        let source = r"
Function GetSpaces(count As Integer) As String
    GetSpaces = Space$(count)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetSpaces"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("count"),
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
                            Identifier ("GetSpaces"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Space$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("count"),
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
    fn space_dollar_with_left() {
        let source = r"
Function FixedWidth(text As String, width As Integer) As String
    Dim temp As String
    temp = text & Space$(width)
    FixedWidth = Left$(temp, width)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("FixedWidth"),
                ParameterList {
                    LeftParenthesis,
                },
                TextKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Comma,
                Whitespace,
                WidthKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("temp"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("temp"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                TextKeyword,
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Space$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            WidthKeyword,
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("FixedWidth"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Left$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("temp"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        WidthKeyword,
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
}
