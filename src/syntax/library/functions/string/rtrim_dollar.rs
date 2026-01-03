//! # `RTrim$` Function
//!
//! The `RTrim$` function in Visual Basic 6 returns a string with trailing (right-side) spaces
//! removed. The dollar sign (`$`) suffix indicates that this function always returns a `String`
//! type, never a `Variant`.
//!
//! ## Syntax
//!
//! ```vb6
//! RTrim$(string)
//! ```
//!
//! ## Parameters
//!
//! - `string` - Required. Any valid string expression. If `string` contains `Null`, `Null` is returned.
//!
//! ## Return Value
//!
//! Returns a `String` with all trailing space characters (ASCII 32) removed from `string`.
//!
//! ## Behavior and Characteristics
//!
//! ### Space Removal
//!
//! - Removes only trailing spaces (ASCII character 32)
//! - Does not remove leading spaces (use `LTrim$` for that)
//! - Does not remove tabs, newlines, or other whitespace characters
//! - If the string contains only spaces, returns an empty string ("")
//! - Preserves spaces in the middle of the string
//!
//! ### Type Differences: `RTrim$` vs `RTrim`
//!
//! - `RTrim$`: Always returns `String` type (never `Variant`)
//! - `RTrim`: Returns `Variant` (can propagate `Null` values)
//! - Use `RTrim$` when you need guaranteed `String` return type
//! - Use `RTrim` when working with potentially `Null` values
//!
//! ## Common Usage Patterns
//!
//! ### 1. Clean User Input
//!
//! ```vb6
//! Function CleanInput(userInput As String) As String
//!     CleanInput = RTrim$(userInput)
//! End Function
//!
//! Dim cleaned As String
//! cleaned = CleanInput("  Hello World  ")  ' Returns "  Hello World"
//! ```
//!
//! ### 2. Format Output for Display
//!
//! ```vb6
//! Sub DisplayData()
//!     Dim dataField As String
//!     dataField = "Value    "
//!     Debug.Print "|" & RTrim$(dataField) & "|"  ' Prints "|Value|"
//! End Sub
//! ```
//!
//! ### 3. Database Field Processing
//!
//! ```vb6
//! Function GetFieldValue(rs As Recordset, fieldName As String) As String
//!     ' Remove trailing spaces from fixed-width database fields
//!     GetFieldValue = RTrim$(rs.Fields(fieldName).Value & "")
//! End Function
//! ```
//!
//! ### 4. Fixed-Width Data Parsing
//!
//! ```vb6
//! Function ParseFixedField(dataLine As String, startPos As Integer, fieldWidth As Integer) As String
//!     Dim rawField As String
//!     rawField = Mid$(dataLine, startPos, fieldWidth)
//!     ParseFixedField = RTrim$(rawField)
//! End Function
//!
//! Dim name As String
//! name = ParseFixedField("John      Doe       ", 1, 10)  ' Returns "John"
//! ```
//!
//! ### 5. Clean File Content
//!
//! ```vb6
//! Function ReadCleanLine(fileNum As Integer) As String
//!     Dim rawLine As String
//!     Line Input #fileNum, rawLine
//!     ReadCleanLine = RTrim$(rawLine)
//! End Function
//! ```
//!
//! ### 6. String Comparison Preparation
//!
//! ```vb6
//! Function CompareValues(value1 As String, value2 As String) As Boolean
//!     ' Remove trailing spaces for accurate comparison
//!     CompareValues = (RTrim$(value1) = RTrim$(value2))
//! End Function
//! ```
//!
//! ### 7. Configuration Value Processing
//!
//! ```vb6
//! Function GetConfigValue(key As String) As String
//!     Dim rawValue As String
//!     rawValue = GetINIString("Settings", key, "")
//!     GetConfigValue = RTrim$(rawValue)
//! End Function
//! ```
//!
//! ### 8. Array Element Cleanup
//!
//! ```vb6
//! Sub CleanStringArray(arr() As String)
//!     Dim i As Integer
//!     For i = LBound(arr) To UBound(arr)
//!         arr(i) = RTrim$(arr(i))
//!     Next i
//! End Sub
//! ```
//!
//! ### 9. Report Generation
//!
//! ```vb6
//! Function FormatReportLine(label As String, value As String) As String
//!     Dim paddedLabel As String
//!     paddedLabel = label & Space(30)
//!     FormatReportLine = Left$(RTrim$(paddedLabel), 30) & value
//! End Function
//! ```
//!
//! ### 10. Logging and Debug Output
//!
//! ```vb6
//! Sub LogMessage(message As String)
//!     Dim timestamp As String
//!     Dim cleanMsg As String
//!     timestamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
//!     cleanMsg = RTrim$(message)
//!     Debug.Print timestamp & " - " & cleanMsg
//! End Sub
//! ```
//!
//! ## Related Functions
//!
//! - `RTrim()` - Returns a `Variant` with trailing spaces removed (can handle `Null`)
//! - `LTrim$()` - Removes leading (left-side) spaces from a string
//! - `Trim$()` - Removes both leading and trailing spaces from a string
//! - `Left$()` - Returns a specified number of characters from the left side
//! - `Right$()` - Returns a specified number of characters from the right side
//! - `Space$()` - Creates a string consisting of the specified number of spaces
//! - `Len()` - Returns the length of a string
//!
//! ## Best Practices
//!
//! ### When to Use `RTrim$` vs `RTrim`
//!
//! ```vb6
//! ' Use RTrim$ when you need a String
//! Dim cleaned As String
//! cleaned = RTrim$(userInput)  ' Type-safe, always returns String
//!
//! ' use RTrim when working with Variants or Null values
//! Dim result As Variant
//! result = RTrim(variantValue)  ' Can propagate Null
//! ```
//!
//! ### Combine with `LTrim$` for Full Cleanup
//!
//! ```vb6
//! ' Remove both leading and trailing spaces
//! Dim fullyClean As String
//! fullyClean = LTrim$(RTrim$(input))
//!
//! ' Or use Trim$ for convenience
//! fullyClean = Trim$(input)
//! ```
//!
//! ### Use for Fixed-Width Fields
//!
//! ```vb6
//! ' Clean up fixed-width database or file fields
//! Dim firstName As String
//! firstName = RTrim$(rs!FirstName)  ' Remove padding spaces
//! ```
//!
//! ### Validate Before Processing
//!
//! ```vb6
//! Function SafeRTrim(value As Variant) As String
//!     If IsNull(value) Then
//!         SafeRTrim = ""
//!     Else
//!         SafeRTrim = RTrim$(CStr(value))
//!     End If
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - `RTrim$` is very efficient and lightweight
//! - Performs a single pass from the end of the string
//! - More efficient than manually removing spaces with loops
//! - No performance penalty for strings without trailing spaces
//!
//! ```vb6
//! ' Efficient: single RTrim$ call
//! Dim cleaned As String
//! cleaned = RTrim$(input)
//!
//! ' Less efficient: manual space removal
//! Dim i As Integer
//! For i = Len(input) To 1 Step -1
//!     If Mid$(input, i, 1) <> " " Then Exit For
//! Next i
//! cleaned = Left$(input, i)
//! ```
//!
//! ## Common Pitfalls
//!
//! ### 1. Only Removes Spaces (ASCII 32)
//!
//! ```vb6
//! Dim text As String
//! text = "Hello" & vbTab  ' Ends with tab character
//!
//! ' RTrim$ does NOT remove tabs
//! Debug.Print RTrim$(text)  ' Still has the tab at the end
//!
//! ' To remove all whitespace, you need custom logic
//! Function RemoveTrailingWhitespace(s As String) As String
//!     Dim i As Integer
//!     For i = Len(s) To 1 Step -1
//!         Select Case Mid$(s, i, 1)
//!             Case " ", vbTab, vbCr, vbLf
//!                 ' Continue
//!             Case Else
//!                 Exit For
//!         End Select
//!     Next i
//!     RemoveTrailingWhitespace = Left$(s, i)
//! End Function
//! ```
//!
//! ### 2. Null Value Handling
//!
//! ```vb6
//! ' RTrim$ with Null causes runtime error
//! Dim result As String
//! result = RTrim$(nullValue)  ' ERROR if nullValue is Null
//!
//! ' Protect against Null
//! If Not IsNull(value) Then
//!     result = RTrim$(value)
//! Else
//!     result = ""
//! End If
//! ```
//!
//! ### 3. Confusing with `Trim$`
//!
//! ```vb6
//! Dim text As String
//! text = "  Hello  "
//!
//! Debug.Print RTrim$(text)   ' "  Hello" (leading spaces remain)
//! Debug.Print LTrim$(text)   ' "Hello  " (trailing spaces remain)
//! Debug.Print Trim$(text)    ' "Hello" (both removed)
//! ```
//!
//! ### 4. Database Field Assumptions
//!
//! ```vb6
//! ' Wrong: assuming all database fields need RTrim
//! value = RTrim$(rs!TextField)  ' May error if field is Null
//!
//! ' Better: handle Null and empty values
//! If IsNull(rs!TextField) Then
//!     value = ""
//! Else
//!     value = RTrim$(rs!TextField & "")
//! End If
//! ```
//!
//! ### 5. Not Checking for Empty Results
//!
//! ```vb6
//! Dim input As String
//! input = "     "  ' Only spaces
//!
//! Dim result As String
//! result = RTrim$(input)  ' Returns "" (empty string)
//!
//! ' Check if result is meaningful
//! If Len(RTrim$(input)) > 0 Then
//!     ' Process non-empty string
//! End If
//! ```
//!
//! ## Limitations
//!
//! - Only removes space characters (ASCII 32), not other whitespace
//! - Cannot handle `Null` values (use `RTrim` variant function instead)
//! - Does not remove leading spaces (use `LTrim$` or `Trim$`)
//! - No option to specify custom characters to remove
//! - Works with strings only, not byte arrays
//! - Does not trim non-breaking spaces (character 160) or other Unicode whitespace

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn rtrim_dollar_simple() {
        let source = r#"
Sub Main()
    result = RTrim$("Hello   ")
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RTrim$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Hello   \""),
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
    fn rtrim_dollar_assignment() {
        let source = r"
Sub Main()
    Dim cleaned As String
    cleaned = RTrim$(userInput)
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
                        Identifier ("cleaned"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("cleaned"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RTrim$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("userInput"),
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
    fn rtrim_dollar_variable() {
        let source = r#"
Sub Main()
    Dim text As String
    Dim result As String
    text = "Sample  "
    result = RTrim$(text)
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
                            TextKeyword,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"Sample  \""),
                        },
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
                            Identifier ("RTrim$"),
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
    fn rtrim_dollar_display_format() {
        let source = r#"
Sub DisplayData()
    Dim dataField As String
    dataField = "Value    "
    Debug.Print "|" & RTrim$(dataField) & "|"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("DisplayData"),
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
                        Identifier ("dataField"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("dataField"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"Value    \""),
                        },
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        StringLiteral ("\"|\""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("RTrim$"),
                        LeftParenthesis,
                        Identifier ("dataField"),
                        RightParenthesis,
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"|\""),
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
    fn rtrim_dollar_database_field() {
        let source = r"
Function GetFieldValue(fieldValue As String) As String
    GetFieldValue = RTrim$(fieldValue)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetFieldValue"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("fieldValue"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
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
                            Identifier ("GetFieldValue"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RTrim$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("fieldValue"),
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
    fn rtrim_dollar_in_condition() {
        let source = r#"
Sub Main()
    If RTrim$(dataValue) = "Expected" Then
        Debug.Print "Match found"
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
                                Identifier ("RTrim$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("dataValue"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"Expected\""),
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
                                StringLiteral ("\"Match found\""),
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
    fn rtrim_dollar_fixed_width() {
        let source = r"
Function ParseFixedField(dataLine As String, startPos As Integer, fieldWidth As Integer) As String
    Dim rawField As String
    rawField = Mid$(dataLine, startPos, fieldWidth)
    ParseFixedField = RTrim$(rawField)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ParseFixedField"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("dataLine"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("startPos"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
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
                        Identifier ("rawField"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("rawField"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Mid$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("dataLine"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("startPos"),
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("ParseFixedField"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RTrim$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("rawField"),
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
    fn rtrim_dollar_comparison() {
        let source = r"
Function CompareValues(value1 As String, value2 As String) As Boolean
    CompareValues = (RTrim$(value1) = RTrim$(value2))
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CompareValues"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("value1"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("value2"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                BooleanKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("CompareValues"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("RTrim$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("value1"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("RTrim$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("value2"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
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
    fn rtrim_dollar_array_cleanup() {
        let source = r"
Sub CleanStringArray(arr() As String)
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        arr(i) = RTrim$(arr(i))
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
                Identifier ("CleanStringArray"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("arr"),
                    LeftParenthesis,
                    RightParenthesis,
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
                        CallExpression {
                            Identifier ("LBound"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("arr"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        CallExpression {
                            Identifier ("UBound"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("arr"),
                                    },
                                },
                            },
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
                                    Identifier ("RTrim$"),
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
    fn rtrim_dollar_multiple_uses() {
        let source = r"
Sub ProcessData()
    Dim firstName As String
    Dim lastName As String
    firstName = RTrim$(rawFirst)
    lastName = RTrim$(rawLast)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ProcessData"),
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
                        Identifier ("firstName"),
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
                        Identifier ("lastName"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("firstName"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RTrim$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("rawFirst"),
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
                            Identifier ("lastName"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RTrim$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("rawLast"),
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
    fn rtrim_dollar_select_case() {
        let source = r#"
Sub Main()
    Select Case RTrim$(status)
        Case "Active"
            Debug.Print "Active record"
        Case "Inactive"
            Debug.Print "Inactive record"
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
                            Identifier ("RTrim$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("status"),
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
                            StringLiteral ("\"Active\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Active record\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\"Inactive\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Inactive record\""),
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
    fn rtrim_dollar_concatenation() {
        let source = r#"
Sub Main()
    Dim output As String
    output = "Name: " & RTrim$(nameField)
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
                            StringLiteralExpression {
                                StringLiteral ("\"Name: \""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("RTrim$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("nameField"),
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
    fn rtrim_dollar_with_ltrim() {
        let source = r"
Sub Main()
    Dim fullyClean As String
    fullyClean = LTrim$(RTrim$(input))
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
                        Identifier ("fullyClean"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("fullyClean"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LTrim$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("RTrim$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    InputKeyword,
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
    fn rtrim_dollar_report_format() {
        let source = r"
Function FormatReportLine(textLabel As String, value As String) As String
    Dim paddedLabel As String
    paddedLabel = textLabel & Space(30)
    FormatReportLine = Left$(RTrim$(paddedLabel), 30) & value
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("FormatReportLine"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("textLabel"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("value"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
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
                        Identifier ("paddedLabel"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("paddedLabel"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("textLabel"),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Space"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("30"),
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
                            Identifier ("FormatReportLine"),
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
                                        CallExpression {
                                            Identifier ("RTrim$"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    IdentifierExpression {
                                                        Identifier ("paddedLabel"),
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
                                            IntegerLiteral ("30"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("value"),
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
    fn rtrim_dollar_logging() {
        let source = r#"
Sub LogMessage(message As String)
    Dim cleanMsg As String
    cleanMsg = RTrim$(message)
    Debug.Print Now & " - " & cleanMsg
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("LogMessage"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("message"),
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
                        Identifier ("cleanMsg"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("cleanMsg"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RTrim$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("message"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("Now"),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\" - \""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("cleanMsg"),
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
    fn rtrim_dollar_in_function() {
        let source = r"
Function CleanInput(userInput As String) As String
    CleanInput = RTrim$(userInput)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CleanInput"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("userInput"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
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
                            Identifier ("CleanInput"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RTrim$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("userInput"),
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
    fn rtrim_dollar_config_value() {
        let source = r#"
Function GetConfigValue(key As String) As String
    Dim rawValue As String
    rawValue = GetINIString("Settings", key, "")
    GetConfigValue = RTrim$(rawValue)
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetConfigValue"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("key"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
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
                        Identifier ("rawValue"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("rawValue"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("GetINIString"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Settings\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("key"),
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetConfigValue"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RTrim$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("rawValue"),
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
    fn rtrim_dollar_empty_check() {
        let source = r#"
Sub Main()
    If Len(RTrim$(input)) > 0 Then
        Debug.Print "Has content"
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
                                            Identifier ("RTrim$"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    IdentifierExpression {
                                                        InputKeyword,
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
                                StringLiteral ("\"Has content\""),
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
    fn rtrim_dollar_file_processing() {
        let source = r"
Function ReadCleanLine(fileNum As Integer) As String
    Dim rawLine As String
    Line Input #fileNum, rawLine
    ReadCleanLine = RTrim$(rawLine)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ReadCleanLine"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("fileNum"),
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
                        Identifier ("rawLine"),
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
                        Identifier ("fileNum"),
                        Comma,
                        Whitespace,
                        Identifier ("rawLine"),
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("ReadCleanLine"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RTrim$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("rawLine"),
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
    fn rtrim_dollar_loop_processing() {
        let source = r"
Sub ProcessLines()
    Dim i As Integer
    Dim cleanLine As String
    For i = 1 To 10
        cleanLine = RTrim$(lines(i))
        Debug.Print cleanLine
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
                Identifier ("ProcessLines"),
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("cleanLine"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
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
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("cleanLine"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("RTrim$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            CallExpression {
                                                Identifier ("lines"),
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
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                Identifier ("cleanLine"),
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
