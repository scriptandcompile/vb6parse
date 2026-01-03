//! # `LCase$` Function
//!
//! Returns a `String` that has been converted to lowercase.
//! The "$" suffix indicates this function returns a `String` type.
//!
//! ## Syntax
//!
//! ```vb
//! LCase$(string)
//! ```
//!
//! ## Parameters
//!
//! - **string**: Required. Any valid string expression. If `string` contains `Null`, `Null` is returned.
//!
//! ## Returns
//!
//! Returns a `String` with all uppercase letters converted to lowercase. Numbers and punctuation
//! are unchanged.
//!
//! ## Remarks
//!
//! - `LCase$` converts all uppercase letters in a string to lowercase.
//! - The "$" suffix explicitly indicates the function returns a `String` type rather than a `Variant`.
//! - Only uppercase letters (A-Z) are affected; lowercase letters and non-alphabetic characters remain unchanged.
//! - `LCase$` is functionally equivalent to `LCase`, but `LCase$` returns a `String` while `LCase` can return a `Variant`.
//! - For better performance when you know the result is a string, use `LCase$`.
//! - If the argument is `Null`, the function returns `Null`.
//! - The conversion is based on the system locale settings.
//! - For international characters, the behavior depends on the current code page.
//! - The inverse function is `UCase$`, which converts strings to uppercase.
//! - Common use cases include case-insensitive comparisons and normalizing user input.
//!
//! ## Typical Uses
//!
//! 1. **Case-insensitive comparisons** - Compare strings without regard to case
//! 2. **User input normalization** - Convert user input to a consistent case
//! 3. **Email address handling** - Normalize email addresses to lowercase
//! 4. **File path comparisons** - Compare file paths on case-insensitive file systems
//! 5. **Username validation** - Normalize usernames to lowercase
//! 6. **Search operations** - Perform case-insensitive searches
//! 7. **Data standardization** - Ensure consistent casing in databases
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Simple conversion
//! Dim result As String
//! result = LCase$("HELLO")  ' Returns "hello"
//! ```
//!
//! ```vb
//! ' Example 2: Mixed case
//! Dim text As String
//! text = LCase$("Hello World")  ' Returns "hello world"
//! ```
//!
//! ```vb
//! ' Example 3: With numbers and punctuation
//! Dim mixed As String
//! mixed = LCase$("ABC123!@#")  ' Returns "abc123!@#"
//! ```
//!
//! ```vb
//! ' Example 4: Already lowercase
//! Dim lower As String
//! lower = LCase$("already lowercase")  ' Returns "already lowercase"
//! ```
//!
//! ## Common Patterns
//!
//! ### Case-Insensitive Comparison
//! ```vb
//! Function CompareIgnoreCase(str1 As String, str2 As String) As Boolean
//!     CompareIgnoreCase = (LCase$(str1) = LCase$(str2))
//! End Function
//! ```
//!
//! ### Normalize User Input
//! ```vb
//! Function NormalizeInput(userInput As String) As String
//!     NormalizeInput = Trim$(LCase$(userInput))
//! End Function
//! ```
//!
//! ### Email Address Normalization
//! ```vb
//! Function NormalizeEmail(email As String) As String
//!     NormalizeEmail = LCase$(Trim$(email))
//! End Function
//! ```
//!
//! ### Case-Insensitive Search
//! ```vb
//! Function ContainsIgnoreCase(text As String, searchFor As String) As Boolean
//!     ContainsIgnoreCase = (InStr(LCase$(text), LCase$(searchFor)) > 0)
//! End Function
//! ```
//!
//! ### Username Validation
//! ```vb
//! Function ValidateUsername(username As String) As String
//!     ' Normalize to lowercase
//!     ValidateUsername = LCase$(Trim$(username))
//! End Function
//! ```
//!
//! ### File Extension Check
//! ```vb
//! Function HasExtension(filename As String, ext As String) As Boolean
//!     Dim fileExt As String
//!     fileExt = LCase$(Right$(filename, Len(ext)))
//!     HasExtension = (fileExt = LCase$(ext))
//! End Function
//! ```
//!
//! ### Dictionary Key Normalization
//! ```vb
//! Sub AddToDictionary(dict As Object, key As String, value As Variant)
//!     dict.Add LCase$(key), value
//! End Sub
//! ```
//!
//! ### Case-Insensitive Replace
//! ```vb
//! Function ReplaceIgnoreCase(text As String, findText As String, replaceWith As String) As String
//!     Dim lowerText As String
//!     Dim lowerFind As String
//!     Dim pos As Long
//!     
//!     lowerText = LCase$(text)
//!     lowerFind = LCase$(findText)
//!     pos = InStr(lowerText, lowerFind)
//!     
//!     If pos > 0 Then
//!         ReplaceIgnoreCase = Left$(text, pos - 1) & replaceWith & Mid$(text, pos + Len(findText))
//!     Else
//!         ReplaceIgnoreCase = text
//!     End If
//! End Function
//! ```
//!
//! ### Sort Key Generation
//! ```vb
//! Function GenerateSortKey(text As String) As String
//!     GenerateSortKey = LCase$(Trim$(text))
//! End Function
//! ```
//!
//! ### Command Parser
//! ```vb
//! Function ParseCommand(input As String) As String
//!     Dim cmd As String
//!     cmd = LCase$(Trim$(input))
//!     
//!     Select Case cmd
//!         Case "help"
//!             ParseCommand = "ShowHelp"
//!         Case "exit", "quit"
//!             ParseCommand = "Exit"
//!         Case Else
//!             ParseCommand = "Unknown"
//!     End Select
//! End Function
//! ```
//!
//! ## Advanced Examples
//!
//! ### Case-Insensitive Collection Lookup
//! ```vb
//! Function FindInCollection(col As Collection, key As String) As Variant
//!     On Error Resume Next
//!     Dim lowerKey As String
//!     Dim item As Variant
//!     Dim i As Long
//!     
//!     lowerKey = LCase$(key)
//!     
//!     ' Try direct lookup first
//!     FindInCollection = col(lowerKey)
//!     
//!     If Err.Number <> 0 Then
//!         ' Not found, search manually
//!         Err.Clear
//!         For i = 1 To col.Count
//!             If LCase$(col(i)) = lowerKey Then
//!                 FindInCollection = col(i)
//!                 Exit Function
//!             End If
//!         Next i
//!     End If
//! End Function
//! ```
//!
//! ### SQL Query Builder
//! ```vb
//! Function BuildWhereClause(fieldName As String, operator As String, value As String) As String
//!     Dim op As String
//!     op = LCase$(Trim$(operator))
//!     
//!     Select Case op
//!         Case "equals", "="
//!             BuildWhereClause = fieldName & " = '" & value & "'"
//!         Case "contains", "like"
//!             BuildWhereClause = fieldName & " LIKE '%" & value & "%'"
//!         Case "startswith"
//!             BuildWhereClause = fieldName & " LIKE '" & value & "%'"
//!         Case Else
//!             BuildWhereClause = ""
//!     End Select
//! End Function
//! ```
//!
//! ### Configuration File Parser
//! ```vb
//! Function ParseConfigLine(line As String, ByRef key As String, ByRef value As String) As Boolean
//!     Dim pos As Long
//!     Dim trimmedLine As String
//!     
//!     trimmedLine = Trim$(line)
//!     
//!     ' Skip comments and empty lines
//!     If Len(trimmedLine) = 0 Or Left$(trimmedLine, 1) = "#" Then
//!         ParseConfigLine = False
//!         Exit Function
//!     End If
//!     
//!     pos = InStr(trimmedLine, "=")
//!     If pos > 0 Then
//!         key = LCase$(Trim$(Left$(trimmedLine, pos - 1)))
//!         value = Trim$(Mid$(trimmedLine, pos + 1))
//!         ParseConfigLine = True
//!     Else
//!         ParseConfigLine = False
//!     End If
//! End Function
//! ```
//!
//! ### Smart String Comparison
//! ```vb
//! Function SmartCompare(str1 As String, str2 As String, caseSensitive As Boolean) As Boolean
//!     If caseSensitive Then
//!         SmartCompare = (str1 = str2)
//!     Else
//!         SmartCompare = (LCase$(str1) = LCase$(str2))
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeLCase(text As String) As String
//!     On Error GoTo ErrorHandler
//!     
//!     If IsNull(text) Then
//!         SafeLCase = ""
//!         Exit Function
//!     End If
//!     
//!     SafeLCase = LCase$(text)
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeLCase = ""
//! End Function
//! ```
//!
//! ## Performance Notes
//!
//! - `LCase$` is a fast operation with minimal overhead
//! - For large strings, the performance is linear with string length
//! - `LCase$` (returns `String`) is slightly faster than `LCase` (returns `Variant`)
//! - When comparing strings, convert both once rather than repeatedly
//! - Consider caching lowercase versions of frequently compared strings
//! - For very large datasets, consider using database-level case-insensitive comparisons
//!
//! ## Best Practices
//!
//! 1. **Use for comparisons** - Always normalize case when doing case-insensitive comparisons
//! 2. **Prefer `LCase$` over `LCase`** - Use `LCase$` when you know the result is a string
//! 3. **Cache results** - Store lowercase versions rather than converting repeatedly
//! 4. **Handle Null** - Check for `Null` values before calling `LCase$`
//! 5. **Combine with Trim** - Often useful to combine `LCase$` with `Trim$` for user input
//! 6. **Document intent** - Make it clear when lowercase conversion is for comparison vs. display
//! 7. **Consider locale** - Be aware that conversion may vary by system locale
//!
//! ## Comparison with Related Functions
//!
//! | Function | Return Type | Conversion | Use Case |
//! |----------|-------------|------------|----------|
//! | `LCase` | `Variant` | To lowercase | When working with `Variant` types |
//! | `LCase$` | `String` | To lowercase | When result is definitely a string |
//! | `UCase` | `Variant` | To uppercase | Convert to uppercase (`Variant`) |
//! | `UCase$` | `String` | To uppercase | Convert to uppercase (`String`) |
//! | `StrConv` | `String` | Various conversions | Complex case conversions |
//!
//! ## Common Use Cases
//!
//! ### Case-Insensitive String Arrays
//!
//! ```vb
//! Function ArrayContains(arr() As String, value As String) As Boolean
//!     Dim i As Long
//!     Dim lowerValue As String
//!     
//!     lowerValue = LCase$(value)
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         If LCase$(arr(i)) = lowerValue Then
//!             ArrayContains = True
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     ArrayContains = False
//! End Function
//! ```
//!
//! ### URL Parameter Parsing
//!
//! ```vb
//! Function GetURLParameter(url As String, paramName As String) As String
//!     Dim lowerURL As String
//!     Dim lowerParam As String
//!     Dim startPos As Long
//!     Dim endPos As Long
//!     
//!     lowerURL = LCase$(url)
//!     lowerParam = LCase$(paramName) & "="
//!     
//!     startPos = InStr(lowerURL, lowerParam)
//!     If startPos > 0 Then
//!         startPos = startPos + Len(lowerParam)
//!         endPos = InStr(startPos, url, "&")
//!         If endPos = 0 Then endPos = Len(url) + 1
//!         GetURLParameter = Mid$(url, startPos, endPos - startPos)
//!     End If
//! End Function
//! ```
//!
//! ## Platform Notes
//!
//! - On Windows, `LCase$` respects the system locale for character conversion
//! - Behavior may vary for extended ASCII and international characters
//! - For ASCII characters (A-Z), behavior is consistent across all platforms
//! - Some characters may convert differently depending on the active code page
//! - Modern Windows systems handle Unicode characters in `LCase$` operations
//!
//! ## Limitations
//!
//! - Conversion is based on system locale; may not work as expected for all Unicode characters
//! - Returns `Null` if the input is `Null` (unlike some other string functions that error)
//! - Does not handle advanced Unicode normalization or case folding
//! - For true Unicode case folding, more sophisticated methods may be needed
//! - Some special characters (like German ß) may not convert in all locales

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn lcase_dollar_simple() {
        let source = r#"
Sub Test()
    result = LCase$("HELLO")
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
                            Identifier ("LCase$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"HELLO\""),
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
    fn lcase_dollar_mixed_case() {
        let source = r#"
Sub Test()
    text = LCase$("Hello World")
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
                        CallExpression {
                            Identifier ("LCase$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Hello World\""),
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
    fn lcase_dollar_with_numbers() {
        let source = r#"
Sub Test()
    mixed = LCase$("ABC123")
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
                            Identifier ("mixed"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LCase$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"ABC123\""),
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
    fn lcase_dollar_compare_function() {
        let source = r"
Function CompareIgnoreCase(str1 As String, str2 As String) As Boolean
    CompareIgnoreCase = (LCase$(str1) = LCase$(str2))
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CompareIgnoreCase"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("str1"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("str2"),
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
                            Identifier ("CompareIgnoreCase"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("LCase$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("str1"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("LCase$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("str2"),
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
    fn lcase_dollar_normalize_input() {
        let source = r"
Function NormalizeInput(userInput As String) As String
    NormalizeInput = Trim$(LCase$(userInput))
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("NormalizeInput"),
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
                            Identifier ("NormalizeInput"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Trim$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("LCase$"),
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
    fn lcase_dollar_email_normalization() {
        let source = r"
Function NormalizeEmail(email As String) As String
    NormalizeEmail = LCase$(Trim$(email))
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("NormalizeEmail"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("email"),
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
                            Identifier ("NormalizeEmail"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LCase$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Trim$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("email"),
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
    fn lcase_dollar_search() {
        let source = r"
Function ContainsIgnoreCase(text As String, searchFor As String) As Boolean
    ContainsIgnoreCase = (InStr(LCase$(text), LCase$(searchFor)) > 0)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ContainsIgnoreCase"),
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
                Identifier ("searchFor"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                BooleanKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("ContainsIgnoreCase"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("InStr"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            CallExpression {
                                                Identifier ("LCase$"),
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
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            CallExpression {
                                                Identifier ("LCase$"),
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        IdentifierExpression {
                                                            Identifier ("searchFor"),
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
    fn lcase_dollar_username_validation() {
        let source = r"
Function ValidateUsername(username As String) As String
    ValidateUsername = LCase$(Trim$(username))
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ValidateUsername"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("username"),
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
                            Identifier ("ValidateUsername"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LCase$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Trim$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("username"),
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
    fn lcase_dollar_file_extension() {
        let source = r"
Function HasExtension(filename As String, ext As String) As Boolean
    Dim fileExt As String
    fileExt = LCase$(Right$(filename, Len(ext)))
    HasExtension = (fileExt = LCase$(ext))
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("HasExtension"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("filename"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("ext"),
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
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("fileExt"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("fileExt"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LCase$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Right$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("filename"),
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                CallExpression {
                                                    LenKeyword,
                                                    LeftParenthesis,
                                                    ArgumentList {
                                                        Argument {
                                                            IdentifierExpression {
                                                                Identifier ("ext"),
                                                            },
                                                        },
                                                    },
                                                    RightParenthesis,
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("HasExtension"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("fileExt"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("LCase$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("ext"),
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
    fn lcase_dollar_dictionary_key() {
        let source = r"
Sub AddToDictionary(dict As Object, key As String, value As Variant)
    lowercaseKey = LCase$(key)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("AddToDictionary"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("dict"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    ObjectKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("key"),
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
                    VariantKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("lowercaseKey"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LCase$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("key"),
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
    fn lcase_dollar_command_parser() {
        let source = r"
Function ParseCommand(input As String) As String
    Dim cmd As String
    cmd = LCase$(Trim$(input))
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ParseCommand"),
                ParameterList {
                    LeftParenthesis,
                },
                InputKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
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
                        Identifier ("cmd"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("cmd"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LCase$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Trim$"),
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
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn lcase_dollar_collection_lookup() {
        let source = r"
Function FindInCollection(col As Collection, key As String) As Variant
    Dim lowerKey As String
    lowerKey = LCase$(key)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("FindInCollection"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("col"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    Identifier ("Collection"),
                    Comma,
                    Whitespace,
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
                VariantKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("lowerKey"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("lowerKey"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LCase$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("key"),
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
    fn lcase_dollar_sql_builder() {
        let source = r"
Function BuildWhereClause(fieldName As String, operator As String) As String
    Dim op As String
    op = LCase$(Trim$(operator))
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("BuildWhereClause"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("fieldName"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("operator"),
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
                        Identifier ("op"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("op"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LCase$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Trim$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("operator"),
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
    fn lcase_dollar_config_parser() {
        let source = r"
Function ParseConfigLine(line As String, key As String) As Boolean
    key = LCase$(Trim$(Left$(line, 10)))
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ParseConfigLine"),
                ParameterList {
                    LeftParenthesis,
                },
                LineKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Comma,
                Whitespace,
                Identifier ("key"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                BooleanKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("key"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LCase$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Trim$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                CallExpression {
                                                    Identifier ("Left$"),
                                                    LeftParenthesis,
                                                    ArgumentList {
                                                        Argument {
                                                            IdentifierExpression {
                                                                LineKeyword,
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
    fn lcase_dollar_smart_compare() {
        let source = r"
Function SmartCompare(str1 As String, str2 As String, caseSensitive As Boolean) As Boolean
    If Not caseSensitive Then
        SmartCompare = (LCase$(str1) = LCase$(str2))
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
                Identifier ("SmartCompare"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("str1"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("str2"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("caseSensitive"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    BooleanKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                BooleanKeyword,
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        UnaryExpression {
                            NotKeyword,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("caseSensitive"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("SmartCompare"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                ParenthesizedExpression {
                                    LeftParenthesis,
                                    BinaryExpression {
                                        CallExpression {
                                            Identifier ("LCase$"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    IdentifierExpression {
                                                        Identifier ("str1"),
                                                    },
                                                },
                                            },
                                            RightParenthesis,
                                        },
                                        Whitespace,
                                        EqualityOperator,
                                        Whitespace,
                                        CallExpression {
                                            Identifier ("LCase$"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    IdentifierExpression {
                                                        Identifier ("str2"),
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
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn lcase_dollar_safe_wrapper() {
        let source = r"
Function SafeLCase(text As String) As String
    SafeLCase = LCase$(text)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("SafeLCase"),
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
                StringKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("SafeLCase"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LCase$"),
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
    fn lcase_dollar_array_contains() {
        let source = r"
Function ArrayContains(arr() As String, value As String) As Boolean
    Dim lowerValue As String
    lowerValue = LCase$(value)
    If LCase$(arr(0)) = lowerValue Then
        ArrayContains = True
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
                Identifier ("ArrayContains"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("arr"),
                    LeftParenthesis,
                    RightParenthesis,
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
                BooleanKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("lowerValue"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("lowerValue"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LCase$"),
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
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("LCase$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
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
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("lowerValue"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("ArrayContains"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BooleanLiteralExpression {
                                    TrueKeyword,
                                },
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
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn lcase_dollar_url_parameter() {
        let source = r"
Function GetURLParameter(url As String, paramName As String) As String
    Dim lowerURL As String
    lowerURL = LCase$(url)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetURLParameter"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("url"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("paramName"),
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
                        Identifier ("lowerURL"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("lowerURL"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LCase$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("url"),
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
    fn lcase_dollar_in_select() {
        let source = r#"
Sub Test()
    Select Case LCase$(command)
        Case "help"
            ShowHelp
        Case "exit"
            ExitApp
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
                            Identifier ("LCase$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("command"),
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
                            StringLiteral ("\"help\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("ShowHelp"),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\"exit\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("ExitApp"),
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
    fn lcase_dollar_in_loop() {
        let source = r"
Sub Test()
    For i = 0 To 10
        items(i) = LCase$(items(i))
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
                            IntegerLiteral ("0"),
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
                                    Identifier ("items"),
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
                                    Identifier ("LCase$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            CallExpression {
                                                Identifier ("items"),
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
}
