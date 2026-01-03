//! # `Right$` Function
//!
//! The `Right$` function in Visual Basic 6 returns a string containing a specified number of
//! characters from the right side (end) of a string. The dollar sign (`$`) suffix indicates
//! that this function always returns a `String` type, never a `Variant`.
//!
//! ## Syntax
//!
//! ```vb6
//! Right$(string, length)
//! ```
//!
//! ## Parameters
//!
//! - `string` - Required. String expression from which the rightmost characters are returned.
//!   If `string` contains `Null`, `Null` is returned.
//! - `length` - Required. Numeric expression indicating how many characters to return. If 0,
//!   a zero-length string ("") is returned. If greater than or equal to the number of characters
//!   in `string`, the entire string is returned.
//!
//! ## Return Value
//!
//! Returns a `String` containing the rightmost `length` characters of `string`.
//!
//! ## Behavior and Characteristics
//!
//! ### Length Handling
//!
//! - If `length` = 0: Returns an empty string ("")
//! - If `length` >= `Len(string)`: Returns the entire string
//! - If `length` < 0: Generates a runtime error (Invalid procedure call or argument)
//! - If `string` is empty (""): Returns an empty string regardless of `length`
//!
//! ### Null Handling
//!
//! - If `string` contains `Null`: Returns `Null`
//! - If `length` is `Null`: Generates a runtime error (Invalid use of Null)
//!
//! ### Type Differences: `Right$` vs `Right`
//!
//! - `Right$`: Always returns `String` type (never `Variant`)
//! - `Right`: Returns `Variant` (can propagate `Null` values)
//! - Use `Right$` when you need guaranteed `String` return type
//! - Use `Right` when working with potentially `Null` values
//!
//! ## Common Usage Patterns
//!
//! ### 1. Extract File Extension
//!
//! ```vb6
//! Function GetExtension(fileName As String) As String
//!     Dim dotPos As Integer
//!     dotPos = InStrRev(fileName, ".")
//!     If dotPos > 0 Then
//!         GetExtension = Right$(fileName, Len(fileName) - dotPos)
//!     Else
//!         GetExtension = ""
//!     End If
//! End Function
//!
//! Dim ext As String
//! ext = GetExtension("document.txt")  ' Returns "txt"
//! ```
//!
//! ### 2. Get Last N Characters
//!
//! ```vb6
//! Dim text As String
//! Dim suffix As String
//! text = "Hello World"
//! suffix = Right$(text, 5)  ' Returns "World"
//! ```
//!
//! ### 3. Extract Account Number Suffix
//!
//! ```vb6
//! Function GetAccountSuffix(accountNum As String) As String
//!     ' Get last 4 digits of account number
//!     GetAccountSuffix = Right$(accountNum, 4)
//! End Function
//!
//! Dim lastFour As String
//! lastFour = GetAccountSuffix("1234567890")  ' Returns "7890"
//! ```
//!
//! ### 4. Pad String to Fixed Width
//!
//! ```vb6
//! Function PadLeft(text As String, width As Integer) As String
//!     Dim padded As String
//!     padded = Space(width) & text
//!     PadLeft = Right$(padded, width)
//! End Function
//!
//! Dim result As String
//! result = PadLeft("42", 5)  ' Returns "   42"
//! ```
//!
//! ### 5. Extract Trailing Digits
//!
//! ```vb6
//! Function GetTrailingNumber(text As String) As String
//!     Dim i As Integer
//!     Dim numChars As Integer
//!     For i = Len(text) To 1 Step -1
//!         If Not IsNumeric(Mid$(text, i, 1)) Then Exit For
//!         numChars = numChars + 1
//!     Next i
//!     If numChars > 0 Then
//!         GetTrailingNumber = Right$(text, numChars)
//!     Else
//!         GetTrailingNumber = ""
//!     End If
//! End Function
//!
//! Dim num As String
//! num = GetTrailingNumber("Item123")  ' Returns "123"
//! ```
//!
//! ### 6. Time Component Extraction
//!
//! ```vb6
//! Function GetSeconds(timeStr As String) As String
//!     ' Extract seconds from "HH:MM:SS" format
//!     GetSeconds = Right$(timeStr, 2)
//! End Function
//!
//! Dim secs As String
//! secs = GetSeconds("14:30:45")  ' Returns "45"
//! ```
//!
//! ### 7. Validate String Suffix
//!
//! ```vb6
//! Function HasExtension(fileName As String, ext As String) As Boolean
//!     Dim fileExt As String
//!     fileExt = Right$(fileName, Len(ext))
//!     HasExtension = (UCase$(fileExt) = UCase$(ext))
//! End Function
//!
//! If HasExtension("report.pdf", ".pdf") Then
//!     Debug.Print "PDF file detected"
//! End If
//! ```
//!
//! ### 8. Extract Domain from Email
//!
//! ```vb6
//! Function GetEmailDomain(email As String) As String
//!     Dim atPos As Integer
//!     atPos = InStr(email, "@")
//!     If atPos > 0 Then
//!         GetEmailDomain = Right$(email, Len(email) - atPos)
//!     Else
//!         GetEmailDomain = ""
//!     End If
//! End Function
//!
//! Dim domain As String
//! domain = GetEmailDomain("user@example.com")  ' Returns "example.com"
//! ```
//!
//! ### 9. Format Currency Display
//!
//! ```vb6
//! Function FormatAmount(amount As String) As String
//!     ' Align decimal values
//!     Dim formatted As String
//!     formatted = Space(15) & amount
//!     FormatAmount = Right$(formatted, 15)
//! End Function
//! ```
//!
//! ### 10. Extract Path Component
//!
//! ```vb6
//! Function GetFileName(fullPath As String) As String
//!     Dim slashPos As Integer
//!     slashPos = InStrRev(fullPath, "\")
//!     If slashPos > 0 Then
//!         GetFileName = Right$(fullPath, Len(fullPath) - slashPos)
//!     Else
//!         GetFileName = fullPath
//!     End If
//! End Function
//!
//! Dim fileName As String
//! fileName = GetFileName("C:\Documents\report.txt")  ' Returns "report.txt"
//! ```
//!
//! ## Related Functions
//!
//! - `Right()` - Returns a `Variant` containing the rightmost characters (can handle `Null`)
//! - `Left$()` - Returns a specified number of characters from the left side of a string
//! - `Mid$()` - Returns a specified number of characters from any position in a string
//! - `Len()` - Returns the number of characters in a string
//! - `InStrRev()` - Finds the position of a substring searching from the end
//! - `Trim$()` - Removes leading and trailing spaces from a string
//! - `LTrim$()` - Removes leading spaces from a string
//! - `RTrim$()` - Removes trailing spaces from a string
//!
//! ## Best Practices
//!
//! ### When to Use `Right$` vs `Right`
//!
//! ```vb6
//! ' Use Right$ when you need a String
//! Dim fileName As String
//! fileName = Right$(fullPath, 10)  ' Type-safe, always returns String
//!
//! ' Use Right when working with Variants or Null values
//! Dim result As Variant
//! result = Right(variantValue, 5)  ' Can propagate Null
//! ```
//!
//! ### Validate Length Parameter
//!
//! ```vb6
//! Function SafeRight(text As String, length As Integer) As String
//!     If length < 0 Then
//!         SafeRight = ""
//!     ElseIf length >= Len(text) Then
//!         SafeRight = text
//!     Else
//!         SafeRight = Right$(text, length)
//!     End If
//! End Function
//! ```
//!
//! ### Check for Empty Strings
//!
//! ```vb6
//! If Len(text) > 0 Then
//!     suffix = Right$(text, 3)
//! Else
//!     suffix = ""
//! End If
//! ```
//!
//! ### Use with `InStrRev` for Parsing
//!
//! ```vb6
//! ' Find last occurrence and extract everything after it
//! Dim pos As Integer
//! pos = InStrRev(fullPath, "\")
//! If pos > 0 Then
//!     fileName = Right$(fullPath, Len(fullPath) - pos)
//! End If
//! ```
//!
//! ## Performance Considerations
//!
//! - `Right$` is very efficient for small to moderate length strings
//! - For very large strings, consider if you really need to extract characters
//! - Using `Right$` in tight loops with large strings may impact performance
//! - Consider caching the length if calling `Len()` repeatedly
//!
//! ```vb6
//! ' Less efficient
//! For i = 1 To 1000
//!     result = Right$(largeString, Len(largeString) - 10)
//! Next i
//!
//! ' More efficient
//! Dim strLen As Long
//! strLen = Len(largeString)
//! For i = 1 To 1000
//!     result = Right$(largeString, strLen - 10)
//! Next i
//! ```
//!
//! ## Common Pitfalls
//!
//! ### 1. Negative Length Values
//!
//! ```vb6
//! ' Runtime error: Invalid procedure call or argument
//! text = Right$("Hello", -1)  ' ERROR!
//!
//! ' Validate first
//! If length >= 0 Then
//!     text = Right$(source, length)
//! End If
//! ```
//!
//! ### 2. Off-by-One Errors
//!
//! ```vb6
//! ' Common mistake: forgetting to account for delimiter position
//! Dim pos As Integer
//! pos = InStrRev(path, "\")
//! ' Wrong: includes the backslash
//! fileName = Right$(path, pos)
//! ' Correct: excludes the backslash
//! fileName = Right$(path, Len(path) - pos)
//! ```
//!
//! ### 3. Not Checking String Length
//!
//! ```vb6
//! ' Potential issue: what if text is shorter than 10 characters?
//! suffix = Right$(text, 10)  ' Returns entire string if text.Length < 10
//!
//! ' Better: check first
//! If Len(text) >= 10 Then
//!     suffix = Right$(text, 10)
//! Else
//!     ' Handle short string case
//!     suffix = text
//! End If
//! ```
//!
//! ### 4. Assuming Fixed Positions
//!
//! ```vb6
//! ' Fragile: assumes extension is always 3 characters
//! ext = Right$(fileName, 3)  ' Fails for ".html", ".jpeg"
//!
//! ' Better: find the dot
//! Dim dotPos As Integer
//! dotPos = InStrRev(fileName, ".")
//! If dotPos > 0 Then
//!     ext = Right$(fileName, Len(fileName) - dotPos)
//! End If
//! ```
//!
//! ### 5. Null Value Handling
//!
//! ```vb6
//! ' Right$ with Null causes runtime error
//! Dim result As String
//! result = Right$(nullValue, 5)  ' ERROR if nullValue is Null
//!
//! ' Protect against Null
//! If Not IsNull(value) Then
//!     result = Right$(value, 5)
//! Else
//!     result = ""
//! End If
//! ```
//!
//! ## Limitations
//!
//! - Cannot handle `Null` values (use `Right` variant function instead)
//! - No built-in trimming of whitespace (combine with `RTrim$` if needed)
//! - Negative `length` values cause runtime errors
//! - Works with characters, not bytes (use `RightB$` for byte-level operations)
//! - No Unicode-specific version (VB6 uses UCS-2 internally)
//! - Cannot extract from right based on delimiter (must calculate length manually)

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn right_dollar_simple() {
        let source = r#"
Sub Main()
    result = Right$("Hello", 3)
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
                            Identifier ("Right$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Hello\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("3"),
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
    fn right_dollar_assignment() {
        let source = r#"
Sub Main()
    Dim suffix As String
    suffix = Right$("Hello World", 5)
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
                        Identifier ("suffix"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("suffix"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Right$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Hello World\""),
                                    },
                                },
                                Comma,
                                Whitespace,
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
    fn right_dollar_variable() {
        let source = r#"
Sub Main()
    Dim text As String
    Dim result As String
    text = "Sample"
    result = Right$(text, 3)
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
                            StringLiteral ("\"Sample\""),
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
                            Identifier ("Right$"),
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
                                    NumericLiteralExpression {
                                        IntegerLiteral ("3"),
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
    fn right_dollar_file_extension() {
        let source = r#"
Function GetExtension(fileName As String) As String
    Dim dotPos As Integer
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        GetExtension = Right$(fileName, Len(fileName) - dotPos)
    End If
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetExtension"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("fileName"),
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
                        Identifier ("dotPos"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("dotPos"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("InStrRev"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("fileName"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\".\""),
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
                            IdentifierExpression {
                                Identifier ("dotPos"),
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
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("GetExtension"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Right$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("fileName"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            BinaryExpression {
                                                CallExpression {
                                                    LenKeyword,
                                                    LeftParenthesis,
                                                    ArgumentList {
                                                        Argument {
                                                            IdentifierExpression {
                                                                Identifier ("fileName"),
                                                            },
                                                        },
                                                    },
                                                    RightParenthesis,
                                                },
                                                Whitespace,
                                                SubtractionOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("dotPos"),
                                                },
                                            },
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
    fn right_dollar_account_suffix() {
        let source = r"
Function GetAccountSuffix(accountNum As String) As String
    GetAccountSuffix = Right$(accountNum, 4)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetAccountSuffix"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("accountNum"),
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
                            Identifier ("GetAccountSuffix"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Right$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("accountNum"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("4"),
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
    fn right_dollar_in_condition() {
        let source = r#"
Sub Main()
    If Right$(fileName, 4) = ".txt" Then
        Debug.Print "Text file"
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
                                Identifier ("Right$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("fileName"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("4"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\".txt\""),
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
                                StringLiteral ("\"Text file\""),
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
    fn right_dollar_pad_left() {
        let source = r"
Function PadLeft(text As String, width As Integer) As String
    Dim padded As String
    padded = Space(width) & text
    PadLeft = Right$(padded, width)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("PadLeft"),
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
                        Identifier ("padded"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("padded"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Space"),
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("PadLeft"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Right$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("padded"),
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

    #[test]
    fn right_dollar_time_extraction() {
        let source = r"
Function GetSeconds(timeStr As String) As String
    GetSeconds = Right$(timeStr, 2)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetSeconds"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("timeStr"),
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
                            Identifier ("GetSeconds"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Right$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("timeStr"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("2"),
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
    fn right_dollar_email_domain() {
        let source = r#"
Function GetEmailDomain(email As String) As String
    Dim atPos As Integer
    atPos = InStr(email, "@")
    If atPos > 0 Then
        GetEmailDomain = Right$(email, Len(email) - atPos)
    End If
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetEmailDomain"),
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
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("atPos"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("atPos"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("InStr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("email"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"@\""),
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
                            IdentifierExpression {
                                Identifier ("atPos"),
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
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("GetEmailDomain"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Right$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("email"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            BinaryExpression {
                                                CallExpression {
                                                    LenKeyword,
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
                                                Whitespace,
                                                SubtractionOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("atPos"),
                                                },
                                            },
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
    fn right_dollar_multiple_uses() {
        let source = r"
Sub ProcessText()
    Dim ext As String
    Dim suffix As String
    ext = Right$(fileName, 3)
    suffix = Right$(accountNum, 4)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ProcessText"),
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
                        Identifier ("ext"),
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
                        Identifier ("suffix"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("ext"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Right$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("fileName"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("3"),
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
                            Identifier ("suffix"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Right$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("accountNum"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("4"),
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
    fn right_dollar_select_case() {
        let source = r#"
Sub Main()
    Select Case Right$(fileName, 4)
        Case ".txt"
            Debug.Print "Text"
        Case ".doc"
            Debug.Print "Document"
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
                            Identifier ("Right$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("fileName"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("4"),
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
                            StringLiteral ("\".txt\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Text\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\".doc\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Document\""),
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
    fn right_dollar_expression_args() {
        let source = r"
Sub Main()
    Dim result As String
    result = Right$(text, Len(text) - 5)
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
                            Identifier ("Right$"),
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
                                        SubtractionOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("5"),
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
    fn right_dollar_concatenation() {
        let source = r#"
Sub Main()
    Dim output As String
    output = "Suffix: " & Right$(text, 5)
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
                                StringLiteral ("\"Suffix: \""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Right$"),
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
                                        NumericLiteralExpression {
                                            IntegerLiteral ("5"),
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
    fn right_dollar_get_filename() {
        let source = r#"
Function GetFileName(fullPath As String) As String
    Dim slashPos As Integer
    slashPos = InStrRev(fullPath, "\")
    If slashPos > 0 Then
        GetFileName = Right$(fullPath, Len(fullPath) - slashPos)
    End If
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetFileName"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("fullPath"),
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
                        Identifier ("slashPos"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("slashPos"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("InStrRev"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("fullPath"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"\\\")"),
                                    },
                                },
                            },
                        },
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("slashPos"),
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
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("GetFileName"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Right$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("fullPath"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            BinaryExpression {
                                                CallExpression {
                                                    LenKeyword,
                                                    LeftParenthesis,
                                                    ArgumentList {
                                                        Argument {
                                                            IdentifierExpression {
                                                                Identifier ("fullPath"),
                                                            },
                                                        },
                                                    },
                                                    RightParenthesis,
                                                },
                                                Whitespace,
                                                SubtractionOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("slashPos"),
                                                },
                                            },
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
    fn right_dollar_validation() {
        let source = r"
Function HasExtension(fileName As String, ext As String) As Boolean
    Dim fileExt As String
    fileExt = Right$(fileName, Len(ext))
    HasExtension = (UCase$(fileExt) = UCase$(ext))
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
                    Identifier ("fileName"),
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
                            Identifier ("Right$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("fileName"),
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
                                CallExpression {
                                    Identifier ("UCase$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("fileExt"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("UCase$"),
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
    fn right_dollar_zero_length() {
        let source = r#"
Sub Main()
    Dim empty As String
    empty = Right$("Hello", 0)
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
                            Identifier ("Right$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Hello\""),
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
    fn right_dollar_full_string() {
        let source = r#"
Sub Main()
    Dim full As String
    full = Right$("Hello", 100)
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
                        Identifier ("full"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("full"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Right$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Hello\""),
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
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn right_dollar_format_amount() {
        let source = r"
Function FormatAmount(amount As String) As String
    Dim formatted As String
    formatted = Space(15) & amount
    FormatAmount = Right$(formatted, 15)
End Function
";
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
                        Identifier ("formatted"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("formatted"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Space"),
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
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("amount"),
                            },
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
                        CallExpression {
                            Identifier ("Right$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("formatted"),
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
    fn right_dollar_with_trim() {
        let source = r"
Sub Main()
    Dim cleaned As String
    cleaned = RTrim$(Right$(dataField, 10))
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
                                    CallExpression {
                                        Identifier ("Right$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("dataField"),
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
    fn right_dollar_loop_processing() {
        let source = r"
Sub ProcessLines()
    Dim i As Integer
    Dim suffix As String
    For i = 1 To 10
        suffix = Right$(lines(i), 5)
        Debug.Print suffix
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
                        Identifier ("suffix"),
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
                                    Identifier ("suffix"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Right$"),
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
                                        Comma,
                                        Whitespace,
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
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                Identifier ("suffix"),
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
