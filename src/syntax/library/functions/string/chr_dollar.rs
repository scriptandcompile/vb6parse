//! # `Chr$` Function
//!
//! Returns a `String` containing the character associated with the specified character code.
//! The dollar sign suffix (`$`) explicitly indicates that this function returns a `String` type.
//!
//! ## Syntax
//!
//! ```vb
//! Chr$(charcode)
//! ```
//!
//! ## Parameters
//!
//! - **`charcode`**: Required. `Long` value that identifies a character. The valid range for
//!   `charcode` is 0-255. For values outside this range, an error occurs.
//!
//! ## Return Value
//!
//! Returns a `String` containing the single character corresponding to the specified character
//! code. For `charcode` values 0-127, this corresponds to the ASCII character set. For values
//! 128-255, this corresponds to the extended ASCII or ANSI character set based on the system's
//! code page.
//!
//! ## Remarks
//!
//! - The `Chr$` function always returns a `String`, while `Chr` (without `$`) can return a `Variant`.
//! - Valid range: 0-255 (Error 5 "Invalid procedure call or argument" for values outside range).
//! - `Chr$(0)` returns a null character (`vbNullChar`).
//! - `Chr$(13)` returns carriage return (`vbCr`).
//! - `Chr$(10)` returns line feed (`vbLf`).
//! - `Chr$(9)` returns tab character (`vbTab`).
//! - Values 0-31 are non-printable control characters.
//! - Values 32-126 are standard printable ASCII characters.
//! - Values 127-255 depend on the system code page (often Windows-1252 in VB6).
//! - The inverse function is `Asc`, which returns the numeric character code of a character.
//! - For better performance when you know the result is a string, use `Chr$` instead of `Chr`.
//!
//! ## Common Character Codes
//!
//! | Code | Character | Constant | Description |
//! |------|-----------|----------|-------------|
//! | 0 | (null) | `vbNullChar` | Null character |
//! | 9 | \t | `vbTab` | Horizontal tab |
//! | 10 | \n | `vbLf` | Line feed |
//! | 13 | \r | `vbCr` | Carriage return |
//! | 32 | (space) | - | Space character |
//! | 34 | " | - | Double quote |
//! | 39 | ' | - | Single quote |
//! | 65 | A | - | Uppercase A |
//! | 97 | a | - | Lowercase a |
//!
//! ## Typical Uses
//!
//! 1. **Line breaks** - Insert carriage returns and line feeds in strings
//! 2. **Special characters** - Add tabs, quotes, and other special characters
//! 3. **Character generation** - Build strings from character codes
//! 4. **Alphabet generation** - Create sequences of characters programmatically
//! 5. **Tab-separated values** - Format data with tab delimiters
//! 6. **Quote escaping** - Insert quotes within strings
//! 7. **File formatting** - Create properly formatted text files
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Get character from code
//! Dim ch As String
//! ch = Chr$(65)  ' Returns "A"
//! ```
//!
//! ```vb
//! ' Example 2: Lowercase letter
//! Dim lower As String
//! lower = Chr$(97)  ' Returns "a"
//! ```
//!
//! ```vb
//! ' Example 3: Special character
//! Dim space As String
//! space = Chr$(32)  ' Returns " "
//! ```
//!
//! ```vb
//! ' Example 4: Line break
//! Dim msg As String
//! msg = "Line 1" & Chr$(13) & Chr$(10) & "Line 2"
//! ```
//!
//! ## Common Patterns
//!
//! ### Multi-line Strings
//! ```vb
//! Function CreateMultiLine() As String
//!     Dim result As String
//!     result = "First Line" & Chr$(13) & Chr$(10)
//!     result = result & "Second Line" & Chr$(13) & Chr$(10)
//!     result = result & "Third Line"
//!     CreateMultiLine = result
//! End Function
//! ```
//!
//! ### Tab-Separated Values
//! ```vb
//! Function CreateTSV(col1 As String, col2 As String, col3 As String) As String
//!     CreateTSV = col1 & Chr$(9) & col2 & Chr$(9) & col3
//! End Function
//! ```
//!
//! ### Generate Alphabet
//! ```vb
//! Function GenerateAlphabet() As String
//!     Dim i As Integer
//!     Dim result As String
//!     For i = 65 To 90
//!         result = result & Chr$(i)
//!     Next i
//!     GenerateAlphabet = result  ' Returns "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
//! End Function
//! ```
//!
//! ### Quote in String
//! ```vb
//! Function AddQuotes(text As String) As String
//!     AddQuotes = Chr$(34) & text & Chr$(34)
//! End Function
//! ```
//!
//! ### CSV Field with Quotes
//! ```vb
//! Function QuoteCSVField(field As String) As String
//!     ' Replace " with ""
//!     Dim quoted As String
//!     quoted = Replace(field, Chr$(34), Chr$(34) & Chr$(34))
//!     QuoteCSVField = Chr$(34) & quoted & Chr$(34)
//! End Function
//! ```
//!
//! ### Null-Terminated String
//! ```vb
//! Function CreateNullTerminated(text As String) As String
//!     CreateNullTerminated = text & Chr$(0)
//! End Function
//! ```
//!
//! ### Password Mask
//! ```vb
//! Function MaskPassword(length As Integer) As String
//!     Dim i As Integer
//!     Dim result As String
//!     For i = 1 To length
//!         result = result & Chr$(42)  ' Asterisk
//!     Next i
//!     MaskPassword = result
//! End Function
//! ```
//!
//! ### Character Range Check
//! ```vb
//! Function IsUpperCase(ch As String) As Boolean
//!     If Len(ch) <> 1 Then Exit Function
//!     Dim code As Integer
//!     code = Asc(ch)
//!     IsUpperCase = (code >= 65 And code <= 90)
//! End Function
//! ```
//!
//! ### Build Character Set
//! ```vb
//! Function GetDigitCharacters() As String
//!     Dim i As Integer
//!     Dim result As String
//!     For i = 48 To 57  ' ASCII codes for 0-9
//!         result = result & Chr$(i)
//!     Next i
//!     GetDigitCharacters = result  ' Returns "0123456789"
//! End Function
//! ```
//!
//! ### Format Output with Alignment
//! ```vb
//! Function AlignRight(text As String, width As Integer) As String
//!     Dim padding As Integer
//!     Dim result As String
//!     padding = width - Len(text)
//!     If padding > 0 Then
//!         Dim i As Integer
//!         For i = 1 To padding
//!             result = result & Chr$(32)  ' Space
//!         Next i
//!     End If
//!     AlignRight = result & text
//! End Function
//! ```
//!
//! ## Advanced Examples
//!
//! ### Format Report Header
//! ```vb
//! Function CreateReportHeader() As String
//!     Dim header As String
//!     header = "Name" & Chr$(9) & "Age" & Chr$(9) & "City" & Chr$(13) & Chr$(10)
//!     header = header & String$(40, Chr$(45))  ' Underline with dashes
//!     CreateReportHeader = header
//! End Function
//! ```
//!
//! ### Parse Character Codes
//! ```vb
//! Function DecodeCharCodes(codes() As Integer) As String
//!     Dim i As Integer
//!     Dim result As String
//!     For i = LBound(codes) To UBound(codes)
//!         result = result & Chr$(codes(i))
//!     Next i
//!     DecodeCharCodes = result
//! End Function
//! ```
//!
//! ### Create Box Drawing
//! ```vb
//! Function CreateBox(width As Integer, height As Integer) As String
//!     Dim result As String
//!     Dim i As Integer
//!     
//!     ' Top line
//!     result = String$(width, Chr$(45)) & Chr$(13) & Chr$(10)
//!     
//!     ' Middle lines
//!     For i = 1 To height - 2
//!         result = result & Chr$(124) & Space$(width - 2) & Chr$(124) & Chr$(13) & Chr$(10)
//!     Next i
//!     
//!     ' Bottom line
//!     result = result & String$(width, Chr$(45))
//!     
//!     CreateBox = result
//! End Function
//! ```
//!
//! ## Differences from Chr
//!
//! | Feature | `Chr$` | `Chr` |
//! |---------|--------|-------|
//! | Return Type | Always `String` | Can return `Variant` |
//! | Performance | Slightly faster | Slightly slower |
//! | Type Safety | Compile-time type checking | Runtime type checking |
//! | Assignment | Can only assign to `String` | Can assign to `Variant` or `String` |
//!
//! ## Related Functions
//!
//! - `Chr`: Returns character as `Variant` instead of `String`
//! - `ChrB$`: Returns byte character for double-byte character sets
//! - `ChrW$`: Returns Unicode character
//! - `Asc`: Returns character code for a character (inverse of `Chr$`)
//! - `AscB`: Returns byte value of first byte in string
//! - `AscW`: Returns Unicode character code
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeChr(code As Long) As String
//!     On Error Resume Next
//!     SafeChr = Chr$(code)
//!     If Err.Number <> 0 Then
//!         SafeChr = ""
//!         Err.Clear
//!     End If
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - `Chr$` is slightly more efficient than `Chr` because it avoids `Variant` overhead
//! - For building strings from many characters, consider using a buffer or `String$` function
//! - Concatenating many `Chr$` calls can be slow; use arrays and `Join` for better performance
//!
//! ## Limitations
//!
//! - Limited to character codes 0-255 (use `ChrW$` for full Unicode support)
//! - Character interpretation depends on system code page
//! - Does not validate that the resulting character is printable
//! - No direct support for multi-byte characters (use `ChrB$` for DBCS)

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn chr_dollar_simple() {
        let source = r"
Sub Test()
    ch = Chr$(65)
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
                            Identifier ("ch"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Chr$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("65"),
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
    fn chr_dollar_lowercase() {
        let source = r"
Sub Test()
    lower = Chr$(97)
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
                            Identifier ("lower"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Chr$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("97"),
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
    fn chr_dollar_space() {
        let source = r"
Sub Test()
    space = Chr$(32)
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
                            Identifier ("space"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Chr$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("32"),
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
    fn chr_dollar_line_break() {
        let source = r#"
Sub Test()
    msg = "Line 1" & Chr$(13) & Chr$(10) & "Line 2"
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
                        BinaryExpression {
                            BinaryExpression {
                                BinaryExpression {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Line 1\""),
                                    },
                                    Whitespace,
                                    Ampersand,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Chr$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("13"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Chr$"),
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
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"Line 2\""),
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
    fn chr_dollar_multi_line_function() {
        let source = r#"
Function CreateMultiLine() As String
    Dim result As String
    result = "First Line" & Chr$(13) & Chr$(10)
    result = result & "Second Line"
    CreateMultiLine = result
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CreateMultiLine"),
                ParameterList {
                    LeftParenthesis,
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
                        BinaryExpression {
                            BinaryExpression {
                                StringLiteralExpression {
                                    StringLiteral ("\"First Line\""),
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Chr$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("13"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Chr$"),
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
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("result"),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"Second Line\""),
                            },
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("CreateMultiLine"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("result"),
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
    fn chr_dollar_tsv() {
        let source = r"
Function CreateTSV(col1 As String, col2 As String) As String
    CreateTSV = col1 & Chr$(9) & col2
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CreateTSV"),
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
                            Identifier ("CreateTSV"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("col1"),
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Chr$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("9"),
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
                                Identifier ("col2"),
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
    fn chr_dollar_alphabet() {
        let source = r"
Function GenerateAlphabet() As String
    Dim i As Integer
    Dim result As String
    For i = 65 To 90
        result = result & Chr$(i)
    Next i
    GenerateAlphabet = result
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GenerateAlphabet"),
                ParameterList {
                    LeftParenthesis,
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
                        Identifier ("result"),
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
                            IntegerLiteral ("65"),
                        },
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("90"),
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
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    Ampersand,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Chr$"),
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GenerateAlphabet"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("result"),
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
    fn chr_dollar_quotes() {
        let source = r"
Function AddQuotes(text As String) As String
    AddQuotes = Chr$(34) & text & Chr$(34)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("AddQuotes"),
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
                            Identifier ("AddQuotes"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("Chr$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("34"),
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
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Chr$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("34"),
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
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn chr_dollar_csv_field() {
        let source = r"
Function QuoteCSVField(field As String) As String
    Dim quoted As String
    quoted = Replace(field, Chr$(34), Chr$(34) & Chr$(34))
    QuoteCSVField = Chr$(34) & quoted & Chr$(34)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("QuoteCSVField"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("field"),
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
                        Identifier ("quoted"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("quoted"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Replace"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("field"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    CallExpression {
                                        Identifier ("Chr$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("34"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    BinaryExpression {
                                        CallExpression {
                                            Identifier ("Chr$"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    NumericLiteralExpression {
                                                        IntegerLiteral ("34"),
                                                    },
                                                },
                                            },
                                            RightParenthesis,
                                        },
                                        Whitespace,
                                        Ampersand,
                                        Whitespace,
                                        CallExpression {
                                            Identifier ("Chr$"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    NumericLiteralExpression {
                                                        IntegerLiteral ("34"),
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("QuoteCSVField"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("Chr$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("34"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("quoted"),
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Chr$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("34"),
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
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn chr_dollar_null_terminated() {
        let source = r"
Function CreateNullTerminated(text As String) As String
    CreateNullTerminated = text & Chr$(0)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CreateNullTerminated"),
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
                            Identifier ("CreateNullTerminated"),
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
                                Identifier ("Chr$"),
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
    fn chr_dollar_password_mask() {
        let source = r"
Function MaskPassword(length As Integer) As String
    Dim i As Integer
    Dim result As String
    For i = 1 To length
        result = result & Chr$(42)
    Next i
    MaskPassword = result
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("MaskPassword"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("length"),
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
                        Identifier ("result"),
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
                        IdentifierExpression {
                            Identifier ("length"),
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
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    Ampersand,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Chr$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("42"),
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
                        NextKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("MaskPassword"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("result"),
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
    fn chr_dollar_digit_characters() {
        let source = r"
Function GetDigitCharacters() As String
    Dim i As Integer
    Dim result As String
    For i = 48 To 57
        result = result & Chr$(i)
    Next i
    GetDigitCharacters = result
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetDigitCharacters"),
                ParameterList {
                    LeftParenthesis,
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
                        Identifier ("result"),
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
                            IntegerLiteral ("48"),
                        },
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("57"),
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
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    Ampersand,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Chr$"),
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetDigitCharacters"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("result"),
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
    fn chr_dollar_align_right() {
        let source = r"
Function AlignRight(text As String, width As Integer) As String
    Dim padding As Integer
    Dim result As String
    padding = width - Len(text)
    If padding > 0 Then
        Dim i As Integer
        For i = 1 To padding
            result = result & Chr$(32)
        Next i
    End If
    AlignRight = result & text
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("AlignRight"),
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
                            Identifier ("padding"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
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
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("padding"),
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
                                IdentifierExpression {
                                    Identifier ("padding"),
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
                                            IdentifierExpression {
                                                Identifier ("result"),
                                            },
                                            Whitespace,
                                            Ampersand,
                                            Whitespace,
                                            CallExpression {
                                                Identifier ("Chr$"),
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        NumericLiteralExpression {
                                                            IntegerLiteral ("32"),
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("AlignRight"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("result"),
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
    fn chr_dollar_report_header() {
        let source = r#"
Function CreateReportHeader() As String
    Dim header As String
    header = "Name" & Chr$(9) & "Age" & Chr$(13) & Chr$(10)
    CreateReportHeader = header
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CreateReportHeader"),
                ParameterList {
                    LeftParenthesis,
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
                        Identifier ("header"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("header"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                BinaryExpression {
                                    BinaryExpression {
                                        StringLiteralExpression {
                                            StringLiteral ("\"Name\""),
                                        },
                                        Whitespace,
                                        Ampersand,
                                        Whitespace,
                                        CallExpression {
                                            Identifier ("Chr$"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    NumericLiteralExpression {
                                                        IntegerLiteral ("9"),
                                                    },
                                                },
                                            },
                                            RightParenthesis,
                                        },
                                    },
                                    Whitespace,
                                    Ampersand,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"Age\""),
                                    },
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Chr$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("13"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Chr$"),
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
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("CreateReportHeader"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("header"),
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
    fn chr_dollar_decode_codes() {
        let source = r"
Function DecodeCharCodes(codes() As Integer) As String
    Dim i As Integer
    Dim result As String
    For i = LBound(codes) To UBound(codes)
        result = result & Chr$(codes(i))
    Next i
    DecodeCharCodes = result
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("DecodeCharCodes"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("codes"),
                    LeftParenthesis,
                    RightParenthesis,
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
                        Identifier ("result"),
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
                        CallExpression {
                            Identifier ("LBound"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("codes"),
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
                                        Identifier ("codes"),
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
                                    Identifier ("result"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    Ampersand,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Chr$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                CallExpression {
                                                    Identifier ("codes"),
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("DecodeCharCodes"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("result"),
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
    fn chr_dollar_box_drawing() {
        let source = r"
Function CreateBox(width As Integer) As String
    Dim result As String
    result = String$(width, Chr$(45)) & Chr$(13) & Chr$(10)
    result = result & Chr$(124) & Space$(width - 2) & Chr$(124)
    CreateBox = result
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CreateBox"),
                ParameterList {
                    LeftParenthesis,
                },
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
                        IdentifierExpression {
                            StringKeyword,
                        },
                    },
                    Unknown,
                    Unknown,
                    WidthStatement {
                        WidthKeyword,
                        Comma,
                        Whitespace,
                        Identifier ("Chr$"),
                        LeftParenthesis,
                        IntegerLiteral ("45"),
                        RightParenthesis,
                        RightParenthesis,
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Chr$"),
                        LeftParenthesis,
                        IntegerLiteral ("13"),
                        RightParenthesis,
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Chr$"),
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        RightParenthesis,
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
                            BinaryExpression {
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("result"),
                                    },
                                    Whitespace,
                                    Ampersand,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Chr$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("124"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
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
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("2"),
                                                },
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Chr$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("124"),
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
                            Identifier ("CreateBox"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("result"),
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
    fn chr_dollar_safe_chr() {
        let source = r#"
Function SafeChr(code As Long) As String
    On Error Resume Next
    SafeChr = Chr$(code)
    If Err.Number <> 0 Then
        SafeChr = ""
        Err.Clear
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
                Identifier ("SafeChr"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("code"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
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
                            Identifier ("SafeChr"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Chr$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("code"),
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
                            MemberAccessExpression {
                                Identifier ("Err"),
                                PeriodOperator,
                                Identifier ("Number"),
                            },
                            Whitespace,
                            InequalityOperator,
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
                                    Identifier ("SafeChr"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                StringLiteralExpression {
                                    StringLiteral ("\"\""),
                                },
                                Newline,
                            },
                            Whitespace,
                            CallStatement {
                                Identifier ("Err"),
                                PeriodOperator,
                                Identifier ("Clear"),
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
    fn chr_dollar_tab_character() {
        let source = r#"
Sub Test()
    data = "Name" & Chr$(9) & "Age"
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
                            Identifier ("data"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                StringLiteralExpression {
                                    StringLiteral ("\"Name\""),
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Chr$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("9"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"Age\""),
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
    fn chr_dollar_multiple_calls() {
        let source = r"
Sub Test()
    text = Chr$(65) & Chr$(66) & Chr$(67)
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
                            TextKeyword,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("Chr$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("65"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Chr$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("66"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Chr$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("67"),
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
    fn chr_dollar_in_condition() {
        let source = r#"
Sub Test()
    If ch = Chr$(32) Then
        MsgBox "Space"
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
                            IdentifierExpression {
                                Identifier ("ch"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("Chr$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("32"),
                                        },
                                    },
                                },
                                RightParenthesis,
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
                                StringLiteral ("\"Space\""),
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
    fn chr_dollar_nested_functions() {
        let source = r"
Sub Test()
    result = UCase$(Chr$(97))
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
                            Identifier ("UCase$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Chr$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("97"),
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
}
