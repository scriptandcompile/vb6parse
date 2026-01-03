//! # `Chr` Function
//!
//! Returns a `String` containing the character associated with the specified character code.
//!
//! ## Syntax
//!
//! ```vb
//! Chr(charcode)
//! ```
//!
//! ## Parameters
//!
//! - **`charcode`**: Required. Long value that identifies a character. The valid range for
//!   `charcode` is 0-255. For values outside this range, an error occurs.
//!
//! ## Return Value
//!
//! Returns a `String` containing the single character corresponding to the specified character
//! code. For charcode values 0-127, this corresponds to the ASCII character set. For values
//! 128-255, this corresponds to the extended ASCII or ANSI character set based on the system's
//! code page.
//!
//! ## Remarks
//!
//! The `Chr` function is the inverse of the `Asc` function. While `Asc` returns the numeric
//! character code of a character, `Chr` returns the character for a given code.
//!
//! **Important Characteristics:**
//!
//! - Valid range: 0-255 (Error 5 "Invalid procedure call or argument" for values outside range)
//! - Chr(0) returns a null character (vbNullChar)
//! - Chr(13) returns carriage return (vbCr)
//! - Chr(10) returns line feed (vbLf)
//! - Chr(9) returns tab character (vbTab)
//! - Values 0-31 are non-printable control characters
//! - Values 32-126 are standard printable ASCII characters
//! - Values 127-255 depend on the system code page (often Windows-1252 in VB6)
//!
//! ## Common Character Codes
//!
//! | Code | Character | Constant | Description |
//! |------|-----------|----------|-------------|
//! | 0 | (null) | vbNullChar | Null character |
//! | 9 | \t | vbTab | Horizontal tab |
//! | 10 | \n | vbLf | Line feed |
//! | 13 | \r | vbCr | Carriage return |
//! | 32 | (space) | - | Space character |
//! | 34 | " | - | Double quote |
//! | 39 | ' | - | Single quote |
//! | 65 | A | - | Uppercase A |
//! | 97 | a | - | Lowercase a |
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Get character from code
//! Dim ch As String
//! ch = Chr(65)  ' Returns "A"
//! ch = Chr(97)  ' Returns "a"
//! ch = Chr(48)  ' Returns "0"
//!
//! ' Special characters
//! ch = Chr(32)  ' Returns space " "
//! ch = Chr(13)  ' Returns carriage return
//! ch = Chr(10)  ' Returns line feed
//! ```
//!
//! ### Line Breaks and Special Characters
//!
//! ```vb
//! ' Create multi-line string
//! Dim msg As String
//! msg = "Line 1" & Chr(13) & Chr(10) & "Line 2"
//! ' Or use the constant
//! msg = "Line 1" & vbCrLf & "Line 2"
//!
//! ' Tab-separated values
//! Dim data As String
//! data = "Name" & Chr(9) & "Age" & Chr(9) & "City"
//! ```
//!
//! ### Building Strings from Codes
//!
//! ```vb
//! ' Build alphabet
//! Dim i As Integer
//! Dim alphabet As String
//! For i = 65 To 90
//!     alphabet = alphabet & Chr(i)
//! Next i
//! ' alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
//! ```
//!
//! ## Common Patterns
//!
//! ### Generating Character Sequences
//!
//! ```vb
//! Function GetAlphabet(uppercase As Boolean) As String
//!     Dim result As String
//!     Dim i As Integer
//!     Dim startCode As Integer
//!     
//!     If uppercase Then
//!         startCode = 65  ' 'A'
//!     Else
//!         startCode = 97  ' 'a'
//!     End If
//!     
//!     For i = startCode To startCode + 25
//!         result = result & Chr(i)
//!     Next i
//!     
//!     GetAlphabet = result
//! End Function
//! ```
//!
//! ### Quote Handling
//!
//! ```vb
//! Function QuoteString(text As String) As String
//!     QuoteString = Chr(34) & text & Chr(34)
//! End Function
//!
//! ' Usage: result = QuoteString("Hello")  ' Returns: "Hello"
//! ```
//!
//! ### CSV Generation
//!
//! ```vb
//! Function CreateCSVRow(ParamArray fields() As Variant) As String
//!     Dim result As String
//!     Dim i As Integer
//!     Dim field As String
//!     
//!     For i = LBound(fields) To UBound(fields)
//!         field = CStr(fields(i))
//!         
//!         ' Quote fields containing commas or quotes
//!         If InStr(field, ",") > 0 Or InStr(field, Chr(34)) > 0 Then
//!             field = Chr(34) & Replace(field, Chr(34), Chr(34) & Chr(34)) & Chr(34)
//!         End If
//!         
//!         If i > LBound(fields) Then result = result & ","
//!         result = result & field
//!     Next i
//!     
//!     CreateCSVRow = result
//! End Function
//! ```
//!
//! ### Control Character Removal
//!
//! ```vb
//! Function RemoveControlChars(text As String) As String
//!     Dim result As String
//!     Dim i As Integer
//!     Dim ch As String
//!     Dim code As Integer
//!     
//!     For i = 1 To Len(text)
//!         ch = Mid(text, i, 1)
//!         code = Asc(ch)
//!         
//!         ' Keep only printable characters (32-126) and common whitespace
//!         If code >= 32 Or code = 9 Or code = 10 Or code = 13 Then
//!             result = result & ch
//!         End If
//!     Next i
//!     
//!     RemoveControlChars = result
//! End Function
//! ```
//!
//! ### String Encoding/Decoding
//!
//! ```vb
//! Function EncodeString(text As String) As String
//!     Dim result As String
//!     Dim i As Integer
//!     
//!     For i = 1 To Len(text)
//!         If i > 1 Then result = result & ","
//!         result = result & CStr(Asc(Mid(text, i, 1)))
//!     Next i
//!     
//!     EncodeString = result
//! End Function
//!
//! Function DecodeString(encoded As String) As String
//!     Dim result As String
//!     Dim codes() As String
//!     Dim i As Integer
//!     
//!     codes = Split(encoded, ",")
//!     For i = LBound(codes) To UBound(codes)
//!         result = result & Chr(CLng(codes(i)))
//!     Next i
//!     
//!     DecodeString = result
//! End Function
//! ```
//!
//! ### Random Character Generation
//!
//! ```vb
//! Function GeneratePassword(length As Integer) As String
//!     Dim result As String
//!     Dim i As Integer
//!     Dim charType As Integer
//!     
//!     Randomize
//!     
//!     For i = 1 To length
//!         charType = Int(Rnd * 3)  ' 0=uppercase, 1=lowercase, 2=digit
//!         
//!         Select Case charType
//!             Case 0  ' Uppercase A-Z
//!                 result = result & Chr(65 + Int(Rnd * 26))
//!             Case 1  ' Lowercase a-z
//!                 result = result & Chr(97 + Int(Rnd * 26))
//!             Case 2  ' Digit 0-9
//!                 result = result & Chr(48 + Int(Rnd * 10))
//!         End Select
//!     Next i
//!     
//!     GeneratePassword = result
//! End Function
//! ```
//!
//! ### Box Drawing Characters
//!
//! ```vb
//! Function DrawBox(width As Integer, height As Integer) As String
//!     Dim result As String
//!     Dim i As Integer
//!     
//!     ' Top border (using extended ASCII box characters)
//!     result = Chr(218)  ' Top-left corner
//!     For i = 1 To width - 2
//!         result = result & Chr(196)  ' Horizontal line
//!     Next i
//!     result = result & Chr(191) & vbCrLf  ' Top-right corner
//!     
//!     ' Middle rows
//!     For i = 1 To height - 2
//!         result = result & Chr(179)  ' Vertical line
//!         result = result & Space(width - 2)
//!         result = result & Chr(179) & vbCrLf  ' Vertical line
//!     Next i
//!     
//!     ' Bottom border
//!     result = result & Chr(192)  ' Bottom-left corner
//!     For i = 1 To width - 2
//!         result = result & Chr(196)  ' Horizontal line
//!     Next i
//!     result = result & Chr(217)  ' Bottom-right corner
//!     
//!     DrawBox = result
//! End Function
//! ```
//!
//! ### Character Case Conversion
//!
//! ```vb
//! Function ToggleCase(text As String) As String
//!     Dim result As String
//!     Dim i As Integer
//!     Dim ch As String
//!     Dim code As Integer
//!     
//!     For i = 1 To Len(text)
//!         ch = Mid(text, i, 1)
//!         code = Asc(ch)
//!         
//!         If code >= 65 And code <= 90 Then
//!             ' Uppercase -> lowercase
//!             result = result & Chr(code + 32)
//!         ElseIf code >= 97 And code <= 122 Then
//!             ' Lowercase -> uppercase
//!             result = result & Chr(code - 32)
//!         Else
//!             result = result & ch
//!         End If
//!     Next i
//!     
//!     ToggleCase = result
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Binary Data Handling
//!
//! ```vb
//! Function BytesToString(bytes() As Byte) As String
//!     Dim result As String
//!     Dim i As Long
//!     
//!     For i = LBound(bytes) To UBound(bytes)
//!         result = result & Chr(bytes(i))
//!     Next i
//!     
//!     BytesToString = result
//! End Function
//!
//! Function StringToBytes(text As String) As Byte()
//!     Dim bytes() As Byte
//!     Dim i As Long
//!     
//!     ReDim bytes(1 To Len(text))
//!     
//!     For i = 1 To Len(text)
//!         bytes(i) = Asc(Mid(text, i, 1))
//!     Next i
//!     
//!     StringToBytes = bytes
//! End Function
//! ```
//!
//! ### Unicode Support (`ChrW` variant)
//!
//! ```vb
//! ' Note: VB6 has ChrW for Unicode characters
//! Function GetUnicodeChar(code As Long) As String
//!     ' For codes 0-255, Chr and ChrW are equivalent
//!     If code <= 255 Then
//!         GetUnicodeChar = Chr(code)
//!     Else
//!         ' For codes > 255, use ChrW (not covered by Chr function)
//!         GetUnicodeChar = ChrW(code)
//!     End If
//! End Function
//! ```
//!
//! ### Escape Sequence Processing
//!
//! ```vb
//! Function ProcessEscapeSequences(text As String) As String
//!     Dim result As String
//!     result = text
//!     
//!     ' Replace common escape sequences
//!     result = Replace(result, "\n", Chr(10))   ' Line feed
//!     result = Replace(result, "\r", Chr(13))   ' Carriage return
//!     result = Replace(result, "\t", Chr(9))    ' Tab
//!     result = Replace(result, "\\", Chr(92))   ' Backslash
//!     result = Replace(result, "\""", Chr(34))  ' Double quote
//!     
//!     ProcessEscapeSequences = result
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeChr(charcode As Long) As String
//!     On Error GoTo ErrorHandler
//!     
//!     If charcode < 0 Or charcode > 255 Then
//!         Err.Raise 5, , "Invalid character code: " & charcode
//!     End If
//!     
//!     SafeChr = Chr(charcode)
//!     Exit Function
//!     
//! ErrorHandler:
//!     MsgBox "Error in Chr: " & Err.Description
//!     SafeChr = ""
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 5** (Invalid procedure call or argument): Character code is outside the range 0-255
//! - **Error 13** (Type mismatch): Argument is not numeric
//!
//! ## Performance Considerations
//!
//! - `Chr` is a fast function with minimal overhead
//! - For building long strings with many `Chr` calls, consider using a `StringBuilder` pattern
//! - Avoid repeated `Chr` calls for the same character code (use a constant instead)
//! - For Unicode support beyond 255, use `ChrW` or `ChrB` functions
//!
//! ## VB6 String Constants vs `Chr`
//!
//! VB6 provides built-in constants for common characters:
//!
//! ```vb
//! ' Prefer constants over Chr for readability
//! vbCr        ' Chr(13) - Carriage return
//! vbLf        ' Chr(10) - Line feed
//! vbCrLf      ' Chr(13) & Chr(10) - Carriage return + line feed
//! vbTab       ' Chr(9) - Tab
//! vbNullChar  ' Chr(0) - Null character
//! vbNullString ' Empty string ""
//! ```
//!
//! ## Limitations
//!
//! - Limited to character codes 0-255 (single-byte characters)
//! - For Unicode beyond 255, use `ChrW` instead
//! - Character interpretation depends on system code page
//! - Control characters (0-31) may not display in UI controls
//! - Extended ASCII (128-255) may vary across systems
//!
//! ## Related Functions
//!
//! - `Asc`: Returns the character code of the first character in a string (inverse of Chr)
//! - `ChrW`: Returns Unicode character for character codes 0-65535
//! - `ChrB`: Returns a byte containing the character code
//! - `AscW`: Returns the Unicode character code
//! - `AscB`: Returns the byte value

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn chr_basic() {
        let source = r"
ch = Chr(65)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("ch"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
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
        ]);
    }

    #[test]
    fn chr_with_variable() {
        let source = r"
result = Chr(code)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
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
        ]);
    }

    #[test]
    fn chr_special_characters() {
        let source = r"
tab = Chr(9)
lf = Chr(10)
cr = Chr(13)
space = Chr(32)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("tab"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
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
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("lf"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
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
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("cr"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
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
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("space"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
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
        ]);
    }

    #[test]
    fn chr_in_concatenation() {
        let source = r#"
msg = "Line 1" & Chr(13) & Chr(10) & "Line 2"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
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
                                Identifier ("Chr"),
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
                            Identifier ("Chr"),
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
        ]);
    }

    #[test]
    fn chr_in_loop() {
        let source = r"
For i = 65 To 90
    alphabet = alphabet & Chr(i)
Next i
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            ForStatement {
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
                            Identifier ("alphabet"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("alphabet"),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Chr"),
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
                },
                NextKeyword,
                Whitespace,
                Identifier ("i"),
                Newline,
            },
        ]);
    }

    #[test]
    fn chr_quote_character() {
        let source = r"
quote = Chr(34)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("quote"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
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
                Newline,
            },
        ]);
    }

    #[test]
    fn chr_in_function() {
        let source = r"
Function QuoteString(text As String) As String
    QuoteString = Chr(34) & text & Chr(34)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("QuoteString"),
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
                            Identifier ("QuoteString"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("Chr"),
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
                                Identifier ("Chr"),
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
    fn chr_with_expression() {
        let source = r"
ch = Chr(65 + offset)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("ch"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            BinaryExpression {
                                NumericLiteralExpression {
                                    IntegerLiteral ("65"),
                                },
                                Whitespace,
                                AdditionOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("offset"),
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
    fn chr_in_if_statement() {
        let source = r"
If ch = Chr(13) Then
    ProcessNewline
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
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
                        Identifier ("Chr"),
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
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("ProcessNewline"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn chr_multiple_calls() {
        let source = r"
line = Chr(218) & Chr(196) & Chr(191)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            LineInputStatement {
                LineKeyword,
                Whitespace,
                EqualityOperator,
                Whitespace,
                Identifier ("Chr"),
                LeftParenthesis,
                IntegerLiteral ("218"),
                RightParenthesis,
                Whitespace,
                Ampersand,
                Whitespace,
                Identifier ("Chr"),
                LeftParenthesis,
                IntegerLiteral ("196"),
                RightParenthesis,
                Whitespace,
                Ampersand,
                Whitespace,
                Identifier ("Chr"),
                LeftParenthesis,
                IntegerLiteral ("191"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn chr_in_assignment() {
        let source = r"
Dim separator As String
separator = Chr(9)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("separator"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("separator"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
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
                Newline,
            },
        ]);
    }

    #[test]
    fn chr_in_select_case() {
        let source = r"
Select Case ch
    Case Chr(13)
        HandleCR
    Case Chr(10)
        HandleLF
End Select
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SelectCaseStatement {
                SelectKeyword,
                Whitespace,
                CaseKeyword,
                Whitespace,
                IdentifierExpression {
                    Identifier ("ch"),
                },
                Newline,
                Whitespace,
                CaseClause {
                    CaseKeyword,
                    Whitespace,
                    Identifier ("Chr"),
                    LeftParenthesis,
                    IntegerLiteral ("13"),
                    RightParenthesis,
                    Newline,
                    StatementList {
                        Whitespace,
                        CallStatement {
                            Identifier ("HandleCR"),
                            Newline,
                        },
                        Whitespace,
                    },
                },
                CaseClause {
                    CaseKeyword,
                    Whitespace,
                    Identifier ("Chr"),
                    LeftParenthesis,
                    IntegerLiteral ("10"),
                    RightParenthesis,
                    Newline,
                    StatementList {
                        Whitespace,
                        CallStatement {
                            Identifier ("HandleLF"),
                            Newline,
                        },
                    },
                },
                EndKeyword,
                Whitespace,
                SelectKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn chr_with_asc() {
        let source = r"
original = Chr(Asc(text))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("original"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Asc"),
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
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn chr_in_while_loop() {
        let source = r"
While i <= 90
    result = result & Chr(i)
    i = i + 1
Wend
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            WhileStatement {
                WhileKeyword,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("i"),
                    },
                    Whitespace,
                    LessThanOrEqualOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("90"),
                    },
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
                                Identifier ("Chr"),
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("i"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("i"),
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
                        },
                        Newline,
                    },
                },
                WendKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn chr_in_do_loop() {
        let source = r"
Do While i < 256
    chars = chars & Chr(i)
    i = i + 1
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DoStatement {
                DoKeyword,
                Whitespace,
                WhileKeyword,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("i"),
                    },
                    Whitespace,
                    LessThanOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("256"),
                    },
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("chars"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("chars"),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Chr"),
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("i"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("i"),
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
                        },
                        Newline,
                    },
                },
                LoopKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn chr_null_character() {
        let source = r"
nullChar = Chr(0)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("nullChar"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
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
        ]);
    }

    #[test]
    fn chr_with_line_continuation() {
        let source = r#"
msg = "Text" & _
      Chr(13) & _
      Chr(10) & _
      "More text"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
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
                                StringLiteral ("\"Text\""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            Underscore,
                            Newline,
                            Whitespace,
                            CallExpression {
                                Identifier ("Chr"),
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
                        Underscore,
                        Newline,
                        Whitespace,
                        CallExpression {
                            Identifier ("Chr"),
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
                    Underscore,
                    Newline,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"More text\""),
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn chr_in_array() {
        let source = r"
chars(i) = Chr(code)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                CallExpression {
                    Identifier ("chars"),
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
                    Identifier ("Chr"),
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
        ]);
    }

    #[test]
    fn chr_in_msgbox() {
        let source = r#"
MsgBox "Line 1" & Chr(13) & "Line 2"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("MsgBox"),
                Whitespace,
                StringLiteral ("\"Line 1\""),
                Whitespace,
                Ampersand,
                Whitespace,
                Identifier ("Chr"),
                LeftParenthesis,
                IntegerLiteral ("13"),
                RightParenthesis,
                Whitespace,
                Ampersand,
                Whitespace,
                StringLiteral ("\"Line 2\""),
                Newline,
            },
        ]);
    }

    #[test]
    fn chr_extended_ascii() {
        let source = r"
boxChar = Chr(196)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("boxChar"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("196"),
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
    fn chr_in_replace() {
        let source = r#"
result = Replace(text, Chr(13), "")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
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
                                TextKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            CallExpression {
                                Identifier ("Chr"),
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
        ]);
    }

    #[test]
    fn chr_with_mod() {
        let source = r"
ch = Chr(value Mod 256)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("ch"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("value"),
                                },
                                Whitespace,
                                ModKeyword,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("256"),
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
    fn chr_in_split() {
        let source = r"
parts = Split(data, Chr(9))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("parts"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Split"),
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
                            CallExpression {
                                Identifier ("Chr"),
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
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn chr_with_cint() {
        let source = r"
ch = Chr(CInt(value))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("ch"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("CInt"),
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
        ]);
    }

    #[test]
    fn chr_with_whitespace() {
        let source = r"
result = Chr( 65 )
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Chr"),
                    LeftParenthesis,
                    ArgumentList {
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("65"),
                            },
                        },
                        Whitespace,
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }
}
