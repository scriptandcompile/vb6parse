//! # `AscW` Function
//!
//! Returns an `Integer` representing the Unicode character code of the first character in a string.
//! The "W" suffix indicates this is the wide (Unicode) version of the `Asc` function.
//!
//! ## Syntax
//!
//! ```vb
//! AscW(string)
//! ```
//!
//! ## Parameters
//!
//! - **string**: Required. Any valid string expression. If the string contains no characters,
//!   a runtime error occurs (Error 5: Invalid procedure call or argument).
//!
//! ## Returns
//!
//! Returns an `Integer` representing the Unicode code point (0-65535) of the first character in the string.
//!
//! ## Remarks
//!
//! - `AscW` returns the Unicode (UTF-16) code point of the first character in a string.
//! - The W suffix stands for "Wide", distinguishing it from the ANSI `AscB` function.
//! - Return values range from 0 to 65535, covering the Basic Multilingual Plane (BMP) of Unicode.
//! - For ASCII characters (0-127), `AscW` and `Asc` return the same values.
//! - For characters in the extended ASCII range (128-255), results match the Latin-1 supplement in Unicode.
//! - If the string is empty (`""`), a runtime error occurs (Error 5).
//! - `AscW` is essential for working with international text and Unicode characters.
//! - The inverse function is `ChrW`, which converts a Unicode code point back to a character.
//! - Characters outside the BMP (above 65535) are represented as surrogate pairs in VB6.
//! - Surrogate pair characters will return the code of the high surrogate (0xD800-0xDBFF).
//!
//! ## Typical Uses
//!
//! 1. **International text processing** - Work with characters from various languages
//! 2. **Unicode character analysis** - Examine Unicode code points in strings
//! 3. **Character validation** - Validate characters are in expected Unicode ranges
//! 4. **Text encoding operations** - Convert between different character encodings
//! 5. **Symbol detection** - Identify mathematical symbols, currency, arrows, etc.
//! 6. **Character range checking** - Determine if characters belong to specific scripts
//! 7. **Multilingual sorting** - Implement custom sort orders based on Unicode values
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Simple ASCII character
//! Dim code As Integer
//! code = AscW("A")  ' Returns 65
//! ```
//!
//! ```vb
//! ' Example 2: Euro symbol
//! Dim euroCode As Integer
//! euroCode = AscW("€")  ' Returns 8364
//! ```
//!
//! ```vb
//! ' Example 3: Greek letter
//! Dim alphaCode As Integer
//! alphaCode = AscW("α")  ' Returns 945
//! ```
//!
//! ```vb
//! ' Example 4: Chinese character
//! Dim hanziCode As Integer
//! hanziCode = AscW("中")  ' Returns 20013
//! ```
//!
//! ## Common Patterns
//!
//! ### Check Unicode Range
//! ```vb
//! Function IsInUnicodeRange(char As String, rangeStart As Long, rangeEnd As Long) As Boolean
//!     If Len(char) = 0 Then Exit Function
//!     Dim code As Long
//!     code = AscW(char)
//!     IsInUnicodeRange = (code >= rangeStart And code <= rangeEnd)
//! End Function
//! ```
//!
//! ### Detect Character Script
//! ```vb
//! Function GetCharacterScript(char As String) As String
//!     If Len(char) = 0 Then Exit Function
//!     
//!     Dim code As Long
//!     code = AscW(char)
//!     
//!     Select Case code
//!         Case 0 To 127
//!             GetCharacterScript = "ASCII"
//!         Case 880 To 1023
//!             GetCharacterScript = "Greek"
//!         Case 1024 To 1279
//!             GetCharacterScript = "Cyrillic"
//!         Case 1424 To 1535
//!             GetCharacterScript = "Hebrew"
//!         Case 1536 To 1791
//!             GetCharacterScript = "Arabic"
//!         Case 19968 To 40959
//!             GetCharacterScript = "CJK"
//!         Case 44032 To 55203
//!             GetCharacterScript = "Hangul"
//!         Case Else
//!             GetCharacterScript = "Other"
//!     End Select
//! End Function
//! ```
//!
//! ### Validate Latin Characters
//! ```vb
//! Function IsLatinChar(char As String) As Boolean
//!     If Len(char) = 0 Then Exit Function
//!     Dim code As Long
//!     code = AscW(char)
//!     ' Basic Latin + Latin-1 Supplement + Latin Extended-A and B
//!     IsLatinChar = (code >= 0 And code <= 591)
//! End Function
//! ```
//!
//! ### Check for Symbol Characters
//! ```vb
//! Function IsSymbol(char As String) As Boolean
//!     If Len(char) = 0 Then Exit Function
//!     
//!     Dim code As Long
//!     code = AscW(char)
//!     
//!     ' Common symbol ranges
//!     IsSymbol = (code >= 8192 And code <= 8303) Or _
//!                (code >= 8352 And code <= 8399) Or _
//!                (code >= 8448 And code <= 8527) Or _
//!                (code >= 8704 And code <= 8959) Or _
//!                (code >= 9632 And code <= 9727)
//! End Function
//! ```
//!
//! ### Compare Unicode Values
//! ```vb
//! Function CompareUnicode(char1 As String, char2 As String) As Integer
//!     If Len(char1) = 0 Or Len(char2) = 0 Then Exit Function
//!     CompareUnicode = AscW(char1) - AscW(char2)
//! End Function
//! ```
//!
//! ### Detect Emoji (BMP only)
//! ```vb
//! Function IsEmojiBMP(char As String) As Boolean
//!     If Len(char) = 0 Then Exit Function
//!     
//!     Dim code As Long
//!     code = AscW(char)
//!     
//!     ' Emoticons and Miscellaneous Symbols
//!     IsEmojiBMP = (code >= 9728 And code <= 9983) Or _
//!                  (code >= 10084 And code <= 10084) Or _
//!                  (code >= 127744 And code <= 128511)
//! End Function
//! ```
//!
//! ### Extract Unicode Array
//! ```vb
//! Function GetUnicodeArray(text As String) As Variant
//!     Dim codes() As Long
//!     Dim i As Long
//!     
//!     If Len(text) = 0 Then Exit Function
//!     
//!     ReDim codes(1 To Len(text))
//!     For i = 1 To Len(text)
//!         codes(i) = AscW(Mid(text, i, 1))
//!     Next i
//!     
//!     GetUnicodeArray = codes
//! End Function
//! ```
//!
//! ### Check Diacritical Marks
//! ```vb
//! Function IsDiacriticalMark(char As String) As Boolean
//!     If Len(char) = 0 Then Exit Function
//!     
//!     Dim code As Long
//!     code = AscW(char)
//!     
//!     ' Combining Diacritical Marks
//!     IsDiacriticalMark = (code >= 768 And code <= 879)
//! End Function
//! ```
//!
//! ### Validate Email Characters
//! ```vb
//! Function IsValidEmailChar(char As String) As Boolean
//!     If Len(char) = 0 Then Exit Function
//!     
//!     Dim code As Long
//!     code = AscW(char)
//!     
//!     ' Alphanumeric, dot, hyphen, underscore, @
//!     IsValidEmailChar = (code >= 48 And code <= 57) Or _
//!                        (code >= 65 And code <= 90) Or _
//!                        (code >= 97 And code <= 122) Or _
//!                        code = 45 Or code = 46 Or code = 64 Or code = 95
//! End Function
//! ```
//!
//! ### Detect Control Characters
//! ```vb
//! Function IsUnicodeControl(char As String) As Boolean
//!     If Len(char) = 0 Then Exit Function
//!     
//!     Dim code As Long
//!     code = AscW(char)
//!     
//!     ' C0 and C1 control characters
//!     IsUnicodeControl = (code >= 0 And code <= 31) Or _
//!                        (code >= 127 And code <= 159)
//! End Function
//! ```
//!
//! ## Advanced Examples
//!
//! ### Unicode Normalization Check
//! ```vb
//! Function CompareNormalized(str1 As String, str2 As String) As Boolean
//!     ' Simple comparison ignoring case
//!     If Len(str1) <> Len(str2) Then
//!         CompareNormalized = False
//!         Exit Function
//!     End If
//!     
//!     Dim i As Long
//!     Dim code1 As Long, code2 As Long
//!     
//!     For i = 1 To Len(str1)
//!         code1 = AscW(Mid(str1, i, 1))
//!         code2 = AscW(Mid(str2, i, 1))
//!         
//!         ' Convert to lowercase if uppercase Latin
//!         If code1 >= 65 And code1 <= 90 Then code1 = code1 + 32
//!         If code2 >= 65 And code2 <= 90 Then code2 = code2 + 32
//!         
//!         If code1 <> code2 Then
//!             CompareNormalized = False
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     CompareNormalized = True
//! End Function
//! ```
//!
//! ### Unicode to HTML Entity
//! ```vb
//! Function UnicodeToHTMLEntity(char As String) As String
//!     If Len(char) = 0 Then Exit Function
//!     
//!     Dim code As Long
//!     code = AscW(char)
//!     
//!     ' Create numeric entity
//!     UnicodeToHTMLEntity = "&#" & code & ";"
//! End Function
//! ```
//!
//! ### Multilingual Text Analyzer
//! ```vb
//! Function AnalyzeText(text As String) As String
//!     Dim i As Long
//!     Dim code As Long
//!     Dim latinCount As Long
//!     Dim cjkCount As Long
//!     Dim arabicCount As Long
//!     Dim otherCount As Long
//!     
//!     For i = 1 To Len(text)
//!         code = AscW(Mid(text, i, 1))
//!         
//!         Select Case code
//!             Case 0 To 591
//!                 latinCount = latinCount + 1
//!             Case 19968 To 40959
//!                 cjkCount = cjkCount + 1
//!             Case 1536 To 1791
//!                 arabicCount = arabicCount + 1
//!             Case Else
//!                 otherCount = otherCount + 1
//!         End Select
//!     Next i
//!     
//!     AnalyzeText = "Latin: " & latinCount & ", CJK: " & cjkCount & _
//!                   ", Arabic: " & arabicCount & ", Other: " & otherCount
//! End Function
//! ```
//!
//! ### Character Category Validator
//! ```vb
//! Function ValidateCategory(text As String, category As String) As Boolean
//!     Dim i As Long
//!     Dim code As Long
//!     
//!     For i = 1 To Len(text)
//!         code = AscW(Mid(text, i, 1))
//!         
//!         Select Case category
//!             Case "Digit"
//!                 If Not (code >= 48 And code <= 57) Then
//!                     ValidateCategory = False
//!                     Exit Function
//!                 End If
//!             Case "UpperLatin"
//!                 If Not (code >= 65 And code <= 90) Then
//!                     ValidateCategory = False
//!                     Exit Function
//!                 End If
//!             Case "LowerLatin"
//!                 If Not (code >= 97 And code <= 122) Then
//!                     ValidateCategory = False
//!                     Exit Function
//!                 End If
//!             Case Else
//!                 ValidateCategory = False
//!                 Exit Function
//!         End Select
//!     Next i
//!     
//!     ValidateCategory = True
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeAscW(text As String) As Long
//!     On Error GoTo ErrorHandler
//!     
//!     If Len(text) = 0 Then
//!         SafeAscW = -1  ' Error indicator
//!         Exit Function
//!     End If
//!     
//!     SafeAscW = AscW(text)
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeAscW = -1
//! End Function
//! ```
//!
//! ## Performance Notes
//!
//! - `AscW` is a very fast operation with minimal overhead
//! - When processing long strings character-by-character, use `Mid` function efficiently
//! - For repeated code point lookups, consider caching results
//! - `AscW` is faster than string comparison for Unicode operations
//! - No significant performance difference between `Asc` and `AscW` on modern systems
//!
//! ## Best Practices
//!
//! 1. **Validate input** - Always check for empty strings before calling `AscW`
//! 2. **Use for Unicode** - Prefer `AscW` over `Asc` when working with international text
//! 3. **Handle errors** - Wrap `AscW` calls in error handlers when processing untrusted input
//! 4. **Document ranges** - Use constants or comments to explain Unicode range checks
//! 5. **Consider normalization** - Be aware that some characters have multiple Unicode representations
//! 6. **Use with `ChrW`** - Pair with `ChrW` for Unicode code point conversions
//! 7. **Test edge cases** - Verify behavior with empty strings, control characters, and non-BMP characters
//!
//! ## Comparison with Related Functions
//!
//! | Function | Returns | Character Set | Use Case |
//! |----------|---------|---------------|----------|
//! | `Asc` | Integer (0-255 or Unicode) | System default | General character codes |
//! | `AscB` | Integer (0-255) | ANSI byte value | Byte-level operations |
//! | `AscW` | Integer (0-65535) | Unicode code point | International text |
//! | `ChrW` | String (Unicode) | Unicode (inverse) | Convert code to character |
//!
//! ## Unicode Ranges Reference
//!
//! Common Unicode ranges that can be detected with `AscW`:
//!
//! - **Basic Latin (ASCII)**: 0-127
//! - **Latin-1 Supplement**: 128-255
//! - **Latin Extended-A**: 256-383
//! - **Greek and Coptic**: 880-1023
//! - **Cyrillic**: 1024-1279
//! - **Hebrew**: 1424-1535
//! - **Arabic**: 1536-1791
//! - **Devanagari**: 2304-2431
//! - **Thai**: 3584-3711
//! - **Tibetan**: 3840-4095
//! - **CJK Unified Ideographs**: 19968-40959
//! - **Hangul Syllables**: 44032-55203
//! - **Currency Symbols**: 8352-8399
//! - **Mathematical Operators**: 8704-8959
//! - **Arrows**: 8592-8703
//! - **Box Drawing**: 9472-9599
//! - **Emoticons**: 9728-9983
//!
//! ## Platform Notes
//!
//! - VB6 uses UTF-16 internally on Windows NT-based systems for Unicode support
//! - On Windows 95/98/ME, Unicode support is limited and may not work correctly
//! - Modern Windows systems (XP and later) have full Unicode support
//! - `AscW` returns consistent values across different code pages
//! - For maximum compatibility, test on target platforms with actual Unicode data
//!
//! ## Limitations
//!
//! - `AscW` only returns code points in the Basic Multilingual Plane (0-65535)
//! - Characters outside the BMP require surrogate pairs and special handling
//! - Combining characters are treated as separate code points
//! - Some characters may display differently depending on available fonts
//! - Grapheme clusters (like emoji with modifiers) are not handled as single units
//! - Runtime error occurs with empty strings
//! - No built-in normalization (characters with multiple representations)

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn ascw_simple_ascii() {
        let source = r#"
Sub Test()
    code = AscW("A")
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
                            Identifier ("code"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"A\""),
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
    fn ascw_euro_symbol() {
        let source = r"
Sub Test()
    euroCode = AscW(euroChar)
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
                            Identifier ("euroCode"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("euroChar"),
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
    fn ascw_greek_letter() {
        let source = r"
Sub Test()
    alphaCode = AscW(greekChar)
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
                            Identifier ("alphaCode"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("greekChar"),
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
    fn ascw_unicode_range_check() {
        let source = r"
Function IsInUnicodeRange(char As String, rangeStart As Long, rangeEnd As Long) As Boolean
    Dim code As Long
    code = AscW(char)
    IsInUnicodeRange = (code >= rangeStart And code <= rangeEnd)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("IsInUnicodeRange"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("char"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("rangeStart"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("rangeEnd"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
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
                        Identifier ("code"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("code"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("char"),
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
                            Identifier ("IsInUnicodeRange"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("code"),
                                    },
                                    Whitespace,
                                    GreaterThanOrEqualOperator,
                                    Whitespace,
                                    IdentifierExpression {
                                        Identifier ("rangeStart"),
                                    },
                                },
                                Whitespace,
                                AndKeyword,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("code"),
                                    },
                                    Whitespace,
                                    LessThanOrEqualOperator,
                                    Whitespace,
                                    IdentifierExpression {
                                        Identifier ("rangeEnd"),
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
    fn ascw_character_script_detection() {
        let source = r#"
Function GetCharacterScript(char As String) As String
    Dim code As Long
    code = AscW(char)
    If code >= 880 And code <= 1023 Then
        GetCharacterScript = "Greek"
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
                Identifier ("GetCharacterScript"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("char"),
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
                        Identifier ("code"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("code"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("char"),
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
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("code"),
                                },
                                Whitespace,
                                GreaterThanOrEqualOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("880"),
                                },
                            },
                            Whitespace,
                            AndKeyword,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("code"),
                                },
                                Whitespace,
                                LessThanOrEqualOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("1023"),
                                },
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("GetCharacterScript"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                StringLiteralExpression {
                                    StringLiteral ("\"Greek\""),
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
    fn ascw_latin_validation() {
        let source = r"
Function IsLatinChar(char As String) As Boolean
    Dim code As Long
    code = AscW(char)
    IsLatinChar = (code >= 0 And code <= 591)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("IsLatinChar"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("char"),
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
                        Identifier ("code"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("code"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("char"),
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
                            Identifier ("IsLatinChar"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("code"),
                                    },
                                    Whitespace,
                                    GreaterThanOrEqualOperator,
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
                                        Identifier ("code"),
                                    },
                                    Whitespace,
                                    LessThanOrEqualOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("591"),
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
    fn ascw_symbol_check() {
        let source = r"
Function IsSymbol(char As String) As Boolean
    Dim code As Long
    code = AscW(char)
    IsSymbol = (code >= 8192 And code <= 8303)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("IsSymbol"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("char"),
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
                        Identifier ("code"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("code"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("char"),
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
                            Identifier ("IsSymbol"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("code"),
                                    },
                                    Whitespace,
                                    GreaterThanOrEqualOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("8192"),
                                    },
                                },
                                Whitespace,
                                AndKeyword,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("code"),
                                    },
                                    Whitespace,
                                    LessThanOrEqualOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("8303"),
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
    fn ascw_unicode_compare() {
        let source = r"
Function CompareUnicode(char1 As String, char2 As String) As Integer
    CompareUnicode = AscW(char1) - AscW(char2)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CompareUnicode"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("char1"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("char2"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
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
                            Identifier ("CompareUnicode"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("AscW"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("char1"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            SubtractionOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("AscW"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("char2"),
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
    fn ascw_emoji_detection() {
        let source = r"
Function IsEmojiBMP(char As String) As Boolean
    Dim code As Long
    code = AscW(char)
    IsEmojiBMP = (code >= 9728 And code <= 9983)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("IsEmojiBMP"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("char"),
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
                        Identifier ("code"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("code"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("char"),
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
                            Identifier ("IsEmojiBMP"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("code"),
                                    },
                                    Whitespace,
                                    GreaterThanOrEqualOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("9728"),
                                    },
                                },
                                Whitespace,
                                AndKeyword,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("code"),
                                    },
                                    Whitespace,
                                    LessThanOrEqualOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("9983"),
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
    fn ascw_unicode_array() {
        let source = r"
Function GetUnicodeArray(text As String) As Variant
    Dim codes() As Long
    Dim i As Long
    For i = 1 To Len(text)
        codes(i) = AscW(Mid(text, i, 1))
    Next i
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetUnicodeArray"),
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
                VariantKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("codes"),
                        LeftParenthesis,
                        RightParenthesis,
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
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
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
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("AscW"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            CallExpression {
                                                MidKeyword,
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
                                                            Identifier ("i"),
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
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn ascw_diacritical_marks() {
        let source = r"
Function IsDiacriticalMark(char As String) As Boolean
    Dim code As Long
    code = AscW(char)
    IsDiacriticalMark = (code >= 768 And code <= 879)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("IsDiacriticalMark"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("char"),
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
                        Identifier ("code"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("code"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("char"),
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
                            Identifier ("IsDiacriticalMark"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("code"),
                                    },
                                    Whitespace,
                                    GreaterThanOrEqualOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("768"),
                                    },
                                },
                                Whitespace,
                                AndKeyword,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("code"),
                                    },
                                    Whitespace,
                                    LessThanOrEqualOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("879"),
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
    fn ascw_email_validation() {
        let source = r"
Function IsValidEmailChar(char As String) As Boolean
    Dim code As Long
    code = AscW(char)
    IsValidEmailChar = (code >= 48 And code <= 57) Or code = 64
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("IsValidEmailChar"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("char"),
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
                        Identifier ("code"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("code"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("char"),
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
                            Identifier ("IsValidEmailChar"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            ParenthesizedExpression {
                                LeftParenthesis,
                                BinaryExpression {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("code"),
                                        },
                                        Whitespace,
                                        GreaterThanOrEqualOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("48"),
                                        },
                                    },
                                    Whitespace,
                                    AndKeyword,
                                    Whitespace,
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("code"),
                                        },
                                        Whitespace,
                                        LessThanOrEqualOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("57"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            OrKeyword,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("code"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("64"),
                                },
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
    fn ascw_control_character() {
        let source = r"
Function IsUnicodeControl(char As String) As Boolean
    Dim code As Long
    code = AscW(char)
    IsUnicodeControl = (code >= 0 And code <= 31)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("IsUnicodeControl"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("char"),
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
                        Identifier ("code"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("code"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("char"),
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
                            Identifier ("IsUnicodeControl"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("code"),
                                    },
                                    Whitespace,
                                    GreaterThanOrEqualOperator,
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
                                        Identifier ("code"),
                                    },
                                    Whitespace,
                                    LessThanOrEqualOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("31"),
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
    fn ascw_normalization_check() {
        let source = r"
Function CompareNormalized(str1 As String, str2 As String) As Boolean
    Dim code1 As Long, code2 As Long
    code1 = AscW(Mid(str1, 1, 1))
    code2 = AscW(Mid(str2, 1, 1))
    CompareNormalized = (code1 = code2)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CompareNormalized"),
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
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("code1"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Comma,
                        Whitespace,
                        Identifier ("code2"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("code1"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        MidKeyword,
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("str1"),
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
                                                    IntegerLiteral ("1"),
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
                            Identifier ("code2"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        MidKeyword,
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("str2"),
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
                                                    IntegerLiteral ("1"),
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
                            Identifier ("CompareNormalized"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("code1"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("code2"),
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
    fn ascw_html_entity() {
        let source = r#"
Function UnicodeToHTMLEntity(char As String) As String
    Dim code As Long
    code = AscW(char)
    UnicodeToHTMLEntity = "&#" & code & ";"
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("UnicodeToHTMLEntity"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("char"),
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
                        Identifier ("code"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("code"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("char"),
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
                            Identifier ("UnicodeToHTMLEntity"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                StringLiteralExpression {
                                    StringLiteral ("\"&#\""),
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("code"),
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\";\""),
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
    fn ascw_text_analyzer() {
        let source = r"
Function AnalyzeText(text As String) As String
    Dim code As Long
    code = AscW(Mid(text, 1, 1))
    If code >= 0 And code <= 591 Then
        latinCount = latinCount + 1
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
                Identifier ("AnalyzeText"),
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
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("code"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("code"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        MidKeyword,
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
                                                    IntegerLiteral ("1"),
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
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("code"),
                                },
                                Whitespace,
                                GreaterThanOrEqualOperator,
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
                                    Identifier ("code"),
                                },
                                Whitespace,
                                LessThanOrEqualOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("591"),
                                },
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("latinCount"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("latinCount"),
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
    fn ascw_category_validator() {
        let source = r"
Function ValidateCategory(text As String, category As String) As Boolean
    Dim code As Long
    code = AscW(Mid(text, 1, 1))
    If code >= 48 And code <= 57 Then
        isDigit = True
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
                Identifier ("ValidateCategory"),
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
                Identifier ("category"),
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
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("code"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("code"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        MidKeyword,
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
                                                    IntegerLiteral ("1"),
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
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("code"),
                                },
                                Whitespace,
                                GreaterThanOrEqualOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("48"),
                                },
                            },
                            Whitespace,
                            AndKeyword,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("code"),
                                },
                                Whitespace,
                                LessThanOrEqualOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("57"),
                                },
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("isDigit"),
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
    fn ascw_safe_wrapper() {
        let source = r"
Function SafeAscW(text As String) As Long
    If Len(text) = 0 Then
        SafeAscW = -1
        Exit Function
    End If
    SafeAscW = AscW(text)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("SafeAscW"),
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
                            EqualityOperator,
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
                                    Identifier ("SafeAscW"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                UnaryExpression {
                                    SubtractionOperator,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                                Newline,
                            },
                            ExitStatement {
                                Whitespace,
                                ExitKeyword,
                                Whitespace,
                                FunctionKeyword,
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
                            Identifier ("SafeAscW"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscW"),
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
    fn ascw_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To Len(text)
        code = AscW(Mid(text, i, 1))
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
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("code"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("AscW"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            CallExpression {
                                                MidKeyword,
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
                                                            Identifier ("i"),
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
    fn ascw_in_conditional() {
        let source = r"
Sub Test()
    If AscW(char) >= 65 And AscW(char) <= 90 Then
        isUpper = True
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
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("AscW"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("char"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                GreaterThanOrEqualOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("65"),
                                },
                            },
                            Whitespace,
                            AndKeyword,
                            Whitespace,
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("AscW"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("char"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                LessThanOrEqualOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("90"),
                                },
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("isUpper"),
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
                SubKeyword,
                Newline,
            },
        ]);
    }
}
