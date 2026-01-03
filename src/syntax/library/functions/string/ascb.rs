//! # `AscB` Function
//!
//! Returns an `Integer` representing the byte value (ANSI code) of the first byte in a string.
//! The "B" suffix indicates this is the byte (ANSI) version of the `Asc` function.
//!
//! ## Syntax
//!
//! ```vb
//! AscB(string)
//! ```
//!
//! ## Parameters
//!
//! - **string**: Required. Any valid string expression. If the string contains no characters,
//!   a runtime error occurs (Error 5: Invalid procedure call or argument).
//!
//! ## Returns
//!
//! Returns an `Integer` (0-255) representing the byte value of the first byte in the string.
//!
//! ## Remarks
//!
//! - `AscB` returns the ANSI byte value of the first byte in a string, not the character code.
//! - The B suffix stands for "Byte", distinguishing it from the Unicode `AscW` function.
//! - For single-byte character sets (ANSI), `AscB` and `Asc` return the same value.
//! - For multi-byte character sets (like DBCS), `AscB` returns only the first byte of a multi-byte character.
//! - The return value is always in the range 0-255.
//! - If the string is empty (`""`), a runtime error occurs (Error 5).
//! - `AscB` is useful for low-level byte operations and working with binary data.
//! - The inverse function is `ChrB`, which converts a byte value back to a character.
//! - For Unicode code points, use `AscW` instead of `AscB`.
//!
//! ## Typical Uses
//!
//! 1. **Byte-level text analysis** - Examine individual bytes in ANSI strings
//! 2. **Binary data processing** - Extract byte values from binary strings
//! 3. **File format parsing** - Read byte values from file headers or data structures
//! 4. **Legacy protocol support** - Work with protocols that use ANSI byte values
//! 5. **Character encoding detection** - Analyze byte patterns in text
//! 6. **Checksum calculations** - Use byte values for checksums or hash calculations
//! 7. **Low-level string comparison** - Compare strings at the byte level
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Simple byte value
//! Dim byteVal As Integer
//! byteVal = AscB("A")  ' Returns 65
//! ```
//!
//! ```vb
//! ' Example 2: Extended ANSI character
//! Dim code As Integer
//! code = AscB("é")  ' Returns 233 (in Windows-1252 code page)
//! ```
//!
//! ```vb
//! ' Example 3: First byte of multi-byte character
//! ' In DBCS (Double Byte Character Set) systems
//! Dim firstByte As Integer
//! firstByte = AscB("中")  ' Returns first byte only (varies by encoding)
//! ```
//!
//! ```vb
//! ' Example 4: Control character
//! Dim tabByte As Integer
//! tabByte = AscB(vbTab)  ' Returns 9
//! ```
//!
//! ## Common Patterns
//!
//! ### Validate ASCII Range
//! ```vb
//! Function IsASCII(char As String) As Boolean
//!     If Len(char) = 0 Then Exit Function
//!     IsASCII = (AscB(char) < 128)
//! End Function
//! ```
//!
//! ### Check for Control Characters
//! ```vb
//! Function IsControlChar(char As String) As Boolean
//!     If Len(char) = 0 Then Exit Function
//!     Dim byteVal As Integer
//!     byteVal = AscB(char)
//!     IsControlChar = (byteVal < 32 Or byteVal = 127)
//! End Function
//! ```
//!
//! ### Compare Byte Values
//! ```vb
//! Function CompareBytes(str1 As String, str2 As String) As Integer
//!     If Len(str1) = 0 Or Len(str2) = 0 Then Exit Function
//!     CompareBytes = AscB(str1) - AscB(str2)
//! End Function
//! ```
//!
//! ### Extract Byte Array
//! ```vb
//! Function GetByteArray(text As String) As Variant
//!     Dim bytes() As Integer
//!     Dim i As Long
//!     
//!     If Len(text) = 0 Then Exit Function
//!     
//!     ReDim bytes(1 To Len(text))
//!     For i = 1 To Len(text)
//!         bytes(i) = AscB(Mid(text, i, 1))
//!     Next i
//!     
//!     GetByteArray = bytes
//! End Function
//! ```
//!
//! ### Calculate Simple Checksum
//! ```vb
//! Function SimpleChecksum(text As String) As Long
//!     Dim i As Long
//!     Dim checksum As Long
//!     
//!     For i = 1 To Len(text)
//!         checksum = checksum + AscB(Mid(text, i, 1))
//!     Next i
//!     
//!     SimpleChecksum = checksum Mod 256
//! End Function
//! ```
//!
//! ### Detect Line Endings
//! ```vb
//! Function DetectLineEnding(text As String) As String
//!     Dim i As Long
//!     Dim byteVal As Integer
//!     
//!     For i = 1 To Len(text)
//!         byteVal = AscB(Mid(text, i, 1))
//!         If byteVal = 13 Then  ' CR
//!             If i < Len(text) And AscB(Mid(text, i + 1, 1)) = 10 Then
//!                 DetectLineEnding = "CRLF"
//!             Else
//!                 DetectLineEnding = "CR"
//!             End If
//!             Exit Function
//!         ElseIf byteVal = 10 Then  ' LF
//!             DetectLineEnding = "LF"
//!             Exit Function
//!         End If
//!     Next i
//! End Function
//! ```
//!
//! ### Hex Dump Generator
//! ```vb
//! Function ByteToHex(char As String) As String
//!     If Len(char) = 0 Then Exit Function
//!     Dim byteVal As Integer
//!     byteVal = AscB(char)
//!     ByteToHex = Right("0" & Hex(byteVal), 2)
//! End Function
//! ```
//!
//! ### Case-Insensitive Byte Compare
//! ```vb
//! Function ByteEqualsIgnoreCase(char1 As String, char2 As String) As Boolean
//!     If Len(char1) = 0 Or Len(char2) = 0 Then Exit Function
//!     
//!     Dim byte1 As Integer, byte2 As Integer
//!     byte1 = AscB(char1)
//!     byte2 = AscB(char2)
//!     
//!     ' Convert uppercase to lowercase (65-90 to 97-122)
//!     If byte1 >= 65 And byte1 <= 90 Then byte1 = byte1 + 32
//!     If byte2 >= 65 And byte2 <= 90 Then byte2 = byte2 + 32
//!     
//!     ByteEqualsIgnoreCase = (byte1 = byte2)
//! End Function
//! ```
//!
//! ### Filter Printable Characters
//! ```vb
//! Function FilterPrintable(text As String) As String
//!     Dim result As String
//!     Dim i As Long
//!     Dim byteVal As Integer
//!     
//!     For i = 1 To Len(text)
//!         byteVal = AscB(Mid(text, i, 1))
//!         If byteVal >= 32 And byteVal <= 126 Then
//!             result = result & Mid(text, i, 1)
//!         End If
//!     Next i
//!     
//!     FilterPrintable = result
//! End Function
//! ```
//!
//! ### Encode for URL
//! ```vb
//! Function NeedsURLEncoding(char As String) As Boolean
//!     If Len(char) = 0 Then Exit Function
//!     
//!     Dim byteVal As Integer
//!     byteVal = AscB(char)
//!     
//!     ' Check if character needs encoding
//!     If (byteVal >= 48 And byteVal <= 57) Or _
//!        (byteVal >= 65 And byteVal <= 90) Or _
//!        (byteVal >= 97 And byteVal <= 122) Then
//!         NeedsURLEncoding = False
//!     Else
//!         NeedsURLEncoding = True
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Examples
//!
//! ### Binary Data Parser
//! ```vb
//! Function ParseBinaryHeader(data As String) As Variant
//!     ' Parse a binary file header (example: BMP format)
//!     Dim header As Variant
//!     ReDim header(1 To 4)
//!     
//!     If Len(data) < 4 Then Exit Function
//!     
//!     ' Read signature bytes
//!     header(1) = AscB(Mid(data, 1, 1))  ' 'B' = 66
//!     header(2) = AscB(Mid(data, 2, 1))  ' 'M' = 77
//!     
//!     ' Read size bytes (little-endian)
//!     header(3) = AscB(Mid(data, 3, 1))
//!     header(4) = AscB(Mid(data, 4, 1))
//!     
//!     ParseBinaryHeader = header
//! End Function
//! ```
//!
//! ### XOR Encryption/Decryption
//! ```vb
//! Function XOREncrypt(text As String, key As String) As String
//!     Dim result As String
//!     Dim i As Long, keyPos As Long
//!     Dim textByte As Integer, keyByte As Integer
//!     
//!     If Len(text) = 0 Or Len(key) = 0 Then Exit Function
//!     
//!     keyPos = 1
//!     For i = 1 To Len(text)
//!         textByte = AscB(Mid(text, i, 1))
//!         keyByte = AscB(Mid(key, keyPos, 1))
//!         
//!         result = result & ChrB(textByte Xor keyByte)
//!         
//!         keyPos = keyPos + 1
//!         If keyPos > Len(key) Then keyPos = 1
//!     Next i
//!     
//!     XOREncrypt = result
//! End Function
//! ```
//!
//! ### CSV Field Parser with Byte Analysis
//! ```vb
//! Function ParseCSVField(field As String) As String
//!     Dim result As String
//!     Dim i As Long
//!     Dim byteVal As Integer
//!     Dim inQuotes As Boolean
//!     
//!     For i = 1 To Len(field)
//!         byteVal = AscB(Mid(field, i, 1))
//!         
//!         Select Case byteVal
//!             Case 34  ' Double quote
//!                 inQuotes = Not inQuotes
//!             Case 44  ' Comma
//!                 If Not inQuotes Then Exit Function
//!                 result = result & Chr(byteVal)
//!             Case Else
//!                 result = result & Chr(byteVal)
//!         End Select
//!     Next i
//!     
//!     ParseCSVField = result
//! End Function
//! ```
//!
//! ### Character Set Validator
//! ```vb
//! Function ValidateCharacterSet(text As String, validSet As String) As Boolean
//!     Dim i As Long, j As Long
//!     Dim textByte As Integer
//!     Dim found As Boolean
//!     
//!     For i = 1 To Len(text)
//!         textByte = AscB(Mid(text, i, 1))
//!         found = False
//!         
//!         For j = 1 To Len(validSet)
//!             If textByte = AscB(Mid(validSet, j, 1)) Then
//!                 found = True
//!                 Exit For
//!             End If
//!         Next j
//!         
//!         If Not found Then
//!             ValidateCharacterSet = False
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     ValidateCharacterSet = True
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeAscB(text As String) As Integer
//!     On Error GoTo ErrorHandler
//!     
//!     If Len(text) = 0 Then
//!         SafeAscB = -1  ' Error indicator
//!         Exit Function
//!     End If
//!     
//!     SafeAscB = AscB(text)
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeAscB = -1
//! End Function
//! ```
//!
//! ## Performance Notes
//!
//! - `AscB` is a very fast operation with minimal overhead
//! - When processing long strings byte-by-byte, consider using `Mid` function efficiently
//! - For repeated byte extraction, the performance is generally good
//! - Avoid calling `AscB` in tight loops if the value can be cached
//! - `AscB` is faster than string comparison for byte-level operations
//!
//! ## Best Practices
//!
//! 1. **Validate input** - Always check for empty strings before calling `AscB`
//! 2. **Use for byte operations** - Prefer `AscB` over `Asc` when working with binary data
//! 3. **Handle errors** - Wrap `AscB` calls in error handlers when processing untrusted input
//! 4. **Document byte values** - Use constants or comments to explain non-obvious byte values
//! 5. **Consider encoding** - Be aware of system code page when working with extended ANSI
//! 6. **Use with `ChrB`** - Pair with `ChrB` for byte-to-character conversions
//! 7. **Test edge cases** - Verify behavior with empty strings, control characters, and extended ANSI
//!
//! ## Comparison with Related Functions
//!
//! | Function | Returns | Character Set | Use Case |
//! |----------|---------|---------------|----------|
//! | `Asc` | Integer (0-255 or Unicode) | System default | General character codes |
//! | `AscB` | Integer (0-255) | ANSI byte value | Byte-level operations |
//! | `AscW` | Integer (0-65535) | Unicode code point | International text |
//! | `ChrB` | String (ANSI) | ANSI (inverse) | Convert byte to character |
//!
//! ## Common Byte Values Reference
//!
//! Some commonly used byte values with `AscB`:
//!
//! - **0**: Null character (NUL)
//! - **9**: Tab (HT)
//! - **10**: Line feed (LF)
//! - **13**: Carriage return (CR)
//! - **32**: Space
//! - **48-57**: Digits '0'-'9'
//! - **65-90**: Uppercase letters 'A'-'Z'
//! - **97-122**: Lowercase letters 'a'-'z'
//! - **127**: Delete (DEL)
//! - **128-255**: Extended ANSI (varies by code page)
//!
//! ## Platform Notes
//!
//! - On Windows systems, `AscB` uses the system's ANSI code page (e.g., Windows-1252)
//! - Different code pages may interpret bytes 128-255 differently
//! - For portable code, stick to ASCII range (0-127) when possible
//! - On older systems (Windows 95/98/ME), ANSI encoding is the native string format
//! - On modern Windows (NT-based), strings are Unicode internally but `AscB` still returns ANSI bytes
//!
//! ## Limitations
//!
//! - Returns only the first byte, not the full character in multi-byte encodings
//! - Cannot handle Unicode characters outside the ANSI range (0-255) properly
//! - Runtime error occurs with empty strings
//! - Code page dependent for extended ANSI characters (128-255)
//! - Not suitable for Unicode text processing (use `AscW` instead)

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn ascb_simple_character() {
        let source = r#"
Sub Test()
    byteVal = AscB("A")
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
                            Identifier ("byteVal"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
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
    fn ascb_extended_ansi() {
        let source = r"
Sub Test()
    code = AscB(extendedChar)
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
                            Identifier ("code"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("extendedChar"),
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
    fn ascb_control_character() {
        let source = r"
Sub Test()
    tabByte = AscB(vbTab)
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
                            Identifier ("tabByte"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("vbTab"),
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
    fn ascb_is_ascii_function() {
        let source = r"
Function IsASCII(char As String) As Boolean
    If Len(char) = 0 Then Exit Function
    IsASCII = (AscB(char) < 128)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("IsASCII"),
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
                                            Identifier ("char"),
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
                        Whitespace,
                        ExitStatement {
                            ExitKeyword,
                            Whitespace,
                            FunctionKeyword,
                            Newline,
                        },
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("IsASCII"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            ParenthesizedExpression {
                                LeftParenthesis,
                                BinaryExpression {
                                    CallExpression {
                                        Identifier ("AscB"),
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
                                    LessThanOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("128"),
                                    },
                                },
                                RightParenthesis,
                            },
                            Newline,
                        },
                        EndKeyword,
                        Whitespace,
                        FunctionKeyword,
                        Newline,
                    },
                },
            },
        ]);
    }

    #[test]
    fn ascb_is_control_char() {
        let source = r"
Function IsControlChar(char As String) As Boolean
    Dim byteVal As Integer
    byteVal = AscB(char)
    IsControlChar = (byteVal < 32 Or byteVal = 127)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("IsControlChar"),
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
                        Identifier ("byteVal"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("byteVal"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
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
                            Identifier ("IsControlChar"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("byteVal"),
                                    },
                                    Whitespace,
                                    LessThanOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("32"),
                                    },
                                },
                                Whitespace,
                                OrKeyword,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("byteVal"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("127"),
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
    fn ascb_compare_bytes() {
        let source = r"
Function CompareBytes(str1 As String, str2 As String) As Integer
    CompareBytes = AscB(str1) - AscB(str2)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CompareBytes"),
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
                IntegerKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("CompareBytes"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("AscB"),
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
                            SubtractionOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("AscB"),
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
    fn ascb_checksum_calculation() {
        let source = r"
Function SimpleChecksum(text As String) As Long
    Dim i As Long
    For i = 1 To Len(text)
        checksum = checksum + AscB(Mid(text, i, 1))
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
                Identifier ("SimpleChecksum"),
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
                                IdentifierExpression {
                                    Identifier ("checksum"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("checksum"),
                                    },
                                    Whitespace,
                                    AdditionOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("AscB"),
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
    fn ascb_detect_line_ending() {
        let source = r#"
Function DetectLineEnding(text As String) As String
    Dim byteVal As Integer
    byteVal = AscB(Mid(text, 1, 1))
    If byteVal = 13 Then
        DetectLineEnding = "CR"
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
                Identifier ("DetectLineEnding"),
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
                        Identifier ("byteVal"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("byteVal"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
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
                            IdentifierExpression {
                                Identifier ("byteVal"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("13"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("DetectLineEnding"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                StringLiteralExpression {
                                    StringLiteral ("\"CR\""),
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
    fn ascb_byte_to_hex() {
        let source = r#"
Function ByteToHex(char As String) As String
    Dim byteVal As Integer
    byteVal = AscB(char)
    ByteToHex = Right("0" & Hex(byteVal), 2)
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ByteToHex"),
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
                        Identifier ("byteVal"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("byteVal"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
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
                            Identifier ("ByteToHex"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Right"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        StringLiteralExpression {
                                            StringLiteral ("\"0\""),
                                        },
                                        Whitespace,
                                        Ampersand,
                                        Whitespace,
                                        CallExpression {
                                            Identifier ("Hex"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    IdentifierExpression {
                                                        Identifier ("byteVal"),
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
    fn ascb_case_insensitive_compare() {
        let source = r"
Function ByteEqualsIgnoreCase(char1 As String, char2 As String) As Boolean
    Dim byte1 As Integer, byte2 As Integer
    byte1 = AscB(char1)
    byte2 = AscB(char2)
    ByteEqualsIgnoreCase = (byte1 = byte2)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ByteEqualsIgnoreCase"),
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
                BooleanKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("byte1"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Comma,
                        Whitespace,
                        Identifier ("byte2"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("byte1"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
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
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("byte2"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
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
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("ByteEqualsIgnoreCase"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("byte1"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("byte2"),
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
    fn ascb_filter_printable() {
        let source = r"
Function FilterPrintable(text As String) As String
    Dim byteVal As Integer
    byteVal = AscB(Mid(text, 1, 1))
    If byteVal >= 32 And byteVal <= 126 Then
        result = result & Mid(text, 1, 1)
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
                Identifier ("FilterPrintable"),
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
                        Identifier ("byteVal"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("byteVal"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
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
                                    Identifier ("byteVal"),
                                },
                                Whitespace,
                                GreaterThanOrEqualOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("32"),
                                },
                            },
                            Whitespace,
                            AndKeyword,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("byteVal"),
                                },
                                Whitespace,
                                LessThanOrEqualOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("126"),
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
    fn ascb_url_encoding_check() {
        let source = r"
Function NeedsURLEncoding(char As String) As Boolean
    Dim byteVal As Integer
    byteVal = AscB(char)
    If byteVal >= 48 And byteVal <= 57 Then
        NeedsURLEncoding = False
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
                Identifier ("NeedsURLEncoding"),
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
                        Identifier ("byteVal"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("byteVal"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
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
                                    Identifier ("byteVal"),
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
                                    Identifier ("byteVal"),
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
                                    Identifier ("NeedsURLEncoding"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BooleanLiteralExpression {
                                    FalseKeyword,
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
    fn ascb_binary_parser() {
        let source = r"
Function ParseBinaryHeader(data As String) As Variant
    Dim header As Variant
    header(1) = AscB(Mid(data, 1, 1))
    header(2) = AscB(Mid(data, 2, 1))
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ParseBinaryHeader"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("data"),
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
                        Identifier ("header"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        VariantKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        CallExpression {
                            Identifier ("header"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        MidKeyword,
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
                        CallExpression {
                            Identifier ("header"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("2"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        MidKeyword,
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
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("2"),
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
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn ascb_xor_encrypt() {
        let source = r"
Function XOREncrypt(text As String, key As String) As String
    Dim textByte As Integer, keyByte As Integer
    textByte = AscB(Mid(text, 1, 1))
    keyByte = AscB(Mid(key, 1, 1))
    result = ChrB(textByte Xor keyByte)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("XOREncrypt"),
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
                Identifier ("key"),
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
                        Identifier ("textByte"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Comma,
                        Whitespace,
                        Identifier ("keyByte"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("textByte"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("keyByte"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        MidKeyword,
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("key"),
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
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("ChrB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("textByte"),
                                        },
                                        Whitespace,
                                        XorKeyword,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("keyByte"),
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
    fn ascb_csv_parser() {
        let source = r"
Function ParseCSVField(field As String) As String
    Dim byteVal As Integer
    byteVal = AscB(Mid(field, 1, 1))
    If byteVal = 34 Then
        inQuotes = True
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
                Identifier ("ParseCSVField"),
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
                        Identifier ("byteVal"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("byteVal"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        MidKeyword,
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
                            IdentifierExpression {
                                Identifier ("byteVal"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("34"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("inQuotes"),
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
    fn ascb_charset_validator() {
        let source = r"
Function ValidateCharacterSet(text As String, validSet As String) As Boolean
    Dim textByte As Integer
    textByte = AscB(Mid(text, 1, 1))
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ValidateCharacterSet"),
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
                Identifier ("validSet"),
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
                        Identifier ("textByte"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("textByte"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
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
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn ascb_safe_wrapper() {
        let source = r"
Function SafeAscB(text As String) As Integer
    If Len(text) = 0 Then
        SafeAscB = -1
        Exit Function
    End If
    SafeAscB = AscB(text)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("SafeAscB"),
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
                IntegerKeyword,
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
                                    Identifier ("SafeAscB"),
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
                            Identifier ("SafeAscB"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
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
    fn ascb_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To Len(text)
        byteVal = AscB(Mid(text, i, 1))
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
                                    Identifier ("byteVal"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("AscB"),
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
    fn ascb_with_mid_function() {
        let source = r"
Sub Test()
    firstByte = AscB(Mid(myString, 1, 1))
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
                            Identifier ("firstByte"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AscB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        MidKeyword,
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("myString"),
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
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn ascb_in_conditional() {
        let source = r"
Sub Test()
    If AscB(char) >= 65 And AscB(char) <= 90 Then
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
                                    Identifier ("AscB"),
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
                                    Identifier ("AscB"),
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
