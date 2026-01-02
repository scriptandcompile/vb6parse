//! # `Hex$` Function
//!
//! The `Hex$` function in Visual Basic 6 returns a string representing the hexadecimal (base-16)
//! value of a number. The dollar sign (`$`) suffix indicates that this function always returns a
//! `String` type, never a `Variant`.
//!
//! ## Syntax
//!
//! ```vb6
//! Hex$(number)
//! ```
//!
//! ## Parameters
//!
//! - `number` - Required. Any valid numeric expression or string expression. If `number` is not a
//!   whole number, it is rounded to the nearest whole number before being evaluated.
//!
//! ## Return Value
//!
//! Returns a `String` representing the hexadecimal value of `number`. The string contains only
//! hexadecimal digits (0-9, A-F) without any prefix (no "0x" or "&H").
//!
//! ## Behavior and Characteristics
//!
//! ### Number Range and Representation
//!
//! - Positive numbers: Returns hexadecimal representation without leading zeros
//! - Negative numbers: Returns two's complement representation
//! - `Byte` values: Up to 2 hex digits (00-FF)
//! - `Integer` values: Up to 4 hex digits (0000-FFFF)
//! - `Long` values: Up to 8 hex digits (00000000-FFFFFFFF)
//! - Zero: Returns "0" (single character)
//! - If `number` contains `Null`, returns `Null`
//!
//! ### Type Differences: `Hex$` vs `Hex`
//!
//! - `Hex$`: Always returns `String` type (never `Variant`)
//! - `Hex`: Returns `Variant` (can propagate `Null` values)
//! - Use `Hex$` when you need guaranteed `String` return type
//! - Use `Hex` when working with potentially `Null` values
//!
//! ### Formatting Characteristics
//!
//! - No "0x" or "&H" prefix in output
//! - Uses uppercase letters (A-F, not a-f)
//! - No leading zeros for positive numbers (except zero itself)
//! - Negative numbers use two's complement representation
//! - Maximum 8 characters for `Long` type
//!
//! ## Common Usage Patterns
//!
//! ### 1. Convert Numbers to Hex Strings
//!
//! ```vb6
//! Function NumberToHex(value As Long) As String
//!     NumberToHex = Hex$(value)
//! End Function
//!
//! Debug.Print NumberToHex(255)      ' "FF"
//! Debug.Print NumberToHex(4096)     ' "1000"
//! Debug.Print NumberToHex(65535)    ' "FFFF"
//! ```
//!
//! ### 2. Display RGB Color Values
//!
//! ```vb6
//! Function ColorToHex(colorValue As Long) As String
//!     Dim hexStr As String
//!     hexStr = Hex$(colorValue)
//!     ' Pad to 6 characters for web colors
//!     ColorToHex = String$(6 - Len(hexStr), "0") & hexStr
//! End Function
//!
//! Dim webColor As String
//! webColor = "#" & ColorToHex(RGB(255, 128, 64))
//! ```
//!
//! ### 3. Debug Memory Addresses
//!
//! ```vb6
//! Function FormatAddress(address As Long) As String
//!     Dim hexAddr As String
//!     hexAddr = Hex$(address)
//!     ' Pad to 8 characters
//!     FormatAddress = "0x" & String$(8 - Len(hexAddr), "0") & hexAddr
//! End Function
//! ```
//!
//! ### 4. Generate Unique Identifiers
//!
//! ```vb6
//! Function GenerateHexID() As String
//!     Randomize
//!     Dim part1 As Long, part2 As Long
//!     part1 = Int(Rnd * &H7FFFFFFF)
//!     part2 = Int(Rnd * &H7FFFFFFF)
//!     GenerateHexID = Hex$(part1) & Hex$(part2)
//! End Function
//! ```
//!
//! ### 5. Format Byte Arrays as Hex Strings
//!
//! ```vb6
//! Function BytesToHex(bytes() As Byte) As String
//!     Dim result As String
//!     Dim i As Integer
//!     Dim hexByte As String
//!     
//!     For i = LBound(bytes) To UBound(bytes)
//!         hexByte = Hex$(bytes(i))
//!         If Len(hexByte) = 1 Then hexByte = "0" & hexByte
//!         result = result & hexByte
//!     Next i
//!     
//!     BytesToHex = result
//! End Function
//! ```
//!
//! ### 6. Log Error Codes in Hex
//!
//! ```vb6
//! Sub LogError(errNum As Long, errDesc As String)
//!     Dim logFile As Integer
//!     logFile = FreeFile
//!     Open "errors.log" For Append As #logFile
//!     Print #logFile, "Error 0x" & Hex$(errNum) & ": " & errDesc
//!     Close #logFile
//! End Sub
//! ```
//!
//! ### 7. Convert Character Codes
//!
//! ```vb6
//! Function CharToHex(ch As String) As String
//!     If Len(ch) > 0 Then
//!         CharToHex = Hex$(Asc(ch))
//!     Else
//!         CharToHex = ""
//!     End If
//! End Function
//!
//! Debug.Print CharToHex("A")  ' "41"
//! Debug.Print CharToHex("Z")  ' "5A"
//! ```
//!
//! ### 8. Create Hexadecimal Dump
//!
//! ```vb6
//! Function HexDump(data As String, Optional bytesPerLine As Integer = 16) As String
//!     Dim result As String
//!     Dim i As Long
//!     Dim hexVal As String
//!     
//!     For i = 1 To Len(data)
//!         hexVal = Hex$(Asc(Mid$(data, i, 1)))
//!         If Len(hexVal) = 1 Then hexVal = "0" & hexVal
//!         result = result & hexVal & " "
//!         
//!         If (i Mod bytesPerLine) = 0 Then
//!             result = result & vbCrLf
//!         End If
//!     Next i
//!     
//!     HexDump = result
//! End Function
//! ```
//!
//! ### 9. Parse and Format Checksums
//!
//! ```vb6
//! Function FormatChecksum(checksum As Long) As String
//!     Dim hexStr As String
//!     hexStr = Hex$(checksum)
//!     ' Pad to 8 characters
//!     FormatChecksum = String$(8 - Len(hexStr), "0") & hexStr
//! End Function
//!
//! Dim crc32 As Long
//! crc32 = CalculateCRC32(fileData)
//! Debug.Print "CRC32: " & FormatChecksum(crc32)
//! ```
//!
//! ### 10. Network Protocol Debugging
//!
//! ```vb6
//! Function FormatPacketHeader(packetType As Byte, packetLen As Integer) As String
//!     Dim typeHex As String, lenHex As String
//!     
//!     typeHex = Hex$(packetType)
//!     If Len(typeHex) = 1 Then typeHex = "0" & typeHex
//!     
//!     lenHex = Hex$(packetLen)
//!     While Len(lenHex) < 4
//!         lenHex = "0" & lenHex
//!     Wend
//!     
//!     FormatPacketHeader = "Type: 0x" & typeHex & " Len: 0x" & lenHex
//! End Function
//! ```
//!
//! ## Related Functions
//!
//! - `Hex()` - Returns a `Variant` containing the hexadecimal value (can handle `Null`)
//! - `Oct$()` - Returns the octal (base-8) representation of a number
//! - `Str$()` - Converts a number to its decimal string representation
//! - `Val()` - Converts a string to a numeric value
//! - `CLng()` - Converts an expression to a `Long` integer
//! - `Asc()` - Returns the character code of the first character in a string
//! - `Chr$()` - Returns the character associated with a character code
//! - `Format$()` - Formats expressions with more control over output
//!
//! ## Best Practices
//!
//! ### Padding Hex Values
//!
//! ```vb6
//! ' Pad to specific width for consistent formatting
//! Function PadHex(value As Long, width As Integer) As String
//!     Dim hexStr As String
//!     hexStr = Hex$(value)
//!     
//!     If Len(hexStr) < width Then
//!         PadHex = String$(width - Len(hexStr), "0") & hexStr
//!     Else
//!         PadHex = hexStr
//!     End If
//! End Function
//!
//! Debug.Print PadHex(255, 4)   ' "00FF"
//! Debug.Print PadHex(4096, 8)  ' "00001000"
//! ```
//!
//! ### Adding Hex Prefix
//!
//! ```vb6
//! Function HexWithPrefix(value As Long) As String
//!     HexWithPrefix = "&H" & Hex$(value)  ' VB6 style
//!     ' Or: HexWithPrefix = "0x" & Hex$(value)  ' C style
//! End Function
//! ```
//!
//! ### Converting Back from Hex String
//!
//! ```vb6
//! Function HexToLong(hexStr As String) As Long
//!     ' Remove any prefix
//!     If Left$(hexStr, 2) = "&H" Or Left$(hexStr, 2) = "0x" Then
//!         hexStr = Mid$(hexStr, 3)
//!     End If
//!     
//!     ' Convert using Val with &H prefix
//!     HexToLong = Val("&H" & hexStr)
//! End Function
//! ```
//!
//! ### Handling Byte Order (Endianness)
//!
//! ```vb6
//! Function LongToHexBytes(value As Long) As String
//!     Dim b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte
//!     
//!     b1 = value And &HFF
//!     b2 = (value \ &H100) And &HFF
//!     b3 = (value \ &H10000) And &HFF
//!     b4 = (value \ &H1000000) And &HFF
//!     
//!     ' Little-endian format
//!     LongToHexBytes = Right$("0" & Hex$(b1), 2) & " " & _
//!                      Right$("0" & Hex$(b2), 2) & " " & _
//!                      Right$("0" & Hex$(b3), 2) & " " & _
//!                      Right$("0" & Hex$(b4), 2)
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - `Hex$` is very fast for converting numbers to hexadecimal strings
//! - No significant performance difference between `Hex` and `Hex$` for non-Null values
//! - String concatenation in loops can be slow; consider building arrays and using `Join`
//! - For large byte arrays, consider buffering output
//!
//! ```vb6
//! ' Efficient for large arrays
//! Function BytesToHexEfficient(bytes() As Byte) As String
//!     Dim chunks() As String
//!     ReDim chunks(UBound(bytes) - LBound(bytes))
//!     
//!     Dim i As Long, idx As Long
//!     For i = LBound(bytes) To UBound(bytes)
//!         chunks(idx) = Right$("0" & Hex$(bytes(i)), 2)
//!         idx = idx + 1
//!     Next i
//!     
//!     BytesToHexEfficient = Join(chunks, "")
//! End Function
//! ```
//!
//! ## Common Pitfalls
//!
//! ### 1. No Automatic Padding
//!
//! ```vb6
//! ' Hex$ does NOT add leading zeros
//! Debug.Print Hex$(15)    ' "F" (not "0F")
//! Debug.Print Hex$(255)   ' "FF" (correct)
//! Debug.Print Hex$(16)    ' "10" (not "0010")
//!
//! ' Must pad manually for consistent width
//! Function PadHex(val As Integer) As String
//!     Dim h As String
//!     h = Hex$(val)
//!     PadHex = String$(4 - Len(h), "0") & h
//! End Function
//! ```
//!
//! ### 2. Negative Numbers Use Two's Complement
//!
//! ```vb6
//! ' Negative numbers are represented in two's complement
//! Debug.Print Hex$(-1)     ' "FFFFFFFF" (Long)
//! Debug.Print Hex$(-256)   ' "FFFFFF00"
//!
//! ' For signed interpretation, check range
//! Function SignedHex(value As Long) As String
//!     If value < 0 Then
//!         SignedHex = "-&H" & Hex$(Abs(value))
//!     Else
//!         SignedHex = "&H" & Hex$(value)
//!     End If
//! End Function
//! ```
//!
//! ### 3. No Prefix in Output
//!
//! ```vb6
//! ' Hex$ does NOT include "&H" or "0x" prefix
//! Dim hexValue As String
//! hexValue = Hex$(255)      ' "FF" (not "&HFF" or "0xFF")
//!
//! ' Add prefix manually if needed
//! hexValue = "&H" & Hex$(255)  ' "&HFF"
//! hexValue = "0x" & Hex$(255)  ' "0xFF"
//! ```
//!
//! ### 4. Uppercase Output Only
//!
//! ```vb6
//! ' Hex$ always returns uppercase A-F
//! Debug.Print Hex$(255)  ' "FF" (not "ff")
//!
//! ' Convert to lowercase if needed
//! hexValue = LCase$(Hex$(255))  ' "ff"
//! ```
//!
//! ### 5. Rounding of Non-Integer Values
//!
//! ```vb6
//! ' Non-integers are rounded before conversion
//! Debug.Print Hex$(15.3)   ' "F" (15 rounded)
//! Debug.Print Hex$(15.7)   ' "10" (16 rounded)
//! Debug.Print Hex$(15.5)   ' "10" (banker's rounding to even)
//!
//! ' Use Fix or Int if you need specific rounding
//! Debug.Print Hex$(Int(15.7))   ' "F" (truncated to 15)
//! Debug.Print Hex$(Fix(15.7))   ' "F" (truncated to 15)
//! ```
//!
//! ### 6. Type Range Limitations
//!
//! ```vb6
//! ' Different types have different ranges
//! Dim b As Byte
//! Dim i As Integer
//! Dim l As Long
//!
//! b = 255
//! Debug.Print Hex$(b)  ' "FF"
//!
//! i = -1
//! Debug.Print Hex$(i)  ' "FFFF" (16-bit two's complement)
//!
//! l = -1
//! Debug.Print Hex$(l)  ' "FFFFFFFF" (32-bit two's complement)
//! ```
//!
//! ## Practical Examples
//!
//! ### Memory Dump Utility
//!
//! ```vb6
//! Sub DumpMemory(startAddr As Long, length As Integer)
//!     Dim i As Integer
//!     Dim addr As Long
//!     Dim byteVal As Byte
//!     Dim line As String
//!     Dim ascii As String
//!     
//!     For i = 0 To length - 1
//!         If (i Mod 16) = 0 Then
//!             If i > 0 Then
//!                 Debug.Print line & "  " & ascii
//!             End If
//!             addr = startAddr + i
//!             line = Right$("00000000" & Hex$(addr), 8) & ": "
//!             ascii = ""
//!         End If
//!         
//!         ' Get byte value (pseudo-code)
//!         byteVal = GetMemoryByte(startAddr + i)
//!         line = line & Right$("0" & Hex$(byteVal), 2) & " "
//!         
//!         If byteVal >= 32 And byteVal <= 126 Then
//!             ascii = ascii & Chr$(byteVal)
//!         Else
//!             ascii = ascii & "."
//!         End If
//!     Next i
//!     
//!     ' Print last line
//!     If ascii <> "" Then
//!         Debug.Print line & String$(3 * (16 - Len(ascii)), " ") & "  " & ascii
//!     End If
//! End Sub
//! ```
//!
//! ### UUID/GUID Formatter
//!
//! ```vb6
//! Function FormatGUID(data1 As Long, data2 As Integer, data3 As Integer, _
//!                     data4() As Byte) As String
//!     Dim result As String
//!     Dim i As Integer
//!     
//!     result = Right$("00000000" & Hex$(data1), 8) & "-"
//!     result = result & Right$("0000" & Hex$(data2), 4) & "-"
//!     result = result & Right$("0000" & Hex$(data3), 4) & "-"
//!     
//!     For i = 0 To 1
//!         result = result & Right$("0" & Hex$(data4(i)), 2)
//!     Next i
//!     result = result & "-"
//!     
//!     For i = 2 To 7
//!         result = result & Right$("0" & Hex$(data4(i)), 2)
//!     Next i
//!     
//!     FormatGUID = result
//! End Function
//! ```
//!
//! ### Color Manipulation
//!
//! ```vb6
//! Function RGBToWebColor(r As Byte, g As Byte, b As Byte) As String
//!     RGBToWebColor = "#" & _
//!                     Right$("0" & Hex$(r), 2) & _
//!                     Right$("0" & Hex$(g), 2) & _
//!                     Right$("0" & Hex$(b), 2)
//! End Function
//!
//! Function WebColorToRGB(webColor As String, r As Byte, g As Byte, b As Byte)
//!     ' Remove # if present
//!     If Left$(webColor, 1) = "#" Then webColor = Mid$(webColor, 2)
//!     
//!     r = Val("&H" & Mid$(webColor, 1, 2))
//!     g = Val("&H" & Mid$(webColor, 3, 2))
//!     b = Val("&H" & Mid$(webColor, 5, 2))
//! End Function
//! ```
//!
//! ## Limitations
//!
//! - Returns only uppercase hexadecimal letters (A-F), not lowercase
//! - Does not include "&H" or "0x" prefix (must add manually)
//! - Does not pad with leading zeros (must pad manually)
//! - Cannot handle `Null` values (use `Hex` variant function instead)
//! - Limited to 32-bit `Long` integer range (no 64-bit support in VB6)
//! - Negative numbers return two's complement representation
//! - Fractional values are rounded before conversion
//! - No direct support for byte-order conversion (endianness)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn hex_dollar_simple() {
        let source = r"
Sub Main()
    result = Hex$(255)
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
                            Identifier ("Hex$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("255"),
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
    fn hex_dollar_assignment() {
        let source = r"
Sub Main()
    Dim hexValue As String
    hexValue = Hex$(value)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Hex$"));
    }

    #[test]
    fn hex_dollar_variable() {
        let source = r"
Sub Main()
    Dim num As Long
    Dim hexStr As String
    num = 4096
    hexStr = Hex$(num)
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
                        Identifier ("num"),
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
                        Identifier ("hexStr"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("num"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("4096"),
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("hexStr"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Hex$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("num"),
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
    fn hex_dollar_number_to_hex() {
        let source = r"
Function NumberToHex(value As Long) As String
    NumberToHex = Hex$(value)
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Hex$"));
    }

    #[test]
    fn hex_dollar_color_conversion() {
        let source = r"
Function ColorToHex(colorValue As Long) As String
    Dim hexStr As String
    hexStr = Hex$(colorValue)
    ColorToHex = hexStr
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ColorToHex"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("colorValue"),
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("hexStr"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("hexStr"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Hex$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("colorValue"),
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
                            Identifier ("ColorToHex"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("hexStr"),
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
    fn hex_dollar_with_padding() {
        let source = r#"
Function PadHex(value As Long, width As Integer) As String
    Dim hexStr As String
    hexStr = Hex$(value)
    PadHex = String$(width - Len(hexStr), "0") & hexStr
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Hex$"));
    }

    #[test]
    fn hex_dollar_debug_address() {
        let source = r#"
Function FormatAddress(address As Long) As String
    FormatAddress = "0x" & Hex$(address)
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("FormatAddress"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("address"),
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("FormatAddress"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"0x\""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Hex$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("address"),
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
    fn hex_dollar_bytes_array() {
        let source = r"
Function BytesToHex(bytes() As Byte) As String
    Dim result As String
    Dim i As Integer
    For i = LBound(bytes) To UBound(bytes)
        result = result & Hex$(bytes(i))
    Next i
    BytesToHex = result
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Hex$"));
    }

    #[test]
    fn hex_dollar_error_logging() {
        let source = r#"
Sub LogError(errNum As Long)
    Debug.Print "Error 0x" & Hex$(errNum)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("LogError"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("errNum"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    LongKeyword,
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
                        StringLiteral ("\"Error 0x\""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Hex$"),
                        LeftParenthesis,
                        Identifier ("errNum"),
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
    fn hex_dollar_char_code() {
        let source = r"
Function CharToHex(ch As String) As String
    If Len(ch) > 0 Then
        CharToHex = Hex$(Asc(ch))
    End If
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Hex$"));
    }

    #[test]
    fn hex_dollar_in_loop() {
        let source = r"
Sub DumpData()
    Dim i As Integer
    For i = 0 To 255
        Debug.Print Hex$(i)
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
                Identifier ("DumpData"),
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
                            IntegerLiteral ("0"),
                        },
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("255"),
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                Identifier ("Hex$"),
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
    fn hex_dollar_concatenation() {
        let source = r#"
Sub Main()
    Dim output As String
    output = "Value: 0x" & Hex$(value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Hex$"));
    }

    #[test]
    fn hex_dollar_checksum() {
        let source = r#"
Function FormatChecksum(checksum As Long) As String
    Dim hexStr As String
    hexStr = Hex$(checksum)
    FormatChecksum = String$(8 - Len(hexStr), "0") & hexStr
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("FormatChecksum"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("checksum"),
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("hexStr"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("hexStr"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Hex$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("checksum"),
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
                            Identifier ("FormatChecksum"),
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
                    Unknown,
                    Whitespace,
                    Unknown,
                    Whitespace,
                    CallStatement {
                        LenKeyword,
                        LeftParenthesis,
                        Identifier ("hexStr"),
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        StringLiteral ("\"0\""),
                        RightParenthesis,
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("hexStr"),
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
    fn hex_dollar_rgb_components() {
        let source = r##"
Function RGBToWebColor(r As Byte, g As Byte, b As Byte) As String
    RGBToWebColor = "#" & Hex$(r) & Hex$(g) & Hex$(b)
End Function
"##;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Hex$"));
    }

    #[test]
    fn hex_dollar_with_prefix() {
        let source = r#"
Function HexWithPrefix(value As Long) As String
    HexWithPrefix = "&H" & Hex$(value)
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("HexWithPrefix"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("value"),
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("HexWithPrefix"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"&H\""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Hex$"),
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
    fn hex_dollar_conditional() {
        let source = r#"
Sub Main()
    If Len(Hex$(value)) > 4 Then
        Debug.Print "Large value"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Hex$"));
    }

    #[test]
    fn hex_dollar_select_case() {
        let source = r#"
Sub Main()
    Select Case Hex$(status)
        Case "FF"
            Debug.Print "All set"
        Case "0"
            Debug.Print "Clear"
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
                            Identifier ("Hex$"),
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
                            StringLiteral ("\"FF\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"All set\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\"0\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Clear\""),
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
    fn hex_dollar_multiple_calls() {
        let source = r#"
Function FormatPacket(pType As Byte, pLen As Integer) As String
    FormatPacket = "Type: 0x" & Hex$(pType) & " Len: 0x" & Hex$(pLen)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Hex$"));
    }

    #[test]
    fn hex_dollar_with_right() {
        let source = r#"
Function PadToTwo(value As Byte) As String
    PadToTwo = Right$("0" & Hex$(value), 2)
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("PadToTwo"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("value"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    ByteKeyword,
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
                            Identifier ("PadToTwo"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Right$"),
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
                                            Identifier ("Hex$"),
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
    fn hex_dollar_guid_format() {
        let source = r#"
Function FormatGUID(data1 As Long) As String
    FormatGUID = Right$("00000000" & Hex$(data1), 8)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Hex$"));
    }
}
