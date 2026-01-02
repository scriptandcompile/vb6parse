//! # `Hex` Function
//!
//! Returns a `String` representing the hexadecimal value of a number.
//!
//! ## Syntax
//!
//! ```vb
//! Hex(number)
//! ```
//!
//! ## Parameters
//!
//! - `number` (Required): Any valid numeric expression or string expression. If `number` is not already a whole number, it is rounded to the nearest whole number before being evaluated.
//!
//! ## Return Value
//!
//! Returns a `String` representing the hexadecimal value of the number. The return value contains hexadecimal digits (0-9, A-F) without the "&H" prefix.
//!
//! ## Remarks
//!
//! The `Hex` function converts a decimal number to its hexadecimal (base 16) string representation:
//!
//! - If `number` is `Null`, `Hex` returns `Null`
//! - If `number` is `Empty`, `Hex` returns "0"
//! - Negative numbers are represented in two's complement form
//! - For Byte values: Returns up to 2 hexadecimal digits
//! - For Integer values: Returns up to 4 hexadecimal digits
//! - For Long values: Returns up to 8 hexadecimal digits
//! - Fractional values are rounded to the nearest integer before conversion
//! - The result does not include the "&H" prefix (use "&H" & `Hex(n)` to include it)
//! - Leading zeros are not included in the result (e.g., 15 returns "F", not "0F")
//! - For hexadecimal to decimal conversion, use the `CLng` or `CInt` functions with "&H" prefix
//!
//! ## Typical Uses
//!
//! 1. **Color Values**: Convert RGB color values to hexadecimal for HTML/CSS
//! 2. **Memory Addresses**: Display memory addresses in hexadecimal format
//! 3. **Debugging**: Show binary data in readable hexadecimal format
//! 4. **Low-Level Programming**: Work with hardware registers and bit patterns
//! 5. **File I/O**: Display byte values when working with binary files
//! 6. **Cryptography**: Show hash values and checksums in hex format
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Simple conversion
//! Debug.Print Hex(255)        ' Prints: FF
//! Debug.Print Hex(16)         ' Prints: 10
//! Debug.Print Hex(0)          ' Prints: 0
//!
//! ' Example 2: Color conversion
//! Dim red As Long
//! red = RGB(255, 0, 0)
//! Debug.Print Hex(red)        ' Prints: FF (on little-endian systems)
//!
//! ' Example 3: Negative numbers (two's complement)
//! Debug.Print Hex(-1)         ' Prints: FFFFFFFF (Long)
//! Debug.Print Hex(-256)       ' Prints: FFFFFF00
//!
//! ' Example 4: With "&H" prefix
//! Dim value As Long
//! value = 42
//! Debug.Print "&H" & Hex(value)  ' Prints: &H2A
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: RGB to hex color string
//! Function RGBToHex(r As Integer, g As Integer, b As Integer) As String
//!     RGBToHex = Right$("0" & Hex(r), 2) & _
//!                Right$("0" & Hex(g), 2) & _
//!                Right$("0" & Hex(b), 2)
//! End Function
//!
//! ' Pattern 2: Pad hex string to specific length
//! Function HexPad(value As Long, length As Integer) As String
//!     HexPad = Right$(String$(length, "0") & Hex(value), length)
//! End Function
//!
//! ' Pattern 3: Display memory dump
//! Sub ShowMemoryDump(bytes() As Byte)
//!     Dim i As Long
//!     Dim line As String
//!     For i = LBound(bytes) To UBound(bytes)
//!         line = line & Right$("0" & Hex(bytes(i)), 2) & " "
//!         If (i + 1) Mod 16 = 0 Then
//!             Debug.Print line
//!             line = ""
//!         End If
//!     Next i
//!     If Len(line) > 0 Then Debug.Print line
//! End Sub
//!
//! ' Pattern 4: Format address with hex
//! Function FormatAddress(addr As Long) As String
//!     FormatAddress = "0x" & Right$("00000000" & Hex(addr), 8)
//! End Function
//!
//! ' Pattern 5: Convert byte array to hex string
//! Function BytesToHex(bytes() As Byte) As String
//!     Dim result As String
//!     Dim i As Long
//!     For i = LBound(bytes) To UBound(bytes)
//!         result = result & Right$("0" & Hex(bytes(i)), 2)
//!     Next i
//!     BytesToHex = result
//! End Function
//!
//! ' Pattern 6: Parse hex string back to number
//! Function HexToLong(hexStr As String) As Long
//!     If Left$(hexStr, 2) = "&H" Then
//!         HexToLong = CLng(hexStr)
//!     Else
//!         HexToLong = CLng("&H" & hexStr)
//!     End If
//! End Function
//!
//! ' Pattern 7: Show bit pattern
//! Sub ShowBitPattern(value As Long)
//!     Debug.Print "Hex: " & Hex(value)
//!     Debug.Print "Dec: " & value
//! End Sub
//!
//! ' Pattern 8: Color component extraction
//! Function GetRedComponent(color As Long) As Integer
//!     GetRedComponent = color And &HFF
//!     Debug.Print "Red: " & Hex(GetRedComponent)
//! End Function
//!
//! ' Pattern 9: Build lookup table
//! Dim hexLookup(0 To 255) As String
//! Sub InitHexLookup()
//!     Dim i As Long
//!     For i = 0 To 255
//!         hexLookup(i) = Right$("0" & Hex(i), 2)
//!     Next i
//! End Sub
//!
//! ' Pattern 10: Format hex with separator
//! Function HexWithSeparator(value As Long, separator As String) As String
//!     Dim hexStr As String
//!     Dim i As Integer
//!     hexStr = Right$("00000000" & Hex(value), 8)
//!     For i = 1 To 7 Step 2
//!         HexWithSeparator = HexWithSeparator & Mid$(hexStr, i, 2)
//!         If i < 7 Then HexWithSeparator = HexWithSeparator & separator
//!     Next i
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Hex dump utility class
//! Public Class HexDumper
//!     Private Const BYTES_PER_LINE As Integer = 16
//!     
//!     Public Function DumpToString(data() As Byte) As String
//!         Dim result As String
//!         Dim i As Long
//!         Dim offset As Long
//!         Dim hexPart As String
//!         Dim asciiPart As String
//!         Dim b As Byte
//!         
//!         For i = LBound(data) To UBound(data)
//!             If i Mod BYTES_PER_LINE = 0 Then
//!                 If i > 0 Then
//!                     result = result & hexPart & "  " & asciiPart & vbCrLf
//!                 End If
//!                 hexPart = Right$("00000000" & Hex(i), 8) & ": "
//!                 asciiPart = ""
//!             End If
//!             
//!             b = data(i)
//!             hexPart = hexPart & Right$("0" & Hex(b), 2) & " "
//!             
//!             If b >= 32 And b <= 126 Then
//!                 asciiPart = asciiPart & Chr$(b)
//!             Else
//!                 asciiPart = asciiPart & "."
//!             End If
//!         Next i
//!         
//!         ' Pad last line
//!         If Len(hexPart) > 10 Then
//!             hexPart = hexPart & String$((BYTES_PER_LINE - Len(asciiPart)) * 3, " ")
//!             result = result & hexPart & "  " & asciiPart
//!         End If
//!         
//!         DumpToString = result
//!     End Function
//! End Class
//!
//! ' Example 2: HTML color converter
//! Public Class ColorConverter
//!     Public Function VBColorToHTML(vbColor As Long) As String
//!         Dim r As Integer, g As Integer, b As Integer
//!         r = vbColor And &HFF
//!         g = (vbColor \ &H100) And &HFF
//!         b = (vbColor \ &H10000) And &HFF
//!         
//!         VBColorToHTML = "#" & _
//!             Right$("0" & Hex(r), 2) & _
//!             Right$("0" & Hex(g), 2) & _
//!             Right$("0" & Hex(b), 2)
//!     End Function
//!     
//!     Public Function HTMLToVBColor(htmlColor As String) As Long
//!         Dim r As Long, g As Long, b As Long
//!         If Left$(htmlColor, 1) = "#" Then htmlColor = Mid$(htmlColor, 2)
//!         
//!         r = CLng("&H" & Mid$(htmlColor, 1, 2))
//!         g = CLng("&H" & Mid$(htmlColor, 3, 2))
//!         b = CLng("&H" & Mid$(htmlColor, 5, 2))
//!         
//!         HTMLToVBColor = RGB(r, g, b)
//!     End Function
//! End Class
//!
//! ' Example 3: Checksum calculator
//! Public Function CalculateChecksum(data() As Byte) As String
//!     Dim checksum As Long
//!     Dim i As Long
//!     
//!     For i = LBound(data) To UBound(data)
//!         checksum = checksum Xor data(i)
//!         checksum = ((checksum And &H7FFFFFFF) * 2) Or (checksum And &H80000000) \ &H80000000
//!     Next i
//!     
//!     CalculateChecksum = Right$("00000000" & Hex(checksum), 8)
//! End Function
//!
//! ' Example 4: Binary file viewer
//! Public Class BinaryFileViewer
//!     Public Function LoadAndDisplayFile(filePath As String) As String
//!         Dim fileNum As Integer
//!         Dim fileData() As Byte
//!         Dim fileSize As Long
//!         Dim result As String
//!         Dim i As Long
//!         
//!         fileNum = FreeFile
//!         Open filePath For Binary As #fileNum
//!         fileSize = LOF(fileNum)
//!         ReDim fileData(0 To fileSize - 1)
//!         Get #fileNum, , fileData
//!         Close #fileNum
//!         
//!         For i = 0 To UBound(fileData)
//!             If i Mod 16 = 0 Then
//!                 result = result & vbCrLf & Right$("00000000" & Hex(i), 8) & ": "
//!             End If
//!             result = result & Right$("0" & Hex(fileData(i)), 2) & " "
//!         Next i
//!         
//!         LoadAndDisplayFile = result
//!     End Function
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! The Hex function generally doesn't raise errors, but type conversion can:
//!
//! - **Type Mismatch (Error 13)**: If the argument cannot be converted to a number
//! - **Overflow (Error 6)**: If the value exceeds Long range before conversion
//! - **Null Propagation**: If the argument is Null, Hex returns Null
//!
//! ```vb
//! On Error Resume Next
//! Dim result As String
//! result = Hex(someValue)
//! If Err.Number <> 0 Then
//!     Debug.Print "Error converting to hex: " & Err.Description
//!     Err.Clear
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: Hex conversion is a fast, native operation
//! - **String Building**: When converting many values, use string builder pattern to avoid repeated concatenation
//! - **Lookup Tables**: For repeated conversions of small numbers (0-255), consider pre-building a lookup table
//! - **Memory**: The returned string length depends on the magnitude of the number
//!
//! ## Best Practices
//!
//! 1. **Padding**: Use Right$() or Format$() to pad hex values to fixed width
//! 2. **Prefixes**: Add "&H" prefix when the value will be parsed back to numeric
//! 3. **Case**: VB6 Hex returns uppercase (A-F); convert to lowercase if needed
//! 4. **Validation**: Validate input before conversion to avoid type mismatch errors
//! 5. **Documentation**: Document whether hex strings include prefixes ("0x", "&H")
//! 6. **Null Handling**: Check for Null before calling Hex if working with Variants
//!
//! ## Comparison with Other Functions
//!
//! | Function | Purpose | Example |
//! |----------|---------|---------|
//! | Hex | Decimal to hexadecimal string | Hex(255) → "FF" |
//! | Oct | Decimal to octal string | Oct(8) → "10" |
//! | Str | Number to decimal string | Str(255) → " 255" |
//! | Format | Number to formatted string | Format(255, "00") → "255" |
//! | CLng("&H...") | Hexadecimal string to Long | CLng("&HFF") → 255 |
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Behavior consistent across Windows platforms
//! - Integer size (2 bytes) and Long size (4 bytes) are fixed
//! - Two's complement representation for negative numbers is standard
//!
//! ## Limitations
//!
//! - Cannot directly convert Currency or Decimal types (convert to Long first)
//! - No built-in support for padding with leading zeros (use Right$ or Format$)
//! - Returns uppercase letters only (A-F, not a-f)
//! - No control over prefix inclusion (always excludes "&H")
//! - Maximum value is Long range (−2,147,483,648 to 2,147,483,647)
//! - Fractional parts are rounded, not truncated
//!
//! ## Related Functions
//!
//! - `Oct`: Returns a String representing the octal value of a number
//! - `Str`: Returns a String representation of a number
//! - `Format`: Returns a Variant (String) formatted according to instructions
//! - `CLng`: Converts an expression to a Long
//! - `CInt`: Converts an expression to an Integer
//! - `Val`: Returns the numbers contained in a string as a numeric value

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn hex_basic() {
        let source = r"
Sub Test()
    result = Hex(255)
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
                            Identifier ("Hex"),
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
    fn hex_in_function() {
        let source = r"
Function ToHexString(value As Long) As String
    ToHexString = Hex(value)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ToHexString"),
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
                            Identifier ("ToHexString"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Hex"),
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
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn hex_with_prefix() {
        let source = r#"
Sub Test()
    hexStr = "&H" & Hex(42)
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
                            Identifier ("hexStr"),
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
                                Identifier ("Hex"),
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
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn hex_debug_print() {
        let source = r"
Sub Test()
    Debug.Print Hex(255)
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
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("Hex"),
                        LeftParenthesis,
                        IntegerLiteral ("255"),
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
    fn hex_if_statement() {
        let source = r#"
Sub Test()
    If Hex(value) = "FF" Then
        Debug.Print "Maximum byte value"
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
                                Identifier ("Hex"),
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
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"FF\""),
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
                                StringLiteral ("\"Maximum byte value\""),
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
    fn hex_for_loop() {
        let source = r"
Sub Test()
    Dim i As Integer
    For i = 0 To 255
        hexValues(i) = Hex(i)
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
                            AssignmentStatement {
                                CallExpression {
                                    Identifier ("hexValues"),
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
                                    Identifier ("Hex"),
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
    fn hex_select_case() {
        let source = r#"
Sub Test()
    Select Case Hex(errorCode)
        Case "FF"
            HandleError
        Case "0"
            Success
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
                            Identifier ("Hex"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("errorCode"),
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
                                    Identifier ("HandleError"),
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
                                    Identifier ("Success"),
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
    fn hex_do_loop() {
        let source = r"
Sub Test()
    Do While Len(Hex(counter)) < 8
        counter = counter + 1
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
                        BinaryExpression {
                            CallExpression {
                                LenKeyword,
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        CallExpression {
                                            Identifier ("Hex"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    IdentifierExpression {
                                                        Identifier ("counter"),
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
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("8"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("counter"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("counter"),
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
    fn hex_class_member() {
        let source = r"
Private Sub Class_Initialize()
    m_hexValue = Hex(initialValue)
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
                            Identifier ("m_hexValue"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Hex"),
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
    fn hex_type_field() {
        let source = r"
Sub Test()
    Dim config As ConfigType
    config.hexCode = Hex(statusValue)
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
                        Identifier ("config"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Identifier ("ConfigType"),
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        MemberAccessExpression {
                            Identifier ("config"),
                            PeriodOperator,
                            Identifier ("hexCode"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Hex"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("statusValue"),
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
    fn hex_collection_add() {
        let source = r#"
Sub Test()
    Dim col As New Collection
    col.Add Hex(value), "HexValue"
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
                        Identifier ("col"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        NewKeyword,
                        Whitespace,
                        Identifier ("Collection"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("col"),
                        PeriodOperator,
                        Identifier ("Add"),
                        Whitespace,
                        Identifier ("Hex"),
                        LeftParenthesis,
                        Identifier ("value"),
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        StringLiteral ("\"HexValue\""),
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
    fn hex_with_statement() {
        let source = r"
Sub Test()
    With myObject
        .HexString = Hex(.Value)
    End With
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
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("myObject"),
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("HexString"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Hex"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    PeriodOperator,
                                                },
                                            },
                                        },
                                    },
                                },
                            },
                            CallStatement {
                                Identifier ("Value"),
                                RightParenthesis,
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
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
    fn hex_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Hex value: " & Hex(errorCode)
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
                        StringLiteral ("\"Hex value: \""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Hex"),
                        LeftParenthesis,
                        Identifier ("errorCode"),
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
    fn hex_property() {
        let source = r"
Property Get HexValue() As String
    HexValue = Hex(m_value)
End Property
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            PropertyStatement {
                PropertyKeyword,
                Whitespace,
                GetKeyword,
                Whitespace,
                Identifier ("HexValue"),
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("HexValue"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Hex"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("m_value"),
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
                PropertyKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn hex_concatenation() {
        let source = r#"
Sub Test()
    result = "0x" & Hex(address)
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
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"0x\""),
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
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn hex_for_each() {
        let source = r"
Sub Test()
    Dim item As Variant
    For Each item In collection
        Debug.Print Hex(item)
    Next
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
                        Identifier ("item"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        VariantKeyword,
                        Newline,
                    },
                    ForEachStatement {
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        EachKeyword,
                        Whitespace,
                        Identifier ("item"),
                        Whitespace,
                        InKeyword,
                        Whitespace,
                        Identifier ("collection"),
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                Identifier ("Hex"),
                                LeftParenthesis,
                                Identifier ("item"),
                                RightParenthesis,
                                Newline,
                            },
                            Whitespace,
                        },
                        NextKeyword,
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
    fn hex_error_handling() {
        let source = r"
Sub Test()
    On Error Resume Next
    hexStr = Hex(value)
    If Err.Number <> 0 Then Err.Clear
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
                            Identifier ("hexStr"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Hex"),
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
                        Whitespace,
                        Identifier ("Err"),
                        PeriodOperator,
                        Identifier ("Clear"),
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

    #[test]
    fn hex_right_function() {
        let source = r#"
Sub Test()
    padded = Right$("00" & Hex(value), 2)
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
                            Identifier ("padded"),
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
                                            StringLiteral ("\"00\""),
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
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn hex_array_assignment() {
        let source = r"
Sub Test()
    Dim hexArray(1 To 10) As String
    hexArray(1) = Hex(255)
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
                        Identifier ("hexArray"),
                        LeftParenthesis,
                        NumericLiteralExpression {
                            IntegerLiteral ("1"),
                        },
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        CallExpression {
                            Identifier ("hexArray"),
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
                            Identifier ("Hex"),
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
    fn hex_function_argument() {
        let source = r"
Sub Test()
    DisplayHexValue Hex(colorValue)
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
                    CallStatement {
                        Identifier ("DisplayHexValue"),
                        Whitespace,
                        Identifier ("Hex"),
                        LeftParenthesis,
                        Identifier ("colorValue"),
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
    fn hex_nested_call() {
        let source = r"
Sub Test()
    length = Len(Hex(value))
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
                            Identifier ("length"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            LenKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Hex"),
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
    fn hex_iif() {
        let source = r"
Sub Test()
    display = IIf(showHex, Hex(value), Str(value))
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
                            Identifier ("display"),
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
                                        Identifier ("showHex"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    CallExpression {
                                        Identifier ("Hex"),
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
                                    CallExpression {
                                        Identifier ("Str"),
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
    fn hex_color_conversion() {
        let source = r#"
Function RGBToHex(r As Integer, g As Integer, b As Integer) As String
    RGBToHex = Right$("0" & Hex(r), 2) & Right$("0" & Hex(g), 2) & Right$("0" & Hex(b), 2)
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("RGBToHex"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("r"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("g"),
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
                StringKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("RGBToHex"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
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
                                                    Identifier ("Hex"),
                                                    LeftParenthesis,
                                                    ArgumentList {
                                                        Argument {
                                                            IdentifierExpression {
                                                                Identifier ("r"),
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
                                Whitespace,
                                Ampersand,
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
                                                    Identifier ("Hex"),
                                                    LeftParenthesis,
                                                    ArgumentList {
                                                        Argument {
                                                            IdentifierExpression {
                                                                Identifier ("g"),
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
                            },
                            Whitespace,
                            Ampersand,
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
                                                Identifier ("Hex"),
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        IdentifierExpression {
                                                            Identifier ("b"),
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
    fn hex_byte_array() {
        let source = r#"
Sub Test()
    Dim bytes(0 To 15) As Byte
    Dim i As Integer
    For i = 0 To 15
        Debug.Print Right$("0" & Hex(bytes(i)), 2);
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
                        Identifier ("bytes"),
                        LeftParenthesis,
                        NumericLiteralExpression {
                            IntegerLiteral ("0"),
                        },
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("15"),
                        },
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        ByteKeyword,
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
                            IntegerLiteral ("15"),
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                Identifier ("Right$"),
                                LeftParenthesis,
                                StringLiteral ("\"0\""),
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                Identifier ("Hex"),
                                LeftParenthesis,
                                Identifier ("bytes"),
                                LeftParenthesis,
                                Identifier ("i"),
                                RightParenthesis,
                                RightParenthesis,
                                Comma,
                                Whitespace,
                                IntegerLiteral ("2"),
                                RightParenthesis,
                                Semicolon,
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
    fn hex_parentheses() {
        let source = r"
Sub Test()
    value = (Hex(number))
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
                        ParenthesizedExpression {
                            LeftParenthesis,
                            CallExpression {
                                Identifier ("Hex"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("number"),
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
    fn hex_format_address() {
        let source = r#"
Function FormatAddress(addr As Long) As String
    FormatAddress = "0x" & Right$("00000000" & Hex(addr), 8)
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
                    Identifier ("addr"),
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
                                Identifier ("Right$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            StringLiteralExpression {
                                                StringLiteral ("\"00000000\""),
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
                                                            Identifier ("addr"),
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
                                            IntegerLiteral ("8"),
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
    fn hex_negative_value() {
        let source = r"
Sub Test()
    Dim negHex As String
    negHex = Hex(-1)
    Debug.Print negHex
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
                        Identifier ("negHex"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("negHex"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Hex"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
                                        },
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
                        Identifier ("negHex"),
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
