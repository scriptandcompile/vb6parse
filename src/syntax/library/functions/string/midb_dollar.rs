//! # `MidB$` Function
//!
//! Returns a `String` containing a specified number of bytes from a string.
//!
//! ## Syntax
//!
//! ```vb6
//! MidB$(string, start[, length])
//! ```
//!
//! ## Parameters
//!
//! - `string`: Required. String expression from which bytes are returned. If `string` contains `Null`, `Null` is returned.
//! - `start`: Required. Byte position in `string` at which the part to be taken begins. If `start` is greater than the number of bytes in `string`, `MidB$` returns a zero-length string ("").
//! - `length`: Optional. Number of bytes to return. If omitted or if there are fewer than `length` bytes in the text (including the byte at `start`), all bytes from `start` to the end of the string are returned.
//!
//! ## Return Value
//!
//! Returns a `String` containing the specified number of bytes from `string`, starting at the byte position `start`. If `length` is omitted, returns all bytes from `start` to the end. Returns an empty string if `start` is greater than the byte length of the string.
//!
//! ## Remarks
//!
//! The `MidB$` function is used with byte data contained in a string. Instead of specifying character positions and counts, `start` and `length` specify byte positions and byte counts.
//!
//! This function is particularly useful when working with:
//! - Binary data stored in strings
//! - ANSI strings requiring byte-level manipulation
//! - Double-byte character set (DBCS) strings
//! - Fixed-position binary protocols
//! - File format parsing at byte level
//! - Network packet extraction
//!
//! `MidB$` is the byte-oriented version of `Mid$`. While `Mid$` counts characters, `MidB$` counts bytes. In single-byte character sets (like standard ASCII), these are equivalent, but in DBCS systems (like Japanese, Chinese, or Korean Windows), characters may occupy multiple bytes.
//!
//! The `MidB$` function always returns a `String`. The `MidB` function returns a `Variant`.
//!
//! Note: Byte positions are 1-based, not 0-based. The first byte is at position 1.
//!
//! ## Typical Uses
//!
//! ### Example 1: Extracting Binary Field
//! ```vb6
//! Dim record As String
//! record = GetBinaryRecord()
//! field = MidB$(record, 5, 4)  ' Get 4 bytes starting at byte 5
//! ```
//!
//! ### Example 2: Parsing Protocol Header
//! ```vb6
//! Dim packet As String
//! packet = ReceivePacket()
//! version = MidB$(packet, 1, 2)   ' First 2 bytes
//! msgType = MidB$(packet, 3, 1)   ' Third byte
//! ```
//!
//! ### Example 3: Reading Fixed-Position Data
//! ```vb6
//! Dim data As String
//! data = binaryData
//! id = MidB$(data, 1, 8)      ' Bytes 1-8: ID
//! name = MidB$(data, 9, 32)   ' Bytes 9-40: Name
//! ```
//!
//! ### Example 4: Extracting GUID Components
//! ```vb6
//! Dim guidBytes As String
//! guidBytes = GetGUID()
//! data1 = MidB$(guidBytes, 1, 4)   ' First DWORD
//! data2 = MidB$(guidBytes, 5, 2)   ' First WORD
//! data3 = MidB$(guidBytes, 7, 2)   ' Second WORD
//! ```
//!
//! ## Common Usage Patterns
//!
//! ### Parsing Binary Structure
//! ```vb6
//! Dim buffer As String
//! buffer = GetBinaryData()
//! magic = MidB$(buffer, 1, 4)        ' Magic number
//! version = MidB$(buffer, 5, 2)      ' Version
//! flags = MidB$(buffer, 7, 1)        ' Flags byte
//! dataLen = MidB$(buffer, 8, 4)      ' Data length
//! ```
//!
//! ### Reading File Header
//! ```vb6
//! Dim fileData As String
//! Open fileName For Binary As #1
//! fileData = Input$(100, #1)
//! Close #1
//! signature = MidB$(fileData, 1, 4)
//! If signature = "RIFF" Then
//!     fileSize = MidB$(fileData, 5, 4)
//! End If
//! ```
//!
//! ### Processing Network Packet
//! ```vb6
//! Dim packet As String
//! packet = Socket.Receive()
//! header = MidB$(packet, 1, 16)
//! payload = MidB$(packet, 17)  ' Rest of packet
//! ```
//!
//! ### Extracting Length-Prefixed String
//! ```vb6
//! Dim message As String
//! message = buffer
//! lenByte = MidB$(message, 1, 1)
//! strLen = AscB(lenByte)
//! text = MidB$(message, 2, strLen)
//! ```
//!
//! ### Reading BMP Image Data
//! ```vb6
//! Dim bmpData As String
//! Open "image.bmp" For Binary As #1
//! bmpData = Input$(LOF(1), #1)
//! Close #1
//! fileType = MidB$(bmpData, 1, 2)     ' "BM"
//! fileSize = MidB$(bmpData, 3, 4)     ' File size
//! offset = MidB$(bmpData, 11, 4)      ' Pixel data offset
//! ```
//!
//! ### Chunked Binary Processing
//! ```vb6
//! Dim data As String, chunk As String
//! Dim pos As Long
//! data = GetBinaryData()
//! pos = 1
//! Do While pos <= LenB(data)
//!     chunk = MidB$(data, pos, 512)
//!     ProcessChunk chunk
//!     pos = pos + 512
//! Loop
//! ```
//!
//! ### Extracting Record Fields
//! ```vb6
//! Dim record As String
//! record = ReadBinaryRecord()
//! ' Fixed-position record layout
//! customerID = MidB$(record, 1, 10)
//! orderDate = MidB$(record, 11, 8)
//! amount = MidB$(record, 19, 8)
//! ```
//!
//! ### Parsing IPv4 Address Bytes
//! ```vb6
//! Dim ipBytes As String
//! ipBytes = GetIPv4Bytes()
//! octet1 = AscB(MidB$(ipBytes, 1, 1))
//! octet2 = AscB(MidB$(ipBytes, 2, 1))
//! octet3 = AscB(MidB$(ipBytes, 3, 1))
//! octet4 = AscB(MidB$(ipBytes, 4, 1))
//! ```
//!
//! ### Reading Bitmap Header Fields
//! ```vb6
//! Dim dibHeader As String
//! dibHeader = MidB$(bmpData, 15, 40)  ' DIB header
//! width = MidB$(dibHeader, 5, 4)
//! height = MidB$(dibHeader, 9, 4)
//! bitsPerPixel = MidB$(dibHeader, 15, 2)
//! ```
//!
//! ### Protocol Message Parsing
//! ```vb6
//! Dim msg As String
//! msg = ReceiveMessage()
//! msgID = MidB$(msg, 1, 4)
//! timestamp = MidB$(msg, 5, 8)
//! sender = MidB$(msg, 13, 16)
//! payload = MidB$(msg, 29)  ' Remaining bytes
//! ```
//!
//! ## Related Functions
//!
//! - `MidB`: Variant version that returns a `Variant`
//! - `Mid$`: Character-based version that counts characters
//! - `LeftB$`: Returns bytes from the left side of a string
//! - `RightB$`: Returns bytes from the right side of a string
//! - `LenB`: Returns the number of bytes in a string
//! - `InStrB`: Finds byte position of a substring
//! - `AscB`: Returns the byte value at a position
//! - `ChrB$`: Returns a string containing a single byte
//!
//! ## Best Practices
//!
//! 1. Use `MidB$` when working with binary data or byte-oriented protocols
//! 2. Be careful with DBCS strings - splitting may corrupt multi-byte characters
//! 3. Always validate byte positions and lengths before extraction
//! 4. Use `LenB` to get byte length, not `Len`
//! 5. Remember that byte positions are 1-based, not 0-based
//! 6. Prefer `MidB$` over `Mid$` for binary file operations
//! 7. Combine with `AscB` and `ChrB$` for byte-level manipulation
//! 8. Document byte offsets and field sizes for binary structures
//! 9. Test DBCS string operations on appropriate language systems
//! 10. Use constants for magic numbers and field offsets
//!
//! ## Performance Considerations
//!
//! - `MidB$` is slightly faster than `Mid$` for binary data
//! - No performance penalty for requesting bytes beyond string end
//! - Direct byte extraction is faster than loops with `AscB`
//! - Extracting large byte ranges is efficient
//! - Minimal overhead compared to `Mid$` in SBCS environments
//!
//! ## Character Set Behavior
//!
//! | Environment | Bytes per Char | Notes |
//! |-------------|----------------|-------|
//! | English Windows | 1 byte | `MidB$` and `Mid$` behave identically |
//! | DBCS Windows | 1-2 bytes | `MidB$` may split multi-byte characters |
//! | Unicode VB6 | 2 bytes | Internal strings are Unicode but converted |
//! | Binary Data | N/A | `MidB$` treats data as raw bytes |
//!
//! ## Common Pitfalls
//!
//! - Using `MidB$` on DBCS strings without checking character boundaries
//! - Confusing byte positions with character positions
//! - Using 0-based indexing (VB6 strings are 1-based)
//! - Not handling `Null` string values (causes runtime error)
//! - Passing negative or zero start position (causes runtime error)
//! - Using `MidB$` for text processing (use `Mid$` instead)
//! - Forgetting that VB6 strings are internally Unicode
//! - Assuming one byte equals one character in all locales
//! - Not validating that `start` position is within bounds
//!
//! ## Limitations
//!
//! - May corrupt DBCS characters if not used carefully
//! - Returns `Null` if the string argument is `Null`
//! - Start position must be positive (1 or greater)
//! - Length parameter cannot be negative
//! - Not suitable for modern Unicode string processing
//! - Limited to VB6's internal string representation
//! - May produce unexpected results with emoji or complex Unicode
//! - Cannot modify the original string (read-only operation)

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn midb_dollar_simple() {
        let source = r#"
Sub Main()
    result = MidB$("Hello", 2, 3)
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
                            Identifier ("MidB$"),
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
                                        IntegerLiteral ("2"),
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
    fn midb_dollar_assignment() {
        let source = r"
Sub Main()
    Dim field As String
    field = MidB$(record, 5, 4)
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
                        Identifier ("field"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("field"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("record"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("5"),
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
    fn midb_dollar_variable() {
        let source = r"
Sub Main()
    data = GetBinaryData()
    bytes = MidB$(data, 1, 8)
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
                            Identifier ("data"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("GetBinaryData"),
                            LeftParenthesis,
                            ArgumentList,
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("bytes"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
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
                                        IntegerLiteral ("8"),
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
    fn midb_dollar_no_length() {
        let source = r"
Sub Main()
    tail = MidB$(data, 10)
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
                            Identifier ("tail"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
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
    fn midb_dollar_protocol_header() {
        let source = r"
Sub Main()
    packet = ReceivePacket()
    version = MidB$(packet, 1, 2)
    msgType = MidB$(packet, 3, 1)
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
                            Identifier ("packet"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("ReceivePacket"),
                            LeftParenthesis,
                            ArgumentList,
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            VersionKeyword,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("packet"),
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
                                        IntegerLiteral ("2"),
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
                            Identifier ("msgType"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("packet"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("3"),
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
    fn midb_dollar_guid_extraction() {
        let source = r"
Sub Main()
    guidBytes = GetGUID()
    data1 = MidB$(guidBytes, 1, 4)
    data2 = MidB$(guidBytes, 5, 2)
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
                            Identifier ("guidBytes"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("GetGUID"),
                            LeftParenthesis,
                            ArgumentList,
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("data1"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("guidBytes"),
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
                                        IntegerLiteral ("4"),
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
                            Identifier ("data2"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("guidBytes"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("5"),
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
    fn midb_dollar_in_condition() {
        let source = r#"
Sub Main()
    If MidB$(fileData, 1, 4) = "RIFF" Then
        Debug.Print "WAV file"
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
                                Identifier ("MidB$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("fileData"),
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
                                StringLiteral ("\"RIFF\""),
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
                                StringLiteral ("\"WAV file\""),
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
    fn midb_dollar_binary_structure() {
        let source = r"
Sub Main()
    buffer = GetBinaryData()
    magic = MidB$(buffer, 1, 4)
    version = MidB$(buffer, 5, 2)
    flags = MidB$(buffer, 7, 1)
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
                            Identifier ("buffer"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("GetBinaryData"),
                            LeftParenthesis,
                            ArgumentList,
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("magic"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("buffer"),
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
                                        IntegerLiteral ("4"),
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
                            VersionKeyword,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("buffer"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("5"),
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("flags"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("buffer"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("7"),
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
    fn midb_dollar_with_lenb() {
        let source = r"
Sub Main()
    If LenB(data) >= 10 Then
        chunk = MidB$(data, 1, 10)
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
                                Identifier ("LenB"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("data"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            GreaterThanOrEqualOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("chunk"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("MidB$"),
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
                                                IntegerLiteral ("10"),
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
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn midb_dollar_loop_processing() {
        let source = r"
Sub Main()
    Dim pos As Long
    pos = 1
    Do While pos <= LenB(data)
        chunk = MidB$(data, pos, 512)
        pos = pos + 512
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
                        Identifier ("pos"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        LongKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("pos"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("1"),
                        },
                        Newline,
                    },
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("pos"),
                            },
                            Whitespace,
                            LessThanOrEqualOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("LenB"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("data"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("chunk"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("MidB$"),
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
                                            IdentifierExpression {
                                                Identifier ("pos"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("512"),
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
                                    Identifier ("pos"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("pos"),
                                    },
                                    Whitespace,
                                    AdditionOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("512"),
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
    fn midb_dollar_select_case() {
        let source = r#"
Sub Main()
    sig = MidB$(data, 1, 4)
    Select Case sig
        Case "RIFF"
            Debug.Print "RIFF file"
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
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("sig"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
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
                                        IntegerLiteral ("4"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("sig"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\"RIFF\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"RIFF file\""),
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
    fn midb_dollar_multiple_uses() {
        let source = r"
Sub Main()
    header = MidB$(buffer, 1, 16)
    payload = MidB$(buffer, 17)
    combined = header & payload
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
                            Identifier ("header"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("buffer"),
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
                                        IntegerLiteral ("16"),
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
                            Identifier ("payload"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("buffer"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("17"),
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
                            Identifier ("combined"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("header"),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("payload"),
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
    fn midb_dollar_in_function() {
        let source = r"
Function ExtractField(data As String, offset As Long) As String
    ExtractField = MidB$(data, offset, 8)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ExtractField"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("data"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("offset"),
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
                            Identifier ("ExtractField"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
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
                                    IdentifierExpression {
                                        Identifier ("offset"),
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
    fn midb_dollar_expression_args() {
        let source = r"
Sub Main()
    n = 5
    result = MidB$(data, n * 2, n + 3)
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
                            Identifier ("n"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("5"),
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
                            Identifier ("MidB$"),
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
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("n"),
                                        },
                                        Whitespace,
                                        MultiplicationOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("2"),
                                        },
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("n"),
                                        },
                                        Whitespace,
                                        AdditionOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("3"),
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
    fn midb_dollar_bmp_parsing() {
        let source = r#"
Sub Main()
    Open "image.bmp" For Binary As #1
    bmpData = Input$(54, #1)
    Close #1
    fileType = MidB$(bmpData, 1, 2)
    fileSize = MidB$(bmpData, 3, 4)
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
                    OpenStatement {
                        Whitespace,
                        OpenKeyword,
                        Whitespace,
                        StringLiteral ("\"image.bmp\""),
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        BinaryKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("bmpData"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            InputKeyword,
                            DollarSign,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("54"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Octothorpe,
                                    },
                                },
                            },
                        },
                    },
                    Unknown,
                    Unknown,
                    Newline,
                    CloseStatement {
                        Whitespace,
                        CloseKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("fileType"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("bmpData"),
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
                                        IntegerLiteral ("2"),
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
                            Identifier ("fileSize"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("bmpData"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("3"),
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
    fn midb_dollar_record_fields() {
        let source = r"
Sub Main()
    record = ReadBinaryRecord()
    customerID = MidB$(record, 1, 10)
    orderDate = MidB$(record, 11, 8)
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
                            Identifier ("record"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("ReadBinaryRecord"),
                            LeftParenthesis,
                            ArgumentList,
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("customerID"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("record"),
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
                                        IntegerLiteral ("10"),
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
                            Identifier ("orderDate"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("record"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("11"),
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
    fn midb_dollar_with_ascb() {
        let source = r"
Sub Main()
    ipBytes = GetIPv4Bytes()
    octet1 = AscB(MidB$(ipBytes, 1, 1))
    octet2 = AscB(MidB$(ipBytes, 2, 1))
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
                            Identifier ("ipBytes"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("GetIPv4Bytes"),
                            LeftParenthesis,
                            ArgumentList,
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("octet1"),
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
                                        Identifier ("MidB$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("ipBytes"),
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
                            Identifier ("octet2"),
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
                                        Identifier ("MidB$"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("ipBytes"),
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
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn midb_dollar_concatenation() {
        let source = r#"
Sub Main()
    prefix = "Header: " & MidB$(data, 1, 4)
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
                            Identifier ("prefix"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"Header: \""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("MidB$"),
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
                                            IntegerLiteral ("4"),
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
    fn midb_dollar_length_prefixed() {
        let source = r"
Sub Main()
    message = buffer
    lenByte = MidB$(message, 1, 1)
    strLen = AscB(lenByte)
    text = MidB$(message, 2, strLen)
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
                            Identifier ("message"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("buffer"),
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("lenByte"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("message"),
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
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("strLen"),
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
                                        Identifier ("lenByte"),
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
                            TextKeyword,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("message"),
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
                                    IdentifierExpression {
                                        Identifier ("strLen"),
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
    fn midb_dollar_network_packet() {
        let source = r"
Sub Main()
    msg = ReceiveMessage()
    msgID = MidB$(msg, 1, 4)
    timestamp = MidB$(msg, 5, 8)
    sender = MidB$(msg, 13, 16)
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
                            Identifier ("msg"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("ReceiveMessage"),
                            LeftParenthesis,
                            ArgumentList,
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("msgID"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("msg"),
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
                                        IntegerLiteral ("4"),
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
                            Identifier ("timestamp"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("msg"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("5"),
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
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("sender"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MidB$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("msg"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("13"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("16"),
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
