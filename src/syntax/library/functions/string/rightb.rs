//! # RightB Function
//!
//! Returns a Variant (String) containing a specified number of bytes from the right side of a string.
//!
//! ## Syntax
//!
//! ```vb
//! RightB(string, length)
//! ```
//!
//! ## Parameters
//!
//! - `string` (Required): String expression from which rightmost bytes are returned
//!   - If string contains Null, Null is returned
//! - `length` (Required): Numeric expression indicating how many bytes to return
//!   - If 0, empty string ("") is returned
//!   - If greater than or equal to number of bytes in string, entire string is returned
//!   - Must be non-negative (negative values cause error)
//!
//! ## Return Value
//!
//! Returns a Variant containing a String:
//! - Contains the specified number of bytes from the right side of the string
//! - Returns empty string if length is 0
//! - Returns entire string if length >= LenB(string)
//! - Returns Null if string argument is Null
//! - Returns Variant type (RightB$ variant returns String type directly)
//!
//! ## Remarks
//!
//! The RightB function extracts bytes from the end of a string:
//!
//! - Returns rightmost bytes up to specified length
//! - Operates on byte level, not character level
//! - Particularly useful with double-byte character sets (DBCS)
//! - Complements LeftB function (which returns leftmost bytes)
//! - Works with MidB function for complete byte-level substring extraction
//! - Extraction from end: RightB("ABC", 2) returns last 2 bytes
//! - Safe with lengths exceeding string byte length (returns full string)
//! - Null propagates through the function
//! - Negative length raises Error 5 (Invalid procedure call or argument)
//! - Common for extracting binary data suffixes, checksums, trailers
//! - More efficient than MidB for right extraction
//! - RightB$ variant returns String type (not Variant) for slight performance gain
//! - Cannot extract from left side (use LeftB for that)
//! - Cannot skip bytes (use MidB for that)
//! - Does not modify original string (strings are immutable)
//!
//! ## Differences from Right Function
//!
//! - Right operates on characters, RightB operates on bytes
//! - In single-byte character sets (SBCS), they are equivalent
//! - In double-byte character sets (DBCS), one character may be multiple bytes
//! - RightB is essential for binary data manipulation
//! - RightB is used with LenB (byte length) rather than Len (character length)
//!
//! ## Typical Uses
//!
//! 1. **Binary Data**: Extract trailing bytes from binary strings
//! 2. **Checksums**: Extract checksum bytes from data
//! 3. **File Trailers**: Extract trailing data from files
//! 4. **DBCS Strings**: Work with Japanese, Chinese, Korean text at byte level
//! 5. **Fixed Byte Records**: Parse trailing fields from binary records
//! 6. **Byte Validation**: Check byte suffixes in data
//! 7. **Binary Structures**: Extract trailing fields from binary structures
//! 8. **Network Data**: Process packet trailers
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Basic byte extraction
//! Dim data As String
//! data = Chr$(65) & Chr$(66) & Chr$(67)  ' "ABC"
//!
//! Debug.Print RightB(data, 2)            ' Last 2 bytes
//! Debug.Print RightB(data, 1)            ' Last byte
//!
//! ' Example 2: Checksum extraction
//! Dim packet As String
//! packet = ReceiveNetworkData()
//!
//! Dim checksum As String
//! checksum = RightB(packet, 4)  ' 4-byte checksum at end
//!
//! ' Example 3: File trailer
//! Dim fileData As String
//! Open "data.bin" For Binary As #1
//! fileData = Input$(LOF(1), #1)
//! Close #1
//!
//! Dim trailer As String
//! trailer = RightB(fileData, 16)  ' 16-byte trailer
//!
//! ' Example 4: DBCS text handling
//! Dim japaneseText As String
//! japaneseText = LoadJapaneseText()  ' Load Japanese text
//!
//! Dim lastBytes As String
//! lastBytes = RightB(japaneseText, 4)  ' Last 4 bytes (may be 2 DBCS chars)
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Extract checksum
//! Function GetChecksum(data As String) As String
//!     If LenB(data) < 4 Then
//!         GetChecksum = ""
//!     Else
//!         GetChecksum = RightB(data, 4)
//!     End If
//! End Function
//!
//! ' Pattern 2: Validate trailer
//! Function ValidateTrailer(data As String, trailer As String) As Boolean
//!     ValidateTrailer = (RightB(data, LenB(trailer)) = trailer)
//! End Function
//!
//! ' Pattern 3: Extract record suffix
//! Function GetRecordSuffix(record As String) As String
//!     ' Last 8 bytes contain record suffix
//!     GetRecordSuffix = RightB(record, 8)
//! End Function
//!
//! ' Pattern 4: Parse packet trailer
//! Sub ParsePacket(packet As String)
//!     Dim payload As String
//!     Dim trailer As String
//!     
//!     trailer = RightB(packet, 8)  ' 8-byte trailer
//!     payload = LeftB(packet, LenB(packet) - 8)  ' All but trailer
//!     
//!     ' Process payload and trailer
//! End Sub
//!
//! ' Pattern 5: Extract file extension bytes
//! Function GetExtensionBytes(fileName As String) As String
//!     ' Assumes extension is last 3 bytes after dot
//!     GetExtensionBytes = RightB(fileName, 3)
//! End Function
//! ```
//!
//! ## Advanced Examples
//!
//! ```vb
//! ' Example: CRC validation
//! Function ValidateCRC(data As String) As Boolean
//!     Dim crc As String
//!     Dim payload As String
//!     
//!     If LenB(data) < 4 Then
//!         ValidateCRC = False
//!         Exit Function
//!     End If
//!     
//!     crc = RightB(data, 4)
//!     payload = LeftB(data, LenB(data) - 4)
//!     
//!     ValidateCRC = (CalculateCRC(payload) = crc)
//! End Function
//!
//! ' Example: Extract GUID data
//! Sub ExtractGUID(guidData As String)
//!     Dim data4 As String
//!     data4 = RightB(guidData, 8)  ' Last 8 bytes of GUID
//!     Debug.Print "Data4: " & BytesToHex(data4)
//! End Sub
//!
//! ' Example: Binary record trailer
//! Function GetRecordTrailer(record As String) As String
//!     ' Records have 12-byte trailer
//!     Const TRAILER_SIZE As Long = 12
//!     
//!     If LenB(record) >= TRAILER_SIZE Then
//!         GetRecordTrailer = RightB(record, TRAILER_SIZE)
//!     Else
//!         GetRecordTrailer = record
//!     End If
//! End Function
//! ```
//!
//! ## See Also
//!
//! - `RightB$`: String-returning variant of RightB
//! - `LeftB`: Returns leftmost bytes from string
//! - `LeftB$`: String-returning variant of LeftB
//! - `MidB`: Returns bytes from middle of string
//! - `MidB$`: String-returning variant of MidB
//! - `LenB`: Returns byte length of string
//! - `Right`: Character-based right extraction
//! - `Left`: Character-based left extraction
//! - `Mid`: Character-based middle extraction

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn rightb_basic() {
        let source = r"
Sub Test()
    result = RightB(data, 4)
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
                            Identifier ("RightB"),
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
    fn rightb_checksum() {
        let source = r"
Function GetChecksum(packet As String) As String
    GetChecksum = RightB(packet, 4)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetChecksum"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("packet"),
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
                            Identifier ("GetChecksum"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RightB"),
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
    fn rightb_trailer() {
        let source = r"
Sub ParseData(data As String)
    trailer = RightB(data, 16)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ParseData"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("data"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("trailer"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RightB"),
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

    #[test]
    fn rightb_in_condition() {
        let source = r"
Sub Test()
    If RightB(data, 4) = checksumBytes Then
        ProcessData
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
                            CallExpression {
                                Identifier ("RightB"),
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
                                            IntegerLiteral ("4"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("checksumBytes"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("ProcessData"),
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
    fn rightb_binary_record() {
        let source = r"
Function GetRecordSuffix(record As String) As String
    GetRecordSuffix = RightB(record, 8)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetRecordSuffix"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("record"),
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
                            Identifier ("GetRecordSuffix"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RightB"),
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
    fn rightb_with_lenb() {
        let source = r"
Sub Test()
    If LenB(data) > 10 Then
        suffix = RightB(data, 10)
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
                            GreaterThanOperator,
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
                                    Identifier ("suffix"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("RightB"),
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
    fn rightb_dbcs_handling() {
        let source = r"
Sub Test()
    Dim jpText As String
    jpText = GetJapaneseText()
    bytes = RightB(jpText, 4)
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
                        Identifier ("jpText"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("jpText"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("GetJapaneseText"),
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
                            Identifier ("RightB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("jpText"),
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
    fn rightb_validation() {
        let source = r"
Function ValidateTrailer(data As String, trailer As String) As Boolean
    ValidateTrailer = (RightB(data, LenB(trailer)) = trailer)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ValidateTrailer"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("data"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("trailer"),
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
                            Identifier ("ValidateTrailer"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("RightB"),
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
                                                Identifier ("LenB"),
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        IdentifierExpression {
                                                            Identifier ("trailer"),
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
                                    Identifier ("trailer"),
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
    fn rightb_crc_validation() {
        let source = r"
Function ValidateCRC(data As String) As Boolean
    Dim crc As String
    Dim payload As String
    
    crc = RightB(data, 4)
    payload = LeftB(data, LenB(data) - 4)
    
    ValidateCRC = (CalculateCRC(payload) = crc)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ValidateCRC"),
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
                BooleanKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("crc"),
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
                        Identifier ("payload"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    Newline,
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("crc"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RightB"),
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
                            Identifier ("payload"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("LeftB"),
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
                                        SubtractionOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("4"),
                                        },
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    Newline,
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("ValidateCRC"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("CalculateCRC"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("payload"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("crc"),
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
    fn rightb_guid_extraction() {
        let source = r"
Sub ExtractGUID(guidData As String)
    data4 = RightB(guidData, 8)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ExtractGUID"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("guidData"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("data4"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("RightB"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("guidData"),
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
}
