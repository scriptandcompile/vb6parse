//! # LeftB Function
//!
//! Returns a Variant (String) containing a specified number of bytes from the left side of a string.
//!
//! ## Syntax
//!
//! ```vb
//! LeftB(string, length)
//! ```
//!
//! ## Parameters
//!
//! - `string` (Required): String expression from which leftmost bytes are returned
//!   - If string contains Null, Null is returned
//! - `length` (Required): Numeric expression indicating how many bytes to return
//!   - If 0, empty string ("") is returned
//!   - If greater than or equal to number of bytes in string, entire string is returned
//!   - Must be non-negative (negative values cause error)
//!
//! ## Return Value
//!
//! Returns a Variant containing a String:
//! - Contains the specified number of bytes from the left side of the string
//! - Returns empty string if length is 0
//! - Returns entire string if length >= LenB(string)
//! - Returns Null if string argument is Null
//! - Returns Variant type (LeftB$ variant returns String type directly)
//!
//! ## Remarks
//!
//! The LeftB function extracts bytes from the beginning of a string:
//!
//! - Returns leftmost bytes up to specified length
//! - Operates on byte level, not character level
//! - Particularly useful with double-byte character sets (DBCS)
//! - Complements RightB function (which returns rightmost bytes)
//! - Works with MidB function for complete byte-level substring extraction
//! - Zero-based extraction: LeftB("ABC", 2) returns first 2 bytes
//! - Safe with lengths exceeding string byte length (returns full string)
//! - Null propagates through the function
//! - Negative length raises Error 5 (Invalid procedure call or argument)
//! - Common for extracting binary data, protocol headers, file signatures
//! - More efficient than MidB(string, 1, length) for left extraction
//! - LeftB$ variant returns String type (not Variant) for slight performance gain
//! - Cannot extract from right side (use RightB for that)
//! - Cannot skip bytes (use MidB for that)
//! - Does not modify original string (strings are immutable)
//!
//! ## Differences from Left Function
//!
//! - Left operates on characters, LeftB operates on bytes
//! - In single-byte character sets (SBCS), they are equivalent
//! - In double-byte character sets (DBCS), one character may be multiple bytes
//! - LeftB is essential for binary data manipulation
//! - LeftB is used with LenB (byte length) rather than Len (character length)
//!
//! ## Typical Uses
//!
//! 1. **Binary Data**: Extract bytes from binary strings
//! 2. **Protocol Headers**: Parse network protocol headers
//! 3. **File Signatures**: Identify file types by magic bytes
//! 4. **DBCS Strings**: Work with Japanese, Chinese, Korean text at byte level
//! 5. **Fixed Byte Records**: Parse fixed-width binary records
//! 6. **Byte Validation**: Check byte prefixes in data
//! 7. **Binary Structures**: Extract fields from binary structures
//! 8. **Network Data**: Process raw network packet data
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Basic byte extraction
//! Dim data As String
//! data = Chr$(65) & Chr$(66) & Chr$(67)  ' "ABC"
//!
//! Debug.Print LeftB(data, 2)            ' First 2 bytes
//! Debug.Print LeftB(data, 1)            ' First byte
//!
//! ' Example 2: File signature checking
//! Dim fileData As String
//! Open "test.exe" For Binary As #1
//! fileData = Input$(2, #1)
//! Close #1
//!
//! If LeftB(fileData, 2) = "MZ" Then
//!     Debug.Print "DOS/Windows executable"
//! End If
//!
//! ' Example 3: Protocol header extraction
//! Dim packet As String
//! packet = ReceiveNetworkData()
//!
//! Dim header As String
//! header = LeftB(packet, 16)  ' 16-byte header
//!
//! ' Example 4: DBCS text handling
//! Dim japaneseText As String
//! japaneseText = LoadJapaneseText()  ' Load Japanese text
//!
//! Dim firstBytes As String
//! firstBytes = LeftB(japaneseText, 4)  ' First 4 bytes (may be 2 DBCS chars)
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Extract binary header
//! Function GetBinaryHeader(data As String, headerSize As Long) As String
//!     If LenB(data) < headerSize Then
//!         GetBinaryHeader = data
//!     Else
//!         GetBinaryHeader = LeftB(data, headerSize)
//!     End If
//! End Function
//!
//! ' Pattern 2: Validate magic bytes
//! Function ValidateMagicBytes(data As String, magic As String) As Boolean
//!     ValidateMagicBytes = (LeftB(data, LenB(magic)) = magic)
//! End Function
//!
//! ' Pattern 3: Extract record ID
//! Function GetRecordID(record As String) As String
//!     ' First 8 bytes contain record ID
//!     GetRecordID = LeftB(record, 8)
//! End Function
//!
//! ' Pattern 4: Parse network packet
//! Sub ParsePacket(packet As String)
//!     Dim header As String
//!     Dim payload As String
//!     
//!     header = LeftB(packet, 20)  ' 20-byte header
//!     payload = MidB(packet, 21)  ' Remaining bytes
//!     
//!     ' Process header and payload
//! End Sub
//! ```
//!
//! ## See Also
//!
//! - `LeftB$`: String-returning variant of LeftB
//! - `RightB`: Returns rightmost bytes from string
//! - `RightB$`: String-returning variant of RightB
//! - `MidB`: Returns bytes from middle of string
//! - `MidB$`: String-returning variant of MidB
//! - `LenB`: Returns byte length of string
//! - `Left`: Character-based left extraction
//! - `Right`: Character-based right extraction
//! - `Mid`: Character-based middle extraction

#[cfg(test)]
mod tests {
    type ConcreteSyntaxTree = crate::parsers::ConcreteSyntaxTree;

    #[test]
    fn leftb_basic() {
        let source = r"
Sub Test()
    result = LeftB(data, 4)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LeftB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn leftb_file_signature() {
        let source = r#"
Function IsExecutable(fileName As String) As Boolean
    Dim fileData As String
    Open fileName For Binary As #1
    fileData = Input$(2, #1)
    Close #1
    IsExecutable = (LeftB(fileData, 2) = "MZ")
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LeftB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn leftb_protocol_header() {
        let source = r"
Sub ParsePacket(packet As String)
    header = LeftB(packet, 16)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LeftB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn leftb_in_condition() {
        let source = r"
Sub Test()
    If LeftB(data, 4) = magicBytes Then
        ProcessData
    End If
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LeftB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn leftb_binary_record() {
        let source = r"
Function GetRecordID(record As String) As String
    GetRecordID = LeftB(record, 8)
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LeftB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn leftb_with_lenb() {
        let source = r"
Sub Test()
    If LenB(data) > 10 Then
        prefix = LeftB(data, 10)
    End If
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LeftB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn leftb_dbcs_handling() {
        let source = r"
Sub Test()
    Dim jpText As String
    jpText = GetJapaneseText()
    bytes = LeftB(jpText, 4)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LeftB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn leftb_validation() {
        let source = r"
Function ValidateMagicBytes(data As String, magic As String) As Boolean
    ValidateMagicBytes = (LeftB(data, LenB(magic)) = magic)
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("LeftB"));
        assert!(text.contains("Identifier"));
    }
}
