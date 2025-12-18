//! # `ChrB` Function
//!
//! Returns a `String` containing the character associated with the specified ANSI character code.
//! The "B" suffix indicates this is the byte (ANSI) version of the `Chr` function.
//!
//! ## Syntax
//!
//! ```vb
//! ChrB(charcode)
//! ```
//!
//! ## Parameters
//!
//! - **charcode**: Required. Long. A numeric expression that identifies a character in the ANSI character set.
//!   Valid values are 0-255.
//!
//! ## Returns
//!
//! Returns a `String` containing a single byte character corresponding to the specified ANSI code.
//!
//! ## Remarks
//!
//! - `ChrB` is used to return ANSI characters (single-byte characters).
//! - The B suffix stands for "Byte", distinguishing it from the Unicode `ChrW` function.
//! - Numbers from 0 to 31 are standard, non-printable ASCII codes. For example, `ChrB(10)` returns a linefeed character.
//! - Numbers from 32 to 127 are standard printable ASCII characters.
//! - Numbers from 128 to 255 are extended ANSI characters (varies by code page).
//! - If charcode is outside the range 0-255, a runtime error occurs (Error 5: Invalid procedure call or argument).
//! - `ChrB` is particularly useful when working with byte arrays or legacy ANSI text.
//! - For Unicode characters, use `ChrW` instead of `ChrB`.
//!
//! ## Typical Uses
//!
//! 1. **Building strings with specific byte values** - Construct ANSI strings byte-by-byte
//! 2. **Creating control characters** - Generate line feeds, carriage returns, tabs, etc.
//! 3. **Low-level text manipulation** - Work with binary data or legacy file formats
//! 4. **ANSI text generation** - Create strings for systems expecting ANSI encoding
//! 5. **Byte array operations** - Convert byte values to string representations
//! 6. **Legacy protocol support** - Work with older communication protocols using ANSI
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Simple character conversion
//! Dim ch As String
//! ch = ChrB(65)  ' Returns "A"
//! ```
//!
//! ```vb
//! ' Example 2: Creating control characters
//! Dim newline As String
//! newline = ChrB(13) & ChrB(10)  ' CR+LF
//! ```
//!
//! ```vb
//! ' Example 3: Building a string from byte codes
//! Dim text As String
//! text = ChrB(72) & ChrB(101) & ChrB(108) & ChrB(108) & ChrB(111)  ' "Hello"
//! ```
//!
//! ```vb
//! ' Example 4: Display a character
//! MsgBox ChrB(65)  ' Displays "A"
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Creating line breaks
//! Dim CRLF As String
//! CRLF = ChrB(13) & ChrB(10)
//! text = "Line 1" & CRLF & "Line 2"
//! ```
//!
//! ```vb
//! ' Pattern 2: Tab-separated values
//! Dim TAB As String
//! TAB = ChrB(9)
//! data = "Name" & TAB & "Age" & TAB & "City"
//! ```
//!
//! ```vb
//! ' Pattern 3: Building ANSI strings from byte array
//! Dim i As Integer
//! Dim result As String
//! Dim bytes() As Byte
//! bytes = Array(72, 101, 108, 108, 111)
//! For i = LBound(bytes) To UBound(bytes)
//!     result = result & ChrB(bytes(i))
//! Next i
//! ```
//!
//! ```vb
//! ' Pattern 4: Creating null-terminated strings
//! Dim nullTerm As String
//! nullTerm = "Hello" & ChrB(0)
//! ```
//!
//! ```vb
//! ' Pattern 5: Character range generation
//! Dim alphabet As String
//! Dim i As Integer
//! For i = 65 To 90
//!     alphabet = alphabet & ChrB(i)
//! Next i  ' "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
//! ```
//!
//! ```vb
//! ' Pattern 6: Control character constants
//! Const NULL_CHAR As String = ChrB(0)
//! Const BELL_CHAR As String = ChrB(7)
//! Const BACKSPACE As String = ChrB(8)
//! Const TAB_CHAR As String = ChrB(9)
//! Const LF_CHAR As String = ChrB(10)
//! Const CR_CHAR As String = ChrB(13)
//! ```
//!
//! ```vb
//! ' Pattern 7: Escape special characters
//! Dim quote As String
//! quote = ChrB(34)  ' Double quote character
//! result = quote & "Hello" & quote  ' "Hello"
//! ```
//!
//! ```vb
//! ' Pattern 8: Binary data to string conversion
//! Function BytesToString(bytes() As Byte) As String
//!     Dim i As Long
//!     Dim result As String
//!     For i = LBound(bytes) To UBound(bytes)
//!         result = result & ChrB(bytes(i))
//!     Next i
//!     BytesToString = result
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 9: Creating delimited strings
//! Dim PIPE As String
//! PIPE = ChrB(124)  ' "|" character
//! data = "Field1" & PIPE & "Field2" & PIPE & "Field3"
//! ```
//!
//! ```vb
//! ' Pattern 10: ASCII art or special symbols
//! Dim box As String
//! box = ChrB(218) & String(10, ChrB(196)) & ChrB(191)  ' Top of box
//! ```
//!
//! ## Advanced Examples
//!
//! ```vb
//! ' Example 1: Complete ANSI string builder
//! Function BuildANSIString(byteCodes() As Integer) As String
//!     Dim i As Integer
//!     Dim result As String
//!     
//!     For i = LBound(byteCodes) To UBound(byteCodes)
//!         If byteCodes(i) >= 0 And byteCodes(i) <= 255 Then
//!             result = result & ChrB(byteCodes(i))
//!         Else
//!             Err.Raise 5, , "Invalid byte code: " & byteCodes(i)
//!         End If
//!     Next i
//!     
//!     BuildANSIString = result
//! End Function
//! ```
//!
//! ```vb
//! ' Example 2: Legacy file format writer
//! Sub WriteLegacyFormat(filename As String, data As String)
//!     Dim f As Integer
//!     Dim i As Integer
//!     Dim checksum As Byte
//!     
//!     f = FreeFile
//!     Open filename For Binary As #f
//!     
//!     ' Write header with STX (Start of Text)
//!     Put #f, , ChrB(2)
//!     
//!     ' Write data
//!     Put #f, , data
//!     
//!     ' Calculate and write checksum
//!     checksum = 0
//!     For i = 1 To Len(data)
//!         checksum = checksum Xor Asc(Mid(data, i, 1))
//!     Next i
//!     Put #f, , ChrB(checksum)
//!     
//!     ' Write ETX (End of Text)
//!     Put #f, , ChrB(3)
//!     
//!     Close #f
//! End Sub
//! ```
//!
//! ```vb
//! ' Example 3: Character encoding converter
//! Function ConvertToANSI(text As String) As String
//!     Dim i As Integer
//!     Dim result As String
//!     Dim charCode As Integer
//!     
//!     For i = 1 To Len(text)
//!         charCode = Asc(Mid(text, i, 1))
//!         If charCode <= 255 Then
//!             result = result & ChrB(charCode)
//!         Else
//!             result = result & ChrB(63)  ' "?" for unmappable chars
//!         End If
//!     Next i
//!     
//!     ConvertToANSI = result
//! End Function
//! ```
//!
//! ```vb
//! ' Example 4: Binary protocol message builder
//! Function CreateProtocolMessage(msgType As Byte, payload As String) As String
//!     Dim msg As String
//!     Dim length As Integer
//!     
//!     length = Len(payload)
//!     
//!     ' SOH (Start of Header)
//!     msg = ChrB(1)
//!     
//!     ' Message type
//!     msg = msg & ChrB(msgType)
//!     
//!     ' Length (2 bytes, little-endian)
//!     msg = msg & ChrB(length Mod 256)
//!     msg = msg & ChrB(length \ 256)
//!     
//!     ' STX (Start of Text)
//!     msg = msg & ChrB(2)
//!     
//!     ' Payload
//!     msg = msg & payload
//!     
//!     ' ETX (End of Text)
//!     msg = msg & ChrB(3)
//!     
//!     CreateProtocolMessage = msg
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeChrB(charcode As Long) As String
//!     On Error GoTo ErrorHandler
//!     
//!     If charcode < 0 Or charcode > 255 Then
//!         Err.Raise 5, , "Character code must be between 0 and 255"
//!     End If
//!     
//!     SafeChrB = ChrB(charcode)
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeChrB = ""
//!     MsgBox "Error converting character code " & charcode & ": " & Err.Description
//! End Function
//! ```
//!
//! ## Performance Notes
//!
//! - `ChrB` is very fast for single-byte character operations.
//! - When building large strings, consider using a byte array and converting once at the end.
//! - `String` concatenation in a loop can be slow; use `StringBuilder` pattern if available.
//! - `ChrB(0)` through `ChrB(31)` are control characters and may not display visibly.
//!
//! ## Best Practices
//!
//! 1. Use named constants for common control characters instead of magic numbers
//! 2. Validate character codes are in the range 0-255 before calling `ChrB`
//! 3. Use `ChrB` for ANSI/byte operations, `ChrW` for Unicode operations
//! 4. Document when using non-printable characters (codes 0-31)
//! 5. Consider code page issues when working with extended ANSI (128-255)
//! 6. Use `vbCrLf` constant instead of `ChrB(13) & ChrB(10)` when possible
//!
//! ## Comparison with Related Functions
//!
//! | Function | Character Set | Return Type | Use Case |
//! |----------|--------------|-------------|----------|
//! | `ChrB`     | ANSI (byte)  | `String (1 byte) | Legacy ANSI text, byte operations |
//! | `Chr`      | ANSI/Unicode | `String`      | General character conversion |
//! | `ChrW`     | Unicode      | `String (2 bytes)` | Unicode character conversion |
//! | `AscB`     | ANSI (byte)  | `Integer`     | Get ANSI code from character |
//!
//! ## Platform Notes
//!
//! - `ChrB` behavior is consistent across Windows platforms.
//! - Extended ANSI characters (128-255) may vary by system code page.
//! - In VB6, strings are internally Unicode but `ChrB` returns ANSI byte values.
//! - `ChrB` is primarily for backward compatibility with older code.
//!
//! ## Common Character Codes
//!
//! - 0: Null character
//! - 7: Bell/beep
//! - 8: Backspace
//! - 9: Tab
//! - 10: Line feed (LF)
//! - 13: Carriage return (CR)
//! - 32: Space
//! - 34: Double quote
//! - 39: Single quote/apostrophe
//! - 65-90: Uppercase A-Z
//! - 97-122: Lowercase a-z
//! - 48-57: Digits 0-9
//!
//! ## Limitations
//!
//! - Only supports ANSI character codes (0-255).
//! - Cannot represent Unicode characters beyond the ANSI range.
//! - Code page dependent for values 128-255.
//! - Runtime error if charcode is outside valid range.

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn chrb_simple_character() {
        let source = r"
Sub Test()
    ch = ChrB(65)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("IntegerLiteral"));
    }

    #[test]
    fn chrb_in_assignment() {
        let source = r"
Sub Test()
    x = ChrB(65)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("IntegerLiteral"));
    }

    #[test]
    fn chrb_control_character() {
        let source = r"
Sub Test()
    newline = ChrB(10)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("IntegerLiteral"));
    }

    #[test]
    fn chrb_concatenation() {
        let source = r"
Sub Test()
    text = ChrB(72) & ChrB(101) & ChrB(108)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrb_crlf_pattern() {
        let source = r"
Sub Test()
    CRLF = ChrB(13) & ChrB(10)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrb_tab_character() {
        let source = r"
Sub Test()
    TAB = ChrB(9)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrb_in_loop() {
        let source = r"
Sub Test()
    For i = 65 To 90
        alphabet = alphabet & ChrB(i)
    Next i
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Identifier"));
    }

    #[test]
    fn chrb_null_character() {
        let source = r#"
Sub Test()
    nullTerm = "Hello" & ChrB(0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("IntegerLiteral"));
    }

    #[test]
    fn chrb_quote_character() {
        let source = r"
Sub Test()
    quote = ChrB(34)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrb_with_variable() {
        let source = r"
Sub Test()
    charCode = 65
    ch = ChrB(charCode)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Identifier"));
    }

    #[test]
    fn chrb_in_function() {
        let source = r"
Function GetChar(code As Integer) As String
    GetChar = ChrB(code)
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Identifier"));
    }

    #[test]
    fn chrb_byte_array_conversion() {
        let source = r"
Sub Test()
    Dim bytes() As Byte
    result = result & ChrB(bytes(i))
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrb_extended_ansi() {
        let source = r"
Sub Test()
    ch = ChrB(128)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("IntegerLiteral"));
    }

    #[test]
    fn chrb_in_conditional() {
        let source = r#"
Sub Test()
    If ch = ChrB(65) Then
        MsgBox "It's A"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrb_constant_definition() {
        let source = r"
Sub Test()
    Const TAB_CHAR As String = ChrB(9)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrb_pipe_delimiter() {
        let source = r#"
Sub Test()
    PIPE = ChrB(124)
    data = "Field1" & PIPE & "Field2"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrb_range_validation() {
        let source = r"
Function SafeChrB(charcode As Long) As String
    If charcode >= 0 And charcode <= 255 Then
        SafeChrB = ChrB(charcode)
    End If
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Identifier"));
    }

    #[test]
    fn chrb_debug_print() {
        let source = r"
Sub Test()
    Debug.Print ChrB(65)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrb_string_builder() {
        let source = r"
Function BuildString() As String
    Dim result As String
    result = ChrB(72) & ChrB(101) & ChrB(108) & ChrB(108) & ChrB(111)
    BuildString = result
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn chrb_protocol_message() {
        let source = r"
Sub CreateMessage()
    msg = ChrB(1) & ChrB(msgType) & ChrB(2) & payload & ChrB(3)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier"));
    }
}
