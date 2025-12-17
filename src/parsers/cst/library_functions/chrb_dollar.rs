//! # `ChrB$` Function
//!
//! Returns a `String` containing the character associated with the specified ANSI character code.
//! The dollar sign suffix (`$`) explicitly indicates that this function returns a `String` type
//! (not a `Variant`), and the "B" suffix indicates this is the byte (ANSI) version.
//!
//! ## Syntax
//!
//! ```vb
//! ChrB$(charcode)
//! ```
//!
//! ## Parameters
//!
//! - **`charcode`**: Required. `Long` value that identifies a character in the ANSI character set.
//!   Valid values are 0-255. For values outside this range, an error occurs.
//!
//! ## Return Value
//!
//! Returns a `String` containing the single byte character corresponding to the specified ANSI code.
//! The return value is always a `String` type (never `Variant`), and represents a single-byte
//! character from the ANSI character set.
//!
//! ## Remarks
//!
//! - The `ChrB$` function combines the behavior of `ChrB` (byte character) with the `$` suffix
//!   (explicit `String` return type).
//! - Valid range: 0-255 (Error 5 "Invalid procedure call or argument" for values outside range).
//! - `ChrB$(0)` returns a null character.
//! - `ChrB$(13)` returns carriage return (`vbCr`).
//! - `ChrB$(10)` returns line feed (`vbLf`).
//! - `ChrB$(9)` returns tab character (`vbTab`).
//! - Values 0-31 are non-printable control characters.
//! - Values 32-126 are standard printable ASCII characters.
//! - Values 127-255 depend on the system code page (often Windows-1252 in VB6).
//! - The inverse function is `AscB`, which returns the numeric byte value of a character.
//! - For better performance when you know the result is a string, use `ChrB$` instead of `ChrB`.
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
//! 1. **Building ANSI strings** - Construct strings from byte values
//! 2. **Line breaks** - Insert carriage returns and line feeds
//! 3. **Special characters** - Add tabs, quotes, and other special characters
//! 4. **Byte-level operations** - Work with binary data or legacy file formats
//! 5. **ANSI text generation** - Create strings for systems expecting ANSI encoding
//! 6. **Legacy protocol support** - Work with older communication protocols
//! 7. **Control characters** - Generate non-printable control characters
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Get character from code
//! Dim ch As String
//! ch = ChrB$(65)  ' Returns "A"
//! ```
//!
//! ```vb
//! ' Example 2: Lowercase letter
//! Dim lower As String
//! lower = ChrB$(97)  ' Returns "a"
//! ```
//!
//! ```vb
//! ' Example 3: Special character
//! Dim space As String
//! space = ChrB$(32)  ' Returns " "
//! ```
//!
//! ```vb
//! ' Example 4: Line break
//! Dim msg As String
//! msg = "Line 1" & ChrB$(13) & ChrB$(10) & "Line 2"
//! ```
//!
//! ## Common Patterns
//!
//! ### Multi-line Strings
//! ```vb
//! Function CreateMultiLine() As String
//!     Dim result As String
//!     result = "First Line" & ChrB$(13) & ChrB$(10)
//!     result = result & "Second Line" & ChrB$(13) & ChrB$(10)
//!     result = result & "Third Line"
//!     CreateMultiLine = result
//! End Function
//! ```
//!
//! ### Tab-Separated Values
//! ```vb
//! Function CreateTSV(col1 As String, col2 As String, col3 As String) As String
//!     CreateTSV = col1 & ChrB$(9) & col2 & ChrB$(9) & col3
//! End Function
//! ```
//!
//! ### Build String from Byte Array
//! ```vb
//! Function BytesToString(bytes() As Byte) As String
//!     Dim i As Integer
//!     Dim result As String
//!     For i = LBound(bytes) To UBound(bytes)
//!         result = result & ChrB$(bytes(i))
//!     Next i
//!     BytesToString = result
//! End Function
//! ```
//!
//! ### Quote in String
//! ```vb
//! Function AddQuotes(text As String) As String
//!     AddQuotes = ChrB$(34) & text & ChrB$(34)
//! End Function
//! ```
//!
//! ### Null-Terminated String
//! ```vb
//! Function CreateNullTerminated(text As String) As String
//!     CreateNullTerminated = text & ChrB$(0)
//! End Function
//! ```
//!
//! ### ANSI Protocol Message
//! ```vb
//! Function CreateProtocolMessage(msgType As Byte, data As String) As String
//!     Dim msg As String
//!     ' SOH (Start of Header)
//!     msg = ChrB$(1)
//!     ' Message type
//!     msg = msg & ChrB$(msgType)
//!     ' STX (Start of Text)
//!     msg = msg & ChrB$(2)
//!     ' Payload
//!     msg = msg & data
//!     ' ETX (End of Text)
//!     msg = msg & ChrB$(3)
//!     CreateProtocolMessage = msg
//! End Function
//! ```
//!
//! ### Generate Alphabet
//! ```vb
//! Function GenerateAlphabet() As String
//!     Dim i As Integer
//!     Dim result As String
//!     For i = 65 To 90
//!         result = result & ChrB$(i)
//!     Next i
//!     GenerateAlphabet = result  ' Returns "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
//! End Function
//! ```
//!
//! ### CSV Field with Quotes
//! ```vb
//! Function QuoteCSVField(field As String) As String
//!     Dim quoted As String
//!     ' Replace " with ""
//!     quoted = Replace(field, ChrB$(34), ChrB$(34) & ChrB$(34))
//!     QuoteCSVField = ChrB$(34) & quoted & ChrB$(34)
//! End Function
//! ```
//!
//! ### Password Mask
//! ```vb
//! Function MaskPassword(length As Integer) As String
//!     Dim i As Integer
//!     Dim result As String
//!     For i = 1 To length
//!         result = result & ChrB$(42)  ' Asterisk
//!     Next i
//!     MaskPassword = result
//! End Function
//! ```
//!
//! ### Character Range
//! ```vb
//! Function GetPrintableChars() As String
//!     Dim i As Integer
//!     Dim result As String
//!     For i = 32 To 126
//!         result = result & ChrB$(i)
//!     Next i
//!     GetPrintableChars = result
//! End Function
//! ```
//!
//! ## Related Functions
//!
//! - `ChrB`: Returns byte character as `Variant` instead of `String`
//! - `Chr$`: Returns ANSI/Unicode character (system dependent)
//! - `ChrW$`: Returns Unicode character (2 bytes)
//! - `AscB`: Returns byte value of first byte in string (inverse of `ChrB$`)
//! - `AscB$`: Not a valid function (there is no `AscB$`)
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeChrB(code As Long) As String
//!     On Error Resume Next
//!     SafeChrB = ChrB$(code)
//!     If Err.Number <> 0 Then
//!         SafeChrB = ""
//!         Err.Clear
//!     End If
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - `ChrB$` is slightly more efficient than `ChrB` because it avoids `Variant` overhead
//! - For building strings from many bytes, consider using a buffer or byte array
//! - Concatenating many `ChrB$` calls can be slow; use arrays and `Join` for better performance
//! - When working with large amounts of byte data, consider `String` function or byte arrays
//!
//! ## Best Practices
//!
//! 1. Use named constants for common control characters instead of magic numbers
//! 2. Validate character codes are in the range 0-255 before calling `ChrB$`
//! 3. Use `ChrB$` for ANSI/byte operations, `ChrW$` for Unicode operations
//! 4. Document when using non-printable characters (codes 0-31)
//! 5. Consider code page issues when working with extended ANSI (128-255)
//! 6. Use `vbCrLf` constant instead of `ChrB$(13) & ChrB$(10)` when possible
//! 7. Prefer `ChrB$` over `ChrB` when you need a `String` result
//!
//! ## Limitations
//!
//! - Limited to character codes 0-255 (use `ChrW$` for full Unicode support)
//! - Character interpretation depends on system code page for values 128-255
//! - Does not validate that the resulting character is printable
//! - No direct support for multi-byte Unicode characters

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn chrb_dollar_simple() {
        let source = r#"
Sub Test()
    ch = ChrB$(65)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_lowercase() {
        let source = r#"
Sub Test()
    lower = ChrB$(97)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_space() {
        let source = r#"
Sub Test()
    space = ChrB$(32)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_line_break() {
        let source = r#"
Sub Test()
    msg = "Line 1" & ChrB$(13) & ChrB$(10) & "Line 2"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_multi_line_function() {
        let source = r#"
Function CreateMultiLine() As String
    Dim result As String
    result = "First Line" & ChrB$(13) & ChrB$(10)
    result = result & "Second Line"
    CreateMultiLine = result
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_tsv() {
        let source = r#"
Function CreateTSV(col1 As String, col2 As String) As String
    CreateTSV = col1 & ChrB$(9) & col2
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_byte_array() {
        let source = r#"
Function BytesToString(bytes() As Byte) As String
    Dim i As Integer
    Dim result As String
    For i = LBound(bytes) To UBound(bytes)
        result = result & ChrB$(bytes(i))
    Next i
    BytesToString = result
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_quotes() {
        let source = r#"
Function AddQuotes(text As String) As String
    AddQuotes = ChrB$(34) & text & ChrB$(34)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_null_terminated() {
        let source = r#"
Function CreateNullTerminated(text As String) As String
    CreateNullTerminated = text & ChrB$(0)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_protocol_message() {
        let source = r#"
Function CreateProtocolMessage(msgType As Byte, data As String) As String
    Dim msg As String
    msg = ChrB$(1)
    msg = msg & ChrB$(msgType)
    msg = msg & ChrB$(2)
    msg = msg & data
    msg = msg & ChrB$(3)
    CreateProtocolMessage = msg
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_alphabet() {
        let source = r#"
Function GenerateAlphabet() As String
    Dim i As Integer
    Dim result As String
    For i = 65 To 90
        result = result & ChrB$(i)
    Next i
    GenerateAlphabet = result
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_csv_field() {
        let source = r#"
Function QuoteCSVField(field As String) As String
    Dim quoted As String
    quoted = Replace(field, ChrB$(34), ChrB$(34) & ChrB$(34))
    QuoteCSVField = ChrB$(34) & quoted & ChrB$(34)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_password_mask() {
        let source = r#"
Function MaskPassword(length As Integer) As String
    Dim i As Integer
    Dim result As String
    For i = 1 To length
        result = result & ChrB$(42)
    Next i
    MaskPassword = result
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_printable_chars() {
        let source = r#"
Function GetPrintableChars() As String
    Dim i As Integer
    Dim result As String
    For i = 32 To 126
        result = result & ChrB$(i)
    Next i
    GetPrintableChars = result
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_safe_chrb() {
        let source = r#"
Function SafeChrB(code As Long) As String
    On Error Resume Next
    SafeChrB = ChrB$(code)
    If Err.Number <> 0 Then
        SafeChrB = ""
        Err.Clear
    End If
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_tab_character() {
        let source = r#"
Sub Test()
    data = "Name" & ChrB$(9) & "Age"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_multiple_calls() {
        let source = r#"
Sub Test()
    text = ChrB$(65) & ChrB$(66) & ChrB$(67)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_in_condition() {
        let source = r#"
Sub Test()
    If ch = ChrB$(32) Then
        MsgBox "Space"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_nested_functions() {
        let source = r#"
Sub Test()
    result = UCase$(ChrB$(97))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_control_chars() {
        let source = r#"
Sub Test()
    CRLF = ChrB$(13) & ChrB$(10)
    TAB = ChrB$(9)
    NULL = ChrB$(0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }

    #[test]
    fn chrb_dollar_with_variable() {
        let source = r#"
Sub Test()
    charCode = 65
    ch = ChrB$(charCode)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("ChrB$"));
    }
}
