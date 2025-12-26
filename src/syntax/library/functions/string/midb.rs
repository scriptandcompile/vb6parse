//! # `MidB` Function
//!
//! The `MidB` function returns a `Variant` (`String`) containing a specified number of bytes from a string.
//! This function operates on byte positions rather than character positions, which is important when working
//! with ANSI strings or when you need byte-level control over string manipulation.
//!
//! ## Syntax
//!
//! ```vb
//! MidB(string, start[, length])
//! ```
//!
//! ## Parameters
//!
//! - `string` - Required. `String` expression from which bytes are returned.
//! - `start` - Required. `Long`. The `Byte` position in string at which the part to be taken begins (1-based).
//! - `length` - Optional. `Long`. Number of bytes to return. If omitted or if there are fewer than `length` bytes in the text (including the byte at `start`), all bytes from the start position to the end of the string are returned.
//!
//! ## Return Value
//!
//! Returns a `Variant` (`String`) containing the specified byte sequence from the input string.
//!
//! ## Behavior
//!
//! - If `start` is greater than the number of bytes in `string`, `MidB` returns a zero-length string ("").
//! - If `start` is less than 1, a runtime error occurs.
//! - If `length` is negative, a runtime error occurs.
//! - The first byte in the string is at position 1.
//! - When working with DBCS (Double-Byte Character Set) strings, `MidB` can split multi-byte characters if not used carefully.
//!
//! ## Difference from Mid
//!
//! The `MidB` function operates on byte positions, while the `Mid` function operates on character positions.
//! For single-byte character sets (like ASCII), they behave identically. For multi-byte character sets
//! (like Unicode or DBCS), `MidB` provides byte-level access which can be useful for binary data manipulation
//! or low-level string operations.
//!
//! ## Examples
//!
//! ```vb
//! ' Extract 3 bytes starting at byte position 2
//! Dim result As Variant
//! result = MidB("Hello World", 2, 3)  ' Returns "ell"
//!
//! ' Extract from byte position 7 to the end
//! result = MidB("Hello World", 7)  ' Returns "World"
//!
//! ' Start position beyond string length
//! result = MidB("Hi", 10)  ' Returns ""
//! ```

#[cfg(test)]
mod tests {
    use crate::parsers::cst::ConcreteSyntaxTree;

    #[test]
    fn midb_simple() {
        let source = r#"
Sub Test()
    Dim result As String
    result = MidB("Hello", 2, 3)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_without_length() {
        let source = r#"
Sub Test()
    Dim result As String
    result = MidB("VB6 Programming", 5)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_with_variables() {
        let source = r"
Sub Test()
    Dim text As String
    Dim pos As Long
    Dim len As Long
    Dim result As String
    result = MidB(text, pos, len)
End Sub
";

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_nested_function() {
        let source = r#"
Sub Test()
    Dim result As String
    result = MidB(UCase$("hello world"), 1, 5)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_concatenation() {
        let source = r#"
Sub Test()
    Dim result As String
    result = "Start: " & MidB("ABCDEFGH", 3, 4) & " :End"
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_comparison() {
        let source = r#"
Sub Test()
    Dim isEqual As Boolean
    isEqual = (MidB("Testing", 2, 3) = "est")
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_if_statement() {
        let source = r#"
Sub Test()
    Dim text As String
    If MidB(text, 1, 4) = "http" Then
        ' Do something
    End If
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_byte_array_processing() {
        let source = r"
Sub Test()
    Dim data As String
    Dim chunk As String
    chunk = MidB(data, 1, 256)
End Sub
";

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_binary_data() {
        let source = r"
Sub Test()
    Dim buffer As String
    Dim header As String
    Dim payload As String
    header = MidB(buffer, 1, 16)
    payload = MidB(buffer, 17)
End Sub
";

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_array_element() {
        let source = r"
Sub Test()
    Dim arr(10) As String
    Dim result As String
    result = MidB(arr(5), 2, 3)
End Sub
";

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_loop() {
        let source = r"
Sub Test()
    Dim i As Long
    Dim text As String
    Dim bytes As String
    For i = 1 To LenB(text) Step 10
        bytes = MidB(text, i, 10)
    Next i
End Sub
";

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_multiple_calls() {
        let source = r"
Sub Test()
    Dim s As String
    Dim result As String
    result = MidB(s, 1, 2) & MidB(s, 5, 3)
End Sub
";

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_function_parameter() {
        let source = r"
Function ProcessBytes(text As String) As String
    ProcessBytes = MidB(text, 1, 10)
End Function
";

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_conversion_combination() {
        let source = r"
Sub Test()
    Dim text As String
    Dim result As String
    result = CStr(MidB(text, 5, 10))
End Sub
";

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_case_statement() {
        let source = r#"
Sub Test()
    Dim data As String
    Dim magic As String
    magic = MidB(data, 1, 4)
    Select Case magic
        Case "RIFF"
            ' WAV file
        Case "PNG"
            ' PNG file
    End Select
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_with_expressions() {
        let source = r"
Sub Test()
    Dim text As String
    Dim offset As Long
    Dim result As String
    result = MidB(text, offset * 2 + 1, 5)
End Sub
";

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_file_header() {
        let source = r"
Sub Test()
    Dim fileData As String
    Dim signature As String
    Dim version As String
    signature = MidB(fileData, 1, 8)
    version = MidB(fileData, 9, 4)
End Sub
";

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_constant_extraction() {
        let source = r#"
Sub Test()
    Const BINARY_DATA As String = "ABCDEF123456"
    Dim part1 As String
    Dim part2 As String
    part1 = MidB(BINARY_DATA, 1, 6)
    part2 = MidB(BINARY_DATA, 7, 6)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_type_field() {
        let source = r"
Type BinaryRecord
    Data As String
    ID As String
End Type

Sub Test()
    Dim rec As BinaryRecord
    Dim bytes As String
    bytes = MidB(rec.Data, 1, 5)
End Sub
";

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }

    #[test]
    fn midb_return_value() {
        let source = r"
Function GetBytes(s As String, pos As Long) As String
    GetBytes = MidB(s, pos, 3)
End Function
";

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("MidBKeyword"));
    }
}
