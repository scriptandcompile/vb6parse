//! # `Mid$` Function
//!
//! The `Mid$` function returns a `String` containing a specified number of characters from a string.
//! The dollar sign suffix (`$`) indicates that this function always returns a `String` type, never a `Variant`.
//!
//! ## Syntax
//!
//! ```vb
//! Mid$(string, start[, length])
//! ```
//!
//! ## Parameters
//!
//! - `string` - Required. `String` expression from which characters are returned.
//! - `start` - Required. `Long`. Character position in `string` at which the part to be taken begins (1-based).
//! - `length` - Optional. `Long`. Number of characters to return. If omitted or if there are fewer than `length` characters in the text (including the character at `start`), all characters from the start position to the end of the string are returned.
//!
//! ## Return Value
//!
//! Returns a `String` containing the specified portion of the input string.
//!
//! ## Behavior
//!
//! - If `start` is greater than the number of characters in `string`, `Mid$` returns a zero-length string ("").
//! - If `start` is less than 1, a runtime error occurs.
//! - If `length` is negative, a runtime error occurs.
//! - The first character in the string is at position 1.
//!
//! ## Difference from Mid
//!
//! The `Mid$` function always returns a `String`, while the `Mid` function (without the dollar sign) can return a `Variant`.
//! In practice, they behave identically in most scenarios, but the dollar sign version may be slightly more efficient
//! as it avoids the overhead of the `Variant` type.
//!
//! ## Examples
//!
//! ```vb
//! ' Extract 3 characters starting at position 2
//! Dim result As String
//! result = Mid$("Hello World", 2, 3)  ' Returns "ell"
//!
//! ' Extract from position 7 to the end
//! result = Mid$("Hello World", 7)  ' Returns "World"
//!
//! ' Start position beyond string length
//! result = Mid$("Hi", 10)  ' Returns ""
//! ```

#[cfg(test)]
mod tests {
    use crate::parsers::cst::ConcreteSyntaxTree;

    #[test]
    fn mid_dollar_simple() {
        let source = r#"
Sub Test()
    Dim result As String
    result = Mid$("Hello", 2, 3)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_without_length() {
        let source = r#"
Sub Test()
    Dim result As String
    result = Mid$("VB6 Programming", 5)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_with_variables() {
        let source = r#"
Sub Test()
    Dim text As String
    Dim pos As Long
    Dim len As Long
    Dim result As String
    result = Mid$(text, pos, len)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_nested_function() {
        let source = r#"
Sub Test()
    Dim result As String
    result = Mid$(UCase$("hello world"), 1, 5)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_concatenation() {
        let source = r#"
Sub Test()
    Dim result As String
    result = "Start: " & Mid$("ABCDEFGH", 3, 4) & " :End"
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_comparison() {
        let source = r#"
Sub Test()
    Dim isEqual As Boolean
    isEqual = (Mid$("Testing", 2, 3) = "est")
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_if_statement() {
        let source = r#"
Sub Test()
    Dim text As String
    If Mid$(text, 1, 4) = "http" Then
        ' Do something
    End If
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_file_extension() {
        let source = r#"
Sub Test()
    Dim filename As String
    Dim ext As String
    ext = Mid$(filename, InStrRev(filename, ".") + 1)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_substring_extract() {
        let source = r#"
Sub Test()
    Dim data As String
    Dim year As String
    Dim month As String
    Dim day As String
    year = Mid$(data, 1, 4)
    month = Mid$(data, 6, 2)
    day = Mid$(data, 9, 2)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_array_element() {
        let source = r#"
Sub Test()
    Dim arr(10) As String
    Dim result As String
    result = Mid$(arr(5), 2, 3)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_loop() {
        let source = r#"
Sub Test()
    Dim i As Long
    Dim text As String
    Dim char As String
    For i = 1 To Len(text)
        char = Mid$(text, i, 1)
    Next i
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_multiple_calls() {
        let source = r#"
Sub Test()
    Dim s As String
    Dim result As String
    result = Mid$(s, 1, 2) & Mid$(s, 5, 3)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_function_parameter() {
        let source = r#"
Function ProcessString(text As String) As String
    ProcessString = Mid$(text, 1, 10)
End Function
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_trim_combination() {
        let source = r#"
Sub Test()
    Dim text As String
    Dim result As String
    result = LTrim$(RTrim$(Mid$(text, 5, 10)))
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_case_statement() {
        let source = r#"
Sub Test()
    Dim code As String
    Dim prefix As String
    prefix = Mid$(code, 1, 3)
    Select Case prefix
        Case "ABC"
            ' Do something
        Case "XYZ"
            ' Do something else
    End Select
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_with_expressions() {
        let source = r#"
Sub Test()
    Dim text As String
    Dim offset As Long
    Dim result As String
    result = Mid$(text, offset + 1, 5)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_tokenizer() {
        let source = r#"
Sub Test()
    Dim line As String
    Dim tokens() As String
    Dim i As Long
    Dim token As String
    For i = 1 To Len(line) Step 10
        token = Mid$(line, i, 10)
    Next i
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_constant_extraction() {
        let source = r#"
Sub Test()
    Const DATA_STRING As String = "ABCDEF123456"
    Dim alpha As String
    Dim numeric As String
    alpha = Mid$(DATA_STRING, 1, 6)
    numeric = Mid$(DATA_STRING, 7, 6)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_type_field() {
        let source = r#"
Type Person
    Name As String
    ID As String
End Type

Sub Test()
    Dim p As Person
    Dim shortID As String
    shortID = Mid$(p.ID, 1, 5)
End Sub
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }

    #[test]
    fn mid_dollar_return_value() {
        let source = r#"
Function GetSubstring(s As String, pos As Long) As String
    GetSubstring = Mid$(s, pos, 3)
End Function
"#;

        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();

        assert!(debug.contains("Identifier") && debug.contains("Mid$"));
    }
}
