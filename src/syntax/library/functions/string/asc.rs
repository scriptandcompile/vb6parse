//! # `Asc` Function
//!
//! Returns an `Integer` representing the character code corresponding to the first letter in a string.
//!
//! ## Syntax
//!
//! ```vb
//! Asc(string)
//! ```
//!
//! ## Parameters
//!
//! - `string` - Required. Any valid string expression. If the string contains no characters, a run-time error occurs.
//!
//! ## Return Value
//!
//! Returns an `Integer` representing the `ANSI` character code of the first character in the string.
//!
//! - For `ANSI` characters (0-127), returns standard `ASCII` values
//! - For extended `ANSI` characters (128-255), returns extended `ASCII` values
//! - Unicode characters are converted to `ANSI` before the code is returned
//!
//! ## Remarks
//!
//! The `Asc` function returns the numeric `ANSI` character code for the first character in a string.
//! This is useful for:
//! - Validating input characters
//! - Performing character-based operations
//! - Converting characters to their numeric representations
//! - Character range checking
//!
//! ### Important Notes
//!
//! 1. **Only First Character**: Only the first character of the string is examined
//! 2. **Empty String Error**: Passing an empty string results in a run-time error (Error 5: Invalid procedure call or argument)
//! 3. **ANSI vs. Unicode**: In VB6, `Asc` returns `ANSI` codes; `AscW` returns Unicode values
//! 4. **Return Type**: Returns `Integer` (16-bit signed), range -32,768 to 32,767, but character codes are 0-255
//! 5. **Case Sensitive**: Upper and lowercase letters have different codes (e.g., "A" = 65, "a" = 97)
//!
//! ### Character Code Ranges
//!
//! - **0-31**: Control characters (non-printable)
//! - **32**: Space
//! - **48-57**: Digits '0' through '9'
//! - **65-90**: Uppercase letters 'A' through 'Z'
//! - **97-122**: Lowercase letters 'a' through 'z'
//! - **128-255**: Extended ANSI characters
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! Dim code As Integer
//! code = Asc("A")          ' Returns 65
//! code = Asc("Apple")      ' Returns 65 (first character only)
//! code = Asc("a")          ' Returns 97
//! code = Asc("0")          ' Returns 48
//! code = Asc(" ")          ' Returns 32 (space)
//! ```
//!
//! ### Character Validation
//!
//! ```vb
//! Function IsDigit(ch As String) As Boolean
//!     Dim code As Integer
//!     code = Asc(ch)
//!     IsDigit = (code >= 48 And code <= 57)
//! End Function
//!
//! Function IsUpperCase(ch As String) As Boolean
//!     Dim code As Integer
//!     code = Asc(ch)
//!     IsUpperCase = (code >= 65 And code <= 90)
//! End Function
//! ```
//!
//! ### Case Conversion Offset
//!
//! ```vb
//! ' Calculate offset between upper and lower case
//! Dim offset As Integer
//! offset = Asc("a") - Asc("A")  ' Returns 32
//! ```
//!
//! ### Character Range Checking
//!
//! ```vb
//! Function IsPrintable(ch As String) As Boolean
//!     Dim code As Integer
//!     code = Asc(ch)
//!     IsPrintable = (code >= 32 And code <= 126)
//! End Function
//! ```
//!
//! ### String Encoding
//!
//! ```vb
//! Function EncodeString(s As String) As String
//!     Dim i As Integer
//!     Dim result As String
//!     For i = 1 To Len(s)
//!         If result <> "" Then result = result & ","
//!         result = result & CStr(Asc(Mid(s, i, 1)))
//!     Next i
//!     EncodeString = result
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### 1. Input Validation
//!
//! ```vb
//! If Asc(userInput) >= 48 And Asc(userInput) <= 57 Then
//!     ' First character is a digit
//! End If
//! ```
//!
//! ### 2. Alphabetic Checking
//!
//! ```vb
//! Dim code As Integer
//! code = Asc(UCase(letter))
//! If code >= 65 And code <= 90 Then
//!     ' It's a letter
//! End If
//! ```
//!
//! ### 3. CSV Parsing Helper
//!
//! ```vb
//! If Asc(field) = 34 Then  ' 34 is double quote
//!     ' Handle quoted field
//! End If
//! ```
//!
//! ### 4. Character Comparison
//!
//! ```vb
//! If Asc(char1) < Asc(char2) Then
//!     ' char1 comes before char2 in ASCII order
//! End If
//! ```
//!
//! ### 5. Special Character Detection
//!
//! ```vb
//! Select Case Asc(ch)
//!     Case 9      ' Tab
//!     Case 10     ' Line feed
//!     Case 13     ' Carriage return
//!     Case 32     ' Space
//! End Select
//! ```
//!
//! ### 6. Keyboard Input Processing
//!
//! ```vb
//! Private Sub Text1_KeyPress(KeyAscii As Integer)
//!     If KeyAscii = 13 Then  ' Enter key
//!         ' Process input
//!     End If
//! End Sub
//! ```
//!
//! ### 7. Character Class Testing
//!
//! ```vb
//! Function IsControl(ch As String) As Boolean
//!     Dim code As Integer
//!     code = Asc(ch)
//!     IsControl = (code < 32 Or code = 127)
//! End Function
//! ```
//!
//! ### 8. Simple Encryption
//!
//! ```vb
//! Function ROT13Char(ch As String) As String
//!     Dim code As Integer
//!     code = Asc(UCase(ch))
//!     If code >= 65 And code <= 90 Then
//!         code = ((code - 65 + 13) Mod 26) + 65
//!         ROT13Char = Chr(code)
//!     Else
//!         ROT13Char = ch
//!     End If
//! End Function
//! ```
//!
//! ## Common Character Codes
//!
//! | Character | Code | Description |
//! |-----------|------|-------------|
//! | Null      | 0    | Null character |
//! | Tab       | 9    | Horizontal tab |
//! | LF        | 10   | Line feed |
//! | CR        | 13   | Carriage return |
//! | Space     | 32   | Space |
//! | !         | 33   | Exclamation mark |
//! | "         | 34   | Double quote |
//! | 0         | 48   | Digit zero |
//! | 9         | 57   | Digit nine |
//! | A         | 65   | Uppercase A |
//! | Z         | 90   | Uppercase Z |
//! | a         | 97   | Lowercase a |
//! | z         | 122  | Lowercase z |
//! | DEL       | 127  | Delete |
//!
//! ## Error Handling
//!
//! ```vb
//! On Error Resume Next
//! code = Asc(inputString)
//! If Err.Number = 5 Then
//!     ' Empty string error
//!     MsgBox "String cannot be empty"
//! End If
//! ```
//!
//! ## Related Functions
//!
//! - `AscB`: Returns the first byte of a string
//! - `AscW`: Returns the Unicode character code
//! - `Chr`: Returns the character for a given character code (inverse of Asc)
//! - `ChrB`: Returns a byte containing the character
//! - `ChrW`: Returns a Unicode character
//! - `InStr`: Finds the position of a character in a string
//! - `Mid`: Extracts a substring
//! - `Left`: Gets leftmost characters
//! - `Right`: Gets rightmost characters
//!
//! ## Performance Notes
//!
//! - Asc is a very fast operation (direct character code lookup)
//! - More efficient than string comparison for single character checks
//! - Use Asc for character-based validation instead of multiple string comparisons
//! - In tight loops, cache Asc results if checking the same character repeatedly
//!
//! ## Parsing Notes
//!
//! The `Asc` function is not a reserved keyword in VB6. It is parsed as a regular
//! function call (`CallExpression`). This module exists primarily for documentation
//! purposes and to provide a comprehensive test suite that validates the parser
//! correctly handles `Asc` function calls in various contexts.

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn asc_simple() {
        let source = r#"
Sub Test()
    code = Asc("A")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_variable() {
        let source = r"
Sub Test()
    result = Asc(userInput)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_in_if_statement() {
        let source = r"
Sub Test()
    If Asc(ch) >= 65 And Asc(ch) <= 90 Then
        valid = True
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_in_select_case() {
        let source = r"
Sub Test()
    Select Case Asc(key)
        Case 13
            ProcessEnter
        Case 27
            ProcessEscape
    End Select
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_in_for_loop() {
        let source = r#"
Sub Test()
    For i = Asc("A") To Asc("Z")
        codes(i) = i
    Next i
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_in_do_loop() {
        let source = r"
Sub Test()
    Do While Asc(buffer) <> 0
        ProcessChar buffer
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_in_while_wend() {
        let source = r"
Sub Test()
    While Asc(ch) <> 13
        ReadNext
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_with_function_call() {
        let source = r"
Sub Test()
    code = Asc(GetFirstChar())
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_with_mid_function() {
        let source = r"
Sub Test()
    charCode = Asc(Mid(text, pos, 1))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_with_property_access() {
        let source = r"
Sub Test()
    value = Asc(Me.TextBox.Text)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_case_insensitive() {
        let source = r#"
Sub Test()
    x = ASC("test")
    y = asc("test")
    z = AsC("test")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_with_line_continuation() {
        let source = r#"
Sub Test()
    result = Asc _
        ("Hello")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_with_whitespace() {
        let source = r#"
Sub Test()
    code = Asc  (  "X"  )
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_in_arithmetic() {
        let source = r#"
Sub Test()
    offset = Asc("a") - Asc("A")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_in_with_block() {
        let source = r"
Sub Test()
    With textData
        code = Asc(.Value)
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_in_array_assignment() {
        let source = r"
Sub Test()
    charCodes(i) = Asc(chars(i))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_in_comparison_chain() {
        let source = r"
Sub Test()
    valid = Asc(ch) >= 48 And Asc(ch) <= 57
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_concatenated_string() {
        let source = r#"
Sub Test()
    code = Asc("Hello" & " World")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_special_characters() {
        let source = r"
Sub Test()
    tab = Asc(vbTab)
    cr = Asc(vbCr)
    lf = Asc(vbLf)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_in_function() {
        let source = r"
Function IsDigit(ch As String) As Boolean
    IsDigit = (Asc(ch) >= 48 And Asc(ch) <= 57)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_in_sub() {
        let source = r"
Sub ProcessKey(key As String)
    If Asc(key) = 13 Then Exit Sub
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_multiple_on_line() {
        let source = r#"
Sub Test()
    a = Asc("A"): b = Asc("B"): c = Asc("C")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_with_ucase() {
        let source = r"
Sub Test()
    code = Asc(UCase(letter))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_in_print_statement() {
        let source = r#"
Sub Test()
    Print Asc("@")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn asc_module_level() {
        let source = r#"Const CHAR_A As Integer = Asc("A")"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/asc");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
