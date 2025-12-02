//! # `LTrim$` Function
//!
//! Returns a `String` containing a copy of a specified string with leading spaces removed.
//!
//! ## Syntax
//!
//! ```vb6
//! LTrim$(string)
//! ```
//!
//! ## Parameters
//!
//! - `string`: Required. Any valid string expression. If `string` contains `Null`, `Null` is returned.
//!
//! ## Return Value
//!
//! Returns a `String` with all leading (left-side) spaces removed. If the string contains only spaces, returns an empty string. Trailing spaces and spaces within the string are preserved.
//!
//! ## Remarks
//!
//! The `LTrim$` function removes leading space characters (ASCII 32) from a string. It's commonly used to clean up user input, process fixed-width data, or normalize strings for comparison.
//!
//! Only the space character (ASCII 32) is removed. Other whitespace characters like tabs, newlines, or non-breaking spaces are not affected by `LTrim$`.
//!
//! `LTrim$` is the string-specific version that always returns a `String`. The `LTrim` function returns a `Variant`.
//!
//! ## Typical Uses
//!
//! ### Example 1: Cleaning User Input
//! ```vb6
//! Dim userInput As String
//! userInput = "  John Doe"
//! cleaned = LTrim$(userInput)  ' "John Doe"
//! ```
//!
//! ### Example 2: Processing Fixed-Width Data
//! ```vb6
//! Dim record As String
//! record = "     12345"
//! id = LTrim$(record)  ' "12345"
//! ```
//!
//! ### Example 3: Normalizing Comparison
//! ```vb6
//! If LTrim$(text1) = LTrim$(text2) Then
//!     Debug.Print "Match (ignoring leading spaces)"
//! End If
//! ```
//!
//! ### Example 4: Parsing Indented Text
//! ```vb6
//! Dim line As String
//! line = "    Code line"
//! code = LTrim$(line)  ' "Code line"
//! ```
//!
//! ## Common Usage Patterns
//!
//! ### Validating Non-Empty Input
//! ```vb6
//! Dim name As String
//! name = txtName.Text
//! If LTrim$(name) = "" Then
//!     MsgBox "Name cannot be empty or spaces only"
//! End If
//! ```
//!
//! ### Processing CSV Fields
//! ```vb6
//! Dim fields() As String
//! fields = Split(csvLine, ",")
//! For i = 0 To UBound(fields)
//!     fields(i) = LTrim$(fields(i))
//! Next i
//! ```
//!
//! ### Reading Indented Configuration
//! ```vb6
//! Dim configLine As String
//! configLine = "    setting=value"
//! setting = LTrim$(configLine)  ' "setting=value"
//! ```
//!
//! ### Extracting List Items
//! ```vb6
//! Dim listItem As String
//! listItem = "  - Item text"
//! text = LTrim$(listItem)  ' "- Item text"
//! ```
//!
//! ### Database Field Cleanup
//! ```vb6
//! Dim dbValue As String
//! dbValue = rs.Fields("name").Value
//! cleanValue = LTrim$(dbValue)
//! ```
//!
//! ### Removing Formatting Spaces
//! ```vb6
//! Dim formatted As String
//! formatted = "     $1,234.56"
//! amount = LTrim$(formatted)  ' "$1,234.56"
//! ```
//!
//! ### Processing Text File Lines
//! ```vb6
//! Dim line As String
//! Open "data.txt" For Input As #1
//! Do Until EOF(1)
//!     Line Input #1, line
//!     line = LTrim$(line)
//!     If Left$(line, 1) <> "#" Then
//!         processLine line
//!     End If
//! Loop
//! Close #1
//! ```
//!
//! ### Normalizing String Arrays
//! ```vb6
//! Dim items() As String
//! items = Split(data, vbCrLf)
//! For i = 0 To UBound(items)
//!     items(i) = LTrim$(items(i))
//! Next i
//! ```
//!
//! ### Removing Padding from Fixed Fields
//! ```vb6
//! Dim fixedRecord As String
//! fixedRecord = "          Customer Name     "
//! name = LTrim$(fixedRecord)  ' "Customer Name     "
//! ```
//!
//! ### Combining with RTrim$ for Full Trim
//! ```vb6
//! Dim text As String
//! text = "  Data  "
//! ' Remove both leading and trailing spaces
//! cleaned = LTrim$(RTrim$(text))  ' "Data"
//! ' Or use Trim$ directly
//! cleaned = Trim$(text)  ' "Data"
//! ```
//!
//! ## Related Functions
//!
//! - `LTrim`: Variant version that returns a `Variant`
//! - `RTrim$`: Removes trailing spaces from a string
//! - `Trim$`: Removes both leading and trailing spaces
//! - `Left$`: Returns characters from the left side of a string
//! - `Len`: Returns the length of a string
//! - `Replace`: Replaces occurrences of a substring
//!
//! ## Best Practices
//!
//! 1. Use `Trim$` instead of `LTrim$` when you want to remove both leading and trailing spaces
//! 2. Always validate input after trimming to check for empty strings
//! 3. Be aware that only space characters (ASCII 32) are removed, not tabs or other whitespace
//! 4. Use `LTrim$` for left-aligned fixed-width fields
//! 5. Combine with validation to prevent injection attacks in SQL or scripts
//! 6. Remember that `LTrim$` preserves internal and trailing spaces
//! 7. Consider using `Replace` for removing other whitespace characters
//! 8. Cache the result if using trimmed value multiple times
//! 9. Use `LTrim$` instead of `LTrim` when you need explicit `String` type
//! 10. Test with edge cases: empty strings, all spaces, no leading spaces
//!
//! ## Performance Considerations
//!
//! - `LTrim$` is a fast operation in VB6
//! - No performance penalty if the string has no leading spaces
//! - More efficient than using `Replace` or manual character removal
//! - Minimal memory allocation if few spaces are removed
//! - Consider caching trimmed values in loops for better performance
//!
//! ## Whitespace Handling
//!
//! | Character | ASCII | Removed by LTrim$ |
//! |-----------|-------|-------------------|
//! | Space | 32 | Yes (if leading) |
//! | Tab | 9 | No |
//! | Newline | 10 | No |
//! | Carriage Return | 13 | No |
//! | Non-breaking Space | 160 | No |
//! | Vertical Tab | 11 | No |
//! | Form Feed | 12 | No |
//!
//! ## Common Pitfalls
//!
//! - Assuming `LTrim$` removes all whitespace characters (it only removes spaces)
//! - Not checking for empty string after trimming spaces-only input
//! - Using `LTrim$` when `Trim$` would be more appropriate
//! - Forgetting that trailing spaces are preserved
//! - Not handling `Null` string values (causes runtime error)
//! - Assuming trimmed string is never empty
//! - Using repeatedly in loops without caching result
//! - Expecting tabs or newlines to be removed
//!
//! ## Limitations
//!
//! - Only removes space characters (ASCII 32), not other whitespace
//! - Cannot specify which characters to remove
//! - Does not remove trailing spaces (use `RTrim$` or `Trim$`)
//! - Returns `Null` if the string argument is `Null`
//! - Cannot remove spaces from the middle of strings
//! - No option to limit how many spaces are removed
//! - Does not normalize multiple internal spaces to single spaces

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn ltrim_dollar_simple() {
        let source = r#"
Sub Main()
    result = LTrim$("  Hello")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_assignment() {
        let source = r#"
Sub Main()
    Dim cleaned As String
    cleaned = LTrim$(userInput)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_variable() {
        let source = r#"
Sub Main()
    text = "  Data"
    trimmed = LTrim$(text)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_in_condition() {
        let source = r#"
Sub Main()
    If LTrim$(name) = "" Then
        MsgBox "Name cannot be empty"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_comparison() {
        let source = r#"
Sub Main()
    If LTrim$(text1) = LTrim$(text2) Then
        Debug.Print "Match"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_user_input() {
        let source = r#"
Sub Main()
    userInput = txtName.Text
    cleaned = LTrim$(userInput)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_fixed_width() {
        let source = r#"
Sub Main()
    record = "     12345"
    id = LTrim$(record)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_in_loop() {
        let source = r#"
Sub Main()
    For i = 0 To UBound(fields)
        fields(i) = LTrim$(fields(i))
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_with_left() {
        let source = r##"
Sub Main()
    textLine = LTrim$(configLine)
    firstChar = Left$(textLine, 1)
End Sub
"##;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_database_field() {
        let source = r#"
Sub Main()
    dbValue = rs.Fields("name").Value
    cleanValue = LTrim$(dbValue)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_select_case() {
        let source = r#"
Sub Main()
    cmd = LTrim$(commandLine)
    Select Case cmd
        Case "START"
            StartProcess
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_multiple_uses() {
        let source = r#"
Sub Main()
    first = LTrim$(field1)
    second = LTrim$(field2)
    combined = first & second
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_in_function() {
        let source = r#"
Function CleanText(text As String) As String
    CleanText = LTrim$(text)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_with_rtrim() {
        let source = r#"
Sub Main()
    text = "  Data  "
    cleaned = LTrim$(RTrim$(text))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_validation() {
        let source = r#"
Sub Main()
    name = txtName.Text
    If LTrim$(name) = "" Then
        valid = False
    Else
        valid = True
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_concatenation() {
        let source = r#"
Sub Main()
    result = "Prefix: " & LTrim$(data)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_file_processing() {
        let source = r#"
Sub Main()
    Line Input #1, dataLine
    dataLine = LTrim$(dataLine)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_with_len() {
        let source = r#"
Sub Main()
    trimmed = LTrim$(text)
    length = Len(trimmed)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_textbox_input() {
        let source = r#"
Sub Main()
    Dim input As String
    input = LTrim$(Text1.Text)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }

    #[test]
    fn ltrim_dollar_split_result() {
        let source = r#"
Sub Main()
    fields = Split(csvLine, ",")
    For i = 0 To UBound(fields)
        fields(i) = LTrim$(fields(i))
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("LTrim$"));
    }
}
