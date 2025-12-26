//! # `Right$` Function
//!
//! The `Right$` function in Visual Basic 6 returns a string containing a specified number of
//! characters from the right side (end) of a string. The dollar sign (`$`) suffix indicates
//! that this function always returns a `String` type, never a `Variant`.
//!
//! ## Syntax
//!
//! ```vb6
//! Right$(string, length)
//! ```
//!
//! ## Parameters
//!
//! - `string` - Required. String expression from which the rightmost characters are returned.
//!   If `string` contains `Null`, `Null` is returned.
//! - `length` - Required. Numeric expression indicating how many characters to return. If 0,
//!   a zero-length string ("") is returned. If greater than or equal to the number of characters
//!   in `string`, the entire string is returned.
//!
//! ## Return Value
//!
//! Returns a `String` containing the rightmost `length` characters of `string`.
//!
//! ## Behavior and Characteristics
//!
//! ### Length Handling
//!
//! - If `length` = 0: Returns an empty string ("")
//! - If `length` >= `Len(string)`: Returns the entire string
//! - If `length` < 0: Generates a runtime error (Invalid procedure call or argument)
//! - If `string` is empty (""): Returns an empty string regardless of `length`
//!
//! ### Null Handling
//!
//! - If `string` contains `Null`: Returns `Null`
//! - If `length` is `Null`: Generates a runtime error (Invalid use of Null)
//!
//! ### Type Differences: `Right$` vs `Right`
//!
//! - `Right$`: Always returns `String` type (never `Variant`)
//! - `Right`: Returns `Variant` (can propagate `Null` values)
//! - Use `Right$` when you need guaranteed `String` return type
//! - Use `Right` when working with potentially `Null` values
//!
//! ## Common Usage Patterns
//!
//! ### 1. Extract File Extension
//!
//! ```vb6
//! Function GetExtension(fileName As String) As String
//!     Dim dotPos As Integer
//!     dotPos = InStrRev(fileName, ".")
//!     If dotPos > 0 Then
//!         GetExtension = Right$(fileName, Len(fileName) - dotPos)
//!     Else
//!         GetExtension = ""
//!     End If
//! End Function
//!
//! Dim ext As String
//! ext = GetExtension("document.txt")  ' Returns "txt"
//! ```
//!
//! ### 2. Get Last N Characters
//!
//! ```vb6
//! Dim text As String
//! Dim suffix As String
//! text = "Hello World"
//! suffix = Right$(text, 5)  ' Returns "World"
//! ```
//!
//! ### 3. Extract Account Number Suffix
//!
//! ```vb6
//! Function GetAccountSuffix(accountNum As String) As String
//!     ' Get last 4 digits of account number
//!     GetAccountSuffix = Right$(accountNum, 4)
//! End Function
//!
//! Dim lastFour As String
//! lastFour = GetAccountSuffix("1234567890")  ' Returns "7890"
//! ```
//!
//! ### 4. Pad String to Fixed Width
//!
//! ```vb6
//! Function PadLeft(text As String, width As Integer) As String
//!     Dim padded As String
//!     padded = Space(width) & text
//!     PadLeft = Right$(padded, width)
//! End Function
//!
//! Dim result As String
//! result = PadLeft("42", 5)  ' Returns "   42"
//! ```
//!
//! ### 5. Extract Trailing Digits
//!
//! ```vb6
//! Function GetTrailingNumber(text As String) As String
//!     Dim i As Integer
//!     Dim numChars As Integer
//!     For i = Len(text) To 1 Step -1
//!         If Not IsNumeric(Mid$(text, i, 1)) Then Exit For
//!         numChars = numChars + 1
//!     Next i
//!     If numChars > 0 Then
//!         GetTrailingNumber = Right$(text, numChars)
//!     Else
//!         GetTrailingNumber = ""
//!     End If
//! End Function
//!
//! Dim num As String
//! num = GetTrailingNumber("Item123")  ' Returns "123"
//! ```
//!
//! ### 6. Time Component Extraction
//!
//! ```vb6
//! Function GetSeconds(timeStr As String) As String
//!     ' Extract seconds from "HH:MM:SS" format
//!     GetSeconds = Right$(timeStr, 2)
//! End Function
//!
//! Dim secs As String
//! secs = GetSeconds("14:30:45")  ' Returns "45"
//! ```
//!
//! ### 7. Validate String Suffix
//!
//! ```vb6
//! Function HasExtension(fileName As String, ext As String) As Boolean
//!     Dim fileExt As String
//!     fileExt = Right$(fileName, Len(ext))
//!     HasExtension = (UCase$(fileExt) = UCase$(ext))
//! End Function
//!
//! If HasExtension("report.pdf", ".pdf") Then
//!     Debug.Print "PDF file detected"
//! End If
//! ```
//!
//! ### 8. Extract Domain from Email
//!
//! ```vb6
//! Function GetEmailDomain(email As String) As String
//!     Dim atPos As Integer
//!     atPos = InStr(email, "@")
//!     If atPos > 0 Then
//!         GetEmailDomain = Right$(email, Len(email) - atPos)
//!     Else
//!         GetEmailDomain = ""
//!     End If
//! End Function
//!
//! Dim domain As String
//! domain = GetEmailDomain("user@example.com")  ' Returns "example.com"
//! ```
//!
//! ### 9. Format Currency Display
//!
//! ```vb6
//! Function FormatAmount(amount As String) As String
//!     ' Align decimal values
//!     Dim formatted As String
//!     formatted = Space(15) & amount
//!     FormatAmount = Right$(formatted, 15)
//! End Function
//! ```
//!
//! ### 10. Extract Path Component
//!
//! ```vb6
//! Function GetFileName(fullPath As String) As String
//!     Dim slashPos As Integer
//!     slashPos = InStrRev(fullPath, "\")
//!     If slashPos > 0 Then
//!         GetFileName = Right$(fullPath, Len(fullPath) - slashPos)
//!     Else
//!         GetFileName = fullPath
//!     End If
//! End Function
//!
//! Dim fileName As String
//! fileName = GetFileName("C:\Documents\report.txt")  ' Returns "report.txt"
//! ```
//!
//! ## Related Functions
//!
//! - `Right()` - Returns a `Variant` containing the rightmost characters (can handle `Null`)
//! - `Left$()` - Returns a specified number of characters from the left side of a string
//! - `Mid$()` - Returns a specified number of characters from any position in a string
//! - `Len()` - Returns the number of characters in a string
//! - `InStrRev()` - Finds the position of a substring searching from the end
//! - `Trim$()` - Removes leading and trailing spaces from a string
//! - `LTrim$()` - Removes leading spaces from a string
//! - `RTrim$()` - Removes trailing spaces from a string
//!
//! ## Best Practices
//!
//! ### When to Use `Right$` vs `Right`
//!
//! ```vb6
//! ' Use Right$ when you need a String
//! Dim fileName As String
//! fileName = Right$(fullPath, 10)  ' Type-safe, always returns String
//!
//! ' Use Right when working with Variants or Null values
//! Dim result As Variant
//! result = Right(variantValue, 5)  ' Can propagate Null
//! ```
//!
//! ### Validate Length Parameter
//!
//! ```vb6
//! Function SafeRight(text As String, length As Integer) As String
//!     If length < 0 Then
//!         SafeRight = ""
//!     ElseIf length >= Len(text) Then
//!         SafeRight = text
//!     Else
//!         SafeRight = Right$(text, length)
//!     End If
//! End Function
//! ```
//!
//! ### Check for Empty Strings
//!
//! ```vb6
//! If Len(text) > 0 Then
//!     suffix = Right$(text, 3)
//! Else
//!     suffix = ""
//! End If
//! ```
//!
//! ### Use with `InStrRev` for Parsing
//!
//! ```vb6
//! ' Find last occurrence and extract everything after it
//! Dim pos As Integer
//! pos = InStrRev(fullPath, "\")
//! If pos > 0 Then
//!     fileName = Right$(fullPath, Len(fullPath) - pos)
//! End If
//! ```
//!
//! ## Performance Considerations
//!
//! - `Right$` is very efficient for small to moderate length strings
//! - For very large strings, consider if you really need to extract characters
//! - Using `Right$` in tight loops with large strings may impact performance
//! - Consider caching the length if calling `Len()` repeatedly
//!
//! ```vb6
//! ' Less efficient
//! For i = 1 To 1000
//!     result = Right$(largeString, Len(largeString) - 10)
//! Next i
//!
//! ' More efficient
//! Dim strLen As Long
//! strLen = Len(largeString)
//! For i = 1 To 1000
//!     result = Right$(largeString, strLen - 10)
//! Next i
//! ```
//!
//! ## Common Pitfalls
//!
//! ### 1. Negative Length Values
//!
//! ```vb6
//! ' Runtime error: Invalid procedure call or argument
//! text = Right$("Hello", -1)  ' ERROR!
//!
//! ' Validate first
//! If length >= 0 Then
//!     text = Right$(source, length)
//! End If
//! ```
//!
//! ### 2. Off-by-One Errors
//!
//! ```vb6
//! ' Common mistake: forgetting to account for delimiter position
//! Dim pos As Integer
//! pos = InStrRev(path, "\")
//! ' Wrong: includes the backslash
//! fileName = Right$(path, pos)
//! ' Correct: excludes the backslash
//! fileName = Right$(path, Len(path) - pos)
//! ```
//!
//! ### 3. Not Checking String Length
//!
//! ```vb6
//! ' Potential issue: what if text is shorter than 10 characters?
//! suffix = Right$(text, 10)  ' Returns entire string if text.Length < 10
//!
//! ' Better: check first
//! If Len(text) >= 10 Then
//!     suffix = Right$(text, 10)
//! Else
//!     ' Handle short string case
//!     suffix = text
//! End If
//! ```
//!
//! ### 4. Assuming Fixed Positions
//!
//! ```vb6
//! ' Fragile: assumes extension is always 3 characters
//! ext = Right$(fileName, 3)  ' Fails for ".html", ".jpeg"
//!
//! ' Better: find the dot
//! Dim dotPos As Integer
//! dotPos = InStrRev(fileName, ".")
//! If dotPos > 0 Then
//!     ext = Right$(fileName, Len(fileName) - dotPos)
//! End If
//! ```
//!
//! ### 5. Null Value Handling
//!
//! ```vb6
//! ' Right$ with Null causes runtime error
//! Dim result As String
//! result = Right$(nullValue, 5)  ' ERROR if nullValue is Null
//!
//! ' Protect against Null
//! If Not IsNull(value) Then
//!     result = Right$(value, 5)
//! Else
//!     result = ""
//! End If
//! ```
//!
//! ## Limitations
//!
//! - Cannot handle `Null` values (use `Right` variant function instead)
//! - No built-in trimming of whitespace (combine with `RTrim$` if needed)
//! - Negative `length` values cause runtime errors
//! - Works with characters, not bytes (use `RightB$` for byte-level operations)
//! - No Unicode-specific version (VB6 uses UCS-2 internally)
//! - Cannot extract from right based on delimiter (must calculate length manually)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn right_dollar_simple() {
        let source = r#"
Sub Main()
    result = Right$("Hello", 3)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_assignment() {
        let source = r#"
Sub Main()
    Dim suffix As String
    suffix = Right$("Hello World", 5)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_variable() {
        let source = r#"
Sub Main()
    Dim text As String
    Dim result As String
    text = "Sample"
    result = Right$(text, 3)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_file_extension() {
        let source = r#"
Function GetExtension(fileName As String) As String
    Dim dotPos As Integer
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        GetExtension = Right$(fileName, Len(fileName) - dotPos)
    End If
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_account_suffix() {
        let source = r"
Function GetAccountSuffix(accountNum As String) As String
    GetAccountSuffix = Right$(accountNum, 4)
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_in_condition() {
        let source = r#"
Sub Main()
    If Right$(fileName, 4) = ".txt" Then
        Debug.Print "Text file"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_pad_left() {
        let source = r"
Function PadLeft(text As String, width As Integer) As String
    Dim padded As String
    padded = Space(width) & text
    PadLeft = Right$(padded, width)
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_time_extraction() {
        let source = r"
Function GetSeconds(timeStr As String) As String
    GetSeconds = Right$(timeStr, 2)
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_email_domain() {
        let source = r#"
Function GetEmailDomain(email As String) As String
    Dim atPos As Integer
    atPos = InStr(email, "@")
    If atPos > 0 Then
        GetEmailDomain = Right$(email, Len(email) - atPos)
    End If
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_multiple_uses() {
        let source = r"
Sub ProcessText()
    Dim ext As String
    Dim suffix As String
    ext = Right$(fileName, 3)
    suffix = Right$(accountNum, 4)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_select_case() {
        let source = r#"
Sub Main()
    Select Case Right$(fileName, 4)
        Case ".txt"
            Debug.Print "Text"
        Case ".doc"
            Debug.Print "Document"
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_expression_args() {
        let source = r"
Sub Main()
    Dim result As String
    result = Right$(text, Len(text) - 5)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_concatenation() {
        let source = r#"
Sub Main()
    Dim output As String
    output = "Suffix: " & Right$(text, 5)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_get_filename() {
        let source = r#"
Function GetFileName(fullPath As String) As String
    Dim slashPos As Integer
    slashPos = InStrRev(fullPath, "\")
    If slashPos > 0 Then
        GetFileName = Right$(fullPath, Len(fullPath) - slashPos)
    End If
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_validation() {
        let source = r"
Function HasExtension(fileName As String, ext As String) As Boolean
    Dim fileExt As String
    fileExt = Right$(fileName, Len(ext))
    HasExtension = (UCase$(fileExt) = UCase$(ext))
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_zero_length() {
        let source = r#"
Sub Main()
    Dim empty As String
    empty = Right$("Hello", 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_full_string() {
        let source = r#"
Sub Main()
    Dim full As String
    full = Right$("Hello", 100)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_format_amount() {
        let source = r"
Function FormatAmount(amount As String) As String
    Dim formatted As String
    formatted = Space(15) & amount
    FormatAmount = Right$(formatted, 15)
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_with_trim() {
        let source = r"
Sub Main()
    Dim cleaned As String
    cleaned = RTrim$(Right$(dataField, 10))
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }

    #[test]
    fn right_dollar_loop_processing() {
        let source = r"
Sub ProcessLines()
    Dim i As Integer
    Dim suffix As String
    For i = 1 To 10
        suffix = Right$(lines(i), 5)
        Debug.Print suffix
    Next i
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Right$"));
    }
}
