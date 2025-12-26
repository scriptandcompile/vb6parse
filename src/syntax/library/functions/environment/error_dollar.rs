//! VB6 `Error$` Function
//!
//! The `Error$` function returns the error message string corresponding to a given error number.
//!
//! ## Syntax
//! ```vb6
//! Error$([errornumber])
//! ```
//!
//! ## Parameters
//! - `errornumber`: Optional. A numeric expression representing an error number. If omitted, returns the error message for the most recent error (`Err.Number`).
//!
//! ## Returns
//! Returns a `String` containing the error message associated with the error number.
//!
//! ## Remarks
//! - `Error$` returns a `String`, while `Error` (without the $) returns a `Variant`.
//! - If `errornumber` is omitted, returns the message for the current error (`Err.Number`).
//! - If the error number is not recognized, returns "Application-defined or object-defined error".
//! - System errors (1-1000) return predefined messages.
//! - User-defined errors typically start at vbObjectError.
//! - Does not raise or clear errors, only retrieves messages.
//! - Can be used to display error messages to users.
//! - Related to the `Err` object and `Error` statement.
//! - Returns an empty string if error number is 0.
//!
//! ## Typical Uses
//! 1. Display error messages in message boxes
//! 2. Log error messages to files or debug output
//! 3. Format custom error messages
//! 4. Retrieve predefined system error messages
//! 5. Build error reporting systems
//! 6. Test error handling code
//! 7. Document expected errors
//! 8. Create user-friendly error dialogs
//!
//! ## Basic Examples
//!
//! ### Example 1: Get current error message
//! ```vb6
//! On Error Resume Next
//! x = 1 / 0
//! MsgBox Error$()
//! ```
//!
//! ### Example 2: Get specific error message
//! ```vb6
//! MsgBox Error$(11) ' "Division by zero"
//! ```
//!
//! ### Example 3: Display error in handler
//! ```vb6
//! On Error GoTo ErrHandler
//! ' code
//! Exit Sub
//! ErrHandler:
//!     MsgBox "Error: " & Error$()
//! ```
//!
//! ### Example 4: Log error message
//! ```vb6
//! Debug.Print "Error occurred: " & Error$()
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Display formatted error
//! ```vb6
//! MsgBox "An error occurred: " & Error$()
//! ```
//!
//! ### Pattern 2: Log error with number
//! ```vb6
//! Debug.Print "Error " & Err.Number & ": " & Error$()
//! ```
//!
//! ### Pattern 3: Custom error message
//! ```vb6
//! If Err.Number <> 0 Then
//!     msg = "Operation failed: " & Error$()
//! End If
//! ```
//!
//! ### Pattern 4: Get specific error text
//! ```vb6
//! errMsg = Error$(53) ' "File not found"
//! ```
//!
//! ### Pattern 5: Error handler logging
//! ```vb6
//! On Error GoTo ErrHandler
//! ' code
//! Exit Sub
//! ErrHandler:
//!     Open "errors.log" For Append As #1
//!     Print #1, Now & ": " & Error$()
//!     Close #1
//! ```
//!
//! ### Pattern 6: Compare error messages
//! ```vb6
//! If Error$() = Error$(11) Then
//!     ' Handle division by zero
//! End If
//! ```
//!
//! ### Pattern 7: Build error report
//! ```vb6
//! report = "Error Number: " & Err.Number & vbCrLf
//! report = report & "Description: " & Error$() & vbCrLf
//! ```
//!
//! ### Pattern 8: Test error messages
//! ```vb6
//! For i = 1 To 100
//!     Debug.Print i & ": " & Error$(i)
//! Next i
//! ```
//!
//! ### Pattern 9: User-friendly error dialog
//! ```vb6
//! MsgBox "Sorry, an error occurred:" & vbCrLf & Error$(), vbExclamation
//! ```
//!
//! ### Pattern 10: Conditional error handling
//! ```vb6
//! If InStr(Error$(), "File") > 0 Then
//!     ' Handle file-related errors
//! End If
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Comprehensive error logger
//! ```vb6
//! Sub LogError()
//!     Dim msg As String
//!     msg = "Error " & Err.Number & " at " & Now & vbCrLf
//!     msg = msg & "Description: " & Error$() & vbCrLf
//!     msg = msg & "Source: " & Err.Source & vbCrLf
//!     Debug.Print msg
//! End Sub
//! ```
//!
//! ### Example 2: Error message translator
//! ```vb6
//! Function GetFriendlyError() As String
//!     Select Case Err.Number
//!         Case 11
//!             GetFriendlyError = "Cannot divide by zero"
//!         Case 53
//!             GetFriendlyError = "The file was not found"
//!         Case Else
//!             GetFriendlyError = Error$()
//!     End Select
//! End Function
//! ```
//!
//! ### Example 3: Error documentation generator
//! ```vb6
//! Sub DocumentErrors()
//!     Dim i As Integer
//!     Open "errors.txt" For Output As #1
//!     For i = 1 To 1000
//!         If Error$(i) <> "" Then
//!             Print #1, i & vbTab & Error$(i)
//!         End If
//!     Next i
//!     Close #1
//! End Sub
//! ```
//!
//! ### Example 4: Error testing utility
//! ```vb6
//! Function TestError(errNum As Integer) As String
//!     On Error Resume Next
//!     Err.Raise errNum
//!     TestError = Error$()
//!     Err.Clear
//! End Function
//! ```
//!
//! ## Error Handling
//! - Returns empty string for error number 0.
//! - Returns "Application-defined or object-defined error" for unrecognized errors.
//! - Does not raise or clear errors itself.
//! - Safe to call at any time.
//!
//! ## Performance Notes
//! - Fast, constant time O(1) lookup.
//! - No side effects on error state.
//! - Safe for repeated calls.
//!
//! ## Best Practices
//! 1. Use with error handlers for user-friendly messages.
//! 2. Combine with `Err.Number` for complete error info.
//! 3. Log `Error$()` output for debugging.
//! 4. Don't rely on exact message text (use error numbers).
//! 5. Provide context with error messages.
//! 6. Use for displaying errors to users.
//! 7. Document which errors your code may encounter.
//! 8. Test error paths with `Error$()` logging.
//! 9. Prefer `Error$()` over `Error` for `String` variables.
//! 10. Clear errors after handling with `Err.Clear`.
//!
//! ## Comparison Table
//!
//! | Function/Statement | Purpose                    | Returns        |
//! |--------------------|----------------------------|----------------|
//! | `Error$`           | Get error message string   | `String`       |
//! | `Error`            | Get error message variant  | `Variant`      |
//! | `Err.Description`  | Current error description  | `String`       |
//! | `Err.Number`       | Current error number       | `Long`         |
//!
//! ## Platform Notes
//! - Available in VB6 and VBA.
//! - Error messages are in English by default.
//! - Some error messages may be locale-specific.
//! - `VBScript` uses `Err.Description` instead.
//!
//! ## Limitations
//! - Returns English messages (may not be localized).
//! - Cannot customize built-in error messages.
//! - Limited to VB6's predefined error numbers.
//! - Does not provide error source or context.
//! - Message text may change between VB versions.

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn error_dollar_current_error() {
        let source = r"
Sub Test()
    On Error Resume Next
    x = 1 / 0
    msg = Error$()
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_specific_error() {
        let source = r"
Sub Test()
    msg = Error$(11)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_in_handler() {
        let source = r#"
Sub Test()
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    msg = "Error: " & Error$()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Error occurred: " & Error$()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_formatted_message() {
        let source = r#"
Sub Test()
    msg = "An error occurred: " & Error$()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_with_number() {
        let source = r#"
Sub Test()
    Debug.Print "Error " & Err.Number & ": " & Error$()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_custom_message() {
        let source = r#"
Sub Test()
    If Err.Number <> 0 Then
        msg = "Operation failed: " & Error$()
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_get_specific() {
        let source = r"
Sub Test()
    errMsg = Error$(53)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_logging() {
        let source = r#"
Sub Test()
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    Open "errors.log" For Append As #1
    Print #1, Now & ": " & Error$()
    Close #1
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_compare_messages() {
        let source = r"
Sub Test()
    If Error$() = Error$(11) Then
        ' Handle division by zero
    End If
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_build_report() {
        let source = r#"
Sub Test()
    report = "Error Number: " & Err.Number & vbCrLf
    report = report & "Description: " & Error$() & vbCrLf
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_messages() {
        let source = r#"
Sub Test()
    For i = 1 To 100
        Debug.Print i & ": " & Error$(i)
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_user_dialog() {
        let source = r#"
Sub Test()
    msg = "Sorry, an error occurred:" & vbCrLf & Error$()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_conditional_handling() {
        let source = r#"
Sub Test()
    If InStr(Error$(), "File") > 0 Then
        ' Handle file-related errors
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_comprehensive_logger() {
        let source = r#"
Sub LogError()
    Dim msg As String
    msg = "Error " & Err.Number & " at " & Now & vbCrLf
    msg = msg & "Description: " & Error$() & vbCrLf
    msg = msg & "Source: " & Err.Source & vbCrLf
    Debug.Print msg
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_friendly_translator() {
        let source = r#"
Function GetFriendlyError() As String
    Select Case Err.Number
        Case 11
            GetFriendlyError = "Cannot divide by zero"
        Case 53
            GetFriendlyError = "The file was not found"
        Case Else
            GetFriendlyError = Error$()
    End Select
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_documentation_generator() {
        let source = r#"
Sub DocumentErrors()
    Dim i As Integer
    Open "errors.txt" For Output As #1
    For i = 1 To 1000
        If Error$(i) <> "" Then
            Print #1, i & vbTab & Error$(i)
        End If
    Next i
    Close #1
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }

    #[test]
    fn error_dollar_testing_utility() {
        let source = r"
Function TestError(errNum As Integer) As String
    On Error Resume Next
    Err.Raise errNum
    TestError = Error$()
    Err.Clear
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Error$"));
    }
}
