//! # `Trim$` Function
//!
//! The `Trim$` function in Visual Basic 6 returns a string with both leading and trailing spaces
//! removed. The dollar sign (`$`) suffix indicates that this function always returns a `String`
//! type, never a `Variant`.
//!
//! ## Syntax
//!
//! ```vb6
//! Trim$(string)
//! ```
//!
//! ## Parameters
//!
//! - `string` - Required. Any valid string expression. If `string` contains `Null`, `Null` is returned.
//!
//! ## Return Value
//!
//! Returns a `String` with all leading and trailing space characters (ASCII 32) removed from `string`.
//!
//! ## Behavior and Characteristics
//!
//! ### Space Removal
//!
//! - Removes both leading (left-side) and trailing (right-side) spaces
//! - Only removes space characters (ASCII character 32)
//! - Does not remove tabs, newlines, or other whitespace characters
//! - If the string contains only spaces, returns an empty string ("")
//! - Preserves spaces in the middle of the string
//!
//! ### Type Differences: `Trim$` vs `Trim`
//!
//! - `Trim$`: Always returns `String` type (never `Variant`)
//! - `Trim`: Returns `Variant` (can propagate `Null` values)
//! - Use `Trim$` when you need guaranteed `String` return type
//! - Use `Trim` when working with potentially `Null` values
//!
//! ## Common Usage Patterns
//!
//! ### 1. Clean User Input
//!
//! ```vb6
//! Function CleanInput(userInput As String) As String
//!     CleanInput = Trim$(userInput)
//! End Function
//!
//! Dim cleaned As String
//! cleaned = CleanInput("  Hello World  ")  ' Returns "Hello World"
//! ```
//!
//! ### 2. Process Text File Lines
//!
//! ```vb6
//! Function ReadCleanLine(fileNum As Integer) As String
//!     Dim rawLine As String
//!     Line Input #fileNum, rawLine
//!     ReadCleanLine = Trim$(rawLine)
//! End Function
//! ```
//!
//! ### 3. Validate Non-Empty Input
//!
//! ```vb6
//! Function IsValidInput(input As String) As Boolean
//!     IsValidInput = (Len(Trim$(input)) > 0)
//! End Function
//!
//! If IsValidInput(txtName.Text) Then
//!     ' Process the input
//! Else
//!     MsgBox "Please enter a value"
//! End If
//! ```
//!
//! ### 4. String Comparison
//!
//! ```vb6
//! Function CompareValues(value1 As String, value2 As String) As Boolean
//!     ' Compare strings ignoring leading/trailing spaces
//!     CompareValues = (Trim$(value1) = Trim$(value2))
//! End Function
//! ```
//!
//! ### 5. Database Field Cleaning
//!
//! ```vb6
//! Function GetFieldValue(rs As Recordset, fieldName As String) As String
//!     If Not IsNull(rs.Fields(fieldName).Value) Then
//!         GetFieldValue = Trim$(rs.Fields(fieldName).Value & "")
//!     Else
//!         GetFieldValue = ""
//!     End If
//! End Function
//! ```
//!
//! ### 6. Configuration File Parsing
//!
//! ```vb6
//! Function ParseConfigLine(configLine As String) As String
//!     Dim equalPos As Integer
//!     equalPos = InStr(configLine, "=")
//!     If equalPos > 0 Then
//!         ParseConfigLine = Trim$(Mid$(configLine, equalPos + 1))
//!     Else
//!         ParseConfigLine = ""
//!     End If
//! End Function
//! ```
//!
//! ### 7. Clean Array Elements
//!
//! ```vb6
//! Sub CleanStringArray(arr() As String)
//!     Dim i As Integer
//!     For i = LBound(arr) To UBound(arr)
//!         arr(i) = Trim$(arr(i))
//!     Next i
//! End Sub
//! ```
//!
//! ### 8. Form Input Processing
//!
//! ```vb6
//! Sub ProcessForm()
//!     Dim userName As String
//!     Dim userEmail As String
//!     
//!     userName = Trim$(txtName.Text)
//!     userEmail = Trim$(txtEmail.Text)
//!     
//!     If Len(userName) > 0 And Len(userEmail) > 0 Then
//!         SaveUserData userName, userEmail
//!     End If
//! End Sub
//! ```
//!
//! ### 9. CSV Field Processing
//!
//! ```vb6
//! Function ParseCSVField(field As String) As String
//!     ' Remove quotes and trim spaces
//!     Dim cleaned As String
//!     cleaned = field
//!     If Left$(cleaned, 1) = """" Then cleaned = Mid$(cleaned, 2)
//!     If Right$(cleaned, 1) = """" Then cleaned = Left$(cleaned, Len(cleaned) - 1)
//!     ParseCSVField = Trim$(cleaned)
//! End Function
//! ```
//!
//! ### 10. Search Query Preparation
//!
//! ```vb6
//! Function PrepareSearchQuery(query As String) As String
//!     Dim cleaned As String
//!     cleaned = Trim$(query)
//!     ' Remove multiple spaces
//!     While InStr(cleaned, "  ") > 0
//!         cleaned = Replace$(cleaned, "  ", " ")
//!     Wend
//!     PrepareSearchQuery = cleaned
//! End Function
//! ```
//!
//! ## Related Functions
//!
//! - `Trim()` - Returns a `Variant` with leading and trailing spaces removed (can handle `Null`)
//! - `LTrim$()` - Removes only leading (left-side) spaces from a string
//! - `RTrim$()` - Removes only trailing (right-side) spaces from a string
//! - `Left$()` - Returns a specified number of characters from the left side
//! - `Right$()` - Returns a specified number of characters from the right side
//! - `Len()` - Returns the length of a string
//! - `Space$()` - Creates a string consisting of the specified number of spaces
//!
//! ## Best Practices
//!
//! ### When to Use `Trim$` vs `LTrim$` vs `RTrim$`
//!
//! ```vb6
//! Dim text As String
//! text = "  Hello  "
//!
//! Debug.Print Trim$(text)   ' "Hello" (both sides trimmed)
//! Debug.Print LTrim$(text)  ' "Hello  " (left side only)
//! Debug.Print RTrim$(text)  ' "  Hello" (right side only)
//! ```
//!
//! ### Use for User Input Validation
//!
//! ```vb6
//! Function ValidateInput(input As String) As Boolean
//!     ' Check if input is meaningful after trimming
//!     Dim cleaned As String
//!     cleaned = Trim$(input)
//!     
//!     If Len(cleaned) = 0 Then
//!         MsgBox "Input cannot be empty or only spaces"
//!         ValidateInput = False
//!     Else
//!         ValidateInput = True
//!     End If
//! End Function
//! ```
//!
//! ### Combine with Other String Functions
//!
//! ```vb6
//! Function NormalizeText(text As String) As String
//!     Dim result As String
//!     result = Trim$(text)
//!     result = UCase$(result)  ' Convert to uppercase
//!     NormalizeText = result
//! End Function
//! ```
//!
//! ### Handle Null Values Safely
//!
//! ```vb6
//! Function SafeTrim(value As Variant) As String
//!     If IsNull(value) Then
//!         SafeTrim = ""
//!     Else
//!         SafeTrim = Trim$(CStr(value))
//!     End If
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - `Trim$` is very efficient and lightweight
//! - Performs a single pass from both ends of the string
//! - More efficient than calling `LTrim$` and `RTrim$` separately
//! - No performance penalty for strings without leading/trailing spaces
//!
//! ```vb6
//! ' Efficient: single Trim$ call
//! Dim cleaned As String
//! cleaned = Trim$(input)
//!
//! ' Less efficient: two function calls
//! cleaned = LTrim$(RTrim$(input))
//!
//! ' Use Trim$ instead of the above
//! ```
//!
//! ## Common Pitfalls
//!
//! ### 1. Only Removes Spaces (ASCII 32)
//!
//! ```vb6
//! Dim text As String
//! text = vbTab & "Hello" & vbTab  ' Tabs at both ends
//!
//! ' Trim$ does NOT remove tabs
//! Debug.Print Trim$(text)  ' Still has tabs!
//!
//! ' To remove all whitespace, use custom function
//! Function TrimAllWhitespace(s As String) As String
//!     Dim i As Integer, j As Integer
//!     
//!     ' Trim from left
//!     For i = 1 To Len(s)
//!         Select Case Mid$(s, i, 1)
//!             Case " ", vbTab, vbCr, vbLf
//!                 ' Continue
//!             Case Else
//!                 Exit For
//!         End Select
//!     Next i
//!     
//!     ' Trim from right
//!     For j = Len(s) To 1 Step -1
//!         Select Case Mid$(s, j, 1)
//!             Case " ", vbTab, vbCr, vbLf
//!                 ' Continue
//!             Case Else
//!                 Exit For
//!         End Select
//!     Next j
//!     
//!     If i <= j Then
//!         TrimAllWhitespace = Mid$(s, i, j - i + 1)
//!     Else
//!         TrimAllWhitespace = ""
//!     End If
//! End Function
//! ```
//!
//! ### 2. Null Value Handling
//!
//! ```vb6
//! ' Trim$ with Null causes runtime error
//! Dim result As String
//! result = Trim$(nullValue)  ' ERROR if nullValue is Null
//!
//! ' Protect against Null
//! If Not IsNull(value) Then
//!     result = Trim$(value)
//! Else
//!     result = ""
//! End If
//! ```
//!
//! ### 3. Empty String vs Spaces-Only String
//!
//! ```vb6
//! Dim input As String
//! input = "     "  ' Only spaces
//!
//! ' Trim$ returns empty string
//! Debug.Print Len(Trim$(input))  ' 0
//!
//! ' Check for meaningful content
//! If Len(Trim$(input)) = 0 Then
//!     Debug.Print "No content"
//! End If
//! ```
//!
//! ### 4. Database Field Assumptions
//!
//! ```vb6
//! ' Wrong: not handling Null
//! value = Trim$(rs!TextField)  ' May error if field is Null
//!
//! ' Better: handle Null explicitly
//! If IsNull(rs!TextField) Then
//!     value = ""
//! Else
//!     value = Trim$(rs!TextField & "")
//! End If
//! ```
//!
//! ### 5. Case Sensitivity
//!
//! ```vb6
//! ' Trim$ does not change case
//! Debug.Print Trim$("  HELLO  ")  ' "HELLO" (not "hello")
//!
//! ' Combine with case conversion if needed
//! Debug.Print UCase$(Trim$("  hello  "))  ' "HELLO"
//! Debug.Print LCase$(Trim$("  HELLO  "))  ' "hello"
//! ```
//!
//! ### 6. Non-Breaking Spaces
//!
//! ```vb6
//! ' Trim$ only removes regular spaces (ASCII 32)
//! ' Non-breaking spaces (Chr$(160)) are NOT removed
//! Dim text As String
//! text = Chr$(160) & "Hello" & Chr$(160)
//! Debug.Print Trim$(text)  ' Still has Chr$(160) at both ends
//! ```
//!
//! ## Practical Examples
//!
//! ### Processing INI File Values
//!
//! ```vb6
//! Function GetINIValue(section As String, key As String, fileName As String) As String
//!     Dim fileNum As Integer
//!     Dim currentLine As String
//!     Dim inSection As Boolean
//!     Dim equalPos As Integer
//!     Dim lineKey As String
//!     
//!     fileNum = FreeFile
//!     Open fileName For Input As #fileNum
//!     
//!     While Not EOF(fileNum)
//!         Line Input #fileNum, currentLine
//!         currentLine = Trim$(currentLine)
//!         
//!         If currentLine = "[" & section & "]" Then
//!             inSection = True
//!         ElseIf Left$(currentLine, 1) = "[" Then
//!             inSection = False
//!         ElseIf inSection Then
//!             equalPos = InStr(currentLine, "=")
//!             If equalPos > 0 Then
//!                 lineKey = Trim$(Left$(currentLine, equalPos - 1))
//!                 If lineKey = key Then
//!                     GetINIValue = Trim$(Mid$(currentLine, equalPos + 1))
//!                     Close #fileNum
//!                     Exit Function
//!                 End If
//!             End If
//!         End If
//!     Wend
//!     
//!     Close #fileNum
//!     GetINIValue = ""
//! End Function
//! ```
//!
//! ### Form Validation
//!
//! ```vb6
//! Function ValidateForm() As Boolean
//!     Dim errors As String
//!     
//!     If Len(Trim$(txtName.Text)) = 0 Then
//!         errors = errors & "Name is required" & vbCrLf
//!     End If
//!     
//!     If Len(Trim$(txtEmail.Text)) = 0 Then
//!         errors = errors & "Email is required" & vbCrLf
//!     End If
//!     
//!     If Len(errors) > 0 Then
//!         MsgBox errors, vbExclamation
//!         ValidateForm = False
//!     Else
//!         ValidateForm = True
//!     End If
//! End Function
//! ```
//!
//! ### Clean Data Import
//!
//! ```vb6
//! Sub ImportCSVData(fileName As String)
//!     Dim fileNum As Integer
//!     Dim currentLine As String
//!     Dim fields() As String
//!     Dim i As Integer
//!     
//!     fileNum = FreeFile
//!     Open fileName For Input As #fileNum
//!     
//!     While Not EOF(fileNum)
//!         Line Input #fileNum, currentLine
//!         fields = Split(currentLine, ",")
//!         
//!         ' Clean all fields
//!         For i = LBound(fields) To UBound(fields)
//!             fields(i) = Trim$(fields(i))
//!         Next i
//!         
//!         ' Process cleaned fields
//!         ProcessRecord fields
//!     Wend
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ## Limitations
//!
//! - Only removes space characters (ASCII 32), not other whitespace
//! - Cannot handle `Null` values (use `Trim` variant function instead)
//! - Does not remove non-breaking spaces (character 160) or Unicode whitespace
//! - No option to specify custom characters to remove
//! - Works with strings only, not byte arrays
//! - Does not change character case (use with `UCase$` or `LCase$` if needed)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn trim_dollar_simple() {
        let source = r#"
Sub Main()
    result = Trim$("  Hello  ")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_assignment() {
        let source = r#"
Sub Main()
    Dim cleaned As String
    cleaned = Trim$(userInput)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_variable() {
        let source = r#"
Sub Main()
    Dim text As String
    Dim result As String
    text = "  Sample  "
    result = Trim$(text)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_clean_input() {
        let source = r#"
Function CleanInput(userInput As String) As String
    CleanInput = Trim$(userInput)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_validate_input() {
        let source = r#"
Function IsValidInput(input As String) As Boolean
    IsValidInput = (Len(Trim$(input)) > 0)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_in_condition() {
        let source = r#"
Sub Main()
    If Trim$(input) = "Expected" Then
        Debug.Print "Match found"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_comparison() {
        let source = r#"
Function CompareValues(value1 As String, value2 As String) As Boolean
    CompareValues = (Trim$(value1) = Trim$(value2))
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_database_field() {
        let source = r#"
Function GetFieldValue(fieldValue As String) As String
    GetFieldValue = Trim$(fieldValue)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_array_cleanup() {
        let source = r#"
Sub CleanStringArray(arr() As String)
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        arr(i) = Trim$(arr(i))
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_multiple_uses() {
        let source = r#"
Sub ProcessForm()
    Dim userName As String
    Dim userEmail As String
    userName = Trim$(txtName.Text)
    userEmail = Trim$(txtEmail.Text)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_select_case() {
        let source = r#"
Sub Main()
    Select Case Trim$(status)
        Case "Active"
            Debug.Print "Active record"
        Case "Inactive"
            Debug.Print "Inactive record"
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_concatenation() {
        let source = r#"
Sub Main()
    Dim output As String
    output = "Name: " & Trim$(nameField)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_config_parsing() {
        let source = r#"
Function ParseConfigLine(configLine As String) As String
    Dim equalPos As Integer
    equalPos = InStr(configLine, "=")
    If equalPos > 0 Then
        ParseConfigLine = Trim$(Mid$(configLine, equalPos + 1))
    End If
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_file_processing() {
        let source = r#"
Function ReadCleanLine(fileNum As Integer) As String
    Dim rawLine As String
    Line Input #fileNum, rawLine
    ReadCleanLine = Trim$(rawLine)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_empty_check() {
        let source = r#"
Sub Main()
    If Len(Trim$(input)) = 0 Then
        Debug.Print "Empty or spaces only"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_with_ucase() {
        let source = r#"
Function NormalizeText(text As String) As String
    NormalizeText = UCase$(Trim$(text))
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_loop_processing() {
        let source = r#"
Sub ProcessLines()
    Dim i As Integer
    Dim cleanLine As String
    For i = 1 To 10
        cleanLine = Trim$(lines(i))
        Debug.Print cleanLine
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_in_function() {
        let source = r#"
Function GetCleanValue(value As String) As String
    GetCleanValue = Trim$(value)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_csv_parsing() {
        let source = r#"
Function ParseCSVField(field As String) As String
    ParseCSVField = Trim$(field)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }

    #[test]
    fn trim_dollar_search_query() {
        let source = r#"
Function PrepareSearchQuery(query As String) As String
    Dim cleaned As String
    cleaned = Trim$(query)
    PrepareSearchQuery = cleaned
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Trim$"));
    }
}
