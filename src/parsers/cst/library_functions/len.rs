//! # `Len` Function
//!
//! Returns a `Long` containing the number of characters in a string or the number of bytes required to store a variable.
//!
//! ## Syntax
//!
//! ```vb
//! Len(string | varname)
//! ```
//!
//! ## Parameters
//!
//! - `string` (Optional): Any valid string expression
//!   - If string contains `Null`, `Null` is returned
//! - `varname` (Optional): Any valid variable name
//!   - For user-defined types, returns size in bytes
//!   - For objects, may return implementation-defined value
//!
//! ## Return Value
//!
//! Returns a Long:
//! - For strings: Number of characters in the string
//! - For variables: Number of bytes required to store the variable
//! - Returns 0 for empty string ("")
//! - Returns `Null` if string argument is `Null`
//! - For fixed-length strings: Returns declared length
//! - For Variant containing string: Returns length of string
//! - For user-defined types: Returns total size in bytes
//!
//! ## Remarks
//!
//! The `Len` function measures string length or variable size:
//!
//! - Most commonly used for string length
//! - Returns character count, not byte count (for strings)
//! - Empty string ("") returns 0
//! - Null propagates through the function
//! - Spaces and special characters are counted
//! - For fixed-length strings (e.g., String * 10), returns declared length
//! - Variable-length strings return actual content length
//! - `LenB` function returns byte count (useful for Unicode/DBCS)
//! - Cannot determine array size (use UBound/LBound instead)
//! - Essential for string manipulation and validation
//! - Used with Left, Right, Mid for substring operations
//! - Common in loops iterating through string characters
//! - Efficient operation with minimal overhead
//! - For user-defined types, returns structure size
//! - Object variables may return unpredictable values
//!
//! ## Typical Uses
//!
//! 1. **Validate Input**: Check if string is empty or within length limits
//! 2. **String Iteration**: Loop through each character in string
//! 3. **Substring Extraction**: Calculate positions for Left/Right/Mid
//! 4. **Text Processing**: Measure text before truncation or padding
//! 5. **File I/O**: Determine buffer sizes for string data
//! 6. **Data Validation**: Ensure fields meet length requirements
//! 7. **Format Verification**: Check string format compliance
//! 8. **Memory Calculation**: Determine structure sizes
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Basic string length
//! Dim text As String
//! text = "Hello World"
//!
//! Debug.Print Len(text)                ' 11 - includes space
//! Debug.Print Len("VB6")               ' 3
//! Debug.Print Len("")                  ' 0 - empty string
//!
//! ' Example 2: Validation
//! Dim password As String
//! password = "secret123"
//!
//! If Len(password) < 8 Then
//!     MsgBox "Password must be at least 8 characters"
//! End If
//!
//! ' Example 3: String iteration
//! Dim i As Long
//! Dim char As String
//! text = "ABC"
//!
//! For i = 1 To Len(text)
//!     char = Mid(text, i, 1)
//!     Debug.Print char
//! Next i
//!
//! ' Example 4: Null handling
//! Dim value As Variant
//! value = Null
//!
//! Debug.Print IsNull(Len(value))       ' True - Null propagates
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Check if string is empty
//! Function IsEmpty(text As String) As Boolean
//!     IsEmpty = (Len(text) = 0)
//! End Function
//!
//! ' Pattern 2: Validate string length range
//! Function ValidateLength(text As String, minLen As Long, maxLen As Long) As Boolean
//!     Dim length As Long
//!     length = Len(text)
//!     ValidateLength = (length >= minLen And length <= maxLen)
//! End Function
//!
//! ' Pattern 3: Pad string to fixed width
//! Function PadRight(text As String, width As Long, Optional padChar As String = " ") As String
//!     If Len(text) >= width Then
//!         PadRight = Left(text, width)
//!     Else
//!         PadRight = text & String(width - Len(text), padChar)
//!     End If
//! End Function
//!
//! ' Pattern 4: Center text in field
//! Function CenterText(text As String, width As Long) As String
//!     Dim padding As Long
//!     
//!     If Len(text) >= width Then
//!         CenterText = Left(text, width)
//!     Else
//!         padding = (width - Len(text)) \ 2
//!         CenterText = String(padding, " ") & text & _
//!                      String(width - Len(text) - padding, " ")
//!     End If
//! End Function
//!
//! ' Pattern 5: Reverse string
//! Function ReverseString(text As String) As String
//!     Dim i As Long
//!     Dim result As String
//!     
//!     result = ""
//!     For i = Len(text) To 1 Step -1
//!         result = result & Mid(text, i, 1)
//!     Next i
//!     
//!     ReverseString = result
//! End Function
//!
//! ' Pattern 6: Count occurrences of character
//! Function CountChar(text As String, searchChar As String) As Long
//!     Dim i As Long
//!     Dim count As Long
//!     
//!     count = 0
//!     For i = 1 To Len(text)
//!         If Mid(text, i, 1) = searchChar Then
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     CountChar = count
//! End Function
//!
//! ' Pattern 7: Truncate with ellipsis
//! Function TruncateText(text As String, maxLength As Long) As String
//!     If Len(text) <= maxLength Then
//!         TruncateText = text
//!     ElseIf maxLength <= 3 Then
//!         TruncateText = Left(text, maxLength)
//!     Else
//!         TruncateText = Left(text, maxLength - 3) & "..."
//!     End If
//! End Function
//!
//! ' Pattern 8: Extract last N characters
//! Function GetLastChars(text As String, count As Long) As String
//!     If count >= Len(text) Then
//!         GetLastChars = text
//!     Else
//!         GetLastChars = Right(text, count)
//!     End If
//! End Function
//!
//! ' Pattern 9: Check minimum length
//! Function HasMinLength(text As String, minLength As Long) As Boolean
//!     HasMinLength = (Len(text) >= minLength)
//! End Function
//!
//! ' Pattern 10: Calculate character position from end
//! Function GetCharFromEnd(text As String, posFromEnd As Long) As String
//!     Dim pos As Long
//!     pos = Len(text) - posFromEnd + 1
//!     
//!     If pos >= 1 And pos <= Len(text) Then
//!         GetCharFromEnd = Mid(text, pos, 1)
//!     Else
//!         GetCharFromEnd = ""
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: String analyzer
//! Public Class StringAnalyzer
//!     Public Function GetStats(text As String) As String
//!         Dim length As Long
//!         Dim words As Long
//!         Dim lines As Long
//!         Dim result As String
//!         
//!         length = Len(text)
//!         words = UBound(Split(text, " ")) + 1
//!         lines = UBound(Split(text, vbCrLf)) + 1
//!         
//!         result = "Length: " & length & vbCrLf
//!         result = result & "Words: " & words & vbCrLf
//!         result = result & "Lines: " & lines
//!         
//!         GetStats = result
//!     End Function
//!     
//!     Public Function GetCharFrequency(text As String) As Object
//!         Dim dict As Object
//!         Dim i As Long
//!         Dim char As String
//!         
//!         Set dict = CreateObject("Scripting.Dictionary")
//!         
//!         For i = 1 To Len(text)
//!             char = Mid(text, i, 1)
//!             If dict.Exists(char) Then
//!                 dict(char) = dict(char) + 1
//!             Else
//!                 dict.Add char, 1
//!             End If
//!         Next i
//!         
//!         Set GetCharFrequency = dict
//!     End Function
//!     
//!     Public Function IsPalindrome(text As String) As Boolean
//!         Dim i As Long
//!         Dim length As Long
//!         
//!         length = Len(text)
//!         For i = 1 To length \ 2
//!             If Mid(text, i, 1) <> Mid(text, length - i + 1, 1) Then
//!                 IsPalindrome = False
//!                 Exit Function
//!             End If
//!         Next i
//!         
//!         IsPalindrome = True
//!     End Function
//! End Class
//!
//! ' Example 2: Text formatter
//! Public Class TextFormatter
//!     Public Function WordWrap(text As String, lineWidth As Long) As String
//!         Dim result As String
//!         Dim currentLine As String
//!         Dim words() As String
//!         Dim i As Long
//!         Dim word As String
//!         
//!         result = ""
//!         currentLine = ""
//!         words = Split(text, " ")
//!         
//!         For i = LBound(words) To UBound(words)
//!             word = words(i)
//!             
//!             If Len(currentLine) = 0 Then
//!                 currentLine = word
//!             ElseIf Len(currentLine) + 1 + Len(word) <= lineWidth Then
//!                 currentLine = currentLine & " " & word
//!             Else
//!                 result = result & currentLine & vbCrLf
//!                 currentLine = word
//!             End If
//!         Next i
//!         
//!         If Len(currentLine) > 0 Then
//!             result = result & currentLine
//!         End If
//!         
//!         WordWrap = result
//!     End Function
//!     
//!     Public Function JustifyText(text As String, width As Long) As String
//!         Dim words() As String
//!         Dim totalLen As Long
//!         Dim gaps As Long
//!         Dim extraSpaces As Long
//!         Dim spacesPerGap As Long
//!         Dim result As String
//!         Dim i As Long
//!         
//!         words = Split(text, " ")
//!         If UBound(words) = 0 Then
//!             JustifyText = text
//!             Exit Function
//!         End If
//!         
//!         ' Calculate total word length
//!         totalLen = 0
//!         For i = LBound(words) To UBound(words)
//!             totalLen = totalLen + Len(words(i))
//!         Next i
//!         
//!         gaps = UBound(words) - LBound(words)
//!         extraSpaces = width - totalLen
//!         spacesPerGap = extraSpaces \ gaps
//!         
//!         result = ""
//!         For i = LBound(words) To UBound(words)
//!             result = result & words(i)
//!             If i < UBound(words) Then
//!                 result = result & String(spacesPerGap, " ")
//!             End If
//!         Next i
//!         
//!         JustifyText = result
//!     End Function
//! End Class
//!
//! ' Example 3: Input validator
//! Public Class InputValidator
//!     Public Function ValidateEmail(email As String) As Boolean
//!         ' Basic email validation
//!         If Len(email) = 0 Then
//!             ValidateEmail = False
//!             Exit Function
//!         End If
//!         
//!         If InStr(email, "@") = 0 Then
//!             ValidateEmail = False
//!             Exit Function
//!         End If
//!         
//!         If InStr(email, ".") = 0 Then
//!             ValidateEmail = False
//!             Exit Function
//!         End If
//!         
//!         ValidateEmail = True
//!     End Function
//!     
//!     Public Function ValidatePassword(password As String) As String
//!         Dim errors As Collection
//!         Dim i As Long
//!         Dim char As String
//!         Dim hasUpper As Boolean
//!         Dim hasLower As Boolean
//!         Dim hasDigit As Boolean
//!         
//!         Set errors = New Collection
//!         
//!         If Len(password) < 8 Then
//!             errors.Add "Password must be at least 8 characters"
//!         End If
//!         
//!         If Len(password) > 128 Then
//!             errors.Add "Password must not exceed 128 characters"
//!         End If
//!         
//!         hasUpper = False
//!         hasLower = False
//!         hasDigit = False
//!         
//!         For i = 1 To Len(password)
//!             char = Mid(password, i, 1)
//!             If char >= "A" And char <= "Z" Then hasUpper = True
//!             If char >= "a" And char <= "z" Then hasLower = True
//!             If char >= "0" And char <= "9" Then hasDigit = True
//!         Next i
//!         
//!         If Not hasUpper Then errors.Add "Password must contain uppercase letter"
//!         If Not hasLower Then errors.Add "Password must contain lowercase letter"
//!         If Not hasDigit Then errors.Add "Password must contain digit"
//!         
//!         If errors.Count = 0 Then
//!             ValidatePassword = ""
//!         Else
//!             Dim msg As String
//!             msg = ""
//!             For i = 1 To errors.Count
//!                 msg = msg & errors(i) & vbCrLf
//!             Next i
//!             ValidatePassword = msg
//!         End If
//!     End Function
//!     
//!     Public Function ValidateLength(text As String, fieldName As String, _
//!                                    minLen As Long, maxLen As Long) As String
//!         Dim length As Long
//!         length = Len(text)
//!         
//!         If length < minLen Then
//!             ValidateLength = fieldName & " must be at least " & minLen & " characters"
//!         ElseIf length > maxLen Then
//!             ValidateLength = fieldName & " must not exceed " & maxLen & " characters"
//!         Else
//!             ValidateLength = ""
//!         End If
//!     End Function
//! End Class
//!
//! ' Example 4: String builder with size tracking
//! Public Class StringBuilder
//!     Private m_buffer As String
//!     Private m_length As Long
//!     
//!     Private Sub Class_Initialize()
//!         m_buffer = ""
//!         m_length = 0
//!     End Sub
//!     
//!     Public Sub Append(text As String)
//!         m_buffer = m_buffer & text
//!         m_length = m_length + Len(text)
//!     End Sub
//!     
//!     Public Sub AppendLine(Optional text As String = "")
//!         Append text & vbCrLf
//!     End Sub
//!     
//!     Public Function ToString() As String
//!         ToString = m_buffer
//!     End Function
//!     
//!     Public Property Get Length() As Long
//!         Length = m_length
//!     End Property
//!     
//!     Public Sub Clear()
//!         m_buffer = ""
//!         m_length = 0
//!     End Sub
//!     
//!     Public Function Substring(startIndex As Long, length As Long) As String
//!         If startIndex < 0 Or startIndex >= m_length Then
//!             Substring = ""
//!             Exit Function
//!         End If
//!         
//!         Substring = Mid(m_buffer, startIndex + 1, length)
//!     End Function
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! Len handles special cases gracefully:
//!
//! ```vb
//! ' Empty string returns 0
//! Debug.Print Len("")                  ' 0
//!
//! ' Null propagates
//! Dim value As Variant
//! value = Null
//! Debug.Print IsNull(Len(value))       ' True
//!
//! ' Spaces are counted
//! Debug.Print Len("   ")               ' 3
//!
//! ' Special characters are counted
//! Debug.Print Len(vbCrLf)              ' 2 - CR + LF
//! Debug.Print Len(vbTab)               ' 1
//!
//! ' Safe pattern with Null check
//! Function SafeLen(value As Variant) As Long
//!     If IsNull(value) Then
//!         SafeLen = 0
//!     Else
//!         SafeLen = Len(value)
//!     End If
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: `Len` is very fast, minimal overhead
//! - **Cache Result**: If checking length multiple times, cache the value
//! - **String Iteration**: For large strings, cache `Len` result in loop variable
//! - **No Side Effects**: `Len` does not modify the string
//!
//! Performance tips:
//! ```vb
//! ' Less efficient - calls Len every iteration
//! For i = 1 To Len(text)
//!     ' process
//! Next i
//!
//! ' More efficient - cache length
//! Dim length As Long
//! length = Len(text)
//! For i = 1 To length
//!     ' process
//! Next i
//! ```
//!
//! ## Best Practices
//!
//! 1. **Validate Input**: Always check string length for validation
//! 2. **`Null` Safety**: Handle `Null` values before calling `Len` if needed
//! 3. **Cache Results**: Store length in variable when used multiple times
//! 4. **Empty Check**: Use ```Len(str) = 0``` to check for empty strings
//! 5. **Loop Optimization**: Cache `Len` result before loops
//! 6. **Bounds Checking**: Verify length before substring operations
//! 7. **Fixed-Length Strings**: Remember `Len` returns declared length, not content
//! 8. **Database Fields**: Validate field lengths before database inserts
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | `Len` | Get string length | `Long` | Character count |
//! | `LenB` | Get byte length | `Long` | Byte count (Unicode/DBCS) |
//! | `UBound` | Get array upper bound | `Long` | Array size calculation |
//! | `InStr` | Find substring | `Long` | Substring position |
//! | `Left`/`Right`/`Mid` | Extract substring | `String` | Substring extraction |
//! ## Len vs LenB
//!
//! ```vb
//! Dim text As String
//! text = "Hello"
//!
//! ' Len - character count
//! Debug.Print Len(text)                ' 5 characters
//!
//! ' LenB - byte count (may differ for Unicode)
//! Debug.Print LenB(text)               ' 10 bytes (2 bytes per char in Unicode)
//!
//! ' For ASCII characters, LenB is typically 2 * Len
//! ' For DBCS characters, varies by character
//! ```
//!
//! ## Fixed-Length vs Variable-Length Strings
//!
//! ```vb
//! ' Variable-length string - Len returns actual content length
//! Dim varStr As String
//! varStr = "ABC"
//! Debug.Print Len(varStr)              ' 3
//!
//! ' Fixed-length string - Len returns declared length
//! Dim fixStr As String * 10
//! fixStr = "ABC"
//! Debug.Print Len(fixStr)              ' 10 (includes padding spaces)
//! ```
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core functions
//! - Returns Long type
//! - Maximum string length in VB6 is approximately 2GB (theoretical)
//! - Practical limit much lower due to memory constraints
//! - For user-defined types, returns total structure size in bytes
//!
//! ## Limitations
//!
//! - Returns character count, not byte count (use LenB for bytes)
//! - Cannot measure array size directly (use UBound - LBound + 1)
//! - For objects, may return unpredictable values
//! - For fixed-length strings, returns declared length, not trimmed length
//! - Does not distinguish between different whitespace characters
//! - No built-in way to get "visible" length (excluding control characters)
//!
//! ## Related Functions
//!
//! - `LenB`: Get byte length of string or variable
//! - `Left`/`Right`/`Mid`: Extract substrings (often used with Len)
//! - `InStr`/`InStrRev`: Find substring positions
//! - `Trim`/`LTrim`/`RTrim`: Remove whitespace
//! - `String`: Create string of repeated characters

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_len_basic() {
        let source = r#"
Sub Test()
    result = Len(myString)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_string_literal() {
        let source = r#"
Sub Test()
    result = Len("Hello World")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_if_statement() {
        let source = r#"
Sub Test()
    If Len(password) < 8 Then
        MsgBox "Too short"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_function_return() {
        let source = r#"
Function GetLength(text As String) As Long
    GetLength = Len(text)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_for_loop() {
        let source = r#"
Sub Test()
    Dim i As Long
    For i = 1 To Len(text)
        Debug.Print Mid(text, i, 1)
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print Len("Test")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Length: " & Len(text)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_variable_assignment() {
        let source = r#"
Sub Test()
    Dim length As Long
    length = Len(myString)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_property_assignment() {
        let source = r#"
Sub Test()
    obj.Length = Len(obj.Text)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_comparison() {
        let source = r#"
Sub Test()
    If Len(str1) = Len(str2) Then
        MsgBox "Same length"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_arithmetic() {
        let source = r#"
Sub Test()
    position = Len(text) - 5
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_in_class() {
        let source = r#"
Private Sub Class_Initialize()
    m_length = Len(m_text)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_with_statement() {
        let source = r#"
Sub Test()
    With record
        .Size = Len(.Data)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_function_argument() {
        let source = r#"
Sub Test()
    Call ValidateLength(Len(input))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_select_case() {
        let source = r#"
Sub Test()
    Select Case Len(code)
        Case 3
            ProcessShort
        Case 10
            ProcessLong
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_elseif() {
        let source = r#"
Sub Test()
    If Len(text) = 0 Then
        HandleEmpty
    ElseIf Len(text) > 100 Then
        HandleLong
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_iif() {
        let source = r#"
Sub Test()
    result = IIf(Len(text) > 0, text, "Empty")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_parentheses() {
        let source = r#"
Sub Test()
    result = (Len(text))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_array_assignment() {
        let source = r#"
Sub Test()
    lengths(i) = Len(strings(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_collection_add() {
        let source = r#"
Sub Test()
    sizes.Add Len(values(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_concatenation() {
        let source = r#"
Sub Test()
    info = "Length: " & Len(data) & " chars"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_while_wend() {
        let source = r#"
Sub Test()
    While Len(buffer) < maxSize
        buffer = buffer & GetData()
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_do_while() {
        let source = r#"
Sub Test()
    Do While Len(input) > 0
        ProcessChar
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_do_until() {
        let source = r#"
Sub Test()
    Do Until Len(result) >= targetLen
        result = result & "x"
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_with_left() {
        let source = r#"
Sub Test()
    prefix = Left(text, Len(text) - 1)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_validation() {
        let source = r#"
Function IsValidLength(text As String) As Boolean
    IsValidLength = (Len(text) >= 3 And Len(text) <= 50)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_len_empty_check() {
        let source = r#"
Sub Test()
    If Len(Trim(input)) = 0 Then
        MsgBox "Please enter a value"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Len"));
        assert!(text.contains("Identifier"));
    }
}
