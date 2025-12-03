//! # Trim Function
//!
//! Returns a String containing a copy of a specified string with both leading and trailing spaces removed.
//!
//! ## Syntax
//!
//! ```vb
//! Trim(string)
//! ```
//!
//! ## Parameters
//!
//! - `string` (Required): String expression from which leading and trailing spaces are to be removed
//!   - Can be any valid string expression
//!   - If string is Null, returns Null
//!   - Empty string returns empty string
//!
//! ## Return Value
//!
//! Returns a String (or Variant):
//! - Copy of string with both leading and trailing spaces removed
//! - Removes only spaces (ASCII 32) from both ends
//! - Does not remove tabs, newlines, or other whitespace characters
//! - Returns Null if input is Null
//! - Returns empty string if input is empty or all spaces
//! - Internal spaces between words are preserved
//! - Equivalent to LTrim(RTrim(string))
//!
//! ## Remarks
//!
//! The Trim function removes both leading and trailing spaces:
//!
//! - Removes only space characters (ASCII 32) from both ends
//! - Does not remove tabs (Chr(9)), line feeds (Chr(10)), or carriage returns (Chr(13))
//! - Does not remove non-breaking spaces or other Unicode whitespace
//! - Internal spaces between words are preserved
//! - Most commonly used string trimming function
//! - Equivalent to calling LTrim(RTrim(string))
//! - More efficient than using `LTrim` and `RTrim` separately
//! - Null input returns Null (propagates Null)
//! - Empty string input returns empty string
//! - String of only spaces returns empty string
//! - Does not modify the original string (returns new string)
//! - Can be used with Variant variables
//! - Essential for data validation and formatting
//! - Used to normalize user input
//! - Common in database operations
//! - Standard practice before comparing strings
//! - Required for cleaning imported data
//! - Part of the VB6 string manipulation library
//! - Available in all VB versions
//! - Related to `LTrim` (removes leading only) and `RTrim` (removes trailing only)
//!
//! ## Typical Uses
//!
//! 1. **Clean User Input**
//!    ```vb
//!    userName = Trim(txtUsername.Text)
//!    ```
//!
//! 2. **Validate Input**
//!    ```vb
//!    If Trim(txtEmail.Text) = "" Then
//!        MsgBox "Email required"
//!    End If
//!    ```
//!
//! 3. **Compare Strings**
//!    ```vb
//!    If Trim(input) = Trim(expected) Then
//!        result = "Match"
//!    End If
//!    ```
//!
//! 4. **Database Fields**
//!    ```vb
//!    customerName = Trim(rs("CustomerName"))
//!    ```
//!
//! 5. **Split Data**
//!    ```vb
//!    parts = Split(Trim(line), ",")
//!    ```
//!
//! 6. **Clean Array**
//!    ```vb
//!    For i = 0 To UBound(items)
//!        items(i) = Trim(items(i))
//!    Next i
//!    ```
//!
//! 7. **Form Validation**
//!    ```vb
//!    If Len(Trim(txtPassword.Text)) < 8 Then
//!        MsgBox "Password too short"
//!    End If
//!    ```
//!
//! 8. **Process CSV**
//!    ```vb
//!    fields = Split(line, ",")
//!    firstName = Trim(fields(0))
//!    lastName = Trim(fields(1))
//!    ```
//!
//! ## Basic Examples
//!
//! ### Example 1: Basic Usage
//! ```vb
//! Dim result As String
//!
//! result = Trim("   Hello   ")        ' Returns "Hello"
//! result = Trim("   Hello World   ")  ' Returns "Hello World"
//! result = Trim("NoSpaces")           ' Returns "NoSpaces"
//! result = Trim("     ")              ' Returns ""
//! result = Trim("")                   ' Returns ""
//! ```
//!
//! ### Example 2: Form Validation
//! ```vb
//! Private Sub cmdSubmit_Click()
//!     Dim userName As String
//!     Dim password As String
//!     
//!     ' Clean inputs
//!     userName = Trim(txtUsername.Text)
//!     password = Trim(txtPassword.Text)
//!     
//!     ' Validate
//!     If userName = "" Then
//!         MsgBox "Username is required", vbExclamation
//!         txtUsername.SetFocus
//!         Exit Sub
//!     End If
//!     
//!     If password = "" Then
//!         MsgBox "Password is required", vbExclamation
//!         txtPassword.SetFocus
//!         Exit Sub
//!     End If
//!     
//!     If Len(password) < 8 Then
//!         MsgBox "Password must be at least 8 characters", vbExclamation
//!         txtPassword.SetFocus
//!         Exit Sub
//!     End If
//!     
//!     ' Process login
//!     ProcessLogin userName, password
//! End Sub
//! ```
//!
//! ### Example 3: CSV Processing
//! ```vb
//! Sub ProcessCSVFile(ByVal filename As String)
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim fields() As String
//!     Dim i As Integer
//!     
//!     fileNum = FreeFile
//!     Open filename For Input As #fileNum
//!     
//!     Do While Not EOF(fileNum)
//!         Line Input #fileNum, line
//!         
//!         ' Split and trim each field
//!         fields = Split(line, ",")
//!         For i = 0 To UBound(fields)
//!             fields(i) = Trim(fields(i))
//!         Next i
//!         
//!         ' Process the cleaned fields
//!         If UBound(fields) >= 2 Then
//!             Debug.Print "Name: " & fields(0) & ", " & fields(1)
//!             Debug.Print "Age: " & fields(2)
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Example 4: Database Operations
//! ```vb
//! Sub LoadCustomers()
//!     Dim rs As ADODB.Recordset
//!     Dim customerName As String
//!     
//!     Set rs = New ADODB.Recordset
//!     rs.Open "SELECT * FROM Customers", conn
//!     
//!     Do While Not rs.EOF
//!         ' Trim database fields (may have padding)
//!         customerName = Trim(rs("CustomerName") & "")
//!         
//!         If customerName <> "" Then
//!             lstCustomers.AddItem customerName
//!         End If
//!         
//!         rs.MoveNext
//!     Loop
//!     
//!     rs.Close
//!     Set rs = Nothing
//! End Sub
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `IsBlank`
//! ```vb
//! Function IsBlank(ByVal text As String) As Boolean
//!     IsBlank = (Trim(text) = "")
//! End Function
//! ```
//!
//! ### Pattern 2: `SafeTrim` (handle Null)
//! ```vb
//! Function SafeTrim(ByVal text As Variant) As String
//!     If IsNull(text) Then
//!         SafeTrim = ""
//!     Else
//!         SafeTrim = Trim(CStr(text))
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 3: `TrimAndUpper`
//! ```vb
//! Function TrimAndUpper(ByVal text As String) As String
//!     TrimAndUpper = UCase(Trim(text))
//! End Function
//! ```
//!
//! ### Pattern 4: `ValidateRequired`
//! ```vb
//! Function ValidateRequired(ByVal value As String, _
//!                          ByVal fieldName As String) As Boolean
//!     If Trim(value) = "" Then
//!         MsgBox fieldName & " is required", vbExclamation
//!         ValidateRequired = False
//!     Else
//!         ValidateRequired = True
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 5: `TrimArray`
//! ```vb
//! Sub TrimArray(arr() As String)
//!     Dim i As Integer
//!     For i = LBound(arr) To UBound(arr)
//!         arr(i) = Trim(arr(i))
//!     Next i
//! End Sub
//! ```
//!
//! ### Pattern 6: `CompareIgnoreSpaces`
//! ```vb
//! Function CompareIgnoreSpaces(ByVal str1 As String, _
//!                             ByVal str2 As String) As Boolean
//!     CompareIgnoreSpaces = (Trim(str1) = Trim(str2))
//! End Function
//! ```
//!
//! ### Pattern 7: `TrimAllControls`
//! ```vb
//! Sub TrimAllControls(ByVal frm As Form)
//!     Dim ctrl As Control
//!     
//!     For Each ctrl In frm.Controls
//!         If TypeOf ctrl Is TextBox Then
//!             ctrl.Text = Trim(ctrl.Text)
//!         ElseIf TypeOf ctrl Is ComboBox Then
//!             ctrl.Text = Trim(ctrl.Text)
//!         End If
//!     Next ctrl
//! End Sub
//! ```
//!
//! ### Pattern 8: `GetCleanValue`
//! ```vb
//! Function GetCleanValue(ByVal ctrl As Control) As String
//!     If TypeOf ctrl Is TextBox Or TypeOf ctrl Is ComboBox Then
//!         GetCleanValue = Trim(ctrl.Text)
//!     Else
//!         GetCleanValue = ""
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 9: `SplitAndTrim`
//! ```vb
//! Function SplitAndTrim(ByVal text As String, _
//!                       ByVal delimiter As String) As String()
//!     Dim parts() As String
//!     Dim i As Integer
//!     
//!     parts = Split(text, delimiter)
//!     For i = 0 To UBound(parts)
//!         parts(i) = Trim(parts(i))
//!     Next i
//!     
//!     SplitAndTrim = parts
//! End Function
//! ```
//!
//! ### Pattern 10: `DefaultIfBlank`
//! ```vb
//! Function DefaultIfBlank(ByVal value As String, _
//!                         ByVal defaultValue As String) As String
//!     If Trim(value) = "" Then
//!         DefaultIfBlank = defaultValue
//!     Else
//!         DefaultIfBlank = Trim(value)
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Examples
//!
//! ### Example 1: Form Validator Class
//! ```vb
//! ' Class: FormValidator
//! Private m_form As Form
//! Private m_errors As Collection
//!
//! Public Sub AttachForm(ByVal frm As Form)
//!     Set m_form = frm
//! End Sub
//!
//! Public Function Validate() As Boolean
//!     Set m_errors = New Collection
//!     
//!     Dim ctrl As Control
//!     For Each ctrl In m_form.Controls
//!         If TypeOf ctrl Is TextBox Then
//!             ValidateTextBox ctrl
//!         End If
//!     Next ctrl
//!     
//!     Validate = (m_errors.Count = 0)
//! End Function
//!
//! Private Sub ValidateTextBox(ByVal txt As TextBox)
//!     Dim value As String
//!     value = Trim(txt.Text)
//!     
//!     ' Check required
//!     If txt.Tag Like "*Required*" Then
//!         If value = "" Then
//!             m_errors.Add "Field '" & txt.Name & "' is required"
//!             Exit Sub
//!         End If
//!     End If
//!     
//!     ' Check minimum length
//!     If txt.Tag Like "*MinLen:*" Then
//!         Dim minLen As Integer
//!         minLen = ExtractNumber(txt.Tag, "MinLen:")
//!         
//!         If Len(value) < minLen Then
//!             m_errors.Add "Field '" & txt.Name & "' must be at least " & _
//!                         minLen & " characters"
//!         End If
//!     End If
//!     
//!     ' Check email format
//!     If txt.Tag Like "*Email*" Then
//!         If value <> "" And InStr(value, "@") = 0 Then
//!             m_errors.Add "Field '" & txt.Name & "' must be a valid email"
//!         End If
//!     End If
//! End Sub
//!
//! Public Property Get Errors() As Collection
//!     Set Errors = m_errors
//! End Property
//!
//! Public Sub ShowErrors()
//!     Dim msg As String
//!     Dim err As Variant
//!     
//!     For Each err In m_errors
//!         msg = msg & err & vbCrLf
//!     Next err
//!     
//!     If msg <> "" Then
//!         MsgBox msg, vbExclamation, "Validation Errors"
//!     End If
//! End Sub
//! ```
//!
//! ### Example 2: Data Cleaner Class
//! ```vb
//! ' Class: DataCleaner
//! Private m_trimSpaces As Boolean
//! Private m_removeLineBreaks As Boolean
//! Private m_collapseSpaces As Boolean
//!
//! Private Sub Class_Initialize()
//!     m_trimSpaces = True
//!     m_removeLineBreaks = False
//!     m_collapseSpaces = False
//! End Sub
//!
//! Public Property Let TrimSpaces(ByVal value As Boolean)
//!     m_trimSpaces = value
//! End Property
//!
//! Public Property Let RemoveLineBreaks(ByVal value As Boolean)
//!     m_removeLineBreaks = value
//! End Property
//!
//! Public Property Let CollapseSpaces(ByVal value As Boolean)
//!     m_collapseSpaces = value
//! End Property
//!
//! Public Function CleanString(ByVal text As String) As String
//!     Dim result As String
//!     result = text
//!     
//!     ' Remove line breaks if requested
//!     If m_removeLineBreaks Then
//!         result = Replace(result, vbCrLf, " ")
//!         result = Replace(result, vbCr, " ")
//!         result = Replace(result, vbLf, " ")
//!     End If
//!     
//!     ' Collapse multiple spaces if requested
//!     If m_collapseSpaces Then
//!         Do While InStr(result, "  ") > 0
//!             result = Replace(result, "  ", " ")
//!         Loop
//!     End If
//!     
//!     ' Trim spaces if requested
//!     If m_trimSpaces Then
//!         result = Trim(result)
//!     End If
//!     
//!     CleanString = result
//! End Function
//!
//! Public Function CleanArray(arr() As String) As String()
//!     Dim result() As String
//!     Dim i As Integer
//!     
//!     ReDim result(LBound(arr) To UBound(arr))
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         result(i) = CleanString(arr(i))
//!     Next i
//!     
//!     CleanArray = result
//! End Function
//! ```
//!
//! ### Example 3: CSV Parser
//! ```vb
//! ' Class: CSVParser
//! Private m_rows As Collection
//!
//! Public Sub ParseFile(ByVal filename As String)
//!     Dim fileNum As Integer
//!     Dim line As String
//!     
//!     Set m_rows = New Collection
//!     
//!     fileNum = FreeFile
//!     Open filename For Input As #fileNum
//!     
//!     Do While Not EOF(fileNum)
//!         Line Input #fileNum, line
//!         
//!         If Trim(line) <> "" Then
//!             m_rows.Add ParseLine(line)
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//! End Sub
//!
//! Private Function ParseLine(ByVal line As String) As Collection
//!     Dim fields As Collection
//!     Dim parts() As String
//!     Dim i As Integer
//!     Dim field As String
//!     
//!     Set fields = New Collection
//!     parts = Split(line, ",")
//!     
//!     For i = 0 To UBound(parts)
//!         field = Trim(parts(i))
//!         
//!         ' Remove quotes if present
//!         If Left(field, 1) = """" And Right(field, 1) = """" Then
//!             field = Mid(field, 2, Len(field) - 2)
//!         End If
//!         
//!         fields.Add field
//!     Next i
//!     
//!     Set ParseLine = fields
//! End Function
//!
//! Public Property Get RowCount() As Long
//!     RowCount = m_rows.Count
//! End Property
//!
//! Public Function GetRow(ByVal index As Long) As Collection
//!     Set GetRow = m_rows(index)
//! End Function
//!
//! Public Function GetField(ByVal row As Long, _
//!                          ByVal col As Long) As String
//!     Dim rowData As Collection
//!     Set rowData = m_rows(row)
//!     GetField = rowData(col)
//! End Function
//! ```
//!
//! ### Example 4: String Utilities Module
//! ```vb
//! ' Module: StringUtils
//!
//! Public Function IsNullOrWhitespace(ByVal text As Variant) As Boolean
//!     If IsNull(text) Then
//!         IsNullOrWhitespace = True
//!     ElseIf VarType(text) = vbString Then
//!         IsNullOrWhitespace = (Trim(text) = "")
//!     Else
//!         IsNullOrWhitespace = False
//!     End If
//! End Function
//!
//! Public Function CoalesceString(ParamArray values() As Variant) As String
//!     Dim i As Integer
//!     Dim value As String
//!     
//!     For i = LBound(values) To UBound(values)
//!         If Not IsNull(values(i)) Then
//!             value = Trim(CStr(values(i)))
//!             If value <> "" Then
//!                 CoalesceString = value
//!                 Exit Function
//!             End If
//!         End If
//!     Next i
//!     
//!     CoalesceString = ""
//! End Function
//!
//! Public Function JoinTrimmed(arr() As String, _
//!                            ByVal delimiter As String) As String
//!     Dim result As String
//!     Dim i As Integer
//!     Dim value As String
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         value = Trim(arr(i))
//!         If value <> "" Then
//!             If result <> "" Then
//!                 result = result & delimiter
//!             End If
//!             result = result & value
//!         End If
//!     Next i
//!     
//!     JoinTrimmed = result
//! End Function
//!
//! Public Function NormalizeWhitespace(ByVal text As String) As String
//!     Dim result As String
//!     
//!     ' Replace all whitespace with single space
//!     result = text
//!     result = Replace(result, vbTab, " ")
//!     result = Replace(result, vbCrLf, " ")
//!     result = Replace(result, vbCr, " ")
//!     result = Replace(result, vbLf, " ")
//!     
//!     ' Collapse multiple spaces
//!     Do While InStr(result, "  ") > 0
//!         result = Replace(result, "  ", " ")
//!     Loop
//!     
//!     ' Trim
//!     NormalizeWhitespace = Trim(result)
//! End Function
//!
//! Public Function TruncateWithEllipsis(ByVal text As String, _
//!                                      ByVal maxLength As Integer) As String
//!     text = Trim(text)
//!     
//!     If Len(text) <= maxLength Then
//!         TruncateWithEllipsis = text
//!     Else
//!         TruncateWithEllipsis = Left(text, maxLength - 3) & "..."
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! ' Trim handles Null gracefully
//! Dim result As Variant
//! result = Trim(Null)  ' Returns Null
//!
//! ' Safe trimming with Null check
//! Function SafeTrim(ByVal value As Variant) As String
//!     If IsNull(value) Then
//!         SafeTrim = ""
//!     Else
//!         SafeTrim = Trim(CStr(value))
//!     End If
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: String trimming is highly optimized
//! - **Creates New String**: Does not modify original (immutable)
//! - **More Efficient**: Using `Trim()` is faster than LTrim(RTrim())
//! - **Cache Results**: Don't call repeatedly in tight loops
//!
//! ## Best Practices
//!
//! 1. **Always trim user input** - Before validation or storage
//! 2. **Trim before comparison** - Ensure consistent string matching
//! 3. **Use with database fields** - Clean data from external sources
//! 4. **Validate after trimming** - Check if empty after removing spaces
//! 5. **Combine with other operations** - Trim then convert case, etc.
//! 6. **Cache trimmed values** - Don't call repeatedly
//! 7. **Handle Null gracefully** - Use `SafeTrim` for Variant types
//! 8. **Document expectations** - Clarify if other whitespace should be removed
//! 9. **Standardize input early** - Trim at entry point
//! 10. **Test edge cases** - Empty strings, all spaces, Null values
//!
//! ## Comparison with Related Functions
//!
//! | Function | Removes Leading | Removes Trailing | Removes Both |
//! |----------|----------------|------------------|--------------|
//! | **Trim** | Yes | Yes | Yes |
//! | **`LTrim`** | Yes | No | No |
//! | **`RTrim`** | No | Yes | No |
//!
//! ## Trim vs `LTrim` vs `RTrim`
//!
//! ```vb
//! Dim text As String
//! text = "   Hello World   "
//!
//! ' Trim - removes both leading and trailing
//! Debug.Print "[" & Trim(text) & "]"       ' [Hello World]
//!
//! ' LTrim - removes leading spaces only
//! Debug.Print "[" & LTrim(text) & "]"      ' [Hello World   ]
//!
//! ' RTrim - removes trailing spaces only
//! Debug.Print "[" & RTrim(text) & "]"      ' [   Hello World]
//!
//! ' Trim is equivalent to LTrim(RTrim())
//! Debug.Print "[" & LTrim(RTrim(text)) & "]"  ' [Hello World]
//! ```
//!
//! ## Whitespace Characters
//!
//! ```vb
//! ' Trim only removes space (ASCII 32)
//! Dim text As String
//!
//! text = "   Hello   "     ' Spaces - REMOVED
//! text = Chr(9) & "Hello" & Chr(9)  ' Tabs - NOT REMOVED
//! text = Chr(10) & "Hello" & Chr(10) ' Line feeds - NOT REMOVED
//! text = Chr(13) & "Hello" & Chr(13) ' CR - NOT REMOVED
//!
//! ' To remove all whitespace:
//! Function TrimAllWhitespace(ByVal text As String) As String
//!     ' Remove line breaks
//!     text = Replace(text, vbCrLf, "")
//!     text = Replace(text, vbCr, "")
//!     text = Replace(text, vbLf, "")
//!     text = Replace(text, vbTab, "")
//!     
//!     ' Trim spaces
//!     TrimAllWhitespace = Trim(text)
//! End Function
//! ```
//!
//! ## Platform Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core library
//! - Works with ANSI and Unicode strings
//! - Only removes ASCII space character (32)
//! - Returns new string (original unchanged)
//! - Handles Null by returning Null
//! - Available in `VBScript`
//! - Same behavior across all Windows versions
//! - More efficient than LTrim(RTrim())
//!
//! ## Limitations
//!
//! - **Only Space Character**: Does not remove tabs, line feeds, etc.
//! - **No Unicode Whitespace**: Does not remove non-breaking spaces, em spaces, etc.
//! - **Creates New String**: Cannot modify string in place
//! - **No Custom Characters**: Cannot specify which characters to remove
//! - **Null Propagation**: Returns Null if input is Null
//!
//! ## Related Functions
//!
//! - `LTrim`: Removes leading spaces from string
//! - `RTrim`: Removes trailing spaces from string
//! - `Left`: Returns leftmost characters
//! - `Right`: Returns rightmost characters
//! - `Mid`: Returns substring from middle
//! - `Replace`: Replaces occurrences of substring
//! - `Space`: Creates string of spaces
//! - `Len`: Returns string length

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_trim_basic() {
        let source = r#"
            Dim result As String
            result = Trim("   Hello   ")
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_variable() {
        let source = r#"
            cleaned = Trim(userInput)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_textbox() {
        let source = r#"
            txtUsername.Text = Trim(txtUsername.Text)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_if_statement() {
        let source = r#"
            If Trim(text) = "" Then
                MsgBox "Empty"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_function_return() {
        let source = r#"
            Function CleanText(s As String) As String
                CleanText = Trim(s)
            End Function
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_validation() {
        let source = r#"
            If Len(Trim(txtPassword.Text)) < 8 Then
                MsgBox "Too short"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_comparison() {
        let source = r#"
            If Trim(input) = Trim(expected) Then
                result = "Match"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_debug_print() {
        let source = r#"
            Debug.Print Trim(text)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_with_statement() {
        let source = r#"
            With record
                .Name = Trim(.Name)
            End With
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_select_case() {
        let source = r#"
            Select Case Trim(input)
                Case ""
                    MsgBox "Empty"
                Case Else
                    Process input
            End Select
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_elseif() {
        let source = r#"
            If text = "" Then
                status = "Empty"
            ElseIf Trim(text) = "" Then
                status = "Whitespace only"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_parentheses() {
        let source = r#"
            result = (Trim(text))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_iif() {
        let source = r#"
            result = IIf(Trim(text) = "", "Empty", "Has data")
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_in_class() {
        let source = r#"
            Private Sub Class_Method()
                m_cleanValue = Trim(m_rawValue)
            End Sub
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_function_argument() {
        let source = r#"
            Call ProcessText(Trim(input))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_property_assignment() {
        let source = r#"
            MyObject.CleanText = Trim(dirtyText)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_array_assignment() {
        let source = r#"
            cleanValues(i) = Trim(rawValues(i))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_for_loop() {
        let source = r#"
            For i = 0 To UBound(items)
                items(i) = Trim(items(i))
            Next i
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_while_wend() {
        let source = r#"
            While Not EOF(1)
                Line Input #1, line
                line = Trim(line)
            Wend
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_do_while() {
        let source = r#"
            Do While i < count
                text = Trim(dataArray(i))
                i = i + 1
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_do_until() {
        let source = r#"
            Do Until Trim(input) <> ""
                input = InputBox("Enter text")
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_msgbox() {
        let source = r#"
            MsgBox Trim(message)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_split() {
        let source = r#"
            parts = Split(Trim(line), ",")
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_database_field() {
        let source = r#"
            customerName = Trim(rs("CustomerName"))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_ucase() {
        let source = r#"
            upperText = UCase(Trim(text))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_concatenation() {
        let source = r#"
            fullName = Trim(firstName) & " " & Trim(lastName)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_trim_label_caption() {
        let source = r#"
            lblName.Caption = Trim(rs.Fields("Name").Value)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Trim"));
        assert!(text.contains("Identifier"));
    }
}
