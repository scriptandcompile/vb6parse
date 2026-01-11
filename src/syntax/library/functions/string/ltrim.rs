//! # `LTrim` Function
//!
//! Returns a String containing a copy of a specified string with leading spaces removed.
//!
//! ## Syntax
//!
//! ```vb
//! LTrim(string)
//! ```
//!
//! ## Parameters
//!
//! - `string` (Required): String expression from which leading spaces are to be removed
//!   - Can be any valid string expression
//!   - If string is Null, returns Null
//!   - Empty string returns empty string
//!
//! ## Return Value
//!
//! Returns a String (or Variant):
//! - Copy of string with leading spaces removed
//! - Removes only spaces (ASCII 32) from the left
//! - Does not remove tabs, newlines, or other whitespace characters
//! - Returns Null if input is Null
//! - Returns empty string if input is empty or all spaces
//! - Trailing spaces are preserved
//! - Internal spaces are preserved
//!
//! ## Remarks
//!
//! The `LTrim` function removes leading spaces:
//!
//! - Removes only space characters (ASCII 32) from the left side
//! - Does not remove tabs (Chr(9)), line feeds (Chr(10)), or carriage returns (Chr(13))
//! - Does not remove non-breaking spaces or other Unicode whitespace
//! - Trailing spaces are not affected
//! - Internal spaces between words are preserved
//! - Often used to clean up user input
//! - Commonly paired with `RTrim` or used with Trim
//! - Null input returns Null (propagates Null)
//! - Empty string input returns empty string
//! - String of only spaces returns empty string
//! - Does not modify the original string (returns new string)
//! - Can be used with Variant variables
//! - Common in data validation and formatting
//! - Used to normalize text from fixed-width fields
//! - Essential for cleaning imported data
//! - Part of the VB6 string manipulation library
//! - Available in all VB versions
//! - Related to `RTrim` (removes trailing spaces) and Trim (removes both)
//!
//! ## Typical Uses
//!
//! 1. **Remove Leading Spaces**
//!    ```vb
//!    cleanText = LTrim("   Hello")
//!    ```
//!
//! 2. **Clean User Input**
//!    ```vb
//!    userName = LTrim(txtUsername.Text)
//!    ```
//!
//! 3. **Process Fixed-Width Data**
//!    ```vb
//!    field = LTrim(Mid(line, 1, 20))
//!    ```
//!
//! 4. **Normalize Text**
//!    ```vb
//!    normalizedText = LTrim(RTrim(inputText))
//!    ```
//!
//! 5. **Data Import Cleanup**
//!    ```vb
//!    value = LTrim(csvField)
//!    ```
//!
//! 6. **Remove Padding**
//!    ```vb
//!    If LTrim(textBox.Text) = "" Then
//!        MsgBox "Required field"
//!    End If
//!    ```
//!
//! 7. **Format Display**
//!    ```vb
//!    lblName.Caption = LTrim(recordset("Name"))
//!    ```
//!
//! 8. **Conditional Processing**
//!    ```vb
//!    If LTrim(line) <> "" Then
//!        ProcessLine line
//!    End If
//!    ```
//!
//! ## Basic Examples
//!
//! ### Example 1: Basic Usage
//! ```vb
//! Dim result As String
//!
//! result = LTrim("   Hello")           ' Returns "Hello"
//! result = LTrim("Hello   ")           ' Returns "Hello   " (trailing preserved)
//! result = LTrim("   Hello World   ")  ' Returns "Hello World   "
//! result = LTrim("NoSpaces")           ' Returns "NoSpaces"
//! result = LTrim("     ")              ' Returns ""
//! result = LTrim("")                   ' Returns ""
//! ```
//!
//! ### Example 2: Clean User Input
//! ```vb
//! Private Sub txtUsername_LostFocus()
//!     ' Remove leading spaces from input
//!     txtUsername.Text = LTrim(txtUsername.Text)
//!     
//!     ' Validate
//!     If LTrim(txtUsername.Text) = "" Then
//!         MsgBox "Username is required", vbExclamation
//!         txtUsername.SetFocus
//!     End If
//! End Sub
//! ```
//!
//! ### Example 3: Process Fixed-Width File
//! ```vb
//! Sub ProcessFixedWidthFile(ByVal filename As String)
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim firstName As String
//!     Dim lastName As String
//!     
//!     fileNum = FreeFile
//!     Open filename For Input As #fileNum
//!     
//!     Do While Not EOF(fileNum)
//!         Line Input #fileNum, line
//!         
//!         ' Extract fields (positions 1-20 and 21-40)
//!         firstName = LTrim(Mid(line, 1, 20))
//!         lastName = LTrim(Mid(line, 21, 20))
//!         
//!         Debug.Print firstName & " " & lastName
//!     Loop
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Example 4: Text Normalization
//! ```vb
//! Function NormalizeText(ByVal text As String) As String
//!     ' Remove leading and trailing spaces
//!     NormalizeText = LTrim(RTrim(text))
//!     
//!     ' Could also use: NormalizeText = Trim(text)
//! End Function
//!
//! ' Usage
//! Dim clean As String
//! clean = NormalizeText("   Hello World   ")  ' Returns "Hello World"
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `FullTrim` (combine with `RTrim`)
//! ```vb
//! Function FullTrim(ByVal text As String) As String
//!     FullTrim = LTrim(RTrim(text))
//!     ' Note: Can also use built-in Trim() function
//! End Function
//! ```
//!
//! ### Pattern 2: `IsBlank` (check for empty or whitespace)
//! ```vb
//! Function IsBlank(ByVal text As String) As Boolean
//!     IsBlank = (LTrim(RTrim(text)) = "")
//! End Function
//! ```
//!
//! ### Pattern 3: `SafeLTrim` (handle Null)
//! ```vb
//! Function SafeLTrim(ByVal text As Variant) As String
//!     If IsNull(text) Then
//!         SafeLTrim = ""
//!     Else
//!         SafeLTrim = LTrim(text)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 4: `CleanInput`
//! ```vb
//! Function CleanInput(ByVal userInput As String) As String
//!     ' Remove leading/trailing spaces and convert to proper case
//!     CleanInput = LTrim(RTrim(userInput))
//!     If CleanInput <> "" Then
//!         CleanInput = UCase(Left(CleanInput, 1)) & LCase(Mid(CleanInput, 2))
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 5: `TrimFields` (process array)
//! ```vb
//! Sub TrimFields(fields() As String)
//!     Dim i As Integer
//!     For i = LBound(fields) To UBound(fields)
//!         fields(i) = LTrim(RTrim(fields(i)))
//!     Next i
//! End Sub
//! ```
//!
//! ### Pattern 6: `ParsePaddedValue`
//! ```vb
//! Function ParsePaddedValue(ByVal paddedText As String) As String
//!     ' Remove leading spaces from fixed-width field
//!     ParsePaddedValue = LTrim(paddedText)
//! End Function
//! ```
//!
//! ### Pattern 7: `ValidateRequired`
//! ```vb
//! Function ValidateRequired(ByVal fieldValue As String, _
//!                          ByVal fieldName As String) As Boolean
//!     If LTrim(RTrim(fieldValue)) = "" Then
//!         MsgBox fieldName & " is required", vbExclamation
//!         ValidateRequired = False
//!     Else
//!         ValidateRequired = True
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 8: `TrimAllControls`
//! ```vb
//! Sub TrimAllControls(ByVal frm As Form)
//!     Dim ctrl As Control
//!     
//!     For Each ctrl In frm.Controls
//!         If TypeOf ctrl Is TextBox Then
//!             ctrl.Text = LTrim(RTrim(ctrl.Text))
//!         End If
//!     Next ctrl
//! End Sub
//! ```
//!
//! ### Pattern 9: `ParseCSVField`
//! ```vb
//! Function ParseCSVField(ByVal field As String) As String
//!     ' Remove quotes and trim
//!     If Left(field, 1) = """" And Right(field, 1) = """" Then
//!         field = Mid(field, 2, Len(field) - 2)
//!     End If
//!     ParseCSVField = LTrim(RTrim(field))
//! End Function
//! ```
//!
//! ### Pattern 10: `RemoveLeadingSpaces`
//! ```vb
//! Sub RemoveLeadingSpaces(ByVal textBox As TextBox)
//!     Dim selStart As Long
//!     selStart = textBox.SelStart
//!     textBox.Text = LTrim(textBox.Text)
//!     textBox.SelStart = selStart
//! End Sub
//! ```
//!
//! ## Advanced Examples
//!
//! ### Example 1: Data Import Processor
//! ```vb
//! ' Class: DataImporter
//! Private m_data As Collection
//!
//! Public Sub ImportFixedWidthFile(ByVal filename As String)
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim record As Dictionary
//!     
//!     Set m_data = New Collection
//!     
//!     fileNum = FreeFile
//!     Open filename For Input As #fileNum
//!     
//!     Do While Not EOF(fileNum)
//!         Line Input #fileNum, line
//!         
//!         If Len(line) >= 60 Then
//!             Set record = New Dictionary
//!             
//!             ' Extract and trim fields
//!             record("ID") = LTrim(Mid(line, 1, 10))
//!             record("Name") = LTrim(Mid(line, 11, 30))
//!             record("City") = LTrim(Mid(line, 41, 20))
//!             
//!             m_data.Add record
//!         End If
//!     Loop
//!     
//!     Close #fileNum
//! End Sub
//!
//! Public Property Get RecordCount() As Long
//!     RecordCount = m_data.Count
//! End Property
//!
//! Public Function GetRecord(ByVal index As Long) As Dictionary
//!     Set GetRecord = m_data(index)
//! End Function
//! ```
//!
//! ### Example 2: Text Field Validator
//! ```vb
//! ' Class: TextValidator
//! Private m_errors As Collection
//!
//! Public Sub ValidateForm(ByVal frm As Form)
//!     Set m_errors = New Collection
//!     
//!     Dim ctrl As Control
//!     For Each ctrl In frm.Controls
//!         If TypeOf ctrl Is TextBox Then
//!             ValidateTextBox ctrl
//!         End If
//!     Next ctrl
//! End Sub
//!
//! Private Sub ValidateTextBox(ByVal txt As TextBox)
//!     Dim trimmed As String
//!     trimmed = LTrim(RTrim(txt.Text))
//!     
//!     ' Check if required (assuming Tag property indicates required)
//!     If txt.Tag = "Required" Then
//!         If trimmed = "" Then
//!             m_errors.Add "Field '" & txt.Name & "' is required"
//!         End If
//!     End If
//!     
//!     ' Check minimum length
//!     If txt.Tag Like "MinLen:*" Then
//!         Dim minLen As Integer
//!         minLen = Val(Mid(txt.Tag, 8))
//!         
//!         If Len(trimmed) < minLen Then
//!             m_errors.Add "Field '" & txt.Name & "' must be at least " & _
//!                         minLen & " characters"
//!         End If
//!     End If
//! End Sub
//!
//! Public Property Get IsValid() As Boolean
//!     IsValid = (m_errors.Count = 0)
//! End Property
//!
//! Public Property Get Errors() As Collection
//!     Set Errors = m_errors
//! End Property
//! ```
//!
//! ### Example 3: String Utilities Module
//! ```vb
//! ' Module: StringUtils
//!
//! Public Function TrimAll(ByVal text As String) As String
//!     TrimAll = LTrim(RTrim(text))
//! End Function
//!
//! Public Function IsNullOrWhitespace(ByVal text As Variant) As Boolean
//!     If IsNull(text) Then
//!         IsNullOrWhitespace = True
//!     ElseIf VarType(text) = vbString Then
//!         IsNullOrWhitespace = (LTrim(RTrim(text)) = "")
//!     Else
//!         IsNullOrWhitespace = False
//!     End If
//! End Function
//!
//! Public Function NormalizeSpaces(ByVal text As String) As String
//!     Dim result As String
//!     Dim i As Integer
//!     Dim lastWasSpace As Boolean
//!     
//!     ' Remove leading spaces
//!     text = LTrim(text)
//!     
//!     ' Collapse multiple spaces to single space
//!     For i = 1 To Len(text)
//!         If Mid(text, i, 1) = " " Then
//!             If Not lastWasSpace Then
//!                 result = result & " "
//!                 lastWasSpace = True
//!             End If
//!         Else
//!             result = result & Mid(text, i, 1)
//!             lastWasSpace = False
//!         End If
//!     Next i
//!     
//!     NormalizeSpaces = RTrim(result)
//! End Function
//!
//! Public Function CleanTextArray(textArray() As String) As String()
//!     Dim i As Integer
//!     Dim result() As String
//!     
//!     ReDim result(LBound(textArray) To UBound(textArray))
//!     
//!     For i = LBound(textArray) To UBound(textArray)
//!         result(i) = LTrim(RTrim(textArray(i)))
//!     Next i
//!     
//!     CleanTextArray = result
//! End Function
//! ```
//!
//! ### Example 4: Form Input Manager
//! ```vb
//! ' Class: FormInputManager
//! Private m_form As Form
//!
//! Public Sub AttachToForm(ByVal frm As Form)
//!     Set m_form = frm
//! End Sub
//!
//! Public Sub TrimAllInputs()
//!     Dim ctrl As Control
//!     
//!     For Each ctrl In m_form.Controls
//!         If TypeOf ctrl Is TextBox Then
//!             ctrl.Text = LTrim(RTrim(ctrl.Text))
//!         ElseIf TypeOf ctrl Is ComboBox Then
//!             ctrl.Text = LTrim(RTrim(ctrl.Text))
//!         End If
//!     Next ctrl
//! End Sub
//!
//! Public Function ValidateRequired() As Boolean
//!     Dim ctrl As Control
//!     Dim trimmed As String
//!     Dim isValid As Boolean
//!     
//!     isValid = True
//!     
//!     For Each ctrl In m_form.Controls
//!         If TypeOf ctrl Is TextBox Then
//!             If ctrl.Tag = "Required" Then
//!                 trimmed = LTrim(RTrim(ctrl.Text))
//!                 
//!                 If trimmed = "" Then
//!                     MsgBox "Field is required: " & ctrl.Name, vbExclamation
//!                     ctrl.SetFocus
//!                     isValid = False
//!                     Exit For
//!                 End If
//!             End If
//!         End If
//!     Next ctrl
//!     
//!     ValidateRequired = isValid
//! End Function
//!
//! Public Function GetCleanValue(ByVal controlName As String) As String
//!     Dim ctrl As Control
//!     
//!     On Error Resume Next
//!     Set ctrl = m_form.Controls(controlName)
//!     
//!     If Not ctrl Is Nothing Then
//!         If TypeOf ctrl Is TextBox Or TypeOf ctrl Is ComboBox Then
//!             GetCleanValue = LTrim(RTrim(ctrl.Text))
//!         End If
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! ' LTrim handles Null gracefully
//! Dim result As Variant
//! result = LTrim(Null)  ' Returns Null
//!
//! ' Safe trimming with Null check
//! Function SafeTrim(ByVal value As Variant) As String
//!     If IsNull(value) Then
//!         SafeTrim = ""
//!     Else
//!         SafeTrim = LTrim(RTrim(CStr(value)))
//!     End If
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: String trimming is highly optimized
//! - **Creates New String**: Does not modify original (immutable)
//! - **Avoid in Tight Loops**: Cache result if using multiple times
//! - **Use `Trim()` Instead**: If removing both leading and trailing spaces
//!
//! ## Best Practices
//!
//! 1. **Use `Trim()` for both sides** - More efficient than LTrim(RTrim())
//! 2. **Validate before use** - Check for Null if using Variant
//! 3. **Clean user input early** - Trim in validation routines
//! 4. **Cache trimmed values** - Don't call repeatedly in loops
//! 5. **Document expectations** - Clarify if tabs/newlines should be removed
//! 6. **Use with database fields** - Clean imported data
//! 7. **Combine with validation** - Check for empty after trimming
//! 8. **Apply to all text inputs** - Standardize data entry
//! 9. **Consider Unicode** - `LTrim` only removes ASCII space (32)
//! 10. **Test edge cases** - Empty strings, all spaces, Null values
//!
//! ## Comparison with Related Functions
//!
//! | Function | Removes Leading | Removes Trailing | Removes Both |
//! |----------|----------------|------------------|--------------|
//! | **`LTrim`** | Yes | No | No |
//! | **`RTrim`** | No | Yes | No |
//! | **Trim** | Yes | Yes | Yes |
//!
//! ## `LTrim` vs `RTrim` vs Trim
//!
//! ```vb
//! Dim text As String
//! text = "   Hello World   "
//!
//! ' LTrim - removes leading spaces only
//! Debug.Print "[" & LTrim(text) & "]"   ' [Hello World   ]
//!
//! ' RTrim - removes trailing spaces only
//! Debug.Print "[" & RTrim(text) & "]"   ' [   Hello World]
//!
//! ' Trim - removes both leading and trailing
//! Debug.Print "[" & Trim(text) & "]"    ' [Hello World]
//!
//! ' Manual equivalent to Trim
//! Debug.Print "[" & LTrim(RTrim(text)) & "]"  ' [Hello World]
//! ```
//!
//! ## Whitespace Characters
//!
//! ```vb
//! ' LTrim only removes space (ASCII 32)
//! Dim text As String
//!
//! text = "   Hello"        ' Spaces - REMOVED
//! text = Chr(9) & "Hello"  ' Tab - NOT REMOVED
//! text = Chr(10) & "Hello" ' Line feed - NOT REMOVED
//! text = Chr(13) & "Hello" ' Carriage return - NOT REMOVED
//! text = Chr(160) & "Hello" ' Non-breaking space - NOT REMOVED
//!
//! ' To remove other whitespace, use custom function
//! Function TrimAllWhitespace(ByVal text As String) As String
//!     Do While Len(text) > 0
//!         Dim ch As String
//!         ch = Left(text, 1)
//!         
//!         If ch = " " Or ch = Chr(9) Or ch = Chr(10) Or ch = Chr(13) Then
//!             text = Mid(text, 2)
//!         Else
//!             Exit Do
//!         End If
//!     Loop
//!     
//!     TrimAllWhitespace = text
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
//! - `RTrim`: Removes trailing spaces from string
//! - `Trim`: Removes both leading and trailing spaces
//! - `Left`: Returns leftmost characters
//! - `Mid`: Returns substring from middle
//! - `Replace`: Replaces occurrences of substring
//! - `Space`: Creates string of spaces
//! - `Len`: Returns string length

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn ltrim_basic() {
        let source = r#"
            Dim result As String
            result = LTrim("   Hello")
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_variable() {
        let source = r"
            cleaned = LTrim(userInput)
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_textbox() {
        let source = r"
            txtUsername.Text = LTrim(txtUsername.Text)
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_if_statement() {
        let source = r#"
            If LTrim(text) = "" Then
                MsgBox "Empty"
            End If
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_function_return() {
        let source = r"
            Function CleanText(s As String) As String
                CleanText = LTrim(s)
            End Function
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_with_rtrim() {
        let source = r"
            fullTrim = LTrim(RTrim(text))
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_mid_extraction() {
        let source = r"
            field = LTrim(Mid(line, 1, 20))
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_debug_print() {
        let source = r"
            Debug.Print LTrim(text)
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_with_statement() {
        let source = r"
            With record
                .Name = LTrim(.Name)
            End With
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_select_case() {
        let source = r#"
            Select Case LTrim(input)
                Case ""
                    MsgBox "Empty"
                Case Else
                    Process input
            End Select
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_elseif() {
        let source = r#"
            If text = "" Then
                status = "Empty"
            ElseIf LTrim(text) = "" Then
                status = "Whitespace only"
            End If
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_parentheses() {
        let source = r"
            result = (LTrim(text))
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_iif() {
        let source = r#"
            result = IIf(LTrim(text) = "", "Empty", "Has data")
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_in_class() {
        let source = r"
            Private Sub Class_Method()
                m_cleanValue = LTrim(m_rawValue)
            End Sub
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_function_argument() {
        let source = r"
            Call ProcessText(LTrim(input))
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_property_assignment() {
        let source = r"
            MyObject.CleanText = LTrim(dirtyText)
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_array_assignment() {
        let source = r"
            cleanValues(i) = LTrim(rawValues(i))
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_for_loop() {
        let source = r"
            For i = 1 To 10
                fields(i) = LTrim(fields(i))
            Next i
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_while_wend() {
        let source = r"
            While Not EOF(1)
                Line Input #1, line
                line = LTrim(line)
            Wend
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_do_while() {
        let source = r"
            Do While i < count
                text = LTrim(dataArray(i))
                i = i + 1
            Loop
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_do_until() {
        let source = r#"
            Do Until LTrim(input) <> ""
                input = InputBox("Enter text")
            Loop
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_msgbox() {
        let source = r"
            MsgBox LTrim(message)
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_concatenation() {
        let source = r#"
            fullName = LTrim(firstName) & " " & LTrim(lastName)
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_comparison() {
        let source = r#"
            If LTrim(txtInput.Text) <> "" Then
                ProcessInput
            End If
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_label_caption() {
        let source = r#"
            lblName.Caption = LTrim(recordset("Name"))
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_validation() {
        let source = r#"
            If LTrim(RTrim(txtUsername.Text)) = "" Then
                MsgBox "Username required"
            End If
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ltrim_recordset_field() {
        let source = r#"
            customerName = LTrim(rs.Fields("CustomerName").Value)
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/ltrim");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
