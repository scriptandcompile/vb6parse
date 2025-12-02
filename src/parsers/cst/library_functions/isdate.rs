//! # `IsDate` Function
//!
//! Returns a Boolean value indicating whether an expression can be converted to a date.
//!
//! ## Syntax
//!
//! ```vb
//! IsDate(expression)
//! ```
//!
//! ## Parameters
//!
//! - `expression` (Required): `Variant` expression to test for date validity
//!
//! ## Return Value
//!
//! Returns a `Boolean`:
//! - `True` if the expression is a date or can be recognized as a valid date
//! - `False` if the expression cannot be converted to a date
//! - Recognizes dates in various formats based on locale settings
//! - Returns `True` for `Date` data type variables
//! - Returns `True` for valid date strings
//! - Returns `False` for `Null`, `Empty`, and invalid date expressions
//!
//! ## Remarks
//!
//! The `IsDate` function determines whether an expression represents a valid date:
//!
//! - Returns `True` for `Date` type variables
//! - Returns `True` for strings that can be interpreted as valid dates
//! - `Date` format recognition depends on locale settings of the system
//! - Recognizes many common `Date` formats (MM/DD/YYYY, DD-MMM-YYYY, etc.)
//! - Returns `False` for `Null` values
//! - Returns `False` for `Empty` variants
//! - Can validate user input before date conversion
//! - Useful for preventing Type Mismatch errors with date operations
//! - Returns `True` for date/time combinations
//! - Returns `True` for time-only values
//! - `Date` range must be valid (e.g., not February 30)
//! - Years typically need to be in valid range (100-9999)
//!
//! ## Typical Uses
//!
//! 1. **Input Validation**: Verify user-entered dates before conversion
//! 2. **Data Type Checking**: Determine if Variant contains date data
//! 3. **Error Prevention**: Avoid Type Mismatch errors in date operations
//! 4. **Database Input**: Validate dates before database insertion
//! 5. **Form Validation**: Check date fields on data entry forms
//! 6. **File Processing**: Validate date columns when importing data
//! 7. **Report Generation**: Ensure date parameters are valid
//! 8. **Date Parsing**: Test multiple date format attempts
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Simple date validation
//! Dim testValue As Variant
//!
//! testValue = "12/25/2025"
//! If IsDate(testValue) Then
//!     Debug.Print "Valid date"  ' This prints
//! End If
//!
//! testValue = "Not a date"
//! If IsDate(testValue) Then
//!     Debug.Print "Valid date"
//! Else
//!     Debug.Print "Invalid date"  ' This prints
//! End If
//!
//! ' Example 2: Validate user input
//! Dim userInput As String
//! userInput = InputBox("Enter a date:")
//!
//! If IsDate(userInput) Then
//!     Dim dateValue As Date
//!     dateValue = CDate(userInput)
//!     MsgBox "You entered: " & Format$(dateValue, "Long Date")
//! Else
//!     MsgBox "Invalid date format", vbExclamation
//! End If
//!
//! ' Example 3: Check various date formats
//! Debug.Print IsDate("12/25/2025")      ' True (MM/DD/YYYY)
//! Debug.Print IsDate("25-Dec-2025")     ' True (DD-MMM-YYYY)
//! Debug.Print IsDate("December 25, 2025") ' True (long format)
//! Debug.Print IsDate("2025-12-25")      ' True (ISO format)
//! Debug.Print IsDate("12:30 PM")        ' True (time only)
//! Debug.Print IsDate("13/45/2025")      ' False (invalid month)
//! Debug.Print IsDate("February 30, 2025") ' False (invalid day)
//!
//! ' Example 4: Validate before date arithmetic
//! Function AddDaysToDate(dateStr As String, days As Long) As Variant
//!     If IsDate(dateStr) Then
//!         AddDaysToDate = CDate(dateStr) + days
//!     Else
//!         AddDaysToDate = Null
//!         MsgBox "Invalid date: " & dateStr, vbExclamation
//!     End If
//! End Function
//!
//! Dim result As Variant
//! result = AddDaysToDate("12/25/2025", 7)
//! If Not IsNull(result) Then
//!     Debug.Print "New date: " & result
//! End If
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Safe date conversion
//! Function SafeCDate(value As Variant) As Variant
//!     If IsDate(value) Then
//!         SafeCDate = CDate(value)
//!     Else
//!         SafeCDate = Null
//!     End If
//! End Function
//!
//! ' Pattern 2: Validate and format date
//! Function FormatIfDate(value As Variant, formatString As String) As String
//!     If IsDate(value) Then
//!         FormatIfDate = Format$(CDate(value), formatString)
//!     Else
//!         FormatIfDate = "N/A"
//!     End If
//! End Function
//!
//! ' Pattern 3: Validate date range
//! Function IsValidDateInRange(dateStr As String, minDate As Date, maxDate As Date) As Boolean
//!     Dim testDate As Date
//!     
//!     If Not IsDate(dateStr) Then
//!         IsValidDateInRange = False
//!         Exit Function
//!     End If
//!     
//!     testDate = CDate(dateStr)
//!     IsValidDateInRange = (testDate >= minDate And testDate <= maxDate)
//! End Function
//!
//! ' Pattern 4: Parse flexible date input
//! Function ParseDate(value As Variant) As Variant
//!     If IsDate(value) Then
//!         ParseDate = CDate(value)
//!     ElseIf IsNumeric(value) Then
//!         ' Try treating as Excel serial date
//!         On Error Resume Next
//!         ParseDate = CDate(CDbl(value))
//!         If Err.Number <> 0 Then ParseDate = Null
//!         On Error GoTo 0
//!     Else
//!         ParseDate = Null
//!     End If
//! End Function
//!
//! ' Pattern 5: Validate date field
//! Function ValidateDateField(fieldValue As Variant, fieldName As String) As Boolean
//!     If IsNull(fieldValue) Or IsEmpty(fieldValue) Then
//!         MsgBox fieldName & " is required", vbExclamation
//!         ValidateDateField = False
//!     ElseIf Not IsDate(fieldValue) Then
//!         MsgBox fieldName & " must be a valid date", vbExclamation
//!         ValidateDateField = False
//!     Else
//!         ValidateDateField = True
//!     End If
//! End Function
//!
//! ' Pattern 6: Extract date from mixed data
//! Function ExtractDate(data As Variant) As Variant
//!     If IsDate(data) Then
//!         ExtractDate = CDate(data)
//!     ElseIf VarType(data) = vbString Then
//!         ' Try to extract date from string
//!         Dim parts() As String
//!         parts = Split(data, " ")
//!         
//!         Dim i As Integer
//!         For i = 0 To UBound(parts)
//!             If IsDate(parts(i)) Then
//!                 ExtractDate = CDate(parts(i))
//!                 Exit Function
//!             End If
//!         Next i
//!         
//!         ExtractDate = Null
//!     Else
//!         ExtractDate = Null
//!     End If
//! End Function
//!
//! ' Pattern 7: Type-safe date comparison
//! Function CompareDates(date1 As Variant, date2 As Variant) As Integer
//!     If Not IsDate(date1) Or Not IsDate(date2) Then
//!         CompareDates = 0  ' Invalid comparison
//!         Exit Function
//!     End If
//!     
//!     Dim d1 As Date, d2 As Date
//!     d1 = CDate(date1)
//!     d2 = CDate(date2)
//!     
//!     If d1 < d2 Then
//!         CompareDates = -1
//!     ElseIf d1 > d2 Then
//!         CompareDates = 1
//!     Else
//!         CompareDates = 0
//!     End If
//! End Function
//!
//! ' Pattern 8: Validate date before database insert
//! Function InsertRecord(recordDate As Variant, description As String) As Boolean
//!     If Not IsDate(recordDate) Then
//!         MsgBox "Invalid date for record", vbCritical
//!         InsertRecord = False
//!         Exit Function
//!     End If
//!     
//!     ' Proceed with database insert
//!     Dim sql As String
//!     sql = "INSERT INTO Records (RecordDate, Description) VALUES (" & _
//!           "#" & CDate(recordDate) & "#, '" & description & "')"
//!     
//!     ' Execute SQL...
//!     InsertRecord = True
//! End Function
//!
//! ' Pattern 9: Handle multiple date formats
//! Function TryParseDateFormats(dateStr As String) As Variant
//!     Dim formats As Variant
//!     Dim i As Integer
//!     
//!     ' Try direct conversion first
//!     If IsDate(dateStr) Then
//!         TryParseDateFormats = CDate(dateStr)
//!         Exit Function
//!     End If
//!     
//!     ' Try reformatting
//!     formats = Array("MM/DD/YYYY", "DD/MM/YYYY", "YYYY-MM-DD")
//!     ' Would need custom parsing logic here
//!     
//!     TryParseDateFormats = Null
//! End Function
//!
//! ' Pattern 10: Validate array of dates
//! Function ValidateDateArray(dates As Variant) As Boolean
//!     Dim i As Long
//!     
//!     If Not IsArray(dates) Then
//!         ValidateDateArray = False
//!         Exit Function
//!     End If
//!     
//!     For i = LBound(dates) To UBound(dates)
//!         If Not IsDate(dates(i)) Then
//!             ValidateDateArray = False
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     ValidateDateArray = True
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Date validation class
//! Public Class DateValidator
//!     Private m_minDate As Date
//!     Private m_maxDate As Date
//!     Private m_allowEmpty As Boolean
//!     
//!     Public Sub Initialize(Optional minDate As Date, Optional maxDate As Date, _
//!                          Optional allowEmpty As Boolean = False)
//!         m_minDate = minDate
//!         m_maxDate = maxDate
//!         m_allowEmpty = allowEmpty
//!     End Sub
//!     
//!     Public Function Validate(value As Variant) As Boolean
//!         ' Check for empty
//!         If IsEmpty(value) Or IsNull(value) Or value = "" Then
//!             Validate = m_allowEmpty
//!             Exit Function
//!         End If
//!         
//!         ' Check if date
//!         If Not IsDate(value) Then
//!             Validate = False
//!             Exit Function
//!         End If
//!         
//!         ' Check range
//!         Dim dateValue As Date
//!         dateValue = CDate(value)
//!         
//!         If m_minDate <> 0 And dateValue < m_minDate Then
//!             Validate = False
//!             Exit Function
//!         End If
//!         
//!         If m_maxDate <> 0 And dateValue > m_maxDate Then
//!             Validate = False
//!             Exit Function
//!         End If
//!         
//!         Validate = True
//!     End Function
//!     
//!     Public Function GetErrorMessage(value As Variant) As String
//!         If IsEmpty(value) Or IsNull(value) Or value = "" Then
//!             If Not m_allowEmpty Then
//!                 GetErrorMessage = "Date is required"
//!             End If
//!         ElseIf Not IsDate(value) Then
//!             GetErrorMessage = "Invalid date format: " & value
//!         Else
//!             Dim dateValue As Date
//!             dateValue = CDate(value)
//!             
//!             If m_minDate <> 0 And dateValue < m_minDate Then
//!                 GetErrorMessage = "Date must be on or after " & m_minDate
//!             ElseIf m_maxDate <> 0 And dateValue > m_maxDate Then
//!                 GetErrorMessage = "Date must be on or before " & m_maxDate
//!             End If
//!         End If
//!     End Function
//! End Class
//!
//! ' Example 2: Smart date parser
//! Public Class SmartDateParser
//!     Public Function Parse(value As Variant) As Variant
//!         ' Handle different input types
//!         If IsDate(value) Then
//!             Parse = CDate(value)
//!             Exit Function
//!         End If
//!         
//!         If IsNumeric(value) Then
//!             ' Try as Excel serial date
//!             On Error Resume Next
//!             Parse = CDate(CDbl(value))
//!             If Err.Number = 0 Then Exit Function
//!             On Error GoTo 0
//!         End If
//!         
//!         If VarType(value) = vbString Then
//!             Parse = ParseString(CStr(value))
//!         Else
//!             Parse = Null
//!         End If
//!     End Function
//!     
//!     Private Function ParseString(str As String) As Variant
//!         ' Remove common prefixes
//!         str = Replace(str, "Date:", "")
//!         str = Trim(str)
//!         
//!         If IsDate(str) Then
//!             ParseString = CDate(str)
//!             Exit Function
//!         End If
//!         
//!         ' Try common transformations
//!         str = Replace(str, ".", "/")  ' 12.25.2025 -> 12/25/2025
//!         If IsDate(str) Then
//!             ParseString = CDate(str)
//!             Exit Function
//!         End If
//!         
//!         ParseString = Null
//!     End Function
//! End Class
//!
//! ' Example 3: Form field validator
//! Public Class FormValidator
//!     Public Function ValidateAllDates(frm As Form) As Boolean
//!         Dim ctl As Control
//!         Dim invalidFields As String
//!         
//!         For Each ctl In frm.Controls
//!             If TypeOf ctl Is TextBox Then
//!                 If ctl.Tag = "DATE" Then
//!                     If ctl.Text <> "" And Not IsDate(ctl.Text) Then
//!                         invalidFields = invalidFields & ctl.Name & vbCrLf
//!                     End If
//!                 End If
//!             End If
//!         Next ctl
//!         
//!         If invalidFields <> "" Then
//!             MsgBox "Invalid date fields:" & vbCrLf & invalidFields, vbExclamation
//!             ValidateAllDates = False
//!         Else
//!             ValidateAllDates = True
//!         End If
//!     End Function
//!     
//!     Public Sub HighlightInvalidDates(frm As Form)
//!         Dim ctl As Control
//!         
//!         For Each ctl In frm.Controls
//!             If TypeOf ctl Is TextBox Then
//!                 If ctl.Tag = "DATE" And ctl.Text <> "" Then
//!                     If IsDate(ctl.Text) Then
//!                         ctl.BackColor = vbWhite
//!                     Else
//!                         ctl.BackColor = RGB(255, 200, 200)  ' Light red
//!                     End If
//!                 End If
//!             End If
//!         Next ctl
//!     End Sub
//! End Class
//!
//! ' Example 4: CSV date column validator
//! Function ValidateCSVDates(filePath As String, dateColumn As Integer) As String
//!     Dim fileNum As Integer
//!     Dim line As String
//!     Dim fields() As String
//!     Dim lineNum As Long
//!     Dim errors As String
//!     
//!     fileNum = FreeFile
//!     Open filePath For Input As fileNum
//!     
//!     lineNum = 0
//!     Do While Not EOF(fileNum)
//!         Line Input #fileNum, line
//!         lineNum = lineNum + 1
//!         
//!         If lineNum = 1 Then
//!             ' Skip header
//!             GoTo NextLine
//!         End If
//!         
//!         fields = Split(line, ",")
//!         
//!         If UBound(fields) >= dateColumn Then
//!             If Not IsDate(fields(dateColumn)) Then
//!                 errors = errors & "Line " & lineNum & ": Invalid date '" & _
//!                         fields(dateColumn) & "'" & vbCrLf
//!             End If
//!         End If
//!         
//! NextLine:
//!     Loop
//!     
//!     Close fileNum
//!     
//!     If errors = "" Then
//!         ValidateCSVDates = "All dates valid"
//!     Else
//!         ValidateCSVDates = "Validation errors:" & vbCrLf & errors
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! The `IsDate` function itself does not raise errors, but it's commonly used to prevent errors:
//!
//! ```vb
//! Function SafeDateOperation(dateStr As String) As Variant
//!     ' Prevent Type Mismatch errors
//!     If Not IsDate(dateStr) Then
//!         MsgBox "Invalid date: " & dateStr, vbCritical
//!         SafeDateOperation = Null
//!         Exit Function
//!     End If
//!     
//!     ' Safe to convert and use
//!     Dim dateValue As Date
//!     dateValue = CDate(dateStr)
//!     
//!     ' Perform date operations
//!     SafeDateOperation = dateValue + 30
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: `IsDate` is a fast check with minimal overhead
//! - **Locale Dependent**: `Date` parsing depends on system locale settings
//! - **Validation Before Conversion**: More efficient than Try/Catch approach with `CDate`
//! - **String Parsing**: `IsDate` must parse string to determine validity
//!
//! ## Best Practices
//!
//! 1. **Always Validate**: Check `IsDate` before `CDate` to prevent Type Mismatch errors
//! 2. **User Input**: Essential for validating user-entered dates
//! 3. **Locale Awareness**: Be aware that date format recognition varies by locale
//! 4. **Clear Messages**: Provide clear error messages when dates are invalid
//! 5. **Range Validation**: Combine `IsDate` with range checks for complete validation
//! 6. **Null Handling**: Remember `IsDate` returns False for `Null` and `Empty`
//! 7. **Database Dates**: Validate before inserting into database date fields
//! 8. **Format Consistency**: Consider standardizing date format after validation
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | `IsDate` | Check if valid date | Boolean | Validate date expressions |
//! | `IsNumeric` | Check if numeric | Boolean | Validate numeric data |
//! | `IsNull` | Check if Null | Boolean | Check for Null values |
//! | `IsEmpty` | Check if uninitialized | Boolean | Check Variant initialization |
//! | `CDate` | Convert to date | Date | Perform conversion |
//! | `VarType` | Get variant type | Integer | Detailed type information |
//! | `DateValue` | Extract date part | Date | Get date without time |
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core functions
//! - Returns `Boolean` type
//! - Date format recognition depends on system locale
//! - Valid date range typically: January 1, 100 to December 31, 9999
//! - Recognizes time-only values as valid dates
//!
//! ## Limitations
//!
//! - Cannot specify expected date format
//! - Locale-dependent interpretation may cause confusion
//! - Does not provide information about why date is invalid
//! - Cannot distinguish between date-only, time-only, or datetime values
//! - May accept ambiguous dates differently on different systems
//! - Does not validate business logic (e.g., "birth date must be in past")
//!
//! ## Related Functions
//!
//! - `CDate`: Convert expression to `Date` type
//! - `DateValue`: Return date part of date/time
//! - `IsNumeric`: Check if numeric
//! - `IsNull`: Check if `Null`
//! - `VarType`: Get detailed type information
//! - `TypeName`: Get type name as `String`
//! - `Format`: Format date for display

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_isdate_basic() {
        let source = r#"
Sub Test()
    result = IsDate(myVariable)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_if_statement() {
        let source = r#"
Sub Test()
    If IsDate(userInput) Then
        dateValue = CDate(userInput)
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_not_condition() {
        let source = r#"
Sub Test()
    If Not IsDate(value) Then
        MsgBox "Invalid date"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_function_return() {
        let source = r#"
Function ValidateDate(v As Variant) As Boolean
    ValidateDate = IsDate(v)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_boolean_and() {
        let source = r#"
Sub Test()
    If IsDate(startDate) And IsDate(endDate) Then
        ProcessDateRange
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_boolean_or() {
        let source = r#"
Sub Test()
    If IsDate(field1) Or IsDate(field2) Then
        ProcessData
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_iif() {
        let source = r#"
Sub Test()
    displayValue = IIf(IsDate(value), Format$(value, "Short Date"), "N/A")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Is valid date: " & IsDate(testValue)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Date validation: " & IsDate(inputValue)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_do_while() {
        let source = r#"
Sub Test()
    Do While Not IsDate(userInput)
        userInput = InputBox("Enter valid date:")
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_do_until() {
        let source = r#"
Sub Test()
    Do Until IsDate(currentValue)
        currentValue = GetNextValue()
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_variable_assignment() {
        let source = r#"
Sub Test()
    Dim isValid As Boolean
    isValid = IsDate(dateString)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_property_assignment() {
        let source = r#"
Sub Test()
    record.IsValidDate = IsDate(record.DateField)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_in_class() {
        let source = r#"
Private Sub Class_Initialize()
    m_isValidDate = IsDate(m_dateValue)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_with_statement() {
        let source = r#"
Sub Test()
    With recordSet
        .IsValid = IsDate(.Fields("OrderDate"))
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_function_argument() {
        let source = r#"
Sub Test()
    Call ProcessIfDate(IsDate(myValue))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_select_case() {
        let source = r#"
Sub Test()
    Select Case True
        Case IsDate(value)
            ConvertDate
        Case Else
            ShowError
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_for_loop() {
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 0 To UBound(dates)
        If IsDate(dates(i)) Then
            validDates = validDates + 1
        End If
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_elseif() {
        let source = r#"
Sub Test()
    If IsNumeric(data) Then
        ProcessNumber
    ElseIf IsDate(data) Then
        ProcessDate
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_concatenation() {
        let source = r#"
Sub Test()
    message = "Valid: " & IsDate(inputText)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_parentheses() {
        let source = r#"
Sub Test()
    result = (IsDate(value))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_array_assignment() {
        let source = r#"
Sub Test()
    validFlags(i) = IsDate(dateValues(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_collection_add() {
        let source = r#"
Sub Test()
    validations.Add IsDate(fields(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_comparison() {
        let source = r#"
Sub Test()
    If IsDate(field1) = IsDate(field2) Then
        MsgBox "Same validity"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_nested_call() {
        let source = r#"
Sub Test()
    result = CStr(IsDate(myVar))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_while_wend() {
        let source = r#"
Sub Test()
    While Not IsDate(input)
        input = GetInput()
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_isdate_error_raise() {
        let source = r#"
Sub Test()
    If Not IsDate(param) Then
        Err.Raise 13, , "Type mismatch"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsDate"));
        assert!(text.contains("Identifier"));
    }
}
