//! # `IsNumeric` Function
//!
//! Returns a `Boolean` value indicating whether an expression can be evaluated as a number.
//!
//! ## Syntax
//!
//! ```vb
//! IsNumeric(expression)
//! ```
//!
//! ## Parameters
//!
//! - `expression` (Required): `Variant` expression to test
//!
//! ## Return Value
//!
//! Returns a `Boolean`:
//! - `True` if the expression can be evaluated as a number
//! - `False` if the expression cannot be evaluated as a number
//! - Returns `True` for numeric strings ("123", "45.67", "-89")
//! - Returns `True` for date/time values (they're stored as numbers)
//! - Returns `True` for `Boolean` values (True = -1, False = 0)
//! - Returns `False` for `Null`
//! - Returns `False` for `Empty`
//! - Returns `False` for non-numeric strings
//! - Recognizes hexadecimal (&H) and octal (&O) notation
//! - Recognizes currency symbols in some locales
//!
//! ## Remarks
//!
//! The `IsNumeric` function is used to determine whether an expression can be converted to a number:
//!
//! - Validates input before numeric conversion
//! - Prevents Type Mismatch errors from `CInt`, `CLng`, `CDbl`, etc.
//! - Recognizes various numeric formats (`Integer`, `Decimal`, and scientific notation)
//! - Locale-dependent for currency and decimal separators
//! - `Date`/`Time` values return `True` (internally stored as `Double`)
//! - `Boolean` values return `True` (`True` = -1, `False` = 0)
//! - Returns `False` for `Null` and `Empty`
//! - Hexadecimal literals (&H10) return `True`
//! - Octal literals (&O77) return `True`
//! - Leading/trailing spaces are ignored
//! - Common in data validation and input processing
//! - Use before converting strings to numbers
//! - Cannot distinguish between integer and floating-point capable strings
//! - `VarType` and `TypeName` provide more detailed type information
//!
//! ## Typical Uses
//!
//! 1. **Input Validation**: Verify user input is numeric before conversion
//! 2. **Data Type Checking**: Determine if `Variant` contains numeric data
//! 3. **Form Validation**: Validate textbox entries contain valid numbers
//! 4. **File Processing**: Validate data from CSV or text files
//! 5. **Error Prevention**: Avoid Type Mismatch errors in calculations
//! 6. **Dynamic Typing**: Handle `Variant` data with unknown types
//! 7. **Database Import**: Validate data before insertion
//! 8. **Report Generation**: Filter numeric values from mixed data
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Basic numeric validation
//! Dim value As Variant
//!
//! value = 123
//! Debug.Print IsNumeric(value)       ' True
//!
//! value = "456"
//! Debug.Print IsNumeric(value)       ' True - numeric string
//!
//! value = "abc"
//! Debug.Print IsNumeric(value)       ' False - non-numeric string
//!
//! ' Example 2: Various data types
//! Debug.Print IsNumeric(123)         ' True - Integer
//! Debug.Print IsNumeric(45.67)       ' True - Double
//! Debug.Print IsNumeric("123")       ' True - numeric string
//! Debug.Print IsNumeric("12.34")     ' True - decimal string
//! Debug.Print IsNumeric("-89")       ' True - negative number
//! Debug.Print IsNumeric("1E+5")      ' True - scientific notation
//! Debug.Print IsNumeric(True)        ' True - Boolean (-1)
//! Debug.Print IsNumeric(False)       ' True - Boolean (0)
//! Debug.Print IsNumeric(#1/1/2025#)  ' True - Date (stored as number)
//! Debug.Print IsNumeric("Hello")     ' False - text
//! Debug.Print IsNumeric(Null)        ' False - Null
//! Debug.Print IsNumeric(Empty)       ' False - Empty
//! Debug.Print IsNumeric("")          ' False - empty string
//! Debug.Print IsNumeric("&H10")      ' True - hexadecimal
//! Debug.Print IsNumeric("&O77")      ' True - octal
//!
//! ' Example 3: Input validation
//! Sub ProcessInput()
//!     Dim userInput As String
//!     Dim numValue As Double
//!     
//!     userInput = InputBox("Enter a number:")
//!     
//!     If IsNumeric(userInput) Then
//!         numValue = CDbl(userInput)
//!         MsgBox "You entered: " & numValue
//!     Else
//!         MsgBox "Invalid number. Please enter a numeric value.", vbExclamation
//!     End If
//! End Sub
//!
//! ' Example 4: Filter numeric values from array
//! Function GetNumericValues(data As Variant) As Variant
//!     Dim result() As Variant
//!     Dim count As Long
//!     Dim i As Long
//!     
//!     If Not IsArray(data) Then
//!         GetNumericValues = Array()
//!         Exit Function
//!     End If
//!     
//!     ReDim result(LBound(data) To UBound(data))
//!     count = LBound(data) - 1
//!     
//!     For i = LBound(data) To UBound(data)
//!         If IsNumeric(data(i)) Then
//!             count = count + 1
//!             result(count) = data(i)
//!         End If
//!     Next i
//!     
//!     If count >= LBound(result) Then
//!         ReDim Preserve result(LBound(result) To count)
//!         GetNumericValues = result
//!     Else
//!         GetNumericValues = Array()
//!     End If
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Safe numeric conversion
//! Function SafeCDbl(value As Variant) As Variant
//!     If IsNumeric(value) Then
//!         SafeCDbl = CDbl(value)
//!     Else
//!         SafeCDbl = Null
//!     End If
//! End Function
//!
//! ' Pattern 2: Numeric with default value
//! Function ToNumber(value As Variant, Optional defaultValue As Double = 0) As Double
//!     If IsNumeric(value) Then
//!         ToNumber = CDbl(value)
//!     Else
//!         ToNumber = defaultValue
//!     End If
//! End Function
//!
//! ' Pattern 3: Validate all values are numeric
//! Function AllNumeric(ParamArray values() As Variant) As Boolean
//!     Dim i As Long
//!     
//!     For i = LBound(values) To UBound(values)
//!         If Not IsNumeric(values(i)) Then
//!             AllNumeric = False
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     AllNumeric = True
//! End Function
//!
//! ' Pattern 4: Count numeric values
//! Function CountNumeric(arr As Variant) As Long
//!     Dim count As Long
//!     Dim i As Long
//!     
//!     If Not IsArray(arr) Then
//!         CountNumeric = 0
//!         Exit Function
//!     End If
//!     
//!     count = 0
//!     For i = LBound(arr) To UBound(arr)
//!         If IsNumeric(arr(i)) Then
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     CountNumeric = count
//! End Function
//!
//! ' Pattern 5: Validate textbox input
//! Function ValidateNumericTextBox(txt As TextBox) As Boolean
//!     If Trim$(txt.Text) = "" Then
//!         MsgBox "Please enter a value", vbExclamation
//!         ValidateNumericTextBox = False
//!     ElseIf Not IsNumeric(txt.Text) Then
//!         MsgBox "Please enter a valid number", vbExclamation
//!         txt.SetFocus
//!         ValidateNumericTextBox = False
//!     Else
//!         ValidateNumericTextBox = True
//!     End If
//! End Function
//!
//! ' Pattern 6: Sum only numeric values
//! Function SumNumeric(values As Variant) As Double
//!     Dim total As Double
//!     Dim i As Long
//!     
//!     If Not IsArray(values) Then
//!         SumNumeric = 0
//!         Exit Function
//!     End If
//!     
//!     total = 0
//!     For i = LBound(values) To UBound(values)
//!         If IsNumeric(values(i)) Then
//!             total = total + CDbl(values(i))
//!         End If
//!     Next i
//!     
//!     SumNumeric = total
//! End Function
//!
//! ' Pattern 7: Parse numeric value with error info
//! Function TryParseNumber(text As String, ByRef result As Double) As Boolean
//!     If IsNumeric(text) Then
//!         result = CDbl(text)
//!         TryParseNumber = True
//!     Else
//!         result = 0
//!         TryParseNumber = False
//!     End If
//! End Function
//!
//! ' Pattern 8: Validate range of values
//! Function ValidateNumericRange(value As Variant, minVal As Double, maxVal As Double) As Boolean
//!     If Not IsNumeric(value) Then
//!         ValidateNumericRange = False
//!         Exit Function
//!     End If
//!     
//!     Dim numVal As Double
//!     numVal = CDbl(value)
//!     
//!     ValidateNumericRange = (numVal >= minVal And numVal <= maxVal)
//! End Function
//!
//! ' Pattern 9: Extract numbers from mixed array
//! Function ExtractNumbers(data As Variant) As Variant
//!     Dim numbers() As Double
//!     Dim count As Long
//!     Dim i As Long
//!     
//!     If Not IsArray(data) Then
//!         ExtractNumbers = Array()
//!         Exit Function
//!     End If
//!     
//!     ReDim numbers(LBound(data) To UBound(data))
//!     count = -1
//!     
//!     For i = LBound(data) To UBound(data)
//!         If IsNumeric(data(i)) Then
//!             count = count + 1
//!             numbers(count) = CDbl(data(i))
//!         End If
//!     Next i
//!     
//!     If count >= 0 Then
//!         ReDim Preserve numbers(0 To count)
//!         ExtractNumbers = numbers
//!     Else
//!         ExtractNumbers = Array()
//!     End If
//! End Function
//!
//! ' Pattern 10: Database field validation
//! Function ValidateNumericField(rs As Recordset, fieldName As String) As Boolean
//!     If IsNull(rs.Fields(fieldName).Value) Then
//!         ValidateNumericField = False
//!     ElseIf IsNumeric(rs.Fields(fieldName).Value) Then
//!         ValidateNumericField = True
//!     Else
//!         ValidateNumericField = False
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Comprehensive form validator
//! Public Class FormValidator
//!     Private m_errors As Collection
//!     
//!     Private Sub Class_Initialize()
//!         Set m_errors = New Collection
//!     End Sub
//!     
//!     Public Function ValidateForm(frm As Form) As Boolean
//!         Dim ctrl As Control
//!         
//!         m_errors.Clear
//!         
//!         For Each ctrl In frm.Controls
//!             If TypeOf ctrl Is TextBox Then
//!                 If ctrl.Tag <> "" Then
//!                     ValidateControl ctrl
//!                 End If
//!             End If
//!         Next ctrl
//!         
//!         ValidateForm = (m_errors.Count = 0)
//!     End Function
//!     
//!     Private Sub ValidateControl(ctrl As Control)
//!         Dim tags As Variant
//!         Dim i As Long
//!         
//!         tags = Split(ctrl.Tag, ",")
//!         
//!         For i = LBound(tags) To UBound(tags)
//!             Select Case Trim$(tags(i))
//!                 Case "required"
//!                     If Trim$(ctrl.Text) = "" Then
//!                         m_errors.Add ctrl.Name & ": Field is required"
//!                     End If
//!                 
//!                 Case "numeric"
//!                     If Trim$(ctrl.Text) <> "" Then
//!                         If Not IsNumeric(ctrl.Text) Then
//!                             m_errors.Add ctrl.Name & ": Must be a number"
//!                         End If
//!                     End If
//!                 
//!                 Case "positive"
//!                     If IsNumeric(ctrl.Text) Then
//!                         If CDbl(ctrl.Text) <= 0 Then
//!                             m_errors.Add ctrl.Name & ": Must be positive"
//!                         End If
//!                     End If
//!                 
//!                 Case "integer"
//!                     If IsNumeric(ctrl.Text) Then
//!                         If CDbl(ctrl.Text) <> CLng(ctrl.Text) Then
//!                             m_errors.Add ctrl.Name & ": Must be an integer"
//!                         End If
//!                     End If
//!             End Select
//!         Next i
//!     End Sub
//!     
//!     Public Function GetErrors() As Collection
//!         Set GetErrors = m_errors
//!     End Function
//! End Class
//!
//! ' Example 2: CSV parser with type detection
//! Public Class CSVParser
//!     Public Function ParseLine(line As String) As Variant
//!         Dim fields() As Variant
//!         Dim parts As Variant
//!         Dim i As Long
//!         
//!         parts = Split(line, ",")
//!         ReDim fields(LBound(parts) To UBound(parts))
//!         
//!         For i = LBound(parts) To UBound(parts)
//!             fields(i) = ConvertField(Trim$(parts(i)))
//!         Next i
//!         
//!         ParseLine = fields
//!     End Function
//!     
//!     Private Function ConvertField(value As String) As Variant
//!         ' Auto-convert to appropriate type
//!         If value = "" Then
//!             ConvertField = Empty
//!         ElseIf UCase$(value) = "NULL" Then
//!             ConvertField = Null
//!         ElseIf IsNumeric(value) Then
//!             ' Determine if integer or floating point
//!             If InStr(value, ".") > 0 Or InStr(value, "E") > 0 Then
//!                 ConvertField = CDbl(value)
//!             Else
//!                 ConvertField = CLng(value)
//!             End If
//!         ElseIf IsDate(value) Then
//!             ConvertField = CDate(value)
//!         Else
//!             ConvertField = value  ' Keep as string
//!         End If
//!     End Function
//! End Class
//!
//! ' Example 3: Data statistics calculator
//! Public Class DataStatistics
//!     Public Function Calculate(data As Variant) As Dictionary
//!         Dim stats As Dictionary
//!         Dim numbers() As Double
//!         Dim count As Long
//!         Dim i As Long
//!         Dim total As Double
//!         Dim mean As Double
//!         Dim variance As Double
//!         
//!         Set stats = CreateObject("Scripting.Dictionary")
//!         
//!         If Not IsArray(data) Then
//!             stats("error") = "Not an array"
//!             Set Calculate = stats
//!             Exit Function
//!         End If
//!         
//!         ' Extract numeric values
//!         ReDim numbers(LBound(data) To UBound(data))
//!         count = -1
//!         
//!         For i = LBound(data) To UBound(data)
//!             If IsNumeric(data(i)) Then
//!                 count = count + 1
//!                 numbers(count) = CDbl(data(i))
//!             End If
//!         Next i
//!         
//!         If count < 0 Then
//!             stats("error") = "No numeric values found"
//!             Set Calculate = stats
//!             Exit Function
//!         End If
//!         
//!         ReDim Preserve numbers(0 To count)
//!         
//!         ' Calculate statistics
//!         total = 0
//!         For i = 0 To count
//!             total = total + numbers(i)
//!         Next i
//!         
//!         mean = total / (count + 1)
//!         
//!         variance = 0
//!         For i = 0 To count
//!             variance = variance + (numbers(i) - mean) ^ 2
//!         Next i
//!         variance = variance / (count + 1)
//!         
//!         stats("count") = count + 1
//!         stats("sum") = total
//!         stats("mean") = mean
//!         stats("variance") = variance
//!         stats("stddev") = Sqr(variance)
//!         stats("min") = MinValue(numbers)
//!         stats("max") = MaxValue(numbers)
//!         
//!         Set Calculate = stats
//!     End Function
//!     
//!     Private Function MinValue(arr() As Double) As Double
//!         Dim min As Double
//!         Dim i As Long
//!         
//!         min = arr(0)
//!         For i = 1 To UBound(arr)
//!             If arr(i) < min Then min = arr(i)
//!         Next i
//!         
//!         MinValue = min
//!     End Function
//!     
//!     Private Function MaxValue(arr() As Double) As Double
//!         Dim max As Double
//!         Dim i As Long
//!         
//!         max = arr(0)
//!         For i = 1 To UBound(arr)
//!             If arr(i) > max Then max = arr(i)
//!         Next i
//!         
//!         MaxValue = max
//!     End Function
//! End Class
//!
//! ' Example 4: Smart input parser
//! Public Class SmartParser
//!     Public Function Parse(input As String) As Variant
//!         Dim trimmed As String
//!         
//!         trimmed = Trim$(input)
//!         
//!         If trimmed = "" Then
//!             Parse = Empty
//!             Exit Function
//!         End If
//!         
//!         ' Try to parse as number
//!         If IsNumeric(trimmed) Then
//!             Parse = ParseAsNumber(trimmed)
//!             Exit Function
//!         End If
//!         
//!         ' Try to parse as date
//!         If IsDate(trimmed) Then
//!             Parse = CDate(trimmed)
//!             Exit Function
//!         End If
//!         
//!         ' Try to parse as boolean
//!         Select Case UCase$(trimmed)
//!             Case "TRUE", "YES", "Y", "1"
//!                 Parse = True
//!                 Exit Function
//!             Case "FALSE", "NO", "N", "0"
//!                 Parse = False
//!                 Exit Function
//!         End Select
//!         
//!         ' Return as string
//!         Parse = trimmed
//!     End Function
//!     
//!     Private Function ParseAsNumber(value As String) As Variant
//!         ' Determine best numeric type
//!         If InStr(value, ".") > 0 Or InStr(value, "E") > 0 Or InStr(value, "e") > 0 Then
//!             ParseAsNumber = CDbl(value)
//!         Else
//!             Dim numVal As Double
//!             numVal = CDbl(value)
//!             
//!             If numVal >= -32768 And numVal <= 32767 And numVal = CLng(numVal) Then
//!                 ParseAsNumber = CInt(numVal)
//!             ElseIf numVal >= -2147483648# And numVal <= 2147483647# And numVal = CLng(numVal) Then
//!                 ParseAsNumber = CLng(numVal)
//!             Else
//!                 ParseAsNumber = numVal
//!             End If
//!         End If
//!     End Function
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! The `IsNumeric` function itself does not raise errors:
//!
//! ```vb
//! ' IsNumeric is safe to call on any value
//! Debug.Print IsNumeric(123)           ' True
//! Debug.Print IsNumeric("456")         ' True
//! Debug.Print IsNumeric("abc")         ' False
//! Debug.Print IsNumeric(Null)          ' False
//! Debug.Print IsNumeric(Empty)         ' False
//!
//! ' Common pattern: validate before conversion
//! Dim text As String
//! text = "123abc"
//!
//! If IsNumeric(text) Then
//!     value = CDbl(text)  ' Safe conversion
//! Else
//!     MsgBox "'" & text & "' is not a valid number", vbExclamation
//! End If
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: `IsNumeric` is a relatively fast check
//! - **Locale-Dependent**: May be slower in some locales due to format parsing
//! - **Cache Results**: If checking the same value multiple times, cache the result
//! - **Prefer Typed Variables**: When possible, use typed variables to avoid checks
//!
//! ## Best Practices
//!
//! 1. **Always Validate Input**: Check `IsNumeric` before converting user input
//! 2. **Handle All Cases**: Account for `Null`, `Empty`, and empty string
//! 3. **Provide Feedback**: Give clear error messages when validation fails
//! 4. **Consider Range**: `IsNumeric` doesn't check if value fits in target type
//! 5. **Locale Awareness**: Be aware of decimal separator differences across locales
//! 6. **Combine Checks**: Often combine with `IsNull`, `IsEmpty` for complete validation
//! 7. **Type-Specific Validation**: Check if integer is needed vs. any numeric
//! 8. **Error Messages**: Provide helpful guidance when validation fails
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | `IsNumeric` | Check if numeric | `Boolean` | Validate numeric data |
//! | `IsDate` | Check if date | `Boolean` | Validate date data |
//! | `IsNull` | Check if `Null` | `Boolean` | Detect `Null` values |
//! | `IsEmpty` | Check if uninitialized | `Boolean` | Detect `Empty` Variants |
//! | `VarType` | Get variant type | `Integer` | Detailed type information |
//! | `TypeName` | Get type name | `String` | Type name as string |
//! | `Val` | Extract number | `Double` | Convert string to number (partial) |
//!
//! ## `IsNumeric` vs `Val` Function
//!
//! ```vb
//! Dim text As String
//!
//! text = "123"
//! Debug.Print IsNumeric(text)    ' True
//! Debug.Print Val(text)          ' 123
//!
//! text = "123abc"
//! Debug.Print IsNumeric(text)    ' False - not fully numeric
//! Debug.Print Val(text)          ' 123 - extracts leading numbers
//!
//! text = "abc123"
//! Debug.Print IsNumeric(text)    ' False
//! Debug.Print Val(text)          ' 0 - no leading numbers
//!
//! ' IsNumeric is stricter - entire expression must be numeric
//! ' Val extracts numeric portion from beginning of string
//! ```
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core functions
//! - Returns `Boolean` type
//! - Locale-dependent for decimal separators and currency
//! - Recognizes hexadecimal (&H) and octal (&O) notation
//! - `Date`/`Time` values return `True`
//! - `Boolean` values return `True`
//!
//! ## Limitations
//!
//! - Does not distinguish between integer and floating-point capable values
//! - Does not check if value fits in target type (`Integer`, `Long`, etc.)
//! - Locale-dependent behavior can cause issues across regions
//! - Cannot validate specific numeric formats (phone numbers, SSN, etc.)
//! - Scientific notation may not be recognized in all locales
//! - Does not validate reasonable ranges for specific use cases
//!
//! ## Related Functions
//!
//! - `IsDate`: Check if expression can be converted to a date
//! - `VarType`: Get detailed Variant type information
//! - `TypeName`: Get type name as string
//! - `Val`: Extract numeric value from string (different behavior)
//! - `CDbl`, `CLng`, `CInt`: Convert to specific numeric types

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn isnumeric_basic() {
        let source = r#"
Sub Test()
    result = IsNumeric(myVariable)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_if_statement() {
        let source = r#"
Sub Test()
    If IsNumeric(value) Then
        numValue = CDbl(value)
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_not_condition() {
        let source = r#"
Sub Test()
    If Not IsNumeric(input) Then
        MsgBox "Please enter a number"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_function_return() {
        let source = r#"
Function IsValidNumber(v As Variant) As Boolean
    IsValidNumber = IsNumeric(v)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_boolean_and() {
        let source = r#"
Sub Test()
    If IsNumeric(value1) And IsNumeric(value2) Then
        total = CDbl(value1) + CDbl(value2)
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_boolean_or() {
        let source = r#"
Sub Test()
    If IsNumeric(field) Or IsDate(field) Then
        ProcessValue field
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_iif() {
        let source = r#"
Sub Test()
    displayValue = IIf(IsNumeric(value), CDbl(value), 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Is numeric: " & IsNumeric(testVar)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Numeric status: " & IsNumeric(userInput)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_do_while() {
        let source = r#"
Sub Test()
    Do While Not IsNumeric(input)
        input = InputBox("Enter a number:")
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_do_until() {
        let source = r#"
Sub Test()
    Do Until IsNumeric(result)
        result = GetNextValue()
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_variable_assignment() {
        let source = r#"
Sub Test()
    Dim isValid As Boolean
    isValid = IsNumeric(dataValue)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_property_assignment() {
        let source = r#"
Sub Test()
    obj.IsNumericValue = IsNumeric(obj.Data)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_in_class() {
        let source = r#"
Private Sub Class_Initialize()
    m_isNumeric = IsNumeric(m_value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_with_statement() {
        let source = r#"
Sub Test()
    With validation
        .IsValid = IsNumeric(.Value)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_function_argument() {
        let source = r#"
Sub Test()
    Call ValidateInput(IsNumeric(txtAmount.Text))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_select_case() {
        let source = r#"
Sub Test()
    Select Case True
        Case IsNumeric(value)
            ProcessNumber value
        Case Else
            ProcessText value
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_for_loop() {
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 0 To UBound(arr)
        If IsNumeric(arr(i)) Then
            total = total + CDbl(arr(i))
        End If
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_elseif() {
        let source = r#"
Sub Test()
    If IsDate(data) Then
        ProcessDate data
    ElseIf IsNumeric(data) Then
        ProcessNumber data
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_concatenation() {
        let source = r#"
Sub Test()
    status = "Valid: " & IsNumeric(variable)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_parentheses() {
        let source = r#"
Sub Test()
    result = (IsNumeric(value))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_array_filter() {
        let source = r#"
Sub Test()
    checks(i) = IsNumeric(values(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_collection_add() {
        let source = r#"
Sub Test()
    numericFlags.Add IsNumeric(data(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_comparison() {
        let source = r#"
Sub Test()
    If IsNumeric(var1) = IsNumeric(var2) Then
        MsgBox "Same type"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_nested_call() {
        let source = r#"
Sub Test()
    result = CStr(IsNumeric(myVar))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_while_wend() {
        let source = r#"
Sub Test()
    While Not IsNumeric(buffer)
        buffer = GetValidInput()
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn isnumeric_input_validation() {
        let source = r#"
Function GetNumber() As Double
    Dim input As String
    input = InputBox("Enter number:")
    
    If IsNumeric(input) Then
        GetNumber = CDbl(input)
    Else
        GetNumber = 0
    End If
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsNumeric"));
        assert!(text.contains("Identifier"));
    }
}
