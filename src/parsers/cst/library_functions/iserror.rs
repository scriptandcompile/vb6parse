//! # `IsError` Function
//!
//! Returns a `Boolean` value indicating whether an expression is an error value.
//!
//! ## Syntax
//!
//! ```vb
//! IsError(expression)
//! ```
//!
//! ## Parameters
//!
//! - `expression` (Required): Variant expression to test
//!
//! ## Return Value
//!
//! Returns a `Boolean`:
//! - `True` if the expression is an error value created by `CVErr`
//! - `False` if the expression is not an error value
//! - Only detects error values created with `CVErr` function
//! - Does not detect runtime errors or error objects
//! - Works with `Variant` variables containing error values
//! - Returns `False` for `Null`, `Empty`, or any non-error value
//!
//! ## Remarks
//!
//! The `IsError` function is used to determine whether a `Variant` expression contains an error value:
//!
//! - Only detects `CVErr` error values (`Variant` subtype `vbError`)
//! - Does not detect `Err` object or runtime errors
//! - Error values are created using `CVErr` function
//! - Useful for propagating errors through `Variant` returns
//! - Common in functions that need to return error indicators
//! - Error values are different from Null or Empty
//! - Can be used to check function return values for errors
//! - Error values preserve error numbers through call chains
//! - Use `CVErr` to create error values, `IsError` to detect them
//! - `VarType(expr) = vbError` provides same functionality
//! - Error values are uncommon in modern VB6 code
//! - Most code uses `Err.Raise` for error handling instead
//!
//! ## Typical Uses
//!
//! 1. **Error Propagation**: Check if function returned an error value
//! 2. **Error Value Detection**: Identify `CVErr` values in `Variant` data
//! 3. **Function Return Checking**: Validate function results
//! 4. **Array Processing**: Detect errors in array elements
//! 5. **Data Validation**: Distinguish errors from valid data
//! 6. **Legacy Code**: Work with older code using `CVErr` pattern
//! 7. **Error Chains**: Propagate errors through multiple function calls
//! 8. **Conditional Logic**: Branch based on error presence
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Create and detect error values
//! Dim result As Variant
//!
//! result = CVErr(5)  ' Create error value with error number 5
//!
//! If IsError(result) Then
//!     Debug.Print "Result is an error"  ' This prints
//!     Debug.Print "Error number: " & CLng(result)  ' Prints: 5
//! End If
//!
//! ' Example 2: Distinguish error from other values
//! Dim testVar As Variant
//!
//! testVar = CVErr(13)
//! Debug.Print IsError(testVar)        ' True - error value
//! testVar = 13
//! Debug.Print IsError(testVar)        ' False - regular number
//! testVar = Null
//! Debug.Print IsError(testVar)        ' False - Null is not error
//! testVar = Empty
//! Debug.Print IsError(testVar)        ' False - Empty is not error
//!
//! ' Example 3: Function returning error or value
//! Function SafeDivide(numerator As Double, denominator As Double) As Variant
//!     If denominator = 0 Then
//!         SafeDivide = CVErr(11)  ' Division by zero error
//!     Else
//!         SafeDivide = numerator / denominator
//!     End If
//! End Function
//!
//! ' Usage
//! Dim result As Variant
//! result = SafeDivide(10, 2)
//!
//! If IsError(result) Then
//!     MsgBox "Error in calculation: " & CLng(result)
//! Else
//!     MsgBox "Result: " & result  ' Prints: 5
//! End If
//!
//! ' Example 4: Process array with error checking
//! Function ProcessValues(values() As Variant) As Variant
//!     Dim i As Integer
//!     Dim total As Double
//!     
//!     total = 0
//!     For i = LBound(values) To UBound(values)
//!         If IsError(values(i)) Then
//!             ProcessValues = CVErr(CLng(values(i)))  ' Propagate error
//!             Exit Function
//!         End If
//!         total = total + values(i)
//!     Next i
//!     
//!     ProcessValues = total
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Safe function call with error checking
//! Function SafeGetValue(source As Variant, key As String) As Variant
//!     On Error Resume Next
//!     SafeGetValue = source(key)
//!     
//!     If Err.Number <> 0 Then
//!         SafeGetValue = CVErr(Err.Number)
//!     End If
//!     On Error GoTo 0
//! End Function
//!
//! ' Usage
//! Dim value As Variant
//! value = SafeGetValue(myDict, "key")
//! If IsError(value) Then
//!     MsgBox "Error: " & CLng(value)
//! End If
//!
//! ' Pattern 2: Coalesce - return first non-error value
//! Function CoalesceValues(ParamArray values() As Variant) As Variant
//!     Dim i As Long
//!     
//!     For i = LBound(values) To UBound(values)
//!         If Not IsError(values(i)) And Not IsNull(values(i)) And Not IsEmpty(values(i)) Then
//!             CoalesceValues = values(i)
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     CoalesceValues = CVErr(xlErrNA)  ' All values were invalid
//! End Function
//!
//! ' Pattern 3: Get error number from error value
//! Function GetErrorNumber(errorValue As Variant) As Long
//!     If IsError(errorValue) Then
//!         GetErrorNumber = CLng(errorValue)
//!     Else
//!         GetErrorNumber = 0  ' No error
//!     End If
//! End Function
//!
//! ' Pattern 4: Chain operations with error propagation
//! Function CalculateResult(a As Variant, b As Variant) As Variant
//!     If IsError(a) Then
//!         CalculateResult = a  ' Propagate first error
//!         Exit Function
//!     End If
//!     
//!     If IsError(b) Then
//!         CalculateResult = b  ' Propagate second error
//!         Exit Function
//!     End If
//!     
//!     ' Perform calculation
//!     CalculateResult = a + b
//! End Function
//!
//! ' Pattern 5: Validate all values before processing
//! Function AllValid(ParamArray values() As Variant) As Boolean
//!     Dim i As Long
//!     
//!     For i = LBound(values) To UBound(values)
//!         If IsError(values(i)) Or IsNull(values(i)) Then
//!             AllValid = False
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     AllValid = True
//! End Function
//!
//! ' Pattern 6: Convert error to message
//! Function ErrorToMessage(value As Variant) As String
//!     If IsError(value) Then
//!         Select Case CLng(value)
//!             Case 5
//!                 ErrorToMessage = "Invalid procedure call"
//!             Case 7
//!                 ErrorToMessage = "Out of memory"
//!             Case 9
//!                 ErrorToMessage = "Subscript out of range"
//!             Case 11
//!                 ErrorToMessage = "Division by zero"
//!             Case 13
//!                 ErrorToMessage = "Type mismatch"
//!             Case Else
//!                 ErrorToMessage = "Error " & CLng(value)
//!         End Select
//!     Else
//!         ErrorToMessage = "No error"
//!     End If
//! End Function
//!
//! ' Pattern 7: Default value for errors
//! Function ValueOrDefault(value As Variant, defaultValue As Variant) As Variant
//!     If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
//!         ValueOrDefault = defaultValue
//!     Else
//!         ValueOrDefault = value
//!     End If
//! End Function
//!
//! ' Pattern 8: Find first error in array
//! Function FindFirstError(arr As Variant) As Variant
//!     Dim i As Long
//!     
//!     If Not IsArray(arr) Then
//!         FindFirstError = CVErr(13)  ' Type mismatch
//!         Exit Function
//!     End If
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         If IsError(arr(i)) Then
//!             FindFirstError = arr(i)
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     FindFirstError = Null  ' No errors found
//! End Function
//!
//! ' Pattern 9: Count errors in array
//! Function CountErrors(arr As Variant) As Long
//!     Dim i As Long
//!     Dim count As Long
//!     
//!     If Not IsArray(arr) Then
//!         CountErrors = 0
//!         Exit Function
//!     End If
//!     
//!     count = 0
//!     For i = LBound(arr) To UBound(arr)
//!         If IsError(arr(i)) Then
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     CountErrors = count
//! End Function
//!
//! ' Pattern 10: Safe numeric conversion
//! Function SafeCDbl(value As Variant) As Variant
//!     On Error Resume Next
//!     Dim result As Double
//!     
//!     result = CDbl(value)
//!     
//!     If Err.Number <> 0 Then
//!         SafeCDbl = CVErr(Err.Number)
//!     Else
//!         SafeCDbl = result
//!     End If
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Error-aware calculator class
//! Public Class SafeCalculator
//!     Public Function Add(a As Variant, b As Variant) As Variant
//!         If IsError(a) Then
//!             Add = a
//!             Exit Function
//!         End If
//!         If IsError(b) Then
//!             Add = b
//!             Exit Function
//!         End If
//!         
//!         On Error Resume Next
//!         Add = a + b
//!         If Err.Number <> 0 Then
//!             Add = CVErr(Err.Number)
//!         End If
//!         On Error GoTo 0
//!     End Function
//!     
//!     Public Function Divide(numerator As Variant, denominator As Variant) As Variant
//!         If IsError(numerator) Then
//!             Divide = numerator
//!             Exit Function
//!         End If
//!         If IsError(denominator) Then
//!             Divide = denominator
//!             Exit Function
//!         End If
//!         
//!         If denominator = 0 Then
//!             Divide = CVErr(11)  ' Division by zero
//!             Exit Function
//!         End If
//!         
//!         Divide = numerator / denominator
//!     End Function
//!     
//!     Public Function GetErrorMessage(errorValue As Variant) As String
//!         If Not IsError(errorValue) Then
//!             GetErrorMessage = "No error"
//!             Exit Function
//!         End If
//!         
//!         GetErrorMessage = "Error " & CLng(errorValue) & ": " & _
//!                          Error$(CLng(errorValue))
//!     End Function
//! End Class
//!
//! ' Example 2: Variant array processor with error handling
//! Public Class VariantArrayProcessor
//!     Public Function Map(arr As Variant, callback As String) As Variant
//!         ' Apply callback function to each element, propagate errors
//!         Dim result() As Variant
//!         Dim i As Long
//!         
//!         If Not IsArray(arr) Then
//!             Map = CVErr(13)  ' Type mismatch
//!             Exit Function
//!         End If
//!         
//!         ReDim result(LBound(arr) To UBound(arr))
//!         
//!         For i = LBound(arr) To UBound(arr)
//!             If IsError(arr(i)) Then
//!                 result(i) = arr(i)  ' Preserve error
//!             Else
//!                 On Error Resume Next
//!                 result(i) = Application.Run(callback, arr(i))
//!                 If Err.Number <> 0 Then
//!                     result(i) = CVErr(Err.Number)
//!                 End If
//!                 On Error GoTo 0
//!             End If
//!         Next i
//!         
//!         Map = result
//!     End Function
//!     
//!     Public Function Filter(arr As Variant) As Variant
//!         ' Remove error values from array
//!         Dim result() As Variant
//!         Dim i As Long, count As Long
//!         
//!         If Not IsArray(arr) Then
//!             Filter = Array()
//!             Exit Function
//!         End If
//!         
//!         ReDim result(LBound(arr) To UBound(arr))
//!         count = LBound(arr) - 1
//!         
//!         For i = LBound(arr) To UBound(arr)
//!             If Not IsError(arr(i)) Then
//!                 count = count + 1
//!                 result(count) = arr(i)
//!             End If
//!         Next i
//!         
//!         If count >= LBound(result) Then
//!             ReDim Preserve result(LBound(result) To count)
//!             Filter = result
//!         Else
//!             Filter = Array()
//!         End If
//!     End Function
//! End Class
//!
//! ' Example 3: Data validator with detailed error reporting
//! Public Class DataValidator
//!     Private m_errors As Collection
//!     
//!     Private Sub Class_Initialize()
//!         Set m_errors = New Collection
//!     End Sub
//!     
//!     Public Function ValidateRecord(record As Variant) As Boolean
//!         Dim i As Long
//!         Dim fieldName As String
//!         
//!         m_errors.Clear
//!         
//!         If Not IsArray(record) Then
//!             ValidateRecord = False
//!             Exit Function
//!         End If
//!         
//!         For i = LBound(record) To UBound(record)
//!             If IsError(record(i)) Then
//!                 m_errors.Add "Field " & i & ": Error " & CLng(record(i))
//!             ElseIf IsNull(record(i)) Then
//!                 m_errors.Add "Field " & i & ": Null value"
//!             ElseIf IsEmpty(record(i)) Then
//!                 m_errors.Add "Field " & i & ": Empty value"
//!             End If
//!         Next i
//!         
//!         ValidateRecord = (m_errors.Count = 0)
//!     End Function
//!     
//!     Public Function GetErrors() As Collection
//!         Set GetErrors = m_errors
//!     End Function
//!     
//!     Public Function GetErrorSummary() As String
//!         Dim msg As String
//!         Dim i As Long
//!         
//!         If m_errors.Count = 0 Then
//!             GetErrorSummary = "No errors"
//!             Exit Function
//!         End If
//!         
//!         msg = "Found " & m_errors.Count & " error(s):" & vbCrLf
//!         For i = 1 To m_errors.Count
//!             msg = msg & "- " & m_errors(i) & vbCrLf
//!         Next i
//!         
//!         GetErrorSummary = msg
//!     End Function
//! End Class
//!
//! ' Example 4: Function composition with error handling
//! Function Compose(value As Variant, ParamArray functions() As Variant) As Variant
//!     ' Apply functions in sequence, stop on first error
//!     Dim i As Long
//!     Dim result As Variant
//!     
//!     result = value
//!     
//!     For i = LBound(functions) To UBound(functions)
//!         If IsError(result) Then
//!             Compose = result  ' Propagate error
//!             Exit Function
//!         End If
//!         
//!         On Error Resume Next
//!         result = Application.Run(functions(i), result)
//!         If Err.Number <> 0 Then
//!             Compose = CVErr(Err.Number)
//!             Exit Function
//!         End If
//!         On Error GoTo 0
//!     Next i
//!     
//!     Compose = result
//! End Function
//!
//! ' Usage:
//! ' result = Compose(10, "DoubleValue", "AddTen", "FormatResult")
//! ' If IsError(result) Then MsgBox "Error in processing"
//! ```
//!
//! ## Error Handling
//!
//! The `IsError` function itself does not raise errors:
//!
//! ```vb
//! ' IsError is safe to call on any value
//! Debug.Print IsError(123)           ' False
//! Debug.Print IsError("text")        ' False
//! Debug.Print IsError(CVErr(5))      ' True
//! Debug.Print IsError(Null)          ' False
//! Debug.Print IsError(Empty)         ' False
//!
//! ' Common pattern: check and extract error number
//! If IsError(value) Then
//!     Dim errNum As Long
//!     errNum = CLng(value)  ' Extract error number
//!     MsgBox Error$(errNum) ' Get error description
//! End If
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: `IsError` is a very fast type check
//! - **Overhead**: `CVErr`/`IsError` pattern has more overhead than `Err.Raise`
//! - **Modern Alternative**: Most code uses structured error handling instead
//! - **Legacy Code**: Primarily seen in older VB6 and Excel VBA code
//!
//! ## Best Practices
//!
//! 1. **Prefer `Err.Raise`**: Use structured error handling for most scenarios
//! 2. **Check Returns**: Always check `IsError` for functions returning `Variant`
//! 3. **Propagate Errors**: Pass error values through call chains when appropriate
//! 4. **Document Behavior**: Clearly document when functions return error values
//! 5. **Extract Numbers**: Use `CLng(errorValue)` to get error number from error value
//! 6. **Combine Checks**: Check `IsError`, `IsNull`, and `IsEmpty` for complete validation
//! 7. **Error Messages**: Convert error numbers to messages for user display
//! 8. **Avoid Overuse**: `CVErr` pattern less common in modern VB6 code
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | `IsError` | Check if `CVErr` value | `Boolean` | Detect error values |
//! | `CVErr` | Create error value | `Variant` (Error) | Return error indicator |
//! | `IsNull` | Check if Null | `Boolean` | Detect Null values |
//! | `IsEmpty` | Check if uninitialized | `Boolean` | Detect Empty Variants |
//! | `VarType` | Get variant type | `Integer` | Detailed type information |
//! | `Err.Raise` | Raise runtime error | N/A | Structured error handling |
//! | `Error$` | Get error description | `String` | Error message from number |
//!
//! ## `CVErr` vs `Err.Raise`
//!
//! ```vb
//! ' CVErr pattern (older style)
//! Function OldStyleDivide(a As Double, b As Double) As Variant
//!     If b = 0 Then
//!         OldStyleDivide = CVErr(11)  ' Return error value
//!     Else
//!         OldStyleDivide = a / b
//!     End If
//! End Function
//!
//! If IsError(result) Then
//!     MsgBox "Error: " & CLng(result)
//! End If
//!
//! ' Err.Raise pattern (modern style)
//! Function ModernDivide(a As Double, b As Double) As Double
//!     If b = 0 Then
//!         Err.Raise 11, , "Division by zero"  ' Raise error
//!     Else
//!         ModernDivide = a / b
//!     End If
//! End Function
//!
//! On Error Resume Next
//! result = ModernDivide(10, 0)
//! If Err.Number <> 0 Then
//!     MsgBox "Error: " & Err.Description
//! End If
//! ```
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core functions
//! - Returns `Boolean` type
//! - Only detects `CVErr` error values (Variant subtype vbError)
//! - Does not detect `Err` object or runtime errors
//! - More common in Excel VBA than desktop VB6
//! - Excel has predefined errors: `xlErrDiv0`, `xlErrNA`, `xlErrName`, `xlErrNull`, `xlErrNum`, `xlErrRef`, `xlErrValue`
//!
//! ## Limitations
//!
//! - Only detects `CVErr` error values, not runtime errors
//! - Does not provide error description (use `Error$` function)
//! - Cannot distinguish different error types beyond number
//! - Less flexible than structured error handling (Try/Catch equivalent)
//! - Error values can be confusing when mixed with normal values
//! - Not widely used in modern VB6 applications
//! - Requires `Variant` return types (cannot use with typed returns)
//!
//! ## Related Functions
//!
//! - `CVErr`: Create error value from error number
//! - `Error$`: Get error description from error number
//! - `IsNull`: Check if `Variant` is `Null`
//! - `IsEmpty`: Check if `Variant` is `Empty`
//! - `VarType`: Get detailed `Variant` type information
//! - `Err.Raise`: Raise runtime error (preferred modern approach)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_iserror_basic() {
        let source = r#"
Sub Test()
    result = IsError(myVariable)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_if_statement() {
        let source = r#"
Sub Test()
    If IsError(value) Then
        HandleError value
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_not_condition() {
        let source = r#"
Sub Test()
    If Not IsError(result) Then
        ProcessResult result
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_function_return() {
        let source = r#"
Function IsValid(v As Variant) As Boolean
    IsValid = Not IsError(v)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_boolean_and() {
        let source = r#"
Sub Test()
    If IsError(value1) And IsError(value2) Then
        ReportErrors
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_boolean_or() {
        let source = r#"
Sub Test()
    If IsError(field) Or IsNull(field) Then
        ShowWarning
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_iif() {
        let source = r#"
Sub Test()
    displayValue = IIf(IsError(value), "Error", value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Is error: " & IsError(testVar)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Error status: " & IsError(myVar)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_do_while() {
        let source = r#"
Sub Test()
    Do While IsError(currentValue)
        currentValue = GetNextValue()
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_do_until() {
        let source = r#"
Sub Test()
    Do Until Not IsError(result)
        result = TryOperation()
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_variable_assignment() {
        let source = r#"
Sub Test()
    Dim hasError As Boolean
    hasError = IsError(dataValue)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_property_assignment() {
        let source = r#"
Sub Test()
    obj.HasError = IsError(obj.Value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_in_class() {
        let source = r#"
Private Sub Class_Initialize()
    m_hasError = IsError(m_data)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_with_statement() {
        let source = r#"
Sub Test()
    With record
        .IsValid = Not IsError(.Data)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_function_argument() {
        let source = r#"
Sub Test()
    Call LogError(IsError(calculation))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_select_case() {
        let source = r#"
Sub Test()
    Select Case True
        Case IsError(value)
            HandleError
        Case Else
            ProcessValue
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_for_loop() {
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 0 To UBound(arr)
        If IsError(arr(i)) Then
            errorCount = errorCount + 1
        End If
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_elseif() {
        let source = r#"
Sub Test()
    If IsNull(data) Then
        ProcessNull
    ElseIf IsError(data) Then
        ProcessError
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_concatenation() {
        let source = r#"
Sub Test()
    status = "Error: " & IsError(variable)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_parentheses() {
        let source = r#"
Sub Test()
    result = (IsError(value))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_array_check() {
        let source = r#"
Sub Test()
    checks(i) = IsError(values(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_collection_add() {
        let source = r#"
Sub Test()
    errorStates.Add IsError(data(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_comparison() {
        let source = r#"
Sub Test()
    If IsError(var1) = IsError(var2) Then
        MsgBox "Same state"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_nested_call() {
        let source = r#"
Sub Test()
    result = CStr(IsError(myVar))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_while_wend() {
        let source = r#"
Sub Test()
    While IsError(buffer)
        buffer = GetValidData()
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_iserror_error_propagation() {
        let source = r#"
Function Calculate(a As Variant) As Variant
    If IsError(a) Then
        Calculate = a
    Else
        Calculate = a * 2
    End If
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IsError"));
        assert!(text.contains("Identifier"));
    }
}
