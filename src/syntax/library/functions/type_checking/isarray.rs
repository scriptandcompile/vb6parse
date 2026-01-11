//! # `IsArray` Function
//!
//! Returns a `Boolean` value indicating whether a variable is an array.
//!
//! ## Syntax
//!
//! ```vb
//! IsArray(varname)
//! ```
//!
//! ## Parameters
//!
//! `varname` (Required): Variable name to test
//!
//! ## Return Value
//!
//! Returns a `Boolean`:
//! - `True` if the variable is an array
//! - `False` if the variable is not an array
//! - `True` for both fixed-size and dynamic arrays
//! - `True` even for unallocated dynamic arrays (dimensioned but not `ReDim`-med)
//!
//! ## Remarks
//!
//! The `IsArray` function is used to determine whether a variable is an array:
//! - Returns `True` for any array variable, regardless of dimensions
//! - Returns `True` for dynamic arrays even before they're allocated with `ReDim`
//! - Returns `False` for all non-array variables
//! - Useful when working with `Variant` variables that might contain arrays
//! - Often used with `ParamArray` parameters to validate input
//! - Can be used with the `Array` function result
//! - Returns `True` for arrays passed as `Variant` parameters
//! - Returns `True` for arrays stored in `Variant` variables
//! - Commonly used in procedures that accept flexible data types
//! - Important for validating function arguments
//!
//! ## Typical Uses
//!
//! **Parameter Validation**: Verify that a `Variant` parameter contains an array
//! **Data Type Detection**: Determine if a `Variant` holds array data
//! **`ParamArray` Handling**: Check individual elements of `ParamArray`
//! **Dynamic Programming**: Handle different data types in generic routines
//! **Array Processing**: Validate data before array operations
//! **Error Prevention**: Avoid runtime errors by checking array status
//! **Flexible Functions**: Create functions that accept both single values and arrays
//! **Type Checking**: Part of comprehensive type validation routines
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Simple array check
//! Dim myArray(1 To 10) As Integer
//! Dim myValue As Integer
//!
//! If IsArray(myArray) Then
//!     Debug.Print "myArray is an array"  ' This prints
//! End If
//!
//! If IsArray(myValue) Then
//!     Debug.Print "myValue is an array"
//! Else
//!     Debug.Print "myValue is not an array"  ' This prints
//! End If
//! ```
//!
//! ```vb
//! ' Example 2: Checking Variant contents
//! Dim myVariant As Variant
//!
//! myVariant = Array(1, 2, 3, 4, 5)
//! If IsArray(myVariant) Then
//!     Debug.Print "Variant contains an array"  ' This prints
//! End If
//!
//! myVariant = 42
//! If IsArray(myVariant) Then
//!     Debug.Print "Variant contains an array"
//! Else
//!     Debug.Print "Variant does not contain an array"  ' This prints
//! End If
//! ```
//!
//! ```vb
//! ' Example 3: Dynamic array before ReDim
//! Dim dynamicArray() As String
//!
//! If IsArray(dynamicArray) Then
//!     Debug.Print "Dynamic array variable is an array"  ' This prints even before ReDim
//! End If
//! ```
//!
//! ```vb
//! ' Example 4: Validating function parameters
//! Function ProcessData(data As Variant) As Long
//!     If IsArray(data) Then
//!         ProcessData = UBound(data) - LBound(data) + 1
//!         Debug.Print "Processing array with " & ProcessData & " elements"
//!     Else
//!         ProcessData = 1
//!         Debug.Print "Processing single value"
//!     End If
//! End Function
//!
//! ' Usage
//! Dim result As Long
//! result = ProcessData(Array(1, 2, 3, 4, 5))  ' Prints: Processing array with 5 elements
//! result = ProcessData(100)                    ' Prints: Processing single value
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Validate array parameter
//! Function SumValues(values As Variant) As Double
//!     Dim i As Long
//!     Dim total As Double
//!     
//!     If Not IsArray(values) Then
//!         Err.Raise 5, , "Parameter must be an array"
//!     End If
//!     
//!     For i = LBound(values) To UBound(values)
//!         total = total + values(i)
//!     Next i
//!     
//!     SumValues = total
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 2: Handle single value or array
//! Sub DisplayData(data As Variant)
//!     Dim i As Long
//!     
//!     If IsArray(data) Then
//!         For i = LBound(data) To UBound(data)
//!             Debug.Print "Item " & i & ": " & data(i)
//!         Next i
//!     Else
//!         Debug.Print "Single value: " & data
//!     End If
//! End Sub
//! ```
//!
//! ```vb
//! ' Pattern 3: Convert single value to array if needed
//! Function EnsureArray(value As Variant) As Variant
//!     If IsArray(value) Then
//!         EnsureArray = value
//!     Else
//!         EnsureArray = Array(value)
//!     End If
//! End Function
//!
//! ```vb
//! ' Pattern 4: Count array elements safely
//! Function GetElementCount(data As Variant) As Long
//!     If IsArray(data) Then
//!         GetElementCount = UBound(data) - LBound(data) + 1
//!     Else
//!         GetElementCount = 1
//!     End If
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 5: Validate before array operation
//! Function GetFirstElement(arr As Variant) As Variant
//!     If Not IsArray(arr) Then
//!         GetFirstElement = arr
//!         Exit Function
//!     End If
//!     
//!     If UBound(arr) >= LBound(arr) Then
//!         GetFirstElement = arr(LBound(arr))
//!     Else
//!         GetFirstElement = Null
//!     End If
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 6: Type checking with IsArray
//! Function DescribeVariable(v As Variant) As String
//!     If IsArray(v) Then
//!         DescribeVariable = "Array with " & (UBound(v) - LBound(v) + 1) & " elements"
//!     ElseIf IsNumeric(v) Then
//!         DescribeVariable = "Numeric value: " & v
//!     ElseIf IsDate(v) Then
//!         DescribeVariable = "Date value: " & v
//!     ElseIf IsNull(v) Then
//!         DescribeVariable = "Null value"
//!     ElseIf IsEmpty(v) Then
//!         DescribeVariable = "Empty variant"
//!     Else
//!         DescribeVariable = "String or object: " & v
//!     End If
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 7: ParamArray validation
//! Sub ProcessItems(ParamArray items() As Variant)
//!     Dim i As Long
//!     
//!     For i = LBound(items) To UBound(items)
//!         If IsArray(items(i)) Then
//!             Debug.Print "Item " & i & " is an array"
//!         Else
//!             Debug.Print "Item " & i & ": " & items(i)
//!         End If
//!     Next i
//! End Sub
//! ```
//!
//! ```vb
//! ' Pattern 8: Safely iterate over data
//! Sub SafeIterate(data As Variant)
//!     Dim i As Long
//!     
//!     If IsArray(data) Then
//!         For i = LBound(data) To UBound(data)
//!             ProcessValue data(i)
//!         Next i
//!     Else
//!         ProcessValue data
//!     End If
//! End Sub
//! ```
//!
//! ```vb
//! ' Pattern 9: Flatten nested arrays
//! Function FlattenArray(arr As Variant) As Variant
//!     Dim result() As Variant
//!     Dim i As Long
//!     Dim count As Long
//!     
//!     If Not IsArray(arr) Then
//!         ReDim result(0 To 0)
//!         result(0) = arr
//!         FlattenArray = result
//!         Exit Function
//!     End If
//!     
//!     count = 0
//!     ReDim result(0 To 100)  ' Initial size
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         If IsArray(arr(i)) Then
//!             ' Handle nested array (recursive)
//!             Dim flattened As Variant
//!             flattened = FlattenArray(arr(i))
//!             ' Add flattened elements to result...
//!         Else
//!             result(count) = arr(i)
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     ReDim Preserve result(0 To count - 1)
//!     FlattenArray = result
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 10: Conditional array processing
//! Function ApplyOperation(data As Variant, operation As String) As Variant
//!     Dim i As Long
//!     Dim result As Variant
//!     
//!     If IsArray(data) Then
//!         ReDim result(LBound(data) To UBound(data))
//!         For i = LBound(data) To UBound(data)
//!             result(i) = PerformOperation(data(i), operation)
//!         Next i
//!         ApplyOperation = result
//!     Else
//!         ApplyOperation = PerformOperation(data, operation)
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Flexible data processor class
//! Public Class DataProcessor
//!     Private m_data As Variant
//!     
//!     Public Sub SetData(data As Variant)
//!         m_data = data
//!     End Sub
//!     
//!     Public Function GetCount() As Long
//!         If IsArray(m_data) Then
//!             GetCount = UBound(m_data) - LBound(m_data) + 1
//!         ElseIf IsEmpty(m_data) Then
//!             GetCount = 0
//!         Else
//!             GetCount = 1
//!         End If
//!     End Function
//!     
//!     Public Function GetSum() As Double
//!         Dim i As Long
//!         Dim total As Double
//!         
//!         If IsArray(m_data) Then
//!             For i = LBound(m_data) To UBound(m_data)
//!                 If IsNumeric(m_data(i)) Then
//!                     total = total + m_data(i)
//!                 End If
//!             Next i
//!         ElseIf IsNumeric(m_data) Then
//!             total = m_data
//!         End If
//!         
//!         GetSum = total
//!     End Function
//!     
//!     Public Function GetAverage() As Double
//!         Dim count As Long
//!         count = GetCount()
//!         
//!         If count > 0 Then
//!             GetAverage = GetSum() / count
//!         End If
//!     End Function
//!     
//!     Public Sub DisplayData()
//!         Dim i As Long
//!         
//!         If IsArray(m_data) Then
//!             Debug.Print "Array with " & GetCount() & " elements:"
//!             For i = LBound(m_data) To UBound(m_data)
//!                 Debug.Print "  [" & i & "] = " & m_data(i)
//!             Next i
//!         ElseIf IsEmpty(m_data) Then
//!             Debug.Print "No data"
//!         Else
//!             Debug.Print "Single value: " & m_data
//!         End If
//!     End Sub
//! End Class
//! ```
//!
//! ```vb
//! ' Example 2: Generic collection converter
//! Public Class CollectionConverter
//!     Public Function ToArray(source As Variant) As Variant
//!         Dim result() As Variant
//!         Dim i As Long
//!         
//!         If IsArray(source) Then
//!             ' Already an array, just return it
//!             ToArray = source
//!         ElseIf TypeName(source) = "Collection" Then
//!             ' Convert collection to array
//!             ReDim result(1 To source.Count)
//!             For i = 1 To source.Count
//!                 result(i) = source(i)
//!             Next i
//!             ToArray = result
//!         Else
//!             ' Single value, wrap in array
//!             ReDim result(0 To 0)
//!             result(0) = source
//!             ToArray = result
//!         End If
//!     End Function
//!     
//!     Public Function ToCollection(source As Variant) As Collection
//!         Dim result As Collection
//!         Dim i As Long
//!         
//!         Set result = New Collection
//!         
//!         If IsArray(source) Then
//!             For i = LBound(source) To UBound(source)
//!                 result.Add source(i)
//!             Next i
//!         Else
//!             result.Add source
//!         End If
//!         
//!         Set ToCollection = result
//!     End Function
//! End Class
//! ```
//!
//! ```vb
//! ' Example 3: Safe array utilities module
//! Public Module ArrayUtils
//!     Public Function SafeUBound(arr As Variant, Optional dimension As Integer = 1) As Long
//!         On Error Resume Next
//!         If IsArray(arr) Then
//!             SafeUBound = UBound(arr, dimension)
//!             If Err.Number <> 0 Then SafeUBound = -1
//!         Else
//!             SafeUBound = -1
//!         End If
//!         On Error GoTo 0
//!     End Function
//!     
//!     Public Function SafeLBound(arr As Variant, Optional dimension As Integer = 1) As Long
//!         On Error Resume Next
//!         If IsArray(arr) Then
//!             SafeLBound = LBound(arr, dimension)
//!             If Err.Number <> 0 Then SafeLBound = 0
//!         Else
//!             SafeLBound = 0
//!         End If
//!         On Error GoTo 0
//!     End Function
//!     
//!     Public Function IsAllocatedArray(arr As Variant) As Boolean
//!         On Error Resume Next
//!         If IsArray(arr) Then
//!             Dim ub As Long
//!             ub = UBound(arr)
//!             IsAllocatedArray = (Err.Number = 0)
//!         Else
//!             IsAllocatedArray = False
//!         End If
//!         On Error GoTo 0
//!     End Function
//!     
//!     Public Function CombineArrays(arr1 As Variant, arr2 As Variant) As Variant
//!         Dim result() As Variant
//!         Dim i As Long, count As Long
//!         
//!         count = 0
//!         
//!         ' Count total elements
//!         If IsArray(arr1) Then
//!             count = count + (UBound(arr1) - LBound(arr1) + 1)
//!         Else
//!             count = count + 1
//!         End If
//!         
//!         If IsArray(arr2) Then
//!             count = count + (UBound(arr2) - LBound(arr2) + 1)
//!         Else
//!             count = count + 1
//!         End If
//!         
//!         ReDim result(0 To count - 1)
//!         
//!         ' Copy elements
//!         count = 0
//!         If IsArray(arr1) Then
//!             For i = LBound(arr1) To UBound(arr1)
//!                 result(count) = arr1(i)
//!                 count = count + 1
//!             Next i
//!         Else
//!             result(count) = arr1
//!             count = count + 1
//!         End If
//!         
//!         If IsArray(arr2) Then
//!             For i = LBound(arr2) To UBound(arr2)
//!                 result(count) = arr2(i)
//!                 count = count + 1
//!             Next i
//!         Else
//!             result(count) = arr2
//!             count = count + 1
//!         End If
//!         
//!         CombineArrays = result
//!     End Function
//! End Module
//! ```
//!
//! ```vb
//! ' Example 4: Flexible function that handles multiple input types
//! Function CalculateTotal(values As Variant, Optional taxRate As Double = 0) As Double
//!     Dim i As Long
//!     Dim subtotal As Double
//!     Dim tax As Double
//!     
//!     subtotal = 0
//!     
//!     If IsArray(values) Then
//!         ' Process array of values
//!         For i = LBound(values) To UBound(values)
//!             If IsNumeric(values(i)) Then
//!                 subtotal = subtotal + values(i)
//!             ElseIf IsArray(values(i)) Then
//!                 ' Handle nested array (recursive call)
//!                 subtotal = subtotal + CalculateTotal(values(i), 0)
//!             End If
//!         Next i
//!     ElseIf IsNumeric(values) Then
//!         ' Process single value
//!         subtotal = values
//!     Else
//!         Err.Raise 13, , "Type mismatch: values must be numeric or array"
//!     End If
//!     
//!     ' Apply tax
//!     If taxRate > 0 Then
//!         tax = subtotal * taxRate
//!         CalculateTotal = subtotal + tax
//!     Else
//!         CalculateTotal = subtotal
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! The `IsArray` function itself does not raise errors, but it's often used in error prevention:
//!
//! ```vb
//! Function SafeArrayOperation(arr As Variant) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     If Not IsArray(arr) Then
//!         Err.Raise 5, , "Invalid procedure call: array expected"
//!     End If
//!     
//!     ' Proceed with array operations
//!     Dim i As Long
//!     For i = LBound(arr) To UBound(arr)
//!         ' Process array elements
//!     Next i
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     MsgBox "Error: " & Err.Description, vbCritical
//!     SafeArrayOperation = Null
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: `IsArray` is a very fast check with minimal overhead
//! - **Type Checking**: More efficient than attempting array operations and handling errors
//! - **Avoid Redundant Checks**: Cache `IsArray` result if checking multiple times
//! - **Early Validation**: Check `IsArray` early to avoid unnecessary processing
//!
//! ## Best Practices
//!
//! 1. **Validate Parameters**: Use `IsArray` to validate `Variant` parameters before array operations
//! 2. **Flexible Functions**: Create functions that gracefully handle both arrays and single values
//! 3. **Clear Error Messages**: Provide informative errors when array is expected but not received
//! 4. **Combine Checks**: Use with other Is functions (`IsNumeric`, `IsNull`, etc.) for complete validation
//! 5. **Document Expectations**: Clearly document whether functions expect arrays or single values
//! 6. **Handle Edge Cases**: Consider unallocated dynamic arrays (`IsArray` returns `True` but `UBound` fails)
//! 7. **Use Early Returns**: Check `IsArray` early and return/exit if validation fails
//! 8. **`ParamArray` Elements**: Remember each `ParamArray` element might be an array
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | `IsArray` | Check if array | `Boolean` | Validate array variables |
//! | `IsEmpty` | Check if uninitialized | `Boolean` | Check Variant initialization |
//! | `IsNull` | Check if Null | `Boolean` | Check for Null values |
//! | `IsNumeric` | Check if numeric | `Boolean` | Validate numeric data |
//! | `IsObject` | Check if object | `Boolean` | Validate object references |
//! | `VarType` | Get variant type | `Integer` | Detailed type information |
//! | `TypeName` | Get type name | `String` | Type name as string |
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core functions
//! - Returns `Boolean` type
//! - Works with all array types (fixed, dynamic, single or multi-dimensional)
//! - Returns `True` for unallocated dynamic arrays
//!
//! ## Limitations
//!
//! - Does not indicate whether dynamic array is allocated (dimensioned vs `ReDim`-med)
//! - Cannot determine number of dimensions
//! - Cannot determine array bounds
//! - Does not validate array contents
//! - Does not distinguish between different array types
//!
//! ## Related Functions
//!
//! - `UBound`: Upper bound of array dimension
//! - `LBound`: Lower bound of array dimension
//! - `Array`: Create `Variant` array
//! - `VarType`: Get detailed type information
//! - `TypeName`: Get type name as string
//! - `IsEmpty`: Check if `Variant` is uninitialized
//! - `IsNull`: Check if `Variant` is `Null`

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn isarray_basic() {
        let source = r"
Sub Test()
    result = IsArray(myVariable)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_if_statement() {
        let source = r#"
Sub Test()
    If IsArray(data) Then
        MsgBox "Data is an array"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_not_condition() {
        let source = r"
Sub Test()
    If Not IsArray(value) Then
        Exit Sub
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_function_return() {
        let source = r"
Function CheckArray(v As Variant) As Boolean
    CheckArray = IsArray(v)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_boolean_and() {
        let source = r"
Sub Test()
    If IsArray(data) And UBound(data) > 0 Then
        ProcessArray data
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_boolean_or() {
        let source = r"
Sub Test()
    If IsArray(data) Or IsNull(data) Then
        handleSpecialCase
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_select_case() {
        let source = r"
Sub Test()
    Select Case True
        Case IsArray(value)
            HandleArray
        Case Else
            HandleSingle
    End Select
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_iif() {
        let source = r"
Sub Test()
    count = IIf(IsArray(data), UBound(data) + 1, 1)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Is array: " & IsArray(myVar)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Array check: " & IsArray(data)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_do_while() {
        let source = r"
Sub Test()
    Do While IsArray(currentData)
        currentData = GetNextData()
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_do_until() {
        let source = r"
Sub Test()
    Do Until IsArray(result)
        result = ProcessNext()
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_variable_assignment() {
        let source = r"
Sub Test()
    Dim isArr As Boolean
    isArr = IsArray(myData)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_property_assignment() {
        let source = r"
Sub Test()
    obj.IsArrayData = IsArray(obj.Data)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_in_class() {
        let source = r"
Private Sub Class_Initialize()
    m_isArray = IsArray(m_data)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_with_statement() {
        let source = r"
Sub Test()
    With dataObject
        .IsValid = IsArray(.Values)
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_function_argument() {
        let source = r"
Sub Test()
    Call ValidateData(IsArray(myVariable))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_comparison() {
        let source = r#"
Sub Test()
    If IsArray(data1) = IsArray(data2) Then
        MsgBox "Same type"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_for_loop() {
        let source = r"
Sub Test()
    Dim i As Integer
    For i = 0 To 10
        If IsArray(items(i)) Then
            ProcessArray items(i)
        End If
    Next i
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_elseif() {
        let source = r"
Sub Test()
    If IsNumeric(data) Then
        ProcessNumber data
    ElseIf IsArray(data) Then
        ProcessArray data
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_concatenation() {
        let source = r#"
Sub Test()
    message = "Type check: " & IsArray(value)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_parentheses() {
        let source = r"
Sub Test()
    result = (IsArray(data))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_array_index() {
        let source = r"
Sub Test()
    checks(i) = IsArray(values(i))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_collection_add() {
        let source = r"
Sub Test()
    results.Add IsArray(data(i))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_nested_call() {
        let source = r"
Sub Test()
    result = CStr(IsArray(myVar))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_error_check() {
        let source = r"
Sub Test()
    If Not IsArray(param) Then
        Err.Raise 5
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isarray_while_wend() {
        let source = r"
Sub Test()
    While IsArray(current)
        current = GetNext()
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isarray",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
