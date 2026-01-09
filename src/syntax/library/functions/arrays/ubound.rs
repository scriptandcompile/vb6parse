//! VB6 `UBound` Function
//!
//! The `UBound` function returns a Long containing the largest available subscript for the indicated dimension of an array.
//!
//! ## Syntax
//! ```vb6
//! UBound(arrayname[, dimension])
//! ```
//!
//! ## Parameters
//! - `arrayname`: Required. Name of the array variable. Follows standard Visual Basic naming conventions.
//! - `dimension`: Optional. Variant (Long). Specifies which dimension's upper bound is returned. Use 1 for the first dimension, 2 for the second, and so on. If `dimension` is omitted, 1 is assumed.
//!
//! ## Returns
//! Returns a `Long` containing the largest available subscript for the specified dimension of the array.
//!
//! ## Remarks
//! The `UBound` function is used to determine the upper limit of an array dimension:
//!
//! - **Dimension parameter**: If omitted, defaults to 1 (first dimension)
//! - **Multi-dimensional arrays**: Use `dimension` parameter to specify which dimension
//! - **Zero-based arrays**: `UBound` returns the upper index regardless of lower bound
//! - **Paired with `LBound`**: Use `LBound` to get the lower bound
//! - **Array size calculation**: Size = `UBound - LBound + 1`
//! - **Dynamic arrays**: Returns current upper bound (changes with `ReDim`)
//! - **Fixed arrays**: Returns the declared upper bound
//! - **Error on uninitialized**: Error 9 (Subscript out of range) if array not initialized
//! - **`ParamArray`**: Works with `ParamArray` arguments to find number of elements
//!
//! ### Common Array Declarations
//! ```vb6
//! Dim arr(5)              ' LBound = 0, UBound = 5 (6 elements)
//! Dim arr(1 To 5)         ' LBound = 1, UBound = 5 (5 elements)
//! Dim arr(10 To 20)       ' LBound = 10, UBound = 20 (11 elements)
//! Dim arr(5, 3)           ' First: 0-5, Second: 0-3
//! Dim arr(1 To 5, 1 To 3) ' First: 1-5, Second: 1-3
//! ```
//!
//! ### Option Base Impact
//! The `Option Base` statement affects default lower bounds:
//! - `Option Base 0`: Default lower bound is 0 (default)
//! - `Option Base 1`: Default lower bound is 1
//! - Explicit bounds (e.g., `1 To 5`) override Option Base
//!
//! ### Dynamic Arrays
//! For dynamic arrays:
//! - Before `ReDim`: Error 9 if accessed
//! - After `ReDim`: Returns current upper bound
//! - `ReDim Preserve`: Can change upper bound while preserving data
//! - `Erase`: Makes array uninitialized again
//!
//! ## Typical Uses
//! 1. **Loop Bounds**: Iterate through all array elements
//! 2. **Array Size**: Calculate the number of elements in an array
//! 3. **Validation**: Check if an index is within valid range
//! 4. **Dynamic Resizing**: Determine current size before `ReDim`
//! 5. **`ParamArray`**: Count variable number of arguments
//! 6. **Array Copying**: Determine target array size
//! 7. **Search Operations**: Set loop limits for array searches
//! 8. **Multi-dimensional**: Navigate complex array structures
//!
//! ## Basic Examples
//!
//! ### Example 1: Simple Array Iteration
//! ```vb6
//! Dim values(10) As Integer
//! Dim i As Integer
//!
//! For i = LBound(values) To UBound(values)
//!     values(i) = i * 2
//! Next i
//! ```
//!
//! ### Example 2: Calculate Array Size
//! ```vb6
//! Function GetArraySize(arr() As Variant) As Long
//!     GetArraySize = UBound(arr) - LBound(arr) + 1
//! End Function
//!
//! ' Usage:
//! Dim myArray(5 To 15) As String
//! Debug.Print GetArraySize(myArray) ' Prints: 11
//! ```
//!
//! ### Example 3: Multi-Dimensional Array
//! ```vb6
//! Sub ProcessMatrix()
//!     Dim matrix(1 To 3, 1 To 4) As Double
//!     Dim row As Integer
//!     Dim col As Integer
//!     
//!     For row = LBound(matrix, 1) To UBound(matrix, 1)
//!         For col = LBound(matrix, 2) To UBound(matrix, 2)
//!             matrix(row, col) = row * col
//!         Next col
//!     Next row
//! End Sub
//! ```
//!
//! ### Example 4: `ParamArray` with `UBound`
//! ```vb6
//! Function Sum(ParamArray values() As Variant) As Double
//!     Dim i As Integer
//!     Dim total As Double
//!     
//!     total = 0
//!     For i = LBound(values) To UBound(values)
//!         total = total + values(i)
//!     Next i
//!     
//!     Sum = total
//! End Function
//!
//! ' Usage: result = Sum(1, 2, 3, 4, 5)
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Safe Array Iteration
//! ```vb6
//! Sub IterateArray(arr() As Variant)
//!     Dim i As Long
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         Debug.Print arr(i)
//!     Next i
//! End Sub
//! ```
//!
//! ### Pattern 2: Check If Array Is Empty
//! ```vb6
//! Function IsArrayEmpty(arr() As Variant) As Boolean
//!     On Error Resume Next
//!     IsArrayEmpty = (UBound(arr) < LBound(arr))
//!     If Err.Number <> 0 Then IsArrayEmpty = True
//! End Function
//! ```
//!
//! ### Pattern 3: Resize Array with Data Preservation
//! ```vb6
//! Sub AddArrayElement(arr() As Variant, newValue As Variant)
//!     Dim newSize As Long
//!     
//!     On Error Resume Next
//!     newSize = UBound(arr) + 1
//!     If Err.Number <> 0 Then
//!         ' Array not initialized
//!         ReDim arr(0 To 0)
//!         newSize = 0
//!     Else
//!         ReDim Preserve arr(LBound(arr) To newSize)
//!     End If
//!     
//!     arr(newSize) = newValue
//! End Sub
//! ```
//!
//! ### Pattern 4: Count Elements in `ParamArray`
//! ```vb6
//! Function CountArgs(ParamArray args() As Variant) As Long
//!     On Error Resume Next
//!     CountArgs = UBound(args) - LBound(args) + 1
//!     If Err.Number <> 0 Then CountArgs = 0
//! End Function
//! ```
//!
//! ### Pattern 5: Validate Array Index
//! ```vb6
//! Function IsValidIndex(arr() As Variant, index As Long) As Boolean
//!     On Error Resume Next
//!     IsValidIndex = (index >= LBound(arr) And index <= UBound(arr))
//!     If Err.Number <> 0 Then IsValidIndex = False
//! End Function
//! ```
//!
//! ### Pattern 6: Copy Array
//! ```vb6
//! Function CopyArray(source() As Variant) As Variant()
//!     Dim dest() As Variant
//!     Dim i As Long
//!     
//!     ReDim dest(LBound(source) To UBound(source))
//!     
//!     For i = LBound(source) To UBound(source)
//!         dest(i) = source(i)
//!     Next i
//!     
//!     CopyArray = dest
//! End Function
//! ```
//!
//! ### Pattern 7: Reverse Array
//! ```vb6
//! Sub ReverseArray(arr() As Variant)
//!     Dim i As Long
//!     Dim j As Long
//!     Dim temp As Variant
//!     
//!     i = LBound(arr)
//!     j = UBound(arr)
//!     
//!     While i < j
//!         temp = arr(i)
//!         arr(i) = arr(j)
//!         arr(j) = temp
//!         i = i + 1
//!         j = j - 1
//!     Wend
//! End Sub
//! ```
//!
//! ### Pattern 8: Find Last Element
//! ```vb6
//! Function GetLastElement(arr() As Variant) As Variant
//!     GetLastElement = arr(UBound(arr))
//! End Function
//! ```
//!
//! ### Pattern 9: Remove Last Element
//! ```vb6
//! Sub RemoveLastElement(arr() As Variant)
//!     Dim newUpper As Long
//!     
//!     newUpper = UBound(arr) - 1
//!     If newUpper >= LBound(arr) Then
//!         ReDim Preserve arr(LBound(arr) To newUpper)
//!     End If
//! End Sub
//! ```
//!
//! ### Pattern 10: Multi-Dimensional Size
//! ```vb6
//! Function GetArrayDimensions(arr As Variant) As Integer
//!     Dim dimension As Integer
//!     
//!     On Error Resume Next
//!     dimension = 1
//!     Do While Err.Number = 0
//!         Dim test As Long
//!         test = UBound(arr, dimension)
//!         dimension = dimension + 1
//!     Loop
//!     
//!     GetArrayDimensions = dimension - 1
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Dynamic Array Manager Class
//! ```vb6
//! ' Class: DynamicArrayManager
//! ' Manages a dynamic array with automatic resizing
//! Option Explicit
//!
//! Private m_Data() As Variant
//! Private m_Initialized As Boolean
//!
//! Public Sub Initialize(Optional initialSize As Long = 10)
//!     ReDim m_Data(0 To initialSize - 1)
//!     m_Initialized = True
//! End Sub
//!
//! Public Sub Add(value As Variant)
//!     Dim newIndex As Long
//!     
//!     If Not m_Initialized Then
//!         Initialize
//!         newIndex = 0
//!     Else
//!         newIndex = UBound(m_Data) + 1
//!         ReDim Preserve m_Data(0 To newIndex)
//!     End If
//!     
//!     m_Data(newIndex) = value
//! End Sub
//!
//! Public Function GetItem(index As Long) As Variant
//!     If index < LBound(m_Data) Or index > UBound(m_Data) Then
//!         Err.Raise 9, , "Index out of range"
//!     End If
//!     
//!     If IsObject(m_Data(index)) Then
//!         Set GetItem = m_Data(index)
//!     Else
//!         GetItem = m_Data(index)
//!     End If
//! End Function
//!
//! Public Sub SetItem(index As Long, value As Variant)
//!     If index < LBound(m_Data) Or index > UBound(m_Data) Then
//!         Err.Raise 9, , "Index out of range"
//!     End If
//!     
//!     m_Data(index) = value
//! End Sub
//!
//! Public Function Count() As Long
//!     If Not m_Initialized Then
//!         Count = 0
//!     Else
//!         Count = UBound(m_Data) - LBound(m_Data) + 1
//!     End If
//! End Function
//!
//! Public Sub Clear()
//!     If m_Initialized Then
//!         Erase m_Data
//!         m_Initialized = False
//!     End If
//! End Sub
//!
//! Public Function ToArray() As Variant()
//!     ToArray = m_Data
//! End Function
//! ```
//!
//! ### Example 2: Array Utilities Module
//! ```vb6
//! ' Module: ArrayUtilities
//! ' Comprehensive array manipulation utilities
//! Option Explicit
//!
//! Public Function ArraySize(arr As Variant) As Long
//!     On Error Resume Next
//!     ArraySize = UBound(arr) - LBound(arr) + 1
//!     If Err.Number <> 0 Then ArraySize = 0
//! End Function
//!
//! Public Function ArrayContains(arr() As Variant, value As Variant) As Boolean
//!     Dim i As Long
//!     
//!     ArrayContains = False
//!     For i = LBound(arr) To UBound(arr)
//!         If arr(i) = value Then
//!             ArrayContains = True
//!             Exit Function
//!         End If
//!     Next i
//! End Function
//!
//! Public Function ArrayIndexOf(arr() As Variant, value As Variant) As Long
//!     Dim i As Long
//!     
//!     ArrayIndexOf = -1
//!     For i = LBound(arr) To UBound(arr)
//!         If arr(i) = value Then
//!             ArrayIndexOf = i
//!             Exit Function
//!         End If
//!     Next i
//! End Function
//!
//! Public Sub ArraySort(arr() As Variant)
//!     Dim i As Long
//!     Dim j As Long
//!     Dim temp As Variant
//!     
//!     For i = LBound(arr) To UBound(arr) - 1
//!         For j = i + 1 To UBound(arr)
//!             If arr(i) > arr(j) Then
//!                 temp = arr(i)
//!                 arr(i) = arr(j)
//!                 arr(j) = temp
//!             End If
//!         Next j
//!     Next i
//! End Sub
//!
//! Public Function ArrayFilter(arr() As Variant, filterValue As Variant) As Variant()
//!     Dim result() As Variant
//!     Dim i As Long
//!     Dim count As Long
//!     
//!     count = 0
//!     For i = LBound(arr) To UBound(arr)
//!         If arr(i) <> filterValue Then
//!             ReDim Preserve result(0 To count)
//!             result(count) = arr(i)
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     ArrayFilter = result
//! End Function
//!
//! Public Function ArraySlice(arr() As Variant, startIndex As Long, _
//!                           endIndex As Long) As Variant()
//!     Dim result() As Variant
//!     Dim i As Long
//!     Dim idx As Long
//!     
//!     ReDim result(0 To endIndex - startIndex)
//!     
//!     idx = 0
//!     For i = startIndex To endIndex
//!         result(idx) = arr(i)
//!         idx = idx + 1
//!     Next i
//!     
//!     ArraySlice = result
//! End Function
//! ```
//!
//! ### Example 3: Matrix Operations Class
//! ```vb6
//! ' Class: MatrixOperations
//! ' Performs operations on 2D arrays
//! Option Explicit
//!
//! Public Function GetRowCount(matrix As Variant) As Long
//!     On Error Resume Next
//!     GetRowCount = UBound(matrix, 1) - LBound(matrix, 1) + 1
//!     If Err.Number <> 0 Then GetRowCount = 0
//! End Function
//!
//! Public Function GetColumnCount(matrix As Variant) As Long
//!     On Error Resume Next
//!     GetColumnCount = UBound(matrix, 2) - LBound(matrix, 2) + 1
//!     If Err.Number <> 0 Then GetColumnCount = 0
//! End Function
//!
//! Public Function GetRow(matrix As Variant, rowIndex As Long) As Variant()
//!     Dim result() As Variant
//!     Dim col As Long
//!     Dim idx As Long
//!     
//!     ReDim result(LBound(matrix, 2) To UBound(matrix, 2))
//!     
//!     For col = LBound(matrix, 2) To UBound(matrix, 2)
//!         result(col) = matrix(rowIndex, col)
//!     Next col
//!     
//!     GetRow = result
//! End Function
//!
//! Public Function GetColumn(matrix As Variant, colIndex As Long) As Variant()
//!     Dim result() As Variant
//!     Dim row As Long
//!     
//!     ReDim result(LBound(matrix, 1) To UBound(matrix, 1))
//!     
//!     For row = LBound(matrix, 1) To UBound(matrix, 1)
//!         result(row) = matrix(row, colIndex)
//!     Next row
//!     
//!     GetColumn = result
//! End Function
//!
//! Public Function TransposeMatrix(matrix As Variant) As Variant
//!     Dim result() As Variant
//!     Dim row As Long
//!     Dim col As Long
//!     
//!     ReDim result(LBound(matrix, 2) To UBound(matrix, 2), _
//!                  LBound(matrix, 1) To UBound(matrix, 1))
//!     
//!     For row = LBound(matrix, 1) To UBound(matrix, 1)
//!         For col = LBound(matrix, 2) To UBound(matrix, 2)
//!             result(col, row) = matrix(row, col)
//!         Next col
//!     Next row
//!     
//!     TransposeMatrix = result
//! End Function
//! ```
//!
//! ### Example 4: Collection to Array Converter
//! ```vb6
//! ' Module: CollectionConverter
//! ' Converts between Collections and Arrays
//! Option Explicit
//!
//! Public Function CollectionToArray(col As Collection) As Variant()
//!     Dim result() As Variant
//!     Dim i As Long
//!     
//!     If col.Count = 0 Then
//!         CollectionToArray = Array()
//!         Exit Function
//!     End If
//!     
//!     ReDim result(1 To col.Count)
//!     
//!     For i = 1 To col.Count
//!         If IsObject(col(i)) Then
//!             Set result(i) = col(i)
//!         Else
//!             result(i) = col(i)
//!         End If
//!     Next i
//!     
//!     CollectionToArray = result
//! End Function
//!
//! Public Function ArrayToCollection(arr() As Variant) As Collection
//!     Dim result As New Collection
//!     Dim i As Long
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         result.Add arr(i)
//!     Next i
//!     
//!     Set ArrayToCollection = result
//! End Function
//!
//! Public Function MergeArrays(ParamArray arrays() As Variant) As Variant()
//!     Dim result() As Variant
//!     Dim totalSize As Long
//!     Dim currentIndex As Long
//!     Dim i As Long
//!     Dim j As Long
//!     Dim arr As Variant
//!     
//!     ' Calculate total size
//!     totalSize = 0
//!     For i = LBound(arrays) To UBound(arrays)
//!         arr = arrays(i)
//!         totalSize = totalSize + (UBound(arr) - LBound(arr) + 1)
//!     Next i
//!     
//!     ' Merge arrays
//!     ReDim result(0 To totalSize - 1)
//!     currentIndex = 0
//!     
//!     For i = LBound(arrays) To UBound(arrays)
//!         arr = arrays(i)
//!         For j = LBound(arr) To UBound(arr)
//!             result(currentIndex) = arr(j)
//!             currentIndex = currentIndex + 1
//!         Next j
//!     Next i
//!     
//!     MergeArrays = result
//! End Function
//! ```
//!
//! ## Error Handling
//! The `UBound` function can raise the following errors:
//!
//! - **Error 9 (Subscript out of range)**: If the array has not been initialized (for dynamic arrays)
//! - **Error 9 (Subscript out of range)**: If `dimension` is less than 1 or greater than the array's number of dimensions
//! - **Error 13 (Type mismatch)**: If the variable is not an array
//! - **Error 5 (Invalid procedure call or argument)**: If dimension parameter is invalid
//!
//! ## Performance Notes
//! - Very fast O(1) operation - directly returns array metadata
//! - No performance difference between dimensions
//! - Safe to call repeatedly in loops
//! - Consider caching value if used extensively in tight loops
//! - No memory allocation or copying involved
//!
//! ## Best Practices
//! 1. **Always use with `LBound`** for complete array bounds information
//! 2. **Check for initialization** with On Error Resume Next for dynamic arrays
//! 3. **Use in For loops** instead of hardcoding array sizes
//! 4. **Specify dimension** explicitly for multi-dimensional arrays
//! 5. **Cache in variables** if used multiple times in tight loops
//! 6. **Validate dimension parameter** when working with multi-dimensional arrays
//! 7. **Handle errors gracefully** for potentially uninitialized arrays
//! 8. **Use for `ParamArray`** to handle variable arguments
//! 9. **Document array bounds** in function comments
//! 10. **Prefer explicit bounds** in array declarations for clarity
//!
//! ## Comparison Table
//!
//! | Function | Purpose | Returns | Notes |
//! |----------|---------|---------|-------|
//! | `UBound` | Upper bound | Long | Largest valid index |
//! | `LBound` | Lower bound | Long | Smallest valid index |
//! | `Array` | Create array | Variant | Returns zero-based array |
//! | `ReDim` | Resize array | N/A | Statement, not function |
//!
//! ## Platform Notes
//! - Available in VB6, VBA, and `VBScript`
//! - Behavior consistent across platforms
//! - Returns Long (32-bit signed integer)
//! - Maximum array size limited by available memory
//! - Multi-dimensional arrays limited to 60 dimensions
//!
//! ## Limitations
//! - Cannot determine if array is initialized without error handling
//! - Does not return array capacity (allocated size vs. used size)
//! - No built-in way to get all dimensions at once
//! - Dimension parameter must be compile-time constant in some contexts
//! - Cannot be used on Collections or other non-array types
//! - Does not work with jagged arrays (arrays of arrays) directly

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn ubound_basic() {
        let source = r"
Sub Test()
    upper = UBound(myArray)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_variable_assignment() {
        let source = r"
Sub Test()
    Dim maxIndex As Long
    maxIndex = UBound(values)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_with_dimension() {
        let source = r"
Sub Test()
    rows = UBound(matrix, 1)
    cols = UBound(matrix, 2)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_for_loop() {
        let source = r"
Sub Test()
    For i = LBound(arr) To UBound(arr)
        Process arr(i)
    Next i
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_function_return() {
        let source = r"
Function GetArraySize(arr() As Variant) As Long
    GetArraySize = UBound(arr) - LBound(arr) + 1
End Function
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_if_statement() {
        let source = r"
Sub Test()
    If index > UBound(data) Then
        Err.Raise 9
    End If
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Array size: " & (UBound(arr) - LBound(arr) + 1)
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn ubound_select_case() {
        let source = r#"
Sub Test()
    Select Case UBound(items)
        Case 0 To 10
            category = "Small"
        Case Else
            category = "Large"
    End Select
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_redim() {
        let source = r"
Sub Test()
    ReDim Preserve arr(LBound(arr) To UBound(arr) + 1)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_function_argument() {
        let source = r"
Sub Test()
    Call ProcessArray(data, UBound(data))
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_comparison() {
        let source = r"
Sub Test()
    If UBound(arr1) > UBound(arr2) Then
        larger = arr1
    End If
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Upper bound: " & UBound(values)
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_do_while() {
        let source = r"
Sub Test()
    Do While i <= UBound(items)
        Process items(i)
        i = i + 1
    Loop
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_do_until() {
        let source = r"
Sub Test()
    Do Until i > UBound(data)
        i = i + 1
    Loop
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_while_wend() {
        let source = r"
Sub Test()
    While idx <= UBound(arr)
        idx = idx + 1
    Wend
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_iif() {
        let source = r#"
Sub Test()
    size = IIf(UBound(arr) > 10, "Large", "Small")
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_with_statement() {
        let source = r"
Sub Test()
    With arrayManager
        .MaxIndex = UBound(.Data)
    End With
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_parentheses() {
        let source = r"
Sub Test()
    result = (UBound(arr) + 1) * 2
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn ubound_error_handling() {
        let source = r"
Sub Test()
    On Error Resume Next
    maxIdx = UBound(dynamicArray)
    If Err.Number <> 0 Then
        maxIdx = -1
    End If
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_property_assignment() {
        let source = r"
Sub Test()
    obj.UpperBound = UBound(obj.Items)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_concatenation() {
        let source = r#"
Sub Test()
    message = "Array has " & UBound(arr) + 1 & " elements"
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_arithmetic() {
        let source = r"
Sub Test()
    lastIndex = UBound(values) - 1
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_print_statement() {
        let source = r"
Sub Test()
    Print #1, UBound(data)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_class_usage() {
        let source = r"
Sub Test()
    Set manager = New ArrayManager
    manager.Size = UBound(manager.Data) + 1
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn ubound_array_bounds() {
        let source = r"
Sub Test()
    lastElement = arr(UBound(arr))
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn ubound_elseif() {
        let source = r#"
Sub Test()
    If UBound(arr) = 0 Then
        result = "Empty"
    ElseIf UBound(arr) < 10 Then
        result = "Small"
    End If
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn ubound_nested_arrays() {
        let source = r"
Sub Test()
    For i = LBound(matrix, 1) To UBound(matrix, 1)
        For j = LBound(matrix, 2) To UBound(matrix, 2)
            total = total + matrix(i, j)
        Next j
    Next i
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn ubound_paramarray() {
        let source = r"
Function Sum(ParamArray values() As Variant) As Double
    For i = LBound(values) To UBound(values)
        total = total + values(i)
    Next i
    Sum = total
End Function
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/arrays/ubound",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
