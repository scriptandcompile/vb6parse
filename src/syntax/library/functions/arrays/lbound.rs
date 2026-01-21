//! ## `LBound` Function
//!
//! Returns a `Long` containing the smallest available subscript for the indicated dimension of an array.
//!
//! ## Syntax
//!
//! ```text
//! LBound(arrayname, [dimension])
//! ```
//!
//! ## Parameters
//!
//! - **arrayname** (Required): Name of the array variable
//! - *dimension** (Optional): `Integer` specifying which dimension's lower bound to return
//!   - If omitted, defaults to 1 (first dimension)
//!   - Must be between 1 and the number of dimensions in the array
//!
//! ## Return Value
//!
//! - Returns a `Long` of the smallest available subscript for the specified dimension.
//! - By default, arrays start at 0 unless `Option Base 1` is specified.
//! - Returns 0 for standard arrays (`Option Base 0`).
//! - Returns 1 for arrays when `Option Base 1` is specified.
//! - For arrays declared with explicit bounds, returns the specified lower bound.
//! - Dynamic arrays preserve their lower bound across `ReDim` operations.
//!
//! ## Remarks
//!
//! The `LBound` function is essential for array processing:
//!
//! - Returns the lower bound (minimum index) of an array dimension
//! - Counterpart to `UBound` (which returns upper bound)
//! - Critical for correctly iterating through arrays
//! - Default lower bound is 0 (unless `Option Base 1`)
//! - Can specify explicit lower bounds:
//!
//! ```vb6
//! Dim arr(5 To 10) 'LBound = 5
//! ```
//!
//! - Works with multi-dimensional arrays using dimension parameter
//! - Omitting dimension parameter returns bound of first dimension
//! - Dynamic arrays must be dimensioned before calling `LBound`
//! - Fixed-size arrays always have bounds available
//! - `ParamArray` parameters always have an `LBound` value of 0.
//! - Essential for writing dimension-agnostic code
//! - Use with `UBound` to determine array size = `UBound` - `LBound` + 1
//! - Safer than assuming arrays start at 0
//! - `ReDim Preserve` maintains lower bounds
//! - Common in For loops:
//!
//! ```vb6
//! For i = LBound(arr) To UBound(arr)
//!
//! Next
//! ```
//!
//! ### Common Errors
//!
//! - **Error 9** (Subscript out of range): If dimension exceeds array dimensions
//! - **Error 9** (Subscript out of range): If array has not been dimensioned (for dynamic arrays)
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: `LBound` is a very fast intrinsic function
//! - **No Overhead**: Direct access to array metadata
//! - **Cache Results**: If using in loops, cache `LBound`/`UBound` values
//! - **Bounds in Loops**: Better to cache than call repeatedly
//!
//! ### Performance Optimization
//!
//! #### Less efficient - Calls `LBound` Every Iteration
//!
//! ```vb6
//! For i = LBound(arr) To UBound(arr)
//!     ' process arr(i)
//! Next i
//! ```
//!
//! #### More efficient for very large loops
//!
//! ```vb6
//! Dim lb As Long, ub As Long
//! lb = LBound(arr)
//! ub = UBound(arr)
//! For i = lb To ub
//!     ' process arr(i)
//! Next i
//! ```
//!
//! ## Typical Uses
//!
//! - **Array Iteration**: Loop through arrays with correct starting index
//! - **Array Size Calculation**: Determine number of elements
//! - **Bounds Validation**: Check if index is within valid range
//! - **Array Copying**: Copy elements with proper bounds
//! - **Multi-dimensional Arrays**: Access correct dimension bounds
//! - **Dynamic Arrays**: Verify array has been dimensioned
//! - **Generic Functions**: Write functions that work with any array bounds
//! - **`Option Base` Handling**: Code that works regardless of `Option Base` setting
//!
//! ## Limitations
//!
//! - Cannot modify array bounds (use `ReDim` for that)
//! - Raises error for undimensioned dynamic arrays
//! - Dimension parameter must be valid (1 to number of dimensions)
//! - Cannot determine if array is fixed-size or dynamic
//! - No way to get all bounds at once (must call separately for each dimension)
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core functions
//! - Returns Long type
//! - Works with all array types (Variant, typed, object arrays)
//! - `ReDim` Preserve maintains lower bounds
//! - `ParamArray` always has `LBound` = 0
//!
//! ## Related Functions
//!
//! - `UBound`: Get upper bound of array dimension
//! - `IsArray`: Check if variable is an array
//! - `ReDim`: Redimension dynamic array
//! - `Array`: Create Variant array
//! - `Split`: Create array from delimited string
//!
//! ## Comparison With Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | `LBound` | Get lower bound | `Long` | Minimum valid index |
//! | `UBound` | Get upper bound | `Long` | Maximum valid index |
//! | `IsArray` | Check if array | `Boolean` | Validate array type |
//! | `Array` | Create array | `Variant` | Initialize arrays |
//! | `ReDim` | Resize array | N/A | Dynamic array sizing |
//!
//! ## Best Practices
//!
//! - **Always Use `LBound`**: Don't assume arrays start at 0
//! - **Dimension Parameter**: Specify dimension for multi-dimensional arrays
//! - **Error Handling**: Handle undimensioned dynamic arrays
//! - **Array Size**: Use ```UBound - LBound + 1``` for element count
//! - **Cache Values**: Store `LBound`/`UBound` in variables for repeated use
//! - **Generic Code**: Write functions that work with any array bounds
//! - **Validate Bounds**: Check if indices are within `LBound` to `UBound` range
//! - **Document Assumptions**: Note expected array bounds in comments
//!
//! ## `LBound` and `Option Base`
//!
//! ### Option Base 0 (default)
//!
//! ```vb6
//! Dim arr1(5) As Integer
//! Debug.Print LBound(arr1)         ' 0
//! Debug.Print UBound(arr1)         ' 5
//! ```
//!
//! ### Option Base 1
//!
//! ```vb6
//! Option Base 1
//! Dim arr2(5) As Integer
//! Debug.Print LBound(arr2)         ' 1
//! Debug.Print UBound(arr2)         ' 5
//! ```
//!
//! ### Explicit Bounds (Overrides Option Base)
//!
//! ```vb6
//! Dim arr3(10 To 20) As Integer
//! Debug.Print LBound(arr3)         ' 10
//! Debug.Print UBound(arr3)         ' 20
//! ```
//!
//! ### Array Size Calculation
//!
//! ```vb6
//! ' Correct way to get array size
//! Function GetArraySize(arr As Variant) As Long
//!     If Not IsArray(arr) Then
//!         GetArraySize = 0
//!     Else
//!         GetArraySize = UBound(arr) - LBound(arr) + 1
//!     End If
//! End Function
//!
//! ' Examples
//! Dim a(0 To 10) As Integer       ' Size = 11
//! Dim b(1 To 10) As Integer       ' Size = 10
//! Dim c(5 To 15) As Integer       ' Size = 11
//!
//! Debug.Print GetArraySize(a)     ' 11
//! Debug.Print GetArraySize(b)     ' 10
//! Debug.Print GetArraySize(c)     ' 11
//! ```
//!
//! ## Examples
//!
//! ### Example 1: Basic Array Iteration
//!
//! ```vb6
//! Dim arr(5) As Integer
//! Dim i As Long
//!
//! For i = LBound(arr) To UBound(arr)
//!     arr(i) = i * 2
//! Next i
//! ```
//!
//! ### Example 2: Explicit Lower Bound
//!
//! ```vb6
//! Dim months(1 To 12) As String
//!
//! Debug.Print LBound(months)           ' 1
//! Debug.Print UBound(months)           ' 12
//! ```
//!
//! ### Example 3: Multi-Dimensional Array
//!
//! ```vb6
//! Dim grid(1 To 10, 1 To 20) As Integer
//!
//! Debug.Print LBound(grid, 1)          ' 1 - first dimension
//! Debug.Print LBound(grid, 2)          ' 1 - second dimension
//! Debug.Print LBound(grid)             ' 1 - defaults to first dimension
//! ```
//!
//! ### Example 4: Calculate Array Size
//!
//! ```vb6
//! Dim values(10 To 50) As Double
//! Dim size As Long
//!
//! size = UBound(values) - LBound(values) + 1
//! Debug.Print size                     ' 41 elements
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Safe Array Iteration
//!
//! ```vb6
//! Sub ProcessArray(arr As Variant)
//!     Dim i As Long
//!     
//!     If Not IsArray(arr) Then Exit Sub
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         Debug.Print arr(i)
//!     Next i
//! End Sub
//! ```
//!
//! ### Pattern 2: Array Size Function
//!
//! ```vb6
//! Function ArraySize(arr As Variant, Optional dimension As Long = 1) As Long
//!     If Not IsArray(arr) Then
//!         ArraySize = 0
//!     Else
//!         ArraySize = UBound(arr, dimension) - LBound(arr, dimension) + 1
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 3: Copy Array With Correct Bounds
//!
//! ```vb6
//! Function CopyArray(source As Variant) As Variant
//!     Dim dest As Variant
//!     Dim i As Long
//!     
//!     If Not IsArray(source) Then Exit Function
//!     
//!     ReDim dest(LBound(source) To UBound(source))
//!     For i = LBound(source) To UBound(source)
//!         If IsObject(source(i)) Then
//!             Set dest(i) = source(i)
//!         Else
//!             dest(i) = source(i)
//!         End If
//!     Next i
//!     
//!     CopyArray = dest
//! End Function
//! ```
//!
//! ### Pattern 4: Check If Array Is Empty
//!
//! ```vb6
//! Function IsArrayEmpty(arr As Variant) As Boolean
//!     On Error Resume Next
//!     IsArrayEmpty = (UBound(arr) < LBound(arr))
//!     If Err.Number <> 0 Then IsArrayEmpty = True
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ### Pattern 5: Array Contains Value
//!
//! ```vb6
//! Function ArrayContains(arr As Variant, value As Variant) As Boolean
//!     Dim i As Long
//!     
//!     If Not IsArray(arr) Then Exit Function
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         If arr(i) = value Then
//!             ArrayContains = True
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     ArrayContains = False
//! End Function
//! ```
//!
//! ### Pattern 6: Find Element Index
//!
//! ```vb6
//! Function FindInArray(arr As Variant, value As Variant) As Long
//!     Dim i As Long
//!     
//!     FindInArray = -1  ' Not found
//!     If Not IsArray(arr) Then Exit Function
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         If arr(i) = value Then
//!             FindInArray = i
//!             Exit Function
//!         End If
//!     Next i
//! End Function
//! ```
//!
//! ### Pattern 7: Reverse Array In Place
//!
//! ```vb6
//! Sub ReverseArray(arr As Variant)
//!     Dim i As Long
//!     Dim j As Long
//!     Dim temp As Variant
//!     
//!     If Not IsArray(arr) Then Exit Sub
//!     
//!     i = LBound(arr)
//!     j = UBound(arr)
//!     
//!     Do While i < j
//!         temp = arr(i)
//!         arr(i) = arr(j)
//!         arr(j) = temp
//!         i = i + 1
//!         j = j - 1
//!     Loop
//! End Sub
//! ```
//!
//! ### Pattern 8: Slice Array
//!
//! ```vb6
//! Function SliceArray(arr As Variant, startIndex As Long, endIndex As Long) As Variant
//!     Dim result() As Variant
//!     Dim i As Long
//!     Dim j As Long
//!     
//!     If Not IsArray(arr) Then Exit Function
//!     If startIndex < LBound(arr) Or endIndex > UBound(arr) Then Exit Function
//!     
//!     ReDim result(0 To endIndex - startIndex)
//!     j = 0
//!     For i = startIndex To endIndex
//!         result(j) = arr(i)
//!         j = j + 1
//!     Next i
//!     
//!     SliceArray = result
//! End Function
//! ```
//!
//! ### Pattern 9: Fill Array With Value
//!
//! ```vb6
//! Sub FillArray(arr As Variant, value As Variant)
//!     Dim i As Long
//!     
//!     If Not IsArray(arr) Then Exit Sub
//!     
//!     For i = LBound(arr) To UBound(arr)
//!         If IsObject(value) Then
//!             Set arr(i) = value
//!         Else
//!             arr(i) = value
//!         End If
//!     Next i
//! End Sub
//! ```
//!
//! ### Pattern 10: Multi-Dimensional Array Iteration
//!
//! ```vb6
//! Sub ProcessGrid(grid As Variant)
//!     Dim i As Long, j As Long
//!     
//!     If Not IsArray(grid) Then Exit Sub
//!     
//!     For i = LBound(grid, 1) To UBound(grid, 1)
//!         For j = LBound(grid, 2) To UBound(grid, 2)
//!             Debug.Print grid(i, j)
//!         Next j
//!     Next i
//! End Sub
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ### Example 1: Generic Array Utilities Class
//!
//! ```vb6
//! Public Class ArrayUtils
//!     Public Function GetSize(arr As Variant, Optional dimension As Long = 1) As Long
//!         On Error GoTo ErrorHandler
//!         
//!         If Not IsArray(arr) Then
//!             GetSize = 0
//!         Else
//!             GetSize = UBound(arr, dimension) - LBound(arr, dimension) + 1
//!         End If
//!         Exit Function
//!         
//!     ErrorHandler:
//!         GetSize = 0
//!     End Function
//!     
//!     Public Function Clone(source As Variant) As Variant
//!         Dim dest As Variant
//!         Dim i As Long
//!         
//!         If Not IsArray(source) Then Exit Function
//!         
//!         ReDim dest(LBound(source) To UBound(source))
//!         For i = LBound(source) To UBound(source)
//!             If IsObject(source(i)) Then
//!                 Set dest(i) = source(i)
//!             Else
//!                 dest(i) = source(i)
//!             End If
//!         Next i
//!         
//!         Clone = dest
//!     End Function
//!     
//!     Public Function IndexOf(arr As Variant, value As Variant) As Long
//!         Dim i As Long
//!         
//!         IndexOf = -1
//!         If Not IsArray(arr) Then Exit Function
//!         
//!         For i = LBound(arr) To UBound(arr)
//!             If arr(i) = value Then
//!                 IndexOf = i
//!                 Exit Function
//!             End If
//!         Next i
//!     End Function
//!     
//!     Public Function Reverse(arr As Variant) As Variant
//!         Dim result As Variant
//!         Dim i As Long
//!         Dim j As Long
//!         
//!         If Not IsArray(arr) Then Exit Function
//!         
//!         ReDim result(LBound(arr) To UBound(arr))
//!         j = UBound(arr)
//!         For i = LBound(arr) To UBound(arr)
//!             result(j) = arr(i)
//!             j = j - 1
//!         Next i
//!         
//!         Reverse = result
//!     End Function
//! End Class
//! ```
//!
//! ### Example 2: Safe array accessor with bounds checking
//!
//! ```vb6
//! Public Class SafeArray
//!     Private m_data As Variant
//!     
//!     Public Sub Initialize(size As Long, Optional lowerBound As Long = 0)
//!         ReDim m_data(lowerBound To lowerBound + size - 1)
//!     End Sub
//!     
//!     Public Property Get Item(index As Long) As Variant
//!         If index < LBound(m_data) Or index > UBound(m_data) Then
//!             Err.Raise 9, "SafeArray", "Index out of bounds: " & index & _
//!                       " (valid range: " & LBound(m_data) & " to " & UBound(m_data) & ")"
//!         End If
//!         
//!         If IsObject(m_data(index)) Then
//!             Set Item = m_data(index)
//!         Else
//!             Item = m_data(index)
//!         End If
//!     End Property
//!     
//!     Public Property Let Item(index As Long, value As Variant)
//!         If index < LBound(m_data) Or index > UBound(m_data) Then
//!             Err.Raise 9, "SafeArray", "Index out of bounds"
//!         End If
//!         
//!         If IsObject(value) Then
//!             Set m_data(index) = value
//!         Else
//!             m_data(index) = value
//!         End If
//!     End Property
//!     
//!     Public Property Get LowerBound() As Long
//!         LowerBound = LBound(m_data)
//!     End Property
//!     
//!     Public Property Get UpperBound() As Long
//!         UpperBound = UBound(m_data)
//!     End Property
//!     
//!     Public Property Get Count() As Long
//!         Count = UBound(m_data) - LBound(m_data) + 1
//!     End Property
//! End Class
//! ```
//!
//! ### Example 3: Matrix Operations Helper
//!
//! ```vb6
//! Public Class MatrixHelper
//!     Public Function GetRowCount(matrix As Variant) As Long
//!         If Not IsArray(matrix) Then
//!             GetRowCount = 0
//!         Else
//!             GetRowCount = UBound(matrix, 1) - LBound(matrix, 1) + 1
//!         End If
//!     End Function
//!     
//!     Public Function GetColumnCount(matrix As Variant) As Long
//!         If Not IsArray(matrix) Then
//!             GetColumnCount = 0
//!         Else
//!             GetColumnCount = UBound(matrix, 2) - LBound(matrix, 2) + 1
//!         End If
//!     End Function
//!     
//!     Public Function GetRow(matrix As Variant, rowIndex As Long) As Variant
//!         Dim result() As Variant
//!         Dim j As Long
//!         Dim k As Long
//!         
//!         If Not IsArray(matrix) Then Exit Function
//!         
//!         ReDim result(LBound(matrix, 2) To UBound(matrix, 2))
//!         For j = LBound(matrix, 2) To UBound(matrix, 2)
//!             result(j) = matrix(rowIndex, j)
//!         Next j
//!         
//!         GetRow = result
//!     End Function
//!     
//!     Public Function GetColumn(matrix As Variant, colIndex As Long) As Variant
//!         Dim result() As Variant
//!         Dim i As Long
//!         
//!         If Not IsArray(matrix) Then Exit Function
//!         
//!         ReDim result(LBound(matrix, 1) To UBound(matrix, 1))
//!         For i = LBound(matrix, 1) To UBound(matrix, 1)
//!             result(i) = matrix(i, colIndex)
//!         Next i
//!         
//!         GetColumn = result
//!     End Function
//!     
//!     Public Sub Fill(matrix As Variant, value As Variant)
//!         Dim i As Long, j As Long
//!         
//!         If Not IsArray(matrix) Then Exit Sub
//!         
//!         For i = LBound(matrix, 1) To UBound(matrix, 1)
//!             For j = LBound(matrix, 2) To UBound(matrix, 2)
//!                 matrix(i, j) = value
//!             Next j
//!         Next i
//!     End Sub
//! End Class
//! ```
//!
//! ### Example 4: Dynamic Array Manager
//!
//! ```vb6
//! Public Class DynamicArray
//!     Private m_data() As Variant
//!     Private m_count As Long
//!     Private m_lowerBound As Long
//!     
//!     Private Sub Class_Initialize()
//!         m_count = 0
//!         m_lowerBound = 0
//!         ReDim m_data(m_lowerBound To m_lowerBound + 9)  ' Initial capacity of 10
//!     End Sub
//!     
//!     Public Sub Add(value As Variant)
//!         If m_count > UBound(m_data) - LBound(m_data) Then
//!             ' Resize array (double capacity)
//!             ReDim Preserve m_data(LBound(m_data) To UBound(m_data) * 2 + 1)
//!         End If
//!         
//!         If IsObject(value) Then
//!             Set m_data(LBound(m_data) + m_count) = value
//!         Else
//!             m_data(LBound(m_data) + m_count) = value
//!         End If
//!         
//!         m_count = m_count + 1
//!     End Sub
//!     
//!     Public Function ToArray() As Variant
//!         Dim result() As Variant
//!         Dim i As Long
//!         
//!         If m_count = 0 Then
//!             ToArray = Array()
//!             Exit Function
//!         End If
//!         
//!         ReDim result(m_lowerBound To m_lowerBound + m_count - 1)
//!         For i = 0 To m_count - 1
//!             If IsObject(m_data(LBound(m_data) + i)) Then
//!                 Set result(m_lowerBound + i) = m_data(LBound(m_data) + i)
//!             Else
//!                 result(m_lowerBound + i) = m_data(LBound(m_data) + i)
//!             End If
//!         Next i
//!         
//!         ToArray = result
//!     End Function
//!     
//!     Public Property Get Count() As Long
//!         Count = m_count
//!     End Property
//!     
//!     Public Property Get Capacity() As Long
//!         Capacity = UBound(m_data) - LBound(m_data) + 1
//!     End Property
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! `LBound` can raise errors in specific cases:
//!
//! ```vb6
//! ' Error 9: Subscript out of range - dimension exceeds array dimensions
//! Dim arr(5, 10) As Integer
//! ' Debug.Print LBound(arr, 3)  ' Error 9 - only 2 dimensions
//! ```
//!
//! ```vb6
//! ' Error 9: Array not dimensioned (dynamic arrays)
//! Dim dynArr() As String
//! ' Debug.Print LBound(dynArr)  ' Error 9 - not dimensioned yet
//! ```
//!
//! ```vb6
//! ' Safe pattern with error handling
//! Function GetLowerBound(arr As Variant, Optional dimension As Long = 1) As Long
//!     On Error Resume Next
//!     GetLowerBound = LBound(arr, dimension)
//!     If Err.Number <> 0 Then
//!         GetLowerBound = -1  ' Indicate error
//!     End If
//!     On Error GoTo 0
//! End Function
//! ```
//!

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn lbound_basic() {
        let source = r"
Sub Test()
    result = LBound(myArray)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_with_dimension() {
        let source = r"
Sub Test()
    result = LBound(matrix, 2)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_for_loop() {
        let source = r"
Sub Test()
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        Debug.Print arr(i)
    Next i
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_function_return() {
        let source = r"
Function GetLowerBound(arr As Variant) As Long
    GetLowerBound = LBound(arr)
End Function
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_if_statement() {
        let source = r"
Sub Test()
    If LBound(data) = 0 Then
        ProcessZeroBased
    End If
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Lower bound: " & LBound(items)
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Array starts at: " & LBound(values)
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_variable_assignment() {
        let source = r"
Sub Test()
    Dim lb As Long
    lb = LBound(myArray)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_property_assignment() {
        let source = r"
Sub Test()
    obj.LowerBound = LBound(obj.Data)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_arithmetic() {
        let source = r"
Sub Test()
    size = UBound(arr) - LBound(arr) + 1
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_in_class() {
        let source = r"
Private Sub Class_Initialize()
    m_lowerBound = LBound(m_data)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_with_statement() {
        let source = r"
Sub Test()
    With container
        .Start = LBound(.Items)
    End With
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_function_argument() {
        let source = r"
Sub Test()
    Call ProcessBound(LBound(data))
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_select_case() {
        let source = r"
Sub Test()
    Select Case LBound(arr)
        Case 0
            ZeroBasedProcessing
        Case 1
            OneBasedProcessing
    End Select
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn lbound_elseif() {
        let source = r"
Sub Test()
    If LBound(arr) = 0 Then
        HandleZero
    ElseIf LBound(arr) = 1 Then
        HandleOne
    End If
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_concatenation() {
        let source = r#"
Sub Test()
    info = "Range: " & LBound(arr) & " to " & UBound(arr)
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_parentheses() {
        let source = r"
Sub Test()
    result = (LBound(values))
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_array_assignment() {
        let source = r"
Sub Test()
    bounds(0) = LBound(data)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_collection_add() {
        let source = r"
Sub Test()
    bounds.Add LBound(arrays(i))
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_comparison() {
        let source = r#"
Sub Test()
    If LBound(arr1) = LBound(arr2) Then
        MsgBox "Same lower bound"
    End If
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_nested_call() {
        let source = r"
Sub Test()
    result = CStr(LBound(data))
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn lbound_while_wend() {
        let source = r"
Sub Test()
    i = LBound(arr)
    While i <= UBound(arr)
        i = i + 1
    Wend
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn lbound_do_while() {
        let source = r"
Sub Test()
    i = LBound(values)
    Do While i <= UBound(values)
        i = i + 1
    Loop
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn lbound_do_until() {
        let source = r"
Sub Test()
    i = LBound(items)
    Do Until i > UBound(items)
        i = i + 1
    Loop
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_redim() {
        let source = r"
Sub Test()
    Dim lb As Long
    lb = LBound(arr)
    ReDim Preserve arr(lb To lb + 20)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn lbound_multi_dimensional() {
        let source = r"
Sub Test()
    For i = LBound(grid, 1) To UBound(grid, 1)
        For j = LBound(grid, 2) To UBound(grid, 2)
            grid(i, j) = 0
        Next j
    Next i
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn lbound_iif() {
        let source = r#"
Sub Test()
    start = IIf(LBound(arr) = 0, "Zero-based", "One-based")
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/lbound");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
