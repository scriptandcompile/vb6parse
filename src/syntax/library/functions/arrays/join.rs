//! # `Join` Function
//!
//! Returns a string created by joining a number of substrings contained in an array.
//!
//! ## Syntax
//!
//! ```vb
//! Join(sourcearray, [delimiter])
//! ```
//!
//! ## Parameters
//!
//! - `sourcearray` (Required): One-dimensional array containing substrings to be joined
//! - `delimiter` (Optional): `String` used to separate the substrings in the returned string
//!   - If omitted, space character (" ") is used
//!   - If empty string (""), items are concatenated with no separator
//!
//! ## Return Value
//!
//! Returns a `String`:
//! - `String` containing all elements of the array joined by the delimiter
//! - `Empty` string ("") if array has zero length
//! - Returns `Null` if `sourcearray` is `Null`
//! - Each array element is converted to `String` before joining
//! - Non-string elements are automatically converted using `Str`/`CStr`
//! - `Empty` array elements become empty strings in result
//! - Trailing/leading spaces in delimiter are preserved
//!
//! ## Remarks
//!
//! The `Join` function is the inverse of the Split function:
//!
//! - Combines array elements into a single string
//! - Only works with one-dimensional arrays
//! - Array elements are converted to strings automatically
//! - Default delimiter is a space (" ")
//! - `Empty` string delimiter concatenates without separators
//! - `Null` array returns `Null` (not an error)
//! - Empty array (zero length) returns empty string
//! - Preserves empty array elements as empty strings
//! - Very efficient for building strings from multiple parts
//! - Much faster than repeated string concatenation in loops
//! - Available in VB6 and VBA (added in VB6/Office 2000)
//! - Common in text processing and file generation
//! - Works with Variant arrays containing mixed types
//! - Does not add delimiter after last element
//!
//! ## Typical Uses
//!
//! 1. **CSV Generation**: Create comma-separated value strings
//! 2. **Path Building**: Combine path components with backslashes
//! 3. **SQL Generation**: Build SQL queries from parts
//! 4. **Text Formatting**: Create formatted text from arrays
//! 5. **File Output**: Generate text file content
//! 6. **URL Building**: Construct URLs from components
//! 7. **String Building**: Efficient alternative to concatenation loops
//! 8. **Report Generation**: Format report lines from data arrays
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Basic join with default delimiter (space)
//! Dim words(2) As String
//! words(0) = "Hello"
//! words(1) = "Visual"
//! words(2) = "Basic"
//!
//! Debug.Print Join(words)              ' "Hello Visual Basic"
//!
//! ' Example 2: Join with custom delimiter
//! Dim values(3) As String
//! values(0) = "apple"
//! values(1) = "banana"
//! values(2) = "cherry"
//! values(3) = "date"
//!
//! Debug.Print Join(values, ", ")       ' "apple, banana, cherry, date"
//! Debug.Print Join(values, " | ")      ' "apple | banana | cherry | date"
//! Debug.Print Join(values, "")         ' "applebananacherrydate"
//!
//! ' Example 3: CSV generation
//! Dim fields(2) As String
//! fields(0) = "John Doe"
//! fields(1) = "Engineer"
//! fields(2) = "50000"
//!
//! Dim csvLine As String
//! csvLine = Join(fields, ",")
//! Debug.Print csvLine                  ' "John Doe,Engineer,50000"
//!
//! ' Example 4: Working with Split and Join
//! Dim original As String
//! Dim parts() As String
//! Dim rebuilt As String
//!
//! original = "one-two-three-four"
//! parts = Split(original, "-")
//! rebuilt = Join(parts, " ")
//! Debug.Print rebuilt                  ' "one two three four"
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Build CSV row
//! Function BuildCSVRow(fields As Variant) As String
//!     BuildCSVRow = Join(fields, ",")
//! End Function
//!
//! ' Pattern 2: Join with line breaks
//! Function JoinLines(lines As Variant) As String
//!     JoinLines = Join(lines, vbCrLf)
//! End Function
//!
//! ' Pattern 3: Build path from components
//! Function BuildPath(ParamArray parts() As Variant) As String
//!     Dim arr() As String
//!     Dim i As Long
//!     
//!     ReDim arr(LBound(parts) To UBound(parts))
//!     For i = LBound(parts) To UBound(parts)
//!         arr(i) = CStr(parts(i))
//!     Next i
//!     
//!     BuildPath = Join(arr, "\")
//! End Function
//!
//! ' Pattern 4: Create comma-separated list
//! Function ToCommaSeparated(items As Variant) As String
//!     If IsArray(items) Then
//!         ToCommaSeparated = Join(items, ", ")
//!     Else
//!         ToCommaSeparated = CStr(items)
//!     End If
//! End Function
//!
//! ' Pattern 5: Build SQL IN clause
//! Function BuildInClause(values As Variant) As String
//!     Dim i As Long
//!     Dim quoted() As String
//!     
//!     If Not IsArray(values) Then Exit Function
//!     
//!     ReDim quoted(LBound(values) To UBound(values))
//!     For i = LBound(values) To UBound(values)
//!         quoted(i) = "'" & Replace(CStr(values(i)), "'", "''") & "'"
//!     Next i
//!     
//!     BuildInClause = Join(quoted, ", ")
//! End Function
//!
//! ' Pattern 6: Join non-empty values only
//! Function JoinNonEmpty(arr As Variant, delimiter As String) As String
//!     Dim result As Collection
//!     Dim i As Long
//!     Dim temp() As String
//!     Dim count As Long
//!     
//!     If Not IsArray(arr) Then Exit Function
//!     
//!     Set result = New Collection
//!     For i = LBound(arr) To UBound(arr)
//!         If Len(arr(i)) > 0 Then
//!             result.Add CStr(arr(i))
//!         End If
//!     Next i
//!     
//!     If result.Count = 0 Then
//!         JoinNonEmpty = ""
//!         Exit Function
//!     End If
//!     
//!     ReDim temp(0 To result.Count - 1)
//!     For i = 1 To result.Count
//!         temp(i - 1) = result(i)
//!     Next i
//!     
//!     JoinNonEmpty = Join(temp, delimiter)
//! End Function
//!
//! ' Pattern 7: Format array for display
//! Function FormatArray(arr As Variant) As String
//!     If Not IsArray(arr) Then
//!         FormatArray = CStr(arr)
//!     Else
//!         FormatArray = "[" & Join(arr, ", ") & "]"
//!     End If
//! End Function
//!
//! ' Pattern 8: Build WHERE clause
//! Function BuildWhereClause(conditions As Variant) As String
//!     If Not IsArray(conditions) Then Exit Function
//!     
//!     If UBound(conditions) < LBound(conditions) Then
//!         BuildWhereClause = ""
//!     Else
//!         BuildWhereClause = Join(conditions, " AND ")
//!     End If
//! End Function
//!
//! ' Pattern 9: Create delimited string with quotes
//! Function JoinQuoted(items As Variant, delimiter As String) As String
//!     Dim i As Long
//!     Dim quoted() As String
//!     
//!     If Not IsArray(items) Then Exit Function
//!     
//!     ReDim quoted(LBound(items) To UBound(items))
//!     For i = LBound(items) To UBound(items)
//!         quoted(i) = Chr(34) & items(i) & Chr(34)  ' Chr(34) = "
//!     Next i
//!     
//!     JoinQuoted = Join(quoted, delimiter)
//! End Function
//!
//! ' Pattern 10: Reverse of Split for round-trip
//! Function ReverseTransform(text As String) As String
//!     Dim parts() As String
//!     Dim i As Long
//!     
//!     parts = Split(text, " ")
//!     
//!     ' Reverse array
//!     For i = LBound(parts) To (UBound(parts) - LBound(parts)) \ 2 + LBound(parts)
//!         Dim temp As String
//!         temp = parts(i)
//!         parts(i) = parts(UBound(parts) - (i - LBound(parts)))
//!         parts(UBound(parts) - (i - LBound(parts))) = temp
//!     Next i
//!     
//!     ReverseTransform = Join(parts, " ")
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: CSV Builder with proper escaping
//! Public Class CSVBuilder
//!     Private m_rows As Collection
//!     
//!     Private Sub Class_Initialize()
//!         Set m_rows = New Collection
//!     End Sub
//!     
//!     Public Sub AddRow(ParamArray values() As Variant)
//!         Dim i As Long
//!         Dim fields() As String
//!         
//!         ReDim fields(LBound(values) To UBound(values))
//!         For i = LBound(values) To UBound(values)
//!             fields(i) = EscapeCSV(CStr(values(i)))
//!         Next i
//!         
//!         m_rows.Add Join(fields, ",")
//!     End Sub
//!     
//!     Private Function EscapeCSV(value As String) As String
//!         If InStr(value, ",") > 0 Or InStr(value, Chr(34)) > 0 Or _
//!            InStr(value, vbCrLf) > 0 Then
//!             ' Need to quote and escape
//!             EscapeCSV = Chr(34) & Replace(value, Chr(34), Chr(34) & Chr(34)) & Chr(34)
//!         Else
//!             EscapeCSV = value
//!         End If
//!     End Function
//!     
//!     Public Function GetCSV() As String
//!         Dim i As Long
//!         Dim lines() As String
//!         
//!         If m_rows.Count = 0 Then
//!             GetCSV = ""
//!             Exit Function
//!         End If
//!         
//!         ReDim lines(0 To m_rows.Count - 1)
//!         For i = 1 To m_rows.Count
//!             lines(i - 1) = m_rows(i)
//!         Next i
//!         
//!         GetCSV = Join(lines, vbCrLf)
//!     End Function
//!     
//!     Public Sub Clear()
//!         Set m_rows = New Collection
//!     End Sub
//! End Class
//!
//! ' Example 2: String builder for efficient concatenation
//! Public Class StringBuilder
//!     Private m_parts As Collection
//!     Private m_delimiter As String
//!     
//!     Private Sub Class_Initialize()
//!         Set m_parts = New Collection
//!         m_delimiter = ""
//!     End Sub
//!     
//!     Public Property Let Delimiter(value As String)
//!         m_delimiter = value
//!     End Property
//!     
//!     Public Sub Append(text As String)
//!         m_parts.Add text
//!     End Sub
//!     
//!     Public Sub AppendLine(text As String)
//!         m_parts.Add text & vbCrLf
//!     End Sub
//!     
//!     Public Function ToString() As String
//!         Dim i As Long
//!         Dim arr() As String
//!         
//!         If m_parts.Count = 0 Then
//!             ToString = ""
//!             Exit Function
//!         End If
//!         
//!         ReDim arr(0 To m_parts.Count - 1)
//!         For i = 1 To m_parts.Count
//!             arr(i - 1) = m_parts(i)
//!         Next i
//!         
//!         ToString = Join(arr, m_delimiter)
//!     End Function
//!     
//!     Public Sub Clear()
//!         Set m_parts = New Collection
//!     End Sub
//!     
//!     Public Function Length() As Long
//!         Length = Len(ToString())
//!     End Function
//! End Class
//!
//! ' Example 3: Query builder using Join
//! Public Class QueryBuilder
//!     Private m_select As Collection
//!     Private m_from As String
//!     Private m_where As Collection
//!     Private m_orderBy As Collection
//!     
//!     Private Sub Class_Initialize()
//!         Set m_select = New Collection
//!         Set m_where = New Collection
//!         Set m_orderBy = New Collection
//!     End Sub
//!     
//!     Public Sub AddField(fieldName As String)
//!         m_select.Add fieldName
//!     End Sub
//!     
//!     Public Sub SetTable(tableName As String)
//!         m_from = tableName
//!     End Sub
//!     
//!     Public Sub AddCondition(condition As String)
//!         m_where.Add condition
//!     End Sub
//!     
//!     Public Sub AddOrderBy(fieldName As String)
//!         m_orderBy.Add fieldName
//!     End Sub
//!     
//!     Public Function BuildSQL() As String
//!         Dim sql As String
//!         Dim fields() As String
//!         Dim conditions() As String
//!         Dim orderFields() As String
//!         Dim i As Long
//!         
//!         ' SELECT clause
//!         If m_select.Count = 0 Then
//!             sql = "SELECT *"
//!         Else
//!             ReDim fields(0 To m_select.Count - 1)
//!             For i = 1 To m_select.Count
//!                 fields(i - 1) = m_select(i)
//!             Next i
//!             sql = "SELECT " & Join(fields, ", ")
//!         End If
//!         
//!         ' FROM clause
//!         If m_from <> "" Then
//!             sql = sql & " FROM " & m_from
//!         End If
//!         
//!         ' WHERE clause
//!         If m_where.Count > 0 Then
//!             ReDim conditions(0 To m_where.Count - 1)
//!             For i = 1 To m_where.Count
//!                 conditions(i - 1) = m_where(i)
//!             Next i
//!             sql = sql & " WHERE " & Join(conditions, " AND ")
//!         End If
//!         
//!         ' ORDER BY clause
//!         If m_orderBy.Count > 0 Then
//!             ReDim orderFields(0 To m_orderBy.Count - 1)
//!             For i = 1 To m_orderBy.Count
//!                 orderFields(i - 1) = m_orderBy(i)
//!             Next i
//!             sql = sql & " ORDER BY " & Join(orderFields, ", ")
//!         End If
//!         
//!         BuildSQL = sql
//!     End Function
//!     
//!     Public Sub Clear()
//!         Set m_select = New Collection
//!         m_from = ""
//!         Set m_where = New Collection
//!         Set m_orderBy = New Collection
//!     End Sub
//! End Class
//!
//! ' Example 4: Report formatter
//! Public Class ReportFormatter
//!     Public Function FormatTable(data As Variant, headers As Variant, _
//!                                  Optional delimiter As String = " | ") As String
//!         Dim lines As Collection
//!         Dim i As Long, j As Long
//!         Dim row() As String
//!         Dim allLines() As String
//!         
//!         Set lines = New Collection
//!         
//!         ' Add header
//!         If IsArray(headers) Then
//!             lines.Add Join(headers, delimiter)
//!             
//!             ' Add separator
//!             ReDim row(LBound(headers) To UBound(headers))
//!             For j = LBound(headers) To UBound(headers)
//!                 row(j) = String(Len(headers(j)), "-")
//!             Next j
//!             lines.Add Join(row, delimiter)
//!         End If
//!         
//!         ' Add data rows
//!         If IsArray(data) Then
//!             For i = LBound(data) To UBound(data)
//!                 If IsArray(data(i)) Then
//!                     lines.Add Join(data(i), delimiter)
//!                 Else
//!                     lines.Add CStr(data(i))
//!                 End If
//!             Next i
//!         End If
//!         
//!         ' Convert collection to array and join
//!         ReDim allLines(0 To lines.Count - 1)
//!         For i = 1 To lines.Count
//!             allLines(i - 1) = lines(i)
//!         Next i
//!         
//!         FormatTable = Join(allLines, vbCrLf)
//!     End Function
//!     
//!     Public Function FormatList(items As Variant, _
//!                                Optional prefix As String = "- ") As String
//!         Dim i As Long
//!         Dim lines() As String
//!         
//!         If Not IsArray(items) Then
//!             FormatList = prefix & CStr(items)
//!             Exit Function
//!         End If
//!         
//!         ReDim lines(LBound(items) To UBound(items))
//!         For i = LBound(items) To UBound(items)
//!             lines(i) = prefix & CStr(items(i))
//!         Next i
//!         
//!         FormatList = Join(lines, vbCrLf)
//!     End Function
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! Join handles several special cases:
//!
//! ```vb
//! ' Empty array returns empty string
//! Dim emptyArr() As String
//! ReDim emptyArr(0 To -1)  ' Zero-length array
//! Debug.Print Join(emptyArr, ",")  ' Returns ""
//!
//! ' Null array returns Null
//! Dim nullArr As Variant
//! nullArr = Null
//! Debug.Print IsNull(Join(nullArr, ","))  ' True
//!
//! ' Works with mixed-type Variant arrays
//! Dim mixed(2) As Variant
//! mixed(0) = 123
//! mixed(1) = "text"
//! mixed(2) = True
//! Debug.Print Join(mixed, "-")  ' "123-text-True"
//!
//! ' Multi-dimensional arrays cause Type Mismatch error
//! Dim multi(1, 1) As String
//! ' Join(multi, ",")  ' Error 13: Type Mismatch
//! ```
//!
//! ## Performance Considerations
//!
//! - **Very Efficient**: Join is much faster than repeated concatenation
//! - **String Building**: Use Join instead of concatenation in loops
//! - **Memory Usage**: Creates single string allocation for result
//! - **Large Arrays**: Handles large arrays efficiently
//!
//! Performance comparison:
//! ```vb
//! ' SLOW: Repeated concatenation
//! Dim result As String
//! For i = 0 To 999
//!     result = result & arr(i) & ","
//! Next i
//!
//! ' FAST: Using Join
//! result = Join(arr, ",")
//! ```
//!
//! ## Best Practices
//!
//! 1. **Use Join for String Building**: Much faster than repeated concatenation
//! 2. **CSV Generation**: Properly escape values containing delimiters
//! 3. **Empty Delimiter**: Use "" to concatenate without separators
//! 4. **Check Array**: Verify array exists before calling `Join`
//! 5. **Null Handling**: Be aware `Join` returns `Null` for `Null` arrays
//! 6. **Line Breaks**: Use `vbCrLf`, `vbLf`, or `vbCr` as delimiter for multi-line text
//! 7. **Collection to String**: Convert `Collection` to array first, then `Join`
//! 8. **Type Conversion**: `Join` automatically converts non-string elements
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Input | Output |
//! |----------|---------|-------|--------|
//! | `Join` | Combine array to string | `Array` | `String` |
//! | `Split` | Split string to array | `String` | `Array` |
//! | `String` concatenation (&) | Combine two strings | `Strings` | `String` |
//! | `Filter` | Filter array elements | `Array` | `Array` |
//! | `UBound`/`LBound` | Get array bounds | `Array` | `Long` |
//!
//! ## `Join` vs `String` Concatenation
//!
//! ```vb
//! Dim arr(2) As String
//! arr(0) = "A"
//! arr(1) = "B"
//! arr(2) = "C"
//!
//! ' Using Join (FAST)
//! result = Join(arr, ",")              ' "A,B,C"
//!
//! ' Using concatenation (SLOW)
//! result = arr(0) & "," & arr(1) & "," & arr(2)  ' "A,B,C"
//!
//! ' For large arrays, Join is dramatically faster
//! ```
//!
//! ## `Join` and `Split` Round-Trip
//!
//! ```vb
//! ' Original string
//! original = "apple,banana,cherry"
//!
//! ' Split into array
//! parts = Split(original, ",")         ' ["apple", "banana", "cherry"]
//!
//! ' Join back to string
//! rebuilt = Join(parts, ",")           ' "apple,banana,cherry"
//!
//! Debug.Print original = rebuilt       ' True - perfect round-trip
//! ```
//!
//! ## Platform and Version Notes
//!
//! - Added in VB6 and Office 2000 VBA
//! - Not available in VB5 or earlier
//! - Part of `VBA.Strings` module
//! - Returns `String` type
//! - Only works with one-dimensional arrays
//! - Automatically converts array elements to `String`
//!
//! ## Limitations
//!
//! - Cannot join multi-dimensional arrays (use loops to flatten first)
//! - Returns `Null` for `Null` array (not empty `String`)
//! - No built-in escaping for CSV (must implement manually)
//! - Cannot skip empty elements automatically
//! - No formatting options for numeric values
//! - Delimiter is applied between all elements (no custom logic)
//!
//! ## Related Functions
//!
//! - `Split`: Split string into array (inverse of `Join`)
//! - `Filter`: Filter array elements based on criteria
//! - `UBound`/`LBound`: Get array bounds
//! - `Array`: Create array from values
//! - `Replace`: Replace substrings in `String`

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn join_basic() {
        let source = r"
Sub Test()
    result = Join(myArray)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_with_delimiter() {
        let source = r#"
Sub Test()
    result = Join(items, ", ")
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_if_statement() {
        let source = r#"
Sub Test()
    If Len(Join(parts, "-")) > 0 Then
        ProcessResult
    End If
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_function_return() {
        let source = r#"
Function GetCSV(fields As Variant) As String
    GetCSV = Join(fields, ",")
End Function
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print Join(values, " | ")
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_msgbox() {
        let source = r"
Sub Test()
    MsgBox Join(names, vbCrLf)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_variable_assignment() {
        let source = r#"
Sub Test()
    Dim combined As String
    combined = Join(words, " ")
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_property_assignment() {
        let source = r"
Sub Test()
    obj.DisplayText = Join(obj.Lines, vbCrLf)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_concatenation() {
        let source = r#"
Sub Test()
    result = "Values: " & Join(data, ", ")
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_in_class() {
        let source = r#"
Private Sub Class_Initialize()
    m_text = Join(m_parts, "")
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_with_statement() {
        let source = r"
Sub Test()
    With builder
        .Output = Join(.Parts, .Delimiter)
    End With
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_function_argument() {
        let source = r"
Sub Test()
    Call ProcessText(Join(lines, vbCrLf))
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_select_case() {
        let source = r#"
Sub Test()
    Select Case Join(tags, ",")
        Case "A,B,C"
            HandleABC
        Case Else
            HandleOther
    End Select
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn join_for_loop() {
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 0 To UBound(dataRows)
        lines(i) = Join(dataRows(i), ",")
    Next i
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    #[allow(clippy::too_many_lines)]
    fn join_elseif() {
        let source = r#"
Sub Test()
    If format = "csv" Then
        output = Join(data, ",")
    ElseIf format = "tsv" Then
        output = Join(data, vbTab)
    End If
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_iif() {
        let source = r#"
Sub Test()
    result = IIf(useComma, Join(arr, ","), Join(arr, ";"))
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_parentheses() {
        let source = r#"
Sub Test()
    result = (Join(values, "-"))
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_array_assignment() {
        let source = r#"
Sub Test()
    csvRows(i) = Join(fields(i), ",")
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_collection_add() {
        let source = r"
Sub Test()
    lines.Add Join(row, vbTab)
End Sub
";
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_comparison() {
        let source = r#"
Sub Test()
    If Join(actual, ",") = Join(expected, ",") Then
        MsgBox "Match!"
    End If
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_nested_call() {
        let source = r#"
Sub Test()
    result = UCase(Join(names, " "))
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_while_wend() {
        let source = r#"
Sub Test()
    While Len(Join(buffer, "")) < maxLen
        buffer.Add GetNextItem()
    Wend
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_do_while() {
        let source = r#"
Sub Test()
    Do While Len(Join(parts, "")) > 0
        ProcessParts
    Loop
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_do_until() {
        let source = r#"
Sub Test()
    Do Until Join(fields, "") = ""
        fields = GetNextFields()
    Loop
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_with_split() {
        let source = r#"
Sub Test()
    parts = Split(text, "-")
    result = Join(parts, " ")
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_csv_builder() {
        let source = r#"
Function BuildCSV(fields As Variant) As String
    BuildCSV = Join(fields, ",")
End Function
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn join_empty_delimiter() {
        let source = r#"
Sub Test()
    concatenated = Join(chars, "")
End Sub
"#;
        let (cst_opt, _failure) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path("../../../../../snapshots/syntax/library/functions/arrays/join");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
