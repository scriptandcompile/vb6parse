//! # `IsNull` Function
//!
//! Returns a `Boolean` value indicating whether an expression contains no valid data (`Null`).
//!
//! ## Syntax
//!
//! ```vb
//! IsNull(expression)
//! ```
//!
//! ## Parameters
//!
//! - `expression` (Required): Variant expression to test
//!
//! ## Return Value
//!
//! Returns a Boolean:
//! - `True` if the expression is `Null`
//! - `False` if the expression contains valid data
//! - `Null` is different from `Empty` (uninitialized)
//! - `Null` is different from zero-length string ("")
//! - `Null` is different from zero (0)
//! - `Null` propagates through expressions ```(Null + 5 = Null)```
//! - Used to detect database `NULL` values
//!
//! ## Remarks
//!
//! The `IsNull` function is used to determine whether an expression evaluates to `Null`:
//!
//! - `Null` represents "no valid data" or "unknown value"
//! - Common in database operations (`NULL` field values)
//! - `Null` is different from `Empty`, zero, or empty string
//! - `Null` propagates: any operation involving `Null` yields `Null`
//! - Use `Null` for missing or unknown data
//! - Cannot compare `Null` with = operator (use `IsNull` instead)
//! - `var = Null` is always `Null` (not `True` or `False`)
//! - Only `IsNull` can reliably test for `Null`
//! - `Null` can be explicitly assigned: `myVar = Null`
//! - Only `Variant` variables can contain `Null`
//! - Common pattern: check `IsNull` before using database field values
//! - ```VarType(expr) = vbNull``` provides same functionality
//! - `Null` is tri-state: `True`, `False`, `Null` (for database three-valued logic)
//!
//! ## Typical Uses
//!
//! 1. **Database `NULL` Handling**: Check if database field contains `NULL`
//! 2. **Data Validation**: Detect missing or unknown values
//! 3. **Error Prevention**: Avoid errors from `Null` propagation
//! 4. **Optional Values**: Represent "not applicable" or "unknown"
//! 5. **Recordset Processing**: Handle `NULL` fields safely
//! 6. **Form Input**: Detect unselected combo boxes or list boxes
//! 7. **API Results**: Check for invalid return values
//! 8. **`Null` Coalescing**: Provide defaults for `Null` values
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Basic Null detection
//! Dim value As Variant
//!
//! value = Null
//!
//! If IsNull(value) Then
//!     Debug.Print "Value is Null"  ' This prints
//! End If
//!
//! ' Example 2: Distinguish Null from other values
//! Dim testVar As Variant
//!
//! testVar = Null
//! Debug.Print IsNull(testVar)        ' True - Null value
//! testVar = Empty
//! Debug.Print IsNull(testVar)        ' False - Empty is not Null
//! testVar = 0
//! Debug.Print IsNull(testVar)        ' False - Zero is not Null
//! testVar = ""
//! Debug.Print IsNull(testVar)        ' False - Empty string is not Null
//! testVar = False
//! Debug.Print IsNull(testVar)        ' False - False is not Null
//!
//! ' Example 3: Database field handling
//! Sub ProcessRecord(rs As Recordset)
//!     Dim email As String
//!     
//!     If IsNull(rs!Email) Then
//!         email = "No email provided"
//!     Else
//!         email = rs!Email
//!     End If
//!     
//!     Debug.Print "Email: " & email
//! End Sub
//!
//! ' Example 4: Null propagation demonstration
//! Dim result As Variant
//! Dim value As Variant
//!
//! value = Null
//! result = value + 10        ' result is Null (Null propagates)
//! result = value & "text"    ' result is Null (Null propagates)
//! result = value * 2         ' result is Null (Null propagates)
//!
//! If IsNull(result) Then
//!     Debug.Print "Result is Null due to propagation"  ' This prints
//! End If
//!
//! ' Cannot use = to test for Null
//! If result = Null Then      ' This condition is always Null (not True!)
//!     Debug.Print "Never prints"
//! End If
//!
//! If IsNull(result) Then     ' Correct way to test
//!     Debug.Print "This prints"
//! End If
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Null coalescing - provide default value
//! Function Coalesce(value As Variant, defaultValue As Variant) As Variant
//!     If IsNull(value) Then
//!         Coalesce = defaultValue
//!     Else
//!         Coalesce = value
//!     End If
//! End Function
//!
//! ' Usage
//! displayValue = Coalesce(rs!Phone, "N/A")
//!
//! ' Pattern 2: Safe database field retrieval
//! Function GetFieldValue(rs As Recordset, fieldName As String, _
//!                        Optional defaultValue As Variant = "") As Variant
//!     If IsNull(rs.Fields(fieldName).Value) Then
//!         GetFieldValue = defaultValue
//!     Else
//!         GetFieldValue = rs.Fields(fieldName).Value
//!     End If
//! End Function
//!
//! ' Pattern 3: Null-safe string concatenation
//! Function NullSafeConcat(ParamArray values() As Variant) As String
//!     Dim result As String
//!     Dim i As Long
//!     
//!     result = ""
//!     For i = LBound(values) To UBound(values)
//!         If Not IsNull(values(i)) Then
//!             result = result & values(i)
//!         End If
//!     Next i
//!     
//!     NullSafeConcat = result
//! End Function
//!
//! ' Pattern 4: Check multiple values for Null
//! Function AnyNull(ParamArray values() As Variant) As Boolean
//!     Dim i As Long
//!     
//!     For i = LBound(values) To UBound(values)
//!         If IsNull(values(i)) Then
//!             AnyNull = True
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     AnyNull = False
//! End Function
//!
//! ' Pattern 5: All values non-Null validation
//! Function AllNonNull(ParamArray values() As Variant) As Boolean
//!     Dim i As Long
//!     
//!     For i = LBound(values) To UBound(values)
//!         If IsNull(values(i)) Or IsEmpty(values(i)) Then
//!             AllNonNull = False
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     AllNonNull = True
//! End Function
//!
//! ' Pattern 6: Null-safe numeric conversion
//! Function NullToZero(value As Variant) As Double
//!     If IsNull(value) Then
//!         NullToZero = 0
//!     Else
//!         NullToZero = CDbl(value)
//!     End If
//! End Function
//!
//! ' Pattern 7: Null-safe comparison
//! Function NullSafeCompare(val1 As Variant, val2 As Variant) As Integer
//!     ' Returns: -1 (less), 0 (equal), 1 (greater), -999 (Null involved)
//!     If IsNull(val1) Or IsNull(val2) Then
//!         NullSafeCompare = -999  ' Indicate Null
//!         Exit Function
//!     End If
//!     
//!     If val1 < val2 Then
//!         NullSafeCompare = -1
//!     ElseIf val1 > val2 Then
//!         NullSafeCompare = 1
//!     Else
//!         NullSafeCompare = 0
//!     End If
//! End Function
//!
//! ' Pattern 8: Count non-Null values
//! Function CountNonNull(arr As Variant) As Long
//!     Dim count As Long
//!     Dim i As Long
//!     
//!     If Not IsArray(arr) Then
//!         CountNonNull = 0
//!         Exit Function
//!     End If
//!     
//!     count = 0
//!     For i = LBound(arr) To UBound(arr)
//!         If Not IsNull(arr(i)) And Not IsEmpty(arr(i)) Then
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     CountNonNull = count
//! End Function
//!
//! ' Pattern 9: Database insert with Null handling
//! Function BuildInsertSQL(table As String, values As Variant) As String
//!     Dim sql As String
//!     Dim i As Long
//!     Dim valueStr As String
//!     
//!     sql = "INSERT INTO " & table & " VALUES ("
//!     
//!     For i = LBound(values) To UBound(values)
//!         If i > LBound(values) Then sql = sql & ", "
//!         
//!         If IsNull(values(i)) Then
//!             valueStr = "NULL"
//!         ElseIf VarType(values(i)) = vbString Then
//!             valueStr = "'" & Replace(values(i), "'", "''") & "'"
//!         Else
//!             valueStr = CStr(values(i))
//!         End If
//!         
//!         sql = sql & valueStr
//!     Next i
//!     
//!     BuildInsertSQL = sql & ")"
//! End Function
//!
//! ' Pattern 10: Form field validation with Null check
//! Function ValidateRequiredField(field As Variant, fieldName As String) As Boolean
//!     If IsNull(field) Then
//!         MsgBox fieldName & " is required", vbExclamation
//!         ValidateRequiredField = False
//!     ElseIf VarType(field) = vbString Then
//!         If Trim$(field) = "" Then
//!             MsgBox fieldName & " cannot be empty", vbExclamation
//!             ValidateRequiredField = False
//!         Else
//!             ValidateRequiredField = True
//!         End If
//!     Else
//!         ValidateRequiredField = True
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Database record processor with Null handling
//! Public Class RecordProcessor
//!     Public Function ProcessRecordset(rs As Recordset) As Collection
//!         Dim results As New Collection
//!         Dim record As Dictionary
//!         Dim fld As Field
//!         
//!         Do While Not rs.EOF
//!             Set record = CreateObject("Scripting.Dictionary")
//!             
//!             For Each fld In rs.Fields
//!                 If IsNull(fld.Value) Then
//!                     record(fld.Name) = Empty  ' Convert Null to Empty
//!                 Else
//!                     record(fld.Name) = fld.Value
//!                 End If
//!             Next fld
//!             
//!             results.Add record
//!             rs.MoveNext
//!         Loop
//!         
//!         Set ProcessRecordset = results
//!     End Function
//!     
//!     Public Function GetNonNullFields(rs As Recordset) As Collection
//!         Dim fields As New Collection
//!         Dim fld As Field
//!         
//!         For Each fld In rs.Fields
//!             If Not IsNull(fld.Value) Then
//!                 fields.Add fld.Name
//!             End If
//!         Next fld
//!         
//!         Set GetNonNullFields = fields
//!     End Function
//! End Class
//!
//! ' Example 2: Null-aware data aggregator
//! Public Class DataAggregator
//!     Public Function Sum(values As Variant) As Variant
//!         ' Sum non-Null values, return Null if all Null
//!         Dim total As Double
//!         Dim count As Long
//!         Dim i As Long
//!         
//!         If Not IsArray(values) Then
//!             Sum = Null
//!             Exit Function
//!         End If
//!         
//!         total = 0
//!         count = 0
//!         
//!         For i = LBound(values) To UBound(values)
//!             If Not IsNull(values(i)) Then
//!                 total = total + values(i)
//!                 count = count + 1
//!             End If
//!         Next i
//!         
//!         If count = 0 Then
//!             Sum = Null  ' All values were Null
//!         Else
//!             Sum = total
//!         End If
//!     End Function
//!     
//!     Public Function Average(values As Variant) As Variant
//!         Dim total As Double
//!         Dim count As Long
//!         Dim i As Long
//!         
//!         If Not IsArray(values) Then
//!             Average = Null
//!             Exit Function
//!         End If
//!         
//!         total = 0
//!         count = 0
//!         
//!         For i = LBound(values) To UBound(values)
//!             If Not IsNull(values(i)) And Not IsEmpty(values(i)) Then
//!                 total = total + values(i)
//!                 count = count + 1
//!             End If
//!         Next i
//!         
//!         If count = 0 Then
//!             Average = Null
//!         Else
//!             Average = total / count
//!         End If
//!     End Function
//! End Class
//!
//! ' Example 3: Form data validator
//! Public Class FormValidator
//!     Private m_errors As Collection
//!     
//!     Private Sub Class_Initialize()
//!         Set m_errors = New Collection
//!     End Sub
//!     
//!     Public Function ValidateForm(form As Form) As Boolean
//!         Dim ctrl As Control
//!         
//!         m_errors.Clear
//!         
//!         For Each ctrl In form.Controls
//!             If TypeOf ctrl Is TextBox Then
//!                 ValidateTextBox ctrl
//!             ElseIf TypeOf ctrl Is ComboBox Then
//!                 ValidateComboBox ctrl
//!             End If
//!         Next ctrl
//!         
//!         ValidateForm = (m_errors.Count = 0)
//!     End Function
//!     
//!     Private Sub ValidateTextBox(txt As TextBox)
//!         If txt.Tag = "required" Then
//!             If IsNull(txt.Value) Or Trim$(txt.Value & "") = "" Then
//!                 m_errors.Add "Field '" & txt.Name & "' is required"
//!             End If
//!         End If
//!     End Sub
//!     
//!     Private Sub ValidateComboBox(cmb As ComboBox)
//!         If cmb.Tag = "required" Then
//!             If IsNull(cmb.Value) Then
//!                 m_errors.Add "Please select a value for '" & cmb.Name & "'"
//!             End If
//!         End If
//!     End Sub
//!     
//!     Public Function GetErrors() As Collection
//!         Set GetErrors = m_errors
//!     End Function
//! End Class
//!
//! ' Example 4: SQL query builder with Null-safe WHERE clause
//! Public Class QueryBuilder
//!     Public Function BuildWhereClause(conditions As Dictionary) As String
//!         Dim sql As String
//!         Dim key As Variant
//!         Dim value As Variant
//!         Dim first As Boolean
//!         
//!         first = True
//!         sql = ""
//!         
//!         For Each key In conditions.Keys
//!             value = conditions(key)
//!             
//!             If Not first Then
//!                 sql = sql & " AND "
//!             Else
//!                 first = False
//!             End If
//!             
//!             If IsNull(value) Then
//!                 sql = sql & key & " IS NULL"
//!             ElseIf VarType(value) = vbString Then
//!                 sql = sql & key & " = '" & Replace(value, "'", "''") & "'"
//!             Else
//!                 sql = sql & key & " = " & value
//!             End If
//!         Next key
//!         
//!         If sql <> "" Then
//!             BuildWhereClause = "WHERE " & sql
//!         Else
//!             BuildWhereClause = ""
//!         End If
//!     End Function
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! The `IsNull` function itself does not raise errors:
//!
//! ```vb
//! ' IsNull is safe to call on any value
//! Debug.Print IsNull(123)           ' False
//! Debug.Print IsNull("text")        ' False
//! Debug.Print IsNull(Null)          ' True
//! Debug.Print IsNull(Empty)         ' False
//! Debug.Print IsNull(rs!Field)      ' True or False depending on field
//!
//! ' Common mistake: using = to test for Null
//! Dim value As Variant
//! value = Null
//!
//! ' WRONG - this doesn't work!
//! If value = Null Then              ' Condition is Null (not True!)
//!     Debug.Print "Never prints"
//! End If
//!
//! ' CORRECT - use IsNull
//! If IsNull(value) Then
//!     Debug.Print "This prints"
//! End If
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: `IsNull` is a very fast type check
//! - **Database Overhead**: `Null` checking is critical for database operations
//! - **Propagation**: Be aware of `Null` propagation in calculations
//! - **Early Check**: Check `IsNull` early to avoid `Null` propagation issues
//!
//! ## Best Practices
//!
//! 1. **Always Check Database Fields**: Use `IsNull` for all database field access
//! 2. **Never Use = Null**: Always use `IsNull`, never `var = Null`
//! 3. **Provide Defaults**: Use `Null` coalescing pattern for display values
//! 4. **Document Null Behavior**: Clearly document when functions can return `Null`
//! 5. **Combine Checks**: Often check both `IsNull` and `IsEmpty` for complete validation
//! 6. **Handle Propagation**: Be aware that `Null` propagates through expressions
//! 7. **Database Inserts**: Convert `Null` to SQL `NULL` in INSERT/UPDATE statements
//! 8. **Form Validation**: Check for `Null` in combo boxes and optional fields
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | `IsNull` | Check if `Null` | `Boolean` | Detect `Null` values |
//! | `IsEmpty` | Check if uninitialized | `Boolean` | Detect `Empty` `Variants` |
//! | `IsError` | Check if error value | `Boolean` | Detect `CVErr` error values |
//! | `IsMissing` | Check if parameter omitted | `Boolean` | Detect missing Optional `Variant` |
//! | `VarType` | Get variant type | `Integer` | Detailed type information |
//! | `TypeName` | Get type name | `String` | Type name as `String` |
//! | `Nz` (Access) | `Null` to zero/string | `Variant` | MS Access `Null` coalescing |
//!
//! ## `Null` vs `Empty` vs `Zero` vs `Empty` `String`
//!
//! ```vb
//! Dim v As Variant
//!
//! ' Null (no valid data)
//! v = Null
//! Debug.Print IsNull(v)          ' True
//! Debug.Print IsEmpty(v)         ' False
//! Debug.Print v = 0              ' Null (not True or False!)
//! Debug.Print v & ""             ' "" (Null coalesces to empty in string context)
//!
//! ' Empty (uninitialized)
//! Dim v2 As Variant
//! Debug.Print IsNull(v2)         ' False
//! Debug.Print IsEmpty(v2)        ' True
//! Debug.Print v2 = 0             ' True (Empty coerces to 0)
//!
//! ' Zero
//! v = 0
//! Debug.Print IsNull(v)          ' False
//! Debug.Print v = 0              ' True
//!
//! ' Empty String
//! v = ""
//! Debug.Print IsNull(v)          ' False
//! Debug.Print v = ""             ' True
//! ```
//!
//! ## Null Propagation
//!
//! ```vb
//! Dim value As Variant
//! Dim result As Variant
//!
//! value = Null
//!
//! ' All arithmetic operations propagate Null
//! result = value + 5             ' result = Null
//! result = value * 2             ' result = Null
//! result = value / 10            ' result = Null
//! result = value ^ 2             ' result = Null
//!
//! ' String concatenation with & doesn't propagate Null
//! result = value & "text"        ' result = "text" (Null becomes "")
//!
//! ' String concatenation with + propagates Null
//! result = value + "text"        ' result = Null
//!
//! ' Logical operations propagate Null
//! result = value And True        ' result = Null
//! result = value Or False        ' result = Null
//!
//! ' Comparison operations return Null
//! result = (value = 5)           ' result = Null (not True or False!)
//! result = (value > 0)           ' result = Null
//! ```
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core functions
//! - Returns `Boolean` type
//! - Only `Variant` variables can contain `Null`
//! - Critical for database programming
//! - MS Access has additional `Nz()` function for `Null` coalescing
//!
//! ## Limitations
//!
//! - Cannot use = operator to test for `Null` (must use `IsNull`)
//! - Only `Variant` type can contain `Null`
//! - `Null` propagates through expressions (can be surprising)
//! - No built-in `Null` coalescing operator (must use `IIf` or custom function)
//! - `Null` in `If` statement is treated as `False` (can be confusing)
//! - No way to distinguish "`Null` from database" vs "assigned `Null`"
//!
//! ## Related Functions
//!
//! - `IsEmpty`: Check if `Variant` is uninitialized (`Empty`)
//! - `VarType`: Get detailed `Variant` type information (vbNull = 1)
//! - `TypeName`: Get type name as string ("Null" for `Null` values)
//! - `IIf`: Can be used for simple `Null` coalescing: `IIf(IsNull(v), default, v)`
//! - `Nz` (Access only): `Null` to zero/string conversion

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn isnull_basic() {
        let source = r"
Sub Test()
    result = IsNull(myVariable)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_if_statement() {
        let source = r"
Sub Test()
    If IsNull(value) Then
        value = defaultValue
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_not_condition() {
        let source = r"
Sub Test()
    If Not IsNull(field) Then
        ProcessField field
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_function_return() {
        let source = r"
Function IsValid(v As Variant) As Boolean
    IsValid = Not IsNull(v)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_boolean_and() {
        let source = r"
Sub Test()
    If IsNull(field1) And IsNull(field2) Then
        ShowError
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_boolean_or() {
        let source = r"
Sub Test()
    If IsNull(value) Or IsEmpty(value) Then
        UseDefault
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_iif() {
        let source = r#"
Sub Test()
    displayValue = IIf(IsNull(value), "N/A", value)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Is null: " & IsNull(testVar)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Field status: " & IsNull(dbField)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_do_while() {
        let source = r"
Sub Test()
    Do While IsNull(currentValue)
        currentValue = GetNextValue()
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_do_until() {
        let source = r"
Sub Test()
    Do Until Not IsNull(result)
        result = TryGetResult()
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_variable_assignment() {
        let source = r"
Sub Test()
    Dim isNull As Boolean
    isNull = IsNull(dataValue)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_property_assignment() {
        let source = r"
Sub Test()
    obj.IsNullValue = IsNull(obj.Data)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_in_class() {
        let source = r"
Private Sub Class_Initialize()
    m_isNull = IsNull(m_value)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_with_statement() {
        let source = r#"
Sub Test()
    With recordset
        .HasNull = IsNull(.Fields("Email"))
    End With
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_function_argument() {
        let source = r"
Sub Test()
    Call ValidateField(IsNull(rs!Name))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_select_case() {
        let source = r"
Sub Test()
    Select Case True
        Case IsNull(value)
            HandleNull
        Case Else
            ProcessValue
    End Select
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_for_loop() {
        let source = r"
Sub Test()
    Dim i As Integer
    For i = 0 To UBound(arr)
        If IsNull(arr(i)) Then
            arr(i) = 0
        End If
    Next i
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_elseif() {
        let source = r"
Sub Test()
    If IsEmpty(data) Then
        ProcessEmpty
    ElseIf IsNull(data) Then
        ProcessNull
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_concatenation() {
        let source = r#"
Sub Test()
    status = "Null: " & IsNull(variable)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_parentheses() {
        let source = r"
Sub Test()
    result = (IsNull(value))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_array_check() {
        let source = r"
Sub Test()
    checks(i) = IsNull(values(i))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_collection_add() {
        let source = r"
Sub Test()
    nullStates.Add IsNull(data(i))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_comparison() {
        let source = r#"
Sub Test()
    If IsNull(var1) = IsNull(var2) Then
        MsgBox "Same state"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_nested_call() {
        let source = r"
Sub Test()
    result = CStr(IsNull(myVar))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_while_wend() {
        let source = r"
Sub Test()
    While IsNull(buffer)
        buffer = ReadNextValue()
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn isnull_database_field() {
        let source = r#"
Sub ProcessRecord(rs As Recordset)
    If IsNull(rs!Email) Then
        email = "N/A"
    Else
        email = rs!Email
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/type_checking/isnull",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
