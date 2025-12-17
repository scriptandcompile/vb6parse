//! # `CallByName` Function
//!
//! Executes a method, sets or returns a property, or sets or returns a field of an object.
//!
//! ## Syntax
//!
//! ```vb
//! CallByName(object, procname, calltype, [args()])
//! ```
//!
//! ## Parameters
//!
//! - `object` - Required. Object expression on which the function will be executed.
//! - `procname` - Required. String expression containing the name of the property, method, or field member of the object.
//! - `calltype` - Required. Member of the `VbCallType` enumeration representing the type of procedure being called.
//! - `args()` - Optional. `Variant` array containing the arguments to be passed to the property, method, or field being called.
//!
//! ## `VbCallType` Constants
//!
//! | Constant | Value | Description |
//! |----------|-------|-------------|
//! | `VbMethod` | 1 | A method is being called |
//! | `VbGet` | 2 | A property value is being retrieved |
//! | `VbLet` | 4 | A property value is being set |
//! | `VbSet` | 8 | A reference to an object is being set |
//!
//! ## Return Value
//!
//! Returns a `Variant` containing the result of the called property or method. For `VbLet` and `VbSet`,
//! the return value is not meaningful.
//!
//! ## Remarks
//!
//! The `CallByName` function allows you to execute object members by name at run time, providing
//! a form of late binding. This is particularly useful for:
//!
//! - Creating generic routines that work with multiple object types
//! - Implementing dynamic property access based on user input
//! - Building reflection-like functionality in VB6
//! - Simplifying repetitive property access code
//!
//! ### Important Notes
//!
//! 1. **Late Binding**: `CallByName` always uses late binding, even if the object variable is early bound
//! 2. **Performance**: Slower than direct method/property calls due to name lookup overhead
//! 3. **Case Insensitive**: The `procname` parameter is case-insensitive
//! 4. **Error Handling**: Raises run-time error if the member doesn't exist
//! 5. **Type Safety**: No compile-time checking of member existence or argument types
//!
//! ### Call Type Details
//!
//! **`VbMethod` (1)**:
//! - Calls a Sub or Function
//! - Returns the function's return value (or Empty for Subs)
//! - Passes arguments in the args array
//!
//! **`VbGet` (2)**:
//! - Retrieves a property value or field value
//! - Can be used with property procedures (Property Get) or public fields
//! - For parameterized properties, pass indices in args array
//!
//! **`VbLet` (4)**:
//! - Sets a `property` value or field value
//! - For simple data types (numbers, strings, etc.)
//! - The new value must be the last element in the args array
//! - For parameterized properties, indices come before the value
//!
//! **`VbSet` (8)**:
//! - Sets an object reference property
//! - Similar to `VbLet` but for object references
//! - Used when you would normally use the Set keyword
//!
//! ## Examples
//!
//! ### Basic Method Call
//!
//! ```vb
//! Dim obj As Object
//! Set obj = CreateObject("Scripting.FileSystemObject")
//!
//! ' Call the GetFolder method
//! Dim folder As Variant
//! folder = CallByName(obj, "GetFolder", VbMethod, "C:\Temp")
//! ```
//!
//! ### Property Get
//!
//! ```vb
//! Dim fs As Object
//! Set fs = CreateObject("Scripting.FileSystemObject")
//!
//! ' Get the Drives property
//! Dim drives As Variant
//! drives = CallByName(fs, "Drives", VbGet)
//! ```
//!
//! ### Property Let
//!
//! ```vb
//! Dim txt As Object
//! Set txt = CreateObject("Scripting.TextStream")
//!
//! ' Set a property value
//! CallByName txt, "Line", VbLet, 10
//! ```
//!
//! ### Property Set (Object Reference)
//!
//! ```vb
//! Dim form As Form
//! Set form = New Form1
//!
//! Dim btn As CommandButton
//! Set btn = New CommandButton
//!
//! ' Set an object property
//! CallByName form, "ActiveControl", VbSet, btn
//! ```
//!
//! ### Dynamic Property Access
//!
//! ```vb
//! Function GetPropertyValue(obj As Object, propName As String) As Variant
//!     GetPropertyValue = CallByName(obj, propName, VbGet)
//! End Function
//!
//! Function SetPropertyValue(obj As Object, propName As String, value As Variant)
//!     CallByName obj, propName, VbLet, value
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### 1. Generic Property Copier
//!
//! ```vb
//! Sub CopyProperties(source As Object, dest As Object, propNames() As String)
//!     Dim i As Integer
//!     Dim value As Variant
//!     
//!     For i = LBound(propNames) To UBound(propNames)
//!         value = CallByName(source, propNames(i), VbGet)
//!         CallByName dest, propNames(i), VbLet, value
//!     Next i
//! End Sub
//! ```
//!
//! ### 2. Form Field Population
//!
//! ```vb
//! Sub PopulateFormFromRecordset(frm As Form, rs As Recordset)
//!     Dim fld As Field
//!     Dim ctl As Control
//!     
//!     For Each fld In rs.Fields
//!         On Error Resume Next
//!         Set ctl = frm.Controls(fld.Name)
//!         If Not ctl Is Nothing Then
//!             CallByName ctl, "Text", VbLet, fld.Value & ""
//!         End If
//!         On Error GoTo 0
//!     Next fld
//! End Sub
//! ```
//!
//! ### 3. Dynamic Method Invocation
//!
//! ```vb
//! Function InvokeMethod(obj As Object, methodName As String, _
//!                      ParamArray args() As Variant) As Variant
//!     Dim argArray() As Variant
//!     Dim i As Integer
//!     
//!     If UBound(args) >= 0 Then
//!         ReDim argArray(LBound(args) To UBound(args))
//!         For i = LBound(args) To UBound(args)
//!             argArray(i) = args(i)
//!         Next i
//!         InvokeMethod = CallByName(obj, methodName, VbMethod, argArray)
//!     Else
//!         InvokeMethod = CallByName(obj, methodName, VbMethod)
//!     End If
//! End Function
//! ```
//!
//! ### 4. Property Name Validation
//!
//! ```vb
//! Function HasProperty(obj As Object, propName As String) As Boolean
//!     On Error Resume Next
//!     Dim temp As Variant
//!     temp = CallByName(obj, propName, VbGet)
//!     HasProperty = (Err.Number = 0)
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ### 5. Bulk Property Setting
//!
//! ```vb
//! Sub SetMultipleProperties(obj As Object, propNames As Variant, _
//!                          propValues As Variant)
//!     Dim i As Integer
//!     
//!     For i = LBound(propNames) To UBound(propNames)
//!         CallByName obj, propNames(i), VbLet, propValues(i)
//!     Next i
//! End Sub
//!
//! ' Usage:
//! SetMultipleProperties myControl, _
//!     Array("Left", "Top", "Width", "Height"), _
//!     Array(100, 100, 200, 50)
//! ```
//!
//! ### 6. Parameterized Property Access
//!
//! ```vb
//! ' Access a property with parameters (like an indexed property)
//! Sub SetIndexedProperty(obj As Object, propName As String, _
//!                       index As Integer, value As Variant)
//!     CallByName obj, propName, VbLet, index, value
//! End Sub
//!
//! Function GetIndexedProperty(obj As Object, propName As String, _
//!                            index As Integer) As Variant
//!     GetIndexedProperty = CallByName(obj, propName, VbGet, index)
//! End Function
//! ```
//!
//! ### 7. Configuration-Driven Object Initialization
//!
//! ```vb
//! Sub InitializeFromConfig(obj As Object, configFile As String)
//!     Dim fs As Object
//!     Dim ts As Object
//!     Dim line As String
//!     Dim parts() As String
//!     
//!     Set fs = CreateObject("Scripting.FileSystemObject")
//!     Set ts = fs.OpenTextFile(configFile)
//!     
//!     Do While Not ts.AtEndOfStream
//!         line = ts.ReadLine
//!         If InStr(line, "=") > 0 Then
//!             parts = Split(line, "=")
//!             CallByName obj, Trim(parts(0)), VbLet, Trim(parts(1))
//!         End If
//!     Loop
//!     
//!     ts.Close
//! End Sub
//! ```
//!
//! ### 8. Error-Safe Property Access
//!
//! ```vb
//! Function SafeGetProperty(obj As Object, propName As String, _
//!                         Optional defaultValue As Variant) As Variant
//!     On Error Resume Next
//!     SafeGetProperty = CallByName(obj, propName, VbGet)
//!     If Err.Number <> 0 Then
//!         If Not IsMissing(defaultValue) Then
//!             SafeGetProperty = defaultValue
//!         Else
//!             SafeGetProperty = Empty
//!         End If
//!     End If
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Calling Methods with Multiple Arguments
//!
//! ```vb
//! Dim obj As Object
//! Set obj = CreateObject("SomeLibrary.SomeClass")
//!
//! ' Call a method with multiple arguments
//! Dim result As Variant
//! result = CallByName(obj, "Calculate", VbMethod, 10, 20, "sum")
//! ```
//!
//! ### Working with Collections
//!
//! ```vb
//! Sub EnumerateCollection(coll As Collection, methodName As String)
//!     Dim item As Variant
//!     For Each item In coll
//!         CallByName item, methodName, VbMethod
//!     Next item
//! End Sub
//! ```
//!
//! ### Building a Simple ORM
//!
//! ```vb
//! Sub SaveObjectToDatabase(obj As Object, tableName As String, _
//!                         propNames() As String)
//!     Dim sql As String
//!     Dim values As String
//!     Dim i As Integer
//!     Dim value As Variant
//!     
//!     sql = "INSERT INTO " & tableName & " ("
//!     values = " VALUES ("
//!     
//!     For i = LBound(propNames) To UBound(propNames)
//!         If i > LBound(propNames) Then
//!             sql = sql & ", "
//!             values = values & ", "
//!         End If
//!         
//!         sql = sql & propNames(i)
//!         value = CallByName(obj, propNames(i), VbGet)
//!         values = values & "'" & value & "'"
//!     Next i
//!     
//!     sql = sql & ")" & values & ")"
//!     ' Execute SQL...
//! End Sub
//! ```
//!
//! ## Error Handling
//!
//! Common errors when using `CallByName`:
//!
//! - **Error 438**: Object doesn't support this property or method
//!   - The specified member doesn't exist
//!   - Check spelling and case (though `CallByName` is case-insensitive)
//!
//! - **Error 450**: Wrong number of arguments or invalid property assignment
//!   - Incorrect number of arguments in the args array
//!   - Using `VbLet` for an object (should use `VbSet`)
//!   - Using `VbSet` for a value type (should use `VbLet`)
//!
//! - **Error 13**: Type mismatch
//!   - Argument types don't match what the member expects
//!
//! ```vb
//! On Error Resume Next
//! result = CallByName(obj, "PropertyName", VbGet)
//! If Err.Number <> 0 Then
//!     MsgBox "Error accessing property: " & Err.Description
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Performance Considerations
//!
//! - `CallByName` is significantly slower than direct member access
//! - Name resolution happens at runtime, not compile time
//! - Consider caching frequently accessed members
//! - Use early binding and direct calls in performance-critical code
//! - `CallByName` is best for scenarios where dynamic access is necessary
//!
//! ## Limitations
//!
//! - Cannot call private members
//! - Cannot call Friend members from outside the project
//! - No `IntelliSense` support for the member being called
//! - No compile-time type checking
//! - Cannot call default members by passing empty string
//! - More difficult to debug than direct calls
//!
//! ## Related Functions
//!
//! - `Eval`: Evaluates an expression (only in VBA, not VB6)
//! - `Execute`: Executes a statement (only in VBA, not VB6)
//! - `GetObject`: Returns a reference to an object
//! - `CreateObject`: Creates an instance of an object
//! - `TypeName`: Returns the type name of a variable
//! - `VarType`: Returns the variant subtype of a variable
//!
//! ## Parsing Notes
//!
//! The `CallByName` function is not a reserved keyword in VB6. It is parsed as a regular
//! function call (`CallExpression`). This module exists primarily for documentation
//! purposes and to provide a comprehensive test suite that validates the parser
//! correctly handles `CallByName` function calls in various contexts.

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn callbyname_simple_method() {
        let source = r#"
Sub Test()
    result = CallByName(obj, "MethodName", VbMethod)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_with_vbget() {
        let source = r#"
Sub Test()
    value = CallByName(obj, "PropertyName", VbGet)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
        assert!(debug.contains("VbGet"));
    }

    #[test]
    fn callbyname_with_vblet() {
        let source = r#"
Sub Test()
    CallByName obj, "PropertyName", VbLet, 42
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
        assert!(debug.contains("VbLet"));
    }

    #[test]
    fn callbyname_with_vbset() {
        let source = r#"
Sub Test()
    CallByName form, "ActiveControl", VbSet, ctrl
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
        assert!(debug.contains("VbSet"));
    }

    #[test]
    fn callbyname_method_with_args() {
        let source = r#"
Sub Test()
    result = CallByName(obj, "Calculate", VbMethod, 10, 20, 30)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_in_loop() {
        let source = r#"
Sub Test()
    For i = 0 To 10
        CallByName obj, "Process", VbMethod, i
    Next i
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_with_string_literal() {
        let source = r#"
Sub Test()
    result = CallByName(obj, "GetFolder", VbMethod, "C:\Temp")
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_case_insensitive() {
        let source = r#"
Sub Test()
    a = CALLBYNAME(obj, "Method", VbMethod)
    b = callbyname(obj, "Method", VbMethod)
    c = CallByName(obj, "Method", VbMethod)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CALLBYNAME") || debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_in_if_statement() {
        let source = r#"
Sub Test()
    If CallByName(obj, "IsValid", VbMethod) Then
        Print "Valid"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    result = CallByName(obj, "PropertyName", VbGet)
    If Err.Number <> 0 Then
        MsgBox "Error"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_in_function() {
        let source = r#"
Function GetProperty(obj As Object, propName As String) As Variant
    GetProperty = CallByName(obj, propName, VbGet)
End Function
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_with_createobject() {
        let source = r#"
Sub Test()
    Set obj = CreateObject("Scripting.FileSystemObject")
    result = CallByName(obj, "GetFolder", VbMethod, "C:\")
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
        assert!(debug.contains("CreateObject"));
    }

    #[test]
    fn callbyname_multiple_in_sequence() {
        let source = r#"
Sub Test()
    value1 = CallByName(obj, "Prop1", VbGet)
    value2 = CallByName(obj, "Prop2", VbGet)
    CallByName obj, "Prop3", VbLet, value1 + value2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_in_with_block() {
        let source = r#"
Sub Test()
    With someObject
        result = CallByName(.Item, "Method", VbMethod)
    End With
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_indexed_property() {
        let source = r#"
Sub Test()
    value = CallByName(obj, "Item", VbGet, 5)
    CallByName obj, "Item", VbLet, 5, "NewValue"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_in_select_case() {
        let source = r#"
Sub Test()
    Select Case callType
        Case VbMethod
            CallByName obj, memberName, VbMethod
        Case VbGet
            result = CallByName(obj, memberName, VbGet)
    End Select
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_with_me() {
        let source = r#"
Sub Test()
    value = CallByName(Me, "Width", VbGet)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_with_concatenation() {
        let source = r#"
Sub Test()
    propName = "Get" & fieldName
    result = CallByName(obj, propName, VbMethod)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_nested_calls() {
        let source = r#"
Sub Test()
    result = CallByName(CallByName(obj, "SubObject", VbGet), "Method", VbMethod)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_with_array_access() {
        let source = r#"
Sub Test()
    For i = 0 To UBound(objects)
        CallByName objects(i), methodName, VbMethod
    Next i
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_with_numeric_constant() {
        let source = r#"
Sub Test()
    result = CallByName(obj, "Method", 1)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_in_do_loop() {
        let source = r#"
Sub Test()
    Do While Not rs.EOF
        CallByName rs, "MoveNext", VbMethod
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_with_whitespace() {
        let source = r#"
Sub Test()
    result = CallByName  (  obj  ,  "Method"  ,  VbMethod  )
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_with_line_continuation() {
        let source = r#"
Sub Test()
    result = CallByName _
        (obj, _
         "MethodName", _
         VbMethod, _
         arg1, arg2)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CallByName"));
    }

    #[test]
    fn callbyname_module_level() {
        let source = r#"Public result As Variant"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Public"));
    }
}
