//! VB6 `TypeName` Function
//!
//! The `TypeName` function returns a string that provides information about the data type of a variable or expression.
//!
//! ## Syntax
//! ```vb6
//! TypeName(varname)
//! ```
//!
//! ## Parameters
//! - `varname`: Required. Name of a variable or expression whose data type is to be determined.
//!
//! ## Returns
//! Returns a `String` describing the data type of the variable or expression. Common return values include:
//! - "Boolean"
//! - "Byte"
//! - "Integer"
//! - "Long"
//! - "Single"
//! - "Double"
//! - "Currency"
//! - "Date"
//! - "String"
//! - "Object"
//! - "Error"
//! - "Empty"
//! - "Null"
//! - "Nothing"
//! - "Variant"
//! - "Unknown"
//! - Custom class or user-defined type names
//!
//! ## Remarks
//! - Returns the type as a string, not the actual type.
//! - For objects, returns the class name or interface name.
//! - For arrays, returns the base type name with "()" appended (e.g., "`Integer()`", "`String()`").
//! - For objects not instantiated, returns "Nothing".
//! - For Null, returns "Null"; for Empty, returns "Empty".
//! - For user-defined types, returns the type name.
//! - For Variant variables, returns the underlying type.
//! - Useful for debugging, logging, and type checking at runtime.
//! - Not case-sensitive.
//!
//! ## Typical Uses
//! 1. Debugging variable types
//! 2. Logging type information
//! 3. Type checking in generic code
//! 4. Handling Variant variables
//! 5. Validating function arguments
//! 6. Reflection-like operations
//! 7. Error handling and reporting
//! 8. Determining array types
//!
//! ## Basic Examples
//!
//! ### Example 1: Get type of variable
//! ```vb6
//! Dim x As Integer
//! MsgBox TypeName(x) ' "Integer"
//! ```
//!
//! ### Example 2: Get type of object
//! ```vb6
//! Dim c As Collection
//! Set c = New Collection
//! MsgBox TypeName(c) ' "Collection"
//! ```
//!
//! ### Example 3: Get type of array
//! ```vb6
//! Dim arr(1 To 5) As String
//! MsgBox TypeName(arr) ' "String()"
//! ```
//!
//! ### Example 4: Get type of Variant
//! ```vb6
//! Dim v As Variant
//! v = 123
//! MsgBox TypeName(v) ' "Integer"
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Check for array
//! ```vb6
//! If Right$(TypeName(var), 2) = "()" Then
//!     MsgBox "It's an array!"
//! End If
//! ```
//!
//! ### Pattern 2: Check for object
//! ```vb6
//! If TypeName(obj) = "Nothing" Then
//!     MsgBox "Object not set!"
//! End If
//! ```
//!
//! ### Pattern 3: Handle Variant types
//! ```vb6
//! If TypeName(v) = "String" Then
//!     ' Handle string
//! End If
//! ```
//!
//! ### Pattern 4: Log variable types
//! ```vb6
//! Debug.Print "Type: " & TypeName(x)
//! ```
//!
//! ### Pattern 5: Validate argument type
//! ```vb6
//! Sub Foo(arg As Variant)
//!     If TypeName(arg) <> "String" Then Err.Raise 5
//! End Sub
//! ```
//!
//! ### Pattern 6: Reflection-like usage
//! ```vb6
//! Dim t As String
//! t = TypeName(obj)
//! If t = "MyClass" Then
//!     ' Do something
//! End If
//! ```
//!
//! ### Pattern 7: Handle Null and Empty
//! ```vb6
//! If TypeName(v) = "Null" Then
//!     ' Handle Null
//! ElseIf TypeName(v) = "Empty" Then
//!     ' Handle Empty
//! End If
//! ```
//!
//! ### Pattern 8: Array type detection
//! ```vb6
//! If InStr(TypeName(arr), "()") > 0 Then
//!     Debug.Print "Array of type: " & Left$(TypeName(arr), Len(TypeName(arr)) - 2)
//! End If
//! ```
//!
//! ### Pattern 9: User-defined type
//! ```vb6
//! Type MyType
//!     x As Integer
//! End Type
//! Dim t As MyType
//! MsgBox TypeName(t) ' "MyType"
//! ```
//!
//! ### Pattern 10: Class type detection
//! ```vb6
//! If TypeName(obj) = "MyClass" Then
//!     ' Handle MyClass
//! End If
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Type checking in generic function
//! ```vb6
//! Function IsString(val As Variant) As Boolean
//!     IsString = (TypeName(val) = "String")
//! End Function
//! ```
//!
//! ### Example 2: Logging all argument types
//! ```vb6
//! Sub LogTypes(ParamArray args() As Variant)
//!     Dim i As Integer
//!     For i = LBound(args) To UBound(args)
//!         Debug.Print "Arg " & i & ": " & TypeName(args(i))
//!     Next i
//! End Sub
//! ```
//!
//! ### Example 3: Reflection for class methods
//! ```vb6
//! If TypeName(obj) = "MyClass" Then
//!     obj.SpecialMethod
//! End If
//! ```
//!
//! ### Example 4: Variant array detection
//! ```vb6
//! Dim v As Variant
//! v = Array(1, 2, 3)
//! If Right$(TypeName(v), 2) = "()" Then
//!     Debug.Print "Variant array"
//! End If
//! ```
//!
//! ## Error Handling
//! - Returns "Unknown" for unsupported types.
//! - Returns "Nothing" for uninitialized object variables.
//! - Returns "Null" for Null values.
//! - Returns "Empty" for uninitialized variables.
//!
//! ## Performance Notes
//! - Fast, constant time O(1).
//! - No side effects.
//!
//! ## Best Practices
//! 1. Use for debugging and logging.
//! 2. Do not use for strict type enforcement.
//! 3. Handle "Nothing", "Null", and "Empty" cases.
//! 4. Use with Variant variables for type safety.
//! 5. Use for generic code and utilities.
//! 6. Document expected type strings.
//! 7. Use with arrays for type detection.
//! 8. Avoid using as a substitute for type declarations.
//! 9. Use for runtime checks, not compile-time.
//! 10. Combine with `VarType` for more detail.
//!
//! ## Comparison Table
//!
//! | Function   | Purpose                | Input      | Returns        |
//! |------------|------------------------|------------|----------------|
//! | `TypeName` | Get type as string     | variable   | String         |
//! | `VarType`  | Get type as constant   | variable   | Integer        |
//! | `IsObject` | Check if is object     | variable   | Boolean        |
//! | `IsArray`  | Check if is array      | variable   | Boolean        |
//!
//! ## Platform Notes
//! - Available in VB6, VBA, `VBScript`
//! - Consistent across platforms
//! - Returns type names in English
//!
//! ## Limitations
//! - Returns only type name as string
//! - Not locale-sensitive
//! - Returns "Unknown" for unsupported types
//! - Not for compile-time type checking
//! - May return user-defined type/class names

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn typename_integer() {
        let source = r"
Sub Test()
    Dim x As Integer
    MsgBox TypeName(x)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_object() {
        let source = r"
Sub Test()
    Dim c As Collection
    Set c = New Collection
    MsgBox TypeName(c)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_array() {
        let source = r"
Sub Test()
    Dim arr(1 To 5) As String
    MsgBox TypeName(arr)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_variant() {
        let source = r"
Sub Test()
    Dim v As Variant
    v = 123
    MsgBox TypeName(v)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_check_array() {
        let source = r#"
Sub Test()
    If Right$(TypeName(var), 2) = "()" Then
        MsgBox "It's an array!"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_check_object() {
        let source = r#"
Sub Test()
    If TypeName(obj) = "Nothing" Then
        MsgBox "Object not set!"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_handle_variant() {
        let source = r#"
Sub Test()
    If TypeName(v) = "String" Then
        ' Handle string
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_log_type() {
        let source = r#"
Sub Test()
    Debug.Print "Type: " & TypeName(x)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_validate_argument() {
        let source = r#"
Sub Foo(arg As Variant)
    If TypeName(arg) <> "String" Then Err.Raise 5
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_reflection() {
        let source = r#"
Sub Test()
    Dim t As String
    t = TypeName(obj)
    If t = "MyClass" Then
        ' Do something
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_null_and_empty() {
        let source = r#"
Sub Test()
    If TypeName(v) = "Null" Then
        ' Handle Null
    ElseIf TypeName(v) = "Empty" Then
        ' Handle Empty
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_array_type_detection() {
        let source = r#"
Sub Test()
    If InStr(TypeName(arr), "()") > 0 Then
        Debug.Print "Array of type: " & Left$(TypeName(arr), Len(TypeName(arr)) - 2)
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_user_defined_type() {
        let source = r"
Type MyType
    x As Integer
End Type
Sub Test()
    Dim t As MyType
    MsgBox TypeName(t)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_class_type_detection() {
        let source = r#"
Sub Test()
    If TypeName(obj) = "MyClass" Then
        ' Handle MyClass
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_isstring_function() {
        let source = r#"
Function IsString(val As Variant) As Boolean
    IsString = (TypeName(val) = "String")
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_logtypes_paramarray() {
        let source = r#"
Sub LogTypes(ParamArray args() As Variant)
    Dim i As Integer
    For i = LBound(args) To UBound(args)
        Debug.Print "Arg " & i & ": " & TypeName(args(i))
    Next i
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_reflection_class_methods() {
        let source = r#"
Sub Test()
    If TypeName(obj) = "MyClass" Then
        obj.SpecialMethod
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn typename_variant_array_detection() {
        let source = r#"
Sub Test()
    Dim v As Variant
    v = Array(1, 2, 3)
    If Right$(TypeName(v), 2) = "()" Then
        Debug.Print "Variant array"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/objects/typename",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
