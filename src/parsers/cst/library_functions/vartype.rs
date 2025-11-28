//! VB6 `VarType` Function
//!
//! The `VarType` function returns an integer constant indicating the subtype of a Variant variable or expression.
//!
//! ## Syntax
//! ```vb6
//! VarType(varname)
//! ```
//!
//! ## Parameters
//! - `varname`: Required. Name of a variable or expression whose Variant subtype is to be determined.
//!
//! ## Returns
//! Returns an `Integer` constant representing the Variant subtype. Common return values:
//! - 0: vbEmpty (uninitialized)
//! - 1: vbNull (Null)
//! - 2: vbInteger
//! - 3: vbLong
//! - 4: vbSingle
//! - 5: vbDouble
//! - 6: vbCurrency
//! - 7: vbDate
//! - 8: vbString
//! - 9: vbObject
//! - 10: vbError
//! - 11: vbBoolean
//! - 12: vbVariant (for arrays)
//! - 13: vbDataObject
//! - 17: vbByte
//! - 8192: vbArray (bitwise OR with base type)
//! - ...and others (see documentation)
//!
//! ## Remarks
//! - Returns a numeric constant, not a string.
//! - For arrays, returns vbArray (8192) bitwise OR'd with the base type (e.g., vbArray + vbInteger = 8194).
//! - For objects, returns vbObject (9) or vbDataObject (13).
//! - For user-defined types, returns vbUserDefinedType (36).
//! - For Empty, returns vbEmpty (0); for Null, returns vbNull (1).
//! - For non-Variant variables, returns the corresponding type constant.
//! - Useful for type checking, debugging, and generic code.
//! - Use with TypeName for string representation.
//!
//! ## Typical Uses
//! 1. Type checking in generic code
//! 2. Handling Variant variables
//! 3. Debugging and logging
//! 4. Validating function arguments
//! 5. Detecting arrays and base types
//! 6. Reflection-like operations
//! 7. Error handling and reporting
//! 8. Working with COM objects
//!
//! ## Basic Examples
//!
//! ### Example 1: Get VarType of Integer
//! ```vb6
//! Dim x As Integer
//! Debug.Print VarType(x) ' 2 (vbInteger)
//! ```
//!
//! ### Example 2: Get VarType of String
//! ```vb6
//! Dim s As String
//! Debug.Print VarType(s) ' 8 (vbString)
//! ```
//!
//! ### Example 3: Get VarType of Array
//! ```vb6
//! Dim arr(1 To 5) As Double
//! Debug.Print VarType(arr) ' 8205 (vbArray + vbDouble)
//! ```
//!
//! ### Example 4: Get VarType of Variant
//! ```vb6
//! Dim v As Variant
//! v = 123
//! Debug.Print VarType(v) ' 2 (vbInteger)
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Check for array
//! ```vb6
//! If VarType(var) And vbArray Then
//!     Debug.Print "It's an array!"
//! End If
//! ```
//!
//! ### Pattern 2: Check for string
//! ```vb6
//! If VarType(x) = vbString Then
//!     ' Handle string
//! End If
//! ```
//!
//! ### Pattern 3: Handle Variant types
//! ```vb6
//! If VarType(v) = vbInteger Then
//!     ' Handle integer
//! End If
//! ```
//!
//! ### Pattern 4: Log variable types
//! ```vb6
//! Debug.Print "VarType: " & VarType(x)
//! ```
//!
//! ### Pattern 5: Validate argument type
//! ```vb6
//! Sub Foo(arg As Variant)
//!     If VarType(arg) <> vbString Then Err.Raise 5
//! End Sub
//! ```
//!
//! ### Pattern 6: Reflection-like usage
//! ```vb6
//! Dim t As Integer
//! t = VarType(obj)
//! If t = vbObject Then
//!     ' Do something
//! End If
//! ```
//!
//! ### Pattern 7: Handle Null and Empty
//! ```vb6
//! If VarType(v) = vbNull Then
//!     ' Handle Null
//! ElseIf VarType(v) = vbEmpty Then
//!     ' Handle Empty
//! End If
//! ```
//!
//! ### Pattern 8: Array type detection
//! ```vb6
//! If (VarType(arr) And vbArray) Then
//!     Debug.Print "Array base type: " & (VarType(arr) - vbArray)
//! End If
//! ```
//!
//! ### Pattern 9: User-defined type
//! ```vb6
//! Type MyType
//!     x As Integer
//! End Type
//! Dim t As MyType
//! Debug.Print VarType(t) ' 36 (vbUserDefinedType)
//! ```
//!
//! ### Pattern 10: Class type detection
//! ```vb6
//! If VarType(obj) = vbObject Then
//!     ' Handle object
//! End If
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Type checking in generic function
//! ```vb6
//! Function IsString(val As Variant) As Boolean
//!     IsString = (VarType(val) = vbString)
//! End Function
//! ```
//!
//! ### Example 2: Logging all argument types
//! ```vb6
//! Sub LogTypes(ParamArray args() As Variant)
//!     Dim i As Integer
//!     For i = LBound(args) To UBound(args)
//!         Debug.Print "Arg " & i & ": " & VarType(args(i))
//!     Next i
//! End Sub
//! ```
//!
//! ### Example 3: Reflection for class methods
//! ```vb6
//! If VarType(obj) = vbObject Then
//!     obj.SpecialMethod
//! End If
//! ```
//!
//! ### Example 4: Variant array detection
//! ```vb6
//! Dim v As Variant
//! v = Array(1, 2, 3)
//! If (VarType(v) And vbArray) Then
//!     Debug.Print "Variant array"
//! End If
//! ```
//!
//! ## Error Handling
//! - Returns vbError (10) for error values.
//! - Returns vbUnknown (0) for unsupported types.
//! - Returns vbEmpty (0) for uninitialized variables.
//! - Returns vbNull (1) for Null values.
//!
//! ## Performance Notes
//! - Fast, constant time O(1).
//! - No side effects.
//!
//! ## Best Practices
//! 1. Use for debugging and logging.
//! 2. Use bitwise AND with vbArray to detect arrays.
//! 3. Use with TypeName for string representation.
//! 4. Handle vbNull, vbEmpty, and vbError cases.
//! 5. Use for generic code and utilities.
//! 6. Document expected type constants.
//! 7. Use for runtime checks, not compile-time.
//! 8. Combine with TypeName for more detail.
//! 9. Use for Variant and object variables.
//! 10. Avoid using as a substitute for type declarations.
//!
//! ## Comparison Table
//!
//! | Function   | Purpose                | Input      | Returns        |
//! |------------|------------------------|------------|----------------|
//! | `VarType`  | Get type as constant   | variable   | Integer        |
//! | `TypeName` | Get type as string     | variable   | String         |
//! | `IsObject` | Check if is object     | variable   | Boolean        |
//! | `IsArray`  | Check if is array      | variable   | Boolean        |
//!
//! ## Platform Notes
//! - Available in VB6, VBA, VBScript
//! - Consistent across platforms
//! - Returns type constants in English
//!
//! ## Limitations
//! - Returns only type constant as integer
//! - Not locale-sensitive
//! - Returns vbUnknown (0) for unsupported types
//! - Not for compile-time type checking
//! - May return user-defined type/class constants

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_vartype_integer() {
        let source = r#"
Sub Test()
    Dim x As Integer
    Debug.Print VarType(x)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_string() {
        let source = r#"
Sub Test()
    Dim s As String
    Debug.Print VarType(s)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_array() {
        let source = r#"
Sub Test()
    Dim arr(1 To 5) As Double
    Debug.Print VarType(arr)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_variant() {
        let source = r#"
Sub Test()
    Dim v As Variant
    v = 123
    Debug.Print VarType(v)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_check_array() {
        let source = r#"
Sub Test()
    If VarType(var) And vbArray Then
        Debug.Print "It's an array!"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_check_string() {
        let source = r#"
Sub Test()
    If VarType(x) = vbString Then
        ' Handle string
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_handle_variant() {
        let source = r#"
Sub Test()
    If VarType(v) = vbInteger Then
        ' Handle integer
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_log_type() {
        let source = r#"
Sub Test()
    Debug.Print "VarType: " & VarType(x)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_validate_argument() {
        let source = r#"
Sub Foo(arg As Variant)
    If VarType(arg) <> vbString Then Err.Raise 5
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_reflection() {
        let source = r#"
Sub Test()
    Dim t As Integer
    t = VarType(obj)
    If t = vbObject Then
        ' Do something
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_null_and_empty() {
        let source = r#"
Sub Test()
    If VarType(v) = vbNull Then
        ' Handle Null
    ElseIf VarType(v) = vbEmpty Then
        ' Handle Empty
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_array_type_detection() {
        let source = r#"
Sub Test()
    If (VarType(arr) And vbArray) Then
        Debug.Print "Array base type: " & (VarType(arr) - vbArray)
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_user_defined_type() {
        let source = r#"
Type MyType
    x As Integer
End Type
Sub Test()
    Dim t As MyType
    Debug.Print VarType(t)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_class_type_detection() {
        let source = r#"
Sub Test()
    If VarType(obj) = vbObject Then
        ' Handle object
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_isstring_function() {
        let source = r#"
Function IsString(val As Variant) As Boolean
    IsString = (VarType(val) = vbString)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_logtypes_paramarray() {
        let source = r#"
Sub LogTypes(ParamArray args() As Variant)
    Dim i As Integer
    For i = LBound(args) To UBound(args)
        Debug.Print "Arg " & i & ": " & VarType(args(i))
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_reflection_class_methods() {
        let source = r#"
Sub Test()
    If VarType(obj) = vbObject Then
        obj.SpecialMethod
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }

    #[test]
    fn test_vartype_variant_array_detection() {
        let source = r#"
Sub Test()
    Dim v As Variant
    v = Array(1, 2, 3)
    If (VarType(v) And vbArray) Then
        Debug.Print "Variant array"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("VarType"));
    }
}
