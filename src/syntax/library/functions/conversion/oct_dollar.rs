//! # `Oct$` Function
//!
//! The `Oct$` function in Visual Basic 6 returns a string representing the octal (base-8) value
//! of a number. The function name stands for "Octal String".
//!
//! ## Syntax
//!
//! ```vb6
//! Oct$(number)
//! ```
//!
//! ## Parameters
//!
//! - `number` - Required. Any valid numeric expression. If `number` is not a whole number, it is
//!   rounded to the nearest whole number before being evaluated.
//!
//! ## Return Value
//!
//! Returns a `String` representing the octal value of the number. The returned string contains
//! only the digits 0-7, without a leading "0" or "&O" prefix.
//!
//! ## Behavior and Characteristics
//!
//! ### Data Type Handling
//!
//! - Accepts any numeric type: `Byte`, `Integer`, `Long`, `Single`, `Double`, `Currency`
//! - Floating-point values are rounded to the nearest integer before conversion
//! - Negative numbers are represented using two's complement notation
//! - Returns unsigned octal representation for the underlying bit pattern
//!
//! ### Range Considerations
//!
//! - `Integer` values: Returns 1-6 octal digits (range: 0 to 177777 for positive, 100000-177777 for negative)
//! - `Long` values: Returns 1-11 octal digits (range: 0 to 17777777777 for positive)
//! - `Byte` values: Returns 1-3 octal digits (range: 0 to 377)
//!
//! ## Common Usage Patterns
//!
//! ### 1. Basic Octal Conversion
//!
//! ```vb6
//! Dim octStr As String
//! octStr = Oct$(64)  ' Returns "100"
//! octStr = Oct$(8)   ' Returns "10"
//! octStr = Oct$(511) ' Returns "777"
//! ```
//!
//! ### 2. Converting Negative Numbers
//!
//! ```vb6
//! Dim octStr As String
//! octStr = Oct$(-1)  ' Returns "177777" (Integer range, two's complement)
//! ```
//!
//! ### 3. File Permission Representation
//!
//! ```vb6
//! Function FormatPermissions(permissions As Integer) As String
//!     ' Unix-style file permissions (e.g., 755, 644)
//!     FormatPermissions = Oct$(permissions)
//! End Function
//!
//! Dim perms As String
//! perms = FormatPermissions(&H1ED)  ' Returns "755"
//! ```
//!
//! ### 4. Bit Mask Display
//!
//! ```vb6
//! Dim flags As Integer
//! Dim octDisplay As String
//! flags = &H1FF
//! octDisplay = "Flags: " & Oct$(flags)  ' "Flags: 777"
//! ```
//!
//! ### 5. Color Component Extraction (Octal)
//!
//! ```vb6
//! Dim colorValue As Long
//! Dim component As Integer
//! colorValue = &HFF8040
//! component = (colorValue And &HFF)
//! Debug.Print Oct$(component)  ' Shows octal representation
//! ```
//!
//! ### 6. Data Structure Field Values
//!
//! ```vb6
//! Type SystemFlags
//!     ReadWrite As Integer
//!     Execute As Integer
//! End Type
//!
//! Dim sysFlags As SystemFlags
//! sysFlags.ReadWrite = &O644  ' Octal literal
//! Debug.Print "RW: " & Oct$(sysFlags.ReadWrite)
//! ```
//!
//! ### 7. Debugging Bit Patterns
//!
//! ```vb6
//! Sub ShowBitPattern(value As Integer)
//!     Debug.Print "Decimal: " & value
//!     Debug.Print "Octal: " & Oct$(value)
//!     Debug.Print "Hex: " & Hex$(value)
//! End Sub
//! ```
//!
//! ### 8. Network Protocol Values
//!
//! ```vb6
//! Dim socketMode As Integer
//! socketMode = &O666  ' Read/write for all
//! Debug.Print "Mode: " & Oct$(socketMode)
//! ```
//!
//! ### 9. Conversion Table Generation
//!
//! ```vb6
//! Sub GenerateOctalTable()
//!     Dim i As Integer
//!     For i = 0 To 64
//!         Debug.Print i & " = " & Oct$(i)
//!     Next i
//! End Sub
//! ```
//!
//! ### 10. Configuration Value Formatting
//!
//! ```vb6
//! Function SaveConfigValue(value As Integer) As String
//!     ' Store configuration as octal string
//!     SaveConfigValue = "CONFIG=" & Oct$(value)
//! End Function
//! ```
//!
//! ## Related Functions
//!
//! - `Hex$()` - Converts a number to hexadecimal (base-16) string representation
//! - `Str$()` - Converts a number to decimal string representation
//! - `Val()` - Converts a string to a numeric value (doesn't parse octal)
//! - `CLng()` - Converts an expression to a `Long` integer
//! - `CInt()` - Converts an expression to an `Integer`
//! - `Format$()` - Provides custom number formatting options
//!
//! ## Best Practices
//!
//! ### When to Use `Oct$`
//!
//! 1. **Unix-style Permissions**: Representing file or directory permissions (e.g., 755, 644)
//! 2. **Bit Pattern Analysis**: When examining data in groups of 3 bits
//! 3. **Legacy System Integration**: Working with systems that use octal notation
//! 4. **Debugging**: Displaying bit patterns in a more compact form than binary
//! 5. **Configuration Files**: Storing numeric values in octal format
//!
//! ### Formatting Output
//!
//! ```vb6
//! ' Add prefix for clarity
//! Debug.Print "Octal: &O" & Oct$(value)
//!
//! ' Pad with leading zeros
//! Debug.Print Right$("000" & Oct$(value), 3)
//! ```
//!
//! ### Type Safety
//!
//! ```vb6
//! ' Explicitly convert to ensure correct range
//! Dim longValue As Long
//! longValue = 1000000
//! Debug.Print Oct$(longValue)  ' Uses Long range
//! ```
//!
//! ## Performance Considerations
//!
//! - `Oct$` is a lightweight function with minimal overhead
//! - String concatenation in loops should use a `String` buffer or array for better performance
//! - For frequent conversions, consider caching results if the same values are converted repeatedly
//!
//! ## Octal Literals in VB6
//!
//! VB6 supports octal literals using the `&O` prefix:
//!
//! ```vb6
//! Dim octValue As Integer
//! octValue = &O777  ' Octal literal (equals 511 decimal)
//! Debug.Print Oct$(octValue)  ' Returns "777"
//! ```
//!
//! ## Common Pitfalls
//!
//! ### 1. No Direct Reverse Function
//!
//! VB6's `Val()` function does not parse octal strings. You need a custom function:
//!
//! ```vb6
//! Function OctVal(octStr As String) As Long
//!     Dim i As Integer
//!     Dim result As Long
//!     For i = 1 To Len(octStr)
//!         result = result * 8 + Val(Mid$(octStr, i, 1))
//!     Next i
//!     OctVal = result
//! End Function
//! ```
//!
//! ### 2. Two's Complement Representation
//!
//! Negative numbers produce two's complement octal strings:
//!
//! ```vb6
//! Debug.Print Oct$(-1)   ' "177777" (for Integer)
//! Debug.Print Oct$(-100) ' Not intuitive without understanding two's complement
//! ```
//!
//! ### 3. Floating-Point Rounding
//!
//! ```vb6
//! Debug.Print Oct$(8.5)  ' "10" (rounds to 8)
//! Debug.Print Oct$(8.6)  ' "11" (rounds to 9)
//! ```
//!
//! ### 4. Leading Zeros Not Included
//!
//! ```vb6
//! Debug.Print Oct$(8)  ' "10", not "010"
//! ' Pad manually if needed
//! Debug.Print Right$("000" & Oct$(8), 3)  ' "010"
//! ```
//!
//! ### 5. No Prefix in Output
//!
//! Unlike some languages, VB6's `Oct$` doesn't include the `&O` prefix:
//!
//! ```vb6
//! Debug.Print Oct$(64)  ' "100", not "&O100"
//! ```
//!
//! ## Limitations
//!
//! - No built-in function to convert octal strings back to numbers (must implement manually)
//! - Cannot specify minimum width or padding (must format manually)
//! - Limited usefulness in modern applications (hexadecimal is more common)
//! - No validation that a string contains valid octal digits
//! - Returns unsigned representation for negative numbers (two's complement)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn oct_dollar_simple() {
        let source = r"
Sub Main()
    result = Oct$(64)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_assignment() {
        let source = r"
Sub Main()
    Dim octStr As String
    octStr = Oct$(511)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_variable() {
        let source = r"
Sub Main()
    Dim num As Integer
    Dim octStr As String
    num = 255
    octStr = Oct$(num)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_negative() {
        let source = r"
Sub Main()
    Dim result As String
    result = Oct$(-1)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_permissions() {
        let source = r"
Function FormatPermissions(permissions As Integer) As String
    FormatPermissions = Oct$(permissions)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_in_condition() {
        let source = r#"
Sub Main()
    If Oct$(value) = "100" Then
        Debug.Print "Is 64"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_bit_mask() {
        let source = r#"
Sub Main()
    Dim flags As Integer
    Dim display As String
    flags = &H1FF
    display = "Flags: " & Oct$(flags)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_loop_table() {
        let source = r#"
Sub GenerateTable()
    Dim i As Integer
    For i = 0 To 64
        Debug.Print i & " = " & Oct$(i)
    Next i
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_in_function() {
        let source = r"
Function GetOctalString(value As Long) As String
    GetOctalString = Oct$(value)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_multiple_calls() {
        let source = r"
Sub ShowConversions()
    Debug.Print Oct$(8)
    Debug.Print Oct$(64)
    Debug.Print Oct$(511)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_select_case() {
        let source = r#"
Sub Main()
    Select Case Oct$(value)
        Case "10"
            Debug.Print "Eight"
        Case "100"
            Debug.Print "Sixty-four"
    End Select
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_color_component() {
        let source = r"
Sub ExtractComponent()
    Dim colorValue As Long
    Dim component As Integer
    colorValue = &HFF8040
    component = (colorValue And &HFF)
    Debug.Print Oct$(component)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_expression_arg() {
        let source = r"
Sub Main()
    Dim result As String
    result = Oct$(value * 8 + offset)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_concatenation() {
        let source = r#"
Sub Main()
    Dim output As String
    output = "Octal: &O" & Oct$(value)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_padded_output() {
        let source = r#"
Sub Main()
    Dim padded As String
    padded = Right$("000" & Oct$(value), 3)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_debug_pattern() {
        let source = r#"
Sub ShowBitPattern(value As Integer)
    Debug.Print "Decimal: " & value
    Debug.Print "Octal: " & Oct$(value)
    Debug.Print "Hex: " & Hex$(value)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_config_value() {
        let source = r#"
Function SaveConfig(value As Integer) As String
    SaveConfig = "CONFIG=" & Oct$(value)
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_floating_point() {
        let source = r"
Sub Main()
    Dim result As String
    result = Oct$(8.6)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_octal_literal() {
        let source = r"
Sub Main()
    Dim octValue As Integer
    Dim result As String
    octValue = &O777
    result = Oct$(octValue)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn oct_dollar_socket_mode() {
        let source = r#"
Sub Main()
    Dim socketMode As Integer
    socketMode = &O666
    Debug.Print "Mode: " & Oct$(socketMode)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/conversion/oct_dollar",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
