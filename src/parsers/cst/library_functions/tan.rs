//! VB6 `Tan` Function
//!
//! The `Tan` function returns the tangent of an angle specified in radians.
//!
//! ## Syntax
//! ```vb6
//! Tan(number)
//! ```
//!
//! ## Parameters
//! - `number`: Required. A numeric expression representing an angle in radians.
//!
//! ## Returns
//! Returns a `Double` representing the tangent of the angle.
//!
//! ## Remarks
//! - The argument must be in radians, not degrees. To convert degrees to radians, multiply by `Pi/180`.
//! - Returns a `Double` value.
//! - If the argument is a multiple of π/2 (except 0), the result is undefined (overflow error).
//! - Returns Null if the argument is Null.
//! - Use `Atn` to get the arctangent (inverse tangent).
//! - The tangent function is periodic with period π.
//! - For very large or very small arguments, floating-point rounding may affect results.
//!
//! ## Typical Uses
//! 1. Trigonometric calculations
//! 2. Geometry and graphics
//! 3. Physics and engineering formulas
//! 4. Calculating slopes and angles
//! 5. Animation and simulation
//! 6. Signal processing
//! 7. Scientific computation
//! 8. Converting between coordinate systems
//!
//! ## Basic Examples
//!
//! ### Example 1: Tangent of 45 degrees
//! ```vb6
//! result = Tan(45 * 3.14159265358979 / 180)
//! ' result = 1
//! ```
//!
//! ### Example 2: Tangent of Pi/4 radians
//! ```vb6
//! result = Tan(3.14159265358979 / 4)
//! ' result = 1
//! ```
//!
//! ### Example 3: Using with Atn
//! ```vb6
//! angle = Atn(1)
//! result = Tan(angle)
//! ' result = 1
//! ```
//!
//! ### Example 4: Handling Null
//! ```vb6
//! result = Tan(Null)
//! ' result = Null
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Convert degrees to radians
//! ```vb6
//! Function DegreesToRadians(degrees As Double) As Double
//!     DegreesToRadians = degrees * 3.14159265358979 / 180
//! End Function
//! result = Tan(DegreesToRadians(60))
//! ```
//!
//! ### Pattern 2: Calculate slope from angle
//! ```vb6
//! Function SlopeFromAngle(angleRadians As Double) As Double
//!     SlopeFromAngle = Tan(angleRadians)
//! End Function
//! ```
//!
//! ### Pattern 3: Use in triangle calculations
//! ```vb6
//! Function OppositeFromAdjacent(adjacent As Double, angleRadians As Double) As Double
//!     OppositeFromAdjacent = adjacent * Tan(angleRadians)
//! End Function
//! ```
//!
//! ### Pattern 4: Animation rotation
//! ```vb6
//! angle = t * 3.14159265358979 / 180
//! y = Tan(angle) * x
//! ```
//!
//! ### Pattern 5: Periodic function
//! ```vb6
//! For i = 0 To 360 Step 45
//!     Debug.Print Tan(i * 3.14159265358979 / 180)
//! Next i
//! ```
//!
//! ### Pattern 6: Error handling for undefined values
//! ```vb6
//! On Error Resume Next
//! result = Tan(3.14159265358979 / 2)
//! If Err.Number <> 0 Then
//!     Debug.Print "Overflow error"
//! End If
//! On Error GoTo 0
//! ```
//!
//! ### Pattern 7: Use with arrays
//! ```vb6
//! For i = LBound(arr) To UBound(arr)
//!     arr(i) = Tan(arr(i))
//! Next i
//! ```
//!
//! ### Pattern 8: Inverse calculation
//! ```vb6
//! angle = Atn(Tan(x))
//! ```
//!
//! ### Pattern 9: Normalize angle
//! ```vb6
//! angle = angle Mod (2 * 3.14159265358979)
//! result = Tan(angle)
//! ```
//!
//! ### Pattern 10: Use in coordinate conversion
//! ```vb6
//! y = r * Tan(theta)
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Trigonometric Table
//! ```vb6
//! For i = 0 To 90 Step 15
//!     Debug.Print "Tan(" & i & ") = " & Tan(i * 3.14159265358979 / 180)
//! Next i
//! ```
//!
//! ### Example 2: Slope Calculation
//! ```vb6
//! Function Slope(degrees As Double) As Double
//!     Slope = Tan(degrees * 3.14159265358979 / 180)
//! End Function
//! ```
//!
//! ### Example 3: Handling Undefined Values
//! ```vb6
//! On Error Resume Next
//! result = Tan(3.14159265358979 / 2)
//! If Err.Number <> 0 Then
//!     result = Null
//! End If
//! On Error GoTo 0
//! ```
//!
//! ### Example 4: Use in Physics Formula
//! ```vb6
//! ' Calculate projectile height
//! height = distance * Tan(angleRadians)
//! ```
//!
//! ## Error Handling
//! - Returns Null if argument is Null.
//! - Overflow error if argument is a multiple of π/2 (except 0).
//!
//! ## Performance Notes
//! - Fast, constant time O(1).
//! - Floating-point rounding may affect results for large/small arguments.
//!
//! ## Best Practices
//! 1. Always use radians, not degrees.
//! 2. Convert degrees to radians as needed.
//! 3. Handle possible overflow for undefined values.
//! 4. Use error handling for edge cases.
//! 5. Test with a range of values.
//! 6. Use with Atn for inverse calculations.
//! 7. Document expected input range.
//! 8. Avoid using with multiples of π/2.
//! 9. Use with arrays for batch calculations.
//! 10. Normalize angles for periodicity.
//!
//! ## Comparison Table
//!
//! | Function | Purpose | Input | Returns |
//! |----------|---------|-------|---------|
//! | `Tan`    | Tangent | radians | Double |
//! | `Atn`    | Arctangent | number | Double |
//! | `Sin`    | Sine | radians | Double |
//! | `Cos`    | Cosine | radians | Double |
//!
//! ## Platform Notes
//! - Available in VB6, VBA, `VBScript`
//! - Consistent across platforms
//! - Returns Double
//!
//! ## Limitations
//! - Argument must be in radians
//! - Undefined for odd multiples of π/2 (except 0)
//! - Returns Null for Null input
//! - No support for complex numbers
//! - Floating-point rounding errors possible

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn tan_basic() {
        let source = r#"
Sub Test()
    result = Tan(0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_45_degrees() {
        let source = r#"
Sub Test()
    result = Tan(45 * 3.14159265358979 / 180)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_pi_over_4() {
        let source = r#"
Sub Test()
    result = Tan(3.14159265358979 / 4)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_with_atn() {
        let source = r#"
Sub Test()
    angle = Atn(1)
    result = Tan(angle)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_null() {
        let source = r#"
Sub Test()
    result = Tan(Null)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_degrees_to_radians() {
        let source = r#"
Function DegreesToRadians(degrees As Double) As Double
    DegreesToRadians = degrees * 3.14159265358979 / 180
End Function
Sub Test()
    result = Tan(DegreesToRadians(60))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_slope_from_angle() {
        let source = r#"
Function SlopeFromAngle(angleRadians As Double) As Double
    SlopeFromAngle = Tan(angleRadians)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_triangle_calculation() {
        let source = r#"
Function OppositeFromAdjacent(adjacent As Double, angleRadians As Double) As Double
    OppositeFromAdjacent = adjacent * Tan(angleRadians)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_animation_rotation() {
        let source = r#"
Sub Animate()
    angle = t * 3.14159265358979 / 180
    y = Tan(angle) * x
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_periodic_function() {
        let source = r#"
Sub Test()
    For i = 0 To 360 Step 45
        Debug.Print Tan(i * 3.14159265358979 / 180)
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    result = Tan(3.14159265358979 / 2)
    If Err.Number <> 0 Then
        Debug.Print "Overflow error"
    End If
    On Error GoTo 0
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_array_usage() {
        let source = r#"
Sub Test()
    For i = LBound(arr) To UBound(arr)
        arr(i) = Tan(arr(i))
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_inverse_calculation() {
        let source = r#"
Sub Test()
    angle = Atn(Tan(x))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_normalize_angle() {
        let source = r#"
Sub Test()
    angle = angle Mod (2 * 3.14159265358979)
    result = Tan(angle)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_coordinate_conversion() {
        let source = r#"
Sub Test()
    y = r * Tan(theta)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_trig_table() {
        let source = r#"
Sub Test()
    For i = 0 To 90 Step 15
        Debug.Print "Tan(" & i & ") = " & Tan(i * 3.14159265358979 / 180)
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_slope_function() {
        let source = r#"
Function Slope(degrees As Double) As Double
    Slope = Tan(degrees * 3.14159265358979 / 180)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_undefined_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    result = Tan(3.14159265358979 / 2)
    If Err.Number <> 0 Then
        result = Null
    End If
    On Error GoTo 0
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }

    #[test]
    fn tan_physics_formula() {
        let source = r#"
Sub Test()
    height = distance * Tan(angleRadians)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Tan"));
    }
}
