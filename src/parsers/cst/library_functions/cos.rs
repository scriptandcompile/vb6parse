//! # Cos Function
//!
//! Returns a Double specifying the cosine of an angle.
//!
//! ## Syntax
//!
//! ```vb
//! Cos(number)
//! ```
//!
//! ## Parameters
//!
//! - **number**: Required. Double or any valid numeric expression that expresses an angle in radians.
//!
//! ## Return Value
//!
//! Returns a Double representing the cosine of the angle. The return value ranges from -1 to 1.
//!
//! ## Remarks
//!
//! The `Cos` function takes an angle in radians and returns the ratio of the length of the
//! adjacent side to the length of the hypotenuse in a right triangle. This is a fundamental
//! trigonometric function used in mathematical calculations, graphics programming, physics
//! simulations, and engineering applications.
//!
//! **Important Characteristics:**
//!
//! - Argument is in radians, not degrees
//! - Return value range: -1 to 1
//! - Cos(0) = 1
//! - Cos(π/2) ≈ 0
//! - Cos(π) = -1
//! - Cos(3π/2) ≈ 0
//! - Cos(2π) = 1
//! - Periodic with period 2π
//! - Even function: Cos(-x) = Cos(x)
//!
//! ## Angle Conversion
//!
//! To convert degrees to radians (required for Cos):
//! ```vb
//! radians = degrees * (π / 180)
//! radians = degrees * 0.0174532925199433
//! ```
//!
//! To convert radians to degrees:
//! ```vb
//! degrees = radians * (180 / π)
//! degrees = radians * 57.2957795130823
//! ```
//!
//! ## Common Angle Values
//!
//! | Degrees | Radians | Cos(angle) |
//! |---------|---------|------------|
//! | 0° | 0 | 1 |
//! | 30° | π/6 | √3/2 ≈ 0.866 |
//! | 45° | π/4 | √2/2 ≈ 0.707 |
//! | 60° | π/3 | 0.5 |
//! | 90° | π/2 | 0 |
//! | 120° | 2π/3 | -0.5 |
//! | 135° | 3π/4 | -√2/2 ≈ -0.707 |
//! | 150° | 5π/6 | -√3/2 ≈ -0.866 |
//! | 180° | π | -1 |
//! | 270° | 3π/2 | 0 |
//! | 360° | 2π | 1 |
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Calculate cosine of an angle in radians
//! Dim result As Double
//! result = Cos(0)        ' Returns 1
//! result = Cos(1.5708)   ' Returns approximately 0 (π/2)
//! result = Cos(3.14159)  ' Returns approximately -1 (π)
//!
//! ' Using with degrees (convert first)
//! Dim angleInDegrees As Double
//! Dim angleInRadians As Double
//! angleInDegrees = 45
//! angleInRadians = angleInDegrees * (3.14159265358979 / 180)
//! result = Cos(angleInRadians)  ' Returns approximately 0.707
//! ```
//!
//! ### Using Pi Constant
//!
//! ```vb
//! Const Pi As Double = 3.14159265358979
//!
//! Dim angle As Double
//! angle = Cos(Pi / 4)     ' 45 degrees, returns ≈0.707
//! angle = Cos(Pi / 2)     ' 90 degrees, returns ≈0
//! angle = Cos(Pi)         ' 180 degrees, returns -1
//! angle = Cos(2 * Pi)     ' 360 degrees, returns 1
//! ```
//!
//! ### Degrees to Radians Conversion
//!
//! ```vb
//! Function DegreesToRadians(degrees As Double) As Double
//!     Const Pi As Double = 3.14159265358979
//!     DegreesToRadians = degrees * (Pi / 180)
//! End Function
//!
//! Function CosDegrees(degrees As Double) As Double
//!     CosDegrees = Cos(DegreesToRadians(degrees))
//! End Function
//!
//! ' Usage
//! Dim result As Double
//! result = CosDegrees(60)  ' Returns 0.5
//! ```
//!
//! ## Common Patterns
//!
//! ### Circle Point Calculation
//!
//! ```vb
//! Function GetCirclePoint(centerX As Double, centerY As Double, _
//!                         radius As Double, angleDegrees As Double) As Point
//!     Const Pi As Double = 3.14159265358979
//!     Dim angleRad As Double
//!     Dim pt As Point
//!     
//!     angleRad = angleDegrees * (Pi / 180)
//!     
//!     pt.X = centerX + radius * Cos(angleRad)
//!     pt.Y = centerY + radius * Sin(angleRad)
//!     
//!     GetCirclePoint = pt
//! End Function
//! ```
//!
//! ### Rotating Points
//!
//! ```vb
//! Function RotatePoint(x As Double, y As Double, angleDegrees As Double) As Point
//!     Const Pi As Double = 3.14159265358979
//!     Dim angleRad As Double
//!     Dim pt As Point
//!     
//!     angleRad = angleDegrees * (Pi / 180)
//!     
//!     pt.X = x * Cos(angleRad) - y * Sin(angleRad)
//!     pt.Y = x * Sin(angleRad) + y * Cos(angleRad)
//!     
//!     RotatePoint = pt
//! End Function
//! ```
//!
//! ### Wave Generation
//!
//! ```vb
//! Function GenerateCosineWave(samples As Integer, amplitude As Double, _
//!                             frequency As Double) As Double()
//!     Const Pi As Double = 3.14159265358979
//!     Dim wave() As Double
//!     Dim i As Integer
//!     Dim angle As Double
//!     
//!     ReDim wave(0 To samples - 1)
//!     
//!     For i = 0 To samples - 1
//!         angle = 2 * Pi * frequency * i / samples
//!         wave(i) = amplitude * Cos(angle)
//!     Next i
//!     
//!     GenerateCosineWave = wave
//! End Function
//! ```
//!
//! ### Harmonic Motion
//!
//! ```vb
//! Function SimpleHarmonicMotion(amplitude As Double, frequency As Double, _
//!                               time As Double) As Double
//!     Const Pi As Double = 3.14159265358979
//!     SimpleHarmonicMotion = amplitude * Cos(2 * Pi * frequency * time)
//! End Function
//! ```
//!
//! ### Distance Calculation
//!
//! ```vb
//! Function DistanceToLine(px As Double, py As Double, _
//!                         angle As Double, distance As Double) As Point
//!     Const Pi As Double = 3.14159265358979
//!     Dim pt As Point
//!     Dim angleRad As Double
//!     
//!     angleRad = angle * (Pi / 180)
//!     
//!     pt.X = px + distance * Cos(angleRad)
//!     pt.Y = py + distance * Sin(angleRad)
//!     
//!     DistanceToLine = pt
//! End Function
//! ```
//!
//! ### Ellipse Points
//!
//! ```vb
//! Function GetEllipsePoint(centerX As Double, centerY As Double, _
//!                          radiusX As Double, radiusY As Double, _
//!                          angleDegrees As Double) As Point
//!     Const Pi As Double = 3.14159265358979
//!     Dim angleRad As Double
//!     Dim pt As Point
//!     
//!     angleRad = angleDegrees * (Pi / 180)
//!     
//!     pt.X = centerX + radiusX * Cos(angleRad)
//!     pt.Y = centerY + radiusY * Sin(angleRad)
//!     
//!     GetEllipsePoint = pt
//! End Function
//! ```
//!
//! ### Polar to Cartesian Conversion
//!
//! ```vb
//! Function PolarToCartesian(radius As Double, angleDegrees As Double) As Point
//!     Const Pi As Double = 3.14159265358979
//!     Dim angleRad As Double
//!     Dim pt As Point
//!     
//!     angleRad = angleDegrees * (Pi / 180)
//!     
//!     pt.X = radius * Cos(angleRad)
//!     pt.Y = radius * Sin(angleRad)
//!     
//!     PolarToCartesian = pt
//! End Function
//! ```
//!
//! ### Clock Hand Position
//!
//! ```vb
//! Function GetClockHandPosition(centerX As Double, centerY As Double, _
//!                               handLength As Double, hours As Integer, _
//!                               minutes As Integer) As Point
//!     Const Pi As Double = 3.14159265358979
//!     Dim angle As Double
//!     Dim angleRad As Double
//!     Dim pt As Point
//!     
//!     ' Calculate angle (12 o'clock = 0 degrees, clockwise)
//!     angle = (hours Mod 12) * 30 + minutes * 0.5  ' 30 degrees per hour
//!     angle = angle - 90  ' Adjust so 0 degrees is at 3 o'clock position
//!     
//!     angleRad = angle * (Pi / 180)
//!     
//!     pt.X = centerX + handLength * Cos(angleRad)
//!     pt.Y = centerY + handLength * Sin(angleRad)
//!     
//!     GetClockHandPosition = pt
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### 3D Rotation (Yaw)
//!
//! ```vb
//! Function RotateYaw(x As Double, y As Double, z As Double, _
//!                    angleDegrees As Double) As Point3D
//!     Const Pi As Double = 3.14159265358979
//!     Dim angleRad As Double
//!     Dim pt As Point3D
//!     
//!     angleRad = angleDegrees * (Pi / 180)
//!     
//!     pt.X = x * Cos(angleRad) - z * Sin(angleRad)
//!     pt.Y = y
//!     pt.Z = x * Sin(angleRad) + z * Cos(angleRad)
//!     
//!     RotateYaw = pt
//! End Function
//! ```
//!
//! ### Fourier Series
//!
//! ```vb
//! Function FourierCosine(x As Double, coefficients() As Double) As Double
//!     Const Pi As Double = 3.14159265358979
//!     Dim result As Double
//!     Dim i As Integer
//!     
//!     result = coefficients(0) / 2  ' a0/2 term
//!     
//!     For i = 1 To UBound(coefficients)
//!         result = result + coefficients(i) * Cos(i * x)
//!     Next i
//!     
//!     FourierCosine = result
//! End Function
//! ```
//!
//! ### Dot Product Calculation
//!
//! ```vb
//! Function DotProduct(x1 As Double, y1 As Double, _
//!                     x2 As Double, y2 As Double) As Double
//!     ' Alternative: using angle between vectors
//!     Dim magnitude1 As Double, magnitude2 As Double
//!     Dim angle As Double
//!     
//!     magnitude1 = Sqr(x1 * x1 + y1 * y1)
//!     magnitude2 = Sqr(x2 * x2 + y2 * y2)
//!     
//!     ' Get angle between vectors using Atn2
//!     angle = Atn2(y2, x2) - Atn2(y1, x1)
//!     
//!     DotProduct = magnitude1 * magnitude2 * Cos(angle)
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeCos(angle As Double) As Double
//!     On Error GoTo ErrorHandler
//!     
//!     SafeCos = Cos(angle)
//!     Exit Function
//!     
//! ErrorHandler:
//!     ' Cos rarely throws errors, but handle overflow
//!     MsgBox "Error calculating cosine: " & Err.Description
//!     SafeCos = 0
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 6** (Overflow): Can occur with extremely large angle values
//! - **Error 13** (Type mismatch): Non-numeric argument
//!
//! ## Performance Considerations
//!
//! - `Cos` is a native function with good performance
//! - For repeated calculations with the same angles, cache results
//! - Consider lookup tables for fixed angle values in performance-critical code
//! - Reducing angle to [0, 2π] range before calling Cos can improve accuracy
//!
//! ## Mathematical Properties
//!
//! ### Pythagorean Identity
//!
//! ```vb
//! ' Cos²(x) + Sin²(x) = 1
//! Function VerifyPythagorean(angle As Double) As Boolean
//!     Dim cosSquared As Double, sinSquared As Double
//!     cosSquared = Cos(angle) ^ 2
//!     sinSquared = Sin(angle) ^ 2
//!     VerifyPythagorean = Abs((cosSquared + sinSquared) - 1) < 0.0000001
//! End Function
//! ```
//!
//! ### Even Function Property
//!
//! ```vb
//! ' Cos(-x) = Cos(x)
//! Function VerifyEvenProperty(angle As Double) As Boolean
//!     VerifyEvenProperty = Abs(Cos(-angle) - Cos(angle)) < 0.0000001
//! End Function
//! ```
//!
//! ### Angle Addition Formula
//!
//! ```vb
//! ' Cos(a + b) = Cos(a)Cos(b) - Sin(a)Sin(b)
//! Function CosSum(angleA As Double, angleB As Double) As Double
//!     CosSum = Cos(angleA) * Cos(angleB) - Sin(angleA) * Sin(angleB)
//! End Function
//! ```
//!
//! ## Limitations
//!
//! - Argument must be in radians (not degrees)
//! - Very large angles may lose precision due to floating-point limitations
//! - Return value is always between -1 and 1
//! - Small rounding errors may occur near critical angles (e.g., π/2)
//! - For angles outside normal range, consider normalizing to [0, 2π]
//!
//! ## Related Functions
//!
//! - `Sin`: Returns the sine of an angle (complementary to cosine)
//! - `Tan`: Returns the tangent of an angle (Sin/Cos)
//! - `Atn`: Returns the arctangent (inverse tangent)
//! - `Acos`: Arc cosine (inverse of cosine, not built-in VB6)
//! - `Sqr`: Square root function (useful for magnitude calculations)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn cos_basic() {
        let source = r#"
result = Cos(angle)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_with_zero() {
        let source = r#"
value = Cos(0)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_with_pi() {
        let source = r#"
Const Pi As Double = 3.14159265358979
result = Cos(Pi)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_degrees_to_radians() {
        let source = r#"
Const Pi As Double = 3.14159265358979
radians = degrees * (Pi / 180)
result = Cos(radians)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_in_function() {
        let source = r#"
Function CosDegrees(degrees As Double) As Double
    Const Pi As Double = 3.14159265358979
    CosDegrees = Cos(degrees * (Pi / 180))
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_circle_calculation() {
        let source = r#"
x = centerX + radius * Cos(angle)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_rotation() {
        let source = r#"
newX = x * Cos(angle) - y * Sin(angle)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_wave_generation() {
        let source = r#"
For i = 0 To samples - 1
    wave(i) = amplitude * Cos(angle)
Next i
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_harmonic_motion() {
        let source = r#"
Const Pi As Double = 3.14159265358979
position = amplitude * Cos(2 * Pi * frequency * time)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_in_assignment() {
        let source = r#"
Dim result As Double
result = Cos(1.5708)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_polar_to_cartesian() {
        let source = r#"
x = radius * Cos(angle)
y = radius * Sin(angle)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_with_expression() {
        let source = r#"
result = Cos(Pi / 4)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_multiple_operations() {
        let source = r#"
value = amplitude * Cos(2 * Pi * frequency * time + phase)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_in_if_statement() {
        let source = r#"
If Cos(angle) > 0 Then
    ProcessPositive
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_pythagorean_identity() {
        let source = r#"
sum = Cos(angle) ^ 2 + Sin(angle) ^ 2
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_in_array() {
        let source = r#"
values(i) = Cos(angles(i))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_ellipse() {
        let source = r#"
ptX = centerX + radiusX * Cos(angle)
ptY = centerY + radiusY * Sin(angle)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_3d_rotation() {
        let source = r#"
newX = x * Cos(angle) - z * Sin(angle)
newZ = x * Sin(angle) + z * Cos(angle)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_in_select_case() {
        let source = r#"
Select Case Cos(angle)
    Case Is > 0.5
        HandleLarge
    Case Else
        HandleSmall
End Select
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_nested_call() {
        let source = r#"
result = Cos(Cos(x))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_in_do_loop() {
        let source = r#"
Do While Cos(angle) > threshold
    angle = angle + step
Loop
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_angle_addition() {
        let source = r#"
result = Cos(a) * Cos(b) - Sin(a) * Sin(b)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_with_abs() {
        let source = r#"
magnitude = Abs(Cos(angle))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_in_print() {
        let source = r#"
Print "Cosine: "; Cos(angle)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn cos_with_whitespace() {
        let source = r#"
result = Cos( angle )
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Cos"));
        assert!(debug.contains("Identifier"));
    }
}
