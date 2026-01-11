//! # Sqr Function
//!
//! Returns a Double specifying the square root of a number.
//!
//! ## Syntax
//!
//! ```vb
//! Sqr(number)
//! ```
//!
//! ## Parameters
//!
//! - `number` - Required. Double or any valid numeric expression greater than or equal to 0.
//!
//! ## Return Value
//!
//! Returns a Double containing the square root of the number.
//!
//! ## Remarks
//!
//! The Sqr function calculates the positive square root of a non-negative number. It's one of the fundamental mathematical functions in VB6 and is commonly used in geometric calculations, statistical analysis, and scientific applications.
//!
//! Key characteristics:
//! - Returns the positive (principal) square root only
//! - Argument must be >= 0 (negative values cause Error 5)
//! - Returns Double for maximum precision
//! - Sqr(0) = 0
//! - Sqr(1) = 1
//! - Sqr(x) * Sqr(x) = x (within floating-point precision)
//! - Inverse operation of squaring: Sqr(x^2) = x (for x >= 0)
//!
//! Mathematical relationships:
//! - Sqr(x * y) = Sqr(x) * Sqr(y)
//! - Sqr(x / y) = Sqr(x) / Sqr(y)
//! - Sqr(x^2) = Abs(x)
//! - x^(1/2) = Sqr(x)
//! - x^(1/n) can be calculated as x^(1/n) = Exp(Log(x) / n)
//!
//! ## Typical Uses
//!
//! 1. **Distance Calculations**: Calculate distance between two points
//! 2. **Pythagorean Theorem**: Find hypotenuse or sides of right triangles
//! 3. **Standard Deviation**: Statistical calculations
//! 4. **Quadratic Equations**: Solve equations using quadratic formula
//! 5. **Normalization**: Normalize vectors and values
//! 6. **Root Mean Square**: Calculate RMS values
//! 7. **Geometric Calculations**: Circle, sphere, and other shape calculations
//! 8. **Physics Simulations**: Velocity, acceleration, and energy calculations
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Calculate square root of a number
//! Dim result As Double
//! result = Sqr(25)
//! ' Returns 5
//! ```
//!
//! ```vb
//! ' Example 2: Distance between two points (Pythagorean theorem)
//! Dim x1 As Double, y1 As Double
//! Dim x2 As Double, y2 As Double
//! Dim distance As Double
//!
//! x1 = 0: y1 = 0
//! x2 = 3: y2 = 4
//! distance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
//! ' Returns 5
//! ```
//!
//! ```vb
//! ' Example 3: Calculate hypotenuse of right triangle
//! Dim sideA As Double
//! Dim sideB As Double
//! Dim hypotenuse As Double
//!
//! sideA = 3
//! sideB = 4
//! hypotenuse = Sqr(sideA ^ 2 + sideB ^ 2)
//! ' Returns 5
//! ```
//!
//! ```vb
//! ' Example 4: Quadratic formula
//! Dim a As Double, b As Double, c As Double
//! Dim discriminant As Double
//! Dim root1 As Double, root2 As Double
//!
//! a = 1: b = -5: c = 6
//! discriminant = b ^ 2 - 4 * a * c
//!
//! If discriminant >= 0 Then
//!     root1 = (-b + Sqr(discriminant)) / (2 * a)
//!     root2 = (-b - Sqr(discriminant)) / (2 * a)
//! End If
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `CalculateDistance2D`
//! Calculate distance between two 2D points
//! ```vb
//! Function CalculateDistance2D(x1 As Double, y1 As Double, _
//!                              x2 As Double, y2 As Double) As Double
//!     Dim dx As Double
//!     Dim dy As Double
//!     
//!     dx = x2 - x1
//!     dy = y2 - y1
//!     
//!     CalculateDistance2D = Sqr(dx * dx + dy * dy)
//! End Function
//! ```
//!
//! ### Pattern 2: `CalculateDistance3D`
//! Calculate distance between two 3D points
//! ```vb
//! Function CalculateDistance3D(x1 As Double, y1 As Double, z1 As Double, _
//!                              x2 As Double, y2 As Double, z2 As Double) As Double
//!     Dim dx As Double, dy As Double, dz As Double
//!     
//!     dx = x2 - x1
//!     dy = y2 - y1
//!     dz = z2 - z1
//!     
//!     CalculateDistance3D = Sqr(dx * dx + dy * dy + dz * dz)
//! End Function
//! ```
//!
//! ### Pattern 3: `CalculateHypotenuse`
//! Calculate hypotenuse of right triangle
//! ```vb
//! Function CalculateHypotenuse(sideA As Double, sideB As Double) As Double
//!     CalculateHypotenuse = Sqr(sideA ^ 2 + sideB ^ 2)
//! End Function
//! ```
//!
//! ### Pattern 4: `CalculateStandardDeviation`
//! Calculate standard deviation of values
//! ```vb
//! Function CalculateStandardDeviation(values() As Double) As Double
//!     Dim i As Integer
//!     Dim sum As Double
//!     Dim mean As Double
//!     Dim sumSquaredDiff As Double
//!     Dim count As Integer
//!     
//!     count = UBound(values) - LBound(values) + 1
//!     
//!     ' Calculate mean
//!     sum = 0
//!     For i = LBound(values) To UBound(values)
//!         sum = sum + values(i)
//!     Next i
//!     mean = sum / count
//!     
//!     ' Calculate sum of squared differences
//!     sumSquaredDiff = 0
//!     For i = LBound(values) To UBound(values)
//!         sumSquaredDiff = sumSquaredDiff + (values(i) - mean) ^ 2
//!     Next i
//!     
//!     CalculateStandardDeviation = Sqr(sumSquaredDiff / count)
//! End Function
//! ```
//!
//! ### Pattern 5: `NormalizeVector`
//! Normalize a 2D vector
//! ```vb
//! Sub NormalizeVector(x As Double, y As Double)
//!     Dim magnitude As Double
//!     
//!     magnitude = Sqr(x * x + y * y)
//!     
//!     If magnitude > 0 Then
//!         x = x / magnitude
//!         y = y / magnitude
//!     End If
//! End Sub
//! ```
//!
//! ### Pattern 6: `CalculateRMS`
//! Calculate root mean square
//! ```vb
//! Function CalculateRMS(values() As Double) As Double
//!     Dim i As Integer
//!     Dim sumSquares As Double
//!     Dim count As Integer
//!     
//!     count = UBound(values) - LBound(values) + 1
//!     sumSquares = 0
//!     
//!     For i = LBound(values) To UBound(values)
//!         sumSquares = sumSquares + values(i) ^ 2
//!     Next i
//!     
//!     CalculateRMS = Sqr(sumSquares / count)
//! End Function
//! ```
//!
//! ### Pattern 7: `SolveQuadratic`
//! Solve quadratic equation ax² + bx + c = 0
//! ```vb
//! Function SolveQuadratic(a As Double, b As Double, c As Double, _
//!                         root1 As Double, root2 As Double) As Boolean
//!     Dim discriminant As Double
//!     
//!     discriminant = b * b - 4 * a * c
//!     
//!     If discriminant < 0 Then
//!         SolveQuadratic = False
//!         Exit Function
//!     End If
//!     
//!     root1 = (-b + Sqr(discriminant)) / (2 * a)
//!     root2 = (-b - Sqr(discriminant)) / (2 * a)
//!     SolveQuadratic = True
//! End Function
//! ```
//!
//! ### Pattern 8: `CalculateCircleRadius`
//! Calculate radius from area
//! ```vb
//! Function CalculateCircleRadius(area As Double) As Double
//!     Const PI As Double = 3.14159265358979
//!     CalculateCircleRadius = Sqr(area / PI)
//! End Function
//! ```
//!
//! ### Pattern 9: `CalculateVelocity`
//! Calculate velocity from kinetic energy
//! ```vb
//! Function CalculateVelocity(kineticEnergy As Double, mass As Double) As Double
//!     ' KE = 1/2 * m * v^2
//!     ' v = Sqr(2 * KE / m)
//!     If mass > 0 Then
//!         CalculateVelocity = Sqr(2 * kineticEnergy / mass)
//!     Else
//!         CalculateVelocity = 0
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 10: `IsPerfectSquare`
//! Check if number is a perfect square
//! ```vb
//! Function IsPerfectSquare(n As Long) As Boolean
//!     Dim sqrtValue As Double
//!     Dim intValue As Long
//!     
//!     If n < 0 Then
//!         IsPerfectSquare = False
//!         Exit Function
//!     End If
//!     
//!     sqrtValue = Sqr(n)
//!     intValue = CLng(sqrtValue)
//!     IsPerfectSquare = (intValue * intValue = n)
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: `GeometryHelper` Class
//! Geometric calculations using Sqr
//! ```vb
//! ' Class: GeometryHelper
//!
//! Public Function DistanceBetweenPoints(x1 As Double, y1 As Double, _
//!                                       x2 As Double, y2 As Double) As Double
//!     DistanceBetweenPoints = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
//! End Function
//!
//! Public Function PointToLineDistance(px As Double, py As Double, _
//!                                     x1 As Double, y1 As Double, _
//!                                     x2 As Double, y2 As Double) As Double
//!     ' Distance from point to line segment
//!     Dim A As Double, B As Double, C As Double
//!     
//!     A = px - x1
//!     B = py - y1
//!     C = x2 - x1
//!     Dim D As Double
//!     D = y2 - y1
//!     
//!     Dim dot As Double
//!     dot = A * C + B * D
//!     Dim lenSq As Double
//!     lenSq = C * C + D * D
//!     
//!     Dim param As Double
//!     param = -1
//!     If lenSq <> 0 Then param = dot / lenSq
//!     
//!     Dim xx As Double, yy As Double
//!     
//!     If param < 0 Then
//!         xx = x1
//!         yy = y1
//!     ElseIf param > 1 Then
//!         xx = x2
//!         yy = y2
//!     Else
//!         xx = x1 + param * C
//!         yy = y1 + param * D
//!     End If
//!     
//!     Dim dx As Double, dy As Double
//!     dx = px - xx
//!     dy = py - yy
//!     
//!     PointToLineDistance = Sqr(dx * dx + dy * dy)
//! End Function
//!
//! Public Function CircleArea(radius As Double) As Double
//!     Const PI As Double = 3.14159265358979
//!     CircleArea = PI * radius * radius
//! End Function
//!
//! Public Function CircleRadiusFromArea(area As Double) As Double
//!     Const PI As Double = 3.14159265358979
//!     CircleRadiusFromArea = Sqr(area / PI)
//! End Function
//!
//! Public Function SphereVolume(radius As Double) As Double
//!     Const PI As Double = 3.14159265358979
//!     SphereVolume = (4# / 3#) * PI * radius ^ 3
//! End Function
//!
//! Public Function SphereRadiusFromVolume(volume As Double) As Double
//!     Const PI As Double = 3.14159265358979
//!     SphereRadiusFromVolume = ((3 * volume) / (4 * PI)) ^ (1# / 3#)
//! End Function
//! ```
//!
//! ### Example 2: `StatisticsCalculator` Class
//! Statistical calculations with Sqr
//! ```vb
//! ' Class: StatisticsCalculator
//! Private m_values() As Double
//! Private m_count As Integer
//!
//! Public Sub AddValue(value As Double)
//!     If m_count > UBound(m_values) Then
//!         ReDim Preserve m_values(0 To m_count * 2)
//!     End If
//!     m_values(m_count) = value
//!     m_count = m_count + 1
//! End Sub
//!
//! Public Sub Clear()
//!     m_count = 0
//!     ReDim m_values(0 To 9)
//! End Sub
//!
//! Private Sub Class_Initialize()
//!     Clear
//! End Sub
//!
//! Public Function GetMean() As Double
//!     Dim sum As Double
//!     Dim i As Integer
//!     
//!     If m_count = 0 Then
//!         GetMean = 0
//!         Exit Function
//!     End If
//!     
//!     sum = 0
//!     For i = 0 To m_count - 1
//!         sum = sum + m_values(i)
//!     Next i
//!     
//!     GetMean = sum / m_count
//! End Function
//!
//! Public Function GetStandardDeviation() As Double
//!     Dim mean As Double
//!     Dim sumSquaredDiff As Double
//!     Dim i As Integer
//!     
//!     If m_count = 0 Then
//!         GetStandardDeviation = 0
//!         Exit Function
//!     End If
//!     
//!     mean = GetMean()
//!     sumSquaredDiff = 0
//!     
//!     For i = 0 To m_count - 1
//!         sumSquaredDiff = sumSquaredDiff + (m_values(i) - mean) ^ 2
//!     Next i
//!     
//!     GetStandardDeviation = Sqr(sumSquaredDiff / m_count)
//! End Function
//!
//! Public Function GetVariance() As Double
//!     Dim sd As Double
//!     sd = GetStandardDeviation()
//!     GetVariance = sd * sd
//! End Function
//!
//! Public Function GetRMS() As Double
//!     Dim sumSquares As Double
//!     Dim i As Integer
//!     
//!     If m_count = 0 Then
//!         GetRMS = 0
//!         Exit Function
//!     End If
//!     
//!     sumSquares = 0
//!     For i = 0 To m_count - 1
//!         sumSquares = sumSquares + m_values(i) ^ 2
//!     Next i
//!     
//!     GetRMS = Sqr(sumSquares / m_count)
//! End Function
//!
//! Public Property Get Count() As Integer
//!     Count = m_count
//! End Property
//! ```
//!
//! ### Example 3: `PhysicsEngine` Module
//! Physics calculations using Sqr
//! ```vb
//! ' Module: PhysicsEngine
//!
//! Public Function CalculateVelocityFromEnergy(kineticEnergy As Double, _
//!                                             mass As Double) As Double
//!     ' v = Sqr(2 * KE / m)
//!     If mass > 0 Then
//!         CalculateVelocityFromEnergy = Sqr(2 * kineticEnergy / mass)
//!     Else
//!         CalculateVelocityFromEnergy = 0
//!     End If
//! End Function
//!
//! Public Function CalculateEscapeVelocity(mass As Double, radius As Double) As Double
//!     ' v_escape = Sqr(2 * G * M / R)
//!     Const G As Double = 6.674E-11  ' Gravitational constant
//!     CalculateEscapeVelocity = Sqr(2 * G * mass / radius)
//! End Function
//!
//! Public Function CalculatePeriod(length As Double) As Double
//!     ' Period of simple pendulum: T = 2π * Sqr(L/g)
//!     Const PI As Double = 3.14159265358979
//!     Const g As Double = 9.81  ' Gravity
//!     CalculatePeriod = 2 * PI * Sqr(length / g)
//! End Function
//!
//! Public Function CalculateFallTime(height As Double) As Double
//!     ' Time to fall from height: t = Sqr(2h/g)
//!     Const g As Double = 9.81  ' Gravity
//!     CalculateFallTime = Sqr(2 * height / g)
//! End Function
//!
//! Public Function CalculateImpactVelocity(height As Double) As Double
//!     ' v = Sqr(2gh)
//!     Const g As Double = 9.81  ' Gravity
//!     CalculateImpactVelocity = Sqr(2 * g * height)
//! End Function
//!
//! Public Function CalculateOrbitalVelocity(mass As Double, radius As Double) As Double
//!     ' v = Sqr(G * M / R)
//!     Const G As Double = 6.674E-11  ' Gravitational constant
//!     CalculateOrbitalVelocity = Sqr(G * mass / radius)
//! End Function
//! ```
//!
//! ### Example 4: `VectorMath` Module
//! Vector operations using Sqr
//! ```vb
//! ' Module: VectorMath
//!
//! Public Function VectorMagnitude2D(x As Double, y As Double) As Double
//!     VectorMagnitude2D = Sqr(x * x + y * y)
//! End Function
//!
//! Public Function VectorMagnitude3D(x As Double, y As Double, z As Double) As Double
//!     VectorMagnitude3D = Sqr(x * x + y * y + z * z)
//! End Function
//!
//! Public Sub NormalizeVector2D(x As Double, y As Double)
//!     Dim magnitude As Double
//!     magnitude = VectorMagnitude2D(x, y)
//!     
//!     If magnitude > 0 Then
//!         x = x / magnitude
//!         y = y / magnitude
//!     End If
//! End Sub
//!
//! Public Sub NormalizeVector3D(x As Double, y As Double, z As Double)
//!     Dim magnitude As Double
//!     magnitude = VectorMagnitude3D(x, y, z)
//!     
//!     If magnitude > 0 Then
//!         x = x / magnitude
//!         y = y / magnitude
//!         z = z / magnitude
//!     End If
//! End Sub
//!
//! Public Function DotProduct2D(x1 As Double, y1 As Double, _
//!                              x2 As Double, y2 As Double) As Double
//!     DotProduct2D = x1 * x2 + y1 * y2
//! End Function
//!
//! Public Function VectorDistance2D(x1 As Double, y1 As Double, _
//!                                  x2 As Double, y2 As Double) As Double
//!     VectorDistance2D = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
//! End Function
//!
//! Public Function VectorAngleBetween(x1 As Double, y1 As Double, _
//!                                    x2 As Double, y2 As Double) As Double
//!     ' Returns angle in radians
//!     Dim dot As Double
//!     Dim mag1 As Double, mag2 As Double
//!     
//!     dot = DotProduct2D(x1, y1, x2, y2)
//!     mag1 = VectorMagnitude2D(x1, y1)
//!     mag2 = VectorMagnitude2D(x2, y2)
//!     
//!     If mag1 > 0 And mag2 > 0 Then
//!         VectorAngleBetween = Atn(Sqr(1 - (dot / (mag1 * mag2)) ^ 2) / (dot / (mag1 * mag2)))
//!     Else
//!         VectorAngleBetween = 0
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! The Sqr function can generate the following errors:
//!
//! - **Error 5** (Invalid procedure call or argument): If number is negative
//! - **Error 13** (Type mismatch): If argument is not numeric
//!
//! Always validate inputs:
//! ```vb
//! On Error Resume Next
//! If value >= 0 Then
//!     result = Sqr(value)
//!     If Err.Number <> 0 Then
//!         MsgBox "Error calculating square root: " & Err.Description
//!     End If
//! Else
//!     MsgBox "Cannot calculate square root of negative number"
//! End If
//! ```
//!
//! ## Performance Considerations
//!
//! - Sqr is a relatively fast operation
//! - Uses hardware floating-point unit when available
//! - For repeated calculations, consider caching results
//! - Sqr(x^2) is slower than Abs(x) for getting absolute value
//! - For integer square roots, consider using Int(Sqr(x))
//!
//! ## Best Practices
//!
//! 1. **Validate Input**: Always ensure argument is non-negative
//! 2. **Use for Distances**: Ideal for Euclidean distance calculations
//! 3. **Combine with ^2**: Use for Pythagorean theorem applications
//! 4. **Handle Zero**: Remember Sqr(0) = 0 is valid
//! 5. **Precision Aware**: Understand floating-point precision limits
//! 6. **Error Handling**: Trap errors for negative inputs
//! 7. **Performance**: Cache results when used repeatedly
//! 8. **Consider Abs**: For Sqr(x^2), use Abs(x) instead
//! 9. **Document Units**: Comment on units in physics calculations
//! 10. **Test Edge Cases**: Test with 0, very small, and very large values
//!
//! ## Comparison with Related Functions
//!
//! | Operation | VB6 Function | Example | Result |
//! |-----------|--------------|---------|--------|
//! | Square root | Sqr(x) | Sqr(25) | 5 |
//! | Cube root | x^(1/3) | 27^(1/3) | 3 |
//! | Nth root | x^(1/n) | 16^(1/4) | 2 |
//! | Power | x^y | 2^3 | 8 |
//! | Absolute value | Abs(x) | Abs(-5) | 5 |
//!
//! ## Platform Considerations
//!
//! - Available in VB6, VBA (all versions)
//! - Part of core mathematical functions
//! - Uses IEEE 754 floating-point arithmetic
//! - Consistent behavior across platforms
//! - Hardware-accelerated on modern processors
//!
//! ## Limitations
//!
//! - Cannot calculate square root of negative numbers (use complex number libraries)
//! - Subject to floating-point precision limits (~15-17 significant digits)
//! - Sqr(x)^2 may not exactly equal x due to rounding
//! - Very large numbers may overflow Double range
//! - Very small numbers may underflow to zero
//!
//! ## Related Functions
//!
//! - `Abs`: Returns absolute value of a number
//! - `Exp`: Returns e raised to a power
//! - `Log`: Returns natural logarithm
//! - `^` operator: Raises number to a power (x^0.5 = Sqr(x))
//!
#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn sqr_basic() {
        let source = r"
Sub Test()
    Dim result As Double
    result = Sqr(25)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("result"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Sqr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("25"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_with_variable() {
        let source = r"
Sub Test()
    Dim value As Double
    Dim sqrtValue As Double
    value = 16
    sqrtValue = Sqr(value)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("value"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("sqrtValue"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("value"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("16"),
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("sqrtValue"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Sqr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("value"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_if_statement() {
        let source = r#"
Sub Test()
    If Sqr(value) > 10 Then
        MsgBox "Large"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Sqr"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("value"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("MsgBox"),
                                Whitespace,
                                StringLiteral ("\"Large\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_function_return() {
        let source = r"
Function CalculateDistance(dx As Double, dy As Double) As Double
    CalculateDistance = Sqr(dx * dx + dy * dy)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CalculateDistance"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("dx"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("dy"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("CalculateDistance"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Sqr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("dx"),
                                            },
                                            Whitespace,
                                            MultiplicationOperator,
                                            Whitespace,
                                            IdentifierExpression {
                                                Identifier ("dx"),
                                            },
                                        },
                                        Whitespace,
                                        AdditionOperator,
                                        Whitespace,
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("dy"),
                                            },
                                            Whitespace,
                                            MultiplicationOperator,
                                            Whitespace,
                                            IdentifierExpression {
                                                Identifier ("dy"),
                                            },
                                        },
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_variable_assignment() {
        let source = r"
Sub Test()
    Dim hypotenuse As Double
    hypotenuse = Sqr(3 ^ 2 + 4 ^ 2)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("hypotenuse"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("hypotenuse"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Sqr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        BinaryExpression {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("3"),
                                            },
                                            Whitespace,
                                            ExponentiationOperator,
                                            Whitespace,
                                            NumericLiteralExpression {
                                                IntegerLiteral ("2"),
                                            },
                                        },
                                        Whitespace,
                                        AdditionOperator,
                                        Whitespace,
                                        BinaryExpression {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("4"),
                                            },
                                            Whitespace,
                                            ExponentiationOperator,
                                            Whitespace,
                                            NumericLiteralExpression {
                                                IntegerLiteral ("2"),
                                            },
                                        },
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Square root: " & Sqr(100)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Square root: \""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Sqr"),
                        LeftParenthesis,
                        IntegerLiteral ("100"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_debug_print() {
        let source = r"
Sub Test()
    Debug.Print Sqr(144)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("Sqr"),
                        LeftParenthesis,
                        IntegerLiteral ("144"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_select_case() {
        let source = r#"
Sub Test()
    Select Case Sqr(value)
        Case Is > 10
            MsgBox "Large"
        Case Is > 5
            MsgBox "Medium"
        Case Else
            MsgBox "Small"
    End Select
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        CallExpression {
                            Identifier ("Sqr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("value"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IsKeyword,
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            IntegerLiteral ("10"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("MsgBox"),
                                    Whitespace,
                                    StringLiteral ("\"Large\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IsKeyword,
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            IntegerLiteral ("5"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("MsgBox"),
                                    Whitespace,
                                    StringLiteral ("\"Medium\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseElseClause {
                            CaseKeyword,
                            Whitespace,
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("MsgBox"),
                                    Whitespace,
                                    StringLiteral ("\"Small\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_class_usage() {
        let source = r"
Class MathHelper
    Public Function GetSquareRoot(n As Double) As Double
        GetSquareRoot = Sqr(n)
    End Function
End Class
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            Unknown,
            Whitespace,
            CallStatement {
                Identifier ("MathHelper"),
                Newline,
            },
            Whitespace,
            FunctionStatement {
                PublicKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("GetSquareRoot"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("n"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetSquareRoot"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Sqr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("n"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
            Unknown,
            Whitespace,
            Unknown,
            Newline,
        ]);
    }

    #[test]
    fn sqr_with_statement() {
        let source = r"
Sub Test()
    With calculator
        Dim root As Double
        root = Sqr(.Value)
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("calculator"),
                        Newline,
                        StatementList {
                            Whitespace,
                            DimStatement {
                                DimKeyword,
                                Whitespace,
                                Identifier ("root"),
                                Whitespace,
                                AsKeyword,
                                Whitespace,
                                DoubleKeyword,
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("root"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Sqr"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                PeriodOperator,
                                            },
                                        },
                                    },
                                },
                            },
                            CallStatement {
                                Identifier ("Value"),
                                RightParenthesis,
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_elseif() {
        let source = r"
Sub Test()
    Dim s As Double
    If value < 0 Then
        s = 0
    ElseIf value = 0 Then
        s = 0
    Else
        s = Sqr(value)
    End If
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("s"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("value"),
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("s"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("0"),
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        ElseIfClause {
                            ElseIfKeyword,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("value"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("0"),
                                },
                            },
                            Whitespace,
                            ThenKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("s"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("0"),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        ElseClause {
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("s"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Sqr"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("value"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_for_loop() {
        let source = r"
Sub Test()
    Dim i As Integer
    For i = 1 To 10
        Debug.Print Sqr(i)
    Next i
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    ForStatement {
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("i"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("1"),
                        },
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                Identifier ("Sqr"),
                                LeftParenthesis,
                                Identifier ("i"),
                                RightParenthesis,
                                Newline,
                            },
                            Whitespace,
                        },
                        NextKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_do_while() {
        let source = r"
Sub Test()
    Do While value > 1
        value = Sqr(value)
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("value"),
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("value"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Sqr"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("value"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        LoopKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_do_until() {
        let source = r"
Sub Test()
    Do Until Sqr(total) < threshold
        total = total - 1
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        UntilKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Sqr"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("total"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("threshold"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("total"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("total"),
                                    },
                                    Whitespace,
                                    SubtractionOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        LoopKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_while_wend() {
        let source = r"
Sub Test()
    While x >= 0
        result = Sqr(x)
        x = x - 1
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("x"),
                            },
                            Whitespace,
                            GreaterThanOrEqualOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("result"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Sqr"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("x"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("x"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("x"),
                                    },
                                    Whitespace,
                                    SubtractionOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_parentheses() {
        let source = r"
Sub Test()
    Dim distance As Double
    distance = (Sqr(dx * dx + dy * dy))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("distance"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("distance"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            CallExpression {
                                Identifier ("Sqr"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    Identifier ("dx"),
                                                },
                                                Whitespace,
                                                MultiplicationOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("dx"),
                                                },
                                            },
                                            Whitespace,
                                            AdditionOperator,
                                            Whitespace,
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    Identifier ("dy"),
                                                },
                                                Whitespace,
                                                MultiplicationOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("dy"),
                                                },
                                            },
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_iif() {
        let source = r"
Sub Test()
    Dim safe As Double
    safe = IIf(value >= 0, Sqr(value), 0)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("safe"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("safe"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("IIf"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("value"),
                                        },
                                        Whitespace,
                                        GreaterThanOrEqualOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("0"),
                                        },
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    CallExpression {
                                        Identifier ("Sqr"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("value"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("0"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_array_assignment() {
        let source = r"
Sub Test()
    Dim roots(10) As Double
    roots(0) = Sqr(25)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("roots"),
                        LeftParenthesis,
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        CallExpression {
                            Identifier ("roots"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("0"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Sqr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("25"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_property_assignment() {
        let source = r"
Class Calculator
    Public SquareRoot As Double
End Class

Sub Test()
    Dim calc As New Calculator
    calc.SquareRoot = Sqr(value)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            Unknown,
            Whitespace,
            CallStatement {
                Identifier ("Calculator"),
                Newline,
            },
            Whitespace,
            DimStatement {
                PublicKeyword,
                Whitespace,
                Identifier ("SquareRoot"),
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                Newline,
            },
            Unknown,
            Whitespace,
            Unknown,
            Newline,
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("calc"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        NewKeyword,
                        Whitespace,
                        Identifier ("Calculator"),
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        MemberAccessExpression {
                            Identifier ("calc"),
                            PeriodOperator,
                            Identifier ("SquareRoot"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Sqr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("value"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_function_argument() {
        let source = r"
Sub ProcessValue(v As Double)
End Sub

Sub Test()
    ProcessValue Sqr(100)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ProcessValue"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("v"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("ProcessValue"),
                        Whitespace,
                        Identifier ("Sqr"),
                        LeftParenthesis,
                        IntegerLiteral ("100"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_concatenation() {
        let source = r#"
Sub Test()
    Dim msg As String
    msg = "The square root is: " & Sqr(value)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("msg"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("msg"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"The square root is: \""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Sqr"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("value"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_comparison() {
        let source = r"
Sub Test()
    Dim isLarge As Boolean
    isLarge = (Sqr(area) > 100)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("isLarge"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        BooleanKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("isLarge"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("Sqr"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("area"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                GreaterThanOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("100"),
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_arithmetic() {
        let source = r"
Sub Test()
    Dim magnitude As Double
    magnitude = Sqr(x * x + y * y + z * z)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("magnitude"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("magnitude"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Sqr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        BinaryExpression {
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    Identifier ("x"),
                                                },
                                                Whitespace,
                                                MultiplicationOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("x"),
                                                },
                                            },
                                            Whitespace,
                                            AdditionOperator,
                                            Whitespace,
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    Identifier ("y"),
                                                },
                                                Whitespace,
                                                MultiplicationOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    Identifier ("y"),
                                                },
                                            },
                                        },
                                        Whitespace,
                                        AdditionOperator,
                                        Whitespace,
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("z"),
                                            },
                                            Whitespace,
                                            MultiplicationOperator,
                                            Whitespace,
                                            IdentifierExpression {
                                                Identifier ("z"),
                                            },
                                        },
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_quadratic_formula() {
        let source = r"
Sub Test()
    Dim discriminant As Double
    Dim root1 As Double
    discriminant = b * b - 4 * a * c
    root1 = (-b + Sqr(discriminant)) / (2 * a)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("discriminant"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("root1"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("discriminant"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("b"),
                                },
                                Whitespace,
                                MultiplicationOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("b"),
                                },
                            },
                            Whitespace,
                            SubtractionOperator,
                            Whitespace,
                            BinaryExpression {
                                BinaryExpression {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("4"),
                                    },
                                    Whitespace,
                                    MultiplicationOperator,
                                    Whitespace,
                                    IdentifierExpression {
                                        Identifier ("a"),
                                    },
                                },
                                Whitespace,
                                MultiplicationOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("c"),
                                },
                            },
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("root1"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            ParenthesizedExpression {
                                LeftParenthesis,
                                BinaryExpression {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        IdentifierExpression {
                                            Identifier ("b"),
                                        },
                                    },
                                    Whitespace,
                                    AdditionOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Sqr"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("discriminant"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            DivisionOperator,
                            Whitespace,
                            ParenthesizedExpression {
                                LeftParenthesis,
                                BinaryExpression {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("2"),
                                    },
                                    Whitespace,
                                    MultiplicationOperator,
                                    Whitespace,
                                    IdentifierExpression {
                                        Identifier ("a"),
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Dim s As Double
    s = Sqr(value)
    If Err.Number <> 0 Then
        MsgBox "Error"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    OnErrorStatement {
                        Whitespace,
                        OnKeyword,
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        ResumeKeyword,
                        Whitespace,
                        NextKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("s"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("s"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Sqr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("value"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            MemberAccessExpression {
                                Identifier ("Err"),
                                PeriodOperator,
                                Identifier ("Number"),
                            },
                            Whitespace,
                            InequalityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("MsgBox"),
                                Whitespace,
                                StringLiteral ("\"Error\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_on_error_goto() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Dim sqrtVal As Double
    sqrtVal = Sqr(number)
    Exit Sub
ErrorHandler:
    MsgBox "Error calculating square root"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    OnErrorStatement {
                        Whitespace,
                        OnKeyword,
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        GotoKeyword,
                        Whitespace,
                        Identifier ("ErrorHandler"),
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("sqrtVal"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("sqrtVal"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Sqr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("number"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    ExitStatement {
                        Whitespace,
                        ExitKeyword,
                        Whitespace,
                        SubKeyword,
                        Newline,
                    },
                    LabelStatement {
                        Identifier ("ErrorHandler"),
                        ColonOperator,
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Error calculating square root\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sqr_standard_deviation() {
        let source = r"
Sub Test()
    Dim variance As Double
    Dim stdDev As Double
    variance = 25
    stdDev = Sqr(variance)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("variance"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("stdDev"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("variance"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("25"),
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("stdDev"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Sqr"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("variance"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }
}
