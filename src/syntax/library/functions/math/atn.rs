//! # `Atn` Function
//!
//! Returns a `Double` specifying the arctangent of a number.
//!
//! ## Syntax
//!
//! ```vb
//! Atn(number)
//! ```
//!
//! ## Parameters
//!
//! - `number` - Required. A `Double` or any valid numeric expression. The tangent value for which you want the angle.
//!
//! ## Return Value
//!
//! Returns a `Double` representing the arctangent of the number in radians.
//!
//! - The result is in the range -π/2 to π/2 radians (-90° to 90°)
//! - To convert from radians to degrees, multiply by 180/π (approximately 57.2957795130823)
//! - To convert from degrees to radians, multiply by π/180 (approximately 0.0174532925199433)
//!
//! ## Remarks
//!
//! The `Atn` function takes the ratio of two sides of a right triangle (opposite/adjacent) and
//! returns the corresponding angle in radians. This is the inverse of the tangent function.
//!
//! ### Important Notes
//!
//! 1. **Return Type**: Always returns `Double` regardless of input type
//! 2. **Radians**: Result is always in radians, not degrees
//! 3. **Range**: Return value is between -π/2 and π/2 (-1.5708 to 1.5708 radians)
//! 4. **Ratio Input**: The argument represents the tangent (opposite/adjacent) of the angle
//! 5. **Inverse Function**: `Atn` is the inverse of `Tan` (`Atn(Tan(x)) = x` for x in valid range)
//!
//! ### Mathematical Relationship
//!
//! ```vb
//! ' For a right triangle:
//! angle = Atn(opposite / adjacent)
//!
//! ' Converting between functions:
//! Tan(Atn(x)) = x    ' Always true
//! Atn(Tan(x)) = x    ' True when -π/2 < x < π/2
//! ```
//!
//! ### Special Values
//!
//! - `Atn(0)` returns 0 (angle is 0 radians or 0°)
//! - `Atn(1)` returns π/4 (approximately 0.785398, which is 45°)
//! - `Atn(-1)` returns -π/4 (approximately -0.785398, which is -45°)
//! - Large positive values approach π/2 (90°)
//! - Large negative values approach -π/2 (-90°)
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! Dim angle As Double
//! angle = Atn(1)              ' Returns π/4 (approx 0.785398) = 45°
//! angle = Atn(0)              ' Returns 0 = 0°
//! angle = Atn(-1)             ' Returns -π/4 (approx -0.785398) = -45°
//! angle = Atn(1.732050808)    ' Returns π/3 (approx 1.047198) = 60°
//! ```
//!
//! ### Converting to Degrees
//!
//! ```vb
//! Function AtnDegrees(tangent As Double) As Double
//!     Const PI As Double = 3.14159265358979
//!     AtnDegrees = Atn(tangent) * 180 / PI
//! End Function
//!
//! ' Usage:
//! angle = AtnDegrees(1)       ' Returns 45 degrees
//! ```
//!
//! ### Calculating Angle from Triangle Sides
//!
//! ```vb
//! Function AngleFromSides(opposite As Double, adjacent As Double) As Double
//!     ' Returns angle in radians
//!     AngleFromSides = Atn(opposite / adjacent)
//! End Function
//!
//! ' For a right triangle with opposite=3, adjacent=4:
//! angle = AngleFromSides(3, 4)    ' Returns angle in radians
//! ```
//!
//! ### Calculating π
//!
//! ```vb
//! Function CalculatePi() As Double
//!     ' Since Atn(1) = π/4, then 4 * Atn(1) = π
//!     CalculatePi = 4 * Atn(1)
//! End Function
//! ```
//!
//! ### Full Circle Angle Calculation (Atn2 Emulation)
//!
//! ```vb
//! Function Atn2(y As Double, x As Double) As Double
//!     ' Emulates atan2 function for full circle (-π to π)
//!     Const PI As Double = 3.14159265358979
//!     
//!     If x > 0 Then
//!         Atn2 = Atn(y / x)
//!     ElseIf x < 0 Then
//!         If y >= 0 Then
//!             Atn2 = Atn(y / x) + PI
//!         Else
//!             Atn2 = Atn(y / x) - PI
//!         End If
//!     ElseIf y > 0 Then
//!         Atn2 = PI / 2
//!     ElseIf y < 0 Then
//!         Atn2 = -PI / 2
//!     Else
//!         Atn2 = 0  ' Undefined, but return 0
//!     End If
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### 1. Slope to Angle Conversion
//!
//! ```vb
//! Dim slope As Double
//! Dim angleRadians As Double
//! Dim angleDegrees As Double
//!
//! slope = 0.5  ' Rise over run
//! angleRadians = Atn(slope)
//! angleDegrees = angleRadians * 180 / (4 * Atn(1))
//! ```
//!
//! ### 2. Direction Calculation
//!
//! ```vb
//! Function GetDirection(deltaX As Double, deltaY As Double) As Double
//!     ' Returns angle in degrees (0-360)
//!     Const PI As Double = 3.14159265358979
//!     Dim angle As Double
//!     
//!     If deltaX = 0 Then
//!         If deltaY > 0 Then
//!             angle = 90
//!         ElseIf deltaY < 0 Then
//!             angle = 270
//!         Else
//!             angle = 0
//!         End If
//!     Else
//!         angle = Atn(deltaY / deltaX) * 180 / PI
//!         If deltaX < 0 Then angle = angle + 180
//!         If angle < 0 Then angle = angle + 360
//!     End If
//!     
//!     GetDirection = angle
//! End Function
//! ```
//!
//! ### 3. Distance and Angle from Origin
//!
//! ```vb
//! Sub GetPolarCoordinates(x As Double, y As Double, _
//!                         distance As Double, angle As Double)
//!     distance = Sqr(x * x + y * y)
//!     
//!     If x = 0 Then
//!         If y > 0 Then
//!             angle = 90
//!         Else
//!             angle = 270
//!         End If
//!     Else
//!         angle = Atn(y / x) * 180 / (4 * Atn(1))
//!         If x < 0 Then angle = angle + 180
//!     End If
//! End Sub
//! ```
//!
//! ### 4. Graphics Rotation
//!
//! ```vb
//! Function RotatePoint(x As Double, y As Double, _
//!                     centerX As Double, centerY As Double, _
//!                     angleDegrees As Double)
//!     Const PI As Double = 3.14159265358979
//!     Dim angleRadians As Double
//!     Dim currentAngle As Double
//!     Dim distance As Double
//!     Dim newAngle As Double
//!     
//!     ' Get current angle and distance from center
//!     distance = Sqr((x - centerX) ^ 2 + (y - centerY) ^ 2)
//!     currentAngle = Atn((y - centerY) / (x - centerX))
//!     
//!     ' Calculate new angle
//!     angleRadians = angleDegrees * PI / 180
//!     newAngle = currentAngle + angleRadians
//!     
//!     ' Calculate new position
//!     x = centerX + distance * Cos(newAngle)
//!     y = centerY + distance * Sin(newAngle)
//! End Function
//! ```
//!
//! ### 5. Navigation Bearing
//!
//! ```vb
//! Function CalculateBearing(lat1 As Double, lon1 As Double, _
//!                          lat2 As Double, lon2 As Double) As Double
//!     Const PI As Double = 3.14159265358979
//!     Dim dLon As Double
//!     Dim y As Double
//!     Dim x As Double
//!     Dim bearing As Double
//!     
//!     dLon = lon2 - lon1
//!     y = Sin(dLon) * Cos(lat2)
//!     x = Cos(lat1) * Sin(lat2) - Sin(lat1) * Cos(lat2) * Cos(dLon)
//!     bearing = Atn(y / x) * 180 / PI
//!     
//!     ' Normalize to 0-360
//!     CalculateBearing = (bearing + 360) Mod 360
//! End Function
//! ```
//!
//! ### 6. Inverse Trigonometry Relationships
//!
//! ```vb
//! ' Calculate arcsine using Atn
//! Function Asin(x As Double) As Double
//!     Asin = Atn(x / Sqr(1 - x * x))
//! End Function
//!
//! ' Calculate arccosine using Atn
//! Function Acos(x As Double) As Double
//!     Const PI As Double = 3.14159265358979
//!     Acos = PI / 2 - Atn(x / Sqr(1 - x * x))
//! End Function
//! ```
//!
//! ### 7. Tangent Line Slope
//!
//! ```vb
//! Function GetTangentAngle(x1 As Double, y1 As Double, _
//!                         x2 As Double, y2 As Double) As Double
//!     Dim slope As Double
//!     Const PI As Double = 3.14159265358979
//!     
//!     If x2 = x1 Then
//!         GetTangentAngle = 90  ' Vertical line
//!     Else
//!         slope = (y2 - y1) / (x2 - x1)
//!         GetTangentAngle = Atn(slope) * 180 / PI
//!     End If
//! End Function
//! ```
//!
//! ### 8. Game Development - Projectile Angle
//!
//! ```vb
//! Function AimAngle(shooterX As Double, shooterY As Double, _
//!                   targetX As Double, targetY As Double) As Double
//!     Const PI As Double = 3.14159265358979
//!     Dim deltaX As Double
//!     Dim deltaY As Double
//!     
//!     deltaX = targetX - shooterX
//!     deltaY = targetY - shooterY
//!     
//!     If deltaX = 0 Then
//!         If deltaY > 0 Then
//!             AimAngle = 90
//!         Else
//!             AimAngle = 270
//!         End If
//!     Else
//!         AimAngle = Atn(deltaY / deltaX) * 180 / PI
//!         If deltaX < 0 Then AimAngle = AimAngle + 180
//!         If AimAngle < 0 Then AimAngle = AimAngle + 360
//!     End If
//! End Function
//! ```
//!
//! ## Common Trigonometric Values
//!
//! | Angle (Degrees) | Angle (Radians) | Tangent | Atn(Tangent) |
//! |-----------------|-----------------|---------|--------------|
//! | 0°              | 0               | 0       | 0            |
//! | 30°             | π/6 ≈ 0.524     | 0.577   | π/6          |
//! | 45°             | π/4 ≈ 0.785     | 1       | π/4          |
//! | 60°             | π/3 ≈ 1.047     | 1.732   | π/3          |
//! | 90°             | π/2 ≈ 1.571     | ∞       | π/2 (limit)  |
//!
//! ## Type Conversion
//!
//! | Input Type | Converted To | Example |
//! |------------|--------------|---------|
//! | `Integer`  | `Double`     | Atn(1) → Atn(1.0) |
//! | `Long`     | `Double`     | Atn(10&) → Atn(10.0) |
//! | `Single`   | `Double`     | Atn(1.5!) → Atn(1.5) |
//! | `Double`   | `Double`     | Atn(1.5#) → 1.5 |
//! | `Currency` | `Double`     | Atn(100@) → Atn(100.0) |
//! | `Variant`  | `Double`     | Depends on content |
//!
//! ## Error Conditions
//!
//! - **Type Mismatch**: If the argument cannot be converted to a numeric value
//! - **No overflow**: Unlike some functions, `Atn` cannot overflow as it's bounded to ±π/2
//!
//! ## Related Functions
//!
//! - `Tan`: Returns the tangent of an angle (inverse of `Atn`)
//! - `Sin`: Returns the sine of an angle
//! - `Cos`: Returns the cosine of an angle
//! - `Sqr`: Returns the square root (used with `Atn` to calculate other inverse trig functions)
//! - `Abs`: Returns absolute value (useful in angle calculations)
//!
//! ## Performance Notes
//!
//! - `Atn` is a fast intrinsic function
//! - Implemented using CPU floating-point instructions
//! - No significant performance difference between `Integer`, `Long`, or `Double` arguments
//! - For repeated π calculations, cache the value of 4 * `Atn`(1) rather than recalculating
//!
//! ## Parsing Notes
//!
//! The `Atn` function is not a reserved keyword in VB6. It is parsed as a regular
//! function call (`CallExpression`). This module exists primarily for documentation
//! purposes and to provide a comprehensive test suite that validates the parser
//! correctly handles `Atn` function calls in various contexts.

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn atn_simple() {
        let source = r"
Sub Test()
    angle = Atn(1)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("angle"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Atn"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
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
    fn atn_with_zero() {
        let source = r"
Sub Test()
    result = Atn(0)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Atn"),
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
    fn atn_with_negative() {
        let source = r"
Sub Test()
    angle = Atn(-1)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("angle"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Atn"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
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
    fn atn_with_variable() {
        let source = r"
Sub Test()
    angle = Atn(tangentValue)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("angle"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Atn"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("tangentValue"),
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
    fn atn_calculate_pi() {
        let source = r"
Sub Test()
    pi = 4 * Atn(1)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("pi"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            NumericLiteralExpression {
                                IntegerLiteral ("4"),
                            },
                            Whitespace,
                            MultiplicationOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("Atn"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
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
    fn atn_to_degrees() {
        let source = r"
Sub Test()
    degrees = Atn(ratio) * 180 / (4 * Atn(1))
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("degrees"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("Atn"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("ratio"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                MultiplicationOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("180"),
                                },
                            },
                            Whitespace,
                            DivisionOperator,
                            Whitespace,
                            ParenthesizedExpression {
                                LeftParenthesis,
                                BinaryExpression {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("4"),
                                    },
                                    Whitespace,
                                    MultiplicationOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Atn"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("1"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
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
    fn atn_in_if_statement() {
        let source = r#"
Sub Test()
    If Atn(slope) > 0.785 Then
        Print "Angle > 45 degrees"
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
                                Identifier ("Atn"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("slope"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                SingleLiteral,
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            PrintStatement {
                                Whitespace,
                                PrintKeyword,
                                Whitespace,
                                StringLiteral ("\"Angle > 45 degrees\""),
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
    fn atn_in_for_loop() {
        let source = r"
Sub Test()
    For i = 0 To 10
        angle = Atn(i / 10)
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
                            IntegerLiteral ("0"),
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
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("angle"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Atn"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    Identifier ("i"),
                                                },
                                                Whitespace,
                                                DivisionOperator,
                                                Whitespace,
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("10"),
                                                },
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
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
    fn atn_with_division() {
        let source = r"
Sub Test()
    angle = Atn(opposite / adjacent)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("angle"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Atn"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("opposite"),
                                        },
                                        Whitespace,
                                        DivisionOperator,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("adjacent"),
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
    fn atn_with_sqr() {
        let source = r"
Sub Test()
    asinValue = Atn(x / Sqr(1 - x * x))
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("asinValue"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Atn"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("x"),
                                        },
                                        Whitespace,
                                        DivisionOperator,
                                        Whitespace,
                                        CallExpression {
                                            Identifier ("Sqr"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    BinaryExpression {
                                                        NumericLiteralExpression {
                                                            IntegerLiteral ("1"),
                                                        },
                                                        Whitespace,
                                                        SubtractionOperator,
                                                        Whitespace,
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
                                                    },
                                                },
                                            },
                                            RightParenthesis,
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
    fn atn_case_insensitive() {
        let source = r"
Sub Test()
    a = ATN(1)
    b = atn(1)
    c = AtN(1)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("a"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("ATN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
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
                            Identifier ("b"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("atn"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
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
                            Identifier ("c"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("AtN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
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
    fn atn_in_select_case() {
        let source = r#"
Sub Test()
    Select Case Atn(value)
        Case Is > 0.5
            Print "High"
        Case Else
            Print "Low"
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
                            Identifier ("Atn"),
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
                            SingleLiteral,
                            Newline,
                            StatementList {
                                PrintStatement {
                                    Whitespace,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"High\""),
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
                                PrintStatement {
                                    Whitespace,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Low\""),
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
    fn atn_in_while_loop() {
        let source = r"
Sub Test()
    While Atn(counter) < 1.5
        counter = counter + 0.1
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
                            CallExpression {
                                Identifier ("Atn"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("counter"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                SingleLiteral,
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("counter"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("counter"),
                                    },
                                    Whitespace,
                                    AdditionOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        SingleLiteral,
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
    fn atn_in_do_loop() {
        let source = r"
Sub Test()
    Do While Atn(x) < threshold
        x = x + step
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
                            CallExpression {
                                Identifier ("Atn"),
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
                                    AdditionOperator,
                                    Whitespace,
                                    IdentifierExpression {
                                        StepKeyword,
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
    fn atn_with_property_access() {
        let source = r"
Sub Test()
    angle = Atn(Me.Slope)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("angle"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Atn"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    MemberAccessExpression {
                                        MeKeyword,
                                        PeriodOperator,
                                        Identifier ("Slope"),
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
    fn atn_with_array_access() {
        let source = r"
Sub Test()
    angles(i) = Atn(tangents(i))
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
                    AssignmentStatement {
                        CallExpression {
                            Identifier ("angles"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("i"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Atn"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("tangents"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("i"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
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
    fn atn_nested_calls() {
        let source = r"
Sub Test()
    result = Atn(Tan(angle))
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Atn"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Tan"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("angle"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
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
    fn atn_in_function() {
        let source = r"
Function AtnDegrees(x As Double) As Double
    AtnDegrees = Atn(x) * 180 / (4 * Atn(1))
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("AtnDegrees"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("x"),
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
                            Identifier ("AtnDegrees"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("Atn"),
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
                                Whitespace,
                                MultiplicationOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("180"),
                                },
                            },
                            Whitespace,
                            DivisionOperator,
                            Whitespace,
                            ParenthesizedExpression {
                                LeftParenthesis,
                                BinaryExpression {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("4"),
                                    },
                                    Whitespace,
                                    MultiplicationOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Atn"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("1"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
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
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn atn_with_whitespace() {
        let source = r"
Sub Test()
    angle = Atn  (  1  )
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("angle"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Atn"),
                            Whitespace,
                            LeftParenthesis,
                            ArgumentList {
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                                Whitespace,
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
    fn atn_with_line_continuation() {
        let source = r"
Sub Test()
    angle = Atn _
        (tangent)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("angle"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Atn"),
                            Whitespace,
                            Underscore,
                            Newline,
                            Whitespace,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("tangent"),
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
    fn atn_multiple_on_line() {
        let source = r"
Sub Test()
    a = Atn(0): b = Atn(1): c = Atn(-1)
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("a"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Atn"),
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
                    },
                    Unknown,
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("b"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Atn"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                    },
                    Unknown,
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("c"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Atn"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    UnaryExpression {
                                        SubtractionOperator,
                                        NumericLiteralExpression {
                                            IntegerLiteral ("1"),
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
    fn atn_in_with_block() {
        let source = r"
Sub Test()
    With Triangle
        angle = Atn(.Opposite / .Adjacent)
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
                        Identifier ("Triangle"),
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("angle"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Atn"),
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
                                Identifier ("Opposite"),
                                Whitespace,
                                DivisionOperator,
                                Whitespace,
                                PeriodOperator,
                                Identifier ("Adjacent"),
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
    fn atn_in_print() {
        let source = r"
Sub Test()
    Print Atn(1)
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
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("Atn"),
                        LeftParenthesis,
                        IntegerLiteral ("1"),
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
    fn atn_module_level() {
        let source = r"Const PI As Double = 4 * Atn(1)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            DimStatement {
                ConstKeyword,
                Whitespace,
                Identifier ("PI"),
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    NumericLiteralExpression {
                        IntegerLiteral ("4"),
                    },
                    Whitespace,
                    MultiplicationOperator,
                    Whitespace,
                    CallExpression {
                        Identifier ("Atn"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                NumericLiteralExpression {
                                    IntegerLiteral ("1"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                },
            },
        ]);
    }

    #[test]
    fn atn_complex_expression() {
        let source = r"
Sub Test()
    bearing = (Atn(deltaY / deltaX) * 180 / pi + 360) Mod 360
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("bearing"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            ParenthesizedExpression {
                                LeftParenthesis,
                                BinaryExpression {
                                    BinaryExpression {
                                        BinaryExpression {
                                            CallExpression {
                                                Identifier ("Atn"),
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        BinaryExpression {
                                                            IdentifierExpression {
                                                                Identifier ("deltaY"),
                                                            },
                                                            Whitespace,
                                                            DivisionOperator,
                                                            Whitespace,
                                                            IdentifierExpression {
                                                                Identifier ("deltaX"),
                                                            },
                                                        },
                                                    },
                                                },
                                                RightParenthesis,
                                            },
                                            Whitespace,
                                            MultiplicationOperator,
                                            Whitespace,
                                            NumericLiteralExpression {
                                                IntegerLiteral ("180"),
                                            },
                                        },
                                        Whitespace,
                                        DivisionOperator,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("pi"),
                                        },
                                    },
                                    Whitespace,
                                    AdditionOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("360"),
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            ModKeyword,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("360"),
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
}
