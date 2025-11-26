/// # Sin Function
///
/// Returns a Double specifying the sine of an angle.
///
/// ## Syntax
///
/// ```vb
/// Sin(number)
/// ```
///
/// ## Parameters
///
/// - `number` - Required. Double or any valid numeric expression that expresses an angle in radians.
///
/// ## Return Value
///
/// Returns a Double value representing the sine of the angle:
/// - Range: -1 to 1 (inclusive)
/// - Sin(0) = 0
/// - Sin(π/2) ≈ 1
/// - Sin(π) ≈ 0
/// - Sin(3π/2) ≈ -1
///
/// ## Remarks
///
/// The Sin function takes an angle in radians and returns the ratio of two sides of a right triangle. The ratio is the length of the side opposite the angle divided by the length of the hypotenuse.
///
/// Key characteristics:
/// - Input is in radians, not degrees
/// - To convert degrees to radians: radians = degrees × (π / 180)
/// - To convert radians to degrees: degrees = radians × (180 / π)
/// - π (Pi) ≈ 3.14159265358979
/// - Use `Atn(1) * 4` to calculate π in VB6
/// - Periodic function: Sin(x) = Sin(x + 2π)
/// - Returns values between -1 and 1
///
/// The sine function is one of the fundamental trigonometric functions:
/// - Sin(x): Sine (this function)
/// - Cos(x): Cosine (use Cos function)
/// - Tan(x): Tangent (use Tan function or Sin(x)/Cos(x))
///
/// Common angles and their sines:
/// - Sin(0°) = Sin(0 rad) = 0
/// - Sin(30°) = Sin(π/6 rad) = 0.5
/// - Sin(45°) = Sin(π/4 rad) ≈ 0.707
/// - Sin(60°) = Sin(π/3 rad) ≈ 0.866
/// - Sin(90°) = Sin(π/2 rad) = 1
/// - Sin(180°) = Sin(π rad) = 0
/// - Sin(270°) = Sin(3π/2 rad) = -1
/// - Sin(360°) = Sin(2π rad) = 0
///
/// ## Typical Uses
///
/// 1. **Wave Generation**: Create sine waves for animations or signals
/// 2. **Circular Motion**: Calculate vertical position in circular paths
/// 3. **Oscillations**: Model periodic oscillating systems
/// 4. **Physics Simulations**: Projectile motion, pendulum swing
/// 5. **Graphics**: Rotation, transformation, curve drawing
/// 6. **Audio Processing**: Sine wave tone generation
/// 7. **Engineering Calculations**: Structural analysis, AC circuits
/// 8. **Game Development**: Movement patterns, trajectories
///
/// ## Basic Examples
///
/// ```vb
/// ' Example 1: Calculate sine of 45 degrees
/// Const PI As Double = 3.14159265358979
/// Dim angle45 As Double
/// Dim sineValue As Double
///
/// angle45 = 45 * (PI / 180)  ' Convert to radians
/// sineValue = Sin(angle45)   ' Returns ≈ 0.707
/// ```
///
/// ```vb
/// ' Example 2: Calculate sine of π/2 (90 degrees)
/// Const PI As Double = 3.14159265358979
/// Dim result As Double
/// result = Sin(PI / 2)  ' Returns 1
/// ```
///
/// ```vb
/// ' Example 3: Use with Atn to calculate π
/// Dim pi As Double
/// Dim sineValue As Double
/// pi = Atn(1) * 4
/// sineValue = Sin(pi)  ' Returns ≈ 0 (very small number)
/// ```
///
/// ```vb
/// ' Example 4: Sine wave generation
/// Dim i As Integer
/// Dim y As Double
/// For i = 0 To 360
///     y = Sin(i * (Atn(1) * 4) / 180)
///     Debug.Print i & " degrees: " & y
/// Next i
/// ```
///
/// ## Common Patterns
///
/// ### Pattern 1: DegreesToRadians
/// Convert degrees to radians for Sin function
/// ```vb
/// Function DegreesToRadians(degrees As Double) As Double
///     Const PI As Double = 3.14159265358979
///     DegreesToRadians = degrees * (PI / 180)
/// End Function
///
/// ' Usage:
/// result = Sin(DegreesToRadians(45))
/// ```
///
/// ### Pattern 2: SinDegrees
/// Sine function that accepts degrees
/// ```vb
/// Function SinDegrees(degrees As Double) As Double
///     Const PI As Double = 3.14159265358979
///     SinDegrees = Sin(degrees * (PI / 180))
/// End Function
/// ```
///
/// ### Pattern 3: GenerateSineWave
/// Generate array of sine wave values
/// ```vb
/// Function GenerateSineWave(samples As Integer, amplitude As Double, _
///                           frequency As Double) As Double()
///     Dim result() As Double
///     Dim i As Integer
///     Dim angle As Double
///     Const PI As Double = 3.14159265358979
///     
///     ReDim result(0 To samples - 1)
///     
///     For i = 0 To samples - 1
///         angle = (i / samples) * 2 * PI * frequency
///         result(i) = amplitude * Sin(angle)
///     Next i
///     
///     GenerateSineWave = result
/// End Function
/// ```
///
/// ### Pattern 4: CircularMotionY
/// Calculate vertical position in circular motion
/// ```vb
/// Function CircularMotionY(centerY As Double, radius As Double, _
///                          angle As Double) As Double
///     ' angle in radians
///     CircularMotionY = centerY + radius * Sin(angle)
/// End Function
/// ```
///
/// ### Pattern 5: OscillatingValue
/// Create oscillating value over time
/// ```vb
/// Function OscillatingValue(time As Double, amplitude As Double, _
///                           frequency As Double, Optional phase As Double = 0) As Double
///     Const PI As Double = 3.14159265358979
///     OscillatingValue = amplitude * Sin(2 * PI * frequency * time + phase)
/// End Function
/// ```
///
/// ### Pattern 6: SineInterpolation
/// Smooth interpolation using sine
/// ```vb
/// Function SineInterpolation(startValue As Double, endValue As Double, _
///                            t As Double) As Double
///     ' t ranges from 0 to 1
///     Dim factor As Double
///     Const PI As Double = 3.14159265358979
///     
///     factor = (1 - Cos(t * PI)) / 2
///     SineInterpolation = startValue + (endValue - startValue) * factor
/// End Function
/// ```
///
/// ### Pattern 7: AngleFromSine
/// Get angle from sine value (inverse sine approximation)
/// ```vb
/// Function ArcSineApprox(sineValue As Double) As Double
///     ' For small angles, asin(x) ≈ x
///     ' For better accuracy, use iterative methods or Atn
///     ' Using Atn for proper arcsin:
///     If Abs(sineValue) >= 1 Then
///         ArcSineApprox = Sgn(sineValue) * Atn(1) * 2
///     Else
///         ArcSineApprox = Atn(sineValue / Sqr(1 - sineValue * sineValue))
///     End If
/// End Function
/// ```
///
/// ### Pattern 8: SineWaveAnalysis
/// Analyze sine wave properties
/// ```vb
/// Sub AnalyzeSineWave(amplitude As Double, frequency As Double, _
///                     ByRef maxVal As Double, ByRef minVal As Double, _
///                     ByRef period As Double)
///     Const PI As Double = 3.14159265358979
///     
///     maxVal = amplitude
///     minVal = -amplitude
///     period = 1 / frequency  ' In seconds or time units
/// End Sub
/// ```
///
/// ### Pattern 9: ProjectileMotionY
/// Calculate vertical position in projectile motion
/// ```vb
/// Function ProjectileY(initialY As Double, velocity As Double, _
///                      angle As Double, time As Double, gravity As Double) As Double
///     ' angle in radians
///     Dim verticalVelocity As Double
///     
///     verticalVelocity = velocity * Sin(angle)
///     ProjectileY = initialY + verticalVelocity * time - 0.5 * gravity * time * time
/// End Function
/// ```
///
/// ### Pattern 10: HarmonicMotion
/// Simple harmonic motion displacement
/// ```vb
/// Function HarmonicDisplacement(amplitude As Double, angularFrequency As Double, _
///                               time As Double, Optional phase As Double = 0) As Double
///     HarmonicDisplacement = amplitude * Sin(angularFrequency * time + phase)
/// End Function
/// ```
///
/// ## Advanced Usage
///
/// ### Example 1: WaveformGenerator Class
/// Generate various waveforms using sine function
/// ```vb
/// ' Class: WaveformGenerator
/// Private Const PI As Double = 3.14159265358979
/// Private m_sampleRate As Long
/// Private m_duration As Double
///
/// Public Sub Initialize(sampleRate As Long, duration As Double)
///     m_sampleRate = sampleRate
///     m_duration = duration
/// End Sub
///
/// Public Function GenerateSineWave(frequency As Double, amplitude As Double) As Double()
///     Dim samples As Long
///     Dim result() As Double
///     Dim i As Long
///     Dim t As Double
///     
///     samples = CLng(m_sampleRate * m_duration)
///     ReDim result(0 To samples - 1)
///     
///     For i = 0 To samples - 1
///         t = i / m_sampleRate
///         result(i) = amplitude * Sin(2 * PI * frequency * t)
///     Next i
///     
///     GenerateSineWave = result
/// End Function
///
/// Public Function GenerateAMWave(carrier As Double, modulator As Double, _
///                                amplitude As Double, modDepth As Double) As Double()
///     ' Amplitude Modulation
///     Dim samples As Long
///     Dim result() As Double
///     Dim i As Long
///     Dim t As Double
///     Dim envelope As Double
///     
///     samples = CLng(m_sampleRate * m_duration)
///     ReDim result(0 To samples - 1)
///     
///     For i = 0 To samples - 1
///         t = i / m_sampleRate
///         envelope = 1 + modDepth * Sin(2 * PI * modulator * t)
///         result(i) = amplitude * envelope * Sin(2 * PI * carrier * t)
///     Next i
///     
///     GenerateAMWave = result
/// End Function
///
/// Public Function GenerateFMWave(carrier As Double, modulator As Double, _
///                                amplitude As Double, modIndex As Double) As Double()
///     ' Frequency Modulation
///     Dim samples As Long
///     Dim result() As Double
///     Dim i As Long
///     Dim t As Double
///     Dim phase As Double
///     
///     samples = CLng(m_sampleRate * m_duration)
///     ReDim result(0 To samples - 1)
///     
///     For i = 0 To samples - 1
///         t = i / m_sampleRate
///         phase = 2 * PI * carrier * t + modIndex * Sin(2 * PI * modulator * t)
///         result(i) = amplitude * Sin(phase)
///     Next i
///     
///     GenerateFMWave = result
/// End Function
///
/// Public Function GenerateHarmonics(fundamental As Double, harmonics As Integer, _
///                                   amplitude As Double) As Double()
///     ' Generate complex tone with harmonics
///     Dim samples As Long
///     Dim result() As Double
///     Dim i As Long
///     Dim h As Integer
///     Dim t As Double
///     Dim value As Double
///     
///     samples = CLng(m_sampleRate * m_duration)
///     ReDim result(0 To samples - 1)
///     
///     For i = 0 To samples - 1
///         t = i / m_sampleRate
///         value = 0
///         
///         For h = 1 To harmonics
///             value = value + (amplitude / h) * Sin(2 * PI * fundamental * h * t)
///         Next h
///         
///         result(i) = value
///     Next i
///     
///     GenerateHarmonics = result
/// End Function
/// ```
///
/// ### Example 2: CircularMotion Module
/// Calculate circular and elliptical motion using trigonometry
/// ```vb
/// ' Module: CircularMotion
/// Private Const PI As Double = 3.14159265358979
///
/// Public Sub GetCircularPosition(centerX As Double, centerY As Double, _
///                                radius As Double, angle As Double, _
///                                ByRef x As Double, ByRef y As Double)
///     ' angle in radians
///     x = centerX + radius * Cos(angle)
///     y = centerY + radius * Sin(angle)
/// End Sub
///
/// Public Sub GetEllipticalPosition(centerX As Double, centerY As Double, _
///                                  radiusX As Double, radiusY As Double, _
///                                  angle As Double, ByRef x As Double, ByRef y As Double)
///     ' angle in radians
///     x = centerX + radiusX * Cos(angle)
///     y = centerY + radiusY * Sin(angle)
/// End Sub
///
/// Public Function CalculateAngularVelocity(rpm As Double) As Double
///     ' Convert revolutions per minute to radians per second
///     CalculateAngularVelocity = (rpm / 60) * 2 * PI
/// End Function
///
/// Public Sub AnimateCircularMotion(centerX As Double, centerY As Double, _
///                                  radius As Double, angularVelocity As Double, _
///                                  time As Double, ByRef x As Double, ByRef y As Double)
///     Dim angle As Double
///     angle = angularVelocity * time
///     
///     x = centerX + radius * Cos(angle)
///     y = centerY + radius * Sin(angle)
/// End Sub
///
/// Public Function CalculateTangentialVelocity(radius As Double, _
///                                             angularVelocity As Double) As Double
///     ' v = r * ω
///     CalculateTangentialVelocity = radius * angularVelocity
/// End Function
///
/// Public Sub GetVelocityComponents(speed As Double, angle As Double, _
///                                  ByRef vx As Double, ByRef vy As Double)
///     ' angle in radians from horizontal
///     vx = speed * Cos(angle)
///     vy = speed * Sin(angle)
/// End Sub
/// ```
///
/// ### Example 3: PhysicsSimulator Class
/// Simulate physics using trigonometric functions
/// ```vb
/// ' Class: PhysicsSimulator
/// Private Const PI As Double = 3.14159265358979
/// Private Const GRAVITY As Double = 9.81  ' m/s²
///
/// Public Function CalculateRange(velocity As Double, angle As Double) As Double
///     ' Projectile range formula: R = v² * sin(2θ) / g
///     ' angle in radians
///     CalculateRange = (velocity * velocity * Sin(2 * angle)) / GRAVITY
/// End Function
///
/// Public Function CalculateMaxHeight(velocity As Double, angle As Double) As Double
///     ' Max height: H = (v * sin(θ))² / (2g)
///     Dim verticalVelocity As Double
///     verticalVelocity = velocity * Sin(angle)
///     CalculateMaxHeight = (verticalVelocity * verticalVelocity) / (2 * GRAVITY)
/// End Function
///
/// Public Function CalculateTimeOfFlight(velocity As Double, angle As Double) As Double
///     ' Time of flight: T = 2 * v * sin(θ) / g
///     CalculateTimeOfFlight = (2 * velocity * Sin(angle)) / GRAVITY
/// End Function
///
/// Public Sub GetProjectilePosition(velocity As Double, angle As Double, _
///                                  time As Double, ByRef x As Double, ByRef y As Double)
///     ' angle in radians
///     Dim vx As Double, vy As Double
///     
///     vx = velocity * Cos(angle)
///     vy = velocity * Sin(angle)
///     
///     x = vx * time
///     y = vy * time - 0.5 * GRAVITY * time * time
/// End Sub
///
/// Public Function CalculatePendulumDisplacement(length As Double, angle0 As Double, _
///                                               time As Double) As Double
///     ' Small angle approximation
///     ' angle(t) = angle0 * cos(ωt) where ω = sqrt(g/L)
///     Dim omega As Double
///     omega = Sqr(GRAVITY / length)
///     
///     ' For small angles, displacement ≈ L * θ
///     CalculatePendulumDisplacement = length * angle0 * Cos(omega * time)
/// End Function
///
/// Public Function CalculateInclinedPlaneForce(mass As Double, angle As Double) As Double
///     ' Force down incline: F = m * g * sin(θ)
///     ' angle in radians
///     CalculateInclinedPlaneForce = mass * GRAVITY * Sin(angle)
/// End Function
/// ```
///
/// ### Example 4: GraphicsHelper Module
/// Graphics and animation helpers using trigonometry
/// ```vb
/// ' Module: GraphicsHelper
/// Private Const PI As Double = 3.14159265358979
///
/// Public Function RotatePointX(x As Double, y As Double, angle As Double, _
///                              centerX As Double, centerY As Double) As Double
///     ' Rotate point around center, return new X
///     ' angle in radians
///     Dim dx As Double, dy As Double
///     
///     dx = x - centerX
///     dy = y - centerY
///     
///     RotatePointX = centerX + dx * Cos(angle) - dy * Sin(angle)
/// End Function
///
/// Public Function RotatePointY(x As Double, y As Double, angle As Double, _
///                              centerX As Double, centerY As Double) As Double
///     ' Rotate point around center, return new Y
///     ' angle in radians
///     Dim dx As Double, dy As Double
///     
///     dx = x - centerX
///     dy = y - centerY
///     
///     RotatePointY = centerY + dx * Sin(angle) + dy * Cos(angle)
/// End Function
///
/// Public Function CreatePulseEffect(time As Double, frequency As Double) As Double
///     ' Create pulsing effect (0 to 1)
///     CreatePulseEffect = (Sin(2 * PI * frequency * time) + 1) / 2
/// End Function
///
/// Public Function CreateFadeInOut(time As Double, duration As Double) As Double
///     ' Smooth fade in and out using sine
///     Dim t As Double
///     t = (time / duration) * PI
///     CreateFadeInOut = Sin(t)
/// End Function
///
/// Public Function EaseInOutSine(t As Double) As Double
///     ' Easing function using sine (t from 0 to 1)
///     EaseInOutSine = -(Cos(PI * t) - 1) / 2
/// End Function
///
/// Public Sub DrawSineWave(picBox As Object, amplitude As Double, _
///                         frequency As Double, Optional phase As Double = 0)
///     Dim x As Integer
///     Dim y As Double
///     Dim prevX As Integer, prevY As Integer
///     Dim width As Integer
///     
///     width = picBox.ScaleWidth
///     
///     For x = 0 To width
///         y = amplitude * Sin(2 * PI * frequency * (x / width) + phase)
///         y = picBox.ScaleHeight / 2 - y  ' Flip Y axis
///         
///         If x > 0 Then
///             picBox.Line (prevX, prevY)-(x, y)
///         End If
///         
///         prevX = x
///         prevY = y
///     Next x
/// End Sub
/// ```
///
/// ## Error Handling
///
/// The Sin function can generate the following errors:
///
/// - **Error 13** (Type mismatch): Argument cannot be interpreted as numeric
/// - **Error 5** (Invalid procedure call): In rare cases with invalid input
///
/// Error handling example:
/// ```vb
/// On Error Resume Next
/// result = Sin(angle)
/// If Err.Number <> 0 Then
///     MsgBox "Error calculating sine: " & Err.Description
/// End If
/// ```
///
/// ## Performance Considerations
///
/// - Sin is a relatively fast mathematical function
/// - Uses hardware FPU for calculation when available
/// - For repeated calculations with same angles, consider caching results
/// - Lookup tables can be faster for real-time applications with limited angle sets
/// - Modern CPUs execute Sin very quickly (microseconds)
///
/// ## Best Practices
///
/// 1. **Use Radians**: Remember Sin takes radians, not degrees
/// 2. **Convert Carefully**: Use consistent conversion factor for degrees↔radians
/// 3. **Cache Pi**: Define PI as a constant rather than calculating repeatedly
/// 4. **Range Awareness**: Sin always returns -1 to 1
/// 5. **Precision**: Be aware of floating-point precision limits
/// 6. **Angle Normalization**: For large angles, consider normalizing to 0-2π
/// 7. **Avoid Division**: Use multiplication by inverse when possible
/// 8. **Test Edge Cases**: Test with 0, π/2, π, 3π/2, 2π
/// 9. **Document Units**: Always document whether angles are in degrees or radians
/// 10. **Combine Functions**: Use with Cos, Tan for complete trigonometric operations
///
/// ## Comparison with Related Functions
///
/// | Function | Input (radians) | Output Range | Description |
/// |----------|-----------------|--------------|-------------|
/// | Sin | angle | -1 to 1 | Sine of angle |
/// | Cos | angle | -1 to 1 | Cosine of angle |
/// | Tan | angle | -∞ to +∞ | Tangent of angle |
/// | Atn | ratio | -π/2 to π/2 | Arctangent (inverse tangent) |
/// | Sqr | number ≥ 0 | ≥ 0 | Square root |
///
/// ## Platform Considerations
///
/// - Available in VB6, VBA (all versions)
/// - Uses system math library
/// - Precision depends on Double data type (IEEE 754)
/// - Results consistent across Windows versions
/// - Very small return values near multiples of π due to floating-point precision
///
/// ## Limitations
///
/// - Input must be in radians (no built-in degree support)
/// - Floating-point precision limits (≈15-17 decimal digits)
/// - Sin(π) returns very small number, not exactly 0
/// - Large angle values may accumulate rounding errors
/// - No complex number support
/// - No automatic angle normalization
///
/// ## Related Functions
///
/// - `Cos`: Returns the cosine of an angle in radians
/// - `Tan`: Returns the tangent of an angle in radians
/// - `Atn`: Returns the arctangent of a number in radians
/// - `Sqr`: Returns the square root (used in inverse sine calculations)
/// - `Abs`: Returns absolute value (useful for angle normalization)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_sin_basic() {
        let source = r#"
Sub Test()
    Dim result As Double
    result = Sin(1.5708)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("result"));
    }

    #[test]
    fn test_sin_with_pi() {
        let source = r#"
Sub Test()
    Dim pi As Double
    Dim sineValue As Double
    pi = Atn(1) * 4
    sineValue = Sin(pi / 2)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("sineValue"));
    }

    #[test]
    fn test_sin_if_statement() {
        let source = r#"
Sub Test()
    If Sin(angle) > 0.5 Then
        MsgBox "Greater than 0.5"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
    }

    #[test]
    fn test_sin_function_return() {
        let source = r#"
Function CalculateSine(angle As Double) As Double
    CalculateSine = Sin(angle)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("CalculateSine"));
    }

    #[test]
    fn test_sin_variable_assignment() {
        let source = r#"
Sub Test()
    Dim y As Double
    y = Sin(x)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("y"));
    }

    #[test]
    fn test_sin_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Sine value: " & Sin(angle)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("MsgBox"));
    }

    #[test]
    fn test_sin_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print Sin(1.0472)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("Debug"));
    }

    #[test]
    fn test_sin_select_case() {
        let source = r#"
Sub Test()
    Select Case Sin(angle)
        Case Is > 0
            MsgBox "Positive"
        Case Is < 0
            MsgBox "Negative"
        Case Else
            MsgBox "Zero"
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
    }

    #[test]
    fn test_sin_class_usage() {
        let source = r#"
Class TrigCalculator
    Public Function GetSine(angle As Double) As Double
        GetSine = Sin(angle)
    End Function
End Class
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("GetSine"));
    }

    #[test]
    fn test_sin_with_statement() {
        let source = r#"
Sub Test()
    With Calculator
        Dim s As Double
        s = Sin(angle)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("s"));
    }

    #[test]
    fn test_sin_elseif() {
        let source = r#"
Sub Test()
    If Sin(a) > 0.9 Then
        MsgBox "High"
    ElseIf Sin(a) > 0.5 Then
        MsgBox "Medium"
    Else
        MsgBox "Low"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
    }

    #[test]
    fn test_sin_for_loop() {
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 0 To 360
        Debug.Print Sin(i)
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
    }

    #[test]
    fn test_sin_do_while() {
        let source = r#"
Sub Test()
    Do While Sin(angle) < 1
        angle = angle + 0.1
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
    }

    #[test]
    fn test_sin_do_until() {
        let source = r#"
Sub Test()
    Do Until Sin(x) > threshold
        x = x + step
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
    }

    #[test]
    fn test_sin_while_wend() {
        let source = r#"
Sub Test()
    While Abs(Sin(angle)) > 0.01
        angle = angle - 0.1
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
    }

    #[test]
    fn test_sin_parentheses() {
        let source = r#"
Sub Test()
    Dim value As Double
    value = (Sin(a) + Sin(b)) / 2
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("value"));
    }

    #[test]
    fn test_sin_iif() {
        let source = r#"
Sub Test()
    Dim msg As String
    msg = IIf(Sin(angle) > 0, "Positive", "Non-positive")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("msg"));
    }

    #[test]
    fn test_sin_array_assignment() {
        let source = r#"
Sub Test()
    Dim waveform(100) As Double
    waveform(0) = Sin(angle)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("waveform"));
    }

    #[test]
    fn test_sin_property_assignment() {
        let source = r#"
Class Point
    Public Y As Double
End Class

Sub Test()
    Dim pt As New Point
    pt.Y = Sin(angle)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
    }

    #[test]
    fn test_sin_function_argument() {
        let source = r#"
Sub ProcessValue(v As Double)
End Sub

Sub Test()
    ProcessValue Sin(angle)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("ProcessValue"));
    }

    #[test]
    fn test_sin_concatenation() {
        let source = r#"
Sub Test()
    Dim output As String
    output = "Sine: " & Sin(angle)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("output"));
    }

    #[test]
    fn test_sin_comparison() {
        let source = r#"
Sub Test()
    Dim isPositive As Boolean
    isPositive = (Sin(angle) > 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("isPositive"));
    }

    #[test]
    fn test_sin_arithmetic() {
        let source = r#"
Sub Test()
    Dim amplitude As Double
    Dim wave As Double
    wave = amplitude * Sin(angle)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("wave"));
    }

    #[test]
    fn test_sin_degrees_conversion() {
        let source = r#"
Sub Test()
    Const PI As Double = 3.14159265358979
    Dim degrees As Double
    Dim radians As Double
    Dim result As Double
    degrees = 45
    radians = degrees * (PI / 180)
    result = Sin(radians)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("result"));
    }

    #[test]
    fn test_sin_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Dim s As Double
    s = Sin(inputValue)
    If Err.Number <> 0 Then
        MsgBox "Error"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("s"));
    }

    #[test]
    fn test_sin_on_error_goto() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Dim sineVal As Double
    sineVal = Sin(angle)
    Exit Sub
ErrorHandler:
    MsgBox "Error calculating sine"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("sineVal"));
    }

    #[test]
    fn test_sin_circular_motion() {
        let source = r#"
Sub Test()
    Dim y As Double
    Dim radius As Double
    y = centerY + radius * Sin(angle)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sin"));
        assert!(debug.contains("y"));
    }
}
