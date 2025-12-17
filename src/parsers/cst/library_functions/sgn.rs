/// # Sgn Function
///
/// Returns an Integer indicating the sign of a number.
///
/// ## Syntax
///
/// ```vb
/// Sgn(number)
/// ```
///
/// ## Parameters
///
/// - `number` - Required. Any valid numeric expression.
///
/// ## Return Value
///
/// Returns an Integer indicating the sign of the number:
/// - Returns **-1** if `number` is negative (less than zero)
/// - Returns **0** if `number` is zero
/// - Returns **1** if `number` is positive (greater than zero)
///
/// ## Remarks
///
/// The Sgn function is used to determine the sign (positive, negative, or zero) of a numeric expression.
/// It is particularly useful for:
/// - Determining the direction of change in values
/// - Implementing sign-dependent logic without complex conditionals
/// - Comparing signs of two numbers
/// - Normalizing values to -1, 0, or 1
/// - Mathematical algorithms requiring sign information
///
/// The sign of `number` determines the return value:
/// - If `number` > 0, Sgn returns 1
/// - If `number` = 0, Sgn returns 0
/// - If `number` < 0, Sgn returns -1
///
/// If `number` is Null, Sgn returns Null.
///
/// The Sgn function is often used in combination with Abs to separate magnitude from direction:
/// - `Abs(number)` gives the magnitude (absolute value)
/// - `Sgn(number)` gives the direction (sign)
/// - `number = Abs(number) * Sgn(number)` (reconstruction)
///
/// ## Typical Uses
///
/// 1. **Direction Detection**: Determine if a value is increasing or decreasing
/// 2. **Sign Comparison**: Compare signs of two numbers
/// 3. **Conditional Logic**: Simplify sign-based branching
/// 4. **Mathematical Operations**: Algorithms requiring sign information
/// 5. **Data Validation**: Check if values are positive, negative, or zero
/// 6. **Trend Analysis**: Determine direction of change in time series
/// 7. **Game Logic**: Movement direction, score changes
/// 8. **Financial Calculations**: Profit/loss direction, balance changes
///
/// ## Basic Examples
///
/// ```vb
/// ' Example 1: Basic sign detection
/// Dim result As Integer
/// result = Sgn(10)      ' Returns 1 (positive)
/// result = Sgn(-5.5)    ' Returns -1 (negative)
/// result = Sgn(0)       ' Returns 0 (zero)
/// ```
///
/// ```vb
/// ' Example 2: Determine direction of change
/// Dim oldValue As Double
/// Dim newValue As Double
/// Dim direction As Integer
///
/// oldValue = 100
/// newValue = 120
/// direction = Sgn(newValue - oldValue)  ' Returns 1 (increasing)
///
/// If direction = 1 Then
///     MsgBox "Value increased"
/// ElseIf direction = -1 Then
///     MsgBox "Value decreased"
/// Else
///     MsgBox "No change"
/// End If
/// ```
///
/// ```vb
/// ' Example 3: Compare signs of two numbers
/// Dim a As Double
/// Dim b As Double
///
/// a = -15
/// b = -20
///
/// If Sgn(a) = Sgn(b) Then
///     MsgBox "Same sign"
/// Else
///     MsgBox "Different signs"
/// End If
/// ```
///
/// ```vb
/// ' Example 4: Normalize to unit sign
/// Dim value As Double
/// Dim unitSign As Integer
///
/// value = 42.7
/// unitSign = Sgn(value)  ' Returns 1
/// ' Now unitSign can be used as a multiplier: 1, 0, or -1
/// ```
///
/// ## Common Patterns
///
/// ### Pattern 1: GetChangeDirection
/// Determine if value increased, decreased, or stayed same
/// ```vb
/// Function GetChangeDirection(oldVal As Double, newVal As Double) As String
///     Dim direction As Integer
///     direction = Sgn(newVal - oldVal)
///     
///     Select Case direction
///         Case 1
///             GetChangeDirection = "Increased"
///         Case -1
///             GetChangeDirection = "Decreased"
///         Case 0
///             GetChangeDirection = "No Change"
///     End Select
/// End Function
/// ```
///
/// ### Pattern 2: SameSign
/// Check if two numbers have the same sign
/// ```vb
/// Function SameSign(a As Double, b As Double) As Boolean
///     ' Both zero, or both have same non-zero sign
///     If a = 0 And b = 0 Then
///         SameSign = True
///     Else
///         SameSign = (Sgn(a) = Sgn(b))
///     End If
/// End Function
/// ```
///
/// ### Pattern 3: OppositeSign
/// Check if two numbers have opposite signs
/// ```vb
/// Function OppositeSign(a As Double, b As Double) As Boolean
///     OppositeSign = (Sgn(a) = -Sgn(b)) And (a <> 0) And (b <> 0)
/// End Function
/// ```
///
/// ### Pattern 4: SignMultiplier
/// Get sign as multiplier for calculations
/// ```vb
/// Function SignMultiplier(value As Double) As Integer
///     SignMultiplier = Sgn(value)
///     ' Returns: -1, 0, or 1 which can be used in calculations
/// End Function
/// ```
///
/// ### Pattern 5: ClampToSign
/// Ensure value has specific sign
/// ```vb
/// Function ClampToSign(value As Double, requiredSign As Integer) As Double
///     If Sgn(value) <> requiredSign Then
///         ClampToSign = Abs(value) * requiredSign
///     Else
///         ClampToSign = value
///     End If
/// End Function
/// ```
///
/// ### Pattern 6: SignString
/// Convert sign to string representation
/// ```vb
/// Function SignString(value As Double) As String
///     Select Case Sgn(value)
///         Case 1
///             SignString = "+"
///         Case -1
///             SignString = "-"
///         Case 0
///             SignString = "0"
///     End Select
/// End Function
/// ```
///
/// ### Pattern 7: CompareBySign
/// Three-way comparison using sign
/// ```vb
/// Function CompareBySign(a As Double, b As Double) As Integer
///     ' Returns: -1 if a < b, 0 if a = b, 1 if a > b
///     CompareBySign = Sgn(a - b)
/// End Function
/// ```
///
/// ### Pattern 8: CountBySign
/// Count positive, negative, and zero values
/// ```vb
/// Sub CountBySign(arr() As Double, ByRef positive As Long, _
///                 ByRef negative As Long, ByRef zero As Long)
///     Dim i As Long
///     positive = 0
///     negative = 0
///     zero = 0
///     
///     For i = LBound(arr) To UBound(arr)
///         Select Case Sgn(arr(i))
///             Case 1
///                 positive = positive + 1
///             Case -1
///                 negative = negative + 1
///             Case 0
///                 zero = zero + 1
///         End Select
///     Next i
/// End Sub
/// ```
///
/// ### Pattern 9: TrendDirection
/// Determine overall trend in series
/// ```vb
/// Function TrendDirection(values() As Double) As String
///     Dim i As Long
///     Dim changes As Long
///     Dim positiveChanges As Long
///     Dim negativeChanges As Long
///     
///     For i = LBound(values) + 1 To UBound(values)
///         Dim change As Integer
///         change = Sgn(values(i) - values(i - 1))
///         
///         If change = 1 Then positiveChanges = positiveChanges + 1
///         If change = -1 Then negativeChanges = negativeChanges + 1
///     Next i
///     
///     If positiveChanges > negativeChanges Then
///         TrendDirection = "Upward"
///     ElseIf negativeChanges > positiveChanges Then
///         TrendDirection = "Downward"
///     Else
///         TrendDirection = "Stable"
///     End If
/// End Function
/// ```
///
/// ### Pattern 10: ApplySignTo
/// Apply sign of one number to another
/// ```vb
/// Function ApplySignTo(magnitude As Double, signSource As Double) As Double
///     ' Take absolute value of magnitude and apply sign from signSource
///     ApplySignTo = Abs(magnitude) * Sgn(signSource)
/// End Function
/// ```
///
/// ## Advanced Usage
///
/// ### Example 1: ChangeAnalyzer Class
/// Analyze changes in data series with trend detection
/// ```vb
/// ' Class: ChangeAnalyzer
/// Private m_values() As Double
/// Private m_count As Long
///
/// Public Sub Initialize(initialCapacity As Long)
///     ReDim m_values(1 To initialCapacity)
///     m_count = 0
/// End Sub
///
/// Public Sub AddValue(value As Double)
///     m_count = m_count + 1
///     
///     If m_count > UBound(m_values) Then
///         ReDim Preserve m_values(1 To m_count * 2)
///     End If
///     
///     m_values(m_count) = value
/// End Sub
///
/// Public Function GetChangeDirection(index As Long) As Integer
///     ' Returns direction of change at index
///     If index < 2 Or index > m_count Then
///         GetChangeDirection = 0
///     Else
///         GetChangeDirection = Sgn(m_values(index) - m_values(index - 1))
///     End If
/// End Function
///
/// Public Function GetOverallTrend() As String
///     Dim i As Long
///     Dim upCount As Long
///     Dim downCount As Long
///     
///     For i = 2 To m_count
///         Select Case Sgn(m_values(i) - m_values(i - 1))
///             Case 1
///                 upCount = upCount + 1
///             Case -1
///                 downCount = downCount + 1
///         End Select
///     Next i
///     
///     If upCount > downCount Then
///         GetOverallTrend = "Upward Trend"
///     ElseIf downCount > upCount Then
///         GetOverallTrend = "Downward Trend"
///     Else
///         GetOverallTrend = "No Clear Trend"
///     End If
/// End Function
///
/// Public Function GetConsecutiveChanges() As Long
///     ' Find longest sequence of consecutive changes in same direction
///     Dim i As Long
///     Dim currentDirection As Integer
///     Dim currentCount As Long
///     Dim maxCount As Long
///     
///     If m_count < 2 Then
///         GetConsecutiveChanges = 0
///         Exit Function
///     End If
///     
///     currentDirection = Sgn(m_values(2) - m_values(1))
///     currentCount = 1
///     maxCount = 1
///     
///     For i = 3 To m_count
///         Dim direction As Integer
///         direction = Sgn(m_values(i) - m_values(i - 1))
///         
///         If direction = currentDirection And direction <> 0 Then
///             currentCount = currentCount + 1
///             If currentCount > maxCount Then maxCount = currentCount
///         Else
///             currentDirection = direction
///             currentCount = 1
///         End If
///     Next i
///     
///     GetConsecutiveChanges = maxCount
/// End Function
///
/// Public Function GetTrendStrength() As Double
///     ' Returns value between -1 and 1 indicating trend strength
///     ' 1 = strong upward, -1 = strong downward, 0 = no trend
///     Dim i As Long
///     Dim sumSign As Long
///     Dim changes As Long
///     
///     For i = 2 To m_count
///         Dim sign As Integer
///         sign = Sgn(m_values(i) - m_values(i - 1))
///         If sign <> 0 Then
///             sumSign = sumSign + sign
///             changes = changes + 1
///         End If
///     Next i
///     
///     If changes > 0 Then
///         GetTrendStrength = CDbl(sumSign) / CDbl(changes)
///     Else
///         GetTrendStrength = 0
///     End If
/// End Function
/// ```
///
/// ### Example 2: SignComparator Module
/// Compare and analyze signs in numeric data
/// ```vb
/// ' Module: SignComparator
///
/// Public Function AllSameSign(values() As Double) As Boolean
///     ' Check if all values have the same sign
///     Dim i As Long
///     Dim firstSign As Integer
///     
///     If UBound(values) < LBound(values) Then
///         AllSameSign = True
///         Exit Function
///     End If
///     
///     ' Get first non-zero sign
///     firstSign = 0
///     For i = LBound(values) To UBound(values)
///         firstSign = Sgn(values(i))
///         If firstSign <> 0 Then Exit For
///     Next i
///     
///     ' Check all others
///     For i = LBound(values) To UBound(values)
///         If Sgn(values(i)) <> 0 And Sgn(values(i)) <> firstSign Then
///             AllSameSign = False
///             Exit Function
///         End If
///     Next i
///     
///     AllSameSign = True
/// End Function
///
/// Public Function GetSignCounts(values() As Double) As String
///     Dim i As Long
///     Dim posCount As Long
///     Dim negCount As Long
///     Dim zeroCount As Long
///     
///     For i = LBound(values) To UBound(values)
///         Select Case Sgn(values(i))
///             Case 1
///                 posCount = posCount + 1
///             Case -1
///                 negCount = negCount + 1
///             Case 0
///                 zeroCount = zeroCount + 1
///         End Select
///     Next i
///     
///     GetSignCounts = "Positive: " & posCount & _
///                     ", Negative: " & negCount & _
///                     ", Zero: " & zeroCount
/// End Function
///
/// Public Function AlternatingSign(values() As Double) As Boolean
///     ' Check if signs alternate (ignoring zeros)
///     Dim i As Long
///     Dim lastSign As Integer
///     
///     For i = LBound(values) To UBound(values)
///         Dim currentSign As Integer
///         currentSign = Sgn(values(i))
///         
///         If currentSign <> 0 Then
///             If lastSign <> 0 And currentSign = lastSign Then
///                 AlternatingSign = False
///                 Exit Function
///             End If
///             lastSign = currentSign
///         End If
///     Next i
///     
///     AlternatingSign = True
/// End Function
///
/// Public Function SignTransitions(values() As Double) As Long
///     ' Count how many times the sign changes
///     Dim i As Long
///     Dim transitions As Long
///     Dim lastSign As Integer
///     
///     For i = LBound(values) To UBound(values)
///         Dim currentSign As Integer
///         currentSign = Sgn(values(i))
///         
///         If currentSign <> 0 Then
///             If lastSign <> 0 And currentSign <> lastSign Then
///                 transitions = transitions + 1
///             End If
///             lastSign = currentSign
///         End If
///     Next i
///     
///     SignTransitions = transitions
/// End Function
/// ```
///
/// ### Example 3: DirectionIndicator Class
/// Track and display directional changes with symbols
/// ```vb
/// ' Class: DirectionIndicator
/// Private m_lastValue As Double
/// Private m_initialized As Boolean
///
/// Public Sub SetInitialValue(value As Double)
///     m_lastValue = value
///     m_initialized = True
/// End Sub
///
/// Public Function UpdateAndGetSymbol(newValue As Double) As String
///     If Not m_initialized Then
///         m_lastValue = newValue
///         m_initialized = True
///         UpdateAndGetSymbol = "―"  ' Neutral symbol
///         Exit Function
///     End If
///     
///     Dim direction As Integer
///     direction = Sgn(newValue - m_lastValue)
///     
///     Select Case direction
///         Case 1
///             UpdateAndGetSymbol = "▲"  ' Up arrow
///         Case -1
///             UpdateAndGetSymbol = "▼"  ' Down arrow
///         Case 0
///             UpdateAndGetSymbol = "―"  ' Neutral
///     End Select
///     
///     m_lastValue = newValue
/// End Function
///
/// Public Function GetDirection() As String
///     ' Not implemented - would require storing current value
///     GetDirection = "N/A"
/// End Function
///
/// Public Function GetChangeText(newValue As Double) As String
///     If Not m_initialized Then
///         GetChangeText = "Initial Value"
///     Else
///         Dim change As Double
///         Dim direction As Integer
///         
///         change = newValue - m_lastValue
///         direction = Sgn(change)
///         
///         Select Case direction
///             Case 1
///                 GetChangeText = "Increased by " & Abs(change)
///             Case -1
///                 GetChangeText = "Decreased by " & Abs(change)
///             Case 0
///                 GetChangeText = "No Change"
///         End Select
///     End If
/// End Function
///
/// Public Sub Reset()
///     m_initialized = False
///     m_lastValue = 0
/// End Sub
/// ```
///
/// ### Example 4: MathSignHelper Module
/// Mathematical operations using sign function
/// ```vb
/// ' Module: MathSignHelper
///
/// Public Function CopySign(magnitude As Double, signSource As Double) As Double
///     ' Copy the sign from signSource to magnitude
///     ' Similar to copysign() in C math library
///     CopySign = Abs(magnitude) * Sgn(signSource)
/// End Function
///
/// Public Function RoundAwayFromZero(value As Double) As Long
///     ' Round away from zero (ceiling for positive, floor for negative)
///     Dim sign As Integer
///     sign = Sgn(value)
///     
///     If sign >= 0 Then
///         RoundAwayFromZero = -Int(-value)  ' Ceiling
///     Else
///         RoundAwayFromZero = Int(value)    ' Floor
///     End If
/// End Function
///
/// Public Function RoundTowardZero(value As Double) As Long
///     ' Round toward zero (floor for positive, ceiling for negative)
///     RoundTowardZero = Fix(value)
/// End Function
///
/// Public Function StepInDirection(value As Double, stepSize As Double, _
///                                 direction As Integer) As Double
///     ' Step in the direction indicated by sign (-1, 0, or 1)
///     StepInDirection = value + (Abs(stepSize) * direction)
/// End Function
///
/// Public Function Clamp(value As Double, minVal As Double, maxVal As Double) As Double
///     ' Clamp value between min and max
///     If value < minVal Then
///         Clamp = minVal
///     ElseIf value > maxVal Then
///         Clamp = maxVal
///     Else
///         Clamp = value
///     End If
/// End Function
///
/// Public Function SignedMin(a As Double, b As Double) As Double
///     ' Return the value with smaller absolute value, preserving sign
///     If Abs(a) < Abs(b) Then
///         SignedMin = a
///     Else
///         SignedMin = b
///     End If
/// End Function
///
/// Public Function SignedMax(a As Double, b As Double) As Double
///     ' Return the value with larger absolute value, preserving sign
///     If Abs(a) > Abs(b) Then
///         SignedMax = a
///     Else
///         SignedMax = b
///     End If
/// End Function
///
/// Public Function CompareNumbers(a As Double, b As Double) As Integer
///     ' Three-way comparison: -1 if a<b, 0 if a=b, 1 if a>b
///     CompareNumbers = Sgn(a - b)
/// End Function
/// ```
///
/// ## Error Handling
///
/// The Sgn function can generate the following errors:
///
/// - **Error 13** (Type mismatch): Argument cannot be interpreted as numeric
/// - **Error 94** (Invalid use of Null): If Null is passed and not properly handled
///
/// Always use error handling when working with user input or uncertain data:
/// ```vb
/// On Error Resume Next
/// result = Sgn(userValue)
/// If Err.Number <> 0 Then
///     MsgBox "Invalid numeric value"
/// End If
/// ```
///
/// ## Performance Considerations
///
/// - Sgn is very fast (simple comparison operation)
/// - Much faster than If...Then...Else chains for sign checking
/// - Use Sgn instead of multiple comparisons when appropriate
/// - No performance penalty for any numeric type
/// - Ideal for use in tight loops
///
/// ## Best Practices
///
/// 1. **Use for Comparisons**: Prefer `Sgn(a - b)` over complex If statements for three-way comparison
/// 2. **Handle Zero Case**: Remember that zero has its own return value (0)
/// 3. **Type Safety**: Ensure argument is numeric to avoid Type Mismatch error
/// 4. **Null Handling**: Check for Null if input might be Null
/// 5. **Sign vs. Value**: Don't confuse sign with actual value
/// 6. **Combine with Abs**: Use together to separate magnitude and direction
/// 7. **Clear Intent**: Use Sgn to make sign-dependent logic more readable
/// 8. **Avoid Redundancy**: Don't use Sgn(Abs(x)) - always returns 1 or 0
/// 9. **Consider Zero**: Zero is neither positive nor negative
/// 10. **Document Usage**: Comment when using Sgn in non-obvious ways
///
/// ## Comparison with Related Functions
///
/// | Function | Purpose | Returns | Zero Handling |
/// |----------|---------|---------|---------------|
/// | Sgn | Get sign | -1, 0, or 1 | Returns 0 |
/// | Abs | Get magnitude | Non-negative number | Returns 0 |
/// | Fix | Truncate to integer | Integer toward zero | Returns 0 |
/// | Int | Round down | Integer (floor) | Returns 0 |
/// | Round | Round to nearest | Rounded number | Returns 0 |
/// | If...Then | Conditional logic | Any value | Requires explicit check |
///
/// ## Platform Considerations
///
/// - Available in VB6, VBA (all versions)
/// - Available in VBScript
/// - Part of core VB language
/// - Consistent behavior across all VB variants
/// - No platform-specific quirks
///
/// ## Limitations
///
/// - Only returns three values: -1, 0, 1
/// - Does not provide magnitude information (use Abs for that)
/// - Cannot distinguish between different magnitudes of same sign
/// - Returns Null if argument is Null (may need special handling)
/// - Not suitable for distinguishing "nearly zero" from actual zero
///
/// ## Related Functions
///
/// - `Abs`: Returns the absolute value (magnitude without sign)
/// - `Fix`: Returns the integer portion of a number, truncating toward zero
/// - `Int`: Returns the integer portion of a number, rounding down
/// - `Round`: Rounds a number to specified decimal places
/// - `IIf`: Conditional function that can implement sign-based logic

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn sgn_basic() {
        let source = r#"
Sub Test()
    Dim result As Integer
    result = Sgn(10)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("result"));
    }

    #[test]
    fn sgn_negative() {
        let source = r#"
Sub Test()
    Dim sign As Integer
    sign = Sgn(-5.5)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("sign"));
    }

    #[test]
    fn sgn_if_statement() {
        let source = r#"
Sub Test()
    If Sgn(value) = 1 Then
        MsgBox "Positive"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
    }

    #[test]
    fn sgn_function_return() {
        let source = r#"
Function GetSign(num As Double) As Integer
    GetSign = Sgn(num)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("GetSign"));
    }

    #[test]
    fn sgn_variable_assignment() {
        let source = r#"
Sub Test()
    Dim direction As Integer
    direction = Sgn(newValue - oldValue)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("direction"));
    }

    #[test]
    fn sgn_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Sign: " & Sgn(value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("MsgBox"));
    }

    #[test]
    fn sgn_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print Sgn(number)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("Debug"));
    }

    #[test]
    fn sgn_select_case() {
        let source = r#"
Sub Test()
    Select Case Sgn(delta)
        Case 1
            MsgBox "Increased"
        Case -1
            MsgBox "Decreased"
        Case 0
            MsgBox "No change"
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
    }

    #[test]
    fn sgn_class_usage() {
        let source = r#"
Class Calculator
    Public Function GetDirection(value As Double) As Integer
        GetDirection = Sgn(value)
    End Function
End Class
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("GetDirection"));
    }

    #[test]
    fn sgn_with_statement() {
        let source = r#"
Sub Test()
    With Calculator
        Dim s As Integer
        s = Sgn(value)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("s"));
    }

    #[test]
    fn sgn_elseif() {
        let source = r#"
Sub Test()
    If Sgn(value) = 1 Then
        MsgBox "Positive"
    ElseIf Sgn(value) = -1 Then
        MsgBox "Negative"
    Else
        MsgBox "Zero"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
    }

    #[test]
    fn sgn_for_loop() {
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 1 To 10
        Debug.Print Sgn(arr(i))
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
    }

    #[test]
    fn sgn_do_while() {
        let source = r#"
Sub Test()
    Do While Sgn(value) = 1
        value = value - 1
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
    }

    #[test]
    fn sgn_do_until() {
        let source = r#"
Sub Test()
    Do Until Sgn(counter) = 0
        counter = counter - 1
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
    }

    #[test]
    fn sgn_while_wend() {
        let source = r#"
Sub Test()
    While Sgn(remaining) > 0
        remaining = remaining - 1
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
    }

    #[test]
    fn sgn_parentheses() {
        let source = r#"
Sub Test()
    Dim result As Integer
    result = (Sgn(a) + Sgn(b))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("result"));
    }

    #[test]
    fn sgn_iif() {
        let source = r#"
Sub Test()
    Dim msg As String
    msg = IIf(Sgn(value) = 1, "Positive", "Not positive")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("msg"));
    }

    #[test]
    fn sgn_array_assignment() {
        let source = r#"
Sub Test()
    Dim signs(10) As Integer
    signs(0) = Sgn(values(0))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("signs"));
    }

    #[test]
    fn sgn_property_assignment() {
        let source = r#"
Class DataPoint
    Public Sign As Integer
End Class

Sub Test()
    Dim pt As New DataPoint
    pt.Sign = Sgn(value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
    }

    #[test]
    fn sgn_function_argument() {
        let source = r#"
Sub ProcessSign(s As Integer)
End Sub

Sub Test()
    ProcessSign Sgn(value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("ProcessSign"));
    }

    #[test]
    fn sgn_concatenation() {
        let source = r#"
Sub Test()
    Dim msg As String
    msg = "Direction: " & Sgn(delta)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("msg"));
    }

    #[test]
    fn sgn_comparison() {
        let source = r#"
Sub Test()
    Dim sameSign As Boolean
    sameSign = (Sgn(a) = Sgn(b))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("sameSign"));
    }

    #[test]
    fn sgn_arithmetic() {
        let source = r#"
Sub Test()
    Dim normalized As Double
    normalized = Abs(value) * Sgn(reference)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("normalized"));
    }

    #[test]
    fn sgn_subtraction() {
        let source = r#"
Sub Test()
    Dim trend As Integer
    trend = Sgn(newVal - oldVal)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("trend"));
    }

    #[test]
    fn sgn_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Dim s As Integer
    s = Sgn(userInput)
    If Err.Number <> 0 Then
        MsgBox "Error"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("s"));
    }

    #[test]
    fn sgn_on_error_goto() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Dim signValue As Integer
    signValue = Sgn(input)
    Exit Sub
ErrorHandler:
    MsgBox "Error getting sign"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("signValue"));
    }

    #[test]
    fn sgn_three_way_comparison() {
        let source = r#"
Sub Test()
    Dim compareResult As Integer
    compareResult = Sgn(value1 - value2)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Sgn"));
        assert!(debug.contains("compareResult"));
    }
}
