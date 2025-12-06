//! # Log Function
//!
//! Returns a Double specifying the natural logarithm of a number.
//!
//! ## Syntax
//!
//! ```vb
//! Log(number)
//! ```
//!
//! ## Parameters
//!
//! - `number` (Required): Double or any valid numeric expression greater than zero
//!   - Must be positive (> 0)
//!   - Cannot be zero or negative
//!   - Error 5 "Invalid procedure call or argument" if number <= 0
//!
//! ## Return Value
//!
//! Returns a Double:
//! - Natural logarithm (base e) of the number
//! - Also known as ln(x) in mathematics
//! - Result can be positive, negative, or zero
//! - Log(1) = 0
//! - Log(e) = 1 where e ≈ 2.71828182845905
//! - For values 0 < x < 1, result is negative
//! - For values x > 1, result is positive
//!
//! ## Remarks
//!
//! The Log function returns the natural logarithm:
//!
//! - Natural logarithm uses base e (Euler's number)
//! - e ≈ 2.71828182845905
//! - Also written as ln(x) in mathematical notation
//! - Inverse operation of Exp function
//! - Log(Exp(x)) = x
//! - Exp(Log(x)) = x (for x > 0)
//! - To calculate logarithms with other bases, use change of base formula
//! - Log base 10: Log(x) / Log(10)
//! - Log base 2: Log(x) / Log(2)
//! - Log base n: Log(x) / Log(n)
//! - Error 5 if argument is zero or negative
//! - Used in scientific and engineering calculations
//! - Common in exponential growth/decay problems
//! - Essential for statistical calculations
//! - Used in information theory (entropy, information content)
//! - Financial calculations (continuous compounding)
//! - Physics (radioactive decay, sound levels)
//! - Can be used to solve exponential equations
//! - Part of VB6's math function library
//! - Available in all VB versions
//!
//! ## Typical Uses
//!
//! 1. **Natural Logarithm**
//!    ```vb
//!    result = Log(10)
//!    ```
//!
//! 2. **Base 10 Logarithm**
//!    ```vb
//!    log10 = Log(x) / Log(10)
//!    ```
//!
//! 3. **Base 2 Logarithm**
//!    ```vb
//!    log2 = Log(x) / Log(2)
//!    ```
//!
//! 4. **Exponential Decay**
//!    ```vb
//!    timeConstant = -1 / Log(decayRate)
//!    ```
//!
//! 5. **Solve for Exponent**
//!    ```vb
//!    exponent = Log(result / initial) / Log(base)
//!    ```
//!
//! 6. **Information Content**
//!    ```vb
//!    bits = -Log(probability) / Log(2)
//!    ```
//!
//! 7. **pH Calculation**
//!    ```vb
//!    pH = -Log(hydrogenIonConcentration) / Log(10)
//!    ```
//!
//! 8. **Continuous Compounding**
//!    ```vb
//!    rate = Log(finalValue / initialValue) / time
//!    ```
//!
//! ## Basic Examples
//!
//! ### Example 1: Natural Logarithm
//! ```vb
//! Dim result As Double
//!
//! result = Log(1)        ' Returns 0
//! result = Log(2.71828)  ' Returns ~1 (ln(e) = 1)
//! result = Log(10)       ' Returns ~2.302585
//! result = Log(100)      ' Returns ~4.605170
//! ```
//!
//! ### Example 2: Base 10 Logarithm
//! ```vb
//! Function Log10(ByVal x As Double) As Double
//!     Log10 = Log(x) / Log(10)
//! End Function
//!
//! ' Usage
//! Dim result As Double
//! result = Log10(100)    ' Returns 2
//! result = Log10(1000)   ' Returns 3
//! result = Log10(10)     ' Returns 1
//! ```
//!
//! ### Example 3: Exponential Growth
//! ```vb
//! ' Calculate doubling time
//! Function DoublingTime(ByVal growthRate As Double) As Double
//!     ' growthRate is the rate per time period (e.g., 0.05 = 5%)
//!     DoublingTime = Log(2) / Log(1 + growthRate)
//! End Function
//!
//! ' Usage
//! Dim years As Double
//! years = DoublingTime(0.07)  ' 7% annual growth
//! MsgBox "Doubling time: " & Format(years, "0.0") & " years"
//! ```
//!
//! ### Example 4: Solve Exponential Equation
//! ```vb
//! ' Solve: base^exponent = result for exponent
//! Function SolveExponent(ByVal base As Double, _
//!                        ByVal result As Double) As Double
//!     If base > 0 And base <> 1 And result > 0 Then
//!         SolveExponent = Log(result) / Log(base)
//!     Else
//!         SolveExponent = 0
//!     End If
//! End Function
//!
//! ' Usage: Solve 2^x = 32
//! Dim x As Double
//! x = SolveExponent(2, 32)  ' Returns 5
//! MsgBox "2^" & x & " = 32"
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Log10 (Base 10 Logarithm)
//! ```vb
//! Function Log10(ByVal x As Double) As Double
//!     If x > 0 Then
//!         Log10 = Log(x) / Log(10)
//!     Else
//!         Err.Raise 5, , "Invalid argument"
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 2: Log2 (Base 2 Logarithm)
//! ```vb
//! Function Log2(ByVal x As Double) As Double
//!     If x > 0 Then
//!         Log2 = Log(x) / Log(2)
//!     Else
//!         Err.Raise 5, , "Invalid argument"
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 3: `LogN` (Logarithm with Any Base)
//! ```vb
//! Function LogN(ByVal x As Double, ByVal base As Double) As Double
//!     If x > 0 And base > 0 And base <> 1 Then
//!         LogN = Log(x) / Log(base)
//!     Else
//!         Err.Raise 5, , "Invalid arguments"
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 4: `SafeLog`
//! ```vb
//! Function SafeLog(ByVal x As Double, _
//!                  Optional ByVal defaultValue As Double = 0) As Double
//!     On Error Resume Next
//!     SafeLog = Log(x)
//!     If Err.Number <> 0 Then
//!         SafeLog = defaultValue
//!         Err.Clear
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 5: `CalculateEntropy`
//! ```vb
//! Function CalculateEntropy(probabilities() As Double) As Double
//!     Dim i As Integer
//!     Dim entropy As Double
//!     Dim p As Double
//!     
//!     entropy = 0
//!     For i = LBound(probabilities) To UBound(probabilities)
//!         p = probabilities(i)
//!         If p > 0 Then
//!             entropy = entropy - p * (Log(p) / Log(2))
//!         End If
//!     Next i
//!     
//!     CalculateEntropy = entropy
//! End Function
//! ```
//!
//! ### Pattern 6: `CalculateHalfLife`
//! ```vb
//! Function CalculateHalfLife(ByVal decayConstant As Double) As Double
//!     If decayConstant > 0 Then
//!         CalculateHalfLife = Log(2) / decayConstant
//!     Else
//!         CalculateHalfLife = 0
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 7: `CalculateDoublingTime`
//! ```vb
//! Function CalculateDoublingTime(ByVal growthRate As Double) As Double
//!     If growthRate > 0 Then
//!         CalculateDoublingTime = Log(2) / Log(1 + growthRate)
//!     Else
//!         CalculateDoublingTime = 0
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 8: `SolveForTime` (Exponential Growth)
//! ```vb
//! Function SolveForTime(ByVal initialValue As Double, _
//!                       ByVal finalValue As Double, _
//!                       ByVal rate As Double) As Double
//!     If initialValue > 0 And finalValue > 0 And rate <> 0 Then
//!         SolveForTime = Log(finalValue / initialValue) / rate
//!     Else
//!         SolveForTime = 0
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 9: `CalculateDecibels`
//! ```vb
//! Function CalculateDecibels(ByVal power As Double, _
//!                            ByVal referencePower As Double) As Double
//!     If power > 0 And referencePower > 0 Then
//!         CalculateDecibels = 10 * (Log(power / referencePower) / Log(10))
//!     Else
//!         CalculateDecibels = 0
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 10: `CalculatePH`
//! ```vb
//! Function CalculatePH(ByVal hydrogenIonConcentration As Double) As Double
//!     If hydrogenIonConcentration > 0 Then
//!         CalculatePH = -(Log(hydrogenIonConcentration) / Log(10))
//!     Else
//!         CalculatePH = 7  ' Neutral pH
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Examples
//!
//! ### Example 1: Logarithm Calculator
//! ```vb
//! ' Module: LogarithmCalculator
//!
//! Public Function NaturalLog(ByVal x As Double) As Double
//!     If x <= 0 Then
//!         Err.Raise 5, "LogarithmCalculator", _
//!                   "Argument must be positive"
//!     End If
//!     NaturalLog = Log(x)
//! End Function
//!
//! Public Function Log10(ByVal x As Double) As Double
//!     If x <= 0 Then
//!         Err.Raise 5, "LogarithmCalculator", _
//!                   "Argument must be positive"
//!     End If
//!     Log10 = Log(x) / Log(10)
//! End Function
//!
//! Public Function Log2(ByVal x As Double) As Double
//!     If x <= 0 Then
//!         Err.Raise 5, "LogarithmCalculator", _
//!                   "Argument must be positive"
//!     End If
//!     Log2 = Log(x) / Log(2)
//! End Function
//!
//! Public Function LogBase(ByVal x As Double, _
//!                         ByVal base As Double) As Double
//!     If x <= 0 Or base <= 0 Or base = 1 Then
//!         Err.Raise 5, "LogarithmCalculator", _
//!                   "Invalid arguments"
//!     End If
//!     LogBase = Log(x) / Log(base)
//! End Function
//!
//! Public Function Antilog(ByVal x As Double) As Double
//!     ' Returns e^x (inverse of Log)
//!     Antilog = Exp(x)
//! End Function
//!
//! Public Function Antilog10(ByVal x As Double) As Double
//!     ' Returns 10^x (inverse of Log10)
//!     Antilog10 = 10 ^ x
//! End Function
//! ```
//!
//! ### Example 2: Exponential Growth Analyzer
//! ```vb
//! ' Class: GrowthAnalyzer
//! Private m_initialValue As Double
//! Private m_currentValue As Double
//! Private m_timeElapsed As Double
//!
//! Public Sub Initialize(ByVal initialValue As Double)
//!     m_initialValue = initialValue
//!     m_currentValue = initialValue
//!     m_timeElapsed = 0
//! End Sub
//!
//! Public Property Let CurrentValue(ByVal value As Double)
//!     m_currentValue = value
//! End Property
//!
//! Public Property Let TimeElapsed(ByVal time As Double)
//!     m_timeElapsed = time
//! End Property
//!
//! Public Property Get GrowthRate() As Double
//!     If m_timeElapsed > 0 And m_initialValue > 0 And m_currentValue > 0 Then
//!         GrowthRate = Log(m_currentValue / m_initialValue) / m_timeElapsed
//!     Else
//!         GrowthRate = 0
//!     End If
//! End Property
//!
//! Public Property Get DoublingTime() As Double
//!     Dim rate As Double
//!     rate = GrowthRate
//!     
//!     If rate > 0 Then
//!         DoublingTime = Log(2) / rate
//!     Else
//!         DoublingTime = 0
//!     End If
//! End Property
//!
//! Public Function ProjectValue(ByVal futureTime As Double) As Double
//!     Dim rate As Double
//!     rate = GrowthRate
//!     
//!     If rate <> 0 Then
//!         ProjectValue = m_initialValue * Exp(rate * futureTime)
//!     Else
//!         ProjectValue = m_initialValue
//!     End If
//! End Function
//!
//! Public Function TimeToReach(ByVal targetValue As Double) As Double
//!     Dim rate As Double
//!     rate = GrowthRate
//!     
//!     If rate > 0 And m_initialValue > 0 And targetValue > 0 Then
//!         TimeToReach = Log(targetValue / m_initialValue) / rate
//!     Else
//!         TimeToReach = 0
//!     End If
//! End Function
//! ```
//!
//! ### Example 3: Sound Level Calculator
//! ```vb
//! ' Module: SoundLevelCalculator
//! Private Const REFERENCE_PRESSURE As Double = 0.00002  ' 20 micropascals
//!
//! Public Function CalculateDecibels(ByVal pressure As Double) As Double
//!     If pressure > 0 Then
//!         CalculateDecibels = 20 * (Log(pressure / REFERENCE_PRESSURE) / Log(10))
//!     Else
//!         CalculateDecibels = 0
//!     End If
//! End Function
//!
//! Public Function CombineSoundLevels(levels() As Double) As Double
//!     Dim i As Integer
//!     Dim sumPressures As Double
//!     Dim pressure As Double
//!     
//!     sumPressures = 0
//!     For i = LBound(levels) To UBound(levels)
//!         ' Convert dB to pressure, sum, then convert back
//!         pressure = REFERENCE_PRESSURE * Exp(levels(i) * Log(10) / 20)
//!         sumPressures = sumPressures + pressure * pressure
//!     Next i
//!     
//!     If sumPressures > 0 Then
//!         CombineSoundLevels = 10 * (Log(sumPressures) / Log(10)) + _
//!                             20 * (Log(REFERENCE_PRESSURE) / Log(10))
//!     Else
//!         CombineSoundLevels = 0
//!     End If
//! End Function
//!
//! Public Function CalculateDistance(ByVal soundLevelAtSource As Double, _
//!                                   ByVal soundLevelAtDistance As Double, _
//!                                   ByVal knownDistance As Double) As Double
//!     Dim ratio As Double
//!     
//!     ' Sound decreases by 6 dB when distance doubles
//!     ratio = (soundLevelAtSource - soundLevelAtDistance) / 6
//!     CalculateDistance = knownDistance * (2 ^ ratio)
//! End Function
//! ```
//!
//! ### Example 4: Financial Calculator
//! ```vb
//! ' Module: FinancialCalculator
//!
//! Public Function CalculateContinuousGrowthRate(ByVal initialValue As Double, _
//!                                                ByVal finalValue As Double, _
//!                                                ByVal years As Double) As Double
//!     If initialValue > 0 And finalValue > 0 And years > 0 Then
//!         CalculateContinuousGrowthRate = Log(finalValue / initialValue) / years
//!     Else
//!         CalculateContinuousGrowthRate = 0
//!     End If
//! End Function
//!
//! Public Function YearsToDouble(ByVal annualRate As Double) As Double
//!     If annualRate > 0 Then
//!         YearsToDouble = Log(2) / Log(1 + annualRate)
//!     Else
//!         YearsToDouble = 0
//!     End If
//! End Function
//!
//! Public Function EffectiveRate(ByVal nominalRate As Double, _
//!                               ByVal compoundingPeriods As Integer) As Double
//!     If compoundingPeriods > 0 Then
//!         EffectiveRate = Exp(compoundingPeriods * _
//!                        Log(1 + nominalRate / compoundingPeriods)) - 1
//!     Else
//!         EffectiveRate = 0
//!     End If
//! End Function
//!
//! Public Function CalculateAPY(ByVal principal As Double, _
//!                              ByVal finalAmount As Double, _
//!                              ByVal years As Double) As Double
//!     If principal > 0 And finalAmount > 0 And years > 0 Then
//!         CalculateAPY = Exp(Log(finalAmount / principal) / years) - 1
//!     Else
//!         CalculateAPY = 0
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! ' Error 5: Invalid procedure call or argument
//! On Error Resume Next
//! result = Log(0)
//! If Err.Number = 5 Then
//!     MsgBox "Cannot take log of zero!"
//! End If
//!
//! result = Log(-10)
//! If Err.Number = 5 Then
//!     MsgBox "Cannot take log of negative number!"
//! End If
//!
//! ' Safe log function
//! Function SafeLog(ByVal x As Double) As Variant
//!     On Error Resume Next
//!     SafeLog = Log(x)
//!     If Err.Number <> 0 Then
//!         SafeLog = Null
//!         Err.Clear
//!     End If
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: Log is a built-in processor instruction
//! - **Cache Constants**: Store Log(10), Log(2) if used repeatedly
//! - **Avoid Division by Zero**: Always validate arguments
//! - **Use Exp for Inverse**: Exp(Log(x)) = x is optimized
//!
//! ## Best Practices
//!
//! 1. **Always validate** that argument is positive
//! 2. **Use error handling** for user input
//! 3. **Cache frequently used** logarithms (Log(10), Log(2))
//! 4. **Document the base** when using change of base formula
//! 5. **Use descriptive names** for logarithm wrapper functions
//! 6. **Consider precision** for very large or small numbers
//! 7. **Check for overflow** in calculations
//! 8. **Use constants** for common values (e, pi, etc.)
//! 9. **Validate results** for domain-specific constraints
//! 10. **Comment formulas** explaining mathematical relationships
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Base | Domain |
//! |----------|---------|------|--------|
//! | **Log** | Natural logarithm | e (2.718...) | x > 0 |
//! | **Exp** | Exponential (e^x) | e | All reals |
//! | **Log10** | Common logarithm | 10 | x > 0 |
//! | **Log2** | Binary logarithm | 2 | x > 0 |
//! | **^** (power) | Exponentiation | Variable | Depends |
//!
//! ## Common Logarithm Identities
//!
//! ```vb
//! ' Log properties
//! Log(x * y) = Log(x) + Log(y)
//! Log(x / y) = Log(x) - Log(y)
//! Log(x ^ y) = y * Log(x)
//! Log(1) = 0
//! Log(e) = 1
//!
//! ' Change of base
//! Log_b(x) = Log(x) / Log(b)
//!
//! ' Inverse relationship
//! Exp(Log(x)) = x  (for x > 0)
//! Log(Exp(x)) = x
//! ```
//!
//! ## Platform Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core library
//! - Uses IEEE 754 double-precision floating point
//! - Precision: approximately 15-16 significant digits
//! - Range: 4.94065645841247E-324 to 1.79769313486232E+308
//! - Behavior identical across Windows versions
//! - CPU-level implementation (very fast)
//!
//! ## Limitations
//!
//! - **Positive Arguments Only**: Cannot compute log of zero or negative numbers
//! - **Floating Point Precision**: Subject to rounding errors
//! - **Very Small Numbers**: May lose precision near zero
//! - **Very Large Numbers**: May overflow in calculations
//! - **No Base Parameter**: Must use change of base formula for other bases
//! - **Error for Invalid Input**: Raises Error 5 instead of returning special value
//!
//! ## Related Functions
//!
//! - `Exp`: Returns e raised to a power (inverse of Log)
//! - `Sqr`: Returns square root
//! - `^`: Exponentiation operator
//! - `Abs`: Returns absolute value
//! - `Sgn`: Returns sign of number

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn log_basic() {
        let source = r#"
            Dim result As Double
            result = Log(10)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_variable() {
        let source = r#"
            naturalLog = Log(x)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_base_10() {
        let source = r#"
            log10 = Log(x) / Log(10)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_if_statement() {
        let source = r#"
            If Log(value) > 0 Then
                MsgBox "Greater than 1"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_function_return() {
        let source = r#"
            Function NaturalLog(x As Double) As Double
                NaturalLog = Log(x)
            End Function
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_error_handling() {
        let source = r#"
            On Error Resume Next
            result = Log(x)
            If Err.Number = 5 Then
                MsgBox "Invalid argument"
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_exponential_growth() {
        let source = r#"
            rate = Log(finalValue / initialValue) / time
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_debug_print() {
        let source = r#"
            Debug.Print "Natural log: " & Log(100)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_with_statement() {
        let source = r#"
            With mathCalc
                .Result = Log(value)
            End With
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_select_case() {
        let source = r#"
            Select Case Log(x)
                Case Is > 0
                    MsgBox "Greater than 1"
                Case Is < 0
                    MsgBox "Less than 1"
            End Select
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_elseif() {
        let source = r#"
            If x <= 0 Then
                result = 0
            ElseIf x = 1 Then
                result = 0
            Else
                result = Log(x)
            End If
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_parentheses() {
        let source = r#"
            result = (Log(x))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_iif() {
        let source = r#"
            result = IIf(x > 0, Log(x), 0)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_in_class() {
        let source = r#"
            Private Sub Class_Method()
                m_logValue = Log(m_value)
            End Sub
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_function_argument() {
        let source = r#"
            Call ProcessLog(Log(value))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_property_assignment() {
        let source = r#"
            MyObject.LogValue = Log(x)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_array_assignment() {
        let source = r#"
            logValues(i) = Log(values(i))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_while_wend() {
        let source = r#"
            While x > 1
                x = Exp(Log(x) - 0.1)
            Wend
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_do_while() {
        let source = r#"
            Do While value > threshold
                value = Log(value) + offset
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_for_loop() {
        let source = r#"
            For i = 1 To 10
                results(i) = Log(i)
            Next i
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_msgbox() {
        let source = r#"
            MsgBox "Log(100) = " & Log(100)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_doubling_time() {
        let source = r#"
            doublingTime = Log(2) / Log(1 + growthRate)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_solve_exponent() {
        let source = r#"
            exponent = Log(result) / Log(base)
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_entropy() {
        let source = r#"
            entropy = entropy - p * (Log(p) / Log(2))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_decibels() {
        let source = r#"
            decibels = 10 * (Log(power / reference) / Log(10))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_ph_calculation() {
        let source = r#"
            pH = -(Log(concentration) / Log(10))
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn log_do_until() {
        let source = r#"
            Do Until Abs(Log(x) - target) < tolerance
                x = x + delta
            Loop
        "#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Log"));
        assert!(text.contains("Identifier"));
    }
}
