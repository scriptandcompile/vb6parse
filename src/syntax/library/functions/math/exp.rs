//! # Exp Function
//!
//! Returns e (the base of natural logarithms) raised to a power.
//!
//! ## Syntax
//!
//! ```vb
//! Exp(number)
//! ```
//!
//! ## Parameters
//!
//! - **number**: Required. A Double or any valid numeric expression representing the exponent.
//!
//! ## Return Value
//!
//! Returns a Double representing e raised to the specified power (e^number).
//! The constant e is approximately 2.718282.
//!
//! ## Remarks
//!
//! The `Exp` function complements the action of the `Log` function and is sometimes
//! referred to as the antilogarithm. It calculates e raised to a power, where e is
//! the base of natural logarithms (approximately 2.718282).
//!
//! **Important Characteristics:**
//!
//! - Returns e^number where e ≈ 2.718282
//! - Inverse of the natural logarithm (Log)
//! - Always returns positive value (e^x > 0 for all x)
//! - Domain: all real numbers
//! - Range: positive real numbers (> 0)
//! - Exp(0) = 1
//! - Exp(1) ≈ 2.718282
//! - For large positive values, can cause overflow
//! - For large negative values, approaches 0
//! - Maximum argument ≈ 709.78 (causes overflow above this)
//! - Minimum useful argument ≈ -708 (returns values very close to 0)
//!
//! ## Mathematical Properties
//!
//! - **Identity**: Exp(0) = 1
//! - **Euler's Number**: Exp(1) = e ≈ 2.718282
//! - **Inverse of Log**: Exp(Log(x)) = x (for x > 0)
//! - **Product Rule**: Exp(a + b) = Exp(a) * Exp(b)
//! - **Power Rule**: Exp(a * b) = Exp(a)^b
//! - **Derivative**: d/dx[Exp(x)] = Exp(x)
//! - **Integral**: ∫Exp(x)dx = Exp(x) + C
//!
//! ## Common Applications
//!
//! - Exponential growth/decay calculations
//! - Compound interest formulas
//! - Population growth models
//! - Radioactive decay
//! - Probability distributions (normal, exponential)
//! - Signal processing
//! - Physics and engineering calculations
//! - Statistics and data analysis
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! Dim result As Double
//!
//! ' Basic exponential calculation
//! result = Exp(1)           ' Returns e ≈ 2.718282
//! result = Exp(0)           ' Returns 1
//! result = Exp(2)           ' Returns e² ≈ 7.389056
//!
//! ' Negative exponents
//! result = Exp(-1)          ' Returns 1/e ≈ 0.367879
//! result = Exp(-2)          ' Returns 1/e² ≈ 0.135335
//! ```
//!
//! ### Exponential Growth
//!
//! ```vb
//! Function ExponentialGrowth(initial As Double, rate As Double, time As Double) As Double
//!     ' Calculate exponential growth: A = A₀ * e^(rt)
//!     ' initial = initial amount
//!     ' rate = growth rate (as decimal, e.g., 0.05 for 5%)
//!     ' time = time period
//!     
//!     ExponentialGrowth = initial * Exp(rate * time)
//! End Function
//!
//! ' Example: Population growth
//! ' Initial population: 1000, growth rate: 3% per year, time: 10 years
//! Dim population As Double
//! population = ExponentialGrowth(1000, 0.03, 10)  ' ≈ 1349.86
//! ```
//!
//! ### Compound Interest
//!
//! ```vb
//! Function ContinuousCompoundInterest(principal As Double, rate As Double, _
//!                                      time As Double) As Double
//!     ' Calculate continuously compounded interest: A = P * e^(rt)
//!     ' principal = initial investment
//!     ' rate = annual interest rate (as decimal)
//!     ' time = time in years
//!     
//!     ContinuousCompoundInterest = principal * Exp(rate * time)
//! End Function
//!
//! ' Example: $1000 at 5% for 10 years
//! Dim amount As Double
//! amount = ContinuousCompoundInterest(1000, 0.05, 10)  ' ≈ $1648.72
//! ```
//!
//! ## Common Patterns
//!
//! ### Radioactive Decay
//!
//! ```vb
//! Function RadioactiveDecay(initialAmount As Double, decayConstant As Double, _
//!                           time As Double) As Double
//!     ' Calculate remaining amount: N = N₀ * e^(-λt)
//!     ' initialAmount = initial quantity
//!     ' decayConstant = decay constant (λ)
//!     ' time = elapsed time
//!     
//!     RadioactiveDecay = initialAmount * Exp(-decayConstant * time)
//! End Function
//!
//! ' Example: Half-life calculation
//! Function HalfLife(decayConstant As Double) As Double
//!     ' t₁/₂ = ln(2) / λ
//!     HalfLife = Log(2) / decayConstant
//! End Function
//! ```
//!
//! ### Normal Distribution
//!
//! ```vb
//! Function NormalDistribution(x As Double, mean As Double, stdDev As Double) As Double
//!     ' Calculate normal (Gaussian) distribution PDF
//!     ' f(x) = (1 / (σ√(2π))) * e^(-(x-μ)²/(2σ²))
//!     
//!     Dim pi As Double
//!     Dim exponent As Double
//!     
//!     pi = 4 * Atn(1)  ' Calculate π
//!     exponent = -((x - mean) ^ 2) / (2 * stdDev ^ 2)
//!     
//!     NormalDistribution = (1 / (stdDev * Sqr(2 * pi))) * Exp(exponent)
//! End Function
//! ```
//!
//! ### Exponential Smoothing
//!
//! ```vb
//! Function ExponentialSmoothing(data() As Double, alpha As Double) As Variant
//!     ' Apply exponential smoothing to data
//!     ' alpha = smoothing factor (0 < α < 1)
//!     
//!     Dim smoothed() As Double
//!     Dim i As Long
//!     
//!     ReDim smoothed(LBound(data) To UBound(data))
//!     
//!     ' First value is same as original
//!     smoothed(LBound(data)) = data(LBound(data))
//!     
//!     ' Apply smoothing formula: S_t = α * Y_t + (1-α) * S_{t-1}
//!     For i = LBound(data) + 1 To UBound(data)
//!         smoothed(i) = alpha * data(i) + (1 - alpha) * smoothed(i - 1)
//!     Next i
//!     
//!     ExponentialSmoothing = smoothed
//! End Function
//! ```
//!
//! ### Temperature Cooling (Newton's Law)
//!
//! ```vb
//! Function CoolingTemperature(initialTemp As Double, ambientTemp As Double, _
//!                             coolingConstant As Double, time As Double) As Double
//!     ' Newton's Law of Cooling: T(t) = T_ambient + (T₀ - T_ambient) * e^(-kt)
//!     ' initialTemp = initial temperature
//!     ' ambientTemp = surrounding temperature
//!     ' coolingConstant = cooling constant (k)
//!     ' time = elapsed time
//!     
//!     CoolingTemperature = ambientTemp + (initialTemp - ambientTemp) * Exp(-coolingConstant * time)
//! End Function
//!
//! ' Example: Coffee cooling from 90°C to room temperature (20°C)
//! Dim temp As Double
//! temp = CoolingTemperature(90, 20, 0.1, 10)  ' Temperature after 10 minutes
//! ```
//!
//! ### Convert Between Log Bases
//!
//! ```vb
//! Function LogBase(number As Double, base As Double) As Double
//!     ' Calculate logarithm with arbitrary base
//!     ' log_base(number) = ln(number) / ln(base)
//!     
//!     If number <= 0 Or base <= 0 Or base = 1 Then
//!         Err.Raise 5, , "Invalid argument"
//!     End If
//!     
//!     LogBase = Log(number) / Log(base)
//! End Function
//!
//! Function PowerWithBase(base As Double, exponent As Double) As Double
//!     ' Calculate base^exponent using Exp and Log
//!     ' base^exponent = e^(exponent * ln(base))
//!     
//!     If base <= 0 Then
//!         Err.Raise 5, , "Base must be positive"
//!     End If
//!     
//!     PowerWithBase = Exp(exponent * Log(base))
//! End Function
//! ```
//!
//! ### Sigmoid Function
//!
//! ```vb
//! Function Sigmoid(x As Double) As Double
//!     ' Logistic sigmoid function: σ(x) = 1 / (1 + e^(-x))
//!     ' Used in neural networks and machine learning
//!     
//!     Sigmoid = 1 / (1 + Exp(-x))
//! End Function
//!
//! Function SigmoidDerivative(x As Double) As Double
//!     ' Derivative of sigmoid: σ'(x) = σ(x) * (1 - σ(x))
//!     Dim s As Double
//!     s = Sigmoid(x)
//!     SigmoidDerivative = s * (1 - s)
//! End Function
//! ```
//!
//! ### Exponential Moving Average
//!
//! ```vb
//! Function CalculateEMA(prices() As Double, periods As Integer) As Variant
//!     ' Calculate Exponential Moving Average
//!     ' Commonly used in financial analysis
//!     
//!     Dim ema() As Double
//!     Dim multiplier As Double
//!     Dim i As Long
//!     
//!     ReDim ema(LBound(prices) To UBound(prices))
//!     
//!     ' Calculate multiplier: 2 / (periods + 1)
//!     multiplier = 2 / (periods + 1)
//!     
//!     ' First EMA is simple moving average
//!     ema(LBound(prices)) = prices(LBound(prices))
//!     
//!     ' Calculate EMA for remaining values
//!     For i = LBound(prices) + 1 To UBound(prices)
//!         ema(i) = (prices(i) - ema(i - 1)) * multiplier + ema(i - 1)
//!     Next i
//!     
//!     CalculateEMA = ema
//! End Function
//! ```
//!
//! ### Black-Scholes Option Pricing
//!
//! ```vb
//! Function BlackScholesCall(stockPrice As Double, strikePrice As Double, _
//!                           timeToExpiry As Double, riskFreeRate As Double, _
//!                           volatility As Double) As Double
//!     ' Simplified Black-Scholes formula for call option
//!     Dim d1 As Double, d2 As Double
//!     Dim pi As Double
//!     
//!     pi = 4 * Atn(1)
//!     
//!     d1 = (Log(stockPrice / strikePrice) + (riskFreeRate + 0.5 * volatility ^ 2) * timeToExpiry) / _
//!          (volatility * Sqr(timeToExpiry))
//!     d2 = d1 - volatility * Sqr(timeToExpiry)
//!     
//!     ' Using normal CDF approximation (simplified)
//!     BlackScholesCall = stockPrice * NormalCDF(d1) - _
//!                        strikePrice * Exp(-riskFreeRate * timeToExpiry) * NormalCDF(d2)
//! End Function
//! ```
//!
//! ### Poisson Distribution
//!
//! ```vb
//! Function PoissonProbability(k As Long, lambda As Double) As Double
//!     ' Calculate Poisson probability: P(X=k) = (λ^k * e^(-λ)) / k!
//!     ' k = number of occurrences
//!     ' lambda = average rate
//!     
//!     Dim i As Long
//!     Dim factorial As Double
//!     
//!     ' Calculate k!
//!     factorial = 1
//!     For i = 2 To k
//!         factorial = factorial * i
//!     Next i
//!     
//!     ' Calculate probability
//!     PoissonProbability = (lambda ^ k * Exp(-lambda)) / factorial
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Hyperbolic Functions
//!
//! ```vb
//! Function Sinh(x As Double) As Double
//!     ' Hyperbolic sine: sinh(x) = (e^x - e^(-x)) / 2
//!     Sinh = (Exp(x) - Exp(-x)) / 2
//! End Function
//!
//! Function Cosh(x As Double) As Double
//!     ' Hyperbolic cosine: cosh(x) = (e^x + e^(-x)) / 2
//!     Cosh = (Exp(x) + Exp(-x)) / 2
//! End Function
//!
//! Function Tanh(x As Double) As Double
//!     ' Hyperbolic tangent: tanh(x) = sinh(x) / cosh(x)
//!     Dim ex As Double
//!     ex = Exp(x)
//!     Tanh = (ex - 1 / ex) / (ex + 1 / ex)
//! End Function
//!
//! Function ArcSinh(x As Double) As Double
//!     ' Inverse hyperbolic sine: asinh(x) = ln(x + √(x² + 1))
//!     ArcSinh = Log(x + Sqr(x * x + 1))
//! End Function
//!
//! Function ArcCosh(x As Double) As Double
//!     ' Inverse hyperbolic cosine: acosh(x) = ln(x + √(x² - 1))
//!     If x < 1 Then
//!         Err.Raise 5, , "Argument must be >= 1"
//!     End If
//!     ArcCosh = Log(x + Sqr(x * x - 1))
//! End Function
//!
//! Function ArcTanh(x As Double) As Double
//!     ' Inverse hyperbolic tangent: atanh(x) = 0.5 * ln((1+x)/(1-x))
//!     If Abs(x) >= 1 Then
//!         Err.Raise 5, , "Argument must be in (-1, 1)"
//!     End If
//!     ArcTanh = 0.5 * Log((1 + x) / (1 - x))
//! End Function
//! ```
//!
//! ### Taylor Series Approximation
//!
//! ```vb
//! Function ExpTaylor(x As Double, terms As Integer) As Double
//!     ' Calculate Exp(x) using Taylor series
//!     ' e^x = 1 + x + x²/2! + x³/3! + x⁴/4! + ...
//!     
//!     Dim result As Double
//!     Dim term As Double
//!     Dim i As Integer
//!     
//!     result = 1  ' First term
//!     term = 1
//!     
//!     For i = 1 To terms
//!         term = term * x / i
//!         result = result + term
//!     Next i
//!     
//!     ExpTaylor = result
//! End Function
//!
//! Sub CompareTaylorWithBuiltIn()
//!     Dim x As Double
//!     Dim terms As Integer
//!     
//!     x = 2
//!     
//!     Debug.Print "Comparing Taylor series with built-in Exp:"
//!     For terms = 1 To 20
//!         Debug.Print "Terms: " & terms & ", Taylor: " & ExpTaylor(x, terms) & _
//!                     ", Built-in: " & Exp(x) & ", Error: " & Abs(ExpTaylor(x, terms) - Exp(x))
//!     Next terms
//! End Sub
//! ```
//!
//! ### Numerical Integration Using Exponential
//!
//! ```vb
//! Function IntegrateExp(lowerBound As Double, upperBound As Double, _
//!                       intervals As Long) As Double
//!     ' Numerical integration of e^x using trapezoidal rule
//!     ' ∫e^x dx from a to b
//!     
//!     Dim h As Double
//!     Dim sum As Double
//!     Dim x As Double
//!     Dim i As Long
//!     
//!     h = (upperBound - lowerBound) / intervals
//!     sum = (Exp(lowerBound) + Exp(upperBound)) / 2
//!     
//!     For i = 1 To intervals - 1
//!         x = lowerBound + i * h
//!         sum = sum + Exp(x)
//!     Next i
//!     
//!     IntegrateExp = sum * h
//! End Function
//!
//! Sub VerifyIntegration()
//!     Dim a As Double, b As Double
//!     Dim numerical As Double
//!     Dim analytical As Double
//!     
//!     a = 0
//!     b = 1
//!     
//!     numerical = IntegrateExp(a, b, 1000)
//!     analytical = Exp(b) - Exp(a)  ' Analytical solution: e^b - e^a
//!     
//!     Debug.Print "Numerical: " & numerical
//!     Debug.Print "Analytical: " & analytical
//!     Debug.Print "Error: " & Abs(numerical - analytical)
//! End Sub
//! ```
//!
//! ### Population Dynamics Model
//!
//! ```vb
//! Function LogisticGrowth(initialPop As Double, carryingCapacity As Double, _
//!                         growthRate As Double, time As Double) As Double
//!     ' Logistic growth model: P(t) = K / (1 + ((K - P₀) / P₀) * e^(-rt))
//!     ' initialPop = initial population (P₀)
//!     ' carryingCapacity = maximum sustainable population (K)
//!     ' growthRate = intrinsic growth rate (r)
//!     ' time = time
//!     
//!     Dim ratio As Double
//!     
//!     ratio = (carryingCapacity - initialPop) / initialPop
//!     LogisticGrowth = carryingCapacity / (1 + ratio * Exp(-growthRate * time))
//! End Function
//!
//! Sub PlotLogisticGrowth()
//!     Dim t As Double
//!     Dim population As Double
//!     
//!     Debug.Print "Time", "Population"
//!     Debug.Print String(40, "-")
//!     
//!     For t = 0 To 50 Step 5
//!         population = LogisticGrowth(100, 10000, 0.1, t)
//!         Debug.Print t, Format(population, "#,##0.00")
//!     Next t
//! End Sub
//! ```
//!
//! ### Complex Exponential (Euler's Formula)
//!
//! ```vb
//! Type ComplexNumber
//!     Real As Double
//!     Imaginary As Double
//! End Type
//!
//! Function ComplexExp(z As ComplexNumber) As ComplexNumber
//!     ' Calculate e^z for complex number z = a + bi
//!     ' e^(a+bi) = e^a * (cos(b) + i*sin(b))  [Euler's formula]
//!     
//!     Dim result As ComplexNumber
//!     Dim magnitude As Double
//!     
//!     magnitude = Exp(z.Real)
//!     result.Real = magnitude * Cos(z.Imaginary)
//!     result.Imaginary = magnitude * Sin(z.Imaginary)
//!     
//!     ComplexExp = result
//! End Function
//!
//! Sub DemonstrateEulerFormula()
//!     Dim z As ComplexNumber
//!     Dim result As ComplexNumber
//!     Dim pi As Double
//!     
//!     pi = 4 * Atn(1)
//!     
//!     ' e^(i*π) = -1 (Euler's identity)
//!     z.Real = 0
//!     z.Imaginary = pi
//!     result = ComplexExp(z)
//!     
//!     Debug.Print "e^(i*π) = " & Format(result.Real, "0.0000") & " + " & _
//!                 Format(result.Imaginary, "0.0000") & "i"
//!     Debug.Print "Should be approximately -1 + 0i"
//! End Sub
//! ```
//!
//! ### Financial Option Greeks
//!
//! ```vb
//! Function CalculateDelta(stockPrice As Double, strikePrice As Double, _
//!                         timeToExpiry As Double, riskFreeRate As Double, _
//!                         volatility As Double) As Double
//!     ' Calculate Delta (rate of change of option price with respect to stock price)
//!     Dim d1 As Double
//!     
//!     d1 = (Log(stockPrice / strikePrice) + (riskFreeRate + 0.5 * volatility ^ 2) * timeToExpiry) / _
//!          (volatility * Sqr(timeToExpiry))
//!     
//!     CalculateDelta = NormalCDF(d1)
//! End Function
//!
//! Function CalculateTheta(stockPrice As Double, strikePrice As Double, _
//!                         timeToExpiry As Double, riskFreeRate As Double, _
//!                         volatility As Double) As Double
//!     ' Calculate Theta (rate of change of option price with respect to time)
//!     ' Involves exponential decay term
//!     Dim d1 As Double, d2 As Double
//!     Dim pi As Double
//!     
//!     pi = 4 * Atn(1)
//!     
//!     d1 = (Log(stockPrice / strikePrice) + (riskFreeRate + 0.5 * volatility ^ 2) * timeToExpiry) / _
//!          (volatility * Sqr(timeToExpiry))
//!     d2 = d1 - volatility * Sqr(timeToExpiry)
//!     
//!     CalculateTheta = -(stockPrice * NormalPDF(d1) * volatility) / (2 * Sqr(timeToExpiry)) - _
//!                      riskFreeRate * strikePrice * Exp(-riskFreeRate * timeToExpiry) * NormalCDF(d2)
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeExp(x As Double) As Double
//!     On Error GoTo ErrorHandler
//!     
//!     ' Check for potential overflow
//!     If x > 709.78 Then
//!         Err.Raise 6, , "Overflow: exponent too large"
//!     End If
//!     
//!     SafeExp = Exp(x)
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 6  ' Overflow
//!             MsgBox "Exponential overflow. Result is too large to represent.", vbExclamation
//!             SafeExp = 0
//!         Case Else
//!             MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
//!             SafeExp = 0
//!     End Select
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 6** (Overflow): Argument too large (> 709.78 approximately)
//! - **Error 13** (Type mismatch): Non-numeric argument
//!
//! ## Performance Considerations
//!
//! - `Exp` is a built-in function, optimized for speed
//! - Hardware-accelerated on most processors
//! - Very fast compared to manual Taylor series calculation
//! - For repeated calculations with same value, consider caching
//! - For small values near 0, consider using Exp(x) - 1 pattern for precision
//!
//! ## Best Practices
//!
//! ### Check for Overflow
//!
//! ```vb
//! ' Good - Check before calculation
//! If x <= 709 Then
//!     result = Exp(x)
//! Else
//!     MsgBox "Value too large for Exp function"
//! End If
//!
//! ' Or use error handling
//! On Error Resume Next
//! result = Exp(x)
//! If Err.Number = 6 Then
//!     MsgBox "Exponential overflow"
//!     result = 0
//! End If
//! On Error GoTo 0
//! ```
//!
//! ### Use with Log for Powers
//!
//! ```vb
//! ' Calculate a^b where a and b are any real numbers
//! ' Good - Use Exp and Log
//! Function Power(base As Double, exponent As Double) As Double
//!     If base <= 0 Then
//!         Err.Raise 5, , "Base must be positive"
//!     End If
//!     Power = Exp(exponent * Log(base))
//! End Function
//!
//! ' For integer exponents, use ^ operator
//! result = base ^ intExponent  ' More efficient
//! ```
//!
//! ## Comparison with Other Functions
//!
//! ### Exp vs ^ Operator
//!
//! ```vb
//! ' Exp - Natural exponential (base e)
//! result = Exp(2)              ' e^2 ≈ 7.389056
//!
//! ' ^ - General power operator
//! result = 2.718282 ^ 2        ' Approximately same
//! result = 10 ^ 2              ' 100 (different base)
//! ```
//!
//! ### Exp vs Log
//!
//! ```vb
//! ' Exp and Log are inverse functions
//! x = 5
//! result = Exp(Log(x))         ' Returns 5
//! result = Log(Exp(x))         ' Returns 5
//! ```
//!
//! ## Limitations
//!
//! - Maximum argument approximately 709.78 (causes overflow)
//! - Minimum useful argument approximately -708 (underflows to 0)
//! - Returns Double (limited precision ~15-16 significant digits)
//! - Cannot directly calculate complex exponentials (requires manual implementation)
//! - Single argument only (unlike some languages with multi-parameter exp functions)
//!
//! ## Related Functions
//!
//! - `Log`: Natural logarithm (inverse of Exp)
//! - `Sqr`: Square root
//! - `^`: Power operator
//! - `Sin`, `Cos`: Trigonometric functions (related via Euler's formula)
//! - `Atn`: Arctangent

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn exp_basic() {
        let source = r"
result = Exp(2)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Exp"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("2"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_zero() {
        let source = r"
result = Exp(0)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Exp"),
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
        ]);
    }

    #[test]
    fn exp_negative() {
        let source = r"
result = Exp(-1)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Exp"),
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
        ]);
    }

    #[test]
    fn exp_variable() {
        let source = r"
y = Exp(x)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("y"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Exp"),
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
        ]);
    }

    #[test]
    fn exp_in_function() {
        let source = r"
Function ExponentialGrowth(rate As Double, time As Double) As Double
    ExponentialGrowth = initial * Exp(rate * time)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("ExponentialGrowth"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("rate"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    Comma,
                    Whitespace,
                },
                TimeKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("ExponentialGrowth"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("initial"),
                            },
                            Whitespace,
                            MultiplicationOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("Exp"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("rate"),
                                            },
                                            Whitespace,
                                            MultiplicationOperator,
                                            Whitespace,
                                            IdentifierExpression {
                                                TimeKeyword,
                                            },
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
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_multiplication() {
        let source = r"
result = principal * Exp(rate * time)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("principal"),
                    },
                    Whitespace,
                    MultiplicationOperator,
                    Whitespace,
                    CallExpression {
                        Identifier ("Exp"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("rate"),
                                    },
                                    Whitespace,
                                    MultiplicationOperator,
                                    Whitespace,
                                    IdentifierExpression {
                                        TimeKeyword,
                                    },
                                },
                            },
                        },
                        RightParenthesis,
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_decay() {
        let source = r"
amount = Exp(-rate * t)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("amount"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Exp"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            BinaryExpression {
                                UnaryExpression {
                                    SubtractionOperator,
                                    IdentifierExpression {
                                        Identifier ("rate"),
                                    },
                                },
                                Whitespace,
                                MultiplicationOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("t"),
                                },
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_sigmoid() {
        let source = r"
result = 1 / (1 + Exp(-x))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    NumericLiteralExpression {
                        IntegerLiteral ("1"),
                    },
                    Whitespace,
                    DivisionOperator,
                    Whitespace,
                    ParenthesizedExpression {
                        LeftParenthesis,
                        BinaryExpression {
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("Exp"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        UnaryExpression {
                                            SubtractionOperator,
                                            IdentifierExpression {
                                                Identifier ("x"),
                                            },
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
        ]);
    }

    #[test]
    fn exp_hyperbolic() {
        let source = r"
sinh = (Exp(x) - Exp(-x)) / 2
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("sinh"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    ParenthesizedExpression {
                        LeftParenthesis,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Exp"),
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
                            SubtractionOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("Exp"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        UnaryExpression {
                                            SubtractionOperator,
                                            IdentifierExpression {
                                                Identifier ("x"),
                                            },
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
                    NumericLiteralExpression {
                        IntegerLiteral ("2"),
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_cosh() {
        let source = r"
cosh = (Exp(x) + Exp(-x)) / 2
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("cosh"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    ParenthesizedExpression {
                        LeftParenthesis,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Exp"),
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
                            AdditionOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("Exp"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        UnaryExpression {
                                            SubtractionOperator,
                                            IdentifierExpression {
                                                Identifier ("x"),
                                            },
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
                    NumericLiteralExpression {
                        IntegerLiteral ("2"),
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_normal_distribution() {
        let source = r"
result = (1 / (stdDev * Sqr(2 * pi))) * Exp(exponent)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    ParenthesizedExpression {
                        LeftParenthesis,
                        BinaryExpression {
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
                            Whitespace,
                            DivisionOperator,
                            Whitespace,
                            ParenthesizedExpression {
                                LeftParenthesis,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("stdDev"),
                                    },
                                    Whitespace,
                                    MultiplicationOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Sqr"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                BinaryExpression {
                                                    NumericLiteralExpression {
                                                        IntegerLiteral ("2"),
                                                    },
                                                    Whitespace,
                                                    MultiplicationOperator,
                                                    Whitespace,
                                                    IdentifierExpression {
                                                        Identifier ("pi"),
                                                    },
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    MultiplicationOperator,
                    Whitespace,
                    CallExpression {
                        Identifier ("Exp"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("exponent"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_with_log() {
        let source = r"
power = Exp(exponent * Log(base))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("power"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Exp"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("exponent"),
                                },
                                Whitespace,
                                MultiplicationOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Log"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                BaseKeyword,
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
        ]);
    }

    #[test]
    fn exp_compound_interest() {
        let source = r"
amount = principal * Exp(rate * time)
interest = amount - principal
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("amount"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("principal"),
                    },
                    Whitespace,
                    MultiplicationOperator,
                    Whitespace,
                    CallExpression {
                        Identifier ("Exp"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("rate"),
                                    },
                                    Whitespace,
                                    MultiplicationOperator,
                                    Whitespace,
                                    IdentifierExpression {
                                        TimeKeyword,
                                    },
                                },
                            },
                        },
                        RightParenthesis,
                    },
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("interest"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("amount"),
                    },
                    Whitespace,
                    SubtractionOperator,
                    Whitespace,
                    IdentifierExpression {
                        Identifier ("principal"),
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_debug_print() {
        let source = r"
Debug.Print Exp(1)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("Debug"),
                PeriodOperator,
                PrintKeyword,
                Whitespace,
                Identifier ("Exp"),
                LeftParenthesis,
                IntegerLiteral ("1"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_in_if() {
        let source = r"
If Exp(x) > threshold Then
    ProcessValue
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Exp"),
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
                    GreaterThanOperator,
                    Whitespace,
                    IdentifierExpression {
                        Identifier ("threshold"),
                    },
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("ProcessValue"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_in_loop() {
        let source = r"
For i = 0 To 10
    result = Exp(i * 0.1)
    Debug.Print result
Next i
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            ForStatement {
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
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Exp"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("i"),
                                        },
                                        Whitespace,
                                        MultiplicationOperator,
                                        Whitespace,
                                        NumericLiteralExpression {
                                            SingleLiteral,
                                        },
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("result"),
                        Newline,
                    },
                },
                NextKeyword,
                Whitespace,
                Identifier ("i"),
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_poisson() {
        let source = r"
probability = (lambda ^ k * Exp(-lambda)) / factorial
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("probability"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    ParenthesizedExpression {
                        LeftParenthesis,
                        BinaryExpression {
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("lambda"),
                                },
                                Whitespace,
                                ExponentiationOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("k"),
                                },
                            },
                            Whitespace,
                            MultiplicationOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("Exp"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        UnaryExpression {
                                            SubtractionOperator,
                                            IdentifierExpression {
                                                Identifier ("lambda"),
                                            },
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
                    IdentifierExpression {
                        Identifier ("factorial"),
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_cooling() {
        let source = r"
temp = ambientTemp + (initialTemp - ambientTemp) * Exp(-coolingConstant * time)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("temp"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("ambientTemp"),
                    },
                    Whitespace,
                    AdditionOperator,
                    Whitespace,
                    BinaryExpression {
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("initialTemp"),
                                },
                                Whitespace,
                                SubtractionOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("ambientTemp"),
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        MultiplicationOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Exp"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        UnaryExpression {
                                            SubtractionOperator,
                                            IdentifierExpression {
                                                Identifier ("coolingConstant"),
                                            },
                                        },
                                        Whitespace,
                                        MultiplicationOperator,
                                        Whitespace,
                                        IdentifierExpression {
                                            TimeKeyword,
                                        },
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_logistic_growth() {
        let source = r"
population = capacity / (1 + ratio * Exp(-growthRate * time))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("population"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("capacity"),
                    },
                    Whitespace,
                    DivisionOperator,
                    Whitespace,
                    ParenthesizedExpression {
                        LeftParenthesis,
                        BinaryExpression {
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("ratio"),
                                },
                                Whitespace,
                                MultiplicationOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Exp"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            BinaryExpression {
                                                UnaryExpression {
                                                    SubtractionOperator,
                                                    IdentifierExpression {
                                                        Identifier ("growthRate"),
                                                    },
                                                },
                                                Whitespace,
                                                MultiplicationOperator,
                                                Whitespace,
                                                IdentifierExpression {
                                                    TimeKeyword,
                                                },
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                        },
                        RightParenthesis,
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_error_handling() {
        let source = r#"
On Error Resume Next
result = Exp(x)
If Err.Number = 6 Then
    MsgBox "Overflow"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            OnErrorStatement {
                OnKeyword,
                Whitespace,
                ErrorKeyword,
                Whitespace,
                ResumeKeyword,
                Whitespace,
                NextKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Exp"),
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
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    MemberAccessExpression {
                        Identifier ("Err"),
                        PeriodOperator,
                        Identifier ("Number"),
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("6"),
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
                        StringLiteral ("\"Overflow\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_taylor_series() {
        let source = r"
term = term * x / i
result = result + term
builtin = Exp(x)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("term"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    BinaryExpression {
                        IdentifierExpression {
                            Identifier ("term"),
                        },
                        Whitespace,
                        MultiplicationOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("x"),
                        },
                    },
                    Whitespace,
                    DivisionOperator,
                    Whitespace,
                    IdentifierExpression {
                        Identifier ("i"),
                    },
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("result"),
                    },
                    Whitespace,
                    AdditionOperator,
                    Whitespace,
                    IdentifierExpression {
                        Identifier ("term"),
                    },
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("builtin"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Exp"),
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
        ]);
    }

    #[test]
    fn exp_black_scholes() {
        let source = r"
callPrice = stockPrice * d1Value - strikePrice * Exp(-rate * time) * d2Value
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("callPrice"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    BinaryExpression {
                        IdentifierExpression {
                            Identifier ("stockPrice"),
                        },
                        Whitespace,
                        MultiplicationOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("d1Value"),
                        },
                    },
                    Whitespace,
                    SubtractionOperator,
                    Whitespace,
                    BinaryExpression {
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("strikePrice"),
                            },
                            Whitespace,
                            MultiplicationOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("Exp"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            UnaryExpression {
                                                SubtractionOperator,
                                                IdentifierExpression {
                                                    Identifier ("rate"),
                                                },
                                            },
                                            Whitespace,
                                            MultiplicationOperator,
                                            Whitespace,
                                            IdentifierExpression {
                                                TimeKeyword,
                                            },
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Whitespace,
                        MultiplicationOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("d2Value"),
                        },
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_complex() {
        let source = r"
magnitude = Exp(z.Real)
result.Real = magnitude * Cos(z.Imaginary)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("magnitude"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Exp"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            MemberAccessExpression {
                                Identifier ("z"),
                                PeriodOperator,
                                Identifier ("Real"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            AssignmentStatement {
                MemberAccessExpression {
                    Identifier ("result"),
                    PeriodOperator,
                    Identifier ("Real"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("magnitude"),
                    },
                    Whitespace,
                    MultiplicationOperator,
                    Whitespace,
                    CallExpression {
                        Identifier ("Cos"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                MemberAccessExpression {
                                    Identifier ("z"),
                                    PeriodOperator,
                                    Identifier ("Imaginary"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_integration() {
        let source = r"
sum = sum + Exp(x)
numerical = sum * h
analytical = Exp(b) - Exp(a)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("sum"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("sum"),
                    },
                    Whitespace,
                    AdditionOperator,
                    Whitespace,
                    CallExpression {
                        Identifier ("Exp"),
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
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("numerical"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("sum"),
                    },
                    Whitespace,
                    MultiplicationOperator,
                    Whitespace,
                    IdentifierExpression {
                        Identifier ("h"),
                    },
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("analytical"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Exp"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("b"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    SubtractionOperator,
                    Whitespace,
                    CallExpression {
                        Identifier ("Exp"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("a"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn exp_nested() {
        let source = r"
result = Exp(Log(Exp(x)))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Exp"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Log"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        CallExpression {
                                            Identifier ("Exp"),
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
        ]);
    }
}
