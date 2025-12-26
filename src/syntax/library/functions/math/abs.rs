//! # `Abs` Function
//!
//! Returns the absolute value of a number.
//!
//! ## Syntax
//!
//! ```vb
//! Abs(number)
//! ```
//!
//! ## Parts
//!
//! - **number**: Required. Any valid numeric expression. If number contains Null, Null is returned;
//!   if it is an uninitialized variable, zero is returned.
//!
//! ## Return Value
//!
//! The return type is the same as the input type, except:
//! - If number is a Variant containing Null, returns Null
//! - If number is an uninitialized Variant, returns 0
//! - The absolute value is always non-negative (>= 0)
//!
//! ## Remarks
//!
//! - **Absolute Value**: The absolute value of a number is its unsigned magnitude. For example,
//!   Abs(-1) and Abs(1) both return 1.
//! - **Type Preservation**: The return type matches the input type. If you pass an Integer, you
//!   get an Integer back. If you pass a Double, you get a Double back.
//! - **Null Handling**: If the argument is Null, the function returns Null.
//! - **Overflow**: For the most negative value of Integer (-32768) or Long (-2147483648), Abs
//!   will cause an overflow error because the positive equivalent is outside the valid range.
//! - **Performance**: Abs is highly optimized and very fast for numeric operations.
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! Dim result As Integer
//! result = Abs(-50)
//! ' result = 50
//! ```
//!
//! ### With Positive Numbers
//!
//! ```vb
//! Dim value As Integer
//! value = Abs(25)
//! ' value = 25 (unchanged)
//! ```
//!
//! ### With Floating Point
//!
//! ```vb
//! Dim distance As Double
//! distance = Abs(-12.75)
//! ' distance = 12.75
//! ```
//!
//! ### With Zero
//!
//! ```vb
//! Dim zero As Integer
//! zero = Abs(0)
//! ' zero = 0
//! ```
//!
//! ### With Expressions
//!
//! ```vb
//! Dim x As Integer, y As Integer
//! x = 10
//! y = 20
//! Dim difference As Integer
//! difference = Abs(x - y)
//! ' difference = 10
//! ```
//!
//! ### Calculating Distance
//!
//! ```vb
//! Function Distance(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
//!     Distance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
//! End Function
//!
//! ' Often used with Abs for 1D distance:
//! Dim dist As Double
//! dist = Abs(x2 - x1)
//! ```
//!
//! ### With Currency
//!
//! ```vb
//! Dim amount As Currency
//! amount = Abs(-1234.56@)
//! ' amount = 1234.56
//! ```
//!
//! ### With Variants
//!
//! ```vb
//! Dim v As Variant
//! v = -42
//! Dim result As Variant
//! result = Abs(v)
//! ' result = 42
//! ```
//!
//! ## Common Patterns
//!
//! ### Ensuring Positive Values
//!
//! ```vb
//! Sub ProcessValue(ByVal input As Integer)
//!     Dim positiveInput As Integer
//!     positiveInput = Abs(input)
//!     ' Always work with positive values
//!     DoSomething positiveInput
//! End Sub
//! ```
//!
//! ### Calculating Difference
//!
//! ```vb
//! Function GetDifference(a As Long, b As Long) As Long
//!     GetDifference = Abs(a - b)
//! End Function
//! ```
//!
//! ### Data Validation
//!
//! ```vb
//! Function IsWithinTolerance(actual As Double, expected As Double, tolerance As Double) As Boolean
//!     IsWithinTolerance = (Abs(actual - expected) <= tolerance)
//! End Function
//! ```
//!
//! ### Financial Calculations
//!
//! ```vb
//! Function CalculateVariance(actual As Currency, budget As Currency) As Currency
//!     CalculateVariance = Abs(actual - budget)
//! End Function
//! ```
//!
//! ### Array Processing
//!
//! ```vb
//! Sub MakeArrayPositive(arr() As Integer)
//!     Dim i As Integer
//!     For i = LBound(arr) To UBound(arr)
//!         arr(i) = Abs(arr(i))
//!     Next i
//! End Sub
//! ```
//!
//! ### Comparison Logic
//!
//! ```vb
//! Function MaxAbsValue(a As Double, b As Double) As Double
//!     If Abs(a) > Abs(b) Then
//!         MaxAbsValue = Abs(a)
//!     Else
//!         MaxAbsValue = Abs(b)
//!     End If
//! End Function
//! ```
//!
//! ### Coordinate Systems
//!
//! ```vb
//! Function ManhattanDistance(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer) As Integer
//!     ManhattanDistance = Abs(x2 - x1) + Abs(y2 - y1)
//! End Function
//! ```
//!
//! ## Related Functions
//!
//! - `Sgn`: Returns the sign of a number (-1, 0, or 1)
//! - `Fix`: Returns the integer portion of a number (truncates toward zero)
//! - `Int`: Returns the integer portion of a number (rounds down)
//! - `Round`: Rounds a number to a specified number of decimal places
//!
//! ## Type Compatibility
//!
//! | Input Type | Return Type | Notes |
//! |------------|-------------|-------|
//! | Byte | Byte | Always positive already |
//! | Integer | Integer | Can overflow at -32768 |
//! | Long | Long | Can overflow at -2147483648 |
//! | Single | Single | Preserves precision |
//! | Double | Double | Preserves precision |
//! | Currency | Currency | Preserves 4 decimal places |
//! | Variant (numeric) | Variant | Type preserved |
//! | Variant (Null) | Null | Returns Null |
//!
//! ## Performance Notes
//!
//! - `Abs` is a very fast intrinsic function
//! - No function call overhead in compiled code
//! - Optimized to CPU instructions where possible
//! - Prefer `Abs` over manual `If`/`Then` checks for performance
//!
//! `Abs` is parsed as a regular function call (`CallExpression`)
//! This module serves as documentation and reference for the Abs function

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn abs_simple_negative() {
        let source = r"
Sub Test()
    x = Abs(-50)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn abs_simple_positive() {
        let source = r"
Sub Test()
    x = Abs(25)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
    }

    #[test]
    fn abs_with_zero() {
        let source = r"
Sub Test()
    x = Abs(0)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
    }

    #[test]
    fn abs_with_variable() {
        let source = r"
Sub Test()
    result = Abs(value)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("value"));
    }

    #[test]
    fn abs_with_expression() {
        let source = r"
Sub Test()
    diff = Abs(x - y)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
    }

    #[test]
    fn abs_floating_point() {
        let source = r"
Sub Test()
    distance = Abs(-12.75)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("12.75"));
    }

    #[test]
    fn abs_in_assignment() {
        let source = r"
Sub Test()
    Dim x As Integer
    x = Abs(-100)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("DimStatement"));
    }

    #[test]
    fn abs_in_if_condition() {
        let source = r"
Sub Test()
    If Abs(value) > 100 Then
        ProcessLargeValue
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn abs_in_comparison() {
        let source = r"
Sub Test()
    If Abs(x - y) < tolerance Then
        DoSomething
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("tolerance"));
    }

    #[test]
    fn abs_nested_call() {
        let source = r"
Sub Test()
    result = Abs(GetValue())
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("GetValue"));
    }

    #[test]
    fn abs_multiple_calls() {
        let source = r"
Sub Test()
    a = Abs(-10)
    b = Abs(-20)
    c = Abs(-30)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("Abs").count();
        assert!(count >= 3);
    }

    #[test]
    fn abs_in_function_return() {
        let source = r"
Function GetDistance() As Double
    GetDistance = Abs(x2 - x1)
End Function
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("FunctionStatement"));
    }

    #[test]
    fn abs_in_loop() {
        let source = r"
Sub Test()
    For i = 1 To 10
        arr(i) = Abs(arr(i))
    Next i
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("ForStatement"));
    }

    #[test]
    fn abs_with_array_element() {
        let source = r"
Sub Test()
    result = Abs(values(index))
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("values"));
    }

    #[test]
    fn abs_with_property_access() {
        let source = r"
Sub Test()
    total = Abs(obj.Value)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("obj"));
    }

    #[test]
    fn abs_case_insensitive() {
        let source = r"
Sub Test()
    x = ABS(-50)
    y = abs(-25)
    z = AbS(-10)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ABS") || debug.contains("abs") || debug.contains("AbS"));
    }

    #[test]
    fn abs_in_print() {
        let source = r"
Sub Test()
    Debug.Print Abs(-42)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
    }

    #[test]
    fn abs_in_select_case() {
        let source = r"
Sub Test()
    Select Case Abs(value)
        Case Is > 100
            ProcessLarge
        Case Else
            ProcessSmall
    End Select
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("SelectCaseStatement"));
    }

    #[test]
    fn abs_with_parenthesized_expression() {
        let source = r"
Sub Test()
    result = Abs((x + y) * 2)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
    }

    #[test]
    fn abs_in_do_loop() {
        let source = r"
Sub Test()
    Do While Abs(delta) > 0.001
        Adjust
    Loop
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("DoStatement"));
    }

    #[test]
    fn abs_with_type_suffix() {
        let source = r"
Sub Test()
    x = Abs(-100%)
    y = Abs(-200&)
    z = Abs(-3.14#)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("Abs").count();
        assert!(count >= 3);
    }

    #[test]
    fn abs_in_while_loop() {
        let source = r"
Sub Test()
    While Abs(current - target) > threshold
        Step
    Wend
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("WhileStatement"));
    }

    #[test]
    fn abs_with_binary_operators() {
        let source = r"
Sub Test()
    result = Abs(a + b - c * d / e)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
    }

    #[test]
    fn abs_in_with_block() {
        let source = r"
Sub Test()
    With myObject
        .Value = Abs(.Delta)
    End With
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("WithStatement"));
    }

    #[test]
    fn abs_currency_literal() {
        let source = r"
Sub Test()
    amount = Abs(-1234.56@)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
    }

    #[test]
    fn abs_in_function_parameter() {
        let source = r"
Sub Test()
    Call ProcessValue(Abs(-50))
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("CallStatement"));
    }

    #[test]
    fn abs_chained_operations() {
        let source = r"
Sub Test()
    result = Abs(x) + Abs(y) - Abs(z)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("Abs").count();
        assert!(count >= 3);
    }

    #[test]
    fn abs_at_module_level() {
        let source = r"
Const MAX_VALUE = Abs(-1000)
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
    }

    #[test]
    fn abs_with_unary_minus() {
        let source = r"
Sub Test()
    x = Abs(-x)
    y = Abs(-(a + b))
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("Abs").count();
        assert!(count >= 2);
    }

    #[test]
    fn abs_preserves_whitespace() {
        let source = r"
Sub Test()
    x = Abs  (  -50  )
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("Abs"));
        assert!(debug.contains("Whitespace"));
    }
}
