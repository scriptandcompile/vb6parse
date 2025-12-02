//! # `FormatNumber` Function
//!
//! Returns an expression formatted as a number.
//!
//! ## Syntax
//!
//! ```vb
//! FormatNumber(expression[, numdigitsafterdecimal[, includeleadingdigit[, useparensfornegativenumbers[, groupdigits]]]])
//! ```
//!
//! ## Parameters
//!
//! - **expression**: Required. Expression to be formatted.
//! - **numdigitsafterdecimal**: Optional. `Numeric` value indicating how many places to the right of the decimal are displayed. Default is -1, which indicates the computer's regional settings are used.
//! - **includeleadingdigit**: Optional. Tristate constant that indicates whether a leading zero is displayed for fractional values. See Settings for values.
//! - **useparensfornegativenumbers**: Optional. Tristate constant that indicates whether to place negative values within parentheses. See Settings for values.
//! - **groupdigits**: Optional. Tristate constant that indicates whether numbers are grouped using the group delimiter specified in the computer's regional settings. See Settings for values.
//!
//! ## Settings
//!
//! The includeleadingdigit, useparensfornegativenumbers, and groupdigits arguments have the following settings:
//!
//! - **vbTrue** (-1): True
//! - **vbFalse** (0): False
//! - **vbUseDefault** (-2): Use the setting from the computer's regional settings
//!
//! ## Return Value
//!
//! Returns a Variant of subtype String containing the expression formatted as a number.
//!
//! ## Remarks
//!
//! The `FormatNumber` function provides a simple way to format numeric values using
//! the system's locale settings. Unlike `FormatCurrency`, it does not add a currency symbol,
//! making it ideal for general numeric display.
//!
//! **Important Characteristics:**
//!
//! - Uses system locale for formatting
//! - Default: 2 decimal places (from regional settings)
//! - Automatically adds thousand separators (if groupdigits=True)
//! - Negative numbers can be displayed with parentheses or minus sign
//! - Leading zeros controlled by regional settings or parameter
//! - No currency symbol added
//! - Returns empty string if expression is Null
//! - More convenient than Format for simple number formatting
//! - Less flexible than Format for custom patterns
//! - Locale-aware (respects user's regional settings)
//!
//! ## Typical Uses
//!
//! - Display numeric values in reports
//! - Format percentages (without % symbol)
//! - Show quantities and measurements
//! - Display statistical values
//! - Format calculated results
//! - Show decimal precision in user interfaces
//! - Display population or large numbers
//! - Format scientific data
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! Dim value As Double
//! value = 1234.567
//!
//! ' Default formatting (2 decimal places, system settings)
//! Debug.Print FormatNumber(value)              ' 1,234.57
//!
//! ' No decimal places
//! Debug.Print FormatNumber(value, 0)           ' 1,235
//!
//! ' Three decimal places
//! Debug.Print FormatNumber(value, 3)           ' 1,234.567
//!
//! ' Four decimal places
//! Debug.Print FormatNumber(value, 4)           ' 1,234.5670
//! ```
//!
//! ### Handling Negative Values
//!
//! ```vb
//! Dim negative As Double
//! negative = -1250.50
//!
//! ' Default negative (with minus sign)
//! Debug.Print FormatNumber(negative)           ' -1,250.50
//!
//! ' Parentheses for negative
//! Debug.Print FormatNumber(negative, 2, , vbTrue)   ' (1,250.50)
//!
//! ' No parentheses (explicit)
//! Debug.Print FormatNumber(negative, 2, , vbFalse)  ' -1,250.50
//! ```
//!
//! ### Control Leading Digits
//!
//! ```vb
//! Dim fraction As Double
//! fraction = 0.75
//!
//! ' With leading zero (default)
//! Debug.Print FormatNumber(fraction)           ' 0.75
//!
//! ' No leading zero
//! Debug.Print FormatNumber(fraction, 2, vbFalse)    ' .75
//!
//! ' Explicit leading zero
//! Debug.Print FormatNumber(fraction, 2, vbTrue)     ' 0.75
//! ```
//!
//! ### Control Grouping
//!
//! ```vb
//! Dim largeNumber As Double
//! largeNumber = 1234567.89
//!
//! ' With grouping (thousands separators)
//! Debug.Print FormatNumber(largeNumber, 2, , , vbTrue)   ' 1,234,567.89
//!
//! ' Without grouping
//! Debug.Print FormatNumber(largeNumber, 2, , , vbFalse)  ' 1234567.89
//! ```
//!
//! ## Common Patterns
//!
//! ### Display Statistical Results
//!
//! ```vb
//! Sub DisplayStatistics(data() As Double)
//!     Dim total As Double
//!     Dim average As Double
//!     Dim i As Long
//!     
//!     total = 0
//!     For i = LBound(data) To UBound(data)
//!         total = total + data(i)
//!     Next i
//!     
//!     average = total / (UBound(data) - LBound(data) + 1)
//!     
//!     Debug.Print "Count: " & FormatNumber(UBound(data) + 1, 0)
//!     Debug.Print "Total: " & FormatNumber(total, 2)
//!     Debug.Print "Average: " & FormatNumber(average, 2)
//! End Sub
//! ```
//!
//! ### Format Population Numbers
//!
//! ```vb
//! Function FormatPopulation(population As Long) As String
//!     FormatPopulation = FormatNumber(population, 0, , , vbTrue)
//! End Function
//!
//! ' Usage
//! Debug.Print "Population: " & FormatPopulation(8500000)  ' Population: 8,500,000
//! ```
//!
//! ### Display Measurement with Precision
//!
//! ```vb
//! Function FormatMeasurement(value As Double, decimals As Integer, _
//!                            unit As String) As String
//!     FormatMeasurement = FormatNumber(value, decimals) & " " & unit
//! End Function
//!
//! ' Usage
//! Debug.Print FormatMeasurement(12.5678, 2, "cm")  ' 12.57 cm
//! Debug.Print FormatMeasurement(98.6, 1, "°F")     ' 98.6 °F
//! ```
//!
//! ### Format Grid/Report Data
//!
//! ```vb
//! Sub PopulateDataGrid(grid As MSFlexGrid, values() As Double)
//!     Dim i As Long
//!     
//!     grid.Rows = UBound(values) + 2
//!     grid.TextMatrix(0, 0) = "Index"
//!     grid.TextMatrix(0, 1) = "Value"
//!     
//!     For i = LBound(values) To UBound(values)
//!         grid.TextMatrix(i + 1, 0) = FormatNumber(i, 0)
//!         grid.TextMatrix(i + 1, 1) = FormatNumber(values(i), 2)
//!     Next i
//! End Sub
//! ```
//!
//! ### Display Percentage (without symbol)
//!
//! ```vb
//! Function FormatPercentValue(value As Double, decimals As Integer) As String
//!     ' Convert to percentage value (multiply by 100)
//!     FormatPercentValue = FormatNumber(value * 100, decimals)
//! End Function
//!
//! ' Usage
//! Debug.Print FormatPercentValue(0.1234, 2)    ' 12.34
//! Debug.Print FormatPercentValue(0.5, 0)       ' 50
//! ```
//!
//! ### Format Score/Rating Display
//!
//! ```vb
//! Function FormatScore(score As Double, maxScore As Double) As String
//!     Dim percentage As Double
//!     percentage = (score / maxScore) * 100
//!     
//!     FormatScore = FormatNumber(score, 1) & " / " & _
//!                   FormatNumber(maxScore, 0) & " (" & _
//!                   FormatNumber(percentage, 1) & "%)"
//! End Function
//!
//! ' Usage
//! Debug.Print FormatScore(87.5, 100)  ' 87.5 / 100 (87.5%)
//! ```
//!
//! ### Display Large Numbers with Suffixes
//!
//! ```vb
//! Function FormatLargeNumber(value As Double) As String
//!     Const Million = 1000000
//!     Const Billion = 1000000000
//!     
//!     If Abs(value) >= Billion Then
//!         FormatLargeNumber = FormatNumber(value / Billion, 2) & "B"
//!     ElseIf Abs(value) >= Million Then
//!         FormatLargeNumber = FormatNumber(value / Million, 2) & "M"
//!     ElseIf Abs(value) >= 1000 Then
//!         FormatLargeNumber = FormatNumber(value / 1000, 2) & "K"
//!     Else
//!         FormatLargeNumber = FormatNumber(value, 2)
//!     End If
//! End Function
//!
//! ' Usage
//! Debug.Print FormatLargeNumber(1500000)       ' 1.50M
//! Debug.Print FormatLargeNumber(2500000000)    ' 2.50B
//! ```
//!
//! ### Format Comparison Display
//!
//! ```vb
//! Function FormatComparison(actual As Double, expected As Double) As String
//!     Dim difference As Double
//!     Dim percentDiff As Double
//!     
//!     difference = actual - expected
//!     If expected <> 0 Then
//!         percentDiff = (difference / expected) * 100
//!     End If
//!     
//!     FormatComparison = "Actual: " & FormatNumber(actual, 2) & vbCrLf & _
//!                        "Expected: " & FormatNumber(expected, 2) & vbCrLf & _
//!                        "Difference: " & FormatNumber(difference, 2, , vbTrue) & vbCrLf & _
//!                        "% Difference: " & FormatNumber(percentDiff, 2, , vbTrue)
//! End Function
//! ```
//!
//! ### `ListBox` Population with Numbers
//!
//! ```vb
//! Sub PopulateNumberList(lst As ListBox, values() As Double, decimals As Integer)
//!     Dim i As Long
//!     
//!     lst.Clear
//!     
//!     For i = LBound(values) To UBound(values)
//!         lst.AddItem FormatNumber(values(i), decimals)
//!     Next i
//! End Sub
//! ```
//!
//! ### Format Database Numeric Display
//!
//! ```vb
//! Function GetFormattedNumber(rs As ADODB.Recordset, fieldName As String, _
//!                             Optional decimals As Integer = 2) As String
//!     If IsNull(rs.Fields(fieldName).Value) Then
//!         GetFormattedNumber = "N/A"
//!     Else
//!         GetFormattedNumber = FormatNumber(rs.Fields(fieldName).Value, decimals)
//!     End If
//! End Function
//! ```
//!
//! ### Display Summary Totals
//!
//! ```vb
//! Sub DisplaySummary(items As Collection)
//!     Dim item As Variant
//!     Dim count As Long
//!     Dim total As Double
//!     Dim average As Double
//!     Dim maxVal As Double
//!     Dim minVal As Double
//!     
//!     count = items.Count
//!     total = 0
//!     maxVal = -1E+308  ' Smallest possible Double
//!     minVal = 1E+308   ' Largest possible Double
//!     
//!     For Each item In items
//!         total = total + item.Value
//!         If item.Value > maxVal Then maxVal = item.Value
//!         If item.Value < minVal Then minVal = item.Value
//!     Next item
//!     
//!     average = total / count
//!     
//!     Debug.Print "Summary Statistics"
//!     Debug.Print String(50, "=")
//!     Debug.Print "Count:   ", FormatNumber(count, 0)
//!     Debug.Print "Total:   ", FormatNumber(total, 2)
//!     Debug.Print "Average: ", FormatNumber(average, 2)
//!     Debug.Print "Maximum: ", FormatNumber(maxVal, 2)
//!     Debug.Print "Minimum: ", FormatNumber(minVal, 2)
//! End Sub
//! ```
//!
//! ## Advanced Usage
//!
//! ### Flexible Number Formatter
//!
//! ```vb
//! Function FormatNumberEx(value As Double, _
//!                         Optional decimals As Integer = 2, _
//!                         Optional useParens As Boolean = False, _
//!                         Optional useGroups As Boolean = True) As String
//!     Dim leadingDigit As VbTriState
//!     Dim parens As VbTriState
//!     Dim groups As VbTriState
//!     
//!     leadingDigit = vbTrue
//!     parens = IIf(useParens, vbTrue, vbFalse)
//!     groups = IIf(useGroups, vbTrue, vbFalse)
//!     
//!     FormatNumberEx = FormatNumber(value, decimals, leadingDigit, parens, groups)
//! End Function
//! ```
//!
//! ### Accounting-Style Formatter
//!
//! ```vb
//! Function FormatAccounting(value As Double, Optional decimals As Integer = 2) As String
//!     ' Always show parentheses for negatives, use grouping
//!     FormatAccounting = FormatNumber(value, decimals, vbTrue, vbTrue, vbTrue)
//! End Function
//!
//! ' Usage
//! Debug.Print FormatAccounting(1234.56)    ' 1,234.56
//! Debug.Print FormatAccounting(-789.12)    ' (789.12)
//! ```
//!
//! ### Dynamic Precision Formatter
//!
//! ```vb
//! Function FormatNumberDynamic(value As Double) As String
//!     ' Adjust precision based on magnitude
//!     If Abs(value) >= 1000 Then
//!         ' Large numbers: no decimals
//!         FormatNumberDynamic = FormatNumber(value, 0)
//!     ElseIf Abs(value) >= 1 Then
//!         ' Regular: 2 decimals
//!         FormatNumberDynamic = FormatNumber(value, 2)
//!     ElseIf Abs(value) >= 0.01 Then
//!         ' Small: 4 decimals
//!         FormatNumberDynamic = FormatNumber(value, 4)
//!     Else
//!         ' Very small: 6 decimals
//!         FormatNumberDynamic = FormatNumber(value, 6)
//!     End If
//! End Function
//! ```
//!
//! ### Table/Report Alignment
//!
//! ```vb
//! Function FormatNumberAligned(value As Double, width As Integer, _
//!                              Optional decimals As Integer = 2) As String
//!     Dim formatted As String
//!     formatted = FormatNumber(value, decimals, vbTrue, vbTrue, vbTrue)
//!     
//!     ' Right-align in field
//!     If Len(formatted) < width Then
//!         FormatNumberAligned = Space(width - Len(formatted)) & formatted
//!     Else
//!         FormatNumberAligned = formatted
//!     End If
//! End Function
//! ```
//!
//! ### Scientific Data Formatter
//!
//! ```vb
//! Function FormatScientificValue(value As Double, _
//!                                significantDigits As Integer) As String
//!     If Abs(value) >= 1000 Or Abs(value) < 0.01 Then
//!         ' Use scientific notation for very large/small
//!         FormatScientificValue = Format(value, "0." & String(significantDigits - 1, "0") & "E+00")
//!     Else
//!         ' Use regular formatting
//!         FormatScientificValue = FormatNumber(value, significantDigits)
//!     End If
//! End Function
//! ```
//!
//! ### Conditional Formatting
//!
//! ```vb
//! Function FormatNumberConditional(value As Double) As String
//!     If value > 0 Then
//!         FormatNumberConditional = "+" & FormatNumber(value, 2)
//!     ElseIf value < 0 Then
//!         FormatNumberConditional = FormatNumber(value, 2, , vbTrue)
//!     Else
//!         FormatNumberConditional = FormatNumber(0, 2)
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeFormatNumber(value As Variant, _
//!                           Optional decimals As Integer = 2) As String
//!     On Error GoTo ErrorHandler
//!     
//!     If IsNull(value) Then
//!         SafeFormatNumber = "N/A"
//!     ElseIf Not IsNumeric(value) Then
//!         SafeFormatNumber = "Invalid"
//!     Else
//!         SafeFormatNumber = FormatNumber(CDbl(value), decimals)
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 13  ' Type mismatch
//!             SafeFormatNumber = "Type Error"
//!         Case 6   ' Overflow
//!             SafeFormatNumber = "Overflow"
//!         Case 5   ' Invalid procedure call
//!             SafeFormatNumber = "Invalid"
//!         Case Else
//!             SafeFormatNumber = "Error"
//!     End Select
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 13** (Type Mismatch): Expression cannot be converted to numeric
//! - **Error 6** (Overflow): Value too large for Double
//! - **Error 5** (Invalid procedure call): Invalid decimal places parameter
//!
//! ## Performance Considerations
//!
//! - `FormatNumber` is fast for simple formatting
//! - Slightly slower than `Format` for custom patterns
//! - Faster than building format strings manually
//! - Locale lookups cached by system
//! - Avoid repeated calls in tight loops if possible
//! - Consider caching formatted values for display
//!
//! ## Best Practices
//!
//! ### Use `FormatNumber` for General Numeric Display
//!
//! ```vb
//! ' Good - Locale-aware, user-friendly
//! lblValue.Caption = FormatNumber(total, 2)
//!
//! ' Less portable - Hard-coded format
//! lblValue.Caption = Format(total, "0.00")
//! ```
//!
//! ### Handle Null Values
//!
//! ```vb
//! ' Good - Check for Null
//! If Not IsNull(value) Then
//!     formatted = FormatNumber(value, 2)
//! Else
//!     formatted = "N/A"
//! End If
//! ```
//!
//! ### Be Consistent with Decimal Places
//!
//! ```vb
//! ' Good - Use constants for consistency
//! Const DECIMAL_PLACES = 2
//! result1 = FormatNumber(value1, DECIMAL_PLACES)
//! result2 = FormatNumber(value2, DECIMAL_PLACES)
//! ```
//!
//! ## Comparison with Other Functions
//!
//! ### `FormatNumber` vs `FormatCurrency`
//!
//! ```vb
//! ' FormatNumber - No currency symbol
//! result = FormatNumber(1234.56)          ' 1,234.56
//!
//! ' FormatCurrency - Adds currency symbol
//! result = FormatCurrency(1234.56)        ' $1,234.56
//! ```
//!
//! ### `FormatNumber` vs `Format`
//!
//! ```vb
//! ' FormatNumber - Simple, predefined
//! result = FormatNumber(1234.56, 2)
//!
//! ' Format - More control, custom patterns
//! result = Format(1234.56, "#,##0.00")
//! ```
//!
//! ### `FormatNumber` vs `FormatPercent`
//!
//! ```vb
//! ' FormatNumber - No percent symbol, no multiplication
//! result = FormatNumber(0.1234 * 100, 2)  ' 12.34
//!
//! ' FormatPercent - Multiplies by 100, adds %
//! result = FormatPercent(0.1234)          ' 12.34%
//! ```
//!
//! ### `FormatNumber` vs `Str`/`CStr`
//!
//! ```vb
//! ' FormatNumber - Full formatting
//! result = FormatNumber(1234.56)          ' 1,234.56
//!
//! ' Str - Basic conversion, no formatting
//! result = Str(1234.56)                   ' " 1234.56"
//!
//! ' CStr - Basic conversion
//! result = CStr(1234.56)                  ' "1234.56"
//! ```
//!
//! ## Limitations
//!
//! - Uses system locale (cannot specify different locale)
//! - All parameters optional, making errors less obvious
//! - Tristate parameters can be confusing
//! - No built-in rounding mode control
//! - Cannot format with custom patterns
//! - No control over decimal/thousand separators
//! - Limited to numeric types
//!
//! ## Regional Settings Impact
//!
//! The `FormatNumber` function behavior varies by locale:
//!
//! - **United States**: 1,234.56
//! - **European Union**: 1.234,56 (note decimal/thousand separators swapped)
//! - **Switzerland**: 1'234.56 (apostrophe as separator)
//! - **India**: 12,34,567.89 (different grouping pattern)
//!
//! ## Related Functions
//!
//! - `Format`: More flexible formatting with custom patterns
//! - `FormatCurrency`: Format numbers as currency
//! - `FormatPercent`: Format numbers as percentages
//! - `FormatDateTime`: Format date/time values
//! - `Round`: Round numbers to specified decimal places
//! - `Int`: Return integer portion of a number
//! - `CDbl`: Convert expression to Double
//! - `IsNumeric`: Check if expression can be converted to numeric

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_formatnumber_basic() {
        let source = r#"
result = FormatNumber(value)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_decimals() {
        let source = r#"
formatted = FormatNumber(value, 2)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_no_decimals() {
        let source = r#"
formatted = FormatNumber(value, 0)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_parens() {
        let source = r#"
result = FormatNumber(negative, 2, , vbTrue)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_all_params() {
        let source = r#"
formatted = FormatNumber(value, 2, vbTrue, vbTrue, vbTrue)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_debug_print() {
        let source = r#"
Debug.Print FormatNumber(total, 2)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_concatenation() {
        let source = r#"
msg = "Total: " & FormatNumber(total, 2)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_in_function() {
        let source = r#"
Function FormatPopulation(population As Long) As String
    FormatPopulation = FormatNumber(population, 0, , , vbTrue)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_grid() {
        let source = r#"
grid.TextMatrix(i + 1, 0) = FormatNumber(i, 0)
grid.TextMatrix(i + 1, 1) = FormatNumber(values(i), 2)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_multiplication() {
        let source = r#"
result = FormatNumber(value * 100, 2)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_division() {
        let source = r#"
formatted = FormatNumber(value / Billion, 2) & "B"
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_listbox() {
        let source = r#"
lst.AddItem FormatNumber(values(i), decimals)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_multiline() {
        let source = r#"
summary = "Actual: " & FormatNumber(actual, 2) & vbCrLf & _
          "Expected: " & FormatNumber(expected, 2) & vbCrLf & _
          "Difference: " & FormatNumber(difference, 2, , vbTrue)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_if_statement() {
        let source = r#"
If Abs(value) >= 1000 Then
    result = FormatNumber(value, 0)
Else
    result = FormatNumber(value, 2)
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_isnull_check() {
        let source = r#"
If Not IsNull(value) Then
    formatted = FormatNumber(value, 2)
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_isnumeric_check() {
        let source = r#"
If IsNumeric(value) Then
    result = FormatNumber(CDbl(value), decimals)
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_error_handling() {
        let source = r#"
On Error GoTo ErrorHandler
formatted = FormatNumber(CDbl(value), decimals)
Exit Function
ErrorHandler:
    formatted = "Error"
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_for_loop() {
        let source = r#"
For i = LBound(values) To UBound(values)
    Debug.Print FormatNumber(values(i), 2)
Next i
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_select_case() {
        let source = r#"
Select Case Abs(value)
    Case Is >= 1000
        result = FormatNumber(value, 0)
    Case Else
        result = FormatNumber(value, 2)
End Select
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_comparison() {
        let source = r#"
comparison = FormatNumber(value1, 2) & " vs " & FormatNumber(value2, 2)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_vbfalse() {
        let source = r#"
result = FormatNumber(fraction, 2, vbFalse)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_grouping() {
        let source = r#"
formatted = FormatNumber(largeNumber, 2, , , vbTrue)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_recordset() {
        let source = r#"
formatted = FormatNumber(rs.Fields(fieldName).Value, decimals)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_iif() {
        let source = r#"
parens = IIf(useParens, vbTrue, vbFalse)
result = FormatNumber(value, decimals, vbTrue, parens, vbTrue)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_with_space() {
        let source = r#"
result = FormatNumber(value, decimals) & " " & unit
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatnumber_subtraction() {
        let source = r#"
difference = FormatNumber(actual - expected, 2, , vbTrue)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatNumber"));
        assert!(debug.contains("Identifier"));
    }
}
