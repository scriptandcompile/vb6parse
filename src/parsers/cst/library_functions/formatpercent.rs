//! FormatPercent Function
//!
//! Returns an expression formatted as a percentage (multiplied by 100) with a trailing % character.
//!
//! # Syntax
//!
//! ```vb
//! FormatPercent(expression[, numdigitsafterdecimal[, includeleadingdigit[, useparensfornegativenumbers[, groupdigits]]]])
//! ```
//!
//! # Parameters
//!
//! - `expression` - Required. Expression to be formatted as a percentage.
//! - `numdigitsafterdecimal` - Optional. Numeric value indicating how many places to the right of the decimal are displayed.
//!   Default value is -1, which indicates that the computer's regional settings are used.
//! - `includeleadingdigit` - Optional. Tristate constant that indicates whether or not a leading zero is displayed for fractional values.
//!   See Settings section for values.
//! - `useparensfornegativenumbers` - Optional. Tristate constant that indicates whether or not to place negative values within parentheses.
//!   See Settings section for values.
//! - `groupdigits` - Optional. Tristate constant that indicates whether or not numbers are grouped using the group delimiter
//!   specified in the computer's regional settings. See Settings section for values.
//!
//! # Settings
//!
//! The `includeleadingdigit`, `useparensfornegativenumbers`, and `groupdigits` arguments have the following settings:
//!
//! | Constant | Value | Description |
//! |----------|-------|-------------|
//! | vbTrue | -1 | True |
//! | vbFalse | 0 | False |
//! | vbUseDefault | -2 | Use the setting from the computer's regional settings |
//!
//! # Return Value
//!
//! Returns a Variant of subtype String containing the expression formatted as a percentage with a trailing % character.
//!
//! # Remarks
//!
//! - The FormatPercent function multiplies the numeric value by 100 and appends a percent sign.
//! - When one or more optional arguments are omitted, the computer's regional settings provide the default values.
//! - The number is formatted according to the computer's locale settings.
//! - If the expression is Null, FormatPercent returns an empty string.
//! - The default number of decimal places is 2 (from regional settings).
//!
//! # Typical Uses
//!
//! - Displaying statistical percentages (success rates, completion rates)
//! - Financial ratios (profit margins, growth rates, interest rates)
//! - Survey results and poll data
//! - Progress indicators and completion percentages
//! - Conversion rates and efficiency metrics
//! - Grade percentages and test scores
//!
//! # Basic Usage Examples
//!
//! ```vb
//! ' Default formatting (2 decimal places from regional settings)
//! result = FormatPercent(0.75)           ' Returns: "75.00%"
//! result = FormatPercent(0.5)            ' Returns: "50.00%"
//! result = FormatPercent(1.25)           ' Returns: "125.00%"
//!
//! ' Handling negative percentages
//! result = FormatPercent(-0.15)          ' Returns: "-15.00%"
//! result = FormatPercent(-0.05, 2, , vbTrue)  ' Returns: "(5.00)%"
//!
//! ' Controlling decimal places
//! result = FormatPercent(0.3333, 0)      ' Returns: "33%"
//! result = FormatPercent(0.3333, 2)      ' Returns: "33.33%"
//! result = FormatPercent(0.3333, 4)      ' Returns: "33.3300%"
//!
//! ' Leading digit control
//! result = FormatPercent(0.005, 2, vbTrue)   ' Returns: "0.50%"
//! result = FormatPercent(0.005, 2, vbFalse)  ' Returns: ".50%"
//! ```
//!
//! # Common Patterns
//!
//! ## 1. Success Rate Display
//!
//! ```vb
//! Dim successful As Long
//! Dim total As Long
//! Dim rate As Double
//!
//! successful = 85
//! total = 100
//! rate = successful / total
//!
//! MsgBox "Success Rate: " & FormatPercent(rate, 1)  ' "Success Rate: 85.0%"
//! ```
//!
//! ## 2. Financial Ratios
//!
//! ```vb
//! Dim profit As Currency
//! Dim revenue As Currency
//! Dim margin As Double
//!
//! profit = 25000
//! revenue = 100000
//! margin = profit / revenue
//!
//! lblMargin.Caption = "Profit Margin: " & FormatPercent(margin, 2)  ' "Profit Margin: 25.00%"
//! ```
//!
//! ## 3. Survey Results
//!
//! ```vb
//! Dim yesVotes As Long
//! Dim totalVotes As Long
//!
//! yesVotes = 347
//! totalVotes = 500
//!
//! txtResult.Text = FormatPercent(yesVotes / totalVotes, 1)  ' "69.4%"
//! ```
//!
//! ## 4. Progress Indicator
//!
//! ```vb
//! Dim completed As Long
//! Dim total As Long
//! Dim progress As Double
//!
//! completed = 45
//! total = 100
//! progress = completed / total
//!
//! lblProgress.Caption = "Progress: " & FormatPercent(progress, 0)  ' "Progress: 45%"
//! ```
//!
//! ## 5. Growth Rate Calculation
//!
//! ```vb
//! Dim currentValue As Double
//! Dim previousValue As Double
//! Dim growth As Double
//!
//! currentValue = 125000
//! previousValue = 100000
//! growth = (currentValue - previousValue) / previousValue
//!
//! MsgBox "Growth: " & FormatPercent(growth, 1)  ' "Growth: 25.0%"
//! ```
//!
//! ## 6. Grade Percentage
//!
//! ```vb
//! Dim score As Long
//! Dim maxScore As Long
//! Dim percentage As Double
//!
//! score = 87
//! maxScore = 100
//! percentage = score / maxScore
//!
//! lblGrade.Caption = "Score: " & FormatPercent(percentage, 0)  ' "Score: 87%"
//! ```
//!
//! ## 7. Comparison Display
//!
//! ```vb
//! Dim actual As Double
//! Dim target As Double
//! Dim achievement As Double
//!
//! actual = 95000
//! target = 100000
//! achievement = actual / target
//!
//! lblStatus.Caption = "Target Achievement: " & FormatPercent(achievement, 1)  ' "Target Achievement: 95.0%"
//! ```
//!
//! ## 8. ListBox Population with Percentages
//!
//! ```vb
//! Dim i As Integer
//! Dim values() As Double
//! Dim total As Double
//!
//! values = Array(15.5, 28.3, 42.1, 14.1)
//! total = 100
//!
//! For i = 0 To UBound(values)
//!     lstResults.AddItem FormatPercent(values(i) / total, 1)
//! Next i
//! ```
//!
//! ## 9. Database Field Display
//!
//! ```vb
//! Dim rs As ADODB.Recordset
//! Set rs = New ADODB.Recordset
//!
//! rs.Open "SELECT CompletionRate FROM Projects", conn
//!
//! While Not rs.EOF
//!     Debug.Print FormatPercent(rs("CompletionRate"), 0)
//!     rs.MoveNext
//! Wend
//! ```
//!
//! ## 10. Multiple Percentages in Report
//!
//! ```vb
//! Dim passed As Long, failed As Long, total As Long
//! Dim report As String
//!
//! passed = 85
//! failed = 15
//! total = passed + failed
//!
//! report = "Passed: " & FormatPercent(passed / total, 1) & vbCrLf & _
//!          "Failed: " & FormatPercent(failed / total, 1)
//! ' "Passed: 85.0%
//! '  Failed: 15.0%"
//! ```
//!
//! # Advanced Usage
//!
//! ## 1. Flexible Percentage Formatter
//!
//! ```vb
//! Function DisplayPercentage(value As Double, Optional precision As Integer = 1) As String
//!     If value < 0.01 Then
//!         DisplayPercentage = FormatPercent(value, 3)  ' More precision for small values
//!     ElseIf value > 1 Then
//!         DisplayPercentage = FormatPercent(value, 0)  ' No decimals for large percentages
//!     Else
//!         DisplayPercentage = FormatPercent(value, precision)
//!     End If
//! End Function
//! ```
//!
//! ## 2. Comparison with Color Coding
//!
//! ```vb
//! Function FormatVariance(actual As Double, target As Double) As String
//!     Dim variance As Double
//!     variance = (actual - target) / target
//!     
//!     FormatVariance = FormatPercent(variance, 1)
//!     
//!     If variance > 0 Then
//!         lblVariance.ForeColor = vbGreen  ' Positive variance
//!     ElseIf variance < 0 Then
//!         lblVariance.ForeColor = vbRed    ' Negative variance
//!     End If
//! End Function
//! ```
//!
//! ## 3. Dynamic Precision Based on Value
//!
//! ```vb
//! Function SmartFormatPercent(value As Double) As String
//!     Dim absValue As Double
//!     absValue = Abs(value)
//!     
//!     Select Case absValue
//!         Case Is < 0.0001
//!             SmartFormatPercent = FormatPercent(value, 4)
//!         Case Is < 0.01
//!             SmartFormatPercent = FormatPercent(value, 3)
//!         Case Is < 1
//!             SmartFormatPercent = FormatPercent(value, 2)
//!         Case Else
//!             SmartFormatPercent = FormatPercent(value, 1)
//!     End Select
//! End Function
//! ```
//!
//! ## 4. Table Alignment with Percentages
//!
//! ```vb
//! Function AlignedPercentage(value As Double, width As Integer) As String
//!     Dim formatted As String
//!     formatted = FormatPercent(value, 2)
//!     AlignedPercentage = Space(width - Len(formatted)) & formatted
//! End Function
//!
//! ' Usage in grid or report
//! For i = 1 To 10
//!     Debug.Print AlignedPercentage(data(i), 12)
//! Next i
//! ```
//!
//! ## 5. Statistical Analysis Display
//!
//! ```vb
//! Type StatisticsResult
//!     Mean As Double
//!     StdDev As Double
//!     Confidence As Double
//! End Type
//!
//! Function FormatStatistics(stats As StatisticsResult) As String
//!     FormatStatistics = "Mean: " & FormatPercent(stats.Mean, 2) & vbCrLf & _
//!                       "Std Dev: " & FormatPercent(stats.StdDev, 2) & vbCrLf & _
//!                       "Confidence: " & FormatPercent(stats.Confidence, 1)
//! End Function
//! ```
//!
//! ## 6. Conditional Formatting for Thresholds
//!
//! ```vb
//! Sub DisplayMetricWithThreshold(metric As Double, threshold As Double, lbl As Label)
//!     lbl.Caption = FormatPercent(metric, 1)
//!     
//!     If metric >= threshold Then
//!         lbl.BackColor = vbGreen
//!         lbl.ForeColor = vbWhite
//!     ElseIf metric >= threshold * 0.8 Then
//!         lbl.BackColor = vbYellow
//!         lbl.ForeColor = vbBlack
//!     Else
//!         lbl.BackColor = vbRed
//!         lbl.ForeColor = vbWhite
//!     End If
//! End Sub
//! ```
//!
//! # Error Handling
//!
//! ```vb
//! Function SafeFormatPercent(value As Variant) As String
//!     On Error GoTo ErrorHandler
//!     
//!     If IsNull(value) Then
//!         SafeFormatPercent = "N/A"
//!         Exit Function
//!     End If
//!     
//!     If Not IsNumeric(value) Then
//!         SafeFormatPercent = "Invalid"
//!         Exit Function
//!     End If
//!     
//!     SafeFormatPercent = FormatPercent(CDbl(value), 2)
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeFormatPercent = "Error"
//! End Function
//! ```
//!
//! Common errors:
//! - **Error 13 (Type mismatch)**: Occurs when expression cannot be converted to a numeric value.
//! - **Error 6 (Overflow)**: Can occur with extremely large or small values.
//! - **Error 5 (Invalid procedure call)**: May occur with invalid argument values.
//!
//! # Performance Considerations
//!
//! - FormatPercent is relatively fast for single values but can be optimized in loops
//! - For large datasets, consider caching formatted values
//! - The multiplication by 100 is automatic and efficient
//! - Regional settings lookup may have slight overhead
//!
//! # Best Practices
//!
//! 1. Use appropriate decimal places for your context (financial: 2, statistics: 1, progress: 0)
//! 2. Consider the audience when choosing precision
//! 3. Handle division by zero before calling FormatPercent
//! 4. Use consistent formatting throughout your application
//! 5. Remember that the input is a decimal (0.5 = 50%)
//! 6. Consider using parentheses for negative values in financial contexts
//!
//! # Comparison with Other Functions
//!
//! - **FormatNumber**: Does not multiply by 100 or add %, gives more control over formatting
//! - **FormatCurrency**: Adds currency symbol instead of %, doesn't multiply by 100
//! - **Format with "%"**: More flexible but requires manual multiplication by 100
//! - **Str/CStr**: No automatic multiplication or % symbol, no locale support
//!
//! # Limitations
//!
//! - Always multiplies by 100 (cannot be disabled)
//! - Always adds trailing % character
//! - Cannot customize the position of the % symbol
//! - Limited control over negative number formatting compared to custom Format strings
//! - Depends on regional settings which may vary across systems
//!
//! # Regional Settings Impact
//!
//! The appearance of formatted percentages varies by locale:
//!
//! - **US (English)**: 75.50%
//! - **European (many)**: 75,50%
//! - **Switzerland**: 75.50%
//!
//! # Related Functions
//!
//! - `FormatNumber` - Formats numbers without percentage conversion
//! - `FormatCurrency` - Formats as currency with symbol
//! - `Format` - General-purpose formatting function
//! - `CDbl` - Converts to Double for percentage calculations
//! - `Round` - Rounds numbers before percentage formatting

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_formatpercent_basic() {
        let source = r#"result = FormatPercent(0.75)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_with_decimals() {
        let source = r#"result = FormatPercent(0.3333, 2)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_no_decimals() {
        let source = r#"result = FormatPercent(0.87, 0)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_with_parentheses() {
        let source = r#"result = FormatPercent(-0.05, 2, , vbTrue)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_all_parameters() {
        let source = r#"result = FormatPercent(0.125, 1, vbTrue, vbFalse, vbTrue)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_in_debug_print() {
        let source = r#"Debug.Print FormatPercent(0.5, 0)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_concatenation() {
        let source = r#"MsgBox "Success Rate: " & FormatPercent(rate, 1)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_in_function() {
        let source = r#"lblMargin.Caption = "Profit Margin: " & FormatPercent(margin, 2)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_division() {
        let source = r#"txtResult.Text = FormatPercent(yesVotes / totalVotes, 1)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_calculation() {
        let source = r#"MsgBox "Growth: " & FormatPercent((currentValue - previousValue) / previousValue, 1)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_assignment() {
        let source = r#"lblGrade.Caption = "Score: " & FormatPercent(score / maxScore, 0)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_listbox() {
        let source = r#"lstResults.AddItem FormatPercent(values(i) / total, 1)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_multiline() {
        let source = r#"report = "Passed: " & FormatPercent(passed / total, 1) & vbCrLf & "Failed: " & FormatPercent(failed / total, 1)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_in_if() {
        let source = r#"If value < 0.01 Then result = FormatPercent(value, 3)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_null_check() {
        let source =
            r#"If IsNull(value) Then Exit Function Else result = FormatPercent(CDbl(value), 2)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_numeric_check() {
        let source =
            r#"If Not IsNumeric(value) Then Exit Function Else result = FormatPercent(value, 2)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_error_handling() {
        let source = r#"On Error GoTo ErrorHandler
result = FormatPercent(value, 2)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_for_loop() {
        let source = r#"For i = 1 To 10
    Debug.Print FormatPercent(data(i), 2)
Next i"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_select_case() {
        let source = r#"Select Case absValue
    Case Is < 0.01
        result = FormatPercent(value, 3)
End Select"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_comparison() {
        let source = r#"lblStatus.Caption = "Achievement: " & FormatPercent(actual / target, 1)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_vbfalse() {
        let source = r#"result = FormatPercent(0.005, 2, vbFalse)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_vbtrue() {
        let source = r#"result = FormatPercent(0.005, 2, vbTrue)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_recordset() {
        let source = r#"Debug.Print FormatPercent(rs("CompletionRate"), 0)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_iif() {
        let source = r#"result = IIf(value > 1, FormatPercent(value, 0), FormatPercent(value, 2))"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_negative() {
        let source = r#"result = FormatPercent(-0.15)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatpercent_large_value() {
        let source = r#"result = FormatPercent(1.25, 2)"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatPercent"));
        assert!(debug.contains("Identifier"));
    }
}
