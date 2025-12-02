//! # `FormatCurrency` Function
//!
//! Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.
//!
//! ## Syntax
//!
//! ```vb
//! FormatCurrency(expression[, numdigitsafterdecimal[, includeleadingdigit[, useparensfornegativenumbers[, groupdigits]]]])
//! ```
//!
//! ## Parameters
//!
//! - **expression**: Required. Expression to be formatted.
//! - **numdigitsafterdecimal**: Optional. Numeric value indicating how many places to the right of the decimal are displayed. Default is -1, which indicates the computer's regional settings are used.
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
//! Returns a `Variant` of subtype `String` containing the expression formatted as a currency value.
//!
//! ## Remarks
//!
//! The `FormatCurrency` function provides a simple way to format numbers as currency values
//! using the system's locale settings. It automatically applies the currency symbol, decimal
//! separator, thousand separator, and negative number formatting according to regional settings.
//!
//! **Important Characteristics:**
//!
//! - Uses system locale for currency symbol and formatting
//! - Default: 2 decimal places (from regional settings)
//! - Automatically adds thousand separators
//! - Negative numbers can be displayed with parentheses or minus sign
//! - Leading zeros controlled by regional settings or parameter
//! - Currency symbol position depends on locale (before or after amount)
//! - Returns empty string if expression is Null
//! - More convenient than `Format` for simple currency formatting
//! - Less flexible than `Format` for custom patterns
//! - Locale-aware (respects user's regional settings)
//!
//! ## Typical Uses
//!
//! - Display prices and monetary amounts
//! - Format financial reports
//! - Show invoice totals
//! - Display account balances
//! - Format transaction amounts
//! - Create currency-formatted exports
//! - Display budget figures
//! - Show cost calculations
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! Dim amount As Double
//! amount = 1234.567
//!
//! ' Default formatting (2 decimal places, system settings)
//! Debug.Print FormatCurrency(amount)           ' $1,234.57
//!
//! ' No decimal places
//! Debug.Print FormatCurrency(amount, 0)        ' $1,235
//!
//! ' Three decimal places
//! Debug.Print FormatCurrency(amount, 3)        ' $1,234.567
//!
//! ' Negative with parentheses
//! Debug.Print FormatCurrency(-500, , , vbTrue) ' ($500.00)
//! ```
//!
//! ### Handling Negative Values
//!
//! ```vb
//! Dim balance As Double
//! balance = -1250.50
//!
//! ' Default negative (with minus sign)
//! Debug.Print FormatCurrency(balance)          ' -$1,250.50
//!
//! ' Parentheses for negative
//! Debug.Print FormatCurrency(balance, 2, , vbTrue)  ' ($1,250.50)
//!
//! ' No parentheses (explicit)
//! Debug.Print FormatCurrency(balance, 2, , vbFalse) ' -$1,250.50
//! ```
//!
//! ### Control Leading Digits
//!
//! ```vb
//! Dim fraction As Double
//! fraction = 0.75
//!
//! ' With leading zero (default)
//! Debug.Print FormatCurrency(fraction)         ' $0.75
//!
//! ' No leading zero
//! Debug.Print FormatCurrency(fraction, 2, vbFalse)  ' $.75
//!
//! ' Explicit leading zero
//! Debug.Print FormatCurrency(fraction, 2, vbTrue)   ' $0.75
//! ```
//!
//! ## Common Patterns
//!
//! ### Format Invoice Line Items
//!
//! ```vb
//! Sub DisplayInvoiceLines(items As Collection)
//!     Dim item As Variant
//!     Dim total As Double
//!     
//!     Debug.Print "Item", "Quantity", "Price", "Amount"
//!     Debug.Print String(60, "-")
//!     
//!     total = 0
//!     For Each item In items
//!         Debug.Print item.Name, _
//!                     item.Quantity, _
//!                     FormatCurrency(item.Price), _
//!                     FormatCurrency(item.Quantity * item.Price)
//!         total = total + (item.Quantity * item.Price)
//!     Next item
//!     
//!     Debug.Print String(60, "-")
//!     Debug.Print "Total:", , , FormatCurrency(total)
//! End Sub
//! ```
//!
//! ### Format Account Balance with Parentheses
//!
//! ```vb
//! Function FormatAccountBalance(balance As Double) As String
//!     ' Show negative balances in parentheses (accounting style)
//!     FormatAccountBalance = FormatCurrency(balance, 2, vbTrue, vbTrue, vbTrue)
//! End Function
//!
//! ' Usage
//! Debug.Print FormatAccountBalance(1500)       ' $1,500.00
//! Debug.Print FormatAccountBalance(-250.50)    ' ($250.50)
//! ```
//!
//! ### Create Price Display
//!
//! ```vb
//! Function FormatPrice(price As Double, showCents As Boolean) As String
//!     If showCents Then
//!         FormatPrice = FormatCurrency(price, 2)
//!     Else
//!         FormatPrice = FormatCurrency(price, 0)
//!     End If
//! End Function
//!
//! ' Usage
//! Debug.Print FormatPrice(19.99, True)         ' $19.99
//! Debug.Print FormatPrice(19.99, False)        ' $20
//! ```
//!
//! ### Display Transaction Summary
//!
//! ```vb
//! Sub ShowTransactionSummary(credits As Double, debits As Double)
//!     Dim balance As Double
//!     
//!     Debug.Print "Transaction Summary"
//!     Debug.Print String(40, "=")
//!     Debug.Print "Credits:  ", FormatCurrency(credits)
//!     Debug.Print "Debits:   ", FormatCurrency(debits, 2, , vbTrue)
//!     Debug.Print String(40, "-")
//!     
//!     balance = credits - debits
//!     Debug.Print "Balance:  ", FormatCurrency(balance, 2, , vbTrue)
//! End Sub
//! ```
//!
//! ### Format Budget Report
//!
//! ```vb
//! Sub PrintBudgetReport()
//!     Dim budgeted As Double
//!     Dim actual As Double
//!     Dim variance As Double
//!     
//!     budgeted = 50000
//!     actual = 48500
//!     variance = actual - budgeted
//!     
//!     Debug.Print "Budget Report"
//!     Debug.Print String(50, "=")
//!     Debug.Print "Budgeted:", FormatCurrency(budgeted, 0)
//!     Debug.Print "Actual:  ", FormatCurrency(actual, 0)
//!     Debug.Print "Variance:", FormatCurrency(variance, 0, , vbTrue)
//!     
//!     If variance < 0 Then
//!         Debug.Print "Status: Under budget"
//!     Else
//!         Debug.Print "Status: Over budget"
//!     End If
//! End Sub
//! ```
//!
//! ### `ListBox`/`ComboBox` Population
//!
//! ```vb
//! Sub PopulatePriceList(lstPrices As ListBox, prices() As Double)
//!     Dim i As Long
//!     
//!     lstPrices.Clear
//!     
//!     For i = LBound(prices) To UBound(prices)
//!         lstPrices.AddItem FormatCurrency(prices(i))
//!     Next i
//! End Sub
//! ```
//!
//! ### Format for Database Display
//!
//! ```vb
//! Function GetFormattedPrice(rs As ADODB.Recordset, fieldName As String) As String
//!     If IsNull(rs.Fields(fieldName).Value) Then
//!         GetFormattedPrice = "N/A"
//!     Else
//!         GetFormattedPrice = FormatCurrency(rs.Fields(fieldName).Value)
//!     End If
//! End Function
//! ```
//!
//! ### Calculate and Display Tax
//!
//! ```vb
//! Function DisplayPriceWithTax(basePrice As Double, taxRate As Double) As String
//!     Dim tax As Double
//!     Dim total As Double
//!     
//!     tax = basePrice * taxRate
//!     total = basePrice + tax
//!     
//!     DisplayPriceWithTax = "Price: " & FormatCurrency(basePrice) & vbCrLf & _
//!                           "Tax: " & FormatCurrency(tax) & vbCrLf & _
//!                           "Total: " & FormatCurrency(total)
//! End Function
//!
//! ' Usage
//! MsgBox DisplayPriceWithTax(100, 0.08)
//! ```
//!
//! ### Format Payment Schedule
//!
//! ```vb
//! Sub ShowPaymentSchedule(loanAmount As Double, months As Integer, rate As Double)
//!     Dim payment As Double
//!     Dim i As Integer
//!     Dim balance As Double
//!     
//!     payment = loanAmount / months
//!     balance = loanAmount
//!     
//!     Debug.Print "Payment Schedule"
//!     Debug.Print String(50, "=")
//!     Debug.Print "Month", "Payment", "Balance"
//!     Debug.Print String(50, "-")
//!     
//!     For i = 1 To months
//!         Debug.Print i, FormatCurrency(payment), FormatCurrency(balance)
//!         balance = balance - payment
//!     Next i
//! End Sub
//! ```
//!
//! ### Compare Values
//!
//! ```vb
//! Function ComparePrices(price1 As Double, price2 As Double) As String
//!     Dim difference As Double
//!     difference = price1 - price2
//!     
//!     ComparePrices = FormatCurrency(price1) & " vs " & FormatCurrency(price2) & _
//!                     " (Difference: " & FormatCurrency(difference, 2, , vbTrue) & ")"
//! End Function
//! ```
//!
//! ### Shopping Cart Total
//!
//! ```vb
//! Function GetCartSummary(items As Collection) As String
//!     Dim item As Variant
//!     Dim subtotal As Double
//!     Dim tax As Double
//!     Dim shipping As Double
//!     Dim total As Double
//!     
//!     subtotal = 0
//!     For Each item In items
//!         subtotal = subtotal + (item.Price * item.Quantity)
//!     Next item
//!     
//!     tax = subtotal * 0.08
//!     shipping = 5.99
//!     total = subtotal + tax + shipping
//!     
//!     GetCartSummary = "Subtotal: " & FormatCurrency(subtotal) & vbCrLf & _
//!                      "Tax:      " & FormatCurrency(tax) & vbCrLf & _
//!                      "Shipping: " & FormatCurrency(shipping) & vbCrLf & _
//!                      "Total:    " & FormatCurrency(total)
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Flexible Currency Formatter
//!
//! ```vb
//! Function FormatCurrencyEx(amount As Double, _
//!                           Optional decimals As Integer = 2, _
//!                           Optional useParens As Boolean = True, _
//!                           Optional useGroups As Boolean = True) As String
//!     Dim leadingDigit As VbTriState
//!     Dim parens As VbTriState
//!     Dim groups As VbTriState
//!     
//!     leadingDigit = vbTrue
//!     parens = IIf(useParens, vbTrue, vbFalse)
//!     groups = IIf(useGroups, vbTrue, vbFalse)
//!     
//!     FormatCurrencyEx = FormatCurrency(amount, decimals, leadingDigit, parens, groups)
//! End Function
//! ```
//!
//! ### Multi-Currency Support
//!
//! ```vb
//! Function FormatMultiCurrency(amount As Double, currencyCode As String) As String
//!     ' Simple multi-currency display (uses FormatCurrency then replaces symbol)
//!     Dim formatted As String
//!     formatted = FormatCurrency(amount)
//!     
//!     Select Case UCase(currencyCode)
//!         Case "USD"
//!             ' Keep default $ symbol
//!             FormatMultiCurrency = formatted
//!         Case "EUR"
//!             FormatMultiCurrency = Replace(formatted, "$", "€")
//!         Case "GBP"
//!             FormatMultiCurrency = Replace(formatted, "$", "£")
//!         Case "JPY"
//!             FormatMultiCurrency = Replace(formatted, "$", "¥")
//!             FormatMultiCurrency = Replace(FormatMultiCurrency, ".00", "")
//!         Case Else
//!             FormatMultiCurrency = currencyCode & " " & Format(amount, "#,##0.00")
//!     End Select
//! End Function
//! ```
//!
//! ### Financial Statement Formatter
//!
//! ```vb
//! Type FinancialLine
//!     Description As String
//!     Amount As Double
//!     IsSubtotal As Boolean
//! End Type
//!
//! Function FormatFinancialStatement(lines() As FinancialLine) As String
//!     Dim result As String
//!     Dim i As Long
//!     Dim line As String
//!     
//!     result = "Financial Statement" & vbCrLf
//!     result = result & String(60, "=") & vbCrLf
//!     
//!     For i = LBound(lines) To UBound(lines)
//!         line = lines(i).Description
//!         
//!         ' Right-align amounts
//!         line = line & Space(40 - Len(line))
//!         line = line & FormatCurrency(lines(i).Amount, 2, vbTrue, vbTrue, vbTrue)
//!         
//!         If lines(i).IsSubtotal Then
//!             result = result & String(60, "-") & vbCrLf
//!         End If
//!         
//!         result = result & line & vbCrLf
//!     Next i
//!     
//!     FormatFinancialStatement = result
//! End Function
//! ```
//!
//! ### Dynamic Precision Formatter
//!
//! ```vb
//! Function FormatCurrencyDynamic(amount As Double) As String
//!     ' Use different precision based on amount magnitude
//!     If Abs(amount) >= 1000000 Then
//!         ' Millions: no decimals
//!         FormatCurrencyDynamic = FormatCurrency(amount, 0) & "M"
//!     ElseIf Abs(amount) >= 1000 Then
//!         ' Thousands: no decimals
//!         FormatCurrencyDynamic = FormatCurrency(amount, 0)
//!     ElseIf Abs(amount) >= 1 Then
//!         ' Regular: 2 decimals
//!         FormatCurrencyDynamic = FormatCurrency(amount, 2)
//!     Else
//!         ' Small amounts: 4 decimals
//!         FormatCurrencyDynamic = FormatCurrency(amount, 4)
//!     End If
//! End Function
//! ```
//!
//! ### Conditional Formatting
//!
//! ```vb
//! Function FormatProfitLoss(amount As Double) As String
//!     ' Format with color indicators for profit/loss
//!     If amount > 0 Then
//!         FormatProfitLoss = "[GREEN]+" & FormatCurrency(amount)
//!     ElseIf amount < 0 Then
//!         FormatProfitLoss = "[RED]" & FormatCurrency(amount, 2, , vbTrue)
//!     Else
//!         FormatProfitLoss = "[BLACK]" & FormatCurrency(0)
//!     End If
//! End Function
//! ```
//!
//! ### Grid/Report Alignment
//!
//! ```vb
//! Function FormatCurrencyAligned(amount As Double, width As Integer) As String
//!     Dim formatted As String
//!     formatted = FormatCurrency(amount, 2, vbTrue, vbTrue, vbTrue)
//!     
//!     ' Right-align in field
//!     If Len(formatted) < width Then
//!         FormatCurrencyAligned = Space(width - Len(formatted)) & formatted
//!     Else
//!         FormatCurrencyAligned = formatted
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeFormatCurrency(value As Variant, _
//!                             Optional decimals As Integer = 2) As String
//!     On Error GoTo ErrorHandler
//!     
//!     If IsNull(value) Then
//!         SafeFormatCurrency = "N/A"
//!     ElseIf Not IsNumeric(value) Then
//!         SafeFormatCurrency = "Invalid"
//!     Else
//!         SafeFormatCurrency = FormatCurrency(CDbl(value), decimals)
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 13  ' Type mismatch
//!             SafeFormatCurrency = "Type Error"
//!         Case 6   ' Overflow
//!             SafeFormatCurrency = "Overflow"
//!         Case Else
//!             SafeFormatCurrency = "Error"
//!     End Select
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 13** (Type Mismatch): Expression cannot be converted to numeric
//! - **Error 6** (Overflow): Value too large for `Double`
//! - **Error 5** (Invalid procedure call): Invalid decimal places parameter
//!
//! ## Performance Considerations
//!
//! - `FormatCurrency` is fast for simple formatting
//! - Slightly slower than `Format` for custom patterns
//! - Faster than building format strings manually
//! - Locale lookups cached by system
//! - Avoid repeated calls in tight loops if possible
//! - Consider caching formatted values for display
//!
//! ## Best Practices
//!
//! ### Use `FormatCurrency` for User-Facing Amounts
//!
//! ```vb
//! ' Good - Locale-aware, user-friendly
//! lblPrice.Caption = FormatCurrency(price)
//!
//! ' Less portable - Hard-coded format
//! lblPrice.Caption = "$" & Format(price, "0.00")
//! ```
//!
//! ### Handle `Null` Values
//!
//! ```vb
//! ' Good - Check for Null
//! If Not IsNull(amount) Then
//!     formatted = FormatCurrency(amount)
//! Else
//!     formatted = "N/A"
//! End If
//! ```
//!
//! ### Be Consistent with Negative Formatting
//!
//! ```vb
//! ' Good - Use same style throughout application
//! Const USE_PARENS = vbTrue
//! formatted = FormatCurrency(balance, 2, , USE_PARENS)
//! ```
//!
//! ## Comparison with Other Functions
//!
//! ### `FormatCurrency` vs `Format`
//!
//! ```vb
//! ' FormatCurrency - Simple, locale-aware
//! result = FormatCurrency(1234.56)
//!
//! ' Format - More control, custom patterns
//! result = Format(1234.56, "$#,##0.00")
//! ```
//!
//! ### `FormatCurrency` vs `FormatNumber`
//!
//! ```vb
//! ' FormatCurrency - Adds currency symbol
//! result = FormatCurrency(1234.56)        ' $1,234.56
//!
//! ' FormatNumber - No currency symbol
//! result = FormatNumber(1234.56)          ' 1,234.56
//! ```
//!
//! ### `FormatCurrency` vs `Str`/`CStr`
//!
//! ```vb
//! ' FormatCurrency - Full formatting
//! result = FormatCurrency(1234.56)        ' $1,234.56
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
//! - Limited to system's currency symbol
//! - Cannot customize symbol position
//! - All parameters optional, making errors less obvious
//! - Tristate parameters can be confusing
//! - No built-in rounding mode control
//! - Cannot format multiple currencies in same session
//!
//! ## Regional Settings Impact
//!
//! The `FormatCurrency` function behavior varies by locale:
//!
//! - **United States**: $1,234.56
//! - **United Kingdom**: £1,234.56
//! - **European Union**: €1.234,56 (note decimal/thousand separators)
//! - **Japan**: ¥1,235 (typically no decimal places)
//!
//! ## Related Functions
//!
//! - `Format`: More flexible formatting with custom patterns
//! - `FormatNumber`: Format numbers without currency symbol
//! - `FormatPercent`: Format numbers as percentages
//! - `FormatDateTime`: Format date/time values
//! - `CCur`: Convert expression to `Currency` type
//! - `CDbl`: Convert expression to `Double` type

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_formatcurrency_basic() {
        let source = r#"
result = FormatCurrency(amount)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_decimals() {
        let source = r#"
formatted = FormatCurrency(value, 2)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_no_decimals() {
        let source = r#"
formatted = FormatCurrency(amount, 0)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_parens() {
        let source = r#"
result = FormatCurrency(balance, 2, , vbTrue)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_all_params() {
        let source = r#"
formatted = FormatCurrency(amount, 2, vbTrue, vbTrue, vbTrue)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_debug_print() {
        let source = r#"
Debug.Print FormatCurrency(price)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_concatenation() {
        let source = r#"
msg = "Total: " & FormatCurrency(total)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_in_function() {
        let source = r#"
Function FormatAccountBalance(balance As Double) As String
    FormatAccountBalance = FormatCurrency(balance, 2, vbTrue, vbTrue, vbTrue)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_if_statement() {
        let source = r#"
If showCents Then
    result = FormatCurrency(price, 2)
Else
    result = FormatCurrency(price, 0)
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_listbox() {
        let source = r#"
lstPrices.AddItem FormatCurrency(prices(i))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_multiline() {
        let source = r#"
summary = "Subtotal: " & FormatCurrency(subtotal) & vbCrLf & _
          "Tax: " & FormatCurrency(tax) & vbCrLf & _
          "Total: " & FormatCurrency(total)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_calculation() {
        let source = r#"
result = FormatCurrency(price * quantity)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_isnull_check() {
        let source = r#"
If Not IsNull(amount) Then
    formatted = FormatCurrency(amount)
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_comparison() {
        let source = r#"
comparison = FormatCurrency(price1) & " vs " & FormatCurrency(price2)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_error_handling() {
        let source = r#"
On Error GoTo ErrorHandler
formatted = FormatCurrency(value, decimals)
Exit Function
ErrorHandler:
    formatted = "N/A"
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_for_loop() {
        let source = r#"
For i = 1 To itemCount
    Debug.Print FormatCurrency(amounts(i))
Next i
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_select_case() {
        let source = r#"
Select Case amount
    Case Is > 1000
        result = FormatCurrency(amount, 0)
    Case Else
        result = FormatCurrency(amount, 2)
End Select
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_msgbox() {
        let source = r#"
MsgBox "Total: " & FormatCurrency(total)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_label_caption() {
        let source = r#"
lblPrice.Caption = FormatCurrency(price)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_recordset() {
        let source = r#"
formatted = FormatCurrency(rs.Fields("Amount").Value)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_vbfalse() {
        let source = r#"
result = FormatCurrency(fraction, 2, vbFalse)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_negative() {
        let source = r#"
formatted = FormatCurrency(balance, 2, , vbFalse)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_subtraction() {
        let source = r#"
difference = FormatCurrency(price1 - price2, 2, , vbTrue)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_addition() {
        let source = r#"
total = FormatCurrency(subtotal + tax + shipping)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_with_cdbl() {
        let source = r#"
result = FormatCurrency(CDbl(value), decimals)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_formatcurrency_iif() {
        let source = r#"
parens = IIf(useParens, vbTrue, vbFalse)
result = FormatCurrency(amount, 2, vbTrue, parens, vbTrue)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatCurrency"));
        assert!(debug.contains("Identifier"));
    }
}
