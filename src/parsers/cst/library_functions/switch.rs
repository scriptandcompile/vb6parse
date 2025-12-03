//! VB6 `Switch` Function
//!
//! The `Switch` function evaluates a list of expressions and returns a value or expression associated with the first expression that is True.
//!
//! ## Syntax
//! ```vb6
//! Switch(expr1, value1[, expr2, value2 ... [, expr_n, value_n]])
//! ```
//!
//! ## Parameters
//! - `expr1, expr2, ..., expr_n`: Required. Variant expressions to be evaluated.
//! - `value1, value2, ..., value_n`: Required. Values or expressions to be returned. If the associated `expr` is True, `Switch` returns the corresponding value.
//!
//! The argument list consists of pairs of expressions and values. The expressions are evaluated from left to right, and the value associated with the first expression to evaluate to True is returned. If the pairs aren't properly matched, a run-time error occurs.
//!
//! ## Returns
//! Returns a `Variant` containing the value associated with the first True expression. If no expression is True, `Switch` returns `Null`.
//!
//! ## Remarks
//! The `Switch` function provides a flexible way to select from multiple alternatives:
//!
//! - **Left-to-right evaluation**: Expressions are evaluated in order until one is True
//! - **Short-circuit evaluation**: Once a True expression is found, remaining expressions are not evaluated
//! - **Null return**: If no expression evaluates to True, returns Null
//! - **Pairs required**: Arguments must come in pairs (expression, value). Odd number of arguments causes Error 450
//! - **Value can be expression**: The value part can be a literal, variable, or expression
//! - **All values evaluated**: Unlike expressions, all values in the argument list may be evaluated (implementation-dependent)
//! - **Type flexibility**: Can return different types for different cases (returns Variant)
//! - **Compare with Select Case**: Switch is an expression (returns value), Select Case is a statement
//! - **Compare with `IIf`**: Switch handles multiple conditions, `IIf` handles only two branches
//!
//! ### Evaluation Behavior
//! - Expressions are evaluated left to right
//! - First True expression stops evaluation of remaining expressions
//! - Values may be evaluated even if their associated expression is False (avoid side effects)
//! - Performance: O(n) where n is number of expression pairs
//!
//! ### When to Use Switch vs Alternatives
//! - **Use Switch** when you have multiple conditions to check and want to return a value
//! - **Use Select Case** for complex branching logic or when executing statements rather than returning values
//! - **Use `IIf`** for simple two-way decisions
//! - **Use nested `IIf`** cautiously (Switch is clearer for 3+ conditions)
//!
//! ## Typical Uses
//! 1. **Conditional Value Selection**: Return different values based on multiple conditions
//! 2. **Grade Calculation**: Assign letter grades based on numeric scores
//! 3. **Status Messages**: Return appropriate messages based on state
//! 4. **Categorization**: Categorize data based on multiple criteria
//! 5. **Default Values**: Provide values with fallback logic
//! 6. **Lookup Logic**: Implement simple lookup tables with conditions
//! 7. **Data Transformation**: Transform data based on multiple rules
//! 8. **Validation Messages**: Return validation results based on checks
//!
//! ## Basic Examples
//!
//! ### Example 1: Simple Value Selection
//! ```vb6
//! Dim result As Variant
//! Dim score As Integer
//!
//! score = 85
//!
//! result = Switch(score >= 90, "A", _
//!                 score >= 80, "B", _
//!                 score >= 70, "C", _
//!                 score >= 60, "D", _
//!                 True, "F")  ' Default case
//!
//! ' result = "B"
//! ```
//!
//! ### Example 2: Status Messages
//! ```vb6
//! Function GetStatusMessage(status As Integer) As String
//!     GetStatusMessage = Switch( _
//!         status = 0, "Idle", _
//!         status = 1, "Processing", _
//!         status = 2, "Complete", _
//!         status = 3, "Error", _
//!         True, "Unknown")
//! End Function
//! ```
//!
//! ### Example 3: Range-Based Categorization
//! ```vb6
//! Function CategorizeAge(age As Integer) As String
//!     CategorizeAge = Switch( _
//!         age < 13, "Child", _
//!         age < 20, "Teenager", _
//!         age < 65, "Adult", _
//!         True, "Senior")
//! End Function
//! ```
//!
//! ### Example 4: Multiple Condition Checks
//! ```vb6
//! Dim priority As String
//! Dim isUrgent As Boolean
//! Dim isImportant As Boolean
//! Dim hasDeadline As Boolean
//!
//! priority = Switch( _
//!     isUrgent And isImportant, "Critical", _
//!     isUrgent, "High", _
//!     isImportant, "Medium", _
//!     hasDeadline, "Low", _
//!     True, "Backlog")
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Grade Assignment
//! ```vb6
//! Function GetLetterGrade(score As Double) As String
//!     GetLetterGrade = Switch( _
//!         score >= 90, "A", _
//!         score >= 80, "B", _
//!         score >= 70, "C", _
//!         score >= 60, "D", _
//!         True, "F")
//! End Function
//! ```
//!
//! ### Pattern 2: Day of Week Name
//! ```vb6
//! Function GetDayName(dayNum As Integer) As String
//!     GetDayName = Switch( _
//!         dayNum = 1, "Sunday", _
//!         dayNum = 2, "Monday", _
//!         dayNum = 3, "Tuesday", _
//!         dayNum = 4, "Wednesday", _
//!         dayNum = 5, "Thursday", _
//!         dayNum = 6, "Friday", _
//!         dayNum = 7, "Saturday", _
//!         True, "Invalid")
//! End Function
//! ```
//!
//! ### Pattern 3: Conditional Formatting
//! ```vb6
//! Function GetColorCode(value As Double, threshold1 As Double, threshold2 As Double) As Long
//!     GetColorCode = Switch( _
//!         value < threshold1, vbRed, _
//!         value < threshold2, vbYellow, _
//!         True, vbGreen)
//! End Function
//! ```
//!
//! ### Pattern 4: Fee Calculation
//! ```vb6
//! Function CalculateShippingFee(weight As Double) As Currency
//!     CalculateShippingFee = Switch( _
//!         weight <= 1, 5.99, _
//!         weight <= 5, 9.99, _
//!         weight <= 10, 14.99, _
//!         weight <= 20, 24.99, _
//!         True, 39.99)
//! End Function
//! ```
//!
//! ### Pattern 5: Error Level Description
//! ```vb6
//! Function GetErrorDescription(errorLevel As Integer) As String
//!     GetErrorDescription = Switch( _
//!         errorLevel = 0, "No error", _
//!         errorLevel = 1, "Warning", _
//!         errorLevel = 2, "Error", _
//!         errorLevel = 3, "Critical error", _
//!         errorLevel = 4, "Fatal error", _
//!         True, "Unknown error level")
//! End Function
//! ```
//!
//! ### Pattern 6: Conditional Default Value
//! ```vb6
//! Function GetConfigValue(key As String, userValue As Variant, defaultValue As Variant) As Variant
//!     GetConfigValue = Switch( _
//!         Not IsNull(userValue), userValue, _
//!         True, defaultValue)
//! End Function
//! ```
//!
//! ### Pattern 7: Multi-Field Validation
//! ```vb6
//! Function ValidateRecord(name As String, age As Integer, email As String) As String
//!     ValidateRecord = Switch( _
//!         Len(name) = 0, "Name is required", _
//!         age < 0 Or age > 120, "Invalid age", _
//!         InStr(email, "@") = 0, "Invalid email", _
//!         True, "Valid")
//! End Function
//! ```
//!
//! ### Pattern 8: Price Tier Selection
//! ```vb6
//! Function GetPriceTier(quantity As Integer) As String
//!     GetPriceTier = Switch( _
//!         quantity >= 1000, "Enterprise", _
//!         quantity >= 100, "Business", _
//!         quantity >= 10, "Professional", _
//!         True, "Individual")
//! End Function
//! ```
//!
//! ### Pattern 9: File Type Detection
//! ```vb6
//! Function GetFileType(extension As String) As String
//!     Dim ext As String
//!     ext = LCase$(extension)
//!     
//!     GetFileType = Switch( _
//!         ext = "txt" Or ext = "log", "Text", _
//!         ext = "doc" Or ext = "docx", "Word", _
//!         ext = "xls" Or ext = "xlsx", "Excel", _
//!         ext = "jpg" Or ext = "png" Or ext = "gif", "Image", _
//!         True, "Unknown")
//! End Function
//! ```
//!
//! ### Pattern 10: Temperature Category
//! ```vb6
//! Function DescribeTemperature(tempF As Double) As String
//!     DescribeTemperature = Switch( _
//!         tempF < 32, "Freezing", _
//!         tempF < 50, "Cold", _
//!         tempF < 70, "Cool", _
//!         tempF < 85, "Comfortable", _
//!         tempF < 100, "Hot", _
//!         True, "Very Hot")
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Dynamic Selector Class
//! ```vb6
//! ' Class: DynamicSelector
//! ' Provides advanced selection logic using Switch
//! Option Explicit
//!
//! Private m_Criteria As Collection
//!
//! Public Sub Initialize()
//!     Set m_Criteria = New Collection
//! End Sub
//!
//! Public Sub AddCriteria(expression As Boolean, value As Variant)
//!     Dim pair As New Collection
//!     pair.Add expression
//!     pair.Add value
//!     m_Criteria.Add pair
//! End Sub
//!
//! Public Function Evaluate() As Variant
//!     Dim item As Variant
//!     Dim expr As Boolean
//!     Dim value As Variant
//!     
//!     For Each item In m_Criteria
//!         expr = item(1)
//!         If expr Then
//!             If IsObject(item(2)) Then
//!                 Set Evaluate = item(2)
//!             Else
//!                 Evaluate = item(2)
//!             End If
//!             Exit Function
//!         End If
//!     Next item
//!     
//!     Evaluate = Null
//! End Function
//!
//! Public Function EvaluateWithDefault(defaultValue As Variant) As Variant
//!     EvaluateWithDefault = Switch( _
//!         Not IsNull(Evaluate()), Evaluate(), _
//!         True, defaultValue)
//! End Function
//! ```
//!
//! ### Example 2: Conditional Formatter Module
//! ```vb6
//! ' Module: ConditionalFormatter
//! ' Formats values based on multiple conditions
//! Option Explicit
//!
//! Public Function FormatCurrency(amount As Currency, showSymbol As Boolean) As String
//!     Dim formatted As String
//!     
//!     formatted = Switch( _
//!         amount < 0, "(" & Format$(Abs(amount), "#,##0.00") & ")", _
//!         amount = 0, "0.00", _
//!         True, Format$(amount, "#,##0.00"))
//!     
//!     If showSymbol Then
//!         FormatCurrency = "$" & formatted
//!     Else
//!         FormatCurrency = formatted
//!     End If
//! End Function
//!
//! Public Function FormatPercentage(value As Double, decimals As Integer) As String
//!     Dim format As String
//!     
//!     format = Switch( _
//!         decimals = 0, "0%", _
//!         decimals = 1, "0.0%", _
//!         decimals = 2, "0.00%", _
//!         True, "0.00%")
//!     
//!     FormatPercentage = Format$(value * 100, format)
//! End Function
//!
//! Public Function FormatFileSize(bytes As Long) As String
//!     FormatFileSize = Switch( _
//!         bytes < 1024, bytes & " B", _
//!         bytes < 1048576, Format$(bytes / 1024, "0.0") & " KB", _
//!         bytes < 1073741824, Format$(bytes / 1048576, "0.0") & " MB", _
//!         True, Format$(bytes / 1073741824, "0.0") & " GB")
//! End Function
//!
//! Public Function GetStatusColor(status As String) As Long
//!     GetStatusColor = Switch( _
//!         status = "Active", vbGreen, _
//!         status = "Pending", vbYellow, _
//!         status = "Inactive", vbGray, _
//!         status = "Error", vbRed, _
//!         True, vbBlack)
//! End Function
//! ```
//!
//! ### Example 3: Business Rules Engine
//! ```vb6
//! ' Class: BusinessRulesEngine
//! ' Evaluates business rules using Switch
//! Option Explicit
//!
//! Public Function CalculateDiscount(customerType As String, orderAmount As Currency, _
//!                                   quantity As Integer) As Double
//!     ' Return discount percentage
//!     CalculateDiscount = Switch( _
//!         customerType = "VIP" And orderAmount > 1000, 0.2, _
//!         customerType = "VIP", 0.15, _
//!         orderAmount > 5000, 0.15, _
//!         orderAmount > 1000, 0.1, _
//!         quantity > 100, 0.1, _
//!         quantity > 50, 0.05, _
//!         True, 0)
//! End Function
//!
//! Public Function DetermineShippingMethod(weight As Double, destination As String, _
//!                                         isExpress As Boolean) As String
//!     DetermineShippingMethod = Switch( _
//!         isExpress And weight < 5, "Express Air", _
//!         isExpress, "Express Ground", _
//!         destination = "International", "International Standard", _
//!         weight > 50, "Freight", _
//!         weight > 10, "Ground", _
//!         True, "Standard Mail")
//! End Function
//!
//! Public Function CalculateTax(amount As Currency, state As String, _
//!                              category As String) As Currency
//!     Dim taxRate As Double
//!     
//!     taxRate = Switch( _
//!         state = "CA" And category = "Food", 0, _
//!         state = "CA", 0.0725, _
//!         state = "NY", 0.08, _
//!         state = "TX", 0.0625, _
//!         state = "FL", 0.06, _
//!         True, 0.05)
//!     
//!     CalculateTax = amount * taxRate
//! End Function
//!
//! Public Function ApprovalRequired(amount As Currency, department As String, _
//!                                  requestor As String) As Boolean
//!     ApprovalRequired = Switch( _
//!         amount > 10000, True, _
//!         amount > 5000 And department <> "Finance", True, _
//!         amount > 1000 And requestor = "Junior", True, _
//!         True, False)
//! End Function
//! ```
//!
//! ### Example 4: Data Categorizer Module
//! ```vb6
//! ' Module: DataCategorizer
//! ' Categorizes data based on complex rules
//! Option Explicit
//!
//! Public Function CategorizeCustomer(totalPurchases As Currency, _
//!                                    yearsMember As Integer, _
//!                                    lastPurchaseDays As Integer) As String
//!     CategorizeCustomer = Switch( _
//!         totalPurchases > 10000 And yearsMember > 5, "Platinum", _
//!         totalPurchases > 5000 And yearsMember > 3, "Gold", _
//!         totalPurchases > 1000 Or yearsMember > 2, "Silver", _
//!         lastPurchaseDays < 90, "Active", _
//!         lastPurchaseDays < 365, "Inactive", _
//!         True, "Dormant")
//! End Function
//!
//! Public Function AssignPriority(severity As Integer, impact As Integer, _
//!                                urgency As Integer) As String
//!     AssignPriority = Switch( _
//!         severity = 1 And impact = 1, "P1 - Critical", _
//!         severity <= 2 And impact <= 2 And urgency = 1, "P2 - High", _
//!         severity <= 3 Or impact <= 3, "P3 - Medium", _
//!         urgency <= 3, "P4 - Low", _
//!         True, "P5 - Backlog")
//! End Function
//!
//! Public Function DetermineRiskLevel(score As Integer, volatility As Double, _
//!                                    exposure As Currency) As String
//!     DetermineRiskLevel = Switch( _
//!         score < 300 Or exposure > 1000000, "High Risk", _
//!         score < 500 Or volatility > 0.3, "Medium Risk", _
//!         score < 700 And volatility > 0.15, "Low-Medium Risk", _
//!         True, "Low Risk")
//! End Function
//!
//! Public Function GetAgeGroup(age As Integer) As String
//!     GetAgeGroup = Switch( _
//!         age < 2, "Infant", _
//!         age < 13, "Child", _
//!         age < 20, "Teenager", _
//!         age < 40, "Young Adult", _
//!         age < 65, "Adult", _
//!         True, "Senior")
//! End Function
//! ```
//!
//! ## Error Handling
//! The `Switch` function can raise the following errors:
//!
//! - **Error 450 (Wrong number of arguments or invalid property assignment)**: If arguments are not provided in pairs (odd number of arguments)
//! - **Error 13 (Type mismatch)**: If expressions cannot be evaluated as Boolean
//! - **Error 94 (Invalid use of Null)**: In some contexts when returned Null is used without checking
//!
//! ## Performance Notes
//! - Evaluates expressions left to right until True is found (short-circuit)
//! - More efficient than nested `IIf` for multiple conditions
//! - All value expressions may be evaluated regardless of condition (avoid side effects)
//! - For large numbers of conditions (10+), Select Case may be more readable
//! - Performance: O(n) where n is the number of condition/value pairs
//!
//! ## Best Practices
//! 1. **Always provide default case** using `True` as the last condition to avoid Null returns
//! 2. **Order conditions properly** - most specific first, most general last
//! 3. **Avoid side effects** in value expressions (they may all be evaluated)
//! 4. **Use for readability** - Switch is clearer than nested `IIf` for 3+ conditions
//! 5. **Check for Null** when using returned value if no default case provided
//! 6. **Keep pairs aligned** for readability (use line continuation)
//! 7. **Limit complexity** - if more than 7-8 pairs, consider Select Case
//! 8. **Document complex logic** when conditions are not self-evident
//! 9. **Test all branches** to ensure correct behavior
//! 10. **Consider Select Case** for executing statements vs returning values
//!
//! ## Comparison Table
//!
//! | Construct | Type | Branches | Returns Value | Short-Circuit |
//! |-----------|------|----------|---------------|---------------|
//! | `Switch` | Function | Multiple | Yes | Yes (conditions) |
//! | `IIf` | Function | 2 | Yes | No |
//! | `Select Case` | Statement | Multiple | No (use with function) | Yes |
//! | `If...Then...Else` | Statement | Multiple | No (use with function) | Yes |
//!
//! ## Platform Notes
//! - Available in VB6, VBA, and `VBScript`
//! - Behavior consistent across platforms
//! - Return type is always Variant
//! - Arguments must come in pairs (expression, value)
//! - Null return when no condition is True (not an error)
//!
//! ## Limitations
//! - Must have even number of arguments (pairs)
//! - Cannot execute statements (only return values)
//! - Returns Variant (must convert if specific type needed)
//! - All value expressions may be evaluated (cannot assume short-circuit for values)
//! - Less readable than Select Case for complex branching
//! - No fall-through behavior like some languages' switch statements
//! - Cannot use ranges directly (must use comparison expressions)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_switch_basic() {
        let source = r#"
Sub Test()
    result = Switch(x = 1, "One", x = 2, "Two", True, "Other")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_variable_assignment() {
        let source = r#"
Sub Test()
    Dim grade As String
    grade = Switch(score >= 90, "A", score >= 80, "B", True, "F")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
        assert!(debug.contains("score"));
    }

    #[test]
    fn test_switch_grade_calculation() {
        let source = r#"
Sub Test()
    grade = Switch(score >= 90, "A", score >= 80, "B", score >= 70, "C", True, "F")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_function_return() {
        let source = r#"
Function GetStatus(code As Integer) As String
    GetStatus = Switch(code = 0, "OK", code = 1, "Warning", True, "Error")
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_if_statement() {
        let source = r#"
Sub Test()
    If Switch(status = 1, True, status = 2, True, True, False) Then
        ProcessItem
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_msgbox() {
        let source = r#"
Sub Test()
    MsgBox Switch(day = 1, "Monday", day = 2, "Tuesday", True, "Other")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_select_case() {
        let source = r#"
Sub Test()
    Select Case Switch(type = 1, "A", type = 2, "B", True, "C")
        Case "A"
            DoSomething
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_for_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        value = Switch(i < 5, "Low", i < 8, "Mid", True, "High")
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_array_assignment() {
        let source = r#"
Sub Test()
    categories(i) = Switch(values(i) < 10, "Small", values(i) < 100, "Medium", True, "Large")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_function_argument() {
        let source = r#"
Sub Test()
    Call DisplayMessage(Switch(level = 1, "Info", level = 2, "Warning", True, "Error"))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_comparison() {
        let source = r#"
Sub Test()
    If Switch(x > 10, True, y > 10, True, True, False) Then
        ProcessData
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print Switch(flag, "Enabled", True, "Disabled")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_do_while() {
        let source = r#"
Sub Test()
    Do While Switch(counter < 10, True, True, False)
        counter = counter + 1
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_do_until() {
        let source = r#"
Sub Test()
    Do Until Switch(status = "Done", True, True, False)
        status = Process()
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_while_wend() {
        let source = r#"
Sub Test()
    While Switch(i < max, True, True, False)
        i = i + 1
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_iif() {
        let source = r#"
Sub Test()
    result = IIf(flag, Switch(x = 1, "A", True, "B"), "Default")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_with_statement() {
        let source = r#"
Sub Test()
    With obj
        .Status = Switch(.Value > 100, "High", .Value > 50, "Medium", True, "Low")
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_parentheses() {
        let source = r#"
Sub Test()
    result = (Switch(a > b, a, True, b))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    category = Switch(val < 0, "Negative", val = 0, "Zero", True, "Positive")
    If Err.Number <> 0 Then
        category = "Error"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_property_assignment() {
        let source = r#"
Sub Test()
    obj.Priority = Switch(urgent, "High", important, "Medium", True, "Low")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_concatenation() {
        let source = r#"
Sub Test()
    message = "Status: " & Switch(code = 0, "OK", code = 1, "Error", True, "Unknown")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_numeric_result() {
        let source = r#"
Sub Test()
    discount = Switch(qty >= 100, 0.2, qty >= 50, 0.1, qty >= 10, 0.05, True, 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_print_statement() {
        let source = r#"
Sub Test()
    Print #1, Switch(type = 1, "Type A", type = 2, "Type B", True, "Unknown")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_class_usage() {
        let source = r#"
Sub Test()
    Set obj = New DataProcessor
    obj.Category = Switch(value > 1000, "A", value > 100, "B", True, "C")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_nested() {
        let source = r#"
Sub Test()
    result = Switch(x > 0, Switch(x > 100, "Large", True, "Small"), True, "Negative")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_elseif() {
        let source = r#"
Sub Test()
    If x = 1 Then
        y = "One"
    ElseIf x = 2 Then
        y = Switch(z > 0, "Two-Positive", True, "Two-Negative")
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }

    #[test]
    fn test_switch_complex_conditions() {
        let source = r#"
Sub Test()
    priority = Switch(urgent And important, "Critical", urgent, "High", important, "Medium", True, "Low")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Switch"));
    }
}
