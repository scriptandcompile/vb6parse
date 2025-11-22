//! # Choose Function
//!
//! Returns a value from a list of choices based on an index.
//!
//! ## Syntax
//!
//! ```vb
//! Choose(index, choice-1[, choice-2, ... [, choice-n]])
//! ```
//!
//! ## Parameters
//!
//! - **index**: Required. Numeric expression (typically Integer or Long) that results in a value 
//!   between 1 and the number of available choices. The index is 1-based.
//!
//! - **choice**: Required. Variant expression containing one of the possible choices. You must 
//!   provide at least one choice argument.
//!
//! ## Return Value
//!
//! Returns a Variant containing the value of the choice at the specified index position. If the 
//! index is less than 1 or greater than the number of choices, `Choose` returns `Null`.
//!
//! ## Remarks
//!
//! The `Choose` function provides a convenient way to select from a list of values based on an 
//! index, similar to a switch/case statement but as an expression. It's particularly useful for:
//!
//! - Mapping numeric codes to string values
//! - Selecting values based on calculated indices
//! - Simplifying multi-way conditional expressions
//! - Implementing lookup tables inline
//!
//! **Important Characteristics:**
//!
//! - The index is 1-based (first choice is at index 1)
//! - Returns `Null` if index is out of range (< 1 or > number of choices)
//! - All choice arguments are evaluated before selection occurs (unlike `IIf`)
//! - Can accept any Variant-compatible expressions as choices
//! - Choices can be of mixed types (numbers, strings, objects, etc.)
//! - The returned value's type depends on the selected choice
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Simple value selection
//! Dim dayType As String
//! dayType = Choose(Weekday(Date), "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
//!
//! ' Numeric selection
//! Dim priority As Integer
//! priority = Choose(level, 1, 5, 10, 50, 100)
//!
//! ' Mixed types
//! Dim result As Variant
//! result = Choose(2, 100, "Text", #1/1/2000#, True)  ' Returns "Text"
//! ```
//!
//! ### Mapping Codes to Descriptions
//!
//! ```vb
//! Function GetStatusDescription(statusCode As Integer) As String
//!     GetStatusDescription = Choose(statusCode, _
//!         "Pending", _
//!         "Approved", _
//!         "Rejected", _
//!         "On Hold", _
//!         "Completed")
//!     
//!     If IsNull(GetStatusDescription) Then
//!         GetStatusDescription = "Unknown"
//!     End If
//! End Function
//! ```
//!
//! ### Dynamic Message Selection
//!
//! ```vb
//! Sub ShowErrorMessage(errorLevel As Integer)
//!     Dim msg As String
//!     msg = Choose(errorLevel, _
//!         "Operation completed successfully.", _
//!         "Warning: Please review the results.", _
//!         "Error: Operation failed.", _
//!         "Critical: System error occurred.")
//!     
//!     MsgBox msg, vbInformation
//! End Sub
//! ```
//!
//! ## Common Patterns
//!
//! ### Month Name Lookup
//!
//! ```vb
//! Function GetMonthName(monthNum As Integer) As String
//!     GetMonthName = Choose(monthNum, _
//!         "January", "February", "March", "April", _
//!         "May", "June", "July", "August", _
//!         "September", "October", "November", "December")
//! End Function
//! ```
//!
//! ### Grade Calculation
//!
//! ```vb
//! Function GetLetterGrade(score As Integer) As String
//!     Dim gradeIndex As Integer
//!     gradeIndex = Int(score / 10) - 5  ' 60-69=1, 70-79=2, etc.
//!     If gradeIndex < 1 Then gradeIndex = 1
//!     If gradeIndex > 5 Then gradeIndex = 5
//!     
//!     GetLetterGrade = Choose(gradeIndex, "F", "D", "C", "B", "A")
//! End Function
//! ```
//!
//! ### Configuration Selection
//!
//! ```vb
//! Function GetServerURL(environment As Integer) As String
//!     ' 1=Development, 2=Testing, 3=Staging, 4=Production
//!     GetServerURL = Choose(environment, _
//!         "http://localhost:8080", _
//!         "http://test.example.com", _
//!         "http://staging.example.com", _
//!         "https://www.example.com")
//! End Function
//! ```
//!
//! ### Color Code Mapping
//!
//! ```vb
//! Function GetColorValue(colorCode As Integer) As Long
//!     GetColorValue = Choose(colorCode, _
//!         vbRed, vbGreen, vbBlue, vbYellow, _
//!         vbMagenta, vbCyan, vbWhite, vbBlack)
//! End Function
//! ```
//!
//! ### Day of Week Operations
//!
//! ```vb
//! Function IsBusinessDay(dayOfWeek As Integer) As Boolean
//!     ' Sunday=1, Monday=2, ... Saturday=7
//!     IsBusinessDay = Choose(dayOfWeek, _
//!         False, True, True, True, True, True, False)
//! End Function
//! ```
//!
//! ### File Extension Mapping
//!
//! ```vb
//! Function GetFileTypeDescription(fileType As Integer) As String
//!     GetFileTypeDescription = Choose(fileType, _
//!         "Text Document", _
//!         "Spreadsheet", _
//!         "Database", _
//!         "Image File", _
//!         "Executable Program")
//! End Function
//! ```
//!
//! ### Priority Level Mapping
//!
//! ```vb
//! Function GetPriorityWeight(priority As Integer) As Double
//!     GetPriorityWeight = Choose(priority, 0.25, 0.5, 1.0, 2.0, 5.0)
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### With Null Handling
//!
//! ```vb
//! Function SafeChoose(index As Integer, ParamArray choices()) As Variant
//!     Dim result As Variant
//!     Dim i As Integer
//!     Dim args As String
//!     
//!     ' Build argument list
//!     For i = LBound(choices) To UBound(choices)
//!         If i > LBound(choices) Then args = args & ", "
//!         args = args & """" & choices(i) & """"
//!     Next i
//!     
//!     ' Use Execute to call Choose dynamically
//!     result = Choose(index, choices(0), choices(1), choices(2))  ' etc.
//!     
//!     If IsNull(result) Then
//!         SafeChoose = "Invalid Index"
//!     Else
//!         SafeChoose = result
//!     End If
//! End Function
//! ```
//!
//! ### Combined with Calculation
//!
//! ```vb
//! Function CalculateDiscount(customerType As Integer, amount As Currency) As Currency
//!     Dim discountRate As Double
//!     discountRate = Choose(customerType, 0.05, 0.1, 0.15, 0.2)
//!     
//!     If IsNull(discountRate) Then
//!         discountRate = 0
//!     End If
//!     
//!     CalculateDiscount = amount * discountRate
//! End Function
//! ```
//!
//! ### Nested Choose
//!
//! ```vb
//! Function GetRegionalGreeting(region As Integer, timeOfDay As Integer) As String
//!     ' region: 1=North, 2=South, 3=East, 4=West
//!     ' timeOfDay: 1=Morning, 2=Afternoon, 3=Evening
//!     
//!     GetRegionalGreeting = Choose(region, _
//!         Choose(timeOfDay, "Good morning, y'all", "Good afternoon, y'all", "Good evening, y'all"), _
//!         Choose(timeOfDay, "Mornin'", "Afternoon", "Evenin'"), _
//!         Choose(timeOfDay, "Good morning", "Good afternoon", "Good evening"), _
//!         Choose(timeOfDay, "Hey, morning!", "Hey, afternoon!", "Hey, evening!"))
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeChooseWithError(index As Integer) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     Dim result As Variant
//!     result = Choose(index, "First", "Second", "Third")
//!     
//!     If IsNull(result) Then
//!         MsgBox "Index out of range: " & index
//!         SafeChooseWithError = Empty
//!     Else
//!         SafeChooseWithError = result
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     MsgBox "Error in Choose: " & Err.Description
//!     SafeChooseWithError = Empty
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 450** (Wrong number of arguments): Occurs if you don't provide at least one choice
//! - **Error 13** (Type mismatch): Can occur if index expression is not numeric
//! - **Null return**: Index is < 1 or > number of choices (not an error, but important to handle)
//!
//! ## Performance Considerations
//!
//! - All choice arguments are evaluated before selection, even if not used
//! - For expensive computations, consider using `Select Case` instead
//! - Choose is most efficient with literal values or simple expressions
//! - For large lookup tables, consider using arrays or collections
//!
//! ## Comparison with Alternatives
//!
//! ### Choose vs. Select Case
//!
//! ```vb
//! ' Using Choose (expression-based)
//! result = Choose(index, "A", "B", "C")
//!
//! ' Using Select Case (statement-based)
//! Select Case index
//!     Case 1: result = "A"
//!     Case 2: result = "B"
//!     Case 3: result = "C"
//!     Case Else: result = Null
//! End Select
//! ```
//!
//! **Choose advantages:**
//! - More concise for simple selections
//! - Can be used as an expression
//! - Good for inline value lookup
//!
//! **Select Case advantages:**
//! - Doesn't evaluate all branches
//! - Better for complex logic
//! - More readable for many cases
//! - Supports range matching
//!
//! ### Choose vs. Array Lookup
//!
//! ```vb
//! ' Using Choose
//! result = Choose(index, "A", "B", "C")
//!
//! ' Using Array
//! Dim choices() As Variant
//! choices = Array("A", "B", "C")
//! If index >= 1 And index <= UBound(choices) + 1 Then
//!     result = choices(index - 1)  ' Array is 0-based, Choose is 1-based
//! Else
//!     result = Null
//! End If
//! ```
//!
//! ## Limitations
//!
//! - All arguments are evaluated even if not selected (performance impact)
//! - Limited to 29 arguments in VB6 (index + 28 choices max in practice)
//! - Index must be numeric (Integer, Long, Byte, etc.)
//! - Returns Null for out-of-range indices (requires explicit checking)
//! - Cannot use named arguments
//! - Not suitable for ranges or complex matching logic
//!
//! ## Related Functions
//!
//! - `Switch`: Similar but uses condition/value pairs instead of index-based selection
//! - `IIf`: Binary choice based on condition (evaluates both branches)
//! - `Select Case`: Statement-based multi-way branching
//! - `Array`: Creates an array that can be indexed (0-based)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_choose_basic() {
        let source = r#"
result = Choose(1, "First", "Second", "Third")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_with_variable_index() {
        let source = r#"
result = Choose(index, "A", "B", "C", "D")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_numeric_values() {
        let source = r#"
value = Choose(level, 1, 5, 10, 50, 100)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_with_expression_index() {
        let source = r#"
result = Choose(x + 1, "Zero", "One", "Two")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_weekday() {
        let source = r#"
dayName = Choose(Weekday(Date), "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_with_line_continuation() {
        let source = r#"
result = Choose(index, _
    "First", _
    "Second", _
    "Third")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_month_names() {
        let source = r#"
monthName = Choose(month, "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_in_if_statement() {
        let source = r#"
If Choose(status, "Pending", "Active", "Closed") = "Active" Then
    Process
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_in_function_call() {
        let source = r#"
MsgBox Choose(errorLevel, "Info", "Warning", "Error")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_with_null_check() {
        let source = r#"
result = Choose(index, "A", "B", "C")
If IsNull(result) Then
    result = "Invalid"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_color_values() {
        let source = r#"
color = Choose(colorCode, vbRed, vbGreen, vbBlue, vbYellow)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_boolean_values() {
        let source = r#"
isValid = Choose(dayOfWeek, False, True, True, True, True, True, False)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_with_function_return() {
        let source = r#"
Function GetStatus(code As Integer) As String
    GetStatus = Choose(code, "Pending", "Approved", "Rejected")
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_nested() {
        let source = r#"
result = Choose(x, Choose(y, "A", "B"), Choose(y, "C", "D"))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_with_calculations() {
        let source = r#"
discount = amount * Choose(customerType, 0.05, 0.1, 0.15, 0.2)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_in_select_case() {
        let source = r#"
Select Case Choose(index, 1, 2, 3)
    Case 1
        DoSomething
    Case 2
        DoOther
End Select
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_with_dates() {
        let source = r#"
dueDate = Choose(priority, #1/1/2000#, #1/15/2000#, #2/1/2000#)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_in_loop() {
        let source = r#"
For i = 1 To 3
    Print Choose(i, "First", "Second", "Third")
Next i
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_mixed_types() {
        let source = r#"
result = Choose(selector, 100, "Text", #1/1/2000#, True, 3.14)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_with_object_properties() {
        let source = r#"
value = Choose(index, obj.Property1, obj.Property2, obj.Property3)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_with_array_elements() {
        let source = r#"
result = Choose(index, arr(0), arr(1), arr(2))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_in_assignment() {
        let source = r#"
Dim msg As String
msg = Choose(errorLevel, "Success", "Warning", "Error", "Critical")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_with_method_calls() {
        let source = r#"
result = Choose(index, obj.Method1(), obj.Method2(), obj.Method3())
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_in_do_loop() {
        let source = r#"
Do While counter < 10
    status = Choose(counter, "A", "B", "C")
    counter = counter + 1
Loop
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_choose_with_whitespace() {
        let source = r#"
result = Choose( index , "First" , "Second" , "Third" )
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Choose"));
        assert!(debug.contains("Identifier"));
    }
}
