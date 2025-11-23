//! # Int Function
//!
//! Returns the integer portion of a number.
//!
//! ## Syntax
//!
//! ```vb
//! Int(number)
//! ```
//!
//! ## Parameters
//!
//! - `number` (Required): Any valid numeric expression. If number contains Null, Null is returned
//!
//! ## Return Value
//!
//! Returns the integer portion of a number:
//! - For positive numbers: Returns the largest integer less than or equal to number
//! - For negative numbers: Returns the first negative integer less than or equal to number
//! - If number is Null: Returns Null
//! - Return type matches the input type (Integer, Long, Single, Double, Currency, Decimal)
//!
//! ## Remarks
//!
//! The Int function truncates toward negative infinity:
//!
//! - Removes the fractional part of a number
//! - For positive numbers, behaves like truncation (same as Fix)
//! - For negative numbers, rounds DOWN (toward negative infinity)
//! - Fix rounds toward zero (always truncates), Int rounds down
//! - Int(-8.4) returns -9, Fix(-8.4) returns -8
//! - Int(8.4) returns 8, Fix(8.4) returns 8
//! - Does not round to nearest integer (use Round for rounding)
//! - The return type preserves the input numeric type
//! - Commonly used with Rnd for generating random integers
//! - For currency calculations, consider using Round or CCur instead
//!
//! ## Typical Uses
//!
//! 1. **Remove Decimals**: Strip fractional part from numbers
//! 2. **Random Integers**: Generate random integer values with Rnd
//! 3. **Array Indices**: Convert floats to valid array indices
//! 4. **Loop Counters**: Ensure integer values for loops
//! 5. **Division Results**: Get whole number quotients
//! 6. **Coordinate Rounding**: Round pixel coordinates
//! 7. **Pagination**: Calculate page numbers
//! 8. **Quantity Calculations**: Ensure whole unit quantities
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Remove decimal portion
//! Dim result As Integer
//! result = Int(8.7)
//! Debug.Print result  ' Prints: 8
//!
//! ' Example 2: Negative number behavior
//! Dim result As Integer
//! result = Int(-8.7)
//! Debug.Print result  ' Prints: -9 (rounds down, not toward zero)
//!
//! ' Example 3: Random integer between 1 and 100
//! Dim randomNum As Integer
//! Randomize
//! randomNum = Int(Rnd * 100) + 1
//!
//! ' Example 4: Calculate whole pages
//! Dim totalItems As Long
//! Dim itemsPerPage As Long
//! Dim totalPages As Long
//! totalItems = 47
//! itemsPerPage = 10
//! totalPages = Int(totalItems / itemsPerPage) + 1
//! Debug.Print totalPages  ' Prints: 5
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Random integer in range
//! Function RandomInteger(minValue As Long, maxValue As Long) As Long
//!     Randomize
//!     RandomInteger = Int((maxValue - minValue + 1) * Rnd) + minValue
//! End Function
//!
//! ' Pattern 2: Get whole number portion
//! Function GetWholeNumber(value As Double) As Long
//!     If value >= 0 Then
//!         GetWholeNumber = Int(value)
//!     Else
//!         ' For negative numbers, Int rounds down
//!         ' Use Fix if you want to truncate toward zero
//!         GetWholeNumber = Int(value)
//!     End If
//! End Function
//!
//! ' Pattern 3: Calculate pages needed
//! Function CalculatePages(totalItems As Long, itemsPerPage As Long) As Long
//!     If itemsPerPage <= 0 Then
//!         CalculatePages = 0
//!         Exit Function
//!     End If
//!     
//!     CalculatePages = Int((totalItems - 1) / itemsPerPage) + 1
//! End Function
//!
//! ' Pattern 4: Round down to nearest multiple
//! Function RoundDownToMultiple(value As Double, multiple As Double) As Double
//!     If multiple = 0 Then
//!         RoundDownToMultiple = value
//!     Else
//!         RoundDownToMultiple = Int(value / multiple) * multiple
//!     End If
//! End Function
//!
//! ' Pattern 5: Extract integer part for display
//! Function FormatNumber(value As Double) As String
//!     Dim wholePart As Long
//!     Dim decimalPart As Double
//!     
//!     wholePart = Int(Abs(value))
//!     decimalPart = Abs(value) - wholePart
//!     
//!     FormatNumber = CStr(wholePart) & "." & _
//!                    Format$(decimalPart, "00")
//! End Function
//!
//! ' Pattern 6: Generate random array index
//! Function RandomArrayIndex(arr As Variant) As Long
//!     Dim lowerBound As Long
//!     Dim upperBound As Long
//!     
//!     lowerBound = LBound(arr)
//!     upperBound = UBound(arr)
//!     
//!     RandomArrayIndex = Int((upperBound - lowerBound + 1) * Rnd) + lowerBound
//! End Function
//!
//! ' Pattern 7: Calculate grid position
//! Sub GetGridPosition(pixelX As Double, pixelY As Double, _
//!                     gridSize As Double, _
//!                     ByRef gridX As Long, ByRef gridY As Long)
//!     gridX = Int(pixelX / gridSize)
//!     gridY = Int(pixelY / gridSize)
//! End Sub
//!
//! ' Pattern 8: Divide and get quotient
//! Function IntegerDivision(dividend As Long, divisor As Long) As Long
//!     If divisor = 0 Then
//!         Err.Raise 11, , "Division by zero"
//!     End If
//!     
//!     IntegerDivision = Int(dividend / divisor)
//! End Function
//!
//! ' Pattern 9: Time to whole seconds
//! Function GetWholeSeconds(timeValue As Double) As Long
//!     Dim secondsDecimal As Double
//!     secondsDecimal = timeValue * 86400  ' Convert days to seconds
//!     GetWholeSeconds = Int(secondsDecimal)
//! End Function
//!
//! ' Pattern 10: Percentage to whole number
//! Function GetWholePercent(value As Double, total As Double) As Long
//!     If total = 0 Then
//!         GetWholePercent = 0
//!     Else
//!         GetWholePercent = Int((value / total) * 100)
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Random number generator class
//! Public Class RandomNumberGenerator
//!     Private m_initialized As Boolean
//!     
//!     Private Sub EnsureInitialized()
//!         If Not m_initialized Then
//!             Randomize
//!             m_initialized = True
//!         End If
//!     End Sub
//!     
//!     Public Function NextInteger(minValue As Long, maxValue As Long) As Long
//!         EnsureInitialized
//!         
//!         If minValue > maxValue Then
//!             Err.Raise 5, , "minValue cannot be greater than maxValue"
//!         End If
//!         
//!         NextInteger = Int((maxValue - minValue + 1) * Rnd) + minValue
//!     End Function
//!     
//!     Public Function NextDouble() As Double
//!         EnsureInitialized
//!         NextDouble = Rnd
//!     End Function
//!     
//!     Public Function NextBoolean() As Boolean
//!         EnsureInitialized
//!         NextBoolean = (Int(Rnd * 2) = 1)
//!     End Function
//!     
//!     Public Function Shuffle(arr As Variant) As Variant
//!         Dim i As Long
//!         Dim j As Long
//!         Dim temp As Variant
//!         Dim result() As Variant
//!         
//!         EnsureInitialized
//!         
//!         ' Copy array
//!         ReDim result(LBound(arr) To UBound(arr))
//!         For i = LBound(arr) To UBound(arr)
//!             result(i) = arr(i)
//!         Next i
//!         
//!         ' Fisher-Yates shuffle
//!         For i = UBound(result) To LBound(result) + 1 Step -1
//!             j = Int((i - LBound(result) + 1) * Rnd) + LBound(result)
//!             temp = result(i)
//!             result(i) = result(j)
//!             result(j) = temp
//!         Next i
//!         
//!         Shuffle = result
//!     End Function
//! End Class
//!
//! ' Example 2: Pagination calculator
//! Public Class PaginationHelper
//!     Private m_totalItems As Long
//!     Private m_itemsPerPage As Long
//!     
//!     Public Property Let TotalItems(value As Long)
//!         m_totalItems = value
//!     End Property
//!     
//!     Public Property Let ItemsPerPage(value As Long)
//!         If value <= 0 Then
//!             Err.Raise 5, , "ItemsPerPage must be greater than zero"
//!         End If
//!         m_itemsPerPage = value
//!     End Property
//!     
//!     Public Property Get PageCount() As Long
//!         If m_itemsPerPage = 0 Then
//!             PageCount = 0
//!         Else
//!             PageCount = Int((m_totalItems - 1) / m_itemsPerPage) + 1
//!         End If
//!     End Property
//!     
//!     Public Function GetPageStartIndex(pageNumber As Long) As Long
//!         If pageNumber < 1 Or pageNumber > PageCount Then
//!             GetPageStartIndex = -1
//!         Else
//!             GetPageStartIndex = (pageNumber - 1) * m_itemsPerPage
//!         End If
//!     End Function
//!     
//!     Public Function GetPageEndIndex(pageNumber As Long) As Long
//!         Dim startIndex As Long
//!         startIndex = GetPageStartIndex(pageNumber)
//!         
//!         If startIndex = -1 Then
//!             GetPageEndIndex = -1
//!         Else
//!             GetPageEndIndex = startIndex + m_itemsPerPage - 1
//!             If GetPageEndIndex >= m_totalItems Then
//!                 GetPageEndIndex = m_totalItems - 1
//!             End If
//!         End If
//!     End Function
//!     
//!     Public Function GetPageForItem(itemIndex As Long) As Long
//!         If itemIndex < 0 Or itemIndex >= m_totalItems Then
//!             GetPageForItem = -1
//!         Else
//!             GetPageForItem = Int(itemIndex / m_itemsPerPage) + 1
//!         End If
//!     End Function
//! End Class
//!
//! ' Example 3: Grid coordinate mapper
//! Public Class GridMapper
//!     Private m_cellWidth As Double
//!     Private m_cellHeight As Double
//!     
//!     Public Sub Initialize(cellWidth As Double, cellHeight As Double)
//!         m_cellWidth = cellWidth
//!         m_cellHeight = cellHeight
//!     End Sub
//!     
//!     Public Sub PixelToGrid(pixelX As Double, pixelY As Double, _
//!                           ByRef gridX As Long, ByRef gridY As Long)
//!         gridX = Int(pixelX / m_cellWidth)
//!         gridY = Int(pixelY / m_cellHeight)
//!     End Sub
//!     
//!     Public Sub GridToPixel(gridX As Long, gridY As Long, _
//!                           ByRef pixelX As Double, ByRef pixelY As Double)
//!         pixelX = gridX * m_cellWidth
//!         pixelY = gridY * m_cellHeight
//!     End Sub
//!     
//!     Public Function SnapToGrid(pixelX As Double, pixelY As Double) As Variant
//!         Dim gridX As Long
//!         Dim gridY As Long
//!         Dim snappedX As Double
//!         Dim snappedY As Double
//!         
//!         PixelToGrid pixelX, pixelY, gridX, gridY
//!         GridToPixel gridX, gridY, snappedX, snappedY
//!         
//!         SnapToGrid = Array(snappedX, snappedY)
//!     End Function
//! End Class
//!
//! ' Example 4: Dice roller simulator
//! Public Class DiceRoller
//!     Public Function Roll(sides As Long, Optional count As Long = 1) As Long
//!         Dim i As Long
//!         Dim total As Long
//!         
//!         Randomize
//!         total = 0
//!         
//!         For i = 1 To count
//!             total = total + Int(Rnd * sides) + 1
//!         Next i
//!         
//!         Roll = total
//!     End Function
//!     
//!     Public Function RollMultiple(sides As Long, count As Long) As Collection
//!         Dim i As Long
//!         Dim result As New Collection
//!         
//!         Randomize
//!         
//!         For i = 1 To count
//!             result.Add Int(Rnd * sides) + 1
//!         Next i
//!         
//!         Set RollMultiple = result
//!     End Function
//!     
//!     Public Function RollWithAdvantage(sides As Long) As Long
//!         Dim roll1 As Long
//!         Dim roll2 As Long
//!         
//!         Randomize
//!         roll1 = Int(Rnd * sides) + 1
//!         roll2 = Int(Rnd * sides) + 1
//!         
//!         RollWithAdvantage = IIf(roll1 > roll2, roll1, roll2)
//!     End Function
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! The Int function can raise errors or return Null:
//!
//! - **Type Mismatch (Error 13)**: If number is not a numeric expression
//! - **Invalid use of Null (Error 94)**: If number is Null and result is assigned to non-Variant
//! - **Overflow (Error 6)**: If result exceeds the range of the target data type
//!
//! ```vb
//! On Error GoTo ErrorHandler
//! Dim result As Long
//! Dim value As Double
//!
//! value = 12.75
//! result = Int(value)
//!
//! Debug.Print "Integer portion: " & result
//! Exit Sub
//!
//! ErrorHandler:
//!     MsgBox "Error in Int: " & Err.Description, vbCritical
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: Int is a very fast built-in function
//! - **Type Preservation**: Return type matches input type
//! - **No Rounding**: Faster than Round (no complex calculation)
//! - **Alternative**: For truncation toward zero, Fix is equivalent for positive numbers
//! - **Currency**: For financial calculations, consider Round or CCur
//!
//! ## Best Practices
//!
//! 1. **Understand Behavior**: Know that Int rounds DOWN (toward negative infinity)
//! 2. **Fix vs Int**: Use Fix for truncation toward zero, Int for floor operation
//! 3. **Random Numbers**: Always Randomize before using Rnd with Int
//! 4. **Type Awareness**: Be aware of return type matching input type
//! 5. **Null Handling**: Use Variant if input might be Null
//! 6. **Array Bounds**: Ensure Int result is within array bounds
//! 7. **Division**: For integer division, consider using \ operator instead
//!
//! ## Comparison with Other Functions
//!
//! | Function | Behavior | Example |
//! |----------|----------|---------|
//! | Int | Rounds down (floor) | Int(-8.7) = -9 |
//! | Fix | Truncates toward zero | Fix(-8.7) = -8 |
//! | Round | Rounds to nearest | Round(-8.7) = -9 |
//! | CLng | Converts to Long with rounding | CLng(-8.7) = -9 |
//! | CInt | Converts to Integer with rounding | CInt(-8.7) = -9 |
//! | \ | Integer division | -87 \ 10 = -8 |
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Consistent behavior across platforms
//! - Return type matches input numeric type
//! - Different from many languages' int() which truncates toward zero
//! - Equivalent to Math.floor() in many other languages
//!
//! ## Limitations
//!
//! - Does not round to nearest (use Round for that)
//! - Behavior with negative numbers can be unexpected (use Fix for truncation)
//! - Return type depends on input type (can cause overflow)
//! - Cannot specify decimal places (always removes all decimals)
//! - No control over rounding direction (always down)
//!
//! ## Related Functions
//!
//! - `Fix`: Returns integer portion, truncating toward zero
//! - `Round`: Rounds to nearest integer or specified decimal places
//! - `CInt`: Converts to Integer with rounding
//! - `CLng`: Converts to Long with rounding
//! - `Rnd`: Random number generator (often used with Int)
//! - `\`: Integer division operator

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_int_basic() {
        let source = r#"
Sub Test()
    result = Int(8.7)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_negative() {
        let source = r#"
Sub Test()
    result = Int(-8.7)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_random() {
        let source = r#"
Sub Test()
    randomNum = Int(Rnd * 100) + 1
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_if_statement() {
        let source = r#"
Sub Test()
    If Int(value) > 10 Then
        Debug.Print "Greater than 10"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_function_return() {
        let source = r#"
Function GetWhole(value As Double) As Long
    GetWhole = Int(value)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_division() {
        let source = r#"
Sub Test()
    pages = Int(totalItems / itemsPerPage) + 1
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_comparison() {
        let source = r#"
Sub Test()
    If Int(price) = expectedPrice Then
        MsgBox "Price matches"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_select_case() {
        let source = r#"
Sub Test()
    Select Case Int(score / 10)
        Case 10, 9
            grade = "A"
        Case 8
            grade = "B"
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_for_loop() {
        let source = r#"
Sub Test()
    Dim i As Long
    For i = 1 To Int(maxValue)
        Debug.Print i
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Integer part: " & Int(value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_array_assignment() {
        let source = r#"
Sub Test()
    wholeNumbers(i) = Int(decimals(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_property_assignment() {
        let source = r#"
Sub Test()
    obj.WholeValue = Int(obj.DecimalValue)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_in_class() {
        let source = r#"
Private Sub Class_Initialize()
    m_wholePart = Int(m_value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_with_statement() {
        let source = r#"
Sub Test()
    With calculator
        .WholeResult = Int(.DecimalResult)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_function_argument() {
        let source = r#"
Sub Test()
    Call ProcessInteger(Int(value))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_concatenation() {
        let source = r#"
Sub Test()
    message = "Whole part: " & Int(number)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_math_expression() {
        let source = r#"
Sub Test()
    gridX = Int(pixelX / gridSize)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_iif() {
        let source = r#"
Sub Test()
    result = IIf(Int(value) > 0, "Positive", "Zero or Negative")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Integer value: " & Int(value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_collection_add() {
        let source = r#"
Sub Test()
    numbers.Add Int(values(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_boolean_expression() {
        let source = r#"
Sub Test()
    isValid = Int(value) >= minValue And Int(value) <= maxValue
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_nested_call() {
        let source = r#"
Sub Test()
    result = CStr(Int(value))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_do_loop() {
        let source = r#"
Sub Test()
    Do While Int(counter) < limit
        counter = counter + 0.5
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_abs() {
        let source = r#"
Sub Test()
    wholePart = Int(Abs(value))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_array_index() {
        let source = r#"
Sub Test()
    index = Int(Rnd * arraySize)
    item = items(index)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_multiple_operations() {
        let source = r#"
Sub Test()
    quotient = Int(dividend / divisor)
    remainder = dividend - (quotient * divisor)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_int_parentheses() {
        let source = r#"
Sub Test()
    value = (Int(number))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Int"));
        assert!(text.contains("Identifier"));
    }
}
