//! VB6 `Str` Function
//!
//! The `Str` function converts a number to a string representation.
//!
//! ## Syntax
//! ```vb6
//! Str(number)
//! ```
//!
//! ## Parameters
//! - `number`: Required. Any valid numeric expression (Long, Integer, Single, Double, Currency, Byte, Boolean, Variant).
//!
//! ## Returns
//! Returns a `String` representing the numeric value. Positive numbers include a leading space for the sign position,
//! negative numbers include a leading minus sign.
//!
//! ## Remarks
//! The `Str` function converts numeric values to their string representation:
//!
//! - **Always includes sign position**: Positive numbers have a leading space, negative numbers have a leading minus sign (`-`)
//! - **No thousands separators**: Returns plain numeric string without commas or other formatting
//! - **Scientific notation**: Very large or very small numbers may use scientific notation (e.g., "1.23E+10")
//! - **Decimal point**: Uses period (`.`) as decimal separator, regardless of locale settings
//! - **Boolean conversion**: `True` becomes `"-1"`, `False` becomes `" 0"`
//! - **Null handling**: If `number` is `Null`, returns `Null` (not empty string)
//! - **Variant support**: Accepts Variant containing numeric values
//! - **No currency symbol**: Currency values converted without `$` or other symbols
//!
//! ### Comparison with Other Functions
//! - **Str vs `CStr`**: `Str` adds leading space for positive numbers, `CStr` does not
//! - **Str vs Format$**: `Str` has no formatting options, `Format$` provides extensive control
//! - **Str vs Val**: `Str` converts number to string, `Val` converts string to number (inverse operations)
//! - **Str vs `LTrim`$**: Often use `LTrim$(Str(number))` to remove leading space from positive numbers
//!
//! ## Typical Uses
//! 1. **Quick Number Display**: Convert numbers to strings for display without formatting
//! 2. **String Concatenation**: Build messages or output by combining numbers with text
//! 3. **File Output**: Write numeric values to text files in plain format
//! 4. **Debug Output**: Display numeric values in debug messages with minimal formatting
//! 5. **Data Export**: Export numeric data to plain text format
//! 6. **Sign Preservation**: Maintain visual alignment when displaying columns of numbers
//! 7. **Legacy Code**: Support older code that relied on `Str` function behavior
//! 8. **Simple Conversion**: Convert numeric types to string without locale-specific formatting
//!
//! ## Basic Examples
//!
//! ### Example 1: Basic Number Conversion
//! ```vb6
//! Dim result As String
//! result = Str(123)      ' Returns " 123" (note leading space)
//! result = Str(-456)     ' Returns "-456" (leading minus sign)
//! result = Str(0)        ' Returns " 0" (leading space)
//! result = Str(3.14159)  ' Returns " 3.14159"
//! ```
//!
//! ### Example 2: Removing Leading Space
//! ```vb6
//! Dim numStr As String
//! Dim value As Long
//! value = 100
//!
//! ' With leading space
//! numStr = Str(value)           ' " 100"
//!
//! ' Without leading space
//! numStr = LTrim$(Str(value))   ' "100"
//!
//! ' Or use CStr instead
//! numStr = CStr(value)          ' "100"
//! ```
//!
//! ### Example 3: String Concatenation
//! ```vb6
//! Dim message As String
//! Dim count As Integer
//! Dim total As Double
//!
//! count = 5
//! total = 123.45
//!
//! message = "Count:" & Str(count) & ", Total:" & Str(total)
//! ' Result: "Count: 5, Total: 123.45" (spaces from Str included)
//!
//! message = "Count:" & LTrim$(Str(count)) & ", Total:" & LTrim$(Str(total))
//! ' Result: "Count:5, Total:123.45" (no extra spaces)
//! ```
//!
//! ### Example 4: Boolean Conversion
//! ```vb6
//! Dim result As String
//! Dim flag As Boolean
//!
//! flag = True
//! result = Str(flag)   ' Returns "-1"
//!
//! flag = False
//! result = Str(flag)   ' Returns " 0"
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Build Display String
//! ```vb6
//! Function BuildSummary(items As Integer, amount As Currency) As String
//!     ' Build summary string with numbers
//!     BuildSummary = "Items:" & LTrim$(Str(items)) & _
//!                    " Amount: $" & LTrim$(Str(amount))
//! End Function
//! ```
//!
//! ### Pattern 2: Write to File
//! ```vb6
//! Sub WriteDataToFile(filename As String, values() As Double)
//!     Dim fileNum As Integer
//!     Dim i As Integer
//!     
//!     fileNum = FreeFile
//!     Open filename For Output As #fileNum
//!     
//!     For i = LBound(values) To UBound(values)
//!         Print #fileNum, LTrim$(Str(values(i)))
//!     Next i
//!     
//!     Close #fileNum
//! End Sub
//! ```
//!
//! ### Pattern 3: Debug Output
//! ```vb6
//! Sub DebugPrintValues(x As Double, y As Double, z As Double)
//!     Debug.Print "X=" & LTrim$(Str(x)) & _
//!                 " Y=" & LTrim$(Str(y)) & _
//!                 " Z=" & LTrim$(Str(z))
//! End Sub
//! ```
//!
//! ### Pattern 4: Array to String
//! ```vb6
//! Function ArrayToString(arr() As Integer) As String
//!     Dim i As Integer
//!     Dim result As String
//!     
//!     result = ""
//!     For i = LBound(arr) To UBound(arr)
//!         If i > LBound(arr) Then result = result & ","
//!         result = result & LTrim$(Str(arr(i)))
//!     Next i
//!     
//!     ArrayToString = result
//! End Function
//! ```
//!
//! ### Pattern 5: Aligned Column Output
//! ```vb6
//! Sub PrintAlignedNumbers(values() As Long)
//!     Dim i As Integer
//!     Dim numStr As String
//!     
//!     For i = LBound(values) To UBound(values)
//!         numStr = Str(values(i))  ' Keep leading space/sign
//!         Debug.Print Right$(Space(10) & numStr, 10)  ' Right-align
//!     Next i
//! End Sub
//! ```
//!
//! ### Pattern 6: Safe Null Handling
//! ```vb6
//! Function SafeStr(value As Variant) As String
//!     ' Handle potential Null values
//!     If IsNull(value) Then
//!         SafeStr = ""
//!     ElseIf IsNumeric(value) Then
//!         SafeStr = LTrim$(Str(value))
//!     Else
//!         SafeStr = CStr(value)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 7: Build CSV Line
//! ```vb6
//! Function BuildCSVLine(id As Long, quantity As Integer, price As Currency) As String
//!     ' Build comma-separated values
//!     BuildCSVLine = LTrim$(Str(id)) & "," & _
//!                    LTrim$(Str(quantity)) & "," & _
//!                    LTrim$(Str(price))
//! End Function
//! ```
//!
//! ### Pattern 8: Number to Padded String
//! ```vb6
//! Function NumberToPaddedString(value As Long, width As Integer) As String
//!     Dim numStr As String
//!     numStr = LTrim$(Str(value))
//!     NumberToPaddedString = Right$(String$(width, "0") & numStr, width)
//! End Function
//! ```
//!
//! ### Pattern 9: Sign-Aware Display
//! ```vb6
//! Function DisplayWithSign(value As Double) As String
//!     Dim numStr As String
//!     numStr = Str(value)  ' Keep sign/space
//!     
//!     If Left$(numStr, 1) = " " Then
//!         DisplayWithSign = "+" & LTrim$(numStr)
//!     Else
//!         DisplayWithSign = numStr  ' Already has minus sign
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 10: Concatenate Multiple Values
//! ```vb6
//! Function ConcatenateValues(ParamArray values() As Variant) As String
//!     Dim i As Integer
//!     Dim result As String
//!     
//!     result = ""
//!     For i = LBound(values) To UBound(values)
//!         If i > LBound(values) Then result = result & " "
//!         result = result & LTrim$(Str(values(i)))
//!     Next i
//!     
//!     ConcatenateValues = result
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Data Export Utility
//! ```vb6
//! ' Class: DataExporter
//! ' Exports numeric data to various text formats
//! Option Explicit
//!
//! Private m_Delimiter As String
//! Private m_IncludeHeader As Boolean
//! Private m_TrimSpaces As Boolean
//!
//! Public Sub Initialize(delimiter As String, includeHeader As Boolean, trimSpaces As Boolean)
//!     m_Delimiter = delimiter
//!     m_IncludeHeader = includeHeader
//!     m_TrimSpaces = trimSpaces
//! End Sub
//!
//! Public Sub ExportToFile(filename As String, data() As Double, headers() As String)
//!     Dim fileNum As Integer
//!     Dim i As Integer
//!     Dim j As Integer
//!     Dim line As String
//!     
//!     fileNum = FreeFile
//!     Open filename For Output As #fileNum
//!     
//!     ' Write header if requested
//!     If m_IncludeHeader Then
//!         line = Join(headers, m_Delimiter)
//!         Print #fileNum, line
//!     End If
//!     
//!     ' Write data rows
//!     For i = LBound(data, 1) To UBound(data, 1)
//!         line = ""
//!         For j = LBound(data, 2) To UBound(data, 2)
//!             If j > LBound(data, 2) Then line = line & m_Delimiter
//!             If m_TrimSpaces Then
//!                 line = line & LTrim$(Str(data(i, j)))
//!             Else
//!                 line = line & Str(data(i, j))
//!             End If
//!         Next j
//!         Print #fileNum, line
//!     Next i
//!     
//!     Close #fileNum
//! End Sub
//!
//! Public Function ConvertRowToString(values() As Variant) As String
//!     Dim i As Integer
//!     Dim result As String
//!     
//!     result = ""
//!     For i = LBound(values) To UBound(values)
//!         If i > LBound(values) Then result = result & m_Delimiter
//!         
//!         If IsNumeric(values(i)) Then
//!             If m_TrimSpaces Then
//!                 result = result & LTrim$(Str(values(i)))
//!             Else
//!                 result = result & Str(values(i))
//!             End If
//!         Else
//!             result = result & CStr(values(i))
//!         End If
//!     Next i
//!     
//!     ConvertRowToString = result
//! End Function
//! ```
//!
//! ### Example 2: Number Formatter Module
//! ```vb6
//! ' Module: NumberFormatter
//! ' Utilities for converting numbers to formatted strings
//! Option Explicit
//!
//! Public Function ToStringWithoutSpace(value As Variant) As String
//!     ' Convert number to string without leading space
//!     If IsNull(value) Then
//!         ToStringWithoutSpace = ""
//!     ElseIf IsNumeric(value) Then
//!         ToStringWithoutSpace = LTrim$(Str(value))
//!     Else
//!         ToStringWithoutSpace = CStr(value)
//!     End If
//! End Function
//!
//! Public Function ToFixedWidth(value As Double, width As Integer) As String
//!     ' Convert number to fixed-width string, right-aligned
//!     Dim numStr As String
//!     numStr = Str(value)  ' Keep sign position
//!     
//!     If Len(numStr) >= width Then
//!         ToFixedWidth = numStr
//!     Else
//!         ToFixedWidth = Space(width - Len(numStr)) & numStr
//!     End If
//! End Function
//!
//! Public Function ToScientific(value As Double, decimals As Integer) As String
//!     ' Convert to scientific notation manually
//!     Dim exponent As Integer
//!     Dim mantissa As Double
//!     Dim absValue As Double
//!     
//!     If value = 0 Then
//!         ToScientific = "0.0E+00"
//!         Exit Function
//!     End If
//!     
//!     absValue = Abs(value)
//!     exponent = Int(Log(absValue) / Log(10))
//!     mantissa = value / (10 ^ exponent)
//!     
//!     ToScientific = LTrim$(Str(Round(mantissa, decimals))) & "E" & _
//!                    IIf(exponent >= 0, "+", "") & LTrim$(Str(exponent))
//! End Function
//!
//! Public Function PadInteger(value As Long, width As Integer) As String
//!     ' Pad integer with leading zeros
//!     Dim numStr As String
//!     numStr = LTrim$(Str(value))
//!     
//!     If Len(numStr) >= width Then
//!         PadInteger = numStr
//!     Else
//!         PadInteger = String$(width - Len(numStr), "0") & numStr
//!     End If
//! End Function
//! ```
//!
//! ### Example 3: Report Builder Class
//! ```vb6
//! ' Class: ReportBuilder
//! ' Builds formatted text reports with numeric data
//! Option Explicit
//!
//! Private m_Lines As Collection
//! Private m_ColumnWidths() As Integer
//! Private m_ColumnCount As Integer
//!
//! Public Sub Initialize(columnCount As Integer)
//!     Set m_Lines = New Collection
//!     m_ColumnCount = columnCount
//!     ReDim m_ColumnWidths(1 To columnCount)
//!     
//!     ' Default column widths
//!     Dim i As Integer
//!     For i = 1 To columnCount
//!         m_ColumnWidths(i) = 10
//!     Next i
//! End Sub
//!
//! Public Sub SetColumnWidth(column As Integer, width As Integer)
//!     If column >= 1 And column <= m_ColumnCount Then
//!         m_ColumnWidths(column) = width
//!     End If
//! End Sub
//!
//! Public Sub AddRow(ParamArray values() As Variant)
//!     Dim i As Integer
//!     Dim line As String
//!     Dim cell As String
//!     Dim colNum As Integer
//!     
//!     line = ""
//!     For i = LBound(values) To UBound(values)
//!         colNum = i + 1
//!         If colNum > m_ColumnCount Then Exit For
//!         
//!         ' Convert value to string
//!         If IsNumeric(values(i)) Then
//!             cell = Str(values(i))  ' Keep sign position for alignment
//!         Else
//!             cell = CStr(values(i))
//!         End If
//!         
//!         ' Pad to column width
//!         If Len(cell) < m_ColumnWidths(colNum) Then
//!             cell = Space(m_ColumnWidths(colNum) - Len(cell)) & cell
//!         End If
//!         
//!         line = line & cell
//!     Next i
//!     
//!     m_Lines.Add line
//! End Sub
//!
//! Public Function GetReport() As String
//!     Dim i As Integer
//!     Dim result As String
//!     
//!     result = ""
//!     For i = 1 To m_Lines.Count
//!         result = result & m_Lines(i) & vbCrLf
//!     Next i
//!     
//!     GetReport = result
//! End Function
//!
//! Public Sub Clear()
//!     Set m_Lines = New Collection
//! End Sub
//! ```
//!
//! ### Example 4: String Builder for Numbers
//! ```vb6
//! ' Module: NumericStringBuilder
//! ' Build strings from numeric arrays and collections
//! Option Explicit
//!
//! Public Function JoinNumbers(numbers() As Double, delimiter As String, trimSpaces As Boolean) As String
//!     Dim i As Integer
//!     Dim result As String
//!     Dim numStr As String
//!     
//!     result = ""
//!     For i = LBound(numbers) To UBound(numbers)
//!         If i > LBound(numbers) Then result = result & delimiter
//!         
//!         numStr = Str(numbers(i))
//!         If trimSpaces Then numStr = LTrim$(numStr)
//!         
//!         result = result & numStr
//!     Next i
//!     
//!     JoinNumbers = result
//! End Function
//!
//! Public Function CreateTable(data() As Variant, rowCount As Integer, colCount As Integer) As String
//!     Dim i As Integer
//!     Dim j As Integer
//!     Dim result As String
//!     Dim row As String
//!     
//!     result = ""
//!     For i = 1 To rowCount
//!         row = ""
//!         For j = 1 To colCount
//!             If j > 1 Then row = row & vbTab
//!             
//!             If IsNumeric(data(i, j)) Then
//!                 row = row & LTrim$(Str(data(i, j)))
//!             Else
//!                 row = row & CStr(data(i, j))
//!             End If
//!         Next j
//!         result = result & row & vbCrLf
//!     Next i
//!     
//!     CreateTable = result
//! End Function
//!
//! Public Function FormatNumberList(numbers As Collection, prefix As String, suffix As String) As String
//!     Dim item As Variant
//!     Dim result As String
//!     Dim first As Boolean
//!     
//!     first = True
//!     result = ""
//!     
//!     For Each item In numbers
//!         If Not first Then result = result & ", "
//!         result = result & prefix & LTrim$(Str(item)) & suffix
//!         first = False
//!     Next item
//!     
//!     FormatNumberList = result
//! End Function
//!
//! Public Function CreateSummaryLine(label As String, value As Double, width As Integer) As String
//!     Dim numStr As String
//!     Dim padding As Integer
//!     
//!     numStr = LTrim$(Str(value))
//!     padding = width - Len(label) - Len(numStr)
//!     
//!     If padding > 0 Then
//!         CreateSummaryLine = label & Space(padding) & numStr
//!     Else
//!         CreateSummaryLine = label & " " & numStr
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//! The `Str` function can raise the following errors:
//!
//! - **Error 13 (Type mismatch)**: If `number` cannot be converted to a numeric value
//! - **Error 94 (Invalid use of Null)**: In some contexts when passing Null without proper handling
//!
//! ## Performance Notes
//! - Very fast for basic numeric types (Integer, Long, Single, Double)
//! - Slightly slower for Variant types due to type checking
//! - No performance overhead from locale or formatting
//! - Direct conversion without intermediate formatting steps
//! - Consider `CStr` if leading space is always unwanted (avoids `LTrim$` call)
//!
//! ## Best Practices
//! 1. **Use `LTrim$`** to remove leading space from positive numbers when space is unwanted
//! 2. **Consider `CStr`** if you never need the leading space (cleaner code)
//! 3. **Preserve sign position** when aligning columns of numbers (keep the `Str` space)
//! 4. **Handle Null values** explicitly with `IsNull` check before calling `Str`
//! 5. **Use `Format$`** instead for formatted output (thousands separators, decimal places)
//! 6. **Document behavior** when `Str` is used in shared code (leading space can surprise developers)
//! 7. **Combine with `Val`** for round-trip conversion (number → string → number)
//! 8. **Avoid locale issues** by using `Str` for culture-invariant numeric strings
//! 9. **Cache results** if converting the same number repeatedly in a loop
//! 10. **Use type-specific functions** like `CStr` when working with non-numeric types
//!
//! ## Comparison Table
//!
//! | Function | Input | Output | Leading Space | Formatting |
//! |----------|-------|--------|---------------|------------|
//! | `Str` | Number | String | Yes (positive) | None |
//! | `CStr` | Any | String | No | None |
//! | `Format$` | Any | String | No | Extensive |
//! | `Val` | String | Double | N/A | N/A |
//! | `LTrim$` | String | String | Removes | None |
//!
//! ## Platform Notes
//! - Available in VB6, VBA, and `VBScript`
//! - Behavior consistent across all platforms
//! - Uses period (`.`) as decimal separator regardless of locale
//! - Scientific notation format is standard across platforms
//! - No localization or culture-specific formatting applied
//!
//! ## Limitations
//! - Always includes leading space for positive numbers (often unwanted)
//! - No formatting options (thousands separators, decimal places, etc.)
//! - Cannot specify decimal precision
//! - Cannot control scientific notation threshold
//! - No currency symbol support
//! - Null input can cause errors if not handled
//! - No culture-aware formatting
//! - Return value length is unpredictable (depends on number magnitude)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn str_basic() {
        let source = r#"
Sub Test()
    result = Str(123)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn str_variable_assignment() {
        let source = r#"
Sub Test()
    Dim s As String
    s = Str(value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
        assert!(debug.contains("value"));
    }

    #[test]
    fn str_negative_number() {
        let source = r#"
Sub Test()
    result = Str(-456)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_decimal() {
        let source = r#"
Sub Test()
    result = Str(3.14159)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_with_ltrim() {
        let source = r#"
Sub Test()
    result = LTrim$(Str(value))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
        assert!(debug.contains("LTrim"));
    }

    #[test]
    fn str_concatenation() {
        let source = r#"
Sub Test()
    message = "Count:" & Str(count)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
        assert!(debug.contains("count"));
    }

    #[test]
    fn str_in_print() {
        let source = r#"
Sub Test()
    Print #1, Str(value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print Str(x)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_if_statement() {
        let source = r#"
Sub Test()
    If Str(value) = " 0" Then
        MsgBox "Zero"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_for_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        result = result & Str(i)
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_array_assignment() {
        let source = r#"
Sub Test()
    arr(i) = Str(values(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_function_argument() {
        let source = r#"
Sub Test()
    Call ProcessString(Str(number))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_function_return() {
        let source = r#"
Function GetNumberString() As String
    GetNumberString = Str(value)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_comparison() {
        let source = r#"
Sub Test()
    If Len(Str(num)) > 5 Then
        MsgBox "Long number"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_select_case() {
        let source = r#"
Sub Test()
    Select Case Str(value)
        Case " 0"
            MsgBox "Zero"
        Case Else
            MsgBox "Other"
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_with_statement() {
        let source = r#"
Sub Test()
    With obj
        .Text = Str(.Value)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_do_while() {
        let source = r#"
Sub Test()
    Do While Len(Str(counter)) < 10
        counter = counter + 1
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_do_until() {
        let source = r#"
Sub Test()
    Do Until Str(value) = " 100"
        value = value + 1
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_while_wend() {
        let source = r#"
Sub Test()
    While i < 10
        output = output & Str(i)
        i = i + 1
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_iif() {
        let source = r#"
Sub Test()
    result = IIf(flag, Str(value1), Str(value2))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_parentheses() {
        let source = r#"
Sub Test()
    result = Str((x + y) * z)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Value: " & Str(total)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_property_assignment() {
        let source = r#"
Sub Test()
    obj.Caption = Str(count)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    result = Str(varValue)
    If Err.Number <> 0 Then
        result = ""
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_class_usage() {
        let source = r#"
Sub Test()
    Set obj = New MyClass
    obj.SetValue Str(counter)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_currency_conversion() {
        let source = r#"
Sub Test()
    Dim amount As Currency
    amount = 123.45
    result = Str(amount)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }

    #[test]
    fn str_csv_building() {
        let source = r#"
Sub Test()
    csvLine = LTrim$(Str(id)) & "," & LTrim$(Str(qty)) & "," & LTrim$(Str(price))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Str"));
    }
}
