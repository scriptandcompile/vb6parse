//! # `Str$` Function
//!
//! The `Str$` function in Visual Basic 6 converts a numeric value to a string representation.
//! The dollar sign (`$`) suffix indicates that this function always returns a `String` type,
//! never a `Variant`.
//!
//! ## Syntax
//!
//! ```vb6
//! Str$(number)
//! ```
//!
//! ## Parameters
//!
//! - `number` - Required. Any valid numeric expression. Can be of type `Byte`, `Integer`, `Long`,
//!   `Single`, `Double`, or `Currency`.
//!
//! ## Return Value
//!
//! Returns a `String` representation of the number. Positive numbers include a leading space
//! for the sign position. Negative numbers include a leading minus sign (-).
//!
//! ## Behavior and Characteristics
//!
//! ### Sign Handling
//!
//! - Positive numbers: Include a leading space (e.g., " 123")
//! - Negative numbers: Include a leading minus sign (e.g., "-123")
//! - Zero: Returns " 0" (with leading space)
//! - The leading space reserves position for the sign
//!
//! ### Numeric Formatting
//!
//! - No thousands separators (e.g., "1000" not "1,000")
//! - Scientific notation for very large or very small numbers
//! - Floating-point numbers may show precision artifacts
//! - No control over decimal places
//!
//! ### Type Differences: `Str$` vs `Str`
//!
//! - `Str$`: Always returns `String` type (never `Variant`)
//! - `Str`: Returns `Variant` containing a string
//! - Use `Str$` when you need guaranteed `String` return type
//! - Use `Str` when working with `Variant` variables
//!
//! ## Common Usage Patterns
//!
//! ### 1. Basic Number to String Conversion
//!
//! ```vb6
//! Dim numStr As String
//! numStr = Str$(123)  ' Returns " 123" (note leading space)
//! numStr = Str$(-45)  ' Returns "-45"
//! ```
//!
//! ### 2. Concatenating Numbers with Text
//!
//! ```vb6
//! Function FormatMessage(count As Integer) As String
//!     FormatMessage = "Found" & Str$(count) & " items"
//! End Function
//!
//! Debug.Print FormatMessage(5)  ' "Found 5 items"
//! ```
//!
//! ### 3. Trimming the Leading Space
//!
//! ```vb6
//! Function NumberToString(value As Long) As String
//!     NumberToString = LTrim$(Str$(value))
//! End Function
//!
//! Dim result As String
//! result = NumberToString(100)  ' Returns "100" (no leading space)
//! ```
//!
//! ### 4. Building Comma-Separated Values
//!
//! ```vb6
//! Function BuildCSV(values() As Integer) As String
//!     Dim i As Integer
//!     Dim result As String
//!     For i = LBound(values) To UBound(values)
//!         If i > LBound(values) Then result = result & ","
//!         result = result & LTrim$(Str$(values(i)))
//!     Next i
//!     BuildCSV = result
//! End Function
//! ```
//!
//! ### 5. Logging and Debug Output
//!
//! ```vb6
//! Sub LogValue(name As String, value As Double)
//!     Debug.Print name & " =" & Str$(value)
//! End Sub
//! ```
//!
//! ### 6. Creating Numeric Labels
//!
//! ```vb6
//! Function CreateLabel(index As Integer) As String
//!     CreateLabel = "Item" & LTrim$(Str$(index))
//! End Function
//!
//! Dim label As String
//! label = CreateLabel(42)  ' Returns "Item42"
//! ```
//!
//! ### 7. File Output Formatting
//!
//! ```vb6
//! Sub WriteDataLine(fileNum As Integer, id As Long, amount As Currency)
//!     Print #fileNum, LTrim$(Str$(id)) & "," & LTrim$(Str$(amount))
//! End Sub
//! ```
//!
//! ### 8. Array Index Display
//!
//! ```vb6
//! Sub ShowArrayContents(arr() As Integer)
//!     Dim i As Integer
//!     For i = LBound(arr) To UBound(arr)
//!         Debug.Print "[" & LTrim$(Str$(i)) & "] = " & LTrim$(Str$(arr(i)))
//!     Next i
//! End Sub
//! ```
//!
//! ### 9. Simple Calculator Display
//!
//! ```vb6
//! Function UpdateDisplay(value As Double) As String
//!     UpdateDisplay = LTrim$(Str$(value))
//! End Function
//! ```
//!
//! ### 10. Building SQL Statements
//!
//! ```vb6
//! Function BuildQuery(userId As Long) As String
//!     BuildQuery = "SELECT * FROM Users WHERE ID = " & LTrim$(Str$(userId))
//! End Function
//! ```
//!
//! ## Related Functions
//!
//! - `Str()` - Returns a `Variant` containing the string representation of a number
//! - `CStr()` - Converts an expression to a `String` (no leading space for positive numbers)
//! - `Format$()` - Provides extensive formatting control for numeric values
//! - `Val()` - Converts a string to a numeric value (inverse operation)
//! - `LTrim$()` - Removes leading spaces (often used with `Str$`)
//! - `Hex$()` - Converts a number to hexadecimal string
//! - `Oct$()` - Converts a number to octal string
//!
//! ## Best Practices
//!
//! ### When to Use `Str$` vs `CStr` vs `Format$`
//!
//! ```vb6
//! Dim value As Integer
//! value = 42
//!
//! ' Str$ includes leading space for positive numbers
//! Debug.Print Str$(value)  ' " 42"
//!
//! ' CStr has no leading space
//! Debug.Print CStr(value)  ' "42"
//!
//! ' Format$ provides control over formatting
//! Debug.Print Format$(value, "000")  ' "042"
//! ```
//!
//! ### Always Trim for Display
//!
//! ```vb6
//! ' Without trim (has leading space for positive numbers)
//! Label1.Caption = Str$(count)  ' " 5"
//!
//! ' With trim (clean output)
//! Label1.Caption = LTrim$(Str$(count))  ' "5"
//!
//! ' Or use CStr instead
//! Label1.Caption = CStr(count)  ' "5"
//! ```
//!
//! ### Use `Format$` for Formatted Output
//!
//! ```vb6
//! ' Str$ has no formatting control
//! Debug.Print Str$(1234.5678)  ' " 1234.5678"
//!
//! ' Format$ provides control
//! Debug.Print Format$(1234.5678, "#,##0.00")  ' "1,234.57"
//! ```
//!
//! ### Handle Negative Numbers
//!
//! ```vb6
//! Function SafeConvert(value As Long) As String
//!     ' Str$ handles negative numbers correctly
//!     SafeConvert = LTrim$(Str$(value))
//!     ' For negative: "-123", for positive: "123"
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - `Str$` is very fast for simple conversions
//! - Faster than `Format$` when formatting is not needed
//! - Similar performance to `CStr`
//! - No significant overhead for any numeric type
//!
//! ```vb6
//! ' Fast: simple conversion
//! For i = 1 To 10000
//!     text = LTrim$(Str$(i))
//! Next i
//!
//! ' Slower: formatted conversion (but more control)
//! For i = 1 To 10000
//!     text = Format$(i, "0000")
//! Next i
//! ```
//!
//! ## Common Pitfalls
//!
//! ### 1. Leading Space for Positive Numbers
//!
//! ```vb6
//! Dim result As String
//! result = Str$(100)  ' " 100" (note the leading space!)
//!
//! ' This can cause problems in comparisons
//! If Str$(100) = "100" Then  ' FALSE! (" 100" <> "100")
//!     Debug.Print "Match"
//! End If
//!
//! ' Use LTrim$ or CStr instead
//! If LTrim$(Str$(100)) = "100" Then  ' TRUE
//!     Debug.Print "Match"
//! End If
//! ```
//!
//! ### 2. Confusion with `CStr`
//!
//! ```vb6
//! Dim value As Integer
//! value = 42
//!
//! Debug.Print Str$(value)   ' " 42" (with space)
//! Debug.Print CStr(value)   ' "42" (no space)
//!
//! ' Know which one you need
//! ```
//!
//! ### 3. No Formatting Control
//!
//! ```vb6
//! Dim amount As Currency
//! amount = 1234.56
//!
//! ' Str$ gives no control
//! Debug.Print Str$(amount)  ' " 1234.56"
//!
//! ' Use Format$ for currency
//! Debug.Print Format$(amount, "$#,##0.00")  ' "$1,234.56"
//! ```
//!
//! ### 4. Floating-Point Precision Issues
//!
//! ```vb6
//! Dim value As Double
//! value = 0.1 + 0.2
//!
//! Debug.Print Str$(value)  ' May show " 0.30000000000000004"
//!
//! ' Use Format$ to control precision
//! Debug.Print Format$(value, "0.00")  ' "0.30"
//! ```
//!
//! ### 5. Not Handling Very Large or Small Numbers
//!
//! ```vb6
//! Dim bigNum As Double
//! bigNum = 1E+20
//!
//! Debug.Print Str$(bigNum)  ' " 1E+20" (scientific notation)
//!
//! ' Be aware of scientific notation in output
//! ```
//!
//! ### 6. Null Values
//!
//! ```vb6
//! ' Str$ cannot handle Null
//! Dim result As String
//! result = Str$(nullValue)  ' Runtime error if nullValue is Null
//!
//! ' Check first
//! If Not IsNull(value) Then
//!     result = Str$(value)
//! Else
//!     result = ""
//! End If
//! ```
//!
//! ## Practical Examples
//!
//! ### Building a Progress Message
//!
//! ```vb6
//! Function ProgressMessage(current As Long, total As Long) As String
//!     ProgressMessage = "Processing item" & Str$(current) & _
//!                      " of" & Str$(total)
//! End Function
//!
//! Debug.Print ProgressMessage(5, 10)  ' "Processing item 5 of 10"
//! ```
//!
//! ### Creating Sequential Filenames
//!
//! ```vb6
//! Function GenerateFileName(baseNameStr As String, index As Integer) As String
//!     GenerateFileName = baseNameStr & LTrim$(Str$(index)) & ".dat"
//! End Function
//!
//! Dim fileName As String
//! fileName = GenerateFileName("data", 1)  ' "data1.dat"
//! ```
//!
//! ### Simple Data Export
//!
//! ```vb6
//! Sub ExportToCSV(data() As Double, fileName As String)
//!     Dim i As Integer
//!     Dim lineData As String
//!     
//!     Open fileName For Output As #1
//!     For i = LBound(data) To UBound(data)
//!         Print #1, LTrim$(Str$(data(i)))
//!     Next i
//!     Close #1
//! End Sub
//! ```
//!
//! ## Limitations
//!
//! - Always includes leading space for positive numbers (use `LTrim$` or `CStr` to remove)
//! - No formatting control (no thousands separators, decimal places, etc.)
//! - Cannot handle `Null` values (use `CStr` with error handling instead)
//! - May produce scientific notation for very large or small numbers
//! - Floating-point precision artifacts may appear in output
//! - No locale-specific formatting (always uses invariant format)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn str_dollar_simple() {
        let source = r"
Sub Main()
    result = Str$(123)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_assignment() {
        let source = r"
Sub Main()
    Dim numStr As String
    numStr = Str$(456)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_variable() {
        let source = r"
Sub Main()
    Dim value As Integer
    Dim text As String
    value = 100
    text = Str$(value)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_negative() {
        let source = r"
Sub Main()
    Dim result As String
    result = Str$(-42)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_with_ltrim() {
        let source = r"
Function NumberToString(value As Long) As String
    NumberToString = LTrim$(Str$(value))
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_concatenation() {
        let source = r#"
Function FormatMessage(count As Integer) As String
    FormatMessage = "Found" & Str$(count) & " items"
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_in_condition() {
        let source = r#"
Sub Main()
    If LTrim$(Str$(value)) = "100" Then
        Debug.Print "Match"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_create_label() {
        let source = r#"
Function CreateLabel(index As Integer) As String
    CreateLabel = "Item" & LTrim$(Str$(index))
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_logging() {
        let source = r#"
Sub LogValue(valueName As String, value As Double)
    Debug.Print valueName & " =" & Str$(value)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_multiple_uses() {
        let source = r"
Sub ProcessValues()
    Dim text1 As String
    Dim text2 As String
    text1 = Str$(10)
    text2 = Str$(20)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_select_case() {
        let source = r#"
Sub Main()
    Select Case LTrim$(Str$(code))
        Case "1"
            Debug.Print "One"
        Case "2"
            Debug.Print "Two"
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_csv_building() {
        let source = r#"
Function BuildCSV(val1 As Integer, val2 As Integer) As String
    BuildCSV = LTrim$(Str$(val1)) & "," & LTrim$(Str$(val2))
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_file_output() {
        let source = r"
Sub WriteData(fileNum As Integer, id As Long)
    Print #fileNum, LTrim$(Str$(id))
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_loop_processing() {
        let source = r#"
Sub ShowArray(arr() As Integer)
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        Debug.Print "[" & Str$(i) & "]"
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_progress_message() {
        let source = r#"
Function ProgressMessage(current As Long, total As Long) As String
    ProgressMessage = "Item" & Str$(current) & " of" & Str$(total)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_filename_generation() {
        let source = r#"
Function GenerateFileName(baseName As String, index As Integer) As String
    GenerateFileName = baseName & LTrim$(Str$(index)) & ".dat"
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_double_value() {
        let source = r"
Sub Main()
    Dim result As String
    Dim value As Double
    value = 123.45
    result = Str$(value)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_currency_value() {
        let source = r"
Sub Main()
    Dim amount As Currency
    Dim text As String
    amount = 1234.56
    text = Str$(amount)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_in_function() {
        let source = r"
Function ConvertNumber(num As Long) As String
    ConvertNumber = Str$(num)
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }

    #[test]
    fn str_dollar_sql_building() {
        let source = r#"
Function BuildQuery(userId As Long) As String
    BuildQuery = "SELECT * FROM Users WHERE ID = " & LTrim$(Str$(userId))
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Str$"));
    }
}
