/// # Spc Function
///
/// Used with the Print # statement or Print method to position output.
///
/// ## Syntax
///
/// ```vb
/// Spc(n)
/// ```
///
/// ## Parameters
///
/// - `n` - Required. Numeric expression specifying the number of spaces to insert before displaying or printing the next expression in the list.
///
/// ## Return Value
///
/// Used only with Print # statement and Print method. Returns a value used to insert space characters in output.
///
/// ## Remarks
///
/// The Spc function is used to insert spaces in output when using the Print # statement (for files) or the Print method (for the Immediate window or printer). Unlike the Space function which returns a string of spaces, Spc is a special formatting function that only works within Print statements.
///
/// Key characteristics:
/// - Only valid within Print # or Print (Debug.Print) statements
/// - Inserts the specified number of space characters
/// - If current print position + n exceeds output line width, Spc skips to the next line
/// - If n is less than the output line width, the next print position is current position + n
/// - If n is greater than the output line width, the next print position is calculated as n Mod width
/// - Cannot be used independently or assigned to variables
/// - Returns a variant used internally by the Print statement
///
/// The Spc function differs from related positioning functions:
/// - **Spc(n)**: Inserts n spaces from current position
/// - **Tab(n)**: Moves to column n (absolute positioning)
/// - **Space(n)**: Returns a string of n spaces (can be used anywhere)
///
/// ## Typical Uses
///
/// 1. **File Output**: Space elements in text file output
/// 2. **Report Formatting**: Create formatted reports with spacing
/// 3. **Column Alignment**: Align data in columns
/// 4. **Debug Output**: Format Debug.Print output
/// 5. **Printer Output**: Format printed output
/// 6. **Data Separation**: Separate data elements with spaces
/// 7. **Fixed-Width Output**: Create fixed-width formatted text
/// 8. **Log Files**: Format log file entries
///
/// ## Basic Examples
///
/// ```vb
/// ' Example 1: Print with spaces between items
/// Dim fileNum As Integer
/// fileNum = FreeFile
/// Open "output.txt" For Output As #fileNum
/// Print #fileNum, "Name"; Spc(10); "Age"; Spc(10); "City"
/// Close #fileNum
/// ' Outputs: "Name          Age          City"
/// ```
///
/// ```vb
/// ' Example 2: Debug output with spacing
/// Debug.Print "Item:"; Spc(5); "Value"
/// ' Outputs: "Item:     Value"
/// ```
///
/// ```vb
/// ' Example 3: Multiple Spc calls in one statement
/// Print #1, "A"; Spc(3); "B"; Spc(5); "C"
/// ' Outputs: "A   B     C"
/// ```
///
/// ```vb
/// ' Example 4: Create formatted columns
/// Dim i As Integer
/// For i = 1 To 3
///     Debug.Print i; Spc(10); i * 10; Spc(10); i * 100
/// Next i
/// ' Outputs aligned columns of numbers
/// ```
///
/// ## Common Patterns
///
/// ### Pattern 1: `FormatFileOutput`
/// Format data in file with consistent spacing
/// ```vb
/// Sub FormatFileOutput(fileNum As Integer, name As String, _
///                      age As Integer, city As String)
///     Print #fileNum, name; Spc(20 - Len(name)); _
///                     age; Spc(10); city
/// End Sub
/// ```
///
/// ### Pattern 2: `PrintAlignedData`
/// Print data with aligned columns
/// ```vb
/// Sub PrintAlignedData(label As String, value As Variant, spacing As Integer)
///     Debug.Print label; Spc(spacing); value
/// End Sub
/// ```
///
/// ### Pattern 3: `CreateTableRow`
/// Create table row with Spc spacing
/// ```vb
/// Sub CreateTableRow(fileNum As Integer, col1 As String, _
///                    col2 As String, col3 As String)
///     Print #fileNum, col1; Spc(15 - Len(col1)); _
///                     col2; Spc(15 - Len(col2)); _
///                     col3
/// End Sub
/// ```
///
/// ### Pattern 4: `FormatLogEntry`
/// Format log entries with timestamps and messages
/// ```vb
/// Sub FormatLogEntry(fileNum As Integer, timestamp As String, _
///                    level As String, message As String)
///     Print #fileNum, timestamp; Spc(5); _
///                     level; Spc(10 - Len(level)); _
///                     message
/// End Sub
/// ```
///
/// ### Pattern 5: `PrintWithIndent`
/// Print text with indentation using Spc
/// ```vb
/// Sub PrintWithIndent(fileNum As Integer, indentLevel As Integer, _
///                     text As String)
///     Print #fileNum, Spc(indentLevel * 4); text
/// End Sub
/// ```
///
/// ### Pattern 6: `FormatKeyValuePair`
/// Format key-value pairs with consistent spacing
/// ```vb
/// Sub FormatKeyValuePair(fileNum As Integer, key As String, _
///                        value As String, Optional totalWidth As Integer = 40)
///     Dim spacesNeeded As Integer
///     spacesNeeded = totalWidth - Len(key) - Len(value)
///     If spacesNeeded < 1 Then spacesNeeded = 1
///     Print #fileNum, key; Spc(spacesNeeded); value
/// End Sub
/// ```
///
/// ### Pattern 7: `PrintHeader`
/// Print formatted header with separators
/// ```vb
/// Sub PrintHeader(fileNum As Integer, title1 As String, _
///                 title2 As String, title3 As String)
///     Print #fileNum, title1; Spc(15 - Len(title1)); _
///                     title2; Spc(15 - Len(title2)); _
///                     title3
///     Print #fileNum, String(15, "-"); Spc(1); _
///                     String(15, "-"); Spc(1); _
///                     String(15, "-")
/// End Sub
/// ```
///
/// ### Pattern 8: `DebugPrintArray`
/// Print array elements with spacing
/// ```vb
/// Sub DebugPrintArray(arr() As Variant)
///     Dim i As Integer
///     For i = LBound(arr) To UBound(arr)
///         Debug.Print arr(i); Spc(5);
///     Next i
///     Debug.Print  ' New line
/// End Sub
/// ```
///
/// ### Pattern 9: `FormatNumericTable`
/// Print numeric data in aligned columns
/// ```vb
/// Sub FormatNumericTable(fileNum As Integer, values() As Double)
///     Dim i As Integer
///     For i = LBound(values) To UBound(values)
///         Print #fileNum, Format(values(i), "0.00"); Spc(10);
///         If (i - LBound(values) + 1) Mod 5 = 0 Then
///             Print #fileNum,  ' New line every 5 values
///         End If
///     Next i
/// End Sub
/// ```
///
/// ### Pattern 10: `PrintReportLine`
/// Print formatted report line
/// ```vb
/// Sub PrintReportLine(fileNum As Integer, lineNum As Integer, _
///                     description As String, amount As Double)
///     Print #fileNum, Format(lineNum, "000"); Spc(5); _
///                     description; Spc(30 - Len(description)); _
///                     Format(amount, "$#,##0.00")
/// End Sub
/// ```
///
/// ## Advanced Usage
///
/// ### Example 1: `ReportWriter` Class
/// Generate formatted text reports with Spc
/// ```vb
/// ' Class: ReportWriter
/// Private m_fileNum As Integer
/// Private m_isOpen As Boolean
///
/// Public Sub OpenReport(fileName As String)
///     m_fileNum = FreeFile
///     Open fileName For Output As #m_fileNum
///     m_isOpen = True
/// End Sub
///
/// Public Sub WriteHeader(title As String, col1 As String, _
///                        col2 As String, col3 As String)
///     If Not m_isOpen Then Exit Sub
///     
///     ' Center title
///     Dim totalWidth As Integer
///     totalWidth = 60
///     Dim leftPad As Integer
///     leftPad = (totalWidth - Len(title)) \ 2
///     Print #m_fileNum, Spc(leftPad); title
///     Print #m_fileNum, String(totalWidth, "=")
///     Print #m_fileNum,
///     
///     ' Column headers
///     Print #m_fileNum, col1; Spc(20 - Len(col1)); _
///                       col2; Spc(20 - Len(col2)); _
///                       col3
///     Print #m_fileNum, String(20, "-"); Spc(1); _
///                       String(20, "-"); Spc(1); _
///                       String(20, "-")
/// End Sub
///
/// Public Sub WriteDataRow(val1 As String, val2 As String, val3 As String)
///     If Not m_isOpen Then Exit Sub
///     
///     Print #m_fileNum, val1; Spc(20 - Len(val1)); _
///                       val2; Spc(20 - Len(val2)); _
///                       val3
/// End Sub
///
/// Public Sub WriteSummary(label As String, value As String)
///     If Not m_isOpen Then Exit Sub
///     
///     Print #m_fileNum,
///     Print #m_fileNum, String(60, "-")
///     Print #m_fileNum, label; Spc(60 - Len(label) - Len(value)); value
/// End Sub
///
/// Public Sub CloseReport()
///     If m_isOpen Then
///         Close #m_fileNum
///         m_isOpen = False
///     End If
/// End Sub
///
/// Private Sub Class_Terminate()
///     CloseReport
/// End Sub
/// ```
///
/// ### Example 2: `LogFileFormatter` Module
/// Format log file entries with timestamps
/// ```vb
/// ' Module: LogFileFormatter
/// Private m_logFile As Integer
/// Private m_logOpen As Boolean
///
/// Public Sub OpenLog(fileName As String)
///     m_logFile = FreeFile
///     Open fileName For Append As #m_logFile
///     m_logOpen = True
/// End Sub
///
/// Public Sub LogInfo(message As String)
///     If Not m_logOpen Then Exit Sub
///     Dim timestamp As String
///     timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
///     Print #m_logFile, timestamp; Spc(5); "INFO"; Spc(10); message
/// End Sub
///
/// Public Sub LogWarning(message As String)
///     If Not m_logOpen Then Exit Sub
///     Dim timestamp As String
///     timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
///     Print #m_logFile, timestamp; Spc(5); "WARNING"; Spc(6); message
/// End Sub
///
/// Public Sub LogError(message As String, Optional errorNum As Long = 0)
///     If Not m_logOpen Then Exit Sub
///     Dim timestamp As String
///     timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
///     Print #m_logFile, timestamp; Spc(5); "ERROR"; Spc(8); message
///     If errorNum <> 0 Then
///         Print #m_logFile, Spc(30); "Error #"; errorNum
///     End If
/// End Sub
///
/// Public Sub LogDebug(category As String, message As String)
///     If Not m_logOpen Then Exit Sub
///     Dim timestamp As String
///     timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
///     Print #m_logFile, timestamp; Spc(5); "DEBUG"; Spc(8); _
///                        "["; category; "]"; Spc(3); message
/// End Sub
///
/// Public Sub LogSeparator()
///     If Not m_logOpen Then Exit Sub
///     Print #m_logFile, String(80, "-")
/// End Sub
///
/// Public Sub CloseLog()
///     If m_logOpen Then
///         Close #m_logFile
///         m_logOpen = False
///     End If
/// End Sub
/// ```
///
/// ### Example 3: `DataTablePrinter` Class
/// Print data in formatted tables
/// ```vb
/// ' Class: DataTablePrinter
/// Private m_columnWidths() As Integer
/// Private m_fileNum As Integer
///
/// Public Sub Initialize(fileNum As Integer, columnWidths() As Integer)
///     Dim i As Integer
///     m_fileNum = fileNum
///     ReDim m_columnWidths(LBound(columnWidths) To UBound(columnWidths))
///     For i = LBound(columnWidths) To UBound(columnWidths)
///         m_columnWidths(i) = columnWidths(i)
///     Next i
/// End Sub
///
/// Public Sub PrintRow(values() As String)
///     Dim i As Integer
///     Dim spaces As Integer
///     
///     For i = LBound(values) To UBound(values)
///         Print #m_fileNum, values(i);
///         
///         If i < UBound(values) Then
///             spaces = m_columnWidths(i) - Len(values(i))
///             If spaces < 1 Then spaces = 1
///             Print #m_fileNum, Spc(spaces);
///         End If
///     Next i
///     Print #m_fileNum,  ' New line
/// End Sub
///
/// Public Sub PrintHeaderRow(headers() As String)
///     Dim i As Integer
///     
///     PrintRow headers
///     
///     ' Print separator
///     For i = LBound(m_columnWidths) To UBound(m_columnWidths)
///         Print #m_fileNum, String(m_columnWidths(i), "-");
///         If i < UBound(m_columnWidths) Then
///             Print #m_fileNum, Spc(1);
///         End If
///     Next i
///     Print #m_fileNum,  ' New line
/// End Sub
///
/// Public Sub PrintRightAligned(values() As String)
///     Dim i As Integer
///     Dim spaces As Integer
///     
///     For i = LBound(values) To UBound(values)
///         spaces = m_columnWidths(i) - Len(values(i))
///         If spaces > 0 Then Print #m_fileNum, Spc(spaces);
///         Print #m_fileNum, values(i);
///         
///         If i < UBound(values) Then
///             Print #m_fileNum, Spc(1);
///         End If
///     Next i
///     Print #m_fileNum,  ' New line
/// End Sub
/// ```
///
/// ### Example 4: `DebugOutputHelper` Module
/// Format debug output with Spc
/// ```vb
/// ' Module: DebugOutputHelper
///
/// Public Sub PrintVariable(varName As String, varValue As Variant)
///     Debug.Print varName; Spc(20 - Len(varName)); "="; Spc(2); varValue
/// End Sub
///
/// Public Sub PrintVariables(ParamArray vars() As Variant)
///     Dim i As Integer
///     For i = LBound(vars) To UBound(vars) Step 2
///         If i + 1 <= UBound(vars) Then
///             PrintVariable CStr(vars(i)), vars(i + 1)
///         End If
///     Next i
/// End Sub
///
/// Public Sub PrintSection(title As String)
///     Debug.Print
///     Debug.Print String(50, "=")
///     Dim leftPad As Integer
///     leftPad = (50 - Len(title)) \ 2
///     Debug.Print Spc(leftPad); title
///     Debug.Print String(50, "=")
///     Debug.Print
/// End Sub
///
/// Public Sub PrintKeyValue(key As String, value As Variant, _
///                          Optional totalWidth As Integer = 40)
///     Dim spacesNeeded As Integer
///     spacesNeeded = totalWidth - Len(key) - Len(CStr(value))
///     If spacesNeeded < 1 Then spacesNeeded = 1
///     Debug.Print key; Spc(spacesNeeded); value
/// End Sub
///
/// Public Sub PrintIndented(level As Integer, text As String)
///     Debug.Print Spc(level * 4); text
/// End Sub
///
/// Public Sub PrintArray(arr() As Variant, Optional itemsPerLine As Integer = 5)
///     Dim i As Integer
///     Dim count As Integer
///     count = 0
///     
///     For i = LBound(arr) To UBound(arr)
///         Debug.Print arr(i); Spc(10);
///         count = count + 1
///         If count >= itemsPerLine Then
///             Debug.Print  ' New line
///             count = 0
///         End If
///     Next i
///     
///     If count > 0 Then Debug.Print  ' Final new line if needed
/// End Sub
/// ```
///
/// ## Error Handling
///
/// The Spc function itself doesn't typically generate errors, but the Print statement it's used with can:
///
/// - **Error 52** (Bad file name or number): If file number is invalid
/// - **Error 54** (Bad file mode): If file not opened for output
/// - **Error 13** (Type mismatch): If n is not numeric
///
/// Always ensure file is properly opened:
/// ```vb
/// On Error Resume Next
/// Print #fileNum, "Data"; Spc(10); "Value"
/// If Err.Number <> 0 Then
///     MsgBox "Error writing to file: " & Err.Description
/// End If
/// ```
///
/// ## Performance Considerations
///
/// - Spc is very efficient for positioning output
/// - More efficient than concatenating `Space()` strings in Print statements
/// - No performance difference between Spc and Space within Print statements
/// - File I/O is the bottleneck, not Spc itself
///
/// ## Best Practices
///
/// 1. **Only in Print**: Use Spc only within Print # or Debug.Print statements
/// 2. **Validate Arguments**: Ensure n is positive and reasonable
/// 3. **Consistent Spacing**: Use constants for column widths
/// 4. **Calculate Dynamically**: Adjust spacing based on content length
/// 5. **Use Space Alternative**: Use `Space()` function if need string result
/// 6. **Combine with Tab**: Use Tab for absolute positioning, Spc for relative
/// 7. **Test Output**: Verify alignment with actual data
/// 8. **Monospace Fonts**: Ensure output viewed in monospace font
/// 9. **Handle Long Data**: Account for data that exceeds expected width
/// 10. **Document Format**: Comment expected column layout
///
/// ## Comparison with Related Functions
///
/// | Function | Usage Context | Positioning | Returns |
/// |----------|--------------|-------------|---------|
/// | Spc(n) | Print statements only | Relative (+n spaces) | Variant (internal) |
/// | Tab(n) | Print statements only | Absolute (column n) | Variant (internal) |
/// | Space(n) | Anywhere | N/A | String of n spaces |
/// | String(n, " ") | Anywhere | N/A | String of n spaces |
///
/// ## Platform Considerations
///
/// - Available in VB6, VBA (all versions)
/// - Part of Print statement syntax
/// - Behavior consistent across platforms
/// - Works with Debug.Print, Print # (files), and Printer.Print
/// - Output width depends on file width setting (default 80 characters)
///
/// ## Limitations
///
/// - Cannot be used outside Print statements
/// - Cannot assign Spc result to variable
/// - Cannot use in string concatenation
/// - Wrapping behavior depends on output width setting
/// - Not suitable for proportional fonts (use with monospace)
/// - Limited to text output scenarios
///
/// ## Related Functions
///
/// - `Tab`: Positions output at absolute column position
/// - `Space`: Returns string of spaces (can be used anywhere)
/// - `Print`: Statement that outputs data
/// - `Width`: Statement that sets output line width
///
#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn spc_basic() {
        let source = r#"
Sub Test()
    Debug.Print "A"; Spc(5); "B"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        StringLiteral ("\"A\""),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("5"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"B\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_with_variable() {
        let source = r#"
Sub Test()
    Dim spaces As Integer
    spaces = 10
    Print #1, "Data"; Spc(spaces); "Value"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("spaces"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("spaces"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        Newline,
                    },
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Data\""),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        Identifier ("spaces"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"Value\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_file_output() {
        let source = r#"
Sub Test()
    Dim fileNum As Integer
    fileNum = FreeFile
    Open "test.txt" For Output As #fileNum
    Print #fileNum, "Name"; Spc(10); "Age"
    Close #fileNum
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("fileNum"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("fileNum"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("FreeFile"),
                        },
                        Newline,
                    },
                    OpenStatement {
                        Whitespace,
                        OpenKeyword,
                        Whitespace,
                        StringLiteral ("\"test.txt\""),
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        OutputKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        Octothorpe,
                        Identifier ("fileNum"),
                        Newline,
                    },
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        Identifier ("fileNum"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Name\""),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"Age\""),
                        Newline,
                    },
                    CloseStatement {
                        Whitespace,
                        CloseKeyword,
                        Whitespace,
                        Octothorpe,
                        Identifier ("fileNum"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Label:"; Spc(15); "Value"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        StringLiteral ("\"Label:\""),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("15"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"Value\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_multiple_calls() {
        let source = r#"
Sub Test()
    Print #1, "A"; Spc(3); "B"; Spc(5); "C"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"A\""),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("3"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"B\""),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("5"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"C\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_in_loop() {
        let source = r"
Sub Test()
    Dim i As Integer
    For i = 1 To 10
        Debug.Print i; Spc(5); i * 10
    Next i
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    ForStatement {
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("i"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("1"),
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
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                Identifier ("i"),
                                Semicolon,
                                Whitespace,
                                Identifier ("Spc"),
                                LeftParenthesis,
                                IntegerLiteral ("5"),
                                RightParenthesis,
                                Semicolon,
                                Whitespace,
                                Identifier ("i"),
                                Whitespace,
                                MultiplicationOperator,
                                Whitespace,
                                IntegerLiteral ("10"),
                                Newline,
                            },
                            Whitespace,
                        },
                        NextKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_calculated_spacing() {
        let source = r#"
Sub Test()
    Dim name As String
    name = "John"
    Print #1, name; Spc(20 - Len(name)); "Age"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        NameKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    NameStatement {
                        Whitespace,
                        NameKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteral ("\"John\""),
                        Newline,
                    },
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        NameKeyword,
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("20"),
                        Whitespace,
                        SubtractionOperator,
                        Whitespace,
                        LenKeyword,
                        LeftParenthesis,
                        NameKeyword,
                        RightParenthesis,
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"Age\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_with_numbers() {
        let source = r"
Sub Test()
    Debug.Print 100; Spc(10); 200; Spc(10); 300
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        IntegerLiteral ("100"),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        IntegerLiteral ("200"),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        IntegerLiteral ("300"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_class_usage() {
        let source = r"
Class Reporter
    Public Sub WriteRow(f As Integer, a As String, b As String)
        Print #f, a; Spc(10); b
    End Sub
End Class
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            Unknown,
            Whitespace,
            CallStatement {
                Identifier ("Reporter"),
                Newline,
            },
            Whitespace,
            SubStatement {
                PublicKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("WriteRow"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("f"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("a"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("b"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        Identifier ("f"),
                        Comma,
                        Whitespace,
                        Identifier ("a"),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        Identifier ("b"),
                        Newline,
                    },
                    Whitespace,
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
            Unknown,
            Whitespace,
            Unknown,
            Newline,
        ]);
    }

    #[test]
    fn spc_if_statement() {
        let source = r#"
Sub Test()
    If condition Then
        Debug.Print "True"; Spc(5); "Yes"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("condition"),
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                StringLiteral ("\"True\""),
                                Semicolon,
                                Whitespace,
                                Identifier ("Spc"),
                                LeftParenthesis,
                                IntegerLiteral ("5"),
                                RightParenthesis,
                                Semicolon,
                                Whitespace,
                                StringLiteral ("\"Yes\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_elseif() {
        let source = r#"
Sub Test()
    If x = 1 Then
        Print #1, "One"; Spc(5); "1"
    ElseIf x = 2 Then
        Print #1, "Two"; Spc(5); "2"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("x"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            PrintStatement {
                                Whitespace,
                                PrintKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                StringLiteral ("\"One\""),
                                Semicolon,
                                Whitespace,
                                Identifier ("Spc"),
                                LeftParenthesis,
                                IntegerLiteral ("5"),
                                RightParenthesis,
                                Semicolon,
                                Whitespace,
                                StringLiteral ("\"1\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        ElseIfClause {
                            ElseIfKeyword,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("x"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("2"),
                                },
                            },
                            Whitespace,
                            ThenKeyword,
                            Newline,
                            StatementList {
                                PrintStatement {
                                    Whitespace,
                                    PrintKeyword,
                                    Whitespace,
                                    Octothorpe,
                                    IntegerLiteral ("1"),
                                    Comma,
                                    Whitespace,
                                    StringLiteral ("\"Two\""),
                                    Semicolon,
                                    Whitespace,
                                    Identifier ("Spc"),
                                    LeftParenthesis,
                                    IntegerLiteral ("5"),
                                    RightParenthesis,
                                    Semicolon,
                                    Whitespace,
                                    StringLiteral ("\"2\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_select_case() {
        let source = r#"
Sub Test()
    Select Case value
        Case 1
            Print #1, "One"; Spc(10); "Value"
        Case 2
            Print #1, "Two"; Spc(10); "Value"
    End Select
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("value"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("1"),
                            Newline,
                            StatementList {
                                PrintStatement {
                                    Whitespace,
                                    PrintKeyword,
                                    Whitespace,
                                    Octothorpe,
                                    IntegerLiteral ("1"),
                                    Comma,
                                    Whitespace,
                                    StringLiteral ("\"One\""),
                                    Semicolon,
                                    Whitespace,
                                    Identifier ("Spc"),
                                    LeftParenthesis,
                                    IntegerLiteral ("10"),
                                    RightParenthesis,
                                    Semicolon,
                                    Whitespace,
                                    StringLiteral ("\"Value\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("2"),
                            Newline,
                            StatementList {
                                PrintStatement {
                                    Whitespace,
                                    PrintKeyword,
                                    Whitespace,
                                    Octothorpe,
                                    IntegerLiteral ("1"),
                                    Comma,
                                    Whitespace,
                                    StringLiteral ("\"Two\""),
                                    Semicolon,
                                    Whitespace,
                                    Identifier ("Spc"),
                                    LeftParenthesis,
                                    IntegerLiteral ("10"),
                                    RightParenthesis,
                                    Semicolon,
                                    Whitespace,
                                    StringLiteral ("\"Value\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_do_while() {
        let source = r"
Sub Test()
    Do While i < 10
        Print #1, i; Spc(5); i * 2
        i = i + 1
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("i"),
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
                        Newline,
                        StatementList {
                            PrintStatement {
                                Whitespace,
                                PrintKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                Identifier ("i"),
                                Semicolon,
                                Whitespace,
                                Identifier ("Spc"),
                                LeftParenthesis,
                                IntegerLiteral ("5"),
                                RightParenthesis,
                                Semicolon,
                                Whitespace,
                                Identifier ("i"),
                                Whitespace,
                                MultiplicationOperator,
                                Whitespace,
                                IntegerLiteral ("2"),
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("i"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("i"),
                                    },
                                    Whitespace,
                                    AdditionOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        LoopKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_do_until() {
        let source = r#"
Sub Test()
    Do Until i >= 10
        Debug.Print "Item"; Spc(8); i
        i = i + 1
    Loop
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        UntilKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("i"),
                            },
                            Whitespace,
                            GreaterThanOrEqualOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                StringLiteral ("\"Item\""),
                                Semicolon,
                                Whitespace,
                                Identifier ("Spc"),
                                LeftParenthesis,
                                IntegerLiteral ("8"),
                                RightParenthesis,
                                Semicolon,
                                Whitespace,
                                Identifier ("i"),
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("i"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("i"),
                                    },
                                    Whitespace,
                                    AdditionOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        LoopKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_while_wend() {
        let source = r"
Sub Test()
    While count < 5
        Print #1, count; Spc(10); count * 10
        count = count + 1
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("count"),
                            },
                            Whitespace,
                            LessThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("5"),
                            },
                        },
                        Newline,
                        StatementList {
                            PrintStatement {
                                Whitespace,
                                PrintKeyword,
                                Whitespace,
                                Octothorpe,
                                IntegerLiteral ("1"),
                                Comma,
                                Whitespace,
                                Identifier ("count"),
                                Semicolon,
                                Whitespace,
                                Identifier ("Spc"),
                                LeftParenthesis,
                                IntegerLiteral ("10"),
                                RightParenthesis,
                                Semicolon,
                                Whitespace,
                                Identifier ("count"),
                                Whitespace,
                                MultiplicationOperator,
                                Whitespace,
                                IntegerLiteral ("10"),
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("count"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("count"),
                                    },
                                    Whitespace,
                                    AdditionOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_with_statement() {
        let source = r#"
Sub Test()
    With reporter
        Print #.FileNum, "Name"; Spc(10); "Value"
    End With
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("reporter"),
                        Newline,
                        StatementList {
                            PrintStatement {
                                Whitespace,
                                PrintKeyword,
                                Whitespace,
                                Octothorpe,
                                PeriodOperator,
                                Identifier ("FileNum"),
                                Comma,
                                Whitespace,
                                StringLiteral ("\"Name\""),
                                Semicolon,
                                Whitespace,
                                Identifier ("Spc"),
                                LeftParenthesis,
                                IntegerLiteral ("10"),
                                RightParenthesis,
                                Semicolon,
                                Whitespace,
                                StringLiteral ("\"Value\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_format_function() {
        let source = r#"
Sub Test()
    Dim amount As Double
    amount = 123.45
    Print #1, "Total:"; Spc(5); Format(amount, "$#,##0.00")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("amount"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("amount"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            SingleLiteral,
                        },
                        Newline,
                    },
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        DateLiteral ("#1, \"Total:\"; Spc(5); Format(amount, \"$#"),
                        Comma,
                        DateLiteral ("##"),
                        SingleLiteral,
                        StringLiteral ("\")"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_zero_spaces() {
        let source = r#"
Sub Test()
    Debug.Print "A"; Spc(0); "B"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        StringLiteral ("\"A\""),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("0"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"B\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_large_number() {
        let source = r#"
Sub Test()
    Print #1, "Start"; Spc(50); "End"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Start\""),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("50"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"End\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_function_call() {
        let source = r"
Sub Test()
    Debug.Print GetLabel(); Spc(10); GetValue()
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("GetLabel"),
                        LeftParenthesis,
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        Identifier ("GetValue"),
                        LeftParenthesis,
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_header_row() {
        let source = r#"
Sub Test()
    Print #1, "Name"; Spc(15); "Age"; Spc(10); "City"
    Print #1, String(15, "-"); Spc(1); String(10, "-"); Spc(1); String(15, "-")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Name\""),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("15"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"Age\""),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"City\""),
                        Newline,
                    },
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringKeyword,
                        LeftParenthesis,
                        IntegerLiteral ("15"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"-\""),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("1"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringKeyword,
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"-\""),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("1"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringKeyword,
                        LeftParenthesis,
                        IntegerLiteral ("15"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"-\""),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_nested_function() {
        let source = r#"
Sub Test()
    Dim width As Integer
    width = 20
    Debug.Print "Item"; Spc(width - Len("Item")); "Value"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    WidthStatement {
                        Whitespace,
                        WidthKeyword,
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IntegerLiteral ("20"),
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        StringLiteral ("\"Item\""),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        WidthKeyword,
                        Whitespace,
                        SubtractionOperator,
                        Whitespace,
                        LenKeyword,
                        LeftParenthesis,
                        StringLiteral ("\"Item\""),
                        RightParenthesis,
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"Value\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Print #fileNum, "Data"; Spc(10); "Value"
    If Err.Number <> 0 Then
        MsgBox "Error"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    OnErrorStatement {
                        Whitespace,
                        OnKeyword,
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        ResumeKeyword,
                        Whitespace,
                        NextKeyword,
                        Newline,
                    },
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        Identifier ("fileNum"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Data\""),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"Value\""),
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            MemberAccessExpression {
                                Identifier ("Err"),
                                PeriodOperator,
                                Identifier ("Number"),
                            },
                            Whitespace,
                            InequalityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
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
                                StringLiteral ("\"Error\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_on_error_goto() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Print #1, "Label"; Spc(spacing); "Value"
    Exit Sub
ErrorHandler:
    MsgBox "Error printing"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    OnErrorStatement {
                        Whitespace,
                        OnKeyword,
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        GotoKeyword,
                        Whitespace,
                        Identifier ("ErrorHandler"),
                        Newline,
                    },
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Label\""),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        Identifier ("spacing"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"Value\""),
                        Newline,
                    },
                    ExitStatement {
                        Whitespace,
                        ExitKeyword,
                        Whitespace,
                        SubKeyword,
                        Newline,
                    },
                    LabelStatement {
                        Identifier ("ErrorHandler"),
                        ColonOperator,
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Error printing\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_indentation() {
        let source = r#"
Sub Test()
    Dim level As Integer
    level = 3
    Print #1, Spc(level * 4); "Indented text"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("level"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("level"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("3"),
                        },
                        Newline,
                    },
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        Identifier ("level"),
                        Whitespace,
                        MultiplicationOperator,
                        Whitespace,
                        IntegerLiteral ("4"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"Indented text\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_report_formatting() {
        let source = r#"
Sub Test()
    Dim lineNum As Integer
    Dim desc As String
    lineNum = 1
    desc = "Item description"
    Print #1, Format(lineNum, "000"); Spc(5); desc; Spc(30 - Len(desc)); "$100.00"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("lineNum"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("desc"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("lineNum"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("1"),
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("desc"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"Item description\""),
                        },
                        Newline,
                    },
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("Format"),
                        LeftParenthesis,
                        Identifier ("lineNum"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"000\""),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("5"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        Identifier ("desc"),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("30"),
                        Whitespace,
                        SubtractionOperator,
                        Whitespace,
                        LenKeyword,
                        LeftParenthesis,
                        Identifier ("desc"),
                        RightParenthesis,
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"$100.00\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn spc_log_entry() {
        let source = r#"
Sub Test()
    Dim timestamp As String
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #1, timestamp; Spc(5); "INFO"; Spc(10); "Application started"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("timestamp"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("timestamp"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Format"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("Now"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"yyyy-mm-dd hh:nn:ss\""),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("timestamp"),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("5"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"INFO\""),
                        Semicolon,
                        Whitespace,
                        Identifier ("Spc"),
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Semicolon,
                        Whitespace,
                        StringLiteral ("\"Application started\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }
}
