/// # Space Function
///
/// Returns a String consisting of the specified number of spaces.
///
/// ## Syntax
///
/// ```vb
/// Space(number)
/// ```
///
/// ## Parameters
///
/// - `number` - Required. Long or any valid numeric expression specifying the number of spaces to return.
///
/// ## Return Value
///
/// Returns a String containing the specified number of space characters (ASCII 32).
///
/// ## Remarks
///
/// The Space function is useful for creating strings with a specific number of spaces, commonly used for:
/// - Formatting output in fixed-width columns
/// - Creating indentation in text
/// - Padding strings to specific lengths
/// - Aligning text in reports or displays
/// - Creating blank lines or spacing in file output
///
/// Key characteristics:
/// - Returns a string of space characters (ASCII value 32)
/// - If `number` is 0, returns an empty string ("")
/// - If `number` is negative, generates Error 5 (Invalid procedure call or argument)
/// - Non-integer values are rounded to the nearest integer
/// - Maximum practical limit is system memory for string storage
///
/// The Space function is related to other string generation functions:
/// - **Space(n)**: Creates n space characters
/// - **String(n, character)**: Creates n repetitions of any character
/// - **String(n, charcode)**: Creates n repetitions of character with given ASCII code
///
/// ## Typical Uses
///
/// 1. **Column Formatting**: Align text in fixed-width columns
/// 2. **Indentation**: Create indented text structures
/// 3. **Padding**: Pad strings to specific widths
/// 4. **Separation**: Add spacing between elements
/// 5. **Report Generation**: Format reports with proper alignment
/// 6. **Text Files**: Create formatted text file output
/// 7. **Display Alignment**: Align data in list boxes or text boxes
/// 8. **Table Creation**: Build ASCII tables with proper spacing
///
/// ## Basic Examples
///
/// ```vb
/// ' Example 1: Create a string of 5 spaces
/// Dim spaces As String
/// spaces = Space(5)
/// ' Returns "     " (5 spaces)
/// ```
///
/// ```vb
/// ' Example 2: Pad a string to 20 characters
/// Dim name As String
/// Dim paddedName As String
/// name = "John"
/// paddedName = name & Space(20 - Len(name))
/// ' Returns "John                " (16 trailing spaces)
/// ```
///
/// ```vb
/// ' Example 3: Create indented text
/// Dim level As Integer
/// Dim text As String
/// level = 3
/// text = Space(level * 4) & "Indented text"
/// ' Returns "            Indented text" (12 spaces for 3 levels of 4-space indent)
/// ```
///
/// ```vb
/// ' Example 4: Format columns in output
/// Dim item As String
/// Dim price As String
/// item = "Apple"
/// price = "$1.99"
/// Debug.Print item & Space(20 - Len(item)) & price
/// ' Outputs: "Apple               $1.99"
/// ```
///
/// ## Common Patterns
///
/// ### Pattern 1: `PadRight`
/// Pad string to specified width (right padding)
/// ```vb
/// Function PadRight(text As String, totalWidth As Integer) As String
///     Dim currentLen As Integer
///     currentLen = Len(text)
///     
///     If currentLen >= totalWidth Then
///         PadRight = text
///     Else
///         PadRight = text & Space(totalWidth - currentLen)
///     End If
/// End Function
/// ```
///
/// ### Pattern 2: `PadLeft`
/// Pad string to specified width (left padding)
/// ```vb
/// Function PadLeft(text As String, totalWidth As Integer) As String
///     Dim currentLen As Integer
///     currentLen = Len(text)
///     
///     If currentLen >= totalWidth Then
///         PadLeft = text
///     Else
///         PadLeft = Space(totalWidth - currentLen) & text
///     End If
/// End Function
/// ```
///
/// ### Pattern 3: Center
/// Center text within specified width
/// ```vb
/// Function Center(text As String, totalWidth As Integer) As String
///     Dim currentLen As Integer
///     Dim leftPadding As Integer
///     Dim rightPadding As Integer
///     
///     currentLen = Len(text)
///     
///     If currentLen >= totalWidth Then
///         Center = text
///         Exit Function
///     End If
///     
///     leftPadding = (totalWidth - currentLen) \ 2
///     rightPadding = totalWidth - currentLen - leftPadding
///     
///     Center = Space(leftPadding) & text & Space(rightPadding)
/// End Function
/// ```
///
/// ### Pattern 4: `CreateIndent`
/// Create indentation for nested structures
/// ```vb
/// Function CreateIndent(level As Integer, Optional spacesPerLevel As Integer = 4) As String
///     If level <= 0 Then
///         CreateIndent = ""
///     Else
///         CreateIndent = Space(level * spacesPerLevel)
///     End If
/// End Function
/// ```
///
/// ### Pattern 5: `FormatColumn`
/// Format text in fixed-width column
/// ```vb
/// Function FormatColumn(text As String, width As Integer, _
///                       Optional alignment As String = "LEFT") As String
///     Select Case UCase(alignment)
///         Case "LEFT"
///             FormatColumn = PadRight(text, width)
///         Case "RIGHT"
///             FormatColumn = PadLeft(text, width)
///         Case "CENTER"
///             FormatColumn = Center(text, width)
///         Case Else
///             FormatColumn = text
///     End Select
/// End Function
/// ```
///
/// ### Pattern 6: `CreateSeparator`
/// Create separator line with spaces
/// ```vb
/// Function CreateSeparator(leftText As String, rightText As String, _
///                          totalWidth As Integer, _
///                          Optional separator As String = " ") As String
///     Dim leftLen As Integer
///     Dim rightLen As Integer
///     Dim middleSpaces As Integer
///     
///     leftLen = Len(leftText)
///     rightLen = Len(rightText)
///     middleSpaces = totalWidth - leftLen - rightLen
///     
///     If middleSpaces < 0 Then middleSpaces = 0
///     
///     CreateSeparator = leftText & String(middleSpaces, separator) & rightText
/// End Function
/// ```
///
/// ### Pattern 7: `BuildTableRow`
/// Build formatted table row
/// ```vb
/// Function BuildTableRow(columns() As String, widths() As Integer) As String
///     Dim row As String
///     Dim i As Integer
///     
///     row = ""
///     For i = LBound(columns) To UBound(columns)
///         If i > LBound(columns) Then row = row & " | "
///         row = row & PadRight(columns(i), widths(i))
///     Next i
///     
///     BuildTableRow = row
/// End Function
/// ```
///
/// ### Pattern 8: `IndentMultiline`
/// Indent all lines in multiline text
/// ```vb
/// Function IndentMultiline(text As String, spaces As Integer) As String
///     Dim lines() As String
///     Dim result As String
///     Dim i As Integer
///     Dim indent As String
///     
///     lines = Split(text, vbCrLf)
///     indent = Space(spaces)
///     result = ""
///     
///     For i = LBound(lines) To UBound(lines)
///         If i > LBound(lines) Then result = result & vbCrLf
///         result = result & indent & lines(i)
///     Next i
///     
///     IndentMultiline = result
/// End Function
/// ```
///
/// ### Pattern 9: `CreateBlankLine`
/// Create blank line with specific spacing
/// ```vb
/// Function CreateBlankLine(width As Integer) As String
///     CreateBlankLine = Space(width)
/// End Function
/// ```
///
/// ### Pattern 10: `AlignNumber`
/// Right-align number in field
/// ```vb
/// Function AlignNumber(value As Variant, width As Integer, _
///                      Optional decimals As Integer = 2) As String
///     Dim formatted As String
///     
///     formatted = Format(value, "0." & String(decimals, "0"))
///     AlignNumber = PadLeft(formatted, width)
/// End Function
/// ```
///
/// ## Advanced Usage
///
/// ### Example 1: `TableFormatter` Class
/// Format data in ASCII tables
/// ```vb
/// ' Class: TableFormatter
/// Private m_columnWidths() As Integer
/// Private m_columnCount As Integer
/// Private m_alignment() As String
///
/// Private Sub Class_Initialize()
///     m_columnCount = 0
/// End Sub
///
/// Public Sub SetColumns(widths() As Integer, Optional alignments As Variant)
///     Dim i As Integer
///     
///     m_columnCount = UBound(widths) - LBound(widths) + 1
///     ReDim m_columnWidths(LBound(widths) To UBound(widths))
///     ReDim m_alignment(LBound(widths) To UBound(widths))
///     
///     For i = LBound(widths) To UBound(widths)
///         m_columnWidths(i) = widths(i)
///         If IsMissing(alignments) Then
///             m_alignment(i) = "LEFT"
///         Else
///             m_alignment(i) = alignments(i)
///         End If
///     Next i
/// End Sub
///
/// Public Function FormatRow(values() As String) As String
///     Dim row As String
///     Dim i As Integer
///     Dim formattedValue As String
///     
///     row = ""
///     For i = LBound(values) To UBound(values)
///         If i > LBound(values) Then row = row & " | "
///         
///         formattedValue = FormatCell(values(i), m_columnWidths(i), m_alignment(i))
///         row = row & formattedValue
///     Next i
///     
///     FormatRow = row
/// End Function
///
/// Private Function FormatCell(value As String, width As Integer, _
///                             alignment As String) As String
///     Dim currentLen As Integer
///     Dim padding As Integer
///     
///     currentLen = Len(value)
///     
///     If currentLen >= width Then
///         FormatCell = Left(value, width)
///         Exit Function
///     End If
///     
///     padding = width - currentLen
///     
///     Select Case UCase(alignment)
///         Case "LEFT"
///             FormatCell = value & Space(padding)
///         Case "RIGHT"
///             FormatCell = Space(padding) & value
///         Case "CENTER"
///             Dim leftPad As Integer
///             Dim rightPad As Integer
///             leftPad = padding \ 2
///             rightPad = padding - leftPad
///             FormatCell = Space(leftPad) & value & Space(rightPad)
///         Case Else
///             FormatCell = value & Space(padding)
///     End Select
/// End Function
///
/// Public Function CreateHeader(headers() As String) As String
///     Dim header As String
///     Dim separator As String
///     Dim i As Integer
///     
///     header = FormatRow(headers)
///     separator = ""
///     
///     For i = LBound(m_columnWidths) To UBound(m_columnWidths)
///         If i > LBound(m_columnWidths) Then separator = separator & "-+-"
///         separator = separator & String(m_columnWidths(i), "-")
///     Next i
///     
///     CreateHeader = header & vbCrLf & separator
/// End Function
///
/// Public Function GetTotalWidth() As Integer
///     Dim total As Integer
///     Dim i As Integer
///     
///     total = 0
///     For i = LBound(m_columnWidths) To UBound(m_columnWidths)
///         total = total + m_columnWidths(i)
///     Next i
///     
///     ' Add separator widths
///     total = total + (m_columnCount - 1) * 3  ' " | " between columns
///     
///     GetTotalWidth = total
/// End Function
/// ```
///
/// ### Example 2: `ReportGenerator` Module
/// Generate formatted text reports
/// ```vb
/// ' Module: ReportGenerator
///
/// Public Function GenerateReport(title As String, data() As Variant, _
///                                columnHeaders() As String, _
///                                columnWidths() As Integer) As String
///     Dim report As String
///     Dim i As Integer
///     Dim j As Integer
///     Dim totalWidth As Integer
///     Dim titleLine As String
///     
///     ' Calculate total width
///     totalWidth = 0
///     For i = LBound(columnWidths) To UBound(columnWidths)
///         totalWidth = totalWidth + columnWidths(i)
///     Next i
///     totalWidth = totalWidth + (UBound(columnWidths) - LBound(columnWidths)) * 3
///     
///     ' Center title
///     titleLine = CenterText(title, totalWidth)
///     report = titleLine & vbCrLf
///     report = report & String(totalWidth, "=") & vbCrLf
///     
///     ' Add header
///     For i = LBound(columnHeaders) To UBound(columnHeaders)
///         If i > LBound(columnHeaders) Then report = report & " | "
///         report = report & PadRight(columnHeaders(i), columnWidths(i))
///     Next i
///     report = report & vbCrLf
///     
///     ' Add separator
///     For i = LBound(columnWidths) To UBound(columnWidths)
///         If i > LBound(columnWidths) Then report = report & "-+-"
///         report = report & String(columnWidths(i), "-")
///     Next i
///     report = report & vbCrLf
///     
///     ' Add data rows
///     For i = LBound(data, 1) To UBound(data, 1)
///         For j = LBound(data, 2) To UBound(data, 2)
///             If j > LBound(data, 2) Then report = report & " | "
///             report = report & PadRight(CStr(data(i, j)), columnWidths(j))
///         Next j
///         report = report & vbCrLf
///     Next i
///     
///     GenerateReport = report
/// End Function
///
/// Private Function CenterText(text As String, width As Integer) As String
///     Dim textLen As Integer
///     Dim leftPad As Integer
///     Dim rightPad As Integer
///     
///     textLen = Len(text)
///     
///     If textLen >= width Then
///         CenterText = text
///         Exit Function
///     End If
///     
///     leftPad = (width - textLen) \ 2
///     rightPad = width - textLen - leftPad
///     
///     CenterText = Space(leftPad) & text & Space(rightPad)
/// End Function
///
/// Private Function PadRight(text As String, width As Integer) As String
///     If Len(text) >= width Then
///         PadRight = Left(text, width)
///     Else
///         PadRight = text & Space(width - Len(text))
///     End If
/// End Function
///
/// Public Function CreateSummaryLine(label As String, value As String, _
///                                   totalWidth As Integer) As String
///     Dim labelLen As Integer
///     Dim valueLen As Integer
///     Dim spacesNeeded As Integer
///     
///     labelLen = Len(label)
///     valueLen = Len(value)
///     spacesNeeded = totalWidth - labelLen - valueLen
///     
///     If spacesNeeded < 1 Then spacesNeeded = 1
///     
///     CreateSummaryLine = label & Space(spacesNeeded) & value
/// End Function
/// ```
///
/// ### Example 3: `CodeFormatter` Class
/// Format source code with indentation
/// ```vb
/// ' Class: CodeFormatter
/// Private m_indentLevel As Integer
/// Private m_spacesPerIndent As Integer
/// Private m_output As String
///
/// Private Sub Class_Initialize()
///     m_indentLevel = 0
///     m_spacesPerIndent = 4
///     m_output = ""
/// End Sub
///
/// Public Property Let SpacesPerIndent(value As Integer)
///     If value > 0 Then m_spacesPerIndent = value
/// End Property
///
/// Public Sub IncreaseIndent()
///     m_indentLevel = m_indentLevel + 1
/// End Sub
///
/// Public Sub DecreaseIndent()
///     If m_indentLevel > 0 Then
///         m_indentLevel = m_indentLevel - 1
///     End If
/// End Sub
///
/// Public Sub AddLine(text As String)
///     Dim indent As String
///     indent = Space(m_indentLevel * m_spacesPerIndent)
///     
///     If m_output <> "" Then m_output = m_output & vbCrLf
///     m_output = m_output & indent & text
/// End Sub
///
/// Public Sub AddBlankLine()
///     If m_output <> "" Then m_output = m_output & vbCrLf
/// End Sub
///
/// Public Function GetOutput() As String
///     GetOutput = m_output
/// End Function
///
/// Public Sub Clear()
///     m_output = ""
///     m_indentLevel = 0
/// End Sub
///
/// Public Sub AddBlock(blockStart As String, blockEnd As String, _
///                     lines() As String)
///     Dim i As Integer
///     
///     AddLine blockStart
///     IncreaseIndent
///     
///     For i = LBound(lines) To UBound(lines)
///         AddLine lines(i)
///     Next i
///     
///     DecreaseIndent
///     AddLine blockEnd
/// End Sub
/// ```
///
/// ### Example 4: `ListBoxFormatter` Module
/// Format items for list box display
/// ```vb
/// ' Module: ListBoxFormatter
///
/// Public Function FormatListItem(item As String, value As String, _
///                                totalWidth As Integer, _
///                                Optional separator As String = " ") As String
///     Dim itemLen As Integer
///     Dim valueLen As Integer
///     Dim separatorLen As Integer
///     Dim spacesNeeded As Integer
///     
///     itemLen = Len(item)
///     valueLen = Len(value)
///     separatorLen = Len(separator)
///     
///     spacesNeeded = totalWidth - itemLen - valueLen
///     
///     If spacesNeeded < 1 Then spacesNeeded = 1
///     
///     FormatListItem = item & Space(spacesNeeded) & value
/// End Function
///
/// Public Sub PopulateFormattedList(lst As ListBox, items() As String, _
///                                  values() As String, width As Integer)
///     Dim i As Integer
///     
///     lst.Clear
///     
///     For i = LBound(items) To UBound(items)
///         lst.AddItem FormatListItem(items(i), values(i), width)
///     Next i
/// End Sub
///
/// Public Function CreateTreeItem(text As String, level As Integer, _
///                                Optional expandSymbol As String = "+") As String
///     Dim indent As String
///     
///     indent = Space(level * 2)
///     
///     If level > 0 Then
///         CreateTreeItem = indent & expandSymbol & " " & text
///     Else
///         CreateTreeItem = text
///     End If
/// End Function
///
/// Public Function AlignCurrency(amount As Double, width As Integer) As String
///     Dim formatted As String
///     
///     formatted = FormatCurrency(amount, 2)
///     AlignCurrency = Space(width - Len(formatted)) & formatted
/// End Function
/// ```
///
/// ## Error Handling
///
/// The Space function can generate the following errors:
///
/// - **Error 5** (Invalid procedure call or argument): If number is negative
/// - **Error 6** (Overflow): If number exceeds Long range
/// - **Error 7** (Out of memory): If resulting string exceeds available memory
/// - **Error 13** (Type mismatch): If number is not numeric
///
/// Always validate inputs:
/// ```vb
/// On Error Resume Next
/// result = Space(count)
/// If Err.Number <> 0 Then
///     MsgBox "Error creating spaces: " & Err.Description
/// End If
/// ```
///
/// ## Performance Considerations
///
/// - Very fast for small to moderate space counts (< 1000)
/// - For large space counts, consider if you really need that many
/// - String concatenation in loops can be slow; build once when possible
/// - Space function is more efficient than repeated string concatenation
/// - Consider caching commonly used space strings
///
/// ## Best Practices
///
/// 1. **Validate Count**: Ensure space count is non-negative
/// 2. **Use Constants**: Define column widths as constants for consistency
/// 3. **Avoid Magic Numbers**: Use named constants instead of literal numbers
/// 4. **Handle Edge Cases**: Check for zero or negative values
/// 5. **Consider Alternatives**: For very large strings, evaluate necessity
/// 6. **Cache Results**: Store frequently used space strings
/// 7. **Document Width**: Comment expected column widths in code
/// 8. **Test Alignment**: Verify output with different data lengths
/// 9. **Use Monospace**: Ensure font is monospace for proper alignment
/// 10. **Combine with Format**: Use with Format function for numeric alignment
///
/// ## Comparison with Related Functions
///
/// | Function | Purpose | Example | Result |
/// |----------|---------|---------|--------|
/// | Space(n) | n spaces | Space(5) | "     " |
/// | String(n, " ") | n of any character | String(5, " ") | "     " |
/// | String(n, 32) | n of ASCII char | String(5, 32) | "     " |
/// | String(n, "*") | n asterisks | String(5, "*") | "*****" |
///
/// ## Platform Considerations
///
/// - Available in VB6, VBA (all versions)
/// - Part of core string functions
/// - Consistent behavior across platforms
/// - Subject to system memory limits
/// - Maximum string length: approximately 2 billion characters (limited by available memory)
///
/// ## Limitations
///
/// - Cannot create negative number of spaces (generates error)
/// - Limited by available system memory
/// - Non-integer values are rounded (e.g., Space(3.7) = Space(4))
/// - Returns empty string for Space(0)
/// - Not suitable for creating non-breaking spaces (use Chr(160) for HTML/Unicode)
///
/// ## Related Functions
///
/// - `String`: Creates a string of repeated characters
/// - `SPC`: Positions output in Print # statements
/// - `Tab`: Positions output at specific column in Print # statements
/// - `LSet`: Left-aligns string within string variable
/// - `RSet`: Right-aligns string within string variable
///
#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn space_basic() {
        let source = r"
Sub Test()
    Dim spaces As String
    spaces = Space(5)
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
                        Identifier ("spaces"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
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
                        CallExpression {
                            Identifier ("Space"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("5"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
    fn space_with_variable() {
        let source = r"
Sub Test()
    Dim count As Integer
    Dim result As String
    count = 10
    result = Space(count)
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
                        Identifier ("count"),
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
                        Identifier ("result"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
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
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Space"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("count"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
    fn space_if_statement() {
        let source = r#"
Sub Test()
    If Len(Space(10)) = 10 Then
        MsgBox "Correct"
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
                            CallExpression {
                                LenKeyword,
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        CallExpression {
                                            Identifier ("Space"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    NumericLiteralExpression {
                                                        IntegerLiteral ("10"),
                                                    },
                                                },
                                            },
                                            RightParenthesis,
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10"),
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
                                StringLiteral ("\"Correct\""),
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
    fn space_function_return() {
        let source = r"
Function CreatePadding(n As Integer) As String
    CreatePadding = Space(n)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CreatePadding"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("n"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    IntegerKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("CreatePadding"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Space"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("n"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn space_variable_assignment() {
        let source = r"
Sub Test()
    Dim padding As String
    padding = Space(20)
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
                        Identifier ("padding"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("padding"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Space"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("20"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
    fn space_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Text" & Space(5) & "More"
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
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Text\""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Space"),
                        LeftParenthesis,
                        IntegerLiteral ("5"),
                        RightParenthesis,
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteral ("\"More\""),
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
    fn space_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Name:" & Space(10) & "Value"
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
                        StringLiteral ("\"Name:\""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Space"),
                        LeftParenthesis,
                        IntegerLiteral ("10"),
                        RightParenthesis,
                        Whitespace,
                        Ampersand,
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
    fn space_select_case() {
        let source = r#"
Sub Test()
    Select Case Len(Space(n))
        Case 5
            MsgBox "Five"
        Case 10
            MsgBox "Ten"
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
                        CallExpression {
                            LenKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Space"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("n"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("5"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("MsgBox"),
                                    Whitespace,
                                    StringLiteral ("\"Five\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IntegerLiteral ("10"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("MsgBox"),
                                    Whitespace,
                                    StringLiteral ("\"Ten\""),
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
    fn space_class_usage() {
        let source = r"
Class Formatter
    Public Function Pad(s As String, width As Integer) As String
        Pad = s & Space(width - Len(s))
    End Function
End Class
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            Unknown,
            Whitespace,
            CallStatement {
                Identifier ("Formatter"),
                Newline,
            },
            Whitespace,
            FunctionStatement {
                PublicKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("Pad"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("s"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    Comma,
                    Whitespace,
                },
                WidthKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("Pad"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("s"),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Space"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                WidthKeyword,
                                            },
                                            Whitespace,
                                            SubtractionOperator,
                                            Whitespace,
                                            CallExpression {
                                                LenKeyword,
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        IdentifierExpression {
                                                            Identifier ("s"),
                                                        },
                                                    },
                                                },
                                                RightParenthesis,
                                            },
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Newline,
                    },
                    Whitespace,
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
            Unknown,
            Whitespace,
            Unknown,
            Newline,
        ]);
    }

    #[test]
    fn space_with_statement() {
        let source = r#"
Sub Test()
    With txtOutput
        .Text = "Data" & Space(10) & "Value"
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
                        Identifier ("txtOutput"),
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    PeriodOperator,
                                },
                                BinaryExpression {
                                    IdentifierExpression {
                                        TextKeyword,
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    BinaryExpression {
                                        BinaryExpression {
                                            StringLiteralExpression {
                                                StringLiteral ("\"Data\""),
                                            },
                                            Whitespace,
                                            Ampersand,
                                            Whitespace,
                                            CallExpression {
                                                Identifier ("Space"),
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        NumericLiteralExpression {
                                                            IntegerLiteral ("10"),
                                                        },
                                                    },
                                                },
                                                RightParenthesis,
                                            },
                                        },
                                        Whitespace,
                                        Ampersand,
                                        Whitespace,
                                        StringLiteralExpression {
                                            StringLiteral ("\"Value\""),
                                        },
                                    },
                                },
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
    fn space_elseif() {
        let source = r"
Sub Test()
    Dim s As String
    If n = 5 Then
        s = Space(5)
    ElseIf n = 10 Then
        s = Space(10)
    Else
        s = Space(20)
    End If
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
                        Identifier ("s"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("n"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("5"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("s"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Space"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("5"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        ElseIfClause {
                            ElseIfKeyword,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("n"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("10"),
                                },
                            },
                            Whitespace,
                            ThenKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("s"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Space"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("10"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        ElseClause {
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("s"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Space"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("20"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
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
    fn space_for_loop() {
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 1 To 5
        Debug.Print Space(i) & "Line"
    Next i
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
                            IntegerLiteral ("5"),
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                Identifier ("Space"),
                                LeftParenthesis,
                                Identifier ("i"),
                                RightParenthesis,
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                StringLiteral ("\"Line\""),
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
    fn space_do_while() {
        let source = r"
Sub Test()
    Dim indent As String
    Do While level > 0
        indent = Space(level * 4)
        level = level - 1
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("indent"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("level"),
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("indent"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Space"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            BinaryExpression {
                                                IdentifierExpression {
                                                    Identifier ("level"),
                                                },
                                                Whitespace,
                                                MultiplicationOperator,
                                                Whitespace,
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("4"),
                                                },
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
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
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("level"),
                                    },
                                    Whitespace,
                                    SubtractionOperator,
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
    fn space_do_until() {
        let source = r"
Sub Test()
    Do Until width >= 50
        padding = padding & Space(10)
        width = width + 10
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
                        UntilKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                WidthKeyword,
                            },
                            Whitespace,
                            GreaterThanOrEqualOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("50"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("padding"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("padding"),
                                    },
                                    Whitespace,
                                    Ampersand,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("Space"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("10"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                Newline,
                            },
                            WidthStatement {
                                Whitespace,
                                WidthKeyword,
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                WidthKeyword,
                                Whitespace,
                                AdditionOperator,
                                Whitespace,
                                IntegerLiteral ("10"),
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
    fn space_while_wend() {
        let source = r"
Sub Test()
    While count < 100
        line = line & Space(5)
        count = count + 5
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
                                IntegerLiteral ("100"),
                            },
                        },
                        Newline,
                        StatementList {
                            LineInputStatement {
                                Whitespace,
                                LineKeyword,
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                LineKeyword,
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                Identifier ("Space"),
                                LeftParenthesis,
                                IntegerLiteral ("5"),
                                RightParenthesis,
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
                                        IntegerLiteral ("5"),
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
    fn space_parentheses() {
        let source = r"
Sub Test()
    Dim formatted As String
    formatted = (name & Space(20 - Len(name)))
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
                        Identifier ("formatted"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("formatted"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                IdentifierExpression {
                                    NameKeyword,
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Space"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            BinaryExpression {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("20"),
                                                },
                                                Whitespace,
                                                SubtractionOperator,
                                                Whitespace,
                                                CallExpression {
                                                    LenKeyword,
                                                    LeftParenthesis,
                                                    ArgumentList {
                                                        Argument {
                                                            IdentifierExpression {
                                                                NameKeyword,
                                                            },
                                                        },
                                                    },
                                                    RightParenthesis,
                                                },
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                            RightParenthesis,
                        },
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
    fn space_iif() {
        let source = r#"
Sub Test()
    Dim padding As String
    padding = IIf(needPadding, Space(10), "")
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
                        Identifier ("padding"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("padding"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("IIf"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("needPadding"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    CallExpression {
                                        Identifier ("Space"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("10"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"\""),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
    fn space_array_assignment() {
        let source = r#"
Sub Test()
    Dim lines(10) As String
    lines(0) = Space(5) & "Header"
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
                        Identifier ("lines"),
                        LeftParenthesis,
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        CallExpression {
                            Identifier ("lines"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("0"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Space"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("5"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"Header\""),
                            },
                        },
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
    fn space_property_assignment() {
        let source = r"
Class TextFormatter
    Public Indent As String
End Class

Sub Test()
    Dim fmt As New TextFormatter
    fmt.Indent = Space(4)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            Unknown,
            Whitespace,
            CallStatement {
                Identifier ("TextFormatter"),
                Newline,
            },
            Whitespace,
            DimStatement {
                PublicKeyword,
                Whitespace,
                Identifier ("Indent"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            Unknown,
            Whitespace,
            Unknown,
            Newline,
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
                        Identifier ("fmt"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        NewKeyword,
                        Whitespace,
                        Identifier ("TextFormatter"),
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        MemberAccessExpression {
                            Identifier ("fmt"),
                            PeriodOperator,
                            Identifier ("Indent"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Space"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("4"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
    fn space_function_argument() {
        let source = r"
Sub ProcessText(s As String)
End Sub

Sub Test()
    ProcessText Space(15)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ProcessText"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("s"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
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
                        Identifier ("ProcessText"),
                        Whitespace,
                        Identifier ("Space"),
                        LeftParenthesis,
                        IntegerLiteral ("15"),
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
    fn space_concatenation() {
        let source = r#"
Sub Test()
    Dim header As String
    header = "Item" & Space(20) & "Price"
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
                        Identifier ("header"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("header"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                StringLiteralExpression {
                                    StringLiteral ("\"Item\""),
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Space"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            NumericLiteralExpression {
                                                IntegerLiteral ("20"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"Price\""),
                            },
                        },
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
    fn space_comparison() {
        let source = r"
Sub Test()
    Dim isCorrect As Boolean
    isCorrect = (Len(Space(10)) = 10)
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
                        Identifier ("isCorrect"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        BooleanKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("isCorrect"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                CallExpression {
                                    LenKeyword,
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            CallExpression {
                                                Identifier ("Space"),
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        NumericLiteralExpression {
                                                            IntegerLiteral ("10"),
                                                        },
                                                    },
                                                },
                                                RightParenthesis,
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("10"),
                                },
                            },
                            RightParenthesis,
                        },
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
    fn space_arithmetic() {
        let source = r"
Sub Test()
    Dim padding As String
    padding = Space(totalWidth - textWidth)
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
                        Identifier ("padding"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("padding"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Space"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("totalWidth"),
                                        },
                                        Whitespace,
                                        SubtractionOperator,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("textWidth"),
                                        },
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
    fn space_indentation() {
        let source = r#"
Sub Test()
    Dim indentedText As String
    indentedText = Space(level * 4) & "Code"
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
                        Identifier ("indentedText"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("indentedText"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("Space"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        BinaryExpression {
                                            IdentifierExpression {
                                                Identifier ("level"),
                                            },
                                            Whitespace,
                                            MultiplicationOperator,
                                            Whitespace,
                                            NumericLiteralExpression {
                                                IntegerLiteral ("4"),
                                            },
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"Code\""),
                            },
                        },
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
    fn space_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Dim s As String
    s = Space(count)
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("s"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("s"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Space"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("count"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
    fn space_on_error_goto() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Dim spacing As String
    spacing = Space(numberOfSpaces)
    Exit Sub
ErrorHandler:
    MsgBox "Error creating spaces"
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("spacing"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("spacing"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Space"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("numberOfSpaces"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
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
                        StringLiteral ("\"Error creating spaces\""),
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
    fn space_column_alignment() {
        let source = r#"
Sub Test()
    Dim name As String
    Dim value As String
    Dim aligned As String
    name = "Item"
    value = "$10.00"
    aligned = name & Space(30 - Len(name)) & value
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
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("value"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("aligned"),
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
                        StringLiteral ("\"Item\""),
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("value"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"$10.00\""),
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("aligned"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                IdentifierExpression {
                                    NameKeyword,
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Space"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            BinaryExpression {
                                                NumericLiteralExpression {
                                                    IntegerLiteral ("30"),
                                                },
                                                Whitespace,
                                                SubtractionOperator,
                                                Whitespace,
                                                CallExpression {
                                                    LenKeyword,
                                                    LeftParenthesis,
                                                    ArgumentList {
                                                        Argument {
                                                            IdentifierExpression {
                                                                NameKeyword,
                                                            },
                                                        },
                                                    },
                                                    RightParenthesis,
                                                },
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("value"),
                            },
                        },
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
