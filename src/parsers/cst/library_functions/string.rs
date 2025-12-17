//! VB6 `String` Function
//!
//! The `String` function returns a string consisting of a repeating character.
//!
//! ## Syntax
//! ```vb6
//! String(number, character)
//! ```
//!
//! ## Parameters
//! - `number`: Required. Long integer specifying the length of the returned string. Must be between 0 and approximately 2 billion (limited by available memory).
//! - `character`: Required. Variant specifying the character code or string expression whose first character is used.
//!   - If `character` is a numeric value, it's treated as a character code (0-255)
//!   - If `character` is a string, only the first character is used
//!
//! ## Returns
//! Returns a `Variant` (String) containing a string of length `number` composed of the repeating character.
//!
//! ## Remarks
//! The `String` function creates strings filled with repeating characters:
//!
//! - **Character code**: If `character` is numeric (0-255), it's interpreted as an ASCII/ANSI character code
//! - **String argument**: If `character` is a string, only the first character is used (rest is ignored)
//! - **Empty string**: If `number` is 0, returns an empty string ("")
//! - **Negative number**: Causes Error 5 (Invalid procedure call or argument)
//! - **Performance**: Very efficient for creating repeating character strings
//! - **Common uses**: Creating separators, padding, borders, rulers, progress bars
//! - **Memory limit**: `number` is limited by available memory (typically up to ~2 billion characters)
//! - **Related function**: `Space` function is equivalent to `String(n, 32)` or `String(n, " ")`
//!
//! ### Character Code Examples
//! - `String(5, 65)` returns "AAAAA" (65 is ASCII code for 'A')
//! - `String(3, 42)` returns "***" (42 is ASCII code for '*')
//! - `String(10, 45)` returns "----------" (45 is ASCII code for '-')
//! - `String(4, 61)` returns "====" (61 is ASCII code for '=')
//!
//! ### String Argument Examples
//! - `String(5, "A")` returns "AAAAA"
//! - `String(3, "*")` returns "***"
//! - `String(10, "-")` returns "----------"
//! - `String(3, "Hello")` returns "HHH" (only first character used)
//!
//! ## Typical Uses
//! 1. **Line Separators**: Create horizontal lines for text output or reports
//! 2. **Padding**: Pad strings to specific widths for alignment
//! 3. **Borders**: Create box borders for text displays
//! 4. **Progress Indicators**: Build progress bars or loading indicators
//! 5. **Masking**: Create mask strings (asterisks for passwords)
//! 6. **Rulers**: Create ruler lines for text editors or debuggers
//! 7. **Fill Characters**: Initialize strings with specific fill characters
//! 8. **Visual Markers**: Create visual separators in console or debug output
//!
//! ## Basic Examples
//!
//! ### Example 1: Creating Line Separators
//! ```vb6
//! Dim separator As String
//!
//! separator = String(50, "-")     ' "--------------------------------------------------"
//! separator = String(40, "=")     ' "========================================"
//! separator = String(30, "*")     ' "******************************"
//! separator = String(20, 45)      ' "--------------------" (45 = '-')
//! ```
//!
//! ### Example 2: Padding Strings
//! ```vb6
//! Dim text As String
//! Dim padded As String
//!
//! text = "Title"
//!
//! ' Left padding with spaces
//! padded = String(10 - Len(text), " ") & text  ' "     Title"
//!
//! ' Right padding with dots
//! padded = text & String(20 - Len(text), ".")  ' "Title..............."
//! ```
//!
//! ### Example 3: Creating Box Borders
//! ```vb6
//! Sub DrawBox(width As Integer)
//!     Dim topBottom As String
//!     Dim side As String
//!     
//!     topBottom = "+" & String(width - 2, "-") & "+"
//!     side = "|" & String(width - 2, " ") & "|"
//!     
//!     Debug.Print topBottom
//!     Debug.Print side
//!     Debug.Print side
//!     Debug.Print topBottom
//! End Sub
//! ```
//!
//! ### Example 4: Progress Bar
//! ```vb6
//! Function CreateProgressBar(percent As Integer, width As Integer) As String
//!     Dim filled As Integer
//!     Dim empty As Integer
//!     
//!     filled = (percent * width) \ 100
//!     empty = width - filled
//!     
//!     CreateProgressBar = "[" & String(filled, "#") & String(empty, " ") & "]"
//!     ' Example: [#####     ] for 50%
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Create Horizontal Line
//! ```vb6
//! Function HorizontalLine(width As Integer, Optional char As String = "-") As String
//!     HorizontalLine = String(width, char)
//! End Function
//! ```
//!
//! ### Pattern 2: Center Text
//! ```vb6
//! Function CenterText(text As String, width As Integer) As String
//!     Dim padding As Integer
//!     Dim leftPad As Integer
//!     Dim rightPad As Integer
//!     
//!     If Len(text) >= width Then
//!         CenterText = Left$(text, width)
//!         Exit Function
//!     End If
//!     
//!     padding = width - Len(text)
//!     leftPad = padding \ 2
//!     rightPad = padding - leftPad
//!     
//!     CenterText = String(leftPad, " ") & text & String(rightPad, " ")
//! End Function
//! ```
//!
//! ### Pattern 3: Pad Number with Zeros
//! ```vb6
//! Function PadWithZeros(value As Long, totalWidth As Integer) As String
//!     Dim numStr As String
//!     numStr = CStr(value)
//!     
//!     If Len(numStr) >= totalWidth Then
//!         PadWithZeros = numStr
//!     Else
//!         PadWithZeros = String(totalWidth - Len(numStr), "0") & numStr
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 4: Create Table Row Separator
//! ```vb6
//! Function TableSeparator(columnWidths() As Integer) As String
//!     Dim i As Integer
//!     Dim result As String
//!     
//!     result = "+"
//!     For i = LBound(columnWidths) To UBound(columnWidths)
//!         result = result & String(columnWidths(i), "-") & "+"
//!     Next i
//!     
//!     TableSeparator = result
//! End Function
//! ```
//!
//! ### Pattern 5: Mask Sensitive Data
//! ```vb6
//! Function MaskString(text As String, Optional maskChar As String = "*") As String
//!     MaskString = String(Len(text), maskChar)
//! End Function
//! ```
//!
//! ### Pattern 6: Indent Text
//! ```vb6
//! Function IndentText(text As String, level As Integer, _
//!                     Optional indentSize As Integer = 4) As String
//!     IndentText = String(level * indentSize, " ") & text
//! End Function
//! ```
//!
//! ### Pattern 7: Create Ruler
//! ```vb6
//! Function CreateRuler(length As Integer) As String
//!     Dim i As Integer
//!     Dim ruler As String
//!     Dim markers As String
//!     
//!     ' Create tick marks every 10 characters
//!     ruler = ""
//!     markers = ""
//!     For i = 1 To length
//!         If i Mod 10 = 0 Then
//!             ruler = ruler & "|"
//!             markers = markers & CStr(i \ 10 Mod 10)
//!         Else
//!             ruler = ruler & "."
//!             markers = markers & " "
//!         End If
//!     Next i
//!     
//!     CreateRuler = markers & vbCrLf & ruler
//! End Function
//! ```
//!
//! ### Pattern 8: Fill to Width
//! ```vb6
//! Function FillToWidth(text As String, width As Integer, _
//!                      Optional fillChar As String = " ") As String
//!     If Len(text) >= width Then
//!         FillToWidth = Left$(text, width)
//!     Else
//!         FillToWidth = text & String(width - Len(text), fillChar)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 9: Create Loading Animation
//! ```vb6
//! Function LoadingBar(step As Integer, totalSteps As Integer, width As Integer) As String
//!     Dim filled As Integer
//!     filled = (step * width) \ totalSteps
//!     LoadingBar = String(filled, "=") & ">" & String(width - filled - 1, " ")
//! End Function
//! ```
//!
//! ### Pattern 10: Duplicate Character
//! ```vb6
//! Function RepeatChar(char As String, count As Integer) As String
//!     ' Alias for String function with clearer name
//!     RepeatChar = String(count, char)
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Text Box Drawer
//! ```vb6
//! ' Class: TextBoxDrawer
//! ' Draws ASCII boxes around text
//! Option Explicit
//!
//! Private m_Width As Integer
//! Private m_Padding As Integer
//!
//! Public Sub Initialize(width As Integer, Optional padding As Integer = 1)
//!     m_Width = width
//!     m_Padding = padding
//! End Sub
//!
//! Public Function DrawBox(title As String, lines() As String) As String
//!     Dim result As String
//!     Dim i As Integer
//!     Dim maxLen As Integer
//!     
//!     ' Find maximum line length
//!     maxLen = Len(title)
//!     For i = LBound(lines) To UBound(lines)
//!         If Len(lines(i)) > maxLen Then maxLen = Len(lines(i))
//!     Next i
//!     
//!     ' Adjust width if needed
//!     If m_Width < maxLen + (m_Padding * 2) + 2 Then
//!         m_Width = maxLen + (m_Padding * 2) + 2
//!     End If
//!     
//!     ' Top border
//!     result = "+" & String(m_Width - 2, "-") & "+" & vbCrLf
//!     
//!     ' Title (centered)
//!     If Len(title) > 0 Then
//!         result = result & "|" & CenterInWidth(title, m_Width - 2) & "|" & vbCrLf
//!         result = result & "+" & String(m_Width - 2, "-") & "+" & vbCrLf
//!     End If
//!     
//!     ' Content lines
//!     For i = LBound(lines) To UBound(lines)
//!         result = result & "|" & PadLine(lines(i), m_Width - 2) & "|" & vbCrLf
//!     Next i
//!     
//!     ' Bottom border
//!     result = result & "+" & String(m_Width - 2, "-") & "+"
//!     
//!     DrawBox = result
//! End Function
//!
//! Private Function CenterInWidth(text As String, width As Integer) As String
//!     Dim leftPad As Integer
//!     Dim rightPad As Integer
//!     Dim totalPad As Integer
//!     
//!     totalPad = width - Len(text)
//!     leftPad = totalPad \ 2
//!     rightPad = totalPad - leftPad
//!     
//!     CenterInWidth = String(leftPad, " ") & text & String(rightPad, " ")
//! End Function
//!
//! Private Function PadLine(text As String, width As Integer) As String
//!     Dim content As String
//!     content = String(m_Padding, " ") & text
//!     PadLine = content & String(width - Len(content), " ")
//! End Function
//! ```
//!
//! ### Example 2: Progress Bar Generator
//! ```vb6
//! ' Class: ProgressBarGenerator
//! ' Creates customizable progress bars
//! Option Explicit
//!
//! Private m_Width As Integer
//! Private m_FilledChar As String
//! Private m_EmptyChar As String
//! Private m_ShowPercent As Boolean
//!
//! Public Sub Initialize(width As Integer, Optional filledChar As String = "#", _
//!                       Optional emptyChar As String = "-", _
//!                       Optional showPercent As Boolean = True)
//!     m_Width = width
//!     m_FilledChar = filledChar
//!     m_EmptyChar = emptyChar
//!     m_ShowPercent = showPercent
//! End Sub
//!
//! Public Function Generate(current As Long, total As Long) As String
//!     Dim percent As Integer
//!     Dim filled As Integer
//!     Dim empty As Integer
//!     Dim bar As String
//!     
//!     If total = 0 Then
//!         Generate = "[" & String(m_Width, m_EmptyChar) & "]"
//!         Exit Function
//!     End If
//!     
//!     percent = CInt((current * 100) / total)
//!     If percent > 100 Then percent = 100
//!     
//!     filled = (percent * m_Width) \ 100
//!     empty = m_Width - filled
//!     
//!     bar = "[" & String(filled, m_FilledChar) & String(empty, m_EmptyChar) & "]"
//!     
//!     If m_ShowPercent Then
//!         bar = bar & " " & Format$(percent, "000") & "%"
//!     End If
//!     
//!     Generate = bar
//! End Function
//!
//! Public Function GenerateIndeterminate(step As Integer) As String
//!     ' Animated indeterminate progress bar
//!     Dim pos As Integer
//!     Dim blockSize As Integer
//!     
//!     blockSize = 5
//!     pos = step Mod (m_Width + blockSize)
//!     
//!     If pos < blockSize Then
//!         GenerateIndeterminate = "[" & String(pos, m_FilledChar) & _
//!                                String(m_Width - pos, m_EmptyChar) & "]"
//!     ElseIf pos < m_Width Then
//!         GenerateIndeterminate = "[" & String(pos - blockSize, m_EmptyChar) & _
//!                                String(blockSize, m_FilledChar) & _
//!                                String(m_Width - pos, m_EmptyChar) & "]"
//!     Else
//!         GenerateIndeterminate = "[" & String(m_Width - (pos - m_Width), m_EmptyChar) & _
//!                                String(pos - m_Width, m_FilledChar) & "]"
//!     End If
//! End Function
//! ```
//!
//! ### Example 3: Table Formatter Module
//! ```vb6
//! ' Module: TableFormatter
//! ' Creates formatted ASCII tables
//! Option Explicit
//!
//! Public Function CreateTable(headers() As String, data() As String, _
//!                             columnWidths() As Integer) As String
//!     Dim result As String
//!     Dim i As Integer
//!     Dim j As Integer
//!     Dim row As Integer
//!     Dim col As Integer
//!     
//!     ' Top border
//!     result = CreateBorder(columnWidths, True) & vbCrLf
//!     
//!     ' Headers
//!     result = result & CreateRow(headers, columnWidths) & vbCrLf
//!     
//!     ' Separator
//!     result = result & CreateBorder(columnWidths, False) & vbCrLf
//!     
//!     ' Data rows
//!     row = LBound(data, 1)
//!     Do While row <= UBound(data, 1)
//!         Dim rowData() As String
//!         ReDim rowData(LBound(headers) To UBound(headers))
//!         
//!         For col = LBound(headers) To UBound(headers)
//!             rowData(col) = data(row, col)
//!         Next col
//!         
//!         result = result & CreateRow(rowData, columnWidths) & vbCrLf
//!         row = row + 1
//!     Loop
//!     
//!     ' Bottom border
//!     result = result & CreateBorder(columnWidths, True)
//!     
//!     CreateTable = result
//! End Function
//!
//! Private Function CreateBorder(columnWidths() As Integer, heavy As Boolean) As String
//!     Dim i As Integer
//!     Dim result As String
//!     Dim corner As String
//!     Dim line As String
//!     
//!     If heavy Then
//!         corner = "+"
//!         line = "="
//!     Else
//!         corner = "+"
//!         line = "-"
//!     End If
//!     
//!     result = corner
//!     For i = LBound(columnWidths) To UBound(columnWidths)
//!         result = result & String(columnWidths(i) + 2, line) & corner
//!     Next i
//!     
//!     CreateBorder = result
//! End Function
//!
//! Private Function CreateRow(cells() As String, columnWidths() As Integer) As String
//!     Dim i As Integer
//!     Dim result As String
//!     Dim cell As String
//!     
//!     result = "|"
//!     For i = LBound(cells) To UBound(cells)
//!         cell = " " & cells(i)
//!         If Len(cell) < columnWidths(i) + 1 Then
//!             cell = cell & String(columnWidths(i) + 1 - Len(cell), " ")
//!         End If
//!         result = result & cell & " |"
//!     Next i
//!     
//!     CreateRow = result
//! End Function
//! ```
//!
//! ### Example 4: Text Padding Utilities
//! ```vb6
//! ' Module: TextPaddingUtils
//! ' Utilities for padding and aligning text
//! Option Explicit
//!
//! Public Function PadLeft(text As String, width As Integer, _
//!                         Optional padChar As String = " ") As String
//!     If Len(text) >= width Then
//!         PadLeft = text
//!     Else
//!         PadLeft = String(width - Len(text), padChar) & text
//!     End If
//! End Function
//!
//! Public Function PadRight(text As String, width As Integer, _
//!                          Optional padChar As String = " ") As String
//!     If Len(text) >= width Then
//!         PadRight = text
//!     Else
//!         PadRight = text & String(width - Len(text), padChar)
//!     End If
//! End Function
//!
//! Public Function PadCenter(text As String, width As Integer, _
//!                           Optional padChar As String = " ") As String
//!     Dim totalPad As Integer
//!     Dim leftPad As Integer
//!     Dim rightPad As Integer
//!     
//!     If Len(text) >= width Then
//!         PadCenter = text
//!         Exit Function
//!     End If
//!     
//!     totalPad = width - Len(text)
//!     leftPad = totalPad \ 2
//!     rightPad = totalPad - leftPad
//!     
//!     PadCenter = String(leftPad, padChar) & text & String(rightPad, padChar)
//! End Function
//!
//! Public Function CreateLine(width As Integer, Optional lineChar As String = "-") As String
//!     CreateLine = String(width, lineChar)
//! End Function
//!
//! Public Function FrameText(text As String, width As Integer, _
//!                           Optional frameChar As String = "*") As String
//!     Dim topBottom As String
//!     Dim middle As String
//!     
//!     topBottom = String(width, frameChar)
//!     middle = frameChar & PadCenter(text, width - 2) & frameChar
//!     
//!     FrameText = topBottom & vbCrLf & middle & vbCrLf & topBottom
//! End Function
//!
//! Public Function Underline(text As String, Optional underlineChar As String = "-") As String
//!     Underline = text & vbCrLf & String(Len(text), underlineChar)
//! End Function
//!
//! Public Function NumberedLine(lineNum As Integer, text As String, _
//!                              totalWidth As Integer) As String
//!     Dim numStr As String
//!     numStr = PadLeft(CStr(lineNum), 4) & ": "
//!     NumberedLine = numStr & text & String(totalWidth - Len(numStr) - Len(text), " ")
//! End Function
//! ```
//!
//! ## Error Handling
//! The `String` function can raise the following errors:
//!
//! - **Error 5 (Invalid procedure call or argument)**: If `number` is negative
//! - **Error 6 (Overflow)**: If `number` is too large (exceeds memory limits)
//! - **Error 13 (Type mismatch)**: If `character` cannot be converted to a valid character
//! - **Error 5 (Invalid procedure call)**: If `character` is a numeric value outside 0-255 range
//!
//! ## Performance Notes
//! - Very fast and efficient for creating repeating character strings
//! - More efficient than concatenating characters in a loop
//! - Memory allocation is done once for the entire string
//! - For very large strings (millions of characters), consider memory constraints
//! - Slightly faster than equivalent loop-based approaches
//!
//! ## Best Practices
//! 1. **Use descriptive variable names** when storing String results (e.g., `separator`, `padding`)
//! 2. **Validate number parameter** to ensure it's non-negative before calling
//! 3. **Use named constants** for common character codes (e.g., `Const ASTERISK = 42`)
//! 4. **Prefer string literals** over character codes for clarity (e.g., `String(5, "*")` vs `String(5, 42)`)
//! 5. **Consider Space function** for creating space-filled strings (`Space(n)` vs `String(n, " ")`)
//! 6. **Cache repeated strings** if used multiple times in a function
//! 7. **Check string length** before padding operations to avoid overflow
//! 8. **Use for visual elements** like separators, borders, and progress indicators
//! 9. **Document character choice** when using character codes instead of string literals
//! 10. **Test with edge cases** like `number = 0` and very large values
//!
//! ## Comparison Table
//!
//! | Function | Purpose | Example | Result |
//! |----------|---------|---------|--------|
//! | `String` | Repeat character | `String(5, "*")` | `"*****"` |
//! | `Space` | Repeat space | `Space(5)` | `"     "` |
//! | `String$` | Type-declared variant | `String$(5, 42)` | `"*****"` |
//!
//! ## Platform Notes
//! - Available in VB6 and VBA
//! - Not available in `VBScript` (use custom function instead)
//! - `String$` variant returns String type directly (slightly faster)
//! - Behavior consistent across all platforms
//! - Character codes follow ANSI/ASCII standard (0-255)
//! - Maximum string length limited by available memory
//!
//! ## Limitations
//! - Cannot create strings with multiple different repeating characters in one call
//! - Character code must be 0-255 (no Unicode character codes)
//! - If `character` is a string, only first character is used (no validation of length)
//! - Very large `number` values can cause out-of-memory errors
//! - No built-in way to create alternating patterns
//! - Cannot specify different characters at different positions

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn string_basic() {
        let source = r#"
Sub Test()
    result = String(50, "-")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_variable_assignment() {
        let source = r#"
Sub Test()
    Dim separator As String
    separator = String(width, char)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
        assert!(debug.contains("width"));
    }

    #[test]
    fn string_character_code() {
        let source = r#"
Sub Test()
    result = String(10, 65)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_asterisk() {
        let source = r#"
Sub Test()
    line = String(20, "*")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_padding() {
        let source = r#"
Sub Test()
    padded = text & String(width - Len(text), " ")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_concatenation() {
        let source = r#"
Sub Test()
    box = "+" & String(width - 2, "-") & "+"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_if_statement() {
        let source = r#"
Sub Test()
    If Len(text) < width Then
        text = text & String(width - Len(text), " ")
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_for_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        Debug.Print String(i, "*")
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_function_return() {
        let source = r#"
Function CreateLine(width As Integer) As String
    CreateLine = String(width, "-")
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_function_argument() {
        let source = r#"
Sub Test()
    Call DrawLine(String(50, "="))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_array_assignment() {
        let source = r#"
Sub Test()
    lines(i) = String(width, "-")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print String(40, "=")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_msgbox() {
        let source = r#"
Sub Test()
    MsgBox String(5, "*") & " Alert " & String(5, "*")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_select_case() {
        let source = r#"
Sub Test()
    Select Case lineType
        Case 1
            line = String(width, "-")
        Case 2
            line = String(width, "=")
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_do_while() {
        let source = r#"
Sub Test()
    Do While Len(buffer) < targetSize
        buffer = buffer & String(10, " ")
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_do_until() {
        let source = r#"
Sub Test()
    Do Until Len(str) >= width
        str = str & String(1, fillChar)
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_while_wend() {
        let source = r#"
Sub Test()
    While i < count
        output = output & String(5, "*")
        i = i + 1
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_iif() {
        let source = r#"
Sub Test()
    line = IIf(style = 1, String(width, "-"), String(width, "="))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_with_statement() {
        let source = r#"
Sub Test()
    With textBox
        .Border = String(.Width, "-")
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_parentheses() {
        let source = r#"
Sub Test()
    result = (String(10, "*") & text & String(10, "*"))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    separator = String(width, char)
    If Err.Number <> 0 Then
        separator = ""
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_property_assignment() {
        let source = r#"
Sub Test()
    obj.Separator = String(100, "=")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_progress_bar() {
        let source = r##"
Sub Test()
    bar = "[" & String(filled, "#") & String(empty, " ") & "]"
End Sub
"##;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_zero_length() {
        let source = r#"
Sub Test()
    empty = String(0, "*")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_expression_length() {
        let source = r#"
Sub Test()
    line = String(maxWidth - currentWidth, " ")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_print_statement() {
        let source = r#"
Sub Test()
    Print #1, String(80, "-")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }

    #[test]
    fn string_class_usage() {
        let source = r#"
Sub Test()
    Set formatter = New TextFormatter
    formatter.SetBorder String(50, "*")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("String"));
    }
}
