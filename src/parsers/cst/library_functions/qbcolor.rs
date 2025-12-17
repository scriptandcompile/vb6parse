//! # `QBColor` Function
//!
//! Returns a Long representing the RGB color code corresponding to the specified color number.
//!
//! ## Syntax
//!
//! ```vb
//! QBColor(color)
//! ```
//!
//! ## Parameters
//!
//! - `color` - Required. Integer in the range 0-15 that represents a color value from the QBasic/DOS era color palette.
//!
//! ## Return Value
//!
//! Returns a `Long` representing the RGB color code that corresponds to the specified `QBasic` color number. The returned value can be used with Visual Basic's color properties.
//!
//! ## Remarks
//!
//! The `QBColor` function provides backward compatibility with `QBasic` and `QuickBASIC` programs by converting the 16-color palette used in DOS applications to RGB values usable in Windows applications.
//!
//! The color argument must be an integer from 0 to 15. Each number corresponds to a specific color from the classic DOS/QBasic palette:
//!
//! | Number | Color Name      | RGB Value      | Hex       |
//! |--------|----------------|----------------|-----------|
//! | 0      | Black          | RGB(0,0,0)     | &H000000  |
//! | 1      | Blue           | RGB(0,0,128)   | &H800000  |
//! | 2      | Green          | RGB(0,128,0)   | &H008000  |
//! | 3      | Cyan           | RGB(0,128,128) | &H808000  |
//! | 4      | Red            | RGB(128,0,0)   | &H000080  |
//! | 5      | Magenta        | RGB(128,0,128) | &H800080  |
//! | 6      | Yellow         | RGB(128,128,0) | &H008080  |
//! | 7      | White          | RGB(192,192,192)| &HC0C0C0 |
//! | 8      | Gray           | RGB(128,128,128)| &H808080 |
//! | 9      | Light Blue     | RGB(0,0,255)   | &HFF0000  |
//! | 10     | Light Green    | RGB(0,255,0)   | &H00FF00  |
//! | 11     | Light Cyan     | RGB(0,255,255) | &HFFFF00  |
//! | 12     | Light Red      | RGB(255,0,0)   | &H0000FF  |
//! | 13     | Light Magenta  | RGB(255,0,255) | &HFF00FF  |
//! | 14     | Light Yellow   | RGB(255,255,0) | &H00FFFF  |
//! | 15     | Bright White   | RGB(255,255,255)| &HFFFFFF |
//!
//! **Important Notes**:
//! - Colors 0-7 are the standard intensity colors
//! - Colors 8-15 are the high intensity (bright) versions
//! - The RGB values use BGR byte order when stored as Long values
//! - Values outside 0-15 will cause an "Invalid procedure call or argument" error
//!
//! ## Typical Uses
//!
//! 1. **Legacy Code Migration**: Converting QBasic/DOS applications to VB6/Windows
//! 2. **Console-Style Interfaces**: Creating retro-style applications with classic color schemes
//! 3. **Educational Programs**: Teaching programming with familiar DOS color palette
//! 4. **Text Display**: Coloring text output in legacy-compatible ways
//! 5. **Form Backgrounds**: Setting form or control colors using `QBasic` conventions
//! 6. **Chart/Graph Colors**: Using classic palette for data visualization
//! 7. **Terminal Emulation**: Emulating DOS/console applications
//! 8. **Game Development**: Retro game development with classic color palette
//!
//! ## Basic Examples
//!
//! ### Example 1: Setting Form Background
//! ```vb
//! ' Set form background to bright blue (QBasic color 9)
//! Form1.BackColor = QBColor(9)
//! ```
//!
//! ### Example 2: Setting Text Color
//! ```vb
//! ' Set label text to bright yellow (QBasic color 14)
//! Label1.ForeColor = QBColor(14)
//! ```
//!
//! ### Example 3: Cycling Through Colors
//! ```vb
//! ' Cycle through all 16 QBasic colors
//! Dim i As Integer
//! For i = 0 To 15
//!     Picture1.Line (i * 20, 0)-(i * 20 + 19, 100), QBColor(i), BF
//! Next i
//! ```
//!
//! ### Example 4: Conditional Coloring
//! ```vb
//! ' Color-code values: green for positive, red for negative
//! If value >= 0 Then
//!     Label1.ForeColor = QBColor(10)  ' Light Green
//! Else
//!     Label1.ForeColor = QBColor(12)  ' Light Red
//! End If
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `ColorByNumber`
//! ```vb
//! Function ColorByNumber(colorNum As Integer) As Long
//!     ' Safely convert color number with bounds checking
//!     If colorNum < 0 Then colorNum = 0
//!     If colorNum > 15 Then colorNum = 15
//!     ColorByNumber = QBColor(colorNum)
//! End Function
//! ```
//!
//! ### Pattern 2: `GetColorName`
//! ```vb
//! Function GetColorName(colorNum As Integer) As String
//!     ' Return descriptive name for QBasic color number
//!     Select Case colorNum
//!         Case 0: GetColorName = "Black"
//!         Case 1: GetColorName = "Blue"
//!         Case 2: GetColorName = "Green"
//!         Case 3: GetColorName = "Cyan"
//!         Case 4: GetColorName = "Red"
//!         Case 5: GetColorName = "Magenta"
//!         Case 6: GetColorName = "Yellow"
//!         Case 7: GetColorName = "White"
//!         Case 8: GetColorName = "Gray"
//!         Case 9: GetColorName = "Light Blue"
//!         Case 10: GetColorName = "Light Green"
//!         Case 11: GetColorName = "Light Cyan"
//!         Case 12: GetColorName = "Light Red"
//!         Case 13: GetColorName = "Light Magenta"
//!         Case 14: GetColorName = "Light Yellow"
//!         Case 15: GetColorName = "Bright White"
//!         Case Else: GetColorName = "Unknown"
//!     End Select
//! End Function
//! ```
//!
//! ### Pattern 3: `ColorPalettePicker`
//! ```vb
//! Sub ShowColorPalette()
//!     Dim i As Integer
//!     Dim x As Integer, y As Integer
//!     
//!     ' Display all 16 colors in a 4x4 grid
//!     For i = 0 To 15
//!         x = (i Mod 4) * 60
//!         y = (i \ 4) * 40
//!         
//!         Picture1.Line (x, y)-(x + 55, y + 35), QBColor(i), BF
//!         Picture1.ForeColor = QBColor(15 - i)  ' Contrast color
//!         Picture1.CurrentX = x + 5
//!         Picture1.CurrentY = y + 10
//!         Picture1.Print i
//!     Next i
//! End Sub
//! ```
//!
//! ### Pattern 4: `ApplyColorScheme`
//! ```vb
//! Sub ApplyColorScheme(bgColor As Integer, fgColor As Integer, _
//!                      Optional ctrl As Control = Nothing)
//!     ' Apply QBasic color scheme to control or form
//!     If ctrl Is Nothing Then
//!         Me.BackColor = QBColor(bgColor)
//!         Me.ForeColor = QBColor(fgColor)
//!     Else
//!         ctrl.BackColor = QBColor(bgColor)
//!         ctrl.ForeColor = QBColor(fgColor)
//!     End If
//! End Sub
//! ```
//!
//! ### Pattern 5: `ValidateColorNumber`
//! ```vb
//! Function ValidateColorNumber(colorNum As Integer) As Boolean
//!     ' Check if color number is in valid range
//!     ValidateColorNumber = (colorNum >= 0 And colorNum <= 15)
//! End Function
//! ```
//!
//! ### Pattern 6: `GetComplementaryColor`
//! ```vb
//! Function GetComplementaryColor(colorNum As Integer) As Long
//!     ' Get a contrasting color for readability
//!     If colorNum >= 0 And colorNum <= 7 Then
//!         ' Dark colors: use white/bright white
//!         GetComplementaryColor = QBColor(15)
//!     Else
//!         ' Light colors: use black
//!         GetComplementaryColor = QBColor(0)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 7: `ColorizeText`
//! ```vb
//! Sub ColorizeText(textBox As TextBox, colorCode As Integer)
//!     ' Apply color with error handling
//!     On Error Resume Next
//!     textBox.ForeColor = QBColor(colorCode)
//!     If Err.Number <> 0 Then
//!         textBox.ForeColor = vbBlack  ' Default to black on error
//!         Err.Clear
//!     End If
//!     On Error GoTo 0
//! End Sub
//! ```
//!
//! ### Pattern 8: `CreateColorGradient`
//! ```vb
//! Sub CreateColorGradient(pic As PictureBox, startColor As Integer, _
//!                        endColor As Integer, steps As Integer)
//!     ' Create gradient using QBasic colors
//!     Dim i As Integer
//!     Dim stepHeight As Single
//!     
//!     stepHeight = pic.ScaleHeight / steps
//!     
//!     For i = 0 To steps - 1
//!         Dim blendColor As Integer
//!         blendColor = startColor + ((endColor - startColor) * i \ steps)
//!         If blendColor < 0 Then blendColor = 0
//!         If blendColor > 15 Then blendColor = 15
//!         
//!         pic.Line (0, i * stepHeight)-(pic.ScaleWidth, (i + 1) * stepHeight), _
//!                  QBColor(blendColor), BF
//!     Next i
//! End Sub
//! ```
//!
//! ### Pattern 9: `HighlightControl`
//! ```vb
//! Sub HighlightControl(ctrl As Control, highlight As Boolean)
//!     ' Toggle control highlighting using QBasic colors
//!     If highlight Then
//!         ctrl.BackColor = QBColor(14)  ' Light Yellow
//!         ctrl.ForeColor = QBColor(0)   ' Black
//!     Else
//!         ctrl.BackColor = QBColor(15)  ' White
//!         ctrl.ForeColor = QBColor(0)   ' Black
//!     End If
//! End Sub
//! ```
//!
//! ### Pattern 10: `ColorCodeStatus`
//! ```vb
//! Function GetStatusColor(status As String) As Long
//!     ' Return color based on status string
//!     Select Case UCase(status)
//!         Case "ERROR", "CRITICAL"
//!             GetStatusColor = QBColor(12)  ' Light Red
//!         Case "WARNING"
//!             GetStatusColor = QBColor(14)  ' Light Yellow
//!         Case "SUCCESS", "OK"
//!             GetStatusColor = QBColor(10)  ' Light Green
//!         Case "INFO"
//!             GetStatusColor = QBColor(9)   ' Light Blue
//!         Case Else
//!             GetStatusColor = QBColor(7)   ' White
//!     End Select
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Console-Style Output Window
//! ```vb
//! ' Create a DOS-style console window with classic colors
//! Class ConsoleWindow
//!     Private m_textBox As TextBox
//!     Private m_currentColor As Integer
//!     Private m_bgColor As Integer
//!     
//!     Public Sub Initialize(textBox As TextBox)
//!         Set m_textBox = textBox
//!         m_currentColor = 7   ' White
//!         m_bgColor = 0        ' Black
//!         
//!         ' Setup for console appearance
//!         With m_textBox
//!             .BackColor = QBColor(m_bgColor)
//!             .ForeColor = QBColor(m_currentColor)
//!             .Font.Name = "Courier New"
//!             .Font.Size = 10
//!             .MultiLine = True
//!             .ScrollBars = vbVertical
//!         End With
//!     End Sub
//!     
//!     Public Sub SetColor(colorNum As Integer)
//!         If colorNum >= 0 And colorNum <= 15 Then
//!             m_currentColor = colorNum
//!         End If
//!     End Sub
//!     
//!     Public Sub SetBackColor(colorNum As Integer)
//!         If colorNum >= 0 And colorNum <= 15 Then
//!             m_bgColor = colorNum
//!             m_textBox.BackColor = QBColor(m_bgColor)
//!         End If
//!     End Sub
//!     
//!     Public Sub Print(text As String)
//!         ' Note: VB6 doesn't support rich text easily in TextBox
//!         ' This is simplified - use RichTextBox for multi-color text
//!         m_textBox.ForeColor = QBColor(m_currentColor)
//!         m_textBox.Text = m_textBox.Text & text & vbCrLf
//!         m_textBox.SelStart = Len(m_textBox.Text)
//!     End Sub
//!     
//!     Public Sub PrintColored(text As String, colorNum As Integer)
//!         Dim oldColor As Integer
//!         oldColor = m_currentColor
//!         SetColor colorNum
//!         Print text
//!         SetColor oldColor
//!     End Sub
//!     
//!     Public Sub Clear()
//!         m_textBox.Text = ""
//!     End Sub
//!     
//!     Public Sub ShowColorTest()
//!         Dim i As Integer
//!         Clear
//!         
//!         For i = 0 To 15
//!             PrintColored "Color " & i & ": " & GetColorName(i), i
//!         Next i
//!     End Sub
//!     
//!     Private Function GetColorName(colorNum As Integer) As String
//!         Select Case colorNum
//!             Case 0: GetColorName = "Black"
//!             Case 1: GetColorName = "Blue"
//!             Case 2: GetColorName = "Green"
//!             Case 3: GetColorName = "Cyan"
//!             Case 4: GetColorName = "Red"
//!             Case 5: GetColorName = "Magenta"
//!             Case 6: GetColorName = "Yellow"
//!             Case 7: GetColorName = "White"
//!             Case 8: GetColorName = "Gray"
//!             Case 9: GetColorName = "Light Blue"
//!             Case 10: GetColorName = "Light Green"
//!             Case 11: GetColorName = "Light Cyan"
//!             Case 12: GetColorName = "Light Red"
//!             Case 13: GetColorName = "Light Magenta"
//!             Case 14: GetColorName = "Light Yellow"
//!             Case 15: GetColorName = "Bright White"
//!             Case Else: GetColorName = "Unknown"
//!         End Select
//!     End Function
//! End Class
//! ```
//!
//! ### Example 2: Color Palette Manager
//! ```vb
//! ' Manage and display QBasic color palette
//! Module ColorPaletteManager
//!     Private Type ColorInfo
//!         Number As Integer
//!         Name As String
//!         RGBValue As Long
//!         HexValue As String
//!     End Type
//!     
//!     Private m_palette(0 To 15) As ColorInfo
//!     Private m_initialized As Boolean
//!     
//!     Public Sub InitializePalette()
//!         Dim i As Integer
//!         
//!         For i = 0 To 15
//!             m_palette(i).Number = i
//!             m_palette(i).Name = GetColorName(i)
//!             m_palette(i).RGBValue = QBColor(i)
//!             m_palette(i).HexValue = "&H" & Right("000000" & Hex(QBColor(i)), 6)
//!         Next i
//!         
//!         m_initialized = True
//!     End Sub
//!     
//!     Public Function GetColorInfo(colorNum As Integer) As String
//!         If Not m_initialized Then InitializePalette
//!         
//!         If colorNum < 0 Or colorNum > 15 Then
//!             GetColorInfo = "Invalid color number"
//!             Exit Function
//!         End If
//!         
//!         With m_palette(colorNum)
//!             GetColorInfo = "Color " & .Number & ": " & .Name & _
//!                           " (RGB: " & .RGBValue & ", Hex: " & .HexValue & ")"
//!         End With
//!     End Function
//!     
//!     Public Function FindColorByName(colorName As String) As Integer
//!         Dim i As Integer
//!         
//!         If Not m_initialized Then InitializePalette
//!         
//!         For i = 0 To 15
//!             If UCase(m_palette(i).Name) = UCase(colorName) Then
//!                 FindColorByName = i
//!                 Exit Function
//!             End If
//!         Next i
//!         
//!         FindColorByName = -1  ' Not found
//!     End Function
//!     
//!     Public Sub DisplayPalette(pic As PictureBox)
//!         Dim i As Integer
//!         Dim x As Single, y As Single
//!         Dim boxSize As Single
//!         
//!         If Not m_initialized Then InitializePalette
//!         
//!         pic.Cls
//!         boxSize = pic.ScaleWidth / 4
//!         
//!         For i = 0 To 15
//!             x = (i Mod 4) * boxSize
//!             y = (i \ 4) * boxSize
//!             
//!             ' Draw color box
//!             pic.Line (x, y)-(x + boxSize - 5, y + boxSize - 5), _
//!                      QBColor(i), BF
//!             
//!             ' Draw color number
//!             pic.ForeColor = GetComplementaryColor(i)
//!             pic.CurrentX = x + 5
//!             pic.CurrentY = y + 5
//!             pic.Print Format(i, "00")
//!         Next i
//!     End Sub
//!     
//!     Public Function ExportPaletteToHTML() As String
//!         Dim html As String
//!         Dim i As Integer
//!         
//!         If Not m_initialized Then InitializePalette
//!         
//!         html = "<table border='1'>" & vbCrLf
//!         html = html & "<tr><th>Number</th><th>Name</th><th>Hex</th><th>Preview</th></tr>" & vbCrLf
//!         
//!         For i = 0 To 15
//!             html = html & "<tr>"
//!             html = html & "<td>" & i & "</td>"
//!             html = html & "<td>" & m_palette(i).Name & "</td>"
//!             html = html & "<td>" & m_palette(i).HexValue & "</td>"
//!             html = html & "<td style='background-color:" & m_palette(i).HexValue & ";'>&nbsp;&nbsp;&nbsp;</td>"
//!             html = html & "</tr>" & vbCrLf
//!         Next i
//!         
//!         html = html & "</table>"
//!         ExportPaletteToHTML = html
//!     End Function
//!     
//!     Private Function GetColorName(colorNum As Integer) As String
//!         ' Implementation same as previous examples
//!         Select Case colorNum
//!             Case 0: GetColorName = "Black"
//!             Case 1: GetColorName = "Blue"
//!             Case 2: GetColorName = "Green"
//!             Case 3: GetColorName = "Cyan"
//!             Case 4: GetColorName = "Red"
//!             Case 5: GetColorName = "Magenta"
//!             Case 6: GetColorName = "Yellow"
//!             Case 7: GetColorName = "White"
//!             Case 8: GetColorName = "Gray"
//!             Case 9: GetColorName = "Light Blue"
//!             Case 10: GetColorName = "Light Green"
//!             Case 11: GetColorName = "Light Cyan"
//!             Case 12: GetColorName = "Light Red"
//!             Case 13: GetColorName = "Light Magenta"
//!             Case 14: GetColorName = "Light Yellow"
//!             Case 15: GetColorName = "Bright White"
//!         End Select
//!     End Function
//!     
//!     Private Function GetComplementaryColor(colorNum As Integer) As Long
//!         If colorNum >= 0 And colorNum <= 7 Then
//!             GetComplementaryColor = QBColor(15)
//!         Else
//!             GetComplementaryColor = QBColor(0)
//!         End If
//!     End Function
//! End Module
//! ```
//!
//! ### Example 3: Retro Game Color Manager
//! ```vb
//! ' Manage colors for retro-style games
//! Class GameColorManager
//!     Private m_playerColor As Integer
//!     Private m_enemyColor As Integer
//!     Private m_backgroundColors() As Integer
//!     Private m_levelColors As Collection
//!     
//!     Public Sub Initialize()
//!         Set m_levelColors = New Collection
//!         ReDim m_backgroundColors(0 To 3)
//!         
//!         ' Default colors
//!         m_playerColor = 14       ' Light Yellow
//!         m_enemyColor = 12        ' Light Red
//!         m_backgroundColors(0) = 0  ' Black
//!         m_backgroundColors(1) = 1  ' Blue
//!         m_backgroundColors(2) = 2  ' Green
//!         m_backgroundColors(3) = 5  ' Magenta
//!     End Sub
//!     
//!     Public Function GetPlayerColor() As Long
//!         GetPlayerColor = QBColor(m_playerColor)
//!     End Function
//!     
//!     Public Function GetEnemyColor() As Long
//!         GetEnemyColor = QBColor(m_enemyColor)
//!     End Function
//!     
//!     Public Function GetBackgroundColor(level As Integer) As Long
//!         Dim colorIndex As Integer
//!         colorIndex = level Mod (UBound(m_backgroundColors) + 1)
//!         GetBackgroundColor = QBColor(m_backgroundColors(colorIndex))
//!     End Function
//!     
//!     Public Sub SetLevelColorScheme(level As Integer, bgColor As Integer, _
//!                                   playerColor As Integer, enemyColor As Integer)
//!         Dim scheme As String
//!         scheme = bgColor & "," & playerColor & "," & enemyColor
//!         
//!         ' Store in collection using level as key
//!         On Error Resume Next
//!         m_levelColors.Remove CStr(level)
//!         On Error GoTo 0
//!         m_levelColors.Add scheme, CStr(level)
//!     End Sub
//!     
//!     Public Sub ApplyLevelColors(level As Integer, gameForm As Form)
//!         Dim scheme As String
//!         Dim colors() As String
//!         
//!         On Error Resume Next
//!         scheme = m_levelColors(CStr(level))
//!         On Error GoTo 0
//!         
//!         If scheme <> "" Then
//!             colors = Split(scheme, ",")
//!             If UBound(colors) = 2 Then
//!                 gameForm.BackColor = QBColor(CInt(colors(0)))
//!                 m_playerColor = CInt(colors(1))
//!                 m_enemyColor = CInt(colors(2))
//!             End If
//!         Else
//!             ' Use default colors
//!             gameForm.BackColor = GetBackgroundColor(level)
//!         End If
//!     End Sub
//!     
//!     Public Function CreateColorCycle(startColor As Integer, endColor As Integer) As Long()
//!         ' Create array of colors for animation
//!         Dim colors() As Long
//!         Dim i As Integer
//!         Dim count As Integer
//!         
//!         count = Abs(endColor - startColor) + 1
//!         ReDim colors(0 To count - 1)
//!         
//!         For i = 0 To count - 1
//!             If startColor <= endColor Then
//!                 colors(i) = QBColor(startColor + i)
//!             Else
//!                 colors(i) = QBColor(startColor - i)
//!             End If
//!         Next i
//!         
//!         CreateColorCycle = colors
//!     End Function
//! End Class
//! ```
//!
//! ### Example 4: Syntax Highlighter
//! ```vb
//! ' Simple syntax highlighter using QBasic colors
//! Class SyntaxHighlighter
//!     Private m_keywordColor As Integer
//!     Private m_stringColor As Integer
//!     Private m_commentColor As Integer
//!     Private m_numberColor As Integer
//!     Private m_normalColor As Integer
//!     
//!     Public Sub Initialize()
//!         m_keywordColor = 9    ' Light Blue
//!         m_stringColor = 12    ' Light Red
//!         m_commentColor = 2    ' Green
//!         m_numberColor = 14    ' Light Yellow
//!         m_normalColor = 15    ' Bright White
//!     End Sub
//!     
//!     Public Function GetKeywordColor() As Long
//!         GetKeywordColor = QBColor(m_keywordColor)
//!     End Function
//!     
//!     Public Function GetStringColor() As Long
//!         GetStringColor = QBColor(m_stringColor)
//!     End Function
//!     
//!     Public Function GetCommentColor() As Long
//!         GetCommentColor = QBColor(m_commentColor)
//!     End Function
//!     
//!     Public Function GetNumberColor() As Long
//!         GetNumberColor = QBColor(m_numberColor)
//!     End Function
//!     
//!     Public Function GetNormalColor() As Long
//!         GetNormalColor = QBColor(m_normalColor)
//!     End Function
//!     
//!     Public Sub SetColorScheme(scheme As String)
//!         ' Apply predefined color schemes
//!         Select Case UCase(scheme)
//!             Case "CLASSIC"
//!                 m_keywordColor = 9
//!                 m_stringColor = 12
//!                 m_commentColor = 2
//!                 m_numberColor = 14
//!                 m_normalColor = 15
//!             Case "PASTEL"
//!                 m_keywordColor = 11
//!                 m_stringColor = 13
//!                 m_commentColor = 10
//!                 m_numberColor = 14
//!                 m_normalColor = 7
//!             Case "CONTRAST"
//!                 m_keywordColor = 14
//!                 m_stringColor = 12
//!                 m_commentColor = 10
//!                 m_numberColor = 11
//!                 m_normalColor = 15
//!         End Select
//!     End Sub
//!     
//!     Public Function GetColorForType(tokenType As String) As Long
//!         Select Case UCase(tokenType)
//!             Case "KEYWORD"
//!                 GetColorForType = GetKeywordColor()
//!             Case "STRING"
//!                 GetColorForType = GetStringColor()
//!             Case "COMMENT"
//!                 GetColorForType = GetCommentColor()
//!             Case "NUMBER"
//!                 GetColorForType = GetNumberColor()
//!             Case Else
//!                 GetColorForType = GetNormalColor()
//!         End Select
//!     End Function
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! The `QBColor` function can raise errors in the following situations:
//!
//! - **Invalid Procedure Call or Argument (Error 5)**: When:
//!   - The color argument is less than 0 or greater than 15
//!   - The argument is not a valid integer
//! - **Type Mismatch (Error 13)**: When the argument cannot be converted to an integer
//!
//! Always validate the color number before calling `QBColor`:
//!
//! ```vb
//! Function SafeQBColor(colorNum As Integer) As Long
//!     On Error Resume Next
//!     SafeQBColor = QBColor(colorNum)
//!     If Err.Number <> 0 Then
//!         SafeQBColor = QBColor(7)  ' Default to white
//!         Err.Clear
//!     End If
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - The `QBColor` function is very fast - it's essentially a lookup table
//! - No performance penalty for repeated calls
//! - Can be called thousands of times without noticeable impact
//! - Consider caching if using the same color multiple times in tight loops
//!
//! ## Best Practices
//!
//! 1. **Validate Input**: Always check that color numbers are in range 0-15
//! 2. **Use Constants**: Define named constants for frequently used colors
//! 3. **Document Color Choices**: Comment why specific colors were chosen
//! 4. **Test Accessibility**: Consider color-blind users when choosing colors
//! 5. **Provide Contrast**: Ensure text is readable against background
//! 6. **Use Complementary Colors**: Pair dark backgrounds with light foregrounds
//! 7. **Consider Modern Alternatives**: For new code, consider `RGB()` function
//! 8. **Legacy Compatibility**: Use `QBColor` when porting QBasic/DOS code
//! 9. **Error Handling**: Wrap `QBColor` calls in error handlers for robustness
//! 10. **Color Naming**: Use `GetColorName` pattern for better code readability
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | **`QBColor`** | `QBasic` color to RGB | Long (RGB value) | Legacy compatibility, retro interfaces |
//! | **RGB** | Create RGB color | Long (RGB value) | Custom colors, modern applications |
//! | **vbRed, vbBlue, etc.** | Predefined constants | Long (RGB value) | Standard colors, quick coding |
//! | **`LoadPicture`** | Load image with colors | Picture object | Complex graphics, photos |
//!
//! ## Platform and Version Notes
//!
//! - Available in VB6 and all versions of VBA
//! - Behavior is consistent across all Windows platforms
//! - Returns BGR byte order (Blue-Green-Red) as is standard for Windows
//! - Color values are identical to `QBasic` and `QuickBASIC`
//! - Not available in VB.NET (use System.Drawing.Color instead)
//!
//! ## Limitations
//!
//! - Limited to 16 predefined colors only
//! - Cannot create custom colors (use RGB function instead)
//! - Color numbers must be exactly 0-15 (no wrapping or modulo)
//! - Colors may not display identically on all monitors
//! - High-intensity colors (8-15) may appear less distinct on some displays
//! - Not suitable for professional graphics requiring precise color control
//!
//! ## Related Functions
//!
//! - `RGB`: Creates custom RGB color values
//! - `Hex`: Converts numbers to hexadecimal strings
//! - `LoadPicture`: Loads images with full color support
//! - Color constants: `vbBlack`, `vbRed`, `vbGreen`, `vbBlue`, etc.

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn qbcolor_basic() {
        let source = r#"
Dim color As Long
color = QBColor(12)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_form_background() {
        let source = r#"
Form1.BackColor = QBColor(9)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_if_statement() {
        let source = r#"
If value > 0 Then
    Label1.ForeColor = QBColor(10)
Else
    Label1.ForeColor = QBColor(12)
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_function_return() {
        let source = r#"
Function GetStatusColor(status As String) As Long
    If status = "OK" Then
        GetStatusColor = QBColor(10)
    Else
        GetStatusColor = QBColor(12)
    End If
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_variable_assignment() {
        let source = r#"
Dim bgColor As Long
bgColor = QBColor(colorNumber)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_msgbox() {
        let source = r#"
MsgBox "Color value: " & QBColor(5)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_debug_print() {
        let source = r#"
Debug.Print "RGB Value: " & QBColor(14)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_select_case() {
        let source = r#"
Select Case colorIndex
    Case 0 To 7
        result = QBColor(colorIndex)
    Case 8 To 15
        result = QBColor(colorIndex)
End Select
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_class_usage() {
        let source = r#"
Private m_color As Long

Public Sub SetColor(colorNum As Integer)
    m_color = QBColor(colorNum)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_with_statement() {
        let source = r#"
With Label1
    .BackColor = QBColor(0)
    .ForeColor = QBColor(15)
End With
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_elseif() {
        let source = r#"
If level < 5 Then
    color = QBColor(2)
ElseIf level < 10 Then
    color = QBColor(10)
Else
    color = QBColor(14)
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_for_loop() {
        let source = r#"
For i = 0 To 15
    Picture1.Line (i * 20, 0)-(i * 20 + 19, 100), QBColor(i), BF
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_do_while() {
        let source = r#"
Do While colorNum <= 15
    colors(colorNum) = QBColor(colorNum)
    colorNum = colorNum + 1
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_do_until() {
        let source = r#"
Do Until i > 15
    palette(i) = QBColor(i)
    i = i + 1
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_while_wend() {
        let source = r#"
While index < 16
    colorArray(index) = QBColor(index)
    index = index + 1
Wend
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_parentheses() {
        let source = r#"
Dim result As Long
result = (QBColor(7))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_iif() {
        let source = r#"
Dim textColor As Long
textColor = IIf(isError, QBColor(12), QBColor(10))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_comparison() {
        let source = r#"
If QBColor(color1) = QBColor(color2) Then
    MsgBox "Same color"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_array_assignment() {
        let source = r#"
Dim colors(15) As Long
colors(i) = QBColor(i)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_property_assignment() {
        let source = r#"
Set obj = New ColorManager
obj.BackgroundColor = QBColor(0)
obj.TextColor = QBColor(15)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_function_argument() {
        let source = r#"
Call SetFormColors(QBColor(0), QBColor(15))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_line_method() {
        let source = r#"
Picture1.Line (0, 0)-(100, 100), QBColor(12), BF
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_concatenation() {
        let source = r#"
Dim msg As String
msg = "Color code: " & QBColor(idx)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_hex_conversion() {
        let source = r#"
Dim hexValue As String
hexValue = Hex(QBColor(15))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_arithmetic() {
        let source = r#"
Dim colorValue As Long
Dim brightness As Long
colorValue = QBColor(index)
brightness = colorValue And &HFF
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_error_handling() {
        let source = r#"
On Error Resume Next
Form1.BackColor = QBColor(colorNum)
If Err.Number <> 0 Then
    Form1.BackColor = QBColor(7)
End If
On Error GoTo 0
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn qbcolor_on_error_goto() {
        let source = r#"
Sub ApplyColor()
    On Error GoTo ErrorHandler
    Dim c As Long
    c = QBColor(colorValue)
    Exit Sub
ErrorHandler:
    MsgBox "Invalid color number"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("QBColor"));
        assert!(text.contains("Identifier"));
    }
}
