//! # RGB Function
//!
//! Returns a Long representing an RGB color value from red, green, and blue color components.
//!
//! ## Syntax
//!
//! ```vb
//! RGB(red, green, blue)
//! ```
//!
//! ## Parameters
//!
//! - `red` - Required. Integer in the range 0-255 that represents the red component of the color.
//! - `green` - Required. Integer in the range 0-255 that represents the green component of the color.
//! - `blue` - Required. Integer in the range 0-255 that represents the blue component of the color.
//!
//! ## Return Value
//!
//! Returns a `Long` representing the RGB color value. The value is calculated as:
//! ```text
//! RGB = red + (green * 256) + (blue * 65536)
//! ```
//!
//! The return value is a Long integer in BGR (Blue-Green-Red) byte order, which is the standard format for Windows color values.
//!
//! ## Remarks
//!
//! The `RGB` function creates custom color values by combining red, green, and blue components. This is the primary way to specify colors programmatically in VB6 beyond the predefined color constants.
//!
//! Each color component must be in the range 0-255:
//! - 0 represents no intensity (color is off)
//! - 255 represents maximum intensity (color is fully on)
//! - Values outside this range are automatically adjusted to fit within 0-255
//!
//! **Important Notes**:
//! - The return value uses BGR byte order (not RGB order) for Windows compatibility
//! - RGB(0, 0, 0) = Black (0x000000)
//! - RGB(255, 255, 255) = White (0xFFFFFF)
//! - RGB(255, 0, 0) = Red (0x0000FF in BGR format)
//! - RGB(0, 255, 0) = Green (0x00FF00)
//! - RGB(0, 0, 255) = Blue (0xFF0000 in BGR format)
//! - Values greater than 255 are treated as 255 (saturated)
//! - Negative values are treated as 0
//!
//! **Color Mixing**:
//! - RGB(255, 0, 0) = Pure Red
//! - RGB(0, 255, 0) = Pure Green
//! - RGB(0, 0, 255) = Pure Blue
//! - RGB(255, 255, 0) = Yellow (Red + Green)
//! - RGB(255, 0, 255) = Magenta (Red + Blue)
//! - RGB(0, 255, 255) = Cyan (Green + Blue)
//! - RGB(128, 128, 128) = Gray (equal components)
//!
//! ## Typical Uses
//!
//! 1. **Custom Colors**: Create specific colors not available as constants
//! 2. **Dynamic Coloring**: Calculate colors based on data values
//! 3. **Gradients**: Create smooth color transitions
//! 4. **Color Schemes**: Define coordinated color palettes
//! 5. **Data Visualization**: Color-code data points, charts, or graphs
//! 6. **User Preferences**: Allow users to select custom colors
//! 7. **Theme Systems**: Implement application-wide color themes
//! 8. **Image Processing**: Manipulate individual pixel colors
//!
//! ## Basic Examples
//!
//! ### Example 1: Primary Colors
//! ```vb
//! ' Set form background to pure red
//! Form1.BackColor = RGB(255, 0, 0)
//!
//! ' Set label to pure green
//! Label1.ForeColor = RGB(0, 255, 0)
//!
//! ' Set button to pure blue
//! Command1.BackColor = RGB(0, 0, 255)
//! ```
//!
//! ### Example 2: Custom Colors
//! ```vb
//! ' Create a nice orange color
//! Dim orange As Long
//! orange = RGB(255, 165, 0)
//!
//! ' Create a purple color
//! Dim purple As Long
//! purple = RGB(128, 0, 128)
//!
//! ' Create a brown color
//! Dim brown As Long
//! brown = RGB(165, 42, 42)
//! ```
//!
//! ### Example 3: Shades of Gray
//! ```vb
//! ' Create various shades of gray
//! Dim lightGray As Long
//! Dim mediumGray As Long
//! Dim darkGray As Long
//!
//! lightGray = RGB(211, 211, 211)
//! mediumGray = RGB(128, 128, 128)
//! darkGray = RGB(64, 64, 64)
//! ```
//!
//! ### Example 4: Data-Driven Coloring
//! ```vb
//! ' Color-code values based on magnitude
//! Dim value As Double
//! Dim cellColor As Long
//!
//! If value > 100 Then
//!     cellColor = RGB(255, 0, 0)      ' Red for high values
//! ElseIf value > 50 Then
//!     cellColor = RGB(255, 255, 0)    ' Yellow for medium values
//! Else
//!     cellColor = RGB(0, 255, 0)      ' Green for low values
//! End If
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `CreateGradient`
//! ```vb
//! Function CreateGradient(startColor As Long, endColor As Long, _
//!                        steps As Integer, stepNum As Integer) As Long
//!     ' Create a color between startColor and endColor
//!     Dim startR As Integer, startG As Integer, startB As Integer
//!     Dim endR As Integer, endG As Integer, endB As Integer
//!     Dim r As Integer, g As Integer, b As Integer
//!     Dim ratio As Double
//!     
//!     ' Extract RGB components from start color
//!     startR = startColor And &HFF
//!     startG = (startColor \ 256) And &HFF
//!     startB = (startColor \ 65536) And &HFF
//!     
//!     ' Extract RGB components from end color
//!     endR = endColor And &HFF
//!     endG = (endColor \ 256) And &HFF
//!     endB = (endColor \ 65536) And &HFF
//!     
//!     ' Calculate interpolation ratio
//!     ratio = stepNum / steps
//!     
//!     ' Interpolate each component
//!     r = startR + ((endR - startR) * ratio)
//!     g = startG + ((endG - startG) * ratio)
//!     b = startB + ((endB - startB) * ratio)
//!     
//!     CreateGradient = RGB(r, g, b)
//! End Function
//! ```
//!
//! ### Pattern 2: `ExtractColorComponents`
//! ```vb
//! Sub ExtractRGB(color As Long, red As Integer, green As Integer, blue As Integer)
//!     ' Extract individual RGB components from a color value
//!     red = color And &HFF
//!     green = (color \ 256) And &HFF
//!     blue = (color \ 65536) And &HFF
//! End Sub
//! ```
//!
//! ### Pattern 3: `LightenColor`
//! ```vb
//! Function LightenColor(color As Long, percent As Double) As Long
//!     ' Lighten a color by a percentage (0-100)
//!     Dim r As Integer, g As Integer, b As Integer
//!     
//!     r = color And &HFF
//!     g = (color \ 256) And &HFF
//!     b = (color \ 65536) And &HFF
//!     
//!     ' Increase each component toward 255
//!     r = r + ((255 - r) * percent / 100)
//!     g = g + ((255 - g) * percent / 100)
//!     b = b + ((255 - b) * percent / 100)
//!     
//!     LightenColor = RGB(r, g, b)
//! End Function
//! ```
//!
//! ### Pattern 4: `DarkenColor`
//! ```vb
//! Function DarkenColor(color As Long, percent As Double) As Long
//!     ' Darken a color by a percentage (0-100)
//!     Dim r As Integer, g As Integer, b As Integer
//!     
//!     r = color And &HFF
//!     g = (color \ 256) And &HFF
//!     b = (color \ 65536) And &HFF
//!     
//!     ' Decrease each component toward 0
//!     r = r - (r * percent / 100)
//!     g = g - (g * percent / 100)
//!     b = b - (b * percent / 100)
//!     
//!     DarkenColor = RGB(r, g, b)
//! End Function
//! ```
//!
//! ### Pattern 5: `BlendColors`
//! ```vb
//! Function BlendColors(color1 As Long, color2 As Long, _
//!                     Optional ratio As Double = 0.5) As Long
//!     ' Blend two colors together (ratio 0.0 = color1, 1.0 = color2)
//!     Dim r1 As Integer, g1 As Integer, b1 As Integer
//!     Dim r2 As Integer, g2 As Integer, b2 As Integer
//!     Dim r As Integer, g As Integer, b As Integer
//!     
//!     ' Extract components
//!     r1 = color1 And &HFF
//!     g1 = (color1 \ 256) And &HFF
//!     b1 = (color1 \ 65536) And &HFF
//!     
//!     r2 = color2 And &HFF
//!     g2 = (color2 \ 256) And &HFF
//!     b2 = (color2 \ 65536) And &HFF
//!     
//!     ' Blend
//!     r = r1 + ((r2 - r1) * ratio)
//!     g = g1 + ((g2 - g1) * ratio)
//!     b = b1 + ((b2 - b1) * ratio)
//!     
//!     BlendColors = RGB(r, g, b)
//! End Function
//! ```
//!
//! ### Pattern 6: `ColorFromHex`
//! ```vb
//! Function ColorFromHex(hexColor As String) As Long
//!     ' Convert hex string like "#FF0000" to RGB color
//!     Dim r As Integer, g As Integer, b As Integer
//!     
//!     ' Remove # if present
//!     If Left(hexColor, 1) = "#" Then
//!         hexColor = Mid(hexColor, 2)
//!     End If
//!     
//!     ' Extract components (assumes format RRGGBB)
//!     r = Val("&H" & Mid(hexColor, 1, 2))
//!     g = Val("&H" & Mid(hexColor, 3, 2))
//!     b = Val("&H" & Mid(hexColor, 5, 2))
//!     
//!     ColorFromHex = RGB(r, g, b)
//! End Function
//! ```
//!
//! ### Pattern 7: `ColorToHex`
//! ```vb
//! Function ColorToHex(color As Long) As String
//!     ' Convert RGB color to hex string
//!     Dim r As Integer, g As Integer, b As Integer
//!     
//!     r = color And &HFF
//!     g = (color \ 256) And &HFF
//!     b = (color \ 65536) And &HFF
//!     
//!     ColorToHex = "#" & _
//!                  Right("0" & Hex(r), 2) & _
//!                  Right("0" & Hex(g), 2) & _
//!                  Right("0" & Hex(b), 2)
//! End Function
//! ```
//!
//! ### Pattern 8: `GetContrastColor`
//! ```vb
//! Function GetContrastColor(bgColor As Long) As Long
//!     ' Return black or white for best contrast
//!     Dim r As Integer, g As Integer, b As Integer
//!     Dim luminance As Double
//!     
//!     r = bgColor And &HFF
//!     g = (bgColor \ 256) And &HFF
//!     b = (bgColor \ 65536) And &HFF
//!     
//!     ' Calculate perceived luminance
//!     luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
//!     
//!     If luminance > 0.5 Then
//!         GetContrastColor = RGB(0, 0, 0)      ' Black
//!     Else
//!         GetContrastColor = RGB(255, 255, 255) ' White
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 9: `TemperatureToColor`
//! ```vb
//! Function TemperatureToColor(temp As Double, minTemp As Double, _
//!                            maxTemp As Double) As Long
//!     ' Map temperature to color (blue = cold, red = hot)
//!     Dim ratio As Double
//!     Dim r As Integer, g As Integer, b As Integer
//!     
//!     ' Normalize temperature to 0-1 range
//!     ratio = (temp - minTemp) / (maxTemp - minTemp)
//!     If ratio < 0 Then ratio = 0
//!     If ratio > 1 Then ratio = 1
//!     
//!     If ratio < 0.5 Then
//!         ' Blue to cyan to green
//!         r = 0
//!         g = ratio * 2 * 255
//!         b = 255 - (ratio * 2 * 255)
//!     Else
//!         ' Green to yellow to red
//!         r = (ratio - 0.5) * 2 * 255
//!         g = 255 - ((ratio - 0.5) * 2 * 255)
//!         b = 0
//!     End If
//!     
//!     TemperatureToColor = RGB(r, g, b)
//! End Function
//! ```
//!
//! ### Pattern 10: `InvertColor`
//! ```vb
//! Function InvertColor(color As Long) As Long
//!     ' Invert a color (negative)
//!     Dim r As Integer, g As Integer, b As Integer
//!     
//!     r = color And &HFF
//!     g = (color \ 256) And &HFF
//!     b = (color \ 65536) And &HFF
//!     
//!     InvertColor = RGB(255 - r, 255 - g, 255 - b)
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Color Palette Manager
//! ```vb
//! ' Comprehensive color palette management system
//! Class ColorPalette
//!     Private m_colors() As Long
//!     Private m_colorNames() As String
//!     Private m_count As Integer
//!     
//!     Public Sub Initialize()
//!         m_count = 0
//!         ReDim m_colors(0 To 99)
//!         ReDim m_colorNames(0 To 99)
//!         LoadDefaultColors
//!     End Sub
//!     
//!     Private Sub LoadDefaultColors()
//!         AddColor "Red", RGB(255, 0, 0)
//!         AddColor "Green", RGB(0, 255, 0)
//!         AddColor "Blue", RGB(0, 0, 255)
//!         AddColor "Yellow", RGB(255, 255, 0)
//!         AddColor "Magenta", RGB(255, 0, 255)
//!         AddColor "Cyan", RGB(0, 255, 255)
//!         AddColor "White", RGB(255, 255, 255)
//!         AddColor "Black", RGB(0, 0, 0)
//!         AddColor "Gray", RGB(128, 128, 128)
//!         AddColor "Orange", RGB(255, 165, 0)
//!         AddColor "Purple", RGB(128, 0, 128)
//!         AddColor "Brown", RGB(165, 42, 42)
//!     End Sub
//!     
//!     Public Sub AddColor(name As String, color As Long)
//!         If m_count > UBound(m_colors) Then
//!             ReDim Preserve m_colors(0 To UBound(m_colors) + 50)
//!             ReDim Preserve m_colorNames(0 To UBound(m_colorNames) + 50)
//!         End If
//!         
//!         m_colorNames(m_count) = name
//!         m_colors(m_count) = color
//!         m_count = m_count + 1
//!     End Sub
//!     
//!     Public Function GetColor(name As String) As Long
//!         Dim i As Integer
//!         
//!         For i = 0 To m_count - 1
//!             If UCase(m_colorNames(i)) = UCase(name) Then
//!                 GetColor = m_colors(i)
//!                 Exit Function
//!             End If
//!         Next i
//!         
//!         GetColor = RGB(0, 0, 0)  ' Default to black
//!     End Function
//!     
//!     Public Function CreateGradientPalette(startColor As Long, endColor As Long, _
//!                                          steps As Integer) As Long()
//!         Dim palette() As Long
//!         Dim i As Integer
//!         Dim r1 As Integer, g1 As Integer, b1 As Integer
//!         Dim r2 As Integer, g2 As Integer, b2 As Integer
//!         Dim r As Integer, g As Integer, b As Integer
//!         Dim ratio As Double
//!         
//!         ReDim palette(0 To steps - 1)
//!         
//!         ' Extract components
//!         r1 = startColor And &HFF
//!         g1 = (startColor \ 256) And &HFF
//!         b1 = (startColor \ 65536) And &HFF
//!         
//!         r2 = endColor And &HFF
//!         g2 = (endColor \ 256) And &HFF
//!         b2 = (endColor \ 65536) And &HFF
//!         
//!         For i = 0 To steps - 1
//!             ratio = i / (steps - 1)
//!             
//!             r = r1 + ((r2 - r1) * ratio)
//!             g = g1 + ((g2 - g1) * ratio)
//!             b = b1 + ((b2 - b1) * ratio)
//!             
//!             palette(i) = RGB(r, g, b)
//!         Next i
//!         
//!         CreateGradientPalette = palette
//!     End Function
//!     
//!     Public Function GetColorCount() As Integer
//!         GetColorCount = m_count
//!     End Function
//!     
//!     Public Function GetColorByIndex(index As Integer) As Long
//!         If index >= 0 And index < m_count Then
//!             GetColorByIndex = m_colors(index)
//!         Else
//!             GetColorByIndex = RGB(0, 0, 0)
//!         End If
//!     End Function
//!     
//!     Public Function GetNameByIndex(index As Integer) As String
//!         If index >= 0 And index < m_count Then
//!             GetNameByIndex = m_colorNames(index)
//!         Else
//!             GetNameByIndex = ""
//!         End If
//!     End Function
//! End Class
//! ```
//!
//! ### Example 2: Heat Map Generator
//! ```vb
//! ' Generate heat map colors for data visualization
//! Module HeatMapColors
//!     Public Function GetHeatMapColor(value As Double, minValue As Double, _
//!                                    maxValue As Double) As Long
//!         ' Returns color from blue (cold) to red (hot)
//!         Dim ratio As Double
//!         Dim r As Integer, g As Integer, b As Integer
//!         
//!         ' Normalize value to 0-1 range
//!         If maxValue = minValue Then
//!             ratio = 0.5
//!         Else
//!             ratio = (value - minValue) / (maxValue - minValue)
//!         End If
//!         
//!         ' Clamp ratio to 0-1
//!         If ratio < 0 Then ratio = 0
//!         If ratio > 1 Then ratio = 1
//!         
//!         ' Calculate color based on ratio
//!         If ratio < 0.25 Then
//!             ' Blue to cyan
//!             r = 0
//!             g = ratio * 4 * 255
//!             b = 255
//!         ElseIf ratio < 0.5 Then
//!             ' Cyan to green
//!             r = 0
//!             g = 255
//!             b = 255 - ((ratio - 0.25) * 4 * 255)
//!         ElseIf ratio < 0.75 Then
//!             ' Green to yellow
//!             r = (ratio - 0.5) * 4 * 255
//!             g = 255
//!             b = 0
//!         Else
//!             ' Yellow to red
//!             r = 255
//!             g = 255 - ((ratio - 0.75) * 4 * 255)
//!             b = 0
//!         End If
//!         
//!         GetHeatMapColor = RGB(r, g, b)
//!     End Function
//!     
//!     Public Sub ApplyHeatMapToRange(dataRange() As Double, controls() As Control)
//!         Dim i As Integer
//!         Dim minVal As Double, maxVal As Double
//!         
//!         ' Find min and max values
//!         minVal = dataRange(LBound(dataRange))
//!         maxVal = dataRange(LBound(dataRange))
//!         
//!         For i = LBound(dataRange) To UBound(dataRange)
//!             If dataRange(i) < minVal Then minVal = dataRange(i)
//!             If dataRange(i) > maxVal Then maxVal = dataRange(i)
//!         Next i
//!         
//!         ' Apply colors
//!         For i = LBound(dataRange) To UBound(dataRange)
//!             controls(i).BackColor = GetHeatMapColor(dataRange(i), minVal, maxVal)
//!             controls(i).ForeColor = GetContrastColor(controls(i).BackColor)
//!         Next i
//!     End Sub
//!     
//!     Private Function GetContrastColor(bgColor As Long) As Long
//!         Dim r As Integer, g As Integer, b As Integer
//!         Dim luminance As Double
//!         
//!         r = bgColor And &HFF
//!         g = (bgColor \ 256) And &HFF
//!         b = (bgColor \ 65536) And &HFF
//!         
//!         luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
//!         
//!         If luminance > 0.5 Then
//!             GetContrastColor = RGB(0, 0, 0)
//!         Else
//!             GetContrastColor = RGB(255, 255, 255)
//!         End If
//!     End Function
//! End Module
//! ```
//!
//! ### Example 3: Theme Manager
//! ```vb
//! ' Application theme management with color schemes
//! Class ThemeManager
//!     Private Type ColorScheme
//!         Background As Long
//!         Foreground As Long
//!         Accent As Long
//!         Highlight As Long
//!         Border As Long
//!     End Type
//!     
//!     Private m_currentTheme As ColorScheme
//!     Private m_themeName As String
//!     
//!     Public Sub SetLightTheme()
//!         m_themeName = "Light"
//!         With m_currentTheme
//!             .Background = RGB(255, 255, 255)    ' White
//!             .Foreground = RGB(0, 0, 0)          ' Black
//!             .Accent = RGB(0, 120, 215)          ' Blue
//!             .Highlight = RGB(255, 255, 0)       ' Yellow
//!             .Border = RGB(192, 192, 192)        ' Light Gray
//!         End With
//!     End Sub
//!     
//!     Public Sub SetDarkTheme()
//!         m_themeName = "Dark"
//!         With m_currentTheme
//!             .Background = RGB(32, 32, 32)       ' Dark Gray
//!             .Foreground = RGB(255, 255, 255)    ' White
//!             .Accent = RGB(0, 120, 215)          ' Blue
//!             .Highlight = RGB(255, 200, 0)       ' Gold
//!             .Border = RGB(64, 64, 64)           ' Medium Gray
//!         End With
//!     End Sub
//!     
//!     Public Sub SetCustomTheme(bg As Long, fg As Long, accent As Long, _
//!                              highlight As Long, border As Long)
//!         m_themeName = "Custom"
//!         With m_currentTheme
//!             .Background = bg
//!             .Foreground = fg
//!             .Accent = accent
//!             .Highlight = highlight
//!             .Border = border
//!         End With
//!     End Sub
//!     
//!     Public Sub ApplyToForm(frm As Form)
//!         Dim ctrl As Control
//!         
//!         frm.BackColor = m_currentTheme.Background
//!         frm.ForeColor = m_currentTheme.Foreground
//!         
//!         For Each ctrl In frm.Controls
//!             ApplyToControl ctrl
//!         Next ctrl
//!     End Sub
//!     
//!     Private Sub ApplyToControl(ctrl As Control)
//!         On Error Resume Next
//!         
//!         ' Apply based on control type
//!         Select Case TypeName(ctrl)
//!             Case "TextBox"
//!                 ctrl.BackColor = m_currentTheme.Background
//!                 ctrl.ForeColor = m_currentTheme.Foreground
//!                 
//!             Case "CommandButton"
//!                 ctrl.BackColor = m_currentTheme.Accent
//!                 ctrl.ForeColor = RGB(255, 255, 255)
//!                 
//!             Case "Label"
//!                 ctrl.BackColor = m_currentTheme.Background
//!                 ctrl.ForeColor = m_currentTheme.Foreground
//!                 
//!             Case "ListBox", "ComboBox"
//!                 ctrl.BackColor = m_currentTheme.Background
//!                 ctrl.ForeColor = m_currentTheme.Foreground
//!         End Select
//!         
//!         On Error GoTo 0
//!     End Sub
//!     
//!     Public Function GetBackgroundColor() As Long
//!         GetBackgroundColor = m_currentTheme.Background
//!     End Function
//!     
//!     Public Function GetForegroundColor() As Long
//!         GetForegroundColor = m_currentTheme.Foreground
//!     End Function
//!     
//!     Public Function GetAccentColor() As Long
//!         GetAccentColor = m_currentTheme.Accent
//!     End Function
//!     
//!     Public Function GetThemeName() As String
//!         GetThemeName = m_themeName
//!     End Function
//! End Class
//! ```
//!
//! ### Example 4: Chart Color Generator
//! ```vb
//! ' Generate distinct colors for charts and graphs
//! Class ChartColorGenerator
//!     Private m_baseHue As Integer
//!     Private m_colorIndex As Integer
//!     
//!     Public Sub Initialize(Optional startHue As Integer = 0)
//!         m_baseHue = startHue
//!         m_colorIndex = 0
//!     End Sub
//!     
//!     Public Function GetNextColor() As Long
//!         Dim hue As Integer
//!         Dim r As Integer, g As Integer, b As Integer
//!         
//!         ' Use golden angle for good color distribution
//!         hue = (m_baseHue + (m_colorIndex * 137)) Mod 360
//!         m_colorIndex = m_colorIndex + 1
//!         
//!         GetNextColor = HSVToRGB(hue, 0.7, 0.9)
//!     End Function
//!     
//!     Private Function HSVToRGB(h As Double, s As Double, v As Double) As Long
//!         ' Convert HSV to RGB
//!         Dim r As Double, g As Double, b As Double
//!         Dim i As Integer
//!         Dim f As Double, p As Double, q As Double, t As Double
//!         
//!         If s = 0 Then
//!             r = v: g = v: b = v
//!         Else
//!             h = h / 60
//!             i = Int(h)
//!             f = h - i
//!             p = v * (1 - s)
//!             q = v * (1 - s * f)
//!             t = v * (1 - s * (1 - f))
//!             
//!             Select Case i
//!                 Case 0: r = v: g = t: b = p
//!                 Case 1: r = q: g = v: b = p
//!                 Case 2: r = p: g = v: b = t
//!                 Case 3: r = p: g = q: b = v
//!                 Case 4: r = t: g = p: b = v
//!                 Case Else: r = v: g = p: b = q
//!             End Select
//!         End If
//!         
//!         HSVToRGB = RGB(r * 255, g * 255, b * 255)
//!     End Function
//!     
//!     Public Sub Reset()
//!         m_colorIndex = 0
//!     End Sub
//!     
//!     Public Function GetColorArray(count As Integer) As Long()
//!         Dim colors() As Long
//!         Dim i As Integer
//!         
//!         ReDim colors(0 To count - 1)
//!         
//!         Reset
//!         For i = 0 To count - 1
//!             colors(i) = GetNextColor()
//!         Next i
//!         
//!         GetColorArray = colors
//!     End Function
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! The `RGB` function automatically handles out-of-range values:
//!
//! - Values greater than 255 are treated as 255 (saturated)
//! - Negative values are treated as 0
//! - Non-integer values are rounded to integers
//! - No error is raised for out-of-range values
//!
//! The function is very robust and rarely requires error handling:
//!
//! ```vb
//! ' RGB automatically clamps values to valid range
//! Dim color As Long
//! color = RGB(300, -50, 150)  ' Treated as RGB(255, 0, 150)
//! ```
//!
//! ## Performance Considerations
//!
//! - The `RGB` function is extremely fast - it's a simple calculation
//! - Can be called thousands of times with negligible performance impact
//! - No need to cache RGB values unless doing complex color calculations
//! - Extracting components from an RGB value is slightly slower than creating it
//!
//! ## Best Practices
//!
//! 1. **Use Named Constants**: Define color constants for reusability
//! 2. **Document Color Choices**: Comment why specific colors were chosen
//! 3. **Consider Accessibility**: Ensure sufficient contrast for readability
//! 4. **Test on Different Displays**: Colors may appear different on various monitors
//! 5. **Use Color Schemes**: Create coordinated palettes rather than random colors
//! 6. **Extract Components Carefully**: Remember BGR byte order when extracting
//! 7. **Validate User Input**: When accepting color values from users
//! 8. **Use Gradients Wisely**: Smooth gradients are more visually appealing
//! 9. **Consider Color Blindness**: Test with color-blind-friendly palettes
//! 10. **Avoid Magic Numbers**: Use `RGB()` rather than numeric color values
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | **RGB** | Create custom color | Long (RGB value) | Precise color control, custom colors |
//! | **`QBColor`** | Get `QBasic` color | Long (RGB value) | Legacy compatibility, 16-color palette |
//! | **vbRed, vbBlue, etc.** | Predefined constants | Long (RGB value) | Quick standard colors |
//! | **`LoadPicture`** | Load image | Picture object | Complex graphics, images |
//!
//! ## Platform and Version Notes
//!
//! - Available in all versions of VB6 and VBA
//! - Behavior consistent across all Windows platforms
//! - Returns BGR byte order (standard for Windows)
//! - Color values are compatible with Windows API functions
//! - Maximum value is &HFFFFFF (16,777,215 colors)
//!
//! ## Limitations
//!
//! - Limited to 24-bit color (16.7 million colors)
//! - No alpha channel support (transparency)
//! - Automatic clamping may mask input errors
//! - BGR byte order can be confusing when extracting components
//! - No built-in color space conversions (HSV, HSL, etc.)
//! - No color validation or named color lookup
//!
//! ## Related Functions
//!
//! - `QBColor`: Returns RGB value for `QBasic` color number (0-15)
//! - Predefined color constants: `vbBlack`, `vbRed`, `vbGreen`, `vbBlue`, `vbWhite`, etc.
//! - `LoadPicture`: Loads images with full color support
//! - `Point`: Returns the RGB color of a specified point on a form or picture box

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn rgb_basic() {
        let source = r#"
Dim color As Long
color = RGB(255, 0, 0)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_all_components() {
        let source = r#"
Dim customColor As Long
customColor = RGB(128, 64, 192)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_if_statement() {
        let source = r#"
If value > 100 Then
    cellColor = RGB(255, 0, 0)
Else
    cellColor = RGB(0, 255, 0)
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_function_return() {
        let source = r#"
Function GetRedColor() As Long
    GetRedColor = RGB(255, 0, 0)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_variable_assignment() {
        let source = r#"
Dim bgColor As Long
bgColor = RGB(red, green, blue)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_form_property() {
        let source = r#"
Form1.BackColor = RGB(200, 200, 200)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_debug_print() {
        let source = r#"
Debug.Print "Color: " & RGB(r, g, b)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_select_case() {
        let source = r#"
Select Case status
    Case "Error"
        color = RGB(255, 0, 0)
    Case "Warning"
        color = RGB(255, 255, 0)
    Case Else
        color = RGB(0, 255, 0)
End Select
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_class_usage() {
        let source = r#"
Private m_backgroundColor As Long

Public Sub SetColor()
    m_backgroundColor = RGB(255, 255, 255)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_with_statement() {
        let source = r#"
With Label1
    .BackColor = RGB(255, 255, 0)
    .ForeColor = RGB(0, 0, 0)
End With
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_elseif() {
        let source = r#"
If temp < 0 Then
    color = RGB(0, 0, 255)
ElseIf temp < 50 Then
    color = RGB(0, 255, 0)
Else
    color = RGB(255, 0, 0)
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_for_loop() {
        let source = r#"
For i = 0 To 255
    gradient(i) = RGB(i, i, i)
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_do_while() {
        let source = r#"
Do While r < 255
    colors(r) = RGB(r, 128, 64)
    r = r + 1
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_do_until() {
        let source = r#"
Do Until index > 10
    palette(index) = RGB(index * 25, 0, 0)
    index = index + 1
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_while_wend() {
        let source = r#"
While count < 100
    shades(count) = RGB(count, count, count)
    count = count + 1
Wend
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_parentheses() {
        let source = r#"
Dim result As Long
result = (RGB(255, 128, 0))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_iif() {
        let source = r#"
Dim textColor As Long
textColor = IIf(isActive, RGB(0, 0, 0), RGB(128, 128, 128))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_array_assignment() {
        let source = r#"
Dim colors(10) As Long
colors(i) = RGB(red(i), green(i), blue(i))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_property_assignment() {
        let source = r#"
Set obj = New ColorManager
obj.PrimaryColor = RGB(255, 0, 0)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_function_argument() {
        let source = r#"
Call SetBackgroundColor(RGB(230, 230, 230))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_concatenation() {
        let source = r#"
Dim msg As String
msg = "Color value: " & RGB(100, 150, 200)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_gradient_calculation() {
        let source = r#"
Dim gradientColor As Long
gradientColor = RGB(startR + (ratio * deltaR), startG, startB)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_comparison() {
        let source = r#"
If RGB(r1, g1, b1) = RGB(r2, g2, b2) Then
    MsgBox "Colors match"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_line_method() {
        let source = r#"
Picture1.Line (0, 0)-(100, 100), RGB(255, 0, 0), BF
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_msgbox() {
        let source = r#"
MsgBox "Color: " & Hex(RGB(255, 128, 64))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_nested_calculation() {
        let source = r#"
Dim blended As Long
blended = RGB((r1 + r2) \ 2, (g1 + g2) \ 2, (b1 + b2) \ 2)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rgb_on_error_goto() {
        let source = r#"
Sub SetColor()
    On Error GoTo ErrorHandler
    Dim c As Long
    c = RGB(redValue, greenValue, blueValue)
    Exit Sub
ErrorHandler:
    MsgBox "Error setting color"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("RGB"));
        assert!(text.contains("Identifier"));
    }
}
