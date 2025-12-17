//! # `FormatDateTime` Function
//!
//! Returns an expression formatted as a date or time.
//!
//! ## Syntax
//!
//! ```vb
//! FormatDateTime(date[, namedformat])
//! ```
//!
//! ## Parameters
//!
//! - **date**: Required. Date expression to be formatted.
//! - **namedformat**: Optional. `Numeric` value that indicates the date/time format used. If omitted, `vbGeneralDate` is used.
//!
//! ## Named Format Constants
//!
//! - **`vbGeneralDate`** (0): Display a date and/or time. For real numbers, display a date and time. If there is no fractional part, display only a date. If there is no integer part, display time only. Date and time display is determined by system settings.
//! - **`vbLongDate`** (1): Display a date using the long date format specified in the computer's regional settings.
//! - **`vbShortDate`** (2): Display a date using the short date format specified in the computer's regional settings.
//! - **`vbLongTime`** (3): Display a time using the time format specified in the computer's regional settings.
//! - **`vbShortTime`** (4): Display a time using the 24-hour format (hh:mm).
//!
//! ## Return Value
//!
//! Returns a Variant of subtype String containing the formatted date/time expression.
//!
//! ## Remarks
//!
//! The `FormatDateTime` function provides a simple way to format date and time values
//! using predefined formats based on the system's regional settings. It's easier to use
//! than Format for common date/time formatting needs.
//!
//! **Important Characteristics:**
//!
//! - Uses system locale for date/time formatting
//! - Only supports 5 predefined formats
//! - Less flexible than Format function
//! - Locale-aware (respects regional settings)
//! - Returns empty string if date is Null
//! - Date-only values show date without time
//! - Time-only values show time without date
//! - Format determined by Windows regional settings
//! - Cannot customize format patterns
//! - Easier for simple date/time display
//!
//! ## Typical Uses
//!
//! - Display dates in user's preferred format
//! - Show timestamps in applications
//! - Format dates for reports
//! - Display file modification times
//! - Show appointment times
//! - Format log timestamps
//! - Display birth dates
//! - Show transaction dates
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! Dim dt As Date
//! dt = #1/15/2025 3:45:30 PM#
//!
//! ' General date (default)
//! Debug.Print FormatDateTime(dt)                    ' 1/15/2025 3:45:30 PM
//! Debug.Print FormatDateTime(dt, vbGeneralDate)     ' 1/15/2025 3:45:30 PM
//!
//! ' Long date
//! Debug.Print FormatDateTime(dt, vbLongDate)        ' Wednesday, January 15, 2025
//!
//! ' Short date
//! Debug.Print FormatDateTime(dt, vbShortDate)       ' 1/15/2025
//!
//! ' Long time
//! Debug.Print FormatDateTime(dt, vbLongTime)        ' 3:45:30 PM
//!
//! ' Short time
//! Debug.Print FormatDateTime(dt, vbShortTime)       ' 15:45
//! ```
//!
//! ### Date Only
//!
//! ```vb
//! Dim dateOnly As Date
//! dateOnly = #1/15/2025#
//!
//! Debug.Print FormatDateTime(dateOnly, vbGeneralDate)  ' 1/15/2025
//! Debug.Print FormatDateTime(dateOnly, vbLongDate)     ' Wednesday, January 15, 2025
//! Debug.Print FormatDateTime(dateOnly, vbShortDate)    ' 1/15/2025
//! ```
//!
//! ### Time Only
//!
//! ```vb
//! Dim timeOnly As Date
//! timeOnly = #3:45:30 PM#
//!
//! Debug.Print FormatDateTime(timeOnly, vbGeneralDate)  ' 3:45:30 PM
//! Debug.Print FormatDateTime(timeOnly, vbLongTime)     ' 3:45:30 PM
//! Debug.Print FormatDateTime(timeOnly, vbShortTime)    ' 15:45
//! ```
//!
//! ## Common Patterns
//!
//! ### Display Current Date/Time
//!
//! ```vb
//! Sub ShowCurrentDateTime()
//!     Debug.Print "Current Date (Long): " & FormatDateTime(Now, vbLongDate)
//!     Debug.Print "Current Date (Short): " & FormatDateTime(Now, vbShortDate)
//!     Debug.Print "Current Time (Long): " & FormatDateTime(Now, vbLongTime)
//!     Debug.Print "Current Time (Short): " & FormatDateTime(Now, vbShortTime)
//! End Sub
//! ```
//!
//! ### Format Label Caption
//!
//! ```vb
//! Sub UpdateDateLabel()
//!     lblCurrentDate.Caption = "Today is " & FormatDateTime(Date, vbLongDate)
//!     lblCurrentTime.Caption = "Time: " & FormatDateTime(Time, vbLongTime)
//! End Sub
//! ```
//!
//! ### Log Entry Formatting
//!
//! ```vb
//! Sub WriteLog(message As String)
//!     Dim logEntry As String
//!     Dim timestamp As String
//!     
//!     timestamp = FormatDateTime(Now, vbGeneralDate)
//!     logEntry = "[" & timestamp & "] " & message
//!     
//!     Debug.Print logEntry
//!     ' Or write to file
//! End Sub
//! ```
//!
//! ### Display File Information
//!
//! ```vb
//! Sub ShowFileInfo(filePath As String)
//!     Dim fileDate As Date
//!     
//!     On Error GoTo ErrorHandler
//!     fileDate = FileDateTime(filePath)
//!     
//!     Debug.Print "File: " & filePath
//!     Debug.Print "Modified: " & FormatDateTime(fileDate, vbLongDate) & _
//!                 " at " & FormatDateTime(fileDate, vbLongTime)
//!     Exit Sub
//!     
//! ErrorHandler:
//!     Debug.Print "Error reading file date"
//! End Sub
//! ```
//!
//! ### Format `ListBox` Items
//!
//! ```vb
//! Sub PopulateDateList(lstDates As ListBox, dates() As Date)
//!     Dim i As Long
//!     
//!     lstDates.Clear
//!     
//!     For i = LBound(dates) To UBound(dates)
//!         lstDates.AddItem FormatDateTime(dates(i), vbShortDate)
//!     Next i
//! End Sub
//! ```
//!
//! ### Display Birthday
//!
//! ```vb
//! Function FormatBirthday(birthDate As Date) As String
//!     Dim age As Integer
//!     
//!     age = DateDiff("yyyy", birthDate, Date)
//!     FormatBirthday = FormatDateTime(birthDate, vbLongDate) & " (Age: " & age & ")"
//! End Function
//!
//! ' Usage
//! Debug.Print FormatBirthday(#3/15/1990#)
//! ```
//!
//! ### Format Appointment Display
//!
//! ```vb
//! Function FormatAppointment(appointmentDate As Date, description As String) As String
//!     Dim dateStr As String
//!     Dim timeStr As String
//!     
//!     dateStr = FormatDateTime(appointmentDate, vbLongDate)
//!     timeStr = FormatDateTime(appointmentDate, vbShortTime)
//!     
//!     FormatAppointment = description & vbCrLf & _
//!                         "Date: " & dateStr & vbCrLf & _
//!                         "Time: " & timeStr
//! End Function
//! ```
//!
//! ### Display Relative Time
//!
//! ```vb
//! Function FormatRelativeDate(dt As Date) As String
//!     Dim daysDiff As Long
//!     
//!     daysDiff = DateDiff("d", dt, Date)
//!     
//!     Select Case daysDiff
//!         Case 0
//!             FormatRelativeDate = "Today at " & FormatDateTime(dt, vbShortTime)
//!         Case 1
//!             FormatRelativeDate = "Yesterday at " & FormatDateTime(dt, vbShortTime)
//!         Case -1
//!             FormatRelativeDate = "Tomorrow at " & FormatDateTime(dt, vbShortTime)
//!         Case Is > 1
//!             FormatRelativeDate = FormatDateTime(dt, vbShortDate)
//!         Case Is < -1
//!             FormatRelativeDate = FormatDateTime(dt, vbShortDate)
//!     End Select
//! End Function
//! ```
//!
//! ### Format Database Display
//!
//! ```vb
//! Function GetFormattedDate(rs As ADODB.Recordset, fieldName As String) As String
//!     If IsNull(rs.Fields(fieldName).Value) Then
//!         GetFormattedDate = "N/A"
//!     Else
//!         GetFormattedDate = FormatDateTime(rs.Fields(fieldName).Value, vbShortDate)
//!     End If
//! End Function
//! ```
//!
//! ### Create Date Range Display
//!
//! ```vb
//! Function FormatDateRange(startDate As Date, endDate As Date) As String
//!     FormatDateRange = FormatDateTime(startDate, vbShortDate) & " - " & _
//!                       FormatDateTime(endDate, vbShortDate)
//! End Function
//!
//! ' Usage
//! Debug.Print FormatDateRange(#1/1/2025#, #1/31/2025#)  ' 1/1/2025 - 1/31/2025
//! ```
//!
//! ### Format Grid Display
//!
//! ```vb
//! Sub PopulateTransactionGrid(grid As MSFlexGrid, transactions As Collection)
//!     Dim i As Long
//!     Dim trans As Variant
//!     
//!     grid.Clear
//!     grid.Rows = transactions.Count + 1
//!     
//!     ' Headers
//!     grid.TextMatrix(0, 0) = "Date"
//!     grid.TextMatrix(0, 1) = "Time"
//!     grid.TextMatrix(0, 2) = "Description"
//!     
//!     i = 1
//!     For Each trans In transactions
//!         grid.TextMatrix(i, 0) = FormatDateTime(trans.TransDate, vbShortDate)
//!         grid.TextMatrix(i, 1) = FormatDateTime(trans.TransDate, vbShortTime)
//!         grid.TextMatrix(i, 2) = trans.Description
//!         i = i + 1
//!     Next trans
//! End Sub
//! ```
//!
//! ## Advanced Usage
//!
//! ### Flexible Date Formatter
//!
//! ```vb
//! Function FormatDateEx(dt As Date, style As String) As String
//!     Select Case LCase(style)
//!         Case "long"
//!             FormatDateEx = FormatDateTime(dt, vbLongDate)
//!         Case "short"
//!             FormatDateEx = FormatDateTime(dt, vbShortDate)
//!         Case "longtime"
//!             FormatDateEx = FormatDateTime(dt, vbLongTime)
//!         Case "shorttime"
//!             FormatDateEx = FormatDateTime(dt, vbShortTime)
//!         Case "full"
//!             FormatDateEx = FormatDateTime(dt, vbLongDate) & " " & _
//!                            FormatDateTime(dt, vbLongTime)
//!         Case "compact"
//!             FormatDateEx = FormatDateTime(dt, vbShortDate) & " " & _
//!                            FormatDateTime(dt, vbShortTime)
//!         Case Else
//!             FormatDateEx = FormatDateTime(dt, vbGeneralDate)
//!     End Select
//! End Function
//! ```
//!
//! ### Multi-Format Display
//!
//! ```vb
//! Sub DisplayAllFormats(dt As Date)
//!     Debug.Print "Date/Time: " & dt
//!     Debug.Print String(60, "=")
//!     Debug.Print "General Date:  ", FormatDateTime(dt, vbGeneralDate)
//!     Debug.Print "Long Date:     ", FormatDateTime(dt, vbLongDate)
//!     Debug.Print "Short Date:    ", FormatDateTime(dt, vbShortDate)
//!     Debug.Print "Long Time:     ", FormatDateTime(dt, vbLongTime)
//!     Debug.Print "Short Time:    ", FormatDateTime(dt, vbShortTime)
//! End Sub
//! ```
//!
//! ### Calendar Event Formatter
//!
//! ```vb
//! Type CalendarEvent
//!     Title As String
//!     StartTime As Date
//!     EndTime As Date
//!     AllDay As Boolean
//! End Type
//!
//! Function FormatCalendarEvent(evt As CalendarEvent) As String
//!     Dim result As String
//!     
//!     result = evt.Title & vbCrLf
//!     result = result & FormatDateTime(evt.StartTime, vbLongDate) & vbCrLf
//!     
//!     If evt.AllDay Then
//!         result = result & "All Day Event"
//!     Else
//!         result = result & FormatDateTime(evt.StartTime, vbShortTime) & " - " & _
//!                  FormatDateTime(evt.EndTime, vbShortTime)
//!     End If
//!     
//!     FormatCalendarEvent = result
//! End Function
//! ```
//!
//! ### Report Header Generator
//!
//! ```vb
//! Function GenerateReportHeader(reportTitle As String) As String
//!     Dim header As String
//!     
//!     header = String(60, "=") & vbCrLf
//!     header = header & reportTitle & vbCrLf
//!     header = header & "Generated: " & FormatDateTime(Now, vbLongDate) & vbCrLf
//!     header = header & "Time: " & FormatDateTime(Now, vbLongTime) & vbCrLf
//!     header = header & String(60, "=") & vbCrLf
//!     
//!     GenerateReportHeader = header
//! End Function
//! ```
//!
//! ### Timestamp Logger
//!
//! ```vb
//! Class TimestampLogger
//!     Private logEntries As Collection
//!     
//!     Private Sub Class_Initialize()
//!         Set logEntries = New Collection
//!     End Sub
//!     
//!     Public Sub LogEvent(message As String, Optional includeTime As Boolean = True)
//!         Dim entry As String
//!         
//!         If includeTime Then
//!             entry = FormatDateTime(Now, vbShortDate) & " " & _
//!                     FormatDateTime(Now, vbLongTime) & " - " & message
//!         Else
//!             entry = FormatDateTime(Now, vbShortDate) & " - " & message
//!         End If
//!         
//!         logEntries.Add entry
//!     End Sub
//!     
//!     Public Function GetLog() As String
//!         Dim entry As Variant
//!         Dim result As String
//!         
//!         For Each entry In logEntries
//!             result = result & entry & vbCrLf
//!         Next entry
//!         
//!         GetLog = result
//!     End Function
//! End Class
//! ```
//!
//! ### Smart Date Display
//!
//! ```vb
//! Function SmartFormatDate(dt As Date, Optional includeTime As Boolean = False) As String
//!     Dim today As Date
//!     Dim diff As Long
//!     
//!     today = Date
//!     diff = DateDiff("d", today, dt)
//!     
//!     If diff = 0 Then
//!         SmartFormatDate = "Today"
//!     ElseIf diff = -1 Then
//!         SmartFormatDate = "Yesterday"
//!     ElseIf diff = 1 Then
//!         SmartFormatDate = "Tomorrow"
//!     ElseIf diff > -7 And diff < 0 Then
//!         SmartFormatDate = Format(dt, "dddd")  ' Day name
//!     Else
//!         SmartFormatDate = FormatDateTime(dt, vbShortDate)
//!     End If
//!     
//!     If includeTime Then
//!         SmartFormatDate = SmartFormatDate & " at " & _
//!                          FormatDateTime(dt, vbShortTime)
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeFormatDateTime(value As Variant, _
//!                             Optional style As VbDateTimeFormat = vbGeneralDate) As String
//!     On Error GoTo ErrorHandler
//!     
//!     If IsNull(value) Then
//!         SafeFormatDateTime = "N/A"
//!     ElseIf IsDate(value) Then
//!         SafeFormatDateTime = FormatDateTime(CDate(value), style)
//!     Else
//!         SafeFormatDateTime = "Invalid Date"
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 13  ' Type mismatch
//!             SafeFormatDateTime = "Type Error"
//!         Case 5   ' Invalid procedure call
//!             SafeFormatDateTime = "Invalid Format"
//!         Case Else
//!             SafeFormatDateTime = "Error"
//!     End Select
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 13** (Type Mismatch): Value is not a valid date
//! - **Error 5** (Invalid procedure call): Invalid format constant
//!
//! ## Performance Considerations
//!
//! - `FormatDateTime` is fast for predefined formats
//! - Faster than `Format` for simple date/time display
//! - Locale lookups cached by system
//! - Minimal overhead for formatting
//! - Consider caching formatted strings for repeated display
//!
//! ## Best Practices
//!
//! ### Use `FormatDateTime` for Simple Formatting
//!
//! ```vb
//! ' Good - Simple and locale-aware
//! formatted = FormatDateTime(dt, vbShortDate)
//!
//! ' Overkill for simple display
//! formatted = Format(dt, "mm/dd/yyyy")
//! ```
//!
//! ### Handle Null Values
//!
//! ```vb
//! ' Good - Check for Null
//! If Not IsNull(dateValue) Then
//!     formatted = FormatDateTime(dateValue, vbShortDate)
//! Else
//!     formatted = "N/A"
//! End If
//! ```
//!
//! ### Use Appropriate Format for Context
//!
//! ```vb
//! ' Good - Long date for formal display
//! lblEventDate.Caption = FormatDateTime(eventDate, vbLongDate)
//!
//! ' Good - Short date for compact display
//! txtDate.Text = FormatDateTime(Date, vbShortDate)
//! ```
//!
//! ## Comparison with Other Functions
//!
//! ### `FormatDateTime` vs `Format`
//!
//! ```vb
//! ' `FormatDateTime` - Predefined formats only
//! result = FormatDateTime(Now, vbLongDate)
//!
//! ' `Format` - Custom patterns possible
//! result = Format(Now, "dddd, mmmm d, yyyy")
//! ```
//!
//! ### `FormatDateTime` vs `DatePart`
//!
//! ```vb
//! ' FormatDateTime - Returns formatted string
//! result = FormatDateTime(Now, vbShortDate)  ' "1/15/2025"
//!
//! ' DatePart - Returns numeric part
//! result = DatePart("m", Now)                ' 1
//! ```
//!
//! ### `FormatDateTime` vs `CStr`/`Str`
//!
//! ```vb
//! ' FormatDateTime - Locale-aware formatting
//! result = FormatDateTime(Now, vbShortDate)  ' "1/15/2025"
//!
//! ' CStr - Default conversion
//! result = CStr(Now)                         ' "1/15/2025 3:45:30 PM"
//! ```
//!
//! ## Limitations
//!
//! - Only 5 predefined formats available
//! - Cannot customize format patterns
//! - Uses system locale (cannot specify different locale)
//! - Less flexible than `Format` function
//! - No control over date/time separators
//! - Cannot specify culture-specific formats
//! - `Format` style depends on Windows settings
//!
//! ## Regional Settings Impact
//!
//! The `FormatDateTime` function behavior varies by locale:
//!
//! ### vbShortDate
//! - **United States**: 1/15/2025
//! - **United Kingdom**: 15/01/2025
//! - **Germany**: 15.01.2025
//! - **Japan**: 2025/01/15
//!
//! ### vbLongDate
//! - **United States**: Wednesday, January 15, 2025
//! - **France**: mercredi 15 janvier 2025
//! - **Germany**: Mittwoch, 15. Januar 2025
//!
//! ### vbShortTime
//! - **Most locales**: 15:45 (24-hour format)
//!
//! ### vbLongTime
//! - **United States**: 3:45:30 PM
//! - **Europe**: 15:45:30
//!
//! ## Related Functions
//!
//! - `Format`: More flexible date/time formatting with custom patterns
//! - `FormatCurrency`: Format numbers as currency
//! - `FormatNumber`: Format numbers without currency symbol
//! - `FormatPercent`: Format numbers as percentages
//! - `DatePart`: Extract specific parts of a date
//! - `DateDiff`: Calculate difference between dates
//! - `CDate`: Convert expression to Date type
//! - `IsDate`: Check if expression can be converted to Date

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn formatdatetime_basic() {
        let source = r#"
result = FormatDateTime(dt)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_general() {
        let source = r#"
formatted = FormatDateTime(dt, vbGeneralDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_longdate() {
        let source = r#"
result = FormatDateTime(dt, vbLongDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_shortdate() {
        let source = r#"
result = FormatDateTime(dt, vbShortDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_longtime() {
        let source = r#"
result = FormatDateTime(dt, vbLongTime)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_shorttime() {
        let source = r#"
result = FormatDateTime(dt, vbShortTime)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_now() {
        let source = r#"
current = FormatDateTime(Now, vbShortDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_date() {
        let source = r#"
today = FormatDateTime(Date, vbLongDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_time() {
        let source = r#"
currentTime = FormatDateTime(Time, vbLongTime)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_debug_print() {
        let source = r#"
Debug.Print FormatDateTime(Now, vbLongDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_concatenation() {
        let source = r#"
msg = "Today is " & FormatDateTime(Date, vbLongDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_label_caption() {
        let source = r#"
lblCurrentDate.Caption = FormatDateTime(Date, vbLongDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_log_entry() {
        let source = r#"
timestamp = FormatDateTime(Now, vbGeneralDate)
logEntry = "[" & timestamp & "] " & message
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_filedatetime() {
        let source = r#"
fileDate = FileDateTime(filePath)
formatted = FormatDateTime(fileDate, vbLongDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_listbox() {
        let source = r#"
lstDates.AddItem FormatDateTime(dates(i), vbShortDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_multiline() {
        let source = r#"
result = "Date: " & FormatDateTime(appointmentDate, vbLongDate) & vbCrLf & _
         "Time: " & FormatDateTime(appointmentDate, vbShortTime)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_select_case() {
        let source = r#"
Select Case style
    Case "long"
        result = FormatDateTime(dt, vbLongDate)
    Case "short"
        result = FormatDateTime(dt, vbShortDate)
End Select
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_if_statement() {
        let source = r#"
If includeTime Then
    result = FormatDateTime(dt, vbGeneralDate)
Else
    result = FormatDateTime(dt, vbShortDate)
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_isnull_check() {
        let source = r#"
If Not IsNull(dateValue) Then
    formatted = FormatDateTime(dateValue, vbShortDate)
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_range() {
        let source = r#"
range = FormatDateTime(startDate, vbShortDate) & " - " & FormatDateTime(endDate, vbShortDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_grid() {
        let source = r#"
grid.TextMatrix(i, 0) = FormatDateTime(trans.TransDate, vbShortDate)
grid.TextMatrix(i, 1) = FormatDateTime(trans.TransDate, vbShortTime)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_for_loop() {
        let source = r#"
For i = LBound(dates) To UBound(dates)
    Debug.Print FormatDateTime(dates(i), vbShortDate)
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_error_handling() {
        let source = r#"
On Error GoTo ErrorHandler
formatted = FormatDateTime(CDate(value), style)
Exit Function
ErrorHandler:
    formatted = "Error"
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_isdate_check() {
        let source = r#"
If IsDate(value) Then
    result = FormatDateTime(CDate(value), vbShortDate)
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_recordset() {
        let source = r#"
formatted = FormatDateTime(rs.Fields("OrderDate").Value, vbShortDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn formatdatetime_function_return() {
        let source = r#"
Function FormatBirthday(birthDate As Date) As String
    FormatBirthday = FormatDateTime(birthDate, vbLongDate)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("FormatDateTime"));
        assert!(debug.contains("Identifier"));
    }
}
