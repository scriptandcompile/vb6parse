//! VB6 `Time` Function
//!
//! The `Time` function returns a Variant (Date) indicating the current system time.
//!
//! ## Syntax
//! ```vb6
//! Time()
//! ```
//! or
//! ```vb6
//! Time
//! ```
//!
//! ## Parameters
//! None. The `Time` function takes no arguments.
//!
//! ## Returns
//! Returns a `Variant` of subtype `Date` containing the current system time. The date portion is set to zero (December 30, 1899).
//!
//! ## Remarks
//! The `Time` function retrieves the current system time:
//!
//! - **No arguments**: Called without parentheses or with empty parentheses
//! - **Date portion**: Always returns date as zero (12/30/1899)
//! - **Time precision**: Returns time to the nearest second (no milliseconds)
//! - **System dependent**: Returns time from system clock
//! - **Time vs Time$**: `Time` returns Variant (Date), `Time$` returns String in format "HH:MM:SS"
//! - **Now function**: Use `Now()` to get current date and time together
//! - **Date function**: Use `Date()` to get current date only
//! - **Timer function**: Use `Timer` for elapsed time measurements (returns seconds since midnight)
//! - **Setting time**: Use `Time` statement (not function) to set system time
//! - **24-hour format**: Internal representation is 24-hour, but display depends on system settings
//!
//! ### Time vs Related Functions
//! - `Time` - Returns current time only (date portion is zero)
//! - `Date` - Returns current date only (time portion is midnight)
//! - `Now` - Returns current date and time together
//! - `Timer` - Returns seconds elapsed since midnight as Single
//! - `Time$` - Returns current time as formatted String
//!
//! ### Time Components
//! Extract time components using:
//! ```vb6
//! currentHour = Hour(Time)      ' 0-23
//! currentMinute = Minute(Time)  ' 0-59
//! currentSecond = Second(Time)  ' 0-59
//! ```
//!
//! ### Time Arithmetic
//! ```vb6
//! ' Add 1 hour to current time
//! newTime = Time + TimeSerial(1, 0, 0)
//!
//! ' Add 30 minutes to current time
//! newTime = DateAdd("n", 30, Time)
//! ```
//!
//! ## Typical Uses
//! 1. **Timestamp Logging**: Record when events occur
//! 2. **Time-based Triggers**: Check current time for scheduled operations
//! 3. **Time Display**: Show current time in user interface
//! 4. **Performance Timing**: Measure operation duration (though Timer is better)
//! 5. **Time Validation**: Check if operation is within allowed time window
//! 6. **Time Calculations**: Calculate time differences or future times
//! 7. **Scheduling**: Determine if tasks should run now
//! 8. **Time Formatting**: Create custom time displays
//!
//! ## Basic Examples
//!
//! ### Example 1: Display Current Time
//! ```vb6
//! Sub ShowCurrentTime()
//!     MsgBox "Current time is: " & Time
//! End Sub
//! ```
//!
//! ### Example 2: Log Event Time
//! ```vb6
//! Sub LogEvent(eventName As String)
//!     Dim logEntry As String
//!     logEntry = Time & " - " & eventName
//!     Debug.Print logEntry
//! End Sub
//! ```
//!
//! ### Example 3: Check Business Hours
//! ```vb6
//! Function IsBusinessHours() As Boolean
//!     Dim currentHour As Integer
//!     currentHour = Hour(Time)
//!     IsBusinessHours = (currentHour >= 9 And currentHour < 17)
//! End Function
//! ```
//!
//! ### Example 4: Format Time Display
//! ```vb6
//! Function GetFormattedTime() As String
//!     GetFormattedTime = Format$(Time, "hh:mm:ss AM/PM")
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Get Current Hour
//! ```vb6
//! Function GetCurrentHour() As Integer
//!     GetCurrentHour = Hour(Time)
//! End Function
//! ```
//!
//! ### Pattern 2: Time-based Greeting
//! ```vb6
//! Function GetGreeting() As String
//!     Dim currentHour As Integer
//!     currentHour = Hour(Time)
//!     
//!     Select Case currentHour
//!         Case 0 To 11
//!             GetGreeting = "Good morning"
//!         Case 12 To 17
//!             GetGreeting = "Good afternoon"
//!         Case Else
//!             GetGreeting = "Good evening"
//!     End Select
//! End Function
//! ```
//!
//! ### Pattern 3: Create Timestamp
//! ```vb6
//! Function CreateTimestamp() As String
//!     CreateTimestamp = Format$(Date, "yyyy-mm-dd") & " " & Format$(Time, "hh:nn:ss")
//! End Function
//! ```
//!
//! ### Pattern 4: Check Time Window
//! ```vb6
//! Function IsWithinTimeWindow(startTime As Date, endTime As Date) As Boolean
//!     Dim currentTime As Date
//!     currentTime = Time
//!     IsWithinTimeWindow = (currentTime >= startTime And currentTime <= endTime)
//! End Function
//! ```
//!
//! ### Pattern 5: Add Time Duration
//! ```vb6
//! Function AddMinutes(minutes As Integer) As Date
//!     AddMinutes = DateAdd("n", minutes, Time)
//! End Function
//! ```
//!
//! ### Pattern 6: Time Until Target
//! ```vb6
//! Function MinutesUntil(targetTime As Date) As Long
//!     MinutesUntil = DateDiff("n", Time, targetTime)
//! End Function
//! ```
//!
//! ### Pattern 7: Round Time to Nearest Interval
//! ```vb6
//! Function RoundToNearest15Minutes() As Date
//!     Dim currentTime As Date
//!     Dim minutes As Integer
//!     
//!     currentTime = Time
//!     minutes = Minute(currentTime)
//!     minutes = ((minutes + 7) \ 15) * 15
//!     
//!     RoundToNearest15Minutes = TimeSerial(Hour(currentTime), minutes, 0)
//! End Function
//! ```
//!
//! ### Pattern 8: Compare Times
//! ```vb6
//! Function IsTimeBefore(compareTime As Date) As Boolean
//!     IsTimeBefore = (Time < compareTime)
//! End Function
//! ```
//!
//! ### Pattern 9: Get Time String
//! ```vb6
//! Function GetTimeString() As String
//!     GetTimeString = Format$(Time, "hh:mm:ss")
//! End Function
//! ```
//!
//! ### Pattern 10: Calculate Elapsed Time
//! ```vb6
//! Function GetElapsedMinutes(startTime As Date) As Long
//!     GetElapsedMinutes = DateDiff("n", startTime, Time)
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Time Tracker Class
//! ```vb6
//! ' Class: TimeTracker
//! ' Tracks operation start/end times and durations
//! Option Explicit
//!
//! Private m_StartTime As Date
//! Private m_EndTime As Date
//! Private m_IsRunning As Boolean
//!
//! Public Sub StartTracking()
//!     m_StartTime = Time
//!     m_IsRunning = True
//! End Sub
//!
//! Public Sub StopTracking()
//!     If Not m_IsRunning Then
//!         Err.Raise 5, , "Tracking not started"
//!     End If
//!     
//!     m_EndTime = Time
//!     m_IsRunning = False
//! End Sub
//!
//! Public Function GetElapsedMinutes() As Long
//!     If m_IsRunning Then
//!         GetElapsedMinutes = DateDiff("n", m_StartTime, Time)
//!     Else
//!         GetElapsedMinutes = DateDiff("n", m_StartTime, m_EndTime)
//!     End If
//! End Function
//!
//! Public Function GetElapsedSeconds() As Long
//!     If m_IsRunning Then
//!         GetElapsedSeconds = DateDiff("s", m_StartTime, Time)
//!     Else
//!         GetElapsedSeconds = DateDiff("s", m_StartTime, m_EndTime)
//!     End If
//! End Function
//!
//! Public Function GetFormattedDuration() As String
//!     Dim totalSeconds As Long
//!     Dim hours As Long
//!     Dim minutes As Long
//!     Dim seconds As Long
//!     
//!     totalSeconds = GetElapsedSeconds()
//!     hours = totalSeconds \ 3600
//!     minutes = (totalSeconds Mod 3600) \ 60
//!     seconds = totalSeconds Mod 60
//!     
//!     GetFormattedDuration = Format$(hours, "00") & ":" & _
//!                           Format$(minutes, "00") & ":" & _
//!                           Format$(seconds, "00")
//! End Function
//!
//! Public Property Get IsRunning() As Boolean
//!     IsRunning = m_IsRunning
//! End Property
//! ```
//!
//! ### Example 2: Schedule Manager Module
//! ```vb6
//! ' Module: ScheduleManager
//! ' Manages time-based scheduling and windows
//! Option Explicit
//!
//! Public Function IsWithinSchedule(scheduleStart As Date, scheduleEnd As Date) As Boolean
//!     Dim currentTime As Date
//!     currentTime = Time
//!     
//!     ' Handle overnight schedules (e.g., 10 PM to 6 AM)
//!     If scheduleStart > scheduleEnd Then
//!         IsWithinSchedule = (currentTime >= scheduleStart Or currentTime <= scheduleEnd)
//!     Else
//!         IsWithinSchedule = (currentTime >= scheduleStart And currentTime <= scheduleEnd)
//!     End If
//! End Function
//!
//! Public Function GetNextScheduledTime(targetTime As Date) As Date
//!     Dim currentTime As Date
//!     currentTime = Time
//!     
//!     If currentTime < targetTime Then
//!         ' Target time is later today
//!         GetNextScheduledTime = Date + targetTime
//!     Else
//!         ' Target time is tomorrow
//!         GetNextScheduledTime = Date + 1 + targetTime
//!     End If
//! End Function
//!
//! Public Function MinutesUntilSchedule(scheduleTime As Date) As Long
//!     Dim currentTime As Date
//!     Dim targetDateTime As Date
//!     
//!     currentTime = Time
//!     
//!     If currentTime < scheduleTime Then
//!         ' Later today
//!         targetDateTime = Date + scheduleTime
//!     Else
//!         ' Tomorrow
//!         targetDateTime = Date + 1 + scheduleTime
//!     End If
//!     
//!     MinutesUntilSchedule = DateDiff("n", Now, targetDateTime)
//! End Function
//!
//! Public Function FormatTimeRemaining(targetTime As Date) As String
//!     Dim minutesLeft As Long
//!     Dim hours As Long
//!     Dim minutes As Long
//!     
//!     minutesLeft = MinutesUntilSchedule(targetTime)
//!     
//!     If minutesLeft < 0 Then
//!         FormatTimeRemaining = "Overdue"
//!         Exit Function
//!     End If
//!     
//!     hours = minutesLeft \ 60
//!     minutes = minutesLeft Mod 60
//!     
//!     If hours > 0 Then
//!         FormatTimeRemaining = hours & "h " & minutes & "m remaining"
//!     Else
//!         FormatTimeRemaining = minutes & "m remaining"
//!     End If
//! End Function
//! ```
//!
//! ### Example 3: Time Logger Class
//! ```vb6
//! ' Class: TimeLogger
//! ' Logs timestamped events
//! Option Explicit
//!
//! Private m_LogEntries As Collection
//!
//! Private Sub Class_Initialize()
//!     Set m_LogEntries = New Collection
//! End Sub
//!
//! Public Sub LogEvent(eventDescription As String)
//!     Dim logEntry As String
//!     logEntry = Format$(Time, "hh:mm:ss") & " - " & eventDescription
//!     m_LogEntries.Add logEntry
//! End Sub
//!
//! Public Sub LogEventWithDetails(eventName As String, details As String)
//!     Dim logEntry As String
//!     logEntry = Format$(Time, "hh:mm:ss") & " - " & eventName & ": " & details
//!     m_LogEntries.Add logEntry
//! End Sub
//!
//! Public Function GetLogEntries() As String
//!     Dim result As String
//!     Dim entry As Variant
//!     
//!     result = "Event Log:" & vbCrLf
//!     result = result & String$(50, "=") & vbCrLf
//!     
//!     For Each entry In m_LogEntries
//!         result = result & entry & vbCrLf
//!     Next entry
//!     
//!     GetLogEntries = result
//! End Function
//!
//! Public Sub ClearLog()
//!     Set m_LogEntries = New Collection
//! End Sub
//!
//! Public Property Get EntryCount() As Long
//!     EntryCount = m_LogEntries.Count
//! End Property
//!
//! Public Function ExportToFile(filename As String) As Boolean
//!     On Error GoTo ErrorHandler
//!     
//!     Dim fileNum As Integer
//!     Dim entry As Variant
//!     
//!     fileNum = FreeFile
//!     Open filename For Output As #fileNum
//!     
//!     Print #fileNum, "Event Log - " & Format$(Date, "yyyy-mm-dd")
//!     Print #fileNum, String$(50, "=")
//!     
//!     For Each entry In m_LogEntries
//!         Print #fileNum, entry
//!     Next entry
//!     
//!     Close #fileNum
//!     ExportToFile = True
//!     Exit Function
//!     
//! ErrorHandler:
//!     ExportToFile = False
//!     If fileNum > 0 Then Close #fileNum
//! End Function
//! ```
//!
//! ### Example 4: Business Hours Validator Module
//! ```vb6
//! ' Module: BusinessHoursValidator
//! ' Validates and manages business hours
//! Option Explicit
//!
//! Private m_BusinessStart As Date
//! Private m_BusinessEnd As Date
//! Private m_LunchStart As Date
//! Private m_LunchEnd As Date
//!
//! Public Sub Initialize(businessStart As Date, businessEnd As Date, _
//!                      Optional lunchStart As Variant, Optional lunchEnd As Variant)
//!     m_BusinessStart = businessStart
//!     m_BusinessEnd = businessEnd
//!     
//!     If Not IsMissing(lunchStart) Then m_LunchStart = lunchStart
//!     If Not IsMissing(lunchEnd) Then m_LunchEnd = lunchEnd
//! End Sub
//!
//! Public Function IsBusinessHours() As Boolean
//!     Dim currentTime As Date
//!     currentTime = Time
//!     
//!     ' Check if within business hours
//!     If currentTime < m_BusinessStart Or currentTime >= m_BusinessEnd Then
//!         IsBusinessHours = False
//!         Exit Function
//!     End If
//!     
//!     ' Check if during lunch (if configured)
//!     If m_LunchStart > 0 And m_LunchEnd > 0 Then
//!         If currentTime >= m_LunchStart And currentTime < m_LunchEnd Then
//!             IsBusinessHours = False
//!             Exit Function
//!         End If
//!     End If
//!     
//!     IsBusinessHours = True
//! End Function
//!
//! Public Function GetBusinessHoursStatus() As String
//!     Dim currentTime As Date
//!     currentTime = Time
//!     
//!     If currentTime < m_BusinessStart Then
//!         GetBusinessHoursStatus = "Before business hours (opens at " & _
//!                                 Format$(m_BusinessStart, "h:mm AM/PM") & ")"
//!     ElseIf currentTime >= m_BusinessEnd Then
//!         GetBusinessHoursStatus = "After business hours (closed at " & _
//!                                 Format$(m_BusinessEnd, "h:mm AM/PM") & ")"
//!     ElseIf m_LunchStart > 0 And currentTime >= m_LunchStart And currentTime < m_LunchEnd Then
//!         GetBusinessHoursStatus = "Lunch break (returns at " & _
//!                                 Format$(m_LunchEnd, "h:mm AM/PM") & ")"
//!     Else
//!         GetBusinessHoursStatus = "Open for business"
//!     End If
//! End Function
//!
//! Public Function MinutesUntilOpen() As Long
//!     Dim currentTime As Date
//!     currentTime = Time
//!     
//!     If currentTime < m_BusinessStart Then
//!         MinutesUntilOpen = DateDiff("n", currentTime, m_BusinessStart)
//!     Else
//!         ' Next business day
//!         Dim nextOpen As Date
//!         nextOpen = DateAdd("d", 1, Date) + m_BusinessStart
//!         MinutesUntilOpen = DateDiff("n", Now, nextOpen)
//!     End If
//! End Function
//!
//! Public Function MinutesUntilClose() As Long
//!     Dim currentTime As Date
//!     currentTime = Time
//!     
//!     If currentTime >= m_BusinessEnd Then
//!         MinutesUntilClose = 0
//!     Else
//!         MinutesUntilClose = DateDiff("n", currentTime, m_BusinessEnd)
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//! The `Time` function typically does not raise errors under normal circumstances:
//!
//! - **No parameters**: Cannot have parameter errors
//! - **System dependent**: Relies on system clock being set correctly
//! - **Always succeeds**: Returns current time from system
//!
//! ## Performance Notes
//! - Very fast operation - direct system call
//! - Negligible performance impact
//! - Can be called repeatedly without concern
//! - For high-precision timing, use `Timer` function instead
//! - Time resolution limited to one second
//!
//! ## Best Practices
//! 1. **Use Now for timestamps** if you need both date and time
//! 2. **Use Timer for performance** measurements (higher precision)
//! 3. **Cache Time value** if using multiple times in tight loop
//! 4. **Use Format$** to display time in specific format
//! 5. **Consider time zones** for distributed applications
//! 6. **Validate system time** is set correctly if critical
//! 7. **Use `TimeSerial`** to create specific time values for comparison
//! 8. **Handle overnight periods** carefully (when start > end time)
//! 9. **Store as Date type** not String for calculations
//! 10. **Document time format** assumptions in code comments
//!
//! ## Comparison Table
//!
//! | Function | Returns | Includes Date | Includes Time | Precision |
//! |----------|---------|---------------|---------------|-----------|
//! | `Time` | Variant (Date) | No (zero date) | Yes | 1 second |
//! | `Date` | Variant (Date) | Yes | No (midnight) | 1 day |
//! | `Now` | Variant (Date) | Yes | Yes | 1 second |
//! | `Timer` | Single | No | Yes (seconds since midnight) | High |
//! | `Time$` | String | No | Yes | 1 second |
//!
//! ## Platform Notes
//! - Available in VB6, VBA, and `VBScript`
//! - Consistent behavior across platforms
//! - Returns time based on system clock
//! - Time format display depends on regional settings
//! - Internal representation is numeric (fraction of a day)
//! - Date portion always 12/30/1899 (zero date)
//!
//! ## Limitations
//! - No millisecond precision (use Timer for better precision)
//! - Cannot get time from different time zone
//! - No built-in UTC time support
//! - Cannot specify which clock source to use
//! - Display format depends on system locale settings
//! - Cannot return time as numeric value directly (use `CDbl` or cast)
//! - No built-in support for daylight saving time handling

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn time_basic() {
        let source = r#"
Sub Test()
    currentTime = Time
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_with_parentheses() {
        let source = r#"
Sub Test()
    currentTime = Time()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_variable_assignment() {
        let source = r#"
Sub Test()
    Dim t As Date
    t = Time
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Current time: " & Time
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print Time
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_function_return() {
        let source = r#"
Function GetCurrentTime() As Date
    GetCurrentTime = Time
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_with_hour() {
        let source = r#"
Sub Test()
    currentHour = Hour(Time)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_comparison() {
        let source = r#"
Sub Test()
    If Time > startTime Then
        DoSomething
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_format() {
        let source = r#"
Sub Test()
    formatted = Format$(Time, "hh:mm:ss")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_dateadd() {
        let source = r#"
Sub Test()
    futureTime = DateAdd("n", 30, Time)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_datediff() {
        let source = r#"
Sub Test()
    elapsed = DateDiff("n", startTime, Time)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_if_statement() {
        let source = r#"
Sub Test()
    If Hour(Time) >= 9 And Hour(Time) < 17 Then
        BusinessHours
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_select_case() {
        let source = r#"
Sub Test()
    Select Case Hour(Time)
        Case 0 To 11
            Morning
        Case 12 To 17
            Afternoon
        Case Else
            Evening
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_function_argument() {
        let source = r#"
Sub Test()
    Call LogTime(Time)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_property_assignment() {
        let source = r#"
Sub Test()
    obj.CurrentTime = Time
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_with_statement() {
        let source = r#"
Sub Test()
    With tracker
        .StartTime = Time
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_concatenation() {
        let source = r#"
Sub Test()
    logEntry = Time & " - Event occurred"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_array_assignment() {
        let source = r#"
Sub Test()
    timestamps(i) = Time
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_print_statement() {
        let source = r#"
Sub Test()
    Print #1, Time
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_elseif() {
        let source = r#"
Sub Test()
    If x = 1 Then
        y = 1
    ElseIf Time > deadline Then
        y = 2
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_do_while() {
        let source = r#"
Sub Test()
    Do While Time < endTime
        Process
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_do_until() {
        let source = r#"
Sub Test()
    Do Until Time >= targetTime
        Wait
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_while_wend() {
        let source = r#"
Sub Test()
    While Time < stopTime
        Continue
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_iif() {
        let source = r#"
Sub Test()
    message = IIf(Time > noon, "PM", "AM")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_arithmetic() {
        let source = r#"
Sub Test()
    futureTime = Time + TimeSerial(1, 0, 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_cdate() {
        let source = r#"
Sub Test()
    timeValue = CDate(Time)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_class_usage() {
        let source = r#"
Sub Test()
    Set logger = New TimeLogger
    logger.LogTime Time
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }

    #[test]
    fn time_minute() {
        let source = r#"
Sub Test()
    currentMinute = Minute(Time)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Time"));
    }
}
