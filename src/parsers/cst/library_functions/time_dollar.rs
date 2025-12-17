//! # `Time$` Function
//!
//! The `Time$` function in Visual Basic 6 returns a string representing the current system time.
//! The dollar sign (`$`) suffix indicates that this function always returns a `String` type,
//! never a `Variant`.
//!
//! ## Syntax
//!
//! ```vb6
//! Time$
//! ```
//!
//! ## Parameters
//!
//! None. `Time$` takes no parameters.
//!
//! ## Return Value
//!
//! Returns a `String` representing the current system time in the format "HH:MM:SS" (24-hour format).
//!
//! ## Behavior and Characteristics
//!
//! ### Time Format
//!
//! - Always returns 24-hour format (e.g., "14:30:45" for 2:30:45 PM)
//! - Format: "HH:MM:SS" where HH is 00-23, MM is 00-59, SS is 00-59
//! - Always includes leading zeros (e.g., "09:05:03")
//! - Does not include AM/PM indicator
//! - Does not include milliseconds or fractional seconds
//!
//! ### Type Differences: `Time$` vs `Time`
//!
//! - `Time$`: Always returns `String` type (never `Variant`)
//! - `Time`: Returns `Variant` containing a Date/Time value
//! - Use `Time$` when you need a string representation
//! - Use `Time` when you need to perform date/time arithmetic
//!
//! ### System Time
//!
//! - Reflects the current system clock time
//! - Updates each time the function is called
//! - Accuracy depends on system clock resolution
//! - No time zone information included
//!
//! ## Common Usage Patterns
//!
//! ### 1. Display Current Time
//!
//! ```vb6
//! Sub ShowTime()
//!     Debug.Print "Current time: " & Time$
//! End Sub
//! ```
//!
//! ### 2. Timestamp for Logging
//!
//! ```vb6
//! Sub LogMessage(message As String)
//!     Dim logEntry As String
//!     logEntry = Time$ & " - " & message
//!     Debug.Print logEntry
//! End Sub
//!
//! LogMessage "Application started"
//! ```
//!
//! ### 3. Time-Based File Naming
//!
//! ```vb6
//! Function GenerateTimeStampedFileName(baseName As String) As String
//!     Dim timeStamp As String
//!     timeStamp = Replace$(Time$, ":", "")
//!     GenerateTimeStampedFileName = baseName & "_" & timeStamp & ".log"
//! End Function
//!
//! ' Generates: "logfile_143045.log" at 2:30:45 PM
//! ```
//!
//! ### 4. Update Time Display
//!
//! ```vb6
//! Sub Timer1_Timer()
//!     ' Update label every second
//!     lblTime.Caption = Time$
//! End Sub
//! ```
//!
//! ### 5. Record Processing Time
//!
//! ```vb6
//! Sub ProcessData()
//!     Dim startTime As String
//!     startTime = Time$
//!     
//!     ' ... processing code ...
//!     
//!     Debug.Print "Started at: " & startTime
//!     Debug.Print "Completed at: " & Time$
//! End Sub
//! ```
//!
//! ### 6. Time-Based Greetings
//!
//! ```vb6
//! Function GetGreeting() As String
//!     Dim currentHour As Integer
//!     currentHour = Hour(Now)
//!     
//!     If currentHour < 12 Then
//!         GetGreeting = "Good morning! Time: " & Time$
//!     ElseIf currentHour < 18 Then
//!         GetGreeting = "Good afternoon! Time: " & Time$
//!     Else
//!         GetGreeting = "Good evening! Time: " & Time$
//!     End If
//! End Function
//! ```
//!
//! ### 7. Audit Trail Entries
//!
//! ```vb6
//! Sub RecordAudit(action As String, userName As String)
//!     Dim auditEntry As String
//!     auditEntry = Date$ & " " & Time$ & " - " & userName & " - " & action
//!     Print #1, auditEntry
//! End Sub
//! ```
//!
//! ### 8. Periodic Task Checking
//!
//! ```vb6
//! Sub CheckScheduledTask()
//!     Dim currentTimeStr As String
//!     currentTimeStr = Time$
//!     
//!     If currentTimeStr = "09:00:00" Then
//!         ' Execute morning task
//!         RunMorningReport
//!     End If
//! End Sub
//! ```
//!
//! ### 9. Status Bar Updates
//!
//! ```vb6
//! Sub UpdateStatusBar()
//!     StatusBar1.Panels(1).Text = "Current Time: " & Time$
//! End Sub
//! ```
//!
//! ### 10. Debug Output with Timestamps
//!
//! ```vb6
//! Sub DebugLog(category As String, message As String)
//!     Debug.Print "[" & Time$ & "] " & category & ": " & message
//! End Sub
//!
//! DebugLog "INFO", "Processing started"
//! ```
//!
//! ## Related Functions
//!
//! - `Time` - Returns a `Variant` containing the current system time (can be used in calculations)
//! - `Date$` - Returns a string representing the current system date
//! - `Now` - Returns the current system date and time as a `Date` value
//! - `Timer` - Returns the number of seconds elapsed since midnight
//! - `Hour()` - Extracts the hour component from a time value
//! - `Minute()` - Extracts the minute component from a time value
//! - `Second()` - Extracts the second component from a time value
//! - `Format$()` - Formats time values with custom formatting
//! - `TimeSerial()` - Creates a time value from hour, minute, and second
//! - `TimeValue()` - Converts a string to a time value
//!
//! ## Best Practices
//!
//! ### When to Use `Time$` vs `Time` vs `Now`
//!
//! ```vb6
//! ' Use Time$ for string display/logging
//! Debug.Print "Current time: " & Time$  ' "14:30:45"
//!
//! ' Use Time for time arithmetic
//! Dim currentTime As Date
//! currentTime = Time
//! Dim laterTime As Date
//! laterTime = currentTime + TimeSerial(1, 0, 0)  ' Add 1 hour
//!
//! ' Use Now for complete timestamp
//! Dim timestamp As Date
//! timestamp = Now  ' Includes both date and time
//! ```
//!
//! ### Combine with `Date$` for Full Timestamp
//!
//! ```vb6
//! Function GetFullTimestamp() As String
//!     GetFullTimestamp = Date$ & " " & Time$
//! End Function
//!
//! Debug.Print GetFullTimestamp()  ' "12/02/2025 14:30:45"
//! ```
//!
//! ### Use `Format$` for Custom Time Formatting
//!
//! ```vb6
//! ' Time$ always returns 24-hour format
//! Debug.Print Time$  ' "14:30:45"
//!
//! ' Use Format$ for 12-hour format or other formats
//! Debug.Print Format$(Now, "hh:mm:ss AM/PM")  ' "02:30:45 PM"
//! Debug.Print Format$(Now, "Long Time")       ' "2:30:45 PM"
//! ```
//!
//! ### Replace Colons for File Names
//!
//! ```vb6
//! Function SafeTimeStamp() As String
//!     ' Colons are invalid in filenames on Windows
//!     SafeTimeStamp = Replace$(Time$, ":", "")
//! End Function
//!
//! Dim fileName As String
//! fileName = "backup_" & SafeTimeStamp() & ".dat"  ' "backup_143045.dat"
//! ```
//!
//! ## Performance Considerations
//!
//! - `Time$` is a system call and relatively fast
//! - Calling repeatedly in tight loops may impact performance
//! - Cache the value if you need it multiple times in the same operation
//!
//! ```vb6
//! ' Less efficient: multiple calls
//! For i = 1 To 1000
//!     Debug.Print Time$ & " - Item " & i
//! Next i
//!
//! ' More efficient: cache the time
//! Dim currentTime As String
//! currentTime = Time$
//! For i = 1 To 1000
//!     Debug.Print currentTime & " - Item " & i
//! Next i
//! ```
//!
//! ## Common Pitfalls
//!
//! ### 1. 24-Hour Format Only
//!
//! ```vb6
//! ' Time$ always uses 24-hour format
//! Debug.Print Time$  ' "14:30:45" (not "2:30:45 PM")
//!
//! ' For 12-hour format, use Format$
//! Debug.Print Format$(Now, "hh:mm:ss AM/PM")  ' "02:30:45 PM"
//! ```
//!
//! ### 2. String Comparison Issues
//!
//! ```vb6
//! ' Comparing time strings can be tricky
//! If Time$ = "9:00:00" Then  ' Will NEVER match!
//!     ' Time$ returns "09:00:00" with leading zero
//! End If
//!
//! ' Correct: include leading zero
//! If Time$ = "09:00:00" Then
//!     ' This works
//! End If
//!
//! ' Better: use Time value for comparisons
//! If Time > TimeSerial(9, 0, 0) Then
//!     ' More reliable
//! End If
//! ```
//!
//! ### 3. Colons in File Names
//!
//! ```vb6
//! ' Invalid filename on Windows (colons not allowed)
//! fileName = "log_" & Time$ & ".txt"  ' "log_14:30:45.txt" - ERROR!
//!
//! ' Remove or replace colons
//! fileName = "log_" & Replace$(Time$, ":", "") & ".txt"  ' "log_143045.txt"
//! fileName = "log_" & Replace$(Time$, ":", "-") & ".txt"  ' "log_14-30-45.txt"
//! ```
//!
//! ### 4. Time Changes During Execution
//!
//! ```vb6
//! ' Problem: time can change between calls
//! Debug.Print "Start: " & Time$
//! ' ... long operation ...
//! Debug.Print "End: " & Time$  ' Different value!
//!
//! ' Solution: capture at start if consistency needed
//! Dim operationTime As String
//! operationTime = Time$
//! Debug.Print "Start: " & operationTime
//! ' ... long operation ...
//! Debug.Print "Started at: " & operationTime
//! Debug.Print "Completed at: " & Time$
//! ```
//!
//! ### 5. No Milliseconds
//!
//! ```vb6
//! ' Time$ only has second precision
//! Debug.Print Time$  ' "14:30:45" (no milliseconds)
//!
//! ' For higher precision, use Timer function
//! Dim startTimer As Single
//! startTimer = Timer
//! ' ... operation ...
//! Debug.Print "Elapsed: " & (Timer - startTimer) & " seconds"
//! ```
//!
//! ### 6. Locale Independence
//!
//! ```vb6
//! ' Time$ format is NOT affected by locale settings
//! ' Always returns "HH:MM:SS" regardless of system locale
//! Debug.Print Time$  ' Always "14:30:45" format
//!
//! ' For locale-specific formatting, use Format$
//! Debug.Print Format$(Now, "Long Time")  ' Respects locale
//! ```
//!
//! ## Practical Examples
//!
//! ### Creating Log Files with Timestamps
//!
//! ```vb6
//! Sub WriteLog(message As String)
//!     Dim logFile As String
//!     Dim timeStamp As String
//!     
//!     logFile = App.Path & "\application.log"
//!     timeStamp = Date$ & " " & Time$
//!     
//!     Open logFile For Append As #1
//!     Print #1, timeStamp & " - " & message
//!     Close #1
//! End Sub
//! ```
//!
//! ### Digital Clock Display
//!
//! ```vb6
//! Private Sub tmrClock_Timer()
//!     lblClock.Caption = Time$
//!     
//!     ' Update every second
//!     tmrClock.Interval = 1000
//! End Sub
//! ```
//!
//! ### Performance Monitoring
//!
//! ```vb6
//! Sub MonitorPerformance()
//!     Dim startTime As Double
//!     Dim endTime As Double
//!     
//!     Debug.Print "Operation started at: " & Time$
//!     startTime = Timer
//!     
//!     ' ... operation to monitor ...
//!     
//!     endTime = Timer
//!     Debug.Print "Operation ended at: " & Time$
//!     Debug.Print "Elapsed time: " & (endTime - startTime) & " seconds"
//! End Sub
//! ```
//!
//! ### Scheduled Task Execution
//!
//! ```vb6
//! Private Sub tmrScheduler_Timer()
//!     Dim currentTimeStr As String
//!     currentTimeStr = Time$
//!     
//!     Select Case currentTimeStr
//!         Case "09:00:00"
//!             RunMorningReport
//!         Case "12:00:00"
//!             RunNoonBackup
//!         Case "17:00:00"
//!             RunEveningCleanup
//!     End Select
//! End Sub
//! ```
//!
//! ## Limitations
//!
//! - Always returns 24-hour format (no AM/PM)
//! - No millisecond or sub-second precision
//! - No time zone information
//! - Format cannot be customized (use `Format$` for that)
//! - Returns string, not suitable for time arithmetic (use `Time` function instead)
//! - Colons in output make it unsuitable for filenames without modification
//! - Cannot be set (read-only; use `Time =` statement to set system time)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn time_dollar_simple() {
        let source = r#"
Sub Main()
    result = Time$
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_assignment() {
        let source = r#"
Sub Main()
    Dim currentTime As String
    currentTime = Time$
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_display() {
        let source = r#"
Sub ShowTime()
    Debug.Print "Current time: " & Time$
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_logging() {
        let source = r#"
Sub LogMessage(message As String)
    Dim logEntry As String
    logEntry = Time$ & " - " & message
    Debug.Print logEntry
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_file_naming() {
        let source = r#"
Function GenerateFileName(baseName As String) As String
    Dim timeStamp As String
    timeStamp = Replace$(Time$, ":", "")
    GenerateFileName = baseName & "_" & timeStamp & ".log"
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_in_condition() {
        let source = r#"
Sub Main()
    If Time$ = "09:00:00" Then
        Debug.Print "Morning task"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_timer_update() {
        let source = r#"
Sub Timer1_Timer()
    lblTime.Caption = Time$
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_with_date() {
        let source = r#"
Function GetFullTimestamp() As String
    GetFullTimestamp = Date$ & " " & Time$
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_processing_time() {
        let source = r#"
Sub ProcessData()
    Dim startTime As String
    startTime = Time$
    Debug.Print "Started at: " & startTime
    Debug.Print "Completed at: " & Time$
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_multiple_uses() {
        let source = r#"
Sub RecordActivity()
    Dim time1 As String
    Dim time2 As String
    time1 = Time$
    time2 = Time$
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_select_case() {
        let source = r#"
Sub CheckTime()
    Select Case Time$
        Case "09:00:00"
            RunMorningTask
        Case "17:00:00"
            RunEveningTask
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_audit_trail() {
        let source = r#"
Sub RecordAudit(action As String, userName As String)
    Dim auditEntry As String
    auditEntry = Date$ & " " & Time$ & " - " & userName & " - " & action
    Print #1, auditEntry
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_status_bar() {
        let source = r#"
Sub UpdateStatusBar()
    StatusBar1.Panels(1).Text = "Current Time: " & Time$
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_debug_log() {
        let source = r#"
Sub DebugLog(category As String, message As String)
    Debug.Print "[" & Time$ & "] " & category & ": " & message
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_in_function() {
        let source = r#"
Function GetCurrentTime() As String
    GetCurrentTime = Time$
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_greeting() {
        let source = r#"
Function GetGreeting() As String
    GetGreeting = "Good morning! Time: " & Time$
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_log_file() {
        let source = r#"
Sub WriteLog(message As String)
    Dim timeStamp As String
    timeStamp = Date$ & " " & Time$
    Print #1, timeStamp & " - " & message
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_loop_processing() {
        let source = r#"
Sub ProcessItems()
    Dim i As Integer
    Dim currentTimeStr As String
    currentTimeStr = Time$
    For i = 1 To 10
        Debug.Print currentTimeStr & " - Item " & i
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_concatenation() {
        let source = r#"
Sub Main()
    Dim output As String
    output = "Time is: " & Time$
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }

    #[test]
    fn time_dollar_scheduler() {
        let source = r#"
Sub CheckSchedule()
    Dim timeCheck As String
    timeCheck = Time$
    If timeCheck >= "09:00:00" Then
        RunScheduledTask
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeKeyword") && debug.contains("DollarSign"));
    }
}
