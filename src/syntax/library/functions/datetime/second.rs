//! # Second Function
//!
//! Returns an Integer specifying a whole number between 0 and 59, inclusive, representing the second of the minute.
//!
//! ## Syntax
//!
//! ```vb
//! Second(time)
//! ```
//!
//! ## Parameters
//!
//! - `time` - Required. Any Variant, numeric expression, string expression, or any combination that can represent a time. If `time` contains Null, Null is returned.
//!
//! ## Return Value
//!
//! Returns an `Integer` between 0 and 59 representing the second of the minute.
//!
//! ## Remarks
//!
//! The `Second` function extracts the seconds component from a time value. It is commonly used for time parsing, logging, and time-based calculations.
//!
//! **Important Notes**:
//! - Returns 0-59 (60 seconds in a minute)
//! - If time contains Null, Second returns Null
//! - If time is not a valid time, runtime error occurs (Error 13: Type mismatch)
//! - Only the time portion is used; date portion is ignored
//! - Accepts Date/Time values, numeric values, and valid time strings
//!
//! **Valid Input Examples**:
//! - Date/Time values: Now, Time, Date
//! - Time strings: "3:45:30 PM", "15:45:30"
//! - Numeric values representing dates: 0.5 (noon), 0.75 (6 PM)
//! - Combined date/time: #1/1/2000 3:45:30 PM#
//!
//! ## Typical Uses
//!
//! 1. **Time Parsing**: Extract seconds from time values
//! 2. **Logging**: Record precise timestamps
//! 3. **Time Calculations**: Compute time differences
//! 4. **Scheduling**: Check if task should run at specific second
//! 5. **Animation**: Time-based animations at second intervals
//! 6. **Performance Timing**: Measure elapsed seconds
//! 7. **Data Validation**: Verify time components
//! 8. **Report Generation**: Format time displays
//!
//! ## Basic Examples
//!
//! ### Example 1: Get Current Second
//! ```vb
//! Dim currentSecond As Integer
//! currentSecond = Second(Now)  ' Returns 0-59
//! ```
//!
//! ### Example 2: Parse Time String
//! ```vb
//! Dim timeStr As String
//! Dim sec As Integer
//!
//! timeStr = "3:45:30 PM"
//! sec = Second(timeStr)  ' Returns 30
//! ```
//!
//! ### Example 3: Check Specific Second
//! ```vb
//! If Second(Now) = 0 Then
//!     MsgBox "New minute started!"
//! End If
//! ```
//!
//! ### Example 4: Extract from Date/Time
//! ```vb
//! Dim dt As Date
//! Dim seconds As Integer
//!
//! dt = #1/15/2000 10:30:45 AM#
//! seconds = Second(dt)  ' Returns 45
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `GetTimeSeconds`
//! ```vb
//! Function GetTimeSeconds(timeValue As Variant) As Integer
//!     ' Safely get seconds from time value
//!     On Error Resume Next
//!     GetTimeSeconds = Second(timeValue)
//!     If Err.Number <> 0 Then
//!         GetTimeSeconds = 0
//!     End If
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ### Pattern 2: `FormatTimeWithSeconds`
//! ```vb
//! Function FormatTimeWithSeconds(timeValue As Date) As String
//!     ' Format time as HH:MM:SS
//!     Dim h As Integer, m As Integer, s As Integer
//!     
//!     h = Hour(timeValue)
//!     m = Minute(timeValue)
//!     s = Second(timeValue)
//!     
//!     FormatTimeWithSeconds = Format(h, "00") & ":" & _
//!                            Format(m, "00") & ":" & _
//!                            Format(s, "00")
//! End Function
//! ```
//!
//! ### Pattern 3: `IsAtSecondInterval`
//! ```vb
//! Function IsAtSecondInterval(timeValue As Date, interval As Integer) As Boolean
//!     ' Check if time is at a specific second interval
//!     ' e.g., IsAtSecondInterval(Now, 15) returns True at :00, :15, :30, :45
//!     IsAtSecondInterval = (Second(timeValue) Mod interval = 0)
//! End Function
//! ```
//!
//! ### Pattern 4: `GetElapsedSeconds`
//! ```vb
//! Function GetElapsedSeconds(startTime As Date, endTime As Date) As Long
//!     ' Get total elapsed seconds between two times
//!     GetElapsedSeconds = DateDiff("s", startTime, endTime)
//! End Function
//! ```
//!
//! ### Pattern 5: `SetSeconds`
//! ```vb
//! Function SetSeconds(timeValue As Date, newSeconds As Integer) As Date
//!     ' Set the seconds component of a time value
//!     Dim h As Integer, m As Integer
//!     
//!     h = Hour(timeValue)
//!     m = Minute(timeValue)
//!     
//!     SetSeconds = TimeSerial(h, m, newSeconds)
//! End Function
//! ```
//!
//! ### Pattern 6: `RoundToNearestMinute`
//! ```vb
//! Function RoundToNearestMinute(timeValue As Date) As Date
//!     ' Round time to nearest minute based on seconds
//!     Dim h As Integer, m As Integer, s As Integer
//!     
//!     h = Hour(timeValue)
//!     m = Minute(timeValue)
//!     s = Second(timeValue)
//!     
//!     If s >= 30 Then
//!         m = m + 1
//!         If m = 60 Then
//!             m = 0
//!             h = h + 1
//!             If h = 24 Then h = 0
//!         End If
//!     End If
//!     
//!     RoundToNearestMinute = TimeSerial(h, m, 0)
//! End Function
//! ```
//!
//! ### Pattern 7: `CompareTimeToSecond`
//! ```vb
//! Function CompareTimeToSecond(time1 As Date, time2 As Date) As Boolean
//!     ' Compare times to the second (ignore milliseconds)
//!     Dim h1 As Integer, m1 As Integer, s1 As Integer
//!     Dim h2 As Integer, m2 As Integer, s2 As Integer
//!     
//!     h1 = Hour(time1): m1 = Minute(time1): s1 = Second(time1)
//!     h2 = Hour(time2): m2 = Minute(time2): s2 = Second(time2)
//!     
//!     CompareTimeToSecond = (h1 = h2 And m1 = m2 And s1 = s2)
//! End Function
//! ```
//!
//! ### Pattern 8: `ValidateTimeSeconds`
//! ```vb
//! Function ValidateTimeSeconds(timeValue As Variant) As Boolean
//!     ' Validate that seconds are in valid range
//!     On Error Resume Next
//!     Dim s As Integer
//!     s = Second(timeValue)
//!     
//!     If Err.Number <> 0 Then
//!         ValidateTimeSeconds = False
//!     Else
//!         ValidateTimeSeconds = (s >= 0 And s <= 59)
//!     End If
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ### Pattern 9: `GetSecondsUntilMinute`
//! ```vb
//! Function GetSecondsUntilMinute(Optional timeValue As Date) As Integer
//!     ' Get seconds remaining until next minute
//!     Dim currentTime As Date
//!     
//!     If timeValue = 0 Then
//!         currentTime = Now
//!     Else
//!         currentTime = timeValue
//!     End If
//!     
//!     GetSecondsUntilMinute = 60 - Second(currentTime)
//! End Function
//! ```
//!
//! ### Pattern 10: `ParseTimeComponents`
//! ```vb
//! Sub ParseTimeComponents(timeValue As Date, hours As Integer, _
//!                        minutes As Integer, seconds As Integer)
//!     ' Parse time into separate components
//!     hours = Hour(timeValue)
//!     minutes = Minute(timeValue)
//!     seconds = Second(timeValue)
//! End Sub
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Precise Timer Class
//! ```vb
//! ' High-precision timer for performance measurement
//! Class PreciseTimer
//!     Private m_startTime As Date
//!     Private m_running As Boolean
//!     
//!     Public Sub Start()
//!         m_startTime = Now
//!         m_running = True
//!     End Sub
//!     
//!     Public Sub Stop()
//!         m_running = False
//!     End Sub
//!     
//!     Public Function GetElapsedSeconds() As Long
//!         ' Get elapsed seconds
//!         Dim endTime As Date
//!         
//!         If m_running Then
//!             endTime = Now
//!         Else
//!             endTime = m_startTime
//!         End If
//!         
//!         GetElapsedSeconds = DateDiff("s", m_startTime, endTime)
//!     End Function
//!     
//!     Public Function GetElapsedTime() As String
//!         ' Get formatted elapsed time
//!         Dim elapsed As Long
//!         Dim hours As Long, minutes As Long, seconds As Long
//!         
//!         elapsed = GetElapsedSeconds()
//!         
//!         hours = elapsed \ 3600
//!         minutes = (elapsed Mod 3600) \ 60
//!         seconds = elapsed Mod 60
//!         
//!         GetElapsedTime = Format(hours, "00") & ":" & _
//!                         Format(minutes, "00") & ":" & _
//!                         Format(seconds, "00")
//!     End Function
//!     
//!     Public Function GetCurrentSecond() As Integer
//!         ' Get current second of running timer
//!         If m_running Then
//!             GetCurrentSecond = Second(Now)
//!         Else
//!             GetCurrentSecond = Second(m_startTime)
//!         End If
//!     End Function
//!     
//!     Public Sub Reset()
//!         m_startTime = Now
//!         m_running = False
//!     End Sub
//!     
//!     Public Function IsRunning() As Boolean
//!         IsRunning = m_running
//!     End Function
//! End Class
//! ```
//!
//! ### Example 2: Time Logger Module
//! ```vb
//! ' Log events with precise timestamps
//! Module TimeLogger
//!     Private Type LogEntry
//!         Timestamp As Date
//!         Message As String
//!         Second As Integer
//!     End Type
//!     
//!     Private m_logEntries() As LogEntry
//!     Private m_count As Integer
//!     
//!     Public Sub InitializeLog()
//!         m_count = 0
//!         ReDim m_logEntries(0 To 999)
//!     End Sub
//!     
//!     Public Sub LogEvent(message As String, Optional timeValue As Date)
//!         Dim entry As LogEntry
//!         Dim logTime As Date
//!         
//!         If m_count > UBound(m_logEntries) Then
//!             ReDim Preserve m_logEntries(0 To UBound(m_logEntries) + 1000)
//!         End If
//!         
//!         If timeValue = 0 Then
//!             logTime = Now
//!         Else
//!             logTime = timeValue
//!         End If
//!         
//!         entry.Timestamp = logTime
//!         entry.Message = message
//!         entry.Second = Second(logTime)
//!         
//!         m_logEntries(m_count) = entry
//!         m_count = m_count + 1
//!     End Sub
//!     
//!     Public Function GetFormattedLog() As String
//!         ' Get formatted log with timestamps
//!         Dim i As Integer
//!         Dim result As String
//!         Dim timeStr As String
//!         
//!         result = "Event Log" & vbCrLf
//!         result = result & String(50, "-") & vbCrLf
//!         
//!         For i = 0 To m_count - 1
//!             timeStr = Format(m_logEntries(i).Timestamp, "yyyy-mm-dd hh:nn:ss")
//!             result = result & timeStr & " | " & m_logEntries(i).Message & vbCrLf
//!         Next i
//!         
//!         GetFormattedLog = result
//!     End Function
//!     
//!     Public Function GetEventsAtSecond(targetSecond As Integer) As String
//!         ' Get all events that occurred at a specific second
//!         Dim i As Integer
//!         Dim result As String
//!         Dim count As Integer
//!         
//!         result = "Events at second " & targetSecond & ":" & vbCrLf
//!         count = 0
//!         
//!         For i = 0 To m_count - 1
//!             If m_logEntries(i).Second = targetSecond Then
//!                 result = result & "  " & _
//!                         Format(m_logEntries(i).Timestamp, "hh:nn:ss") & " | " & _
//!                         m_logEntries(i).Message & vbCrLf
//!                 count = count + 1
//!             End If
//!         Next i
//!         
//!         result = result & vbCrLf & "Total: " & count & " events"
//!         GetEventsAtSecond = result
//!     End Function
//!     
//!     Public Function GetLogCount() As Integer
//!         GetLogCount = m_count
//!     End Function
//!     
//!     Public Sub ClearLog()
//!         m_count = 0
//!         ReDim m_logEntries(0 To 999)
//!     End Sub
//! End Module
//! ```
//!
//! ### Example 3: Time-Based Task Scheduler
//! ```vb
//! ' Schedule tasks to run at specific seconds
//! Class TaskScheduler
//!     Private Type ScheduledTask
//!         TaskName As String
//!         TargetSecond As Integer
//!         Interval As Integer
//!         LastRun As Date
//!         Enabled As Boolean
//!     End Type
//!     
//!     Private m_tasks() As ScheduledTask
//!     Private m_count As Integer
//!     
//!     Public Sub Initialize()
//!         m_count = 0
//!         ReDim m_tasks(0 To 99)
//!     End Sub
//!     
//!     Public Sub AddTask(taskName As String, targetSecond As Integer, _
//!                       Optional interval As Integer = 60)
//!         ' Add task to run at specific second
//!         If m_count > UBound(m_tasks) Then
//!             ReDim Preserve m_tasks(0 To UBound(m_tasks) + 50)
//!         End If
//!         
//!         m_tasks(m_count).TaskName = taskName
//!         m_tasks(m_count).TargetSecond = targetSecond
//!         m_tasks(m_count).Interval = interval
//!         m_tasks(m_count).LastRun = 0
//!         m_tasks(m_count).Enabled = True
//!         
//!         m_count = m_count + 1
//!     End Sub
//!     
//!     Public Function CheckTasks() As String
//!         ' Check which tasks should run now
//!         Dim currentTime As Date
//!         Dim currentSecond As Integer
//!         Dim i As Integer
//!         Dim tasksToRun As String
//!         Dim elapsedSeconds As Long
//!         
//!         currentTime = Now
//!         currentSecond = Second(currentTime)
//!         tasksToRun = ""
//!         
//!         For i = 0 To m_count - 1
//!             If m_tasks(i).Enabled Then
//!                 If currentSecond = m_tasks(i).TargetSecond Then
//!                     If m_tasks(i).LastRun = 0 Then
//!                         ' First run
//!                         tasksToRun = tasksToRun & m_tasks(i).TaskName & ";"
//!                         m_tasks(i).LastRun = currentTime
//!                     Else
//!                         ' Check interval
//!                         elapsedSeconds = DateDiff("s", m_tasks(i).LastRun, currentTime)
//!                         If elapsedSeconds >= m_tasks(i).Interval Then
//!                             tasksToRun = tasksToRun & m_tasks(i).TaskName & ";"
//!                             m_tasks(i).LastRun = currentTime
//!                         End If
//!                     End If
//!                 End If
//!             End If
//!         Next i
//!         
//!         CheckTasks = tasksToRun
//!     End Function
//!     
//!     Public Sub EnableTask(taskName As String)
//!         Dim i As Integer
//!         For i = 0 To m_count - 1
//!             If m_tasks(i).TaskName = taskName Then
//!                 m_tasks(i).Enabled = True
//!                 Exit Sub
//!             End If
//!         Next i
//!     End Sub
//!     
//!     Public Sub DisableTask(taskName As String)
//!         Dim i As Integer
//!         For i = 0 To m_count - 1
//!             If m_tasks(i).TaskName = taskName Then
//!                 m_tasks(i).Enabled = False
//!                 Exit Sub
//!             End If
//!         Next i
//!     End Sub
//! End Class
//! ```
//!
//! ### Example 4: Animation Timer
//! ```vb
//! ' Time-based animation controller using seconds
//! Class AnimationTimer
//!     Private m_startTime As Date
//!     Private m_duration As Integer  ' Duration in seconds
//!     Private m_running As Boolean
//!     
//!     Public Sub StartAnimation(durationSeconds As Integer)
//!         m_startTime = Now
//!         m_duration = durationSeconds
//!         m_running = True
//!     End Sub
//!     
//!     Public Function GetProgress() As Double
//!         ' Get animation progress (0.0 to 1.0)
//!         Dim elapsed As Long
//!         
//!         If Not m_running Then
//!             GetProgress = 0
//!             Exit Function
//!         End If
//!         
//!         elapsed = DateDiff("s", m_startTime, Now)
//!         
//!         If elapsed >= m_duration Then
//!             GetProgress = 1
//!             m_running = False
//!         Else
//!             GetProgress = elapsed / m_duration
//!         End If
//!     End Function
//!     
//!     Public Function GetElapsedSeconds() As Integer
//!         ' Get elapsed seconds since animation start
//!         If m_running Then
//!             GetElapsedSeconds = DateDiff("s", m_startTime, Now)
//!         Else
//!             GetElapsedSeconds = 0
//!         End If
//!     End Function
//!     
//!     Public Function GetRemainingSeconds() As Integer
//!         ' Get remaining seconds in animation
//!         Dim elapsed As Integer
//!         
//!         If Not m_running Then
//!             GetRemainingSeconds = 0
//!             Exit Function
//!         End If
//!         
//!         elapsed = DateDiff("s", m_startTime, Now)
//!         
//!         If elapsed >= m_duration Then
//!             GetRemainingSeconds = 0
//!             m_running = False
//!         Else
//!             GetRemainingSeconds = m_duration - elapsed
//!         End If
//!     End Function
//!     
//!     Public Function IsComplete() As Boolean
//!         IsComplete = Not m_running
//!     End Function
//!     
//!     Public Sub Stop()
//!         m_running = False
//!     End Sub
//!     
//!     Public Function GetCurrentSecond() As Integer
//!         GetCurrentSecond = Second(Now)
//!     End Function
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! The `Second` function generates errors in specific situations:
//!
//! **Error 13: Type mismatch**
//! - Occurs when the argument cannot be interpreted as a date/time value
//!
//! **Error 94: Invalid use of Null**
//! - Can occur if Null is passed and not handled properly
//!
//! Example error handling:
//!
//! ```vb
//! On Error Resume Next
//! Dim sec As Integer
//! sec = Second(userInput)
//! If Err.Number <> 0 Then
//!     MsgBox "Invalid time value"
//!     sec = 0
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Performance Considerations
//!
//! - `Second` is very fast - simple extraction operation
//! - No significant performance considerations
//! - Can be called thousands of times with minimal impact
//! - Consider caching if calling repeatedly on same value
//!
//! ## Best Practices
//!
//! 1. **Validate Input**: Check that input is valid date/time before calling
//! 2. **Handle Null**: Check for Null if data source is uncertain
//! 3. **Use with Other Time Functions**: Combine with Hour, Minute for complete parsing
//! 4. **Format Consistently**: Use consistent time formatting in your application
//! 5. **Consider Time Zones**: Be aware of time zone issues in time calculations
//! 6. **Document Precision**: Clarify if milliseconds matter in your application
//! 7. **Use for Validation**: Validate time components are in expected ranges
//! 8. **Cache Now Values**: If using Now multiple times, cache to avoid timing issues
//! 9. **Test Edge Cases**: Test with midnight, end of minute, etc.
//! 10. **Use `TimeSerial`**: Combine with `TimeSerial` to construct times
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Range |
//! |----------|---------|---------|-------|
//! | **Second** | Get seconds | Integer | 0-59 |
//! | **Minute** | Get minutes | Integer | 0-59 |
//! | **Hour** | Get hours | Integer | 0-23 |
//! | **Day** | Get day | Integer | 1-31 |
//! | **Month** | Get month | Integer | 1-12 |
//! | **Year** | Get year | Integer | 100-9999 |
//! | **`TimeSerial`** | Create time | Date | Constructs time from H:M:S |
//! | **`DatePart`** | Get date part | Variant | Flexible date/time extraction |
//!
//! ## Platform and Version Notes
//!
//! - Available in all versions of VB6 and VBA
//! - Behavior consistent across all platforms
//! - In VB.NET, replaced by DateTime.Second property
//! - Returns Integer type (not Long)
//! - Always returns 0-59 range
//!
//! ## Limitations
//!
//! - No millisecond precision
//! - Cannot distinguish leap seconds
//! - Range limited to 0-59 (no support for values outside this range)
//! - Date portion of Date/Time value is ignored
//! - Cannot be used as `LValue` (cannot assign to Second)
//!
//! ## Related Functions
//!
//! - `Minute`: Returns the minute of the hour (0-59)
//! - `Hour`: Returns the hour of the day (0-23)
//! - `Now`: Returns the current system date and time
//! - `Time`: Returns the current system time
//! - `TimeSerial`: Returns a Date value for a specific time
//! - `DatePart`: Returns a specified part of a given date
//! - `Timer`: Returns seconds since midnight as Single

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn second_basic() {
        let source = r"
Dim sec As Integer
sec = Second(Now)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("sec"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("sec"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Second"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("Now"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn second_time_string() {
        let source = r#"
Dim seconds As Integer
seconds = Second("3:45:30 PM")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("seconds"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("seconds"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Second"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"3:45:30 PM\""),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn second_if_statement() {
        let source = r#"
If Second(Now) = 0 Then
    MsgBox "New minute!"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Second"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("Now"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    EqualityOperator,
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
                        StringLiteral ("\"New minute!\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn second_function_return() {
        let source = r"
Function GetSeconds(timeValue As Date) As Integer
    GetSeconds = Second(timeValue)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetSeconds"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("timeValue"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DateKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetSeconds"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Second"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("timeValue"),
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
    fn second_variable_assignment() {
        let source = r"
Dim s As Integer
s = Second(currentTime)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("s"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("s"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Second"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("currentTime"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn second_msgbox() {
        let source = r#"
MsgBox "Second: " & Second(Time)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("MsgBox"),
                Whitespace,
                StringLiteral ("\"Second: \""),
                Whitespace,
                Ampersand,
                Whitespace,
                Identifier ("Second"),
                LeftParenthesis,
                TimeKeyword,
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn second_debug_print() {
        let source = r"
Debug.Print Second(Now)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("Debug"),
                PeriodOperator,
                PrintKeyword,
                Whitespace,
                Identifier ("Second"),
                LeftParenthesis,
                Identifier ("Now"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn second_select_case() {
        let source = r#"
Select Case Second(Now)
    Case 0 To 15
        msg = "First quarter"
    Case 16 To 30
        msg = "Second quarter"
End Select
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SelectCaseStatement {
                SelectKeyword,
                Whitespace,
                CaseKeyword,
                Whitespace,
                CallExpression {
                    Identifier ("Second"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("Now"),
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
                    IntegerLiteral ("0"),
                    Whitespace,
                    ToKeyword,
                    Whitespace,
                    IntegerLiteral ("15"),
                    Newline,
                    StatementList {
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("msg"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"First quarter\""),
                            },
                            Newline,
                        },
                        Whitespace,
                    },
                },
                CaseClause {
                    CaseKeyword,
                    Whitespace,
                    IntegerLiteral ("16"),
                    Whitespace,
                    ToKeyword,
                    Whitespace,
                    IntegerLiteral ("30"),
                    Newline,
                    StatementList {
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("msg"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"Second quarter\""),
                            },
                            Newline,
                        },
                    },
                },
                EndKeyword,
                Whitespace,
                SelectKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn second_class_usage() {
        let source = r"
Private m_seconds As Integer

Public Sub UpdateTime()
    m_seconds = Second(Now)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_seconds"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
            Newline,
            SubStatement {
                PublicKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("UpdateTime"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("m_seconds"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Second"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("Now"),
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
    fn second_with_statement() {
        let source = r"
With timeData
    .Seconds = Second(.TimeValue)
End With
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            WithStatement {
                WithKeyword,
                Whitespace,
                Identifier ("timeData"),
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            PeriodOperator,
                        },
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("Seconds"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("Second"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            PeriodOperator,
                                        },
                                    },
                                },
                            },
                        },
                    },
                    CallStatement {
                        Identifier ("TimeValue"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                WithKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn second_elseif() {
        let source = r#"
If Second(t) < 30 Then
    half = "First"
ElseIf Second(t) >= 30 Then
    half = "Second"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Second"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("t"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    LessThanOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("30"),
                    },
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("half"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"First\""),
                        },
                        Newline,
                    },
                },
                ElseIfClause {
                    ElseIfKeyword,
                    Whitespace,
                    BinaryExpression {
                        CallExpression {
                            Identifier ("Second"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("t"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        GreaterThanOrEqualOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("30"),
                        },
                    },
                    Whitespace,
                    ThenKeyword,
                    Newline,
                    StatementList {
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("half"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"Second\""),
                            },
                            Newline,
                        },
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn second_for_loop() {
        let source = r"
For i = 1 To 10
    timestamps(i) = Second(times(i))
Next i
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            ForStatement {
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
                    AssignmentStatement {
                        CallExpression {
                            Identifier ("timestamps"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("i"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Second"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("times"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("i"),
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
                    },
                },
                NextKeyword,
                Whitespace,
                Identifier ("i"),
                Newline,
            },
        ]);
    }

    #[test]
    fn second_do_while() {
        let source = r"
Do While Second(Now) < 30
    DoSomething
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DoStatement {
                DoKeyword,
                Whitespace,
                WhileKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Second"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("Now"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    LessThanOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("30"),
                    },
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("DoSomething"),
                        Newline,
                    },
                },
                LoopKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn second_do_until() {
        let source = r"
Do Until Second(currentTime) = 0
    currentTime = Now
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DoStatement {
                DoKeyword,
                Whitespace,
                UntilKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Second"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("currentTime"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    EqualityOperator,
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
                            Identifier ("currentTime"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("Now"),
                        },
                        Newline,
                    },
                },
                LoopKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn second_while_wend() {
        let source = r"
While Second(Now) > 45
    Wait
Wend
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            WhileStatement {
                WhileKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Second"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("Now"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    GreaterThanOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("45"),
                    },
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Wait"),
                        Newline,
                    },
                },
                WendKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn second_parentheses() {
        let source = r"
Dim val As Integer
val = (Second(Now))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("val"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("val"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                ParenthesizedExpression {
                    LeftParenthesis,
                    CallExpression {
                        Identifier ("Second"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("Now"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn second_iif() {
        let source = r#"
Dim display As String
display = IIf(Second(Now) < 10, "0" & Second(Now), Second(Now))
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("display"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("display"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("IIf"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("Second"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("Now"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                LessThanOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("10"),
                                },
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BinaryExpression {
                                StringLiteralExpression {
                                    StringLiteral ("\"0\""),
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                CallExpression {
                                    Identifier ("Second"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("Now"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            CallExpression {
                                Identifier ("Second"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("Now"),
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
            },
        ]);
    }

    #[test]
    fn second_array_assignment() {
        let source = r"
Dim seconds(10) As Integer
seconds(i) = Second(timeArray(i))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("seconds"),
                LeftParenthesis,
                NumericLiteralExpression {
                    IntegerLiteral ("10"),
                },
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
            AssignmentStatement {
                CallExpression {
                    Identifier ("seconds"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("i"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Second"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("timeArray"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("i"),
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
            },
        ]);
    }

    #[test]
    fn second_property_assignment() {
        let source = r"
Set obj = New TimeData
obj.SecondValue = Second(Now)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SetStatement {
                SetKeyword,
                Whitespace,
                Identifier ("obj"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                NewKeyword,
                Whitespace,
                Identifier ("TimeData"),
                Newline,
            },
            AssignmentStatement {
                MemberAccessExpression {
                    Identifier ("obj"),
                    PeriodOperator,
                    Identifier ("SecondValue"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Second"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("Now"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn second_function_argument() {
        let source = r"
Call LogTime(Second(Now))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                CallKeyword,
                Whitespace,
                Identifier ("LogTime"),
                LeftParenthesis,
                Identifier ("Second"),
                LeftParenthesis,
                Identifier ("Now"),
                RightParenthesis,
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn second_concatenation() {
        let source = r#"
Dim msg As String
msg = "Seconds: " & Second(Time)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("msg"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("msg"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    StringLiteralExpression {
                        StringLiteral ("\"Seconds: \""),
                    },
                    Whitespace,
                    Ampersand,
                    Whitespace,
                    CallExpression {
                        Identifier ("Second"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    TimeKeyword,
                                },
                            },
                        },
                        RightParenthesis,
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn second_comparison() {
        let source = r#"
If Second(time1) = Second(time2) Then
    MsgBox "Same second"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("Second"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("time1"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    CallExpression {
                        Identifier ("Second"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("time2"),
                                },
                            },
                        },
                        RightParenthesis,
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
                        StringLiteral ("\"Same second\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn second_modulo() {
        let source = r"
If Second(Now) Mod 15 = 0 Then
    UpdateDisplay
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            IfStatement {
                IfKeyword,
                Whitespace,
                BinaryExpression {
                    BinaryExpression {
                        CallExpression {
                            Identifier ("Second"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("Now"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        ModKeyword,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("15"),
                        },
                    },
                    Whitespace,
                    EqualityOperator,
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
                        Identifier ("UpdateDisplay"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn second_with_format() {
        let source = r#"
Dim formatted As String
formatted = Format(Second(Now), "00")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
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
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("formatted"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Format"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Second"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("Now"),
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
                                StringLiteral ("\"00\""),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn second_date_literal() {
        let source = r"
Dim s As Integer
s = Second(#1/15/2000 10:30:45 AM#)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("s"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("s"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Second"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            LiteralExpression {
                                DateLiteral ("#1/15/2000 10:30:45 AM#"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn second_error_handling() {
        let source = r"
On Error Resume Next
Dim sec As Integer
sec = Second(userInput)
If Err.Number <> 0 Then
    sec = 0
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            OnErrorStatement {
                OnKeyword,
                Whitespace,
                ErrorKeyword,
                Whitespace,
                ResumeKeyword,
                Whitespace,
                NextKeyword,
                Newline,
            },
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("sec"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("sec"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Second"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("userInput"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            IfStatement {
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
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("sec"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("0"),
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                IfKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn second_on_error_goto() {
        let source = r#"
Sub GetTimeSecond()
    On Error GoTo ErrorHandler
    Dim s As Integer
    s = Second(timeValue)
    Exit Sub
ErrorHandler:
    MsgBox "Error getting second"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("GetTimeSecond"),
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
                        Identifier ("s"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
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
                            Identifier ("Second"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("timeValue"),
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
                        StringLiteral ("\"Error getting second\""),
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
