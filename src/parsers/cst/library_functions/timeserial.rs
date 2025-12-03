//! VB6 `TimeSerial` Function
//!
//! The `TimeSerial` function returns a Variant (Date) containing the time for a specific hour, minute, and second.
//!
//! ## Syntax
//! ```vb6
//! TimeSerial(hour, minute, second)
//! ```
//!
//! ## Parameters
//! - `hour`: Required. Integer from 0 to 23, representing the hour. Values outside this range are normalized.
//! - `minute`: Required. Integer representing the minute. Values outside 0-59 are normalized.
//! - `second`: Required. Integer representing the second. Values outside 0-59 are normalized.
//!
//! ## Returns
//! Returns a `Variant` of subtype `Date` containing a time value. The date portion is set to zero (December 30, 1899).
//!
//! ## Remarks
//! The `TimeSerial` function creates time values from component parts:
//!
//! - **24-hour format**: Hour parameter uses 24-hour format (0-23)
//! - **Normalization**: Values outside normal ranges are automatically adjusted
//! - **Date portion**: Always returns zero date (12/30/1899)
//! - **Overflow handling**: Excess values roll over (e.g., 90 seconds = 1 minute 30 seconds)
//! - **Negative values**: Can use negative values to subtract time
//! - **Calculation flexibility**: Can use expressions for any parameter
//! - **Time arithmetic**: Ideal for adding/subtracting time intervals
//! - **Companion to `DateSerial`**: `TimeSerial` for time, `DateSerial` for dates
//! - **Type returned**: Returns Variant (Date), not a numeric type
//!
//! ### Normalization Examples
//! ```vb6
//! ' These all produce valid times through normalization:
//! TimeSerial(0, 0, 90)      ' = 00:01:30 (90 seconds = 1 min 30 sec)
//! TimeSerial(0, 90, 0)      ' = 01:30:00 (90 minutes = 1 hour 30 min)
//! TimeSerial(25, 0, 0)      ' = 01:00:00 (25 hours = 1 AM next day)
//! TimeSerial(0, -30, 0)     ' = 23:30:00 (previous day)
//! TimeSerial(12, 30, -60)   ' = 12:29:00 (subtract 60 seconds)
//! ```
//!
//! ### Time Arithmetic
//! ```vb6
//! ' Add 2 hours to current time
//! newTime = Time + TimeSerial(2, 0, 0)
//!
//! ' Subtract 30 minutes
//! newTime = Time + TimeSerial(0, -30, 0)
//!
//! ' Add 1 hour 15 minutes
//! newTime = Time + TimeSerial(1, 15, 0)
//! ```
//!
//! ### Creating Specific Times
//! ```vb6
//! ' 8:30 AM
//! morning = TimeSerial(8, 30, 0)
//!
//! ' Noon
//! noon = TimeSerial(12, 0, 0)
//!
//! ' 11:59:59 PM
//! lastSecond = TimeSerial(23, 59, 59)
//!
//! ' Midnight
//! midnight = TimeSerial(0, 0, 0)
//! ```
//!
//! ## Typical Uses
//! 1. **Create Time Values**: Build time from components
//! 2. **Time Arithmetic**: Add/subtract hours, minutes, seconds
//! 3. **Schedule Times**: Define specific times for scheduling
//! 4. **Time Comparison**: Create reference times for comparison
//! 5. **Time Calculations**: Calculate time differences
//! 6. **Business Hours**: Define opening/closing times
//! 7. **Time Intervals**: Represent durations
//! 8. **Alarm Times**: Set specific alarm or reminder times
//!
//! ## Basic Examples
//!
//! ### Example 1: Create Specific Time
//! ```vb6
//! Sub CreateTime()
//!     Dim businessOpen As Date
//!     businessOpen = TimeSerial(9, 0, 0)  ' 9:00 AM
//!     MsgBox "Opens at: " & Format$(businessOpen, "hh:mm AM/PM")
//! End Sub
//! ```
//!
//! ### Example 2: Add Time to Current Time
//! ```vb6
//! Function AddHours(hours As Integer) As Date
//!     AddHours = Time + TimeSerial(hours, 0, 0)
//! End Function
//! ```
//!
//! ### Example 3: Calculate Time Difference
//! ```vb6
//! Function GetTimeDuration(hours As Integer, minutes As Integer) As Date
//!     GetTimeDuration = TimeSerial(hours, minutes, 0)
//! End Function
//! ```
//!
//! ### Example 4: Check If Time Is Between Range
//! ```vb6
//! Function IsInTimeRange(checkTime As Date, startHour As Integer, endHour As Integer) As Boolean
//!     Dim startTime As Date
//!     Dim endTime As Date
//!     
//!     startTime = TimeSerial(startHour, 0, 0)
//!     endTime = TimeSerial(endHour, 0, 0)
//!     
//!     IsInTimeRange = (checkTime >= startTime And checkTime < endTime)
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Add Minutes to Time
//! ```vb6
//! Function AddMinutes(baseTime As Date, minutes As Integer) As Date
//!     AddMinutes = baseTime + TimeSerial(0, minutes, 0)
//! End Function
//! ```
//!
//! ### Pattern 2: Add Seconds to Time
//! ```vb6
//! Function AddSeconds(baseTime As Date, seconds As Integer) As Date
//!     AddSeconds = baseTime + TimeSerial(0, 0, seconds)
//! End Function
//! ```
//!
//! ### Pattern 3: Create Time from Total Seconds
//! ```vb6
//! Function SecondsToTime(totalSeconds As Long) As Date
//!     Dim hours As Long
//!     Dim minutes As Long
//!     Dim seconds As Long
//!     
//!     hours = totalSeconds \ 3600
//!     minutes = (totalSeconds Mod 3600) \ 60
//!     seconds = totalSeconds Mod 60
//!     
//!     SecondsToTime = TimeSerial(hours, minutes, seconds)
//! End Function
//! ```
//!
//! ### Pattern 4: Round Time to Nearest Interval
//! ```vb6
//! Function RoundToNearestMinutes(t As Date, intervalMinutes As Integer) As Date
//!     Dim totalMinutes As Long
//!     Dim roundedMinutes As Long
//!     
//!     totalMinutes = Hour(t) * 60 + Minute(t)
//!     roundedMinutes = ((totalMinutes + intervalMinutes \ 2) \ intervalMinutes) * intervalMinutes
//!     
//!     RoundToNearestMinutes = TimeSerial(roundedMinutes \ 60, roundedMinutes Mod 60, 0)
//! End Function
//! ```
//!
//! ### Pattern 5: Calculate Elapsed Time
//! ```vb6
//! Function CalculateElapsedTime(startTime As Date, endTime As Date) As Date
//!     Dim diffSeconds As Long
//!     
//!     diffSeconds = DateDiff("s", startTime, endTime)
//!     CalculateElapsedTime = TimeSerial(0, 0, diffSeconds)
//! End Function
//! ```
//!
//! ### Pattern 6: Get Noon Time
//! ```vb6
//! Function GetNoon() As Date
//!     GetNoon = TimeSerial(12, 0, 0)
//! End Function
//! ```
//!
//! ### Pattern 7: Get Midnight Time
//! ```vb6
//! Function GetMidnight() As Date
//!     GetMidnight = TimeSerial(0, 0, 0)
//! End Function
//! ```
//!
//! ### Pattern 8: Create Business Hours Range
//! ```vb6
//! Sub GetBusinessHours(ByRef openTime As Date, ByRef closeTime As Date)
//!     openTime = TimeSerial(9, 0, 0)    ' 9 AM
//!     closeTime = TimeSerial(17, 0, 0)  ' 5 PM
//! End Sub
//! ```
//!
//! ### Pattern 9: Add Time Duration
//! ```vb6
//! Function AddDuration(baseTime As Date, hours As Integer, minutes As Integer, seconds As Integer) As Date
//!     AddDuration = baseTime + TimeSerial(hours, minutes, seconds)
//! End Function
//! ```
//!
//! ### Pattern 10: Normalize Time Components
//! ```vb6
//! Function NormalizeTime(hours As Integer, minutes As Integer, seconds As Integer) As Date
//!     NormalizeTime = TimeSerial(hours, minutes, seconds)
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Time Calculator Class
//! ```vb6
//! ' Class: TimeCalculator
//! ' Performs various time calculations and manipulations
//! Option Explicit
//!
//! Public Function AddTime(baseTime As Date, hours As Integer, minutes As Integer, seconds As Integer) As Date
//!     AddTime = baseTime + TimeSerial(hours, minutes, seconds)
//! End Function
//!
//! Public Function SubtractTime(baseTime As Date, hours As Integer, minutes As Integer, seconds As Integer) As Date
//!     SubtractTime = baseTime + TimeSerial(-hours, -minutes, -seconds)
//! End Function
//!
//! Public Function GetTimeBetween(startTime As Date, endTime As Date) As Date
//!     Dim diffSeconds As Long
//!     Dim hours As Long
//!     Dim minutes As Long
//!     Dim seconds As Long
//!     
//!     diffSeconds = DateDiff("s", startTime, endTime)
//!     
//!     hours = diffSeconds \ 3600
//!     minutes = (diffSeconds Mod 3600) \ 60
//!     seconds = diffSeconds Mod 60
//!     
//!     GetTimeBetween = TimeSerial(hours, minutes, seconds)
//! End Function
//!
//! Public Function RoundToQuarterHour(t As Date) As Date
//!     Dim totalMinutes As Long
//!     Dim roundedMinutes As Long
//!     
//!     totalMinutes = Hour(t) * 60 + Minute(t)
//!     roundedMinutes = ((totalMinutes + 7) \ 15) * 15
//!     
//!     RoundToQuarterHour = TimeSerial(roundedMinutes \ 60, roundedMinutes Mod 60, 0)
//! End Function
//!
//! Public Function TruncateToMinute(t As Date) As Date
//!     TruncateToMinute = TimeSerial(Hour(t), Minute(t), 0)
//! End Function
//!
//! Public Function TruncateToHour(t As Date) As Date
//!     TruncateToHour = TimeSerial(Hour(t), 0, 0)
//! End Function
//!
//! Public Function CreateTimeFromSeconds(totalSeconds As Long) As Date
//!     CreateTimeFromSeconds = TimeSerial(0, 0, totalSeconds)
//! End Function
//!
//! Public Function CreateTimeFromMinutes(totalMinutes As Long) As Date
//!     CreateTimeFromMinutes = TimeSerial(0, totalMinutes, 0)
//! End Function
//! ```
//!
//! ### Example 2: Schedule Manager Module
//! ```vb6
//! ' Module: ScheduleManager
//! ' Manages schedules and time-based operations
//! Option Explicit
//!
//! Private Type ScheduleEntry
//!     Name As String
//!     StartTime As Date
//!     EndTime As Date
//!     Active As Boolean
//! End Type
//!
//! Private m_Schedules() As ScheduleEntry
//! Private m_ScheduleCount As Long
//!
//! Public Sub AddSchedule(name As String, startHour As Integer, startMinute As Integer, _
//!                       endHour As Integer, endMinute As Integer)
//!     ReDim Preserve m_Schedules(m_ScheduleCount)
//!     
//!     m_Schedules(m_ScheduleCount).Name = name
//!     m_Schedules(m_ScheduleCount).StartTime = TimeSerial(startHour, startMinute, 0)
//!     m_Schedules(m_ScheduleCount).EndTime = TimeSerial(endHour, endMinute, 0)
//!     m_Schedules(m_ScheduleCount).Active = True
//!     
//!     m_ScheduleCount = m_ScheduleCount + 1
//! End Sub
//!
//! Public Function IsScheduleActive(name As String) As Boolean
//!     Dim i As Long
//!     Dim currentTime As Date
//!     
//!     currentTime = Time
//!     
//!     For i = 0 To m_ScheduleCount - 1
//!         If m_Schedules(i).Name = name And m_Schedules(i).Active Then
//!             If m_Schedules(i).StartTime <= m_Schedules(i).EndTime Then
//!                 ' Normal schedule (same day)
//!                 IsScheduleActive = (currentTime >= m_Schedules(i).StartTime And _
//!                                   currentTime < m_Schedules(i).EndTime)
//!             Else
//!                 ' Overnight schedule
//!                 IsScheduleActive = (currentTime >= m_Schedules(i).StartTime Or _
//!                                   currentTime < m_Schedules(i).EndTime)
//!             End If
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     IsScheduleActive = False
//! End Function
//!
//! Public Function GetScheduleDuration(name As String) As Date
//!     Dim i As Long
//!     Dim diffSeconds As Long
//!     
//!     For i = 0 To m_ScheduleCount - 1
//!         If m_Schedules(i).Name = name Then
//!             diffSeconds = DateDiff("s", m_Schedules(i).StartTime, m_Schedules(i).EndTime)
//!             If diffSeconds < 0 Then diffSeconds = diffSeconds + 86400  ' Add 24 hours
//!             GetScheduleDuration = TimeSerial(0, 0, diffSeconds)
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     GetScheduleDuration = TimeSerial(0, 0, 0)
//! End Function
//! ```
//!
//! ### Example 3: Time Range Validator Class
//! ```vb6
//! ' Class: TimeRangeValidator
//! ' Validates times against allowed ranges
//! Option Explicit
//!
//! Private m_AllowedStart As Date
//! Private m_AllowedEnd As Date
//! Private m_AllowOvernight As Boolean
//!
//! Public Sub SetAllowedRange(startHour As Integer, startMinute As Integer, _
//!                           endHour As Integer, endMinute As Integer)
//!     m_AllowedStart = TimeSerial(startHour, startMinute, 0)
//!     m_AllowedEnd = TimeSerial(endHour, endMinute, 0)
//!     m_AllowOvernight = (m_AllowedStart > m_AllowedEnd)
//! End Sub
//!
//! Public Function IsTimeAllowed(checkTime As Date) As Boolean
//!     If m_AllowOvernight Then
//!         ' Overnight range (e.g., 10 PM to 6 AM)
//!         IsTimeAllowed = (checkTime >= m_AllowedStart Or checkTime < m_AllowedEnd)
//!     Else
//!         ' Normal range (e.g., 9 AM to 5 PM)
//!         IsTimeAllowed = (checkTime >= m_AllowedStart And checkTime < m_AllowedEnd)
//!     End If
//! End Function
//!
//! Public Function GetNextAllowedTime(fromTime As Date) As Date
//!     If IsTimeAllowed(fromTime) Then
//!         GetNextAllowedTime = fromTime
//!     Else
//!         ' Return start of next allowed window
//!         If fromTime < m_AllowedStart Then
//!             GetNextAllowedTime = m_AllowedStart
//!         Else
//!             ' Must wait until tomorrow's start time
//!             GetNextAllowedTime = DateAdd("d", 1, Date) + m_AllowedStart
//!         End If
//!     End If
//! End Function
//!
//! Public Function GetTimeUntilAllowed(fromTime As Date) As Date
//!     Dim nextAllowed As Date
//!     Dim diffSeconds As Long
//!     
//!     nextAllowed = GetNextAllowedTime(fromTime)
//!     diffSeconds = DateDiff("s", fromTime, nextAllowed)
//!     
//!     GetTimeUntilAllowed = TimeSerial(0, 0, diffSeconds)
//! End Function
//! ```
//!
//! ### Example 4: Time Interval Generator Module
//! ```vb6
//! ' Module: TimeIntervalGenerator
//! ' Generates time intervals for scheduling
//! Option Explicit
//!
//! Public Function GenerateTimeIntervals(startHour As Integer, endHour As Integer, _
//!                                      intervalMinutes As Integer) As Collection
//!     Dim intervals As New Collection
//!     Dim currentTime As Date
//!     Dim endTime As Date
//!     
//!     currentTime = TimeSerial(startHour, 0, 0)
//!     endTime = TimeSerial(endHour, 0, 0)
//!     
//!     Do While currentTime < endTime
//!         intervals.Add currentTime
//!         currentTime = currentTime + TimeSerial(0, intervalMinutes, 0)
//!     Loop
//!     
//!     Set GenerateTimeIntervals = intervals
//! End Function
//!
//! Public Function GenerateWorkDaySchedule(startHour As Integer, endHour As Integer, _
//!                                        taskDurationMinutes As Integer) As Variant
//!     Dim schedule() As Date
//!     Dim currentTime As Date
//!     Dim endTime As Date
//!     Dim index As Long
//!     Dim maxSlots As Long
//!     
//!     currentTime = TimeSerial(startHour, 0, 0)
//!     endTime = TimeSerial(endHour, 0, 0)
//!     
//!     maxSlots = DateDiff("n", currentTime, endTime) \ taskDurationMinutes
//!     ReDim schedule(maxSlots - 1)
//!     
//!     index = 0
//!     Do While currentTime < endTime And index < maxSlots
//!         schedule(index) = currentTime
//!         currentTime = currentTime + TimeSerial(0, taskDurationMinutes, 0)
//!         index = index + 1
//!     Loop
//!     
//!     GenerateWorkDaySchedule = schedule
//! End Function
//!
//! Public Function CreateAppointmentSlots(startHour As Integer, endHour As Integer, _
//!                                       slotDuration As Integer, breakDuration As Integer) As Collection
//!     Dim slots As New Collection
//!     Dim currentTime As Date
//!     Dim endTime As Date
//!     
//!     currentTime = TimeSerial(startHour, 0, 0)
//!     endTime = TimeSerial(endHour, 0, 0)
//!     
//!     Do While currentTime + TimeSerial(0, slotDuration, 0) <= endTime
//!         slots.Add currentTime
//!         currentTime = currentTime + TimeSerial(0, slotDuration + breakDuration, 0)
//!     Loop
//!     
//!     Set CreateAppointmentSlots = slots
//! End Function
//! ```
//!
//! ## Error Handling
//! The `TimeSerial` function can raise the following errors:
//!
//! - **Error 5 (Invalid procedure call)**: If parameters result in invalid time after normalization
//! - **Error 13 (Type mismatch)**: If non-numeric arguments provided
//! - **Error 6 (Overflow)**: If extreme values cause numeric overflow
//!
//! ## Performance Notes
//! - Fast operation - simple calculation
//! - Constant time O(1) complexity
//! - No significant overhead from normalization
//! - Efficient for time arithmetic
//! - Safe to call repeatedly
//!
//! ## Best Practices
//! 1. **Use for time creation** rather than parsing strings
//! 2. **Leverage normalization** for time arithmetic (e.g., negative minutes to subtract)
//! 3. **Store as Date type** for compatibility with other date/time functions
//! 4. **Use 24-hour format** for hour parameter (0-23)
//! 5. **Combine with Date** for complete date/time values
//! 6. **Use for relative times** (intervals, durations)
//! 7. **Format for display** with Format$ function
//! 8. **Document time assumptions** (e.g., time zone, 24-hour format)
//! 9. **Validate inputs** if accepting user-provided values
//! 10. **Use `DateAdd`** for more complex date/time arithmetic
//!
//! ## Comparison Table
//!
//! | Function | Purpose | Parameters | Returns |
//! |----------|---------|------------|---------|
//! | `TimeSerial` | Create time from components | hour, minute, second | Date (time only) |
//! | `DateSerial` | Create date from components | year, month, day | Date (date only) |
//! | `TimeValue` | Parse time from string | time string | Date (time only) |
//! | `DateValue` | Parse date from string | date string | Date (date only) |
//! | `CDate` | Convert to date | expression | Date |
//!
//! ## Platform Notes
//! - Available in VB6, VBA, and `VBScript`
//! - Consistent behavior across platforms
//! - Automatic normalization of values
//! - Date portion always zero (12/30/1899)
//! - Works with standard Date type
//! - Compatible with all date/time functions
//!
//! ## Limitations
//! - Returns only time portion (date is zero)
//! - Cannot directly create date and time together (use `DateSerial` + `TimeSerial`)
//! - No timezone support
//! - No daylight saving time handling
//! - Limited to standard time resolution (seconds)
//! - Cannot create times with milliseconds
//! - Normalization may produce unexpected results if not understood

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_timeserial_basic() {
        let source = r#"
Sub Test()
    t = TimeSerial(12, 30, 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_variable_assignment() {
        let source = r#"
Sub Test()
    Dim myTime As Date
    myTime = TimeSerial(9, 0, 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_all_parameters() {
        let source = r#"
Sub Test()
    result = TimeSerial(14, 45, 30)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_function_return() {
        let source = r#"
Function GetNoon() As Date
    GetNoon = TimeSerial(12, 0, 0)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_addition() {
        let source = r#"
Sub Test()
    newTime = Time + TimeSerial(2, 0, 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_subtraction() {
        let source = r#"
Sub Test()
    earlierTime = Time + TimeSerial(-1, 0, 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_comparison() {
        let source = r#"
Sub Test()
    If Time > TimeSerial(17, 0, 0) Then
        AfterHours
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Opening time: " & TimeSerial(9, 0, 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print TimeSerial(currentHour, currentMinute, currentSecond)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_format() {
        let source = r#"
Sub Test()
    formatted = Format$(TimeSerial(15, 30, 0), "hh:mm AM/PM")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_if_statement() {
        let source = r#"
Sub Test()
    If currentTime >= TimeSerial(9, 0, 0) And currentTime < TimeSerial(17, 0, 0) Then
        BusinessHours
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_select_case() {
        let source = r#"
Sub Test()
    Select Case Time
        Case Is >= TimeSerial(0, 0, 0) And Is < TimeSerial(12, 0, 0)
            Morning
        Case Else
            Afternoon
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_function_argument() {
        let source = r#"
Sub Test()
    Call ProcessTime(TimeSerial(12, 0, 0))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_property_assignment() {
        let source = r#"
Sub Test()
    obj.StartTime = TimeSerial(8, 30, 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_with_statement() {
        let source = r#"
Sub Test()
    With schedule
        .Start = TimeSerial(9, 0, 0)
        .End = TimeSerial(17, 0, 0)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_array_assignment() {
        let source = r#"
Sub Test()
    times(i) = TimeSerial(hours(i), minutes(i), 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_print_statement() {
        let source = r#"
Sub Test()
    Print #1, TimeSerial(10, 15, 30)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_elseif() {
        let source = r#"
Sub Test()
    If x = 1 Then
        y = 1
    ElseIf Time > TimeSerial(18, 0, 0) Then
        y = 2
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_do_while() {
        let source = r#"
Sub Test()
    Do While Time < TimeSerial(17, 0, 0)
        Work
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_do_until() {
        let source = r#"
Sub Test()
    Do Until Time >= TimeSerial(9, 0, 0)
        Wait
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_while_wend() {
        let source = r#"
Sub Test()
    While Time < TimeSerial(12, 0, 0)
        Process
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_iif() {
        let source = r#"
Sub Test()
    greeting = IIf(Time < TimeSerial(12, 0, 0), "Morning", "Afternoon")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_parentheses() {
        let source = r#"
Sub Test()
    result = (TimeSerial(12, 0, 0))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_class_usage() {
        let source = r#"
Sub Test()
    Set calculator = New TimeCalculator
    result = calculator.AddTime(Now, TimeSerial(1, 30, 0))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_negative_values() {
        let source = r#"
Sub Test()
    earlier = Time + TimeSerial(0, -30, 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_hour_minute_second() {
        let source = r#"
Sub Test()
    myTime = TimeSerial(Hour(dt), Minute(dt), Second(dt))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_concatenation() {
        let source = r#"
Sub Test()
    display = "Time: " & TimeSerial(14, 30, 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn test_timeserial_multiple_calls() {
        let source = r#"
Sub Test()
    startTime = TimeSerial(9, 0, 0)
    endTime = TimeSerial(17, 0, 0)
    duration = endTime - startTime
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeSerial"));
    }
}
