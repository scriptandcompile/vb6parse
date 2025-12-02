//! # `Hour` Function
//!
//! Returns an Integer specifying a whole number between 0 and 23, inclusive, representing the hour of the day.
//!
//! ## Syntax
//!
//! ```vb
//! Hour(time)
//! ```
//!
//! ## Parameters
//!
//! - `time` (Required): Any `Variant`, numeric expression, string expression, or any combination that can represent a time. If `time` contains `Null`, `Null` is returned.
//!
//! ## Return Value
//!
//! Returns an `Integer` from 0 to 23 representing the hour of the day. The hour is returned in 24-hour format (0 = midnight, 23 = 11 PM).
//!
//! ## Remarks
//!
//! The `Hour` function extracts the hour component from a date/time value:
//!
//! - Returns values from 0 (midnight) to 23 (11 PM)
//! - Uses 24-hour format, not 12-hour AM/PM format
//! - If `time` is `Null`, the function returns `Null`
//! - If `time` contains only a date with no time component, returns 0
//! - Works with `Date` variables, time strings, and numeric date/time values
//! - Can be used with `Now`, `Time`, `TimeValue`, and other date/time functions
//! - For 12-hour format with AM/PM, use `Format$` function instead
//! - Complements `Minute` and `Second` functions for complete time extraction
//!
//! ## Typical Uses
//!
//! 1. **Time-Based Logic**: Determine if an event occurred during business hours
//! 2. **Scheduling**: Check if current time is within a specific hour range
//! 3. **Data Analysis**: Analyze activity patterns by hour of day
//! 4. **Logging**: Extract hour from timestamp for categorization
//! 5. **User Interface**: Display hour portion of time separately
//! 6. **Time Validation**: Verify time values fall within expected ranges
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Get current hour
//! Dim currentHour As Integer
//! currentHour = Hour(Now)
//! Debug.Print "Current hour: " & currentHour
//!
//! ' Example 2: Extract hour from time string
//! Dim h As Integer
//! h = Hour("14:30:00")    ' Returns 14 (2 PM)
//!
//! ' Example 3: Check for business hours
//! If Hour(Now) >= 9 And Hour(Now) < 17 Then
//!     Debug.Print "Within business hours"
//! End If
//!
//! ' Example 4: Get hour from Date variable
//! Dim myDate As Date
//! myDate = #1/15/2024 3:45:00 PM#
//! Debug.Print Hour(myDate)    ' Returns 15
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Check if current time is within business hours
//! Function IsBusinessHours() As Boolean
//!     Dim h As Integer
//!     h = Hour(Now)
//!     IsBusinessHours = (h >= 9 And h < 17)
//! End Function
//!
//! ' Pattern 2: Categorize time of day
//! Function GetTimeOfDay() As String
//!     Dim h As Integer
//!     h = Hour(Now)
//!     
//!     If h >= 0 And h < 6 Then
//!         GetTimeOfDay = "Night"
//!     ElseIf h >= 6 And h < 12 Then
//!         GetTimeOfDay = "Morning"
//!     ElseIf h >= 12 And h < 18 Then
//!         GetTimeOfDay = "Afternoon"
//!     Else
//!         GetTimeOfDay = "Evening"
//!     End If
//! End Function
//!
//! ' Pattern 3: Format time in 12-hour format with AM/PM
//! Function Format12Hour(timeValue As Date) As String
//!     Dim h As Integer
//!     Dim ampm As String
//!     
//!     h = Hour(timeValue)
//!     
//!     If h >= 12 Then
//!         ampm = "PM"
//!         If h > 12 Then h = h - 12
//!     Else
//!         ampm = "AM"
//!         If h = 0 Then h = 12
//!     End If
//!     
//!     Format12Hour = h & ":" & Right$("0" & Minute(timeValue), 2) & " " & ampm
//! End Function
//!
//! ' Pattern 4: Round time to nearest hour
//! Function RoundToNearestHour(timeValue As Date) As Date
//!     If Minute(timeValue) >= 30 Then
//!         RoundToNearestHour = DateSerial(Year(timeValue), Month(timeValue), Day(timeValue)) + _
//!                              TimeSerial(Hour(timeValue) + 1, 0, 0)
//!     Else
//!         RoundToNearestHour = DateSerial(Year(timeValue), Month(timeValue), Day(timeValue)) + _
//!                              TimeSerial(Hour(timeValue), 0, 0)
//!     End If
//! End Function
//!
//! ' Pattern 5: Calculate hours elapsed
//! Function HoursElapsed(startTime As Date, endTime As Date) As Integer
//!     Dim totalHours As Double
//!     totalHours = (endTime - startTime) * 24
//!     HoursElapsed = Int(totalHours)
//! End Function
//!
//! ' Pattern 6: Check if after specific hour
//! Function IsAfterHour(checkTime As Date, hourThreshold As Integer) As Boolean
//!     IsAfterHour = (Hour(checkTime) >= hourThreshold)
//! End Function
//!
//! ' Pattern 7: Get hour range description
//! Function GetHourRange(timeValue As Date) As String
//!     Dim h As Integer
//!     h = Hour(timeValue)
//!     GetHourRange = h & ":00 - " & h & ":59"
//! End Function
//!
//! ' Pattern 8: Validate time is within allowed hours
//! Function IsWithinAllowedHours(checkTime As Date, startHour As Integer, endHour As Integer) As Boolean
//!     Dim h As Integer
//!     h = Hour(checkTime)
//!     
//!     If startHour <= endHour Then
//!         IsWithinAllowedHours = (h >= startHour And h <= endHour)
//!     Else
//!         ' Handle overnight range (e.g., 22:00 to 6:00)
//!         IsWithinAllowedHours = (h >= startHour Or h <= endHour)
//!     End If
//! End Function
//!
//! ' Pattern 9: Count hourly occurrences
//! Sub CountByHour(timestamps() As Date, hourCounts() As Long)
//!     Dim i As Long
//!     Dim h As Integer
//!     
//!     ReDim hourCounts(0 To 23)
//!     
//!     For i = LBound(timestamps) To UBound(timestamps)
//!         h = Hour(timestamps(i))
//!         hourCounts(h) = hourCounts(h) + 1
//!     Next i
//! End Sub
//!
//! ' Pattern 10: Create time from hour
//! Function CreateTimeFromHour(hourValue As Integer, Optional minuteValue As Integer = 0) As Date
//!     CreateTimeFromHour = TimeSerial(hourValue, minuteValue, 0)
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Business hours scheduler
//! Public Class BusinessScheduler
//!     Private Const BUSINESS_START As Integer = 9
//!     Private Const BUSINESS_END As Integer = 17
//!     Private Const LUNCH_START As Integer = 12
//!     Private Const LUNCH_END As Integer = 13
//!     
//!     Public Function IsAvailable(checkTime As Date) As Boolean
//!         Dim h As Integer
//!         h = Hour(checkTime)
//!         
//!         If h < BUSINESS_START Or h >= BUSINESS_END Then
//!             IsAvailable = False
//!         ElseIf h >= LUNCH_START And h < LUNCH_END Then
//!             IsAvailable = False
//!         Else
//!             IsAvailable = True
//!         End If
//!     End Function
//!     
//!     Public Function GetNextAvailableSlot(currentTime As Date) As Date
//!         Dim h As Integer
//!         Dim nextSlot As Date
//!         
//!         h = Hour(currentTime)
//!         
//!         If h < BUSINESS_START Then
//!             nextSlot = DateSerial(Year(currentTime), Month(currentTime), Day(currentTime)) + _
//!                        TimeSerial(BUSINESS_START, 0, 0)
//!         ElseIf h >= LUNCH_START And h < LUNCH_END Then
//!             nextSlot = DateSerial(Year(currentTime), Month(currentTime), Day(currentTime)) + _
//!                        TimeSerial(LUNCH_END, 0, 0)
//!         ElseIf h >= BUSINESS_END Then
//!             nextSlot = DateSerial(Year(currentTime), Month(currentTime), Day(currentTime) + 1) + _
//!                        TimeSerial(BUSINESS_START, 0, 0)
//!         Else
//!             nextSlot = currentTime
//!         End If
//!         
//!         GetNextAvailableSlot = nextSlot
//!     End Function
//! End Class
//!
//! ' Example 2: Hourly activity analyzer
//! Public Class ActivityAnalyzer
//!     Private m_hourlyData(0 To 23) As Long
//!     
//!     Public Sub RecordActivity(activityTime As Date)
//!         Dim h As Integer
//!         h = Hour(activityTime)
//!         m_hourlyData(h) = m_hourlyData(h) + 1
//!     End Sub
//!     
//!     Public Function GetPeakHour() As Integer
//!         Dim i As Integer
//!         Dim maxCount As Long
//!         Dim peakHour As Integer
//!         
//!         maxCount = 0
//!         peakHour = 0
//!         
//!         For i = 0 To 23
//!             If m_hourlyData(i) > maxCount Then
//!                 maxCount = m_hourlyData(i)
//!                 peakHour = i
//!             End If
//!         Next i
//!         
//!         GetPeakHour = peakHour
//!     End Function
//!     
//!     Public Function GetActivityInRange(startHour As Integer, endHour As Integer) As Long
//!         Dim i As Integer
//!         Dim total As Long
//!         
//!         total = 0
//!         For i = startHour To endHour
//!             total = total + m_hourlyData(i)
//!         Next i
//!         
//!         GetActivityInRange = total
//!     End Function
//! End Class
//!
//! ' Example 3: Time slot manager
//! Public Class TimeSlotManager
//!     Private Type TimeSlot
//!         StartHour As Integer
//!         EndHour As Integer
//!         IsAvailable As Boolean
//!     End Type
//!     
//!     Private m_slots() As TimeSlot
//!     
//!     Public Sub Initialize()
//!         Dim i As Integer
//!         ReDim m_slots(0 To 23)
//!         
//!         For i = 0 To 23
//!             m_slots(i).StartHour = i
//!             m_slots(i).EndHour = i
//!             m_slots(i).IsAvailable = True
//!         Next i
//!     End Sub
//!     
//!     Public Function BookSlot(bookingTime As Date) As Boolean
//!         Dim h As Integer
//!         h = Hour(bookingTime)
//!         
//!         If m_slots(h).IsAvailable Then
//!             m_slots(h).IsAvailable = False
//!             BookSlot = True
//!         Else
//!             BookSlot = False
//!         End If
//!     End Function
//!     
//!     Public Function FindNextAvailable(afterTime As Date) As Integer
//!         Dim h As Integer
//!         Dim i As Integer
//!         
//!         h = Hour(afterTime)
//!         
//!         For i = h To 23
//!             If m_slots(i).IsAvailable Then
//!                 FindNextAvailable = i
//!                 Exit Function
//!             End If
//!         Next i
//!         
//!         FindNextAvailable = -1  ' No available slots
//!     End Function
//! End Class
//!
//! ' Example 4: Shift time calculator
//! Public Function CalculateShiftHours(clockIn As Date, clockOut As Date) As Double
//!     Dim totalHours As Double
//!     Dim inHour As Integer
//!     Dim outHour As Integer
//!     
//!     inHour = Hour(clockIn)
//!     outHour = Hour(clockOut)
//!     
//!     If clockOut < clockIn Then
//!         ' Overnight shift
//!         totalHours = (clockOut + 1 - clockIn) * 24
//!     Else
//!         totalHours = (clockOut - clockIn) * 24
//!     End If
//!     
//!     CalculateShiftHours = totalHours
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! The Hour function can raise errors in certain situations:
//!
//! - **Type Mismatch (Error 13)**: If the argument cannot be interpreted as a date/time value
//! - **Invalid Procedure Call (Error 5)**: If the date value is invalid
//! - **Null Propagation**: If the argument is `Null`, the function returns `Null` (not an error)
//!
//! ```vb
//! On Error Resume Next
//! Dim h As Integer
//! h = Hour(someValue)
//! If Err.Number <> 0 Then
//!     Debug.Print "Error extracting hour: " & Err.Description
//!     h = 0  ' Default value
//!     Err.Clear
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: Hour extraction is a very fast, native operation
//! - **No Overhead**: Minimal performance impact even in tight loops
//! - **Caching**: If checking the same time repeatedly, cache the `Hour()` result
//! - **Comparison**: Direct `Hour()` comparisons are faster than full time comparisons
//!
//! ## Best Practices
//!
//! 1. **24-Hour Format**: Remember `Hour` returns 0-23, not 1-12 with AM/PM
//! 2. **Midnight**: `Hour(#12:00:00 AM#)` returns 0, not 12
//! 3. **Noon**: `Hour(#12:00:00 PM#)` returns 12
//! 4. **Validation**: Validate hour ranges (0-23) when accepting user input
//! 5. **Date-Only Values**: The hour component of a date-only value (no time) is always 0
//! 6. **Null Handling**: Check for `Null` when working with `Variant` date values
//! 7. **Time Zones**: `Hour` doesn't handle time zones; use explicit conversion if needed
//!
//! ## Comparison with Other Functions
//!
//! | Function | Purpose | Return Range |
//! |----------|---------|--------------|
//! | `Hour` | Extract hour from time | 0-23 (24-hour) |
//! | `Minute` | Extract minute from time | 0-59 |
//! | `Second` | Extract second from time | 0-59 |
//! | `Day` | Extract day from date | 1-31 |
//! | `Month` | Extract month from date | 1-12 |
//! | `Year` | Extract year from date | 100-9999 |
//! | `Weekday` | Get day of week | 1-7 |
//!
//! ## Conversion Examples
//!
//! ```vb
//! ' Convert 24-hour to 12-hour with AM/PM
//! Function To12Hour(hour24 As Integer) As String
//!     Dim hour12 As Integer
//!     Dim ampm As String
//!     
//!     If hour24 >= 12 Then
//!         ampm = "PM"
//!         hour12 = IIf(hour24 = 12, 12, hour24 - 12)
//!     Else
//!         ampm = "AM"
//!         hour12 = IIf(hour24 = 0, 12, hour24)
//!     End If
//!     
//!     To12Hour = hour12 & " " & ampm
//! End Function
//!
//! ' Convert 12-hour to 24-hour
//! Function To24Hour(hour12 As Integer, isPM As Boolean) As Integer
//!     If isPM Then
//!         To24Hour = IIf(hour12 = 12, 12, hour12 + 12)
//!     Else
//!         To24Hour = IIf(hour12 = 12, 0, hour12)
//!     End If
//! End Function
//! ```
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Consistent behavior across Windows platforms
//! - Uses system locale for time interpretation when parsing strings
//! - Returns Integer (2-byte signed integer, range: -32,768 to 32,767)
//!
//! ## Limitations
//!
//! - Returns 24-hour format only (0-23)
//! - No built-in AM/PM support (use Format$ for that)
//! - Cannot extract fractional hours (use full date arithmetic for precision)
//! - No time zone awareness (always uses local time interpretation)
//! - Date-only values always return hour 0
//! - Cannot handle dates before 1/1/100 or after 12/31/9999
//!
//! ## Related Functions
//!
//! - `Minute`: Returns the minute of the hour (0-59)
//! - `Second`: Returns the second of the minute (0-59)
//! - `Now`: Returns the current date and time
//! - `Time`: Returns the current system time
//! - `TimeSerial`: Returns a Date for a specified hour, minute, and second
//! - `TimeValue`: Converts a string to a time value
//! - `Format`: Formats a date/time with custom formatting including AM/PM

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_hour_basic() {
        let source = r#"
Sub Test()
    h = Hour(Now)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_in_function() {
        let source = r#"
Function GetCurrentHour() As Integer
    GetCurrentHour = Hour(Now)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_if_statement() {
        let source = r#"
Sub Test()
    If Hour(Now) >= 9 And Hour(Now) < 17 Then
        Debug.Print "Business hours"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print Hour(Time)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_select_case() {
        let source = r#"
Sub Test()
    Select Case Hour(Now)
        Case 0 To 5
            Debug.Print "Night"
        Case 6 To 11
            Debug.Print "Morning"
        Case Else
            Debug.Print "Afternoon/Evening"
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_for_loop() {
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 0 To 23
        If Hour(Now) = i Then Exit For
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_do_loop() {
        let source = r#"
Sub Test()
    Do While Hour(Now) < 17
        DoWork
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_class_member() {
        let source = r#"
Private Sub Class_Initialize()
    m_hour = Hour(Now)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_type_field() {
        let source = r#"
Sub Test()
    Dim timeInfo As TimeType
    timeInfo.currentHour = Hour(Now)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_collection_add() {
        let source = r#"
Sub Test()
    Dim col As New Collection
    col.Add Hour(Now)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_with_statement() {
        let source = r#"
Sub Test()
    With timeObject
        .HourValue = Hour(.TimeStamp)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Current hour: " & Hour(Now)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_property() {
        let source = r#"
Property Get CurrentHour() As Integer
    CurrentHour = Hour(Now)
End Property
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_concatenation() {
        let source = r#"
Sub Test()
    msg = "The hour is " & Hour(Now)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_for_each() {
        let source = r#"
Sub Test()
    Dim dt As Variant
    For Each dt In dateCollection
        Debug.Print Hour(dt)
    Next
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    h = Hour(someDate)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_comparison() {
        let source = r#"
Sub Test()
    If Hour(startTime) < Hour(endTime) Then
        Debug.Print "Same day"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_array_assignment() {
        let source = r#"
Sub Test()
    Dim hours(1 To 10) As Integer
    hours(1) = Hour(Now)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_function_argument() {
        let source = r#"
Sub Test()
    ProcessHour Hour(Now)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_nested_call() {
        let source = r#"
Sub Test()
    result = CStr(Hour(Now))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_iif() {
        let source = r#"
Sub Test()
    period = IIf(Hour(Now) < 12, "AM", "PM")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_time_literal() {
        let source = r#"
Sub Test()
    h = Hour(#3:45:00 PM#)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_variable() {
        let source = r#"
Sub Test()
    Dim myTime As Date
    myTime = Now
    h = Hour(myTime)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_parentheses() {
        let source = r#"
Sub Test()
    value = (Hour(Now))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_business_logic() {
        let source = r#"
Function IsBusinessHours() As Boolean
    Dim h As Integer
    h = Hour(Now)
    IsBusinessHours = (h >= 9 And h < 17)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_time_range() {
        let source = r#"
Sub Test()
    Dim h As Integer
    h = Hour(Now)
    If h >= 0 And h <= 23 Then
        Debug.Print "Valid hour"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_hour_math_operation() {
        let source = r#"
Sub Test()
    hoursUntilMidnight = 24 - Hour(Now)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Hour"));
        assert!(text.contains("Identifier"));
    }
}
