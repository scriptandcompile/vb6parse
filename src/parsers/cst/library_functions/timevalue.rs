//! VB6 `TimeValue` Function
//!
//! The `TimeValue` function returns a Variant (Date) containing the time value represented by a string.
//!
//! ## Syntax
//! ```vb6
//! TimeValue(time)
//! ```
//!
//! ## Parameters
//! - `time`: Required. String expression representing a time (e.g., "14:30:00", "2:30 PM"). Can also be a Variant containing a string or a numeric value representing a time.
//!
//! ## Returns
//! Returns a `Variant` of subtype `Date` containing the time value. The date portion is set to zero (December 30, 1899).
//!
//! ## Remarks
//! The `TimeValue` function parses a string and returns the corresponding time value:
//!
//! - **String input**: Accepts time strings in various formats (e.g., "14:30", "2:30 PM", "23:59:59")
//! - **Date portion**: Always returns zero date (12/30/1899)
//! - **Locale aware**: Accepts time formats based on system locale
//! - **Numeric input**: Accepts numeric values representing fractional days (e.g., 0.5 = 12:00 PM)
//! - **Null handling**: Returns Null if input is Null
//! - **Error handling**: Error 13 (Type mismatch) if string cannot be parsed as a time
//! - **No milliseconds**: Only parses up to seconds
//! - **Companion to `DateValue`**: Use `DateValue` for date-only parsing
//! - **Type returned**: Returns Variant (Date), not a string
//!
//! ### Accepted Time Formats
//! - 24-hour: "14:30", "23:59:59"
//! - 12-hour: "2:30 PM", "11:59:59 PM"
//! - With/without seconds: "8:00", "8:00:00"
//! - With/without AM/PM: "7:15", "7:15 AM"
//! - Numeric: 0.75 (6:00 PM)
//!
//! ### Invalid Inputs
//! - "25:00" (invalid hour)
//! - "13:60" (invalid minute)
//! - "not a time" (not parseable)
//!
//! ## Typical Uses
//! 1. **Parse User Input**: Convert time strings to time values
//! 2. **Time Comparison**: Compare parsed times to other time values
//! 3. **Scheduling**: Store and use times from configuration or user entry
//! 4. **Time Calculations**: Perform arithmetic with parsed times
//! 5. **Validation**: Check if input is a valid time
//! 6. **Database Import**: Parse time strings from data sources
//! 7. **Time Filtering**: Filter records by parsed time
//! 8. **Time Formatting**: Convert string to time for display or calculation
//!
//! ## Basic Examples
//!
//! ### Example 1: Parse Time String
//! ```vb6
//! Sub ParseTime()
//!     Dim t As Date
//!     t = TimeValue("14:30:00")
//!     MsgBox "Parsed time: " & Format$(t, "hh:mm AM/PM")
//! End Sub
//! ```
//!
//! ### Example 2: Parse 12-hour Time
//! ```vb6
//! Function ParseNoon() As Date
//!     ParseNoon = TimeValue("12:00 PM")
//! End Function
//! ```
//!
//! ### Example 3: Validate Time Input
//! ```vb6
//! Function IsValidTime(inputStr As String) As Boolean
//!     On Error Resume Next
//!     Dim t As Date
//!     t = TimeValue(inputStr)
//!     IsValidTime = (Err.Number = 0)
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ### Example 4: Compare Parsed Time
//! ```vb6
//! Function IsAfterNoon(timeStr As String) As Boolean
//!     IsAfterNoon = (TimeValue(timeStr) > TimeValue("12:00 PM"))
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Parse and Add Minutes
//! ```vb6
//! Function AddMinutes(timeStr As String, minutes As Integer) As Date
//!     AddMinutes = TimeValue(timeStr) + TimeSerial(0, minutes, 0)
//! End Function
//! ```
//!
//! ### Pattern 2: Parse and Compare to Now
//! ```vb6
//! Function IsPast(timeStr As String) As Boolean
//!     IsPast = (TimeValue(timeStr) < Time)
//! End Function
//! ```
//!
//! ### Pattern 3: Parse and Format
//! ```vb6
//! Function FormatParsedTime(timeStr As String) As String
//!     FormatParsedTime = Format$(TimeValue(timeStr), "hh:mm:ss")
//! End Function
//! ```
//!
//! ### Pattern 4: Parse with Error Handling
//! ```vb6
//! Function TryParseTime(timeStr As String, ByRef result As Date) As Boolean
//!     On Error Resume Next
//!     result = TimeValue(timeStr)
//!     TryParseTime = (Err.Number = 0)
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ### Pattern 5: Parse Numeric Time
//! ```vb6
//! Function ParseNumericTime(numericValue As Double) As Date
//!     ParseNumericTime = TimeValue(numericValue)
//! End Function
//! ```
//!
//! ### Pattern 6: Parse and Set Property
//! ```vb6
//! Sub SetStartTime(obj As Object, timeStr As String)
//!     obj.StartTime = TimeValue(timeStr)
//! End Sub
//! ```
//!
//! ### Pattern 7: Parse Array of Times
//! ```vb6
//! Function ParseTimeArray(timeStrings() As String) As Variant
//!     Dim i As Integer
//!     Dim times() As Date
//!     ReDim times(LBound(timeStrings) To UBound(timeStrings))
//!     For i = LBound(timeStrings) To UBound(timeStrings)
//!         times(i) = TimeValue(timeStrings(i))
//!     Next i
//!     ParseTimeArray = times
//! End Function
//! ```
//!
//! ### Pattern 8: Filter Valid Times
//! ```vb6
//! Function FilterValidTimes(timeStrings() As String) As Collection
//!     Dim validTimes As New Collection
//!     Dim i As Integer
//!     For i = LBound(timeStrings) To UBound(timeStrings)
//!         If IsValidTime(timeStrings(i)) Then
//!             validTimes.Add TimeValue(timeStrings(i))
//!         End If
//!     Next i
//!     Set FilterValidTimes = validTimes
//! End Function
//! ```
//!
//! ### Pattern 9: Parse and Use in Schedule
//! ```vb6
//! Sub AddScheduleEntry(schedule As Collection, timeStr As String, description As String)
//!     Dim entry As New Collection
//!     entry.Add TimeValue(timeStr), "Time"
//!     entry.Add description, "Description"
//!     schedule.Add entry
//! End Sub
//! ```
//!
//! ### Pattern 10: Parse and Compare Range
//! ```vb6
//! Function IsWithinRange(timeStr As String, startStr As String, endStr As String) As Boolean
//!     Dim t As Date
//!     t = TimeValue(timeStr)
//!     IsWithinRange = (t >= TimeValue(startStr) And t <= TimeValue(endStr))
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Time Parser Class
//! ```vb6
//! ' Class: TimeParser
//! ' Parses and validates time strings
//! Option Explicit
//!
//! Public Function Parse(timeStr As String) As Date
//!     Parse = TimeValue(timeStr)
//! End Function
//!
//! Public Function TryParse(timeStr As String, ByRef result As Date) As Boolean
//!     On Error Resume Next
//!     result = TimeValue(timeStr)
//!     TryParse = (Err.Number = 0)
//!     On Error GoTo 0
//! End Function
//!
//! Public Function IsValid(timeStr As String) As Boolean
//!     On Error Resume Next
//!     Dim t As Date
//!     t = TimeValue(timeStr)
//!     IsValid = (Err.Number = 0)
//!     On Error GoTo 0
//! End Function
//!
//! Public Function Format(timeStr As String, formatStr As String) As String
//!     Format = Format$(TimeValue(timeStr), formatStr)
//! End Function
//! ```
//!
//! ### Example 2: Schedule Importer Module
//! ```vb6
//! ' Module: ScheduleImporter
//! ' Imports and parses schedule times from data
//! Option Explicit
//!
//! Public Function ImportSchedule(times() As String, descriptions() As String) As Collection
//!     Dim schedule As New Collection
//!     Dim i As Integer
//!     For i = LBound(times) To UBound(times)
//!         Dim entry As New Collection
//!         entry.Add TimeValue(times(i)), "Time"
//!         entry.Add descriptions(i), "Description"
//!         schedule.Add entry
//!     Next i
//!     Set ImportSchedule = schedule
//! End Function
//!
//! Public Function GetEarliestTime(times() As String) As Date
//!     Dim minTime As Date
//!     Dim i As Integer
//!     minTime = TimeValue(times(LBound(times)))
//!     For i = LBound(times) + 1 To UBound(times)
//!         If TimeValue(times(i)) < minTime Then minTime = TimeValue(times(i))
//!     Next i
//!     GetEarliestTime = minTime
//! End Function
//!
//! Public Function GetLatestTime(times() As String) As Date
//!     Dim maxTime As Date
//!     Dim i As Integer
//!     maxTime = TimeValue(times(LBound(times)))
//!     For i = LBound(times) + 1 To UBound(times)
//!         If TimeValue(times(i)) > maxTime Then maxTime = TimeValue(times(i))
//!     Next i
//!     GetLatestTime = maxTime
//! End Function
//! ```
//!
//! ### Example 3: Time Validator Class
//! ```vb6
//! ' Class: TimeValidator
//! ' Validates and normalizes time strings
//! Option Explicit
//!
//! Public Function IsValid(timeStr As String) As Boolean
//!     On Error Resume Next
//!     Dim t As Date
//!     t = TimeValue(timeStr)
//!     IsValid = (Err.Number = 0)
//!     On Error GoTo 0
//! End Function
//!
//! Public Function Normalize(timeStr As String) As String
//!     On Error Resume Next
//!     Dim t As Date
//!     t = TimeValue(timeStr)
//!     If Err.Number = 0 Then
//!         Normalize = Format$(t, "hh:mm:ss")
//!     Else
//!         Normalize = ""
//!     End If
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ### Example 4: Time Range Analyzer Module
//! ```vb6
//! ' Module: TimeRangeAnalyzer
//! ' Analyzes and compares time ranges
//! Option Explicit
//!
//! Public Function GetTimeRange(times() As String) As String
//!     Dim minTime As Date, maxTime As Date
//!     Dim i As Integer
//!     minTime = TimeValue(times(LBound(times)))
//!     maxTime = minTime
//!     For i = LBound(times) + 1 To UBound(times)
//!         If TimeValue(times(i)) < minTime Then minTime = TimeValue(times(i))
//!         If TimeValue(times(i)) > maxTime Then maxTime = TimeValue(times(i))
//!     Next i
//!     GetTimeRange = Format$(minTime, "hh:mm") & " - " & Format$(maxTime, "hh:mm")
//! End Function
//!
//! Public Function IsWithinRange(timeStr As String, startStr As String, endStr As String) As Boolean
//!     Dim t As Date
//!     t = TimeValue(timeStr)
//!     IsWithinRange = (t >= TimeValue(startStr) And t <= TimeValue(endStr))
//! End Function
//! ```
//!
//! ## Error Handling
//! The `TimeValue` function can raise the following errors:
//!
//! - **Error 13 (Type mismatch)**: If the string cannot be parsed as a time
//! - **Returns Null**: If the input is Null (not an error)
//!
//! ## Performance Notes
//! - Fast operation - simple parsing
//! - Constant time O(1) complexity for valid strings
//! - Locale-dependent parsing speed
//! - Safe to call repeatedly
//!
//! ## Best Practices
//! 1. **Validate input** before calling if user-provided
//! 2. **Handle Null** explicitly when working with nullable fields
//! 3. **Use Format$** for display after parsing
//! 4. **Test with different locales** for internationalization
//! 5. **Use `TryParse` pattern** for robust error handling
//! 6. **Document expected formats** for maintainability
//! 7. **Store as Date type** for calculations
//! 8. **Use with `TimeSerial`** for arithmetic
//! 9. **Test edge cases** (midnight, noon, invalid times)
//! 10. **Combine with `DateValue`** for full date/time parsing
//!
//! ## Comparison Table
//!
//! | Function | Purpose | Input | Returns |
//! |----------|---------|-------|---------|
//! | `TimeValue` | Parse time from string | time string | Date (time only) |
//! | `DateValue` | Parse date from string | date string | Date (date only) |
//! | `TimeSerial` | Create time from components | hour, minute, second | Date (time only) |
//! | `DateSerial` | Create date from components | year, month, day | Date (date only) |
//! | `CDate` | Convert to date | expression | Date |
//!
//! ## Platform Notes
//! - Available in VB6, VBA, and `VBScript`
//! - Consistent behavior across platforms
//! - Locale-aware parsing
//! - Date portion always zero (12/30/1899)
//! - Compatible with all date/time functions
//!
//! ## Limitations
//! - Only parses up to seconds (no milliseconds)
//! - Locale-dependent parsing (may fail on some formats)
//! - Cannot parse date and time together (use `CDate` for full datetime)
//! - No timezone support
//! - No daylight saving time handling
//! - Returns only time portion (date is zero)
//! - Error on invalid input (not Null)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_timevalue_basic() {
        let source = r#"
Sub Test()
    t = TimeValue("14:30:00")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_12_hour() {
        let source = r#"
Sub Test()
    noon = TimeValue("12:00 PM")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_variable_assignment() {
        let source = r#"
Sub Test()
    Dim myTime As Date
    myTime = TimeValue("8:15")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_function_return() {
        let source = r#"
Function ParseTime(s As String) As Date
    ParseTime = TimeValue(s)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_add_minutes() {
        let source = r#"
Sub Test()
    newTime = TimeValue("10:00") + TimeSerial(0, 30, 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_comparison() {
        let source = r#"
Sub Test()
    If TimeValue("15:00") > TimeValue("12:00") Then
        Afternoon
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Parsed: " & TimeValue("7:45 AM")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print TimeValue("23:59:59")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_format() {
        let source = r#"
Sub Test()
    formatted = Format$(TimeValue("18:30"), "hh:mm AM/PM")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_if_statement() {
        let source = r#"
Sub Test()
    If TimeValue("6:00 PM") > TimeValue("12:00 PM") Then
        Evening
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_select_case() {
        let source = r#"
Sub Test()
    Select Case TimeValue("8:00")
        Case Is < TimeValue("12:00")
            Morning
        Case Else
            Afternoon
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_function_argument() {
        let source = r#"
Sub Test()
    Call SetStartTime(obj, "9:00 AM")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("SetStartTime"));
        // The function SetStartTime uses TimeValue internally
    }

    #[test]
    fn test_timevalue_property_assignment() {
        let source = r#"
Sub Test()
    obj.Time = TimeValue("13:45")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_with_statement() {
        let source = r#"
Sub Test()
    With schedule
        .Start = TimeValue("8:00")
        .End = TimeValue("17:00")
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_array_assignment() {
        let source = r#"
Sub Test()
    times(i) = TimeValue(timeStrings(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_print_statement() {
        let source = r#"
Sub Test()
    Print #1, TimeValue("10:15:30")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_elseif() {
        let source = r#"
Sub Test()
    If x = 1 Then
        y = 1
    ElseIf TimeValue("18:00") > TimeValue("12:00") Then
        y = 2
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_do_while() {
        let source = r#"
Sub Test()
    Do While TimeValue("15:00") > TimeValue("12:00")
        Process
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_do_until() {
        let source = r#"
Sub Test()
    Do Until TimeValue("8:00") >= TimeValue("17:00")
        Wait
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_while_wend() {
        let source = r#"
Sub Test()
    While TimeValue("9:00") < TimeValue("17:00")
        Work
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_iif() {
        let source = r#"
Sub Test()
    greeting = IIf(TimeValue("8:00") < TimeValue("12:00"), "Morning", "Afternoon")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_parentheses() {
        let source = r#"
Sub Test()
    result = (TimeValue("12:00"))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn test_timevalue_class_usage() {
        let source = r#"
Sub Test()
    Set parser = New TimeParser
    t = parser.Parse("14:30")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Parse"));
        // The Parse method uses TimeValue internally
    }

    #[test]
    fn test_timevalue_numeric_input() {
        let source = r#"
Sub Test()
    t = TimeValue(0.5)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("TimeValue"));
    }
}
