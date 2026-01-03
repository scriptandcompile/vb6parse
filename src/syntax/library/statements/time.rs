//! # Time Statement
//!
//! Sets the system time.
//!
//! ## Syntax
//!
//! ```vb
//! Time = time
//! ```
//!
//! ## Parts
//!
//! - **time**: Required. Any numeric expression, string expression, or any combination that can represent a time.
//!
//! ## Remarks
//!
//! - **System Time**: The Time statement sets the computer's system time to the specified time value.
//! - **Time Format**: Accepts times in various formats including "HH:MM:SS", "HH:MM", or numeric values representing time.
//! - **24-Hour Format**: You can use 24-hour format (e.g., "13:30" for 1:30 PM) or 12-hour format with AM/PM.
//! - **Permissions**: Changing the system time may require administrator privileges on some operating systems.
//! - **String Expression**: When using a string, it should be in a valid time format that VB6 can interpret.
//! - **Numeric Expression**: Numeric values represent the fractional portion of a day (e.g., 0.5 = noon).
//! - **Current Date Preserved**: Setting the time does not affect the system date.
//! - **Time Function**: Use the Time function (without assignment) to retrieve the current system time.
//! - **Now Function**: The Now function returns both date and time; Time$ returns only the time portion.
//! - **Error Handling**: Invalid time values will generate a run-time error.
//!
//! ## Common Uses
//!
//! - **Time Synchronization**: Set system time from network time server
//! - **Testing**: Set specific times for testing time-dependent code
//! - **Kiosk Applications**: Reset time for demo or kiosk systems
//! - **Simulation**: Simulate different times of day for testing
//! - **Time Adjustment**: Correct system time drift
//!
//! ## Examples
//!
//! ### Set Time to Specific Hour and Minute
//!
//! ```vb
//! Time = "14:30:00"  ' Set to 2:30 PM
//! ```
//!
//! ### Set Time Using String
//!
//! ```vb
//! Time = "9:15 AM"
//! ```
//!
//! ### Set Time to Midnight
//!
//! ```vb
//! Time = "00:00:00"
//! ```
//!
//! ### Set Time to Noon
//!
//! ```vb
//! Time = "12:00:00"
//! ```
//!
//! ### Set Time Using Variable
//!
//! ```vb
//! Dim newTime As String
//! newTime = "15:45:30"
//! Time = newTime
//! ```
//!
//! ### Set Time Using `TimeValue` Function
//!
//! ```vb
//! Time = TimeValue("3:30 PM")
//! ```
//!
//! ### Set Time Using Current Time Plus Offset
//!
//! ```vb
//! Time = Time + TimeValue("00:15:00")  ' Add 15 minutes
//! ```
//!
//! ### Set Time from User Input
//!
//! ```vb
//! Dim userTime As String
//! userTime = InputBox("Enter new time (HH:MM:SS):")
//! If IsDate(userTime) Then
//!     Time = userTime
//! Else
//!     MsgBox "Invalid time format"
//! End If
//! ```
//!
//! ### Set Time with Error Handling
//!
//! ```vb
//! On Error Resume Next
//! Time = "10:30:00"
//! If Err.Number <> 0 Then
//!     MsgBox "Failed to set time: " & Err.Description
//! End If
//! On Error GoTo 0
//! ```
//!
//! ### Set Time Using Now Function
//!
//! ```vb
//! Time = Now  ' Sets time to current time (redundant but valid)
//! ```
//!
//! ### Set Time in Sub
//!
//! ```vb
//! Sub SetApplicationTime()
//!     Time = "08:00:00"  ' Set to 8 AM
//! End Sub
//! ```
//!
//! ### Set Time Conditionally
//!
//! ```vb
//! If Hour(Time) > 17 Then
//!     Time = "08:00:00"  ' Reset to morning
//! End If
//! ```
//!
//! ### Set Time Using `TimeSerial`
//!
//! ```vb
//! Time = TimeSerial(14, 30, 0)  ' 2:30 PM
//! ```
//!
//! ### Set Time with Concatenation
//!
//! ```vb
//! Dim hours As String
//! Dim minutes As String
//! hours = "09"
//! minutes = "45"
//! Time = hours & ":" & minutes & ":00"
//! ```
//!
//! ### Set Time for Testing
//!
//! ```vb
//! ' Set specific time for testing time-dependent code
//! Time = "23:59:59"  ' One second before midnight
//! TestMidnightRollover
//! ```
//!
//! ### Set Time in Class Module
//!
//! ```vb
//! Private Sub Class_Initialize()
//!     Time = "12:00:00"  ' Reset to noon on initialization
//! End Sub
//! ```
//!
//! ### Set Time Using Format
//!
//! ```vb
//! Dim timeStr As String
//! timeStr = Format(Now, "hh:mm:ss")
//! Time = timeStr
//! ```
//!
//! ### Set Time in Loop
//!
//! ```vb
//! For i = 0 To 23
//!     Time = TimeSerial(i, 0, 0)
//!     ProcessHourlyTask
//! Next i
//! ```
//!
//! ### Set Time with Validation
//!
//! ```vb
//! Function SetSystemTime(newTime As String) As Boolean
//!     On Error GoTo ErrorHandler
//!     
//!     If IsDate(newTime) Then
//!         Time = newTime
//!         SetSystemTime = True
//!     Else
//!         SetSystemTime = False
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     SetSystemTime = False
//! End Function
//! ```
//!
//! ## Important Notes
//!
//! - **Administrator Rights**: Setting system time may require elevated permissions
//! - **System Impact**: Changing system time affects all applications and scheduled tasks
//! - **Time Zones**: Time is set in local time zone, not UTC
//! - **Date Unchanged**: Only the time portion is modified; the date remains unchanged
//! - **Validation**: Always validate user input before setting system time
//! - **Error Handling**: Use error handling as time setting can fail due to permissions
//! - **Testing Only**: In production, avoid changing system time; use application-level time variables instead
//! - **Numeric Values**: 0 = midnight, 0.5 = noon, 0.75 = 6 PM
//!
//! ## Time Formats Accepted
//!
//! - "HH:MM:SS" - Full time with seconds (e.g., "14:30:45")
//! - "HH:MM" - Hour and minute (e.g., "14:30")
//! - "HH:MM AM/PM" - 12-hour format (e.g., "2:30 PM")
//! - Numeric - Fractional day value (e.g., 0.5 for noon)
//! - TimeSerial(hour, minute, second) - Function result
//! - TimeValue(string) - Converted string
//!
//! ## Common Errors
//!
//! - **Error 5**: Invalid procedure call or argument - occurs with invalid time format
//! - **Error 70**: Permission denied - occurs without sufficient privileges
//! - **Error 13**: Type mismatch - occurs with incompatible data types
//!
//! ## Best Practices
//!
//! - Always use error handling when setting system time
//! - Validate time strings before assignment using `IsDate()`
//! - Use `TimeSerial` or `TimeValue` for programmatic time construction
//! - Consider using application-level time variables instead of changing system time
//! - Document why system time is being changed in production code
//! - Test time-setting code with various formats and edge cases
//! - Be aware of time zone and daylight saving time implications
//! - Consider user permissions and UAC on modern Windows systems
//!
//! ## See Also
//!
//! - `Time` function (retrieve current system time)
//! - `Date` statement (set system date)
//! - `Now` function (get current date and time)
//! - `TimeSerial` function (create time from components)
//! - `TimeValue` function (convert string to time)
//! - `Hour`, `Minute`, `Second` functions (extract time components)
//!
//! ## References
//!
//! - [Time Statement - Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/time-statement)

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses a Time statement.
    pub(crate) fn parse_time_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::TimeStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn time_simple() {
        let source = r#"
Sub Test()
    Time = "14:30:00"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
        assert!(debug.contains("TimeKeyword"));
    }

    #[test]
    fn time_at_module_level() {
        let source = "Time = \"12:00:00\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_with_am_pm() {
        let source = r#"
Sub Test()
    Time = "9:15 AM"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_midnight() {
        let source = r#"
Sub Test()
    Time = "00:00:00"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_noon() {
        let source = r#"
Sub Test()
    Time = "12:00:00"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_with_variable() {
        let source = r"
Sub Test()
    Time = newTime
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
        assert!(debug.contains("newTime"));
    }

    #[test]
    fn time_with_timevalue() {
        let source = r#"
Sub Test()
    Time = TimeValue("3:30 PM")
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
        assert!(debug.contains("TimeValue"));
    }

    #[test]
    fn time_with_offset() {
        let source = r#"
Sub Test()
    Time = Time + TimeValue("00:15:00")
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_with_comment() {
        let source = r#"
Sub Test()
    Time = "10:30:00" ' Set to 10:30 AM
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
        assert!(debug.contains("' Set to 10:30 AM"));
    }

    #[test]
    fn time_preserves_whitespace() {
        let source = "Time   =   \"12:00:00\"  \n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_with_now() {
        let source = r"
Sub Test()
    Time = Now
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
        assert!(debug.contains("Now"));
    }

    #[test]
    fn time_in_if_statement() {
        let source = r#"
If condition Then
    Time = "08:00:00"
End If
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_with_timeserial() {
        let source = r"
Sub Test()
    Time = TimeSerial(14, 30, 0)
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
        assert!(debug.contains("TimeSerial"));
    }

    #[test]
    fn time_with_concatenation() {
        let source = r#"
Sub Test()
    Time = hours & ":" & minutes & ":00"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_in_function() {
        let source = r#"
Function SetSystemTime() As Boolean
    Time = "12:00:00"
    SetSystemTime = True
End Function
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_in_error_handler() {
        let source = r#"
On Error Resume Next
Time = "10:30:00"
If Err.Number <> 0 Then
    MsgBox "Error"
End If
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_in_class_module() {
        let source = r#"
Private Sub Class_Initialize()
    Time = "12:00:00"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.cls", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_with_format() {
        let source = r#"
Sub Test()
    Time = Format(Now, "hh:mm:ss")
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
        assert!(debug.contains("Format"));
    }

    #[test]
    fn time_in_loop() {
        let source = r"
For i = 0 To 23
    Time = TimeSerial(i, 0, 0)
Next i
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_with_isdate() {
        let source = r"
Sub Test()
    If IsDate(userTime) Then
        Time = userTime
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
        assert!(debug.contains("IsDate"));
    }

    #[test]
    fn time_in_select_case() {
        let source = r#"
Select Case timeType
    Case 1
        Time = "08:00:00"
    Case 2
        Time = "12:00:00"
End Select
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_multiple_on_same_line() {
        let source = "Time = \"08:00:00\": Time = \"12:00:00\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_in_with_block() {
        let source = r"
With timeSettings
    Time = .DefaultTime
End With
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_with_hour_function() {
        let source = r#"
Sub Test()
    If Hour(Time) > 17 Then
        Time = "08:00:00"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_numeric_value() {
        let source = r"
Sub Test()
    Time = 0.5
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_with_line_continuation() {
        let source = r#"
Sub Test()
    Time = _
        "14:30:00"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_inline_if() {
        let source = "If resetTime Then Time = \"00:00:00\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_testing_midnight() {
        let source = r#"
Sub Test()
    Time = "23:59:59"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_with_expression() {
        let source = r"
Sub Test()
    Time = baseTime + offset
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }

    #[test]
    fn time_case_insensitive() {
        let source = "time = \"12:00:00\"\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("TimeStatement"));
    }
}
