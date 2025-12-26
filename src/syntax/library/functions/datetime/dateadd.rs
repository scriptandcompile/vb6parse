//! # `DateAdd` Function
//!
//! Returns a `Variant` (`Date`) containing a date to which a specified time interval has been added.
//!
//! ## Syntax
//!
//! ```vb
//! DateAdd(interval, number, date)
//! ```
//!
//! ## Parameters
//!
//! - **`interval`**: Required. `String` expression that is the interval of time you want to add.
//!   See the Interval Settings section for valid values.
//! - **`number`**: Required. `Numeric` expression that is the number of intervals you want to add.
//!   Can be positive (to get dates in the future) or negative (to get dates in the past).
//! - **`date`**: Required. `Variant` (`Date`) or literal representing the date to which the interval is added.
//!
//! ## Interval Settings
//!
//! The `interval` parameter can have the following values:
//!
//! | Setting | Description |
//! |---------|-------------|
//! | "yyyy" | Year |
//! | "q" | Quarter |
//! | "m" | Month |
//! | "y" | Day of year |
//! | "d" | Day |
//! | "w" | Weekday |
//! | "ww" | Week of year |
//! | "h" | Hour |
//! | "n" | Minute |
//! | "s" | Second |
//!
//! ## Return Value
//!
//! Returns a `Variant` of subtype `Date` containing the result of adding the specified interval
//! to the given date. Returns Null if any parameter is Null.
//!
//! ## Remarks
//!
//! The `DateAdd` function is used to add or subtract a specified time interval from a date.
//! You can use it to calculate future or past dates relative to a known date.
//!
//! **Important Characteristics:**
//!
//! - Negative numbers subtract intervals (dates in the past)
//! - Positive numbers add intervals (dates in the future)
//! - Handles month-end dates intelligently (e.g., adding 1 month to Jan 31 gives Feb 28/29)
//! - When adding months, if the resulting day doesn't exist, uses last day of month
//! - Respects daylight saving time transitions
//! - Week ("ww") interval treats Sunday as the first day of the week
//! - Weekday ("w") interval is equivalent to day ("d") interval
//! - Day of year ("y") interval is equivalent to day ("d") interval
//!
//! ## Month and Year Calculations
//!
//! When adding months or years, `DateAdd` ensures the result is valid:
//! - Jan 31 + 1 month = Feb 28 (or 29 in leap year)
//! - Jan 31 + 2 months = Mar 31
//! - Aug 31 - 3 months = May 31
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Add days to a date
//! Dim futureDate As Date
//! futureDate = DateAdd("d", 30, Date)
//! MsgBox "30 days from now: " & futureDate
//!
//! ' Subtract days from a date
//! Dim pastDate As Date
//! pastDate = DateAdd("d", -7, Date)
//! MsgBox "A week ago: " & pastDate
//!
//! ' Add months
//! Dim nextMonth As Date
//! nextMonth = DateAdd("m", 1, Date)
//! MsgBox "One month from now: " & nextMonth
//! ```
//!
//! ### Different Time Intervals
//!
//! ```vb
//! Dim startDate As Date
//! startDate = #1/15/2025#
//!
//! ' Add years
//! MsgBox "Next year: " & DateAdd("yyyy", 1, startDate)
//!
//! ' Add quarters
//! MsgBox "Next quarter: " & DateAdd("q", 1, startDate)
//!
//! ' Add weeks
//! MsgBox "Next week: " & DateAdd("ww", 1, startDate)
//!
//! ' Add hours
//! MsgBox "In 6 hours: " & DateAdd("h", 6, startDate)
//!
//! ' Add minutes
//! MsgBox "In 90 minutes: " & DateAdd("n", 90, startDate)
//!
//! ' Add seconds
//! MsgBox "In 3600 seconds: " & DateAdd("s", 3600, startDate)
//! ```
//!
//! ### Working with Past Dates
//!
//! ```vb
//! ' Calculate date 90 days ago
//! Dim quarterAgo As Date
//! quarterAgo = DateAdd("d", -90, Date)
//!
//! ' Calculate date 1 year ago
//! Dim yearAgo As Date
//! yearAgo = DateAdd("yyyy", -1, Date)
//!
//! ' Calculate date 3 months ago
//! Dim threeMonthsAgo As Date
//! threeMonthsAgo = DateAdd("m", -3, Date)
//! ```
//!
//! ## Common Patterns
//!
//! ### Due Date Calculation
//!
//! ```vb
//! Function CalculateDueDate(invoiceDate As Date, terms As Integer) As Date
//!     ' NET 30, NET 60, etc.
//!     CalculateDueDate = DateAdd("d", terms, invoiceDate)
//! End Function
//!
//! ' Usage
//! Dim invoice As Date
//! Dim dueDate As Date
//! invoice = Date
//! dueDate = CalculateDueDate(invoice, 30)  ' Due in 30 days
//! ```
//!
//! ### Age-Based Eligibility
//!
//! ```vb
//! Function IsOldEnough(birthDate As Date, requiredAge As Integer) As Boolean
//!     Dim eligibilityDate As Date
//!     eligibilityDate = DateAdd("yyyy", requiredAge, birthDate)
//!     IsOldEnough = (Date >= eligibilityDate)
//! End Function
//!
//! ' Usage
//! If IsOldEnough(#5/10/2005#, 18) Then
//!     MsgBox "Eligible"
//! End If
//! ```
//!
//! ### Expiration Date Setting
//!
//! ```vb
//! Function SetExpirationDate(startDate As Date, months As Integer) As Date
//!     SetExpirationDate = DateAdd("m", months, startDate)
//! End Function
//!
//! ' Set license to expire in 12 months
//! Dim license As Date
//! license = Date
//! Dim expires As Date
//! expires = SetExpirationDate(license, 12)
//! ```
//!
//! ### Meeting Schedule
//!
//! ```vb
//! Function GetNextMeeting(lastMeeting As Date, interval As String, count As Integer) As Date
//!     GetNextMeeting = DateAdd(interval, count, lastMeeting)
//! End Function
//!
//! ' Weekly meeting
//! Dim nextWeekly As Date
//! nextWeekly = GetNextMeeting(#1/15/2025#, "ww", 1)
//!
//! ' Monthly meeting
//! Dim nextMonthly As Date
//! nextMonthly = GetNextMeeting(#1/15/2025#, "m", 1)
//! ```
//!
//! ### Subscription Renewal
//!
//! ```vb
//! Sub CalculateRenewalDates()
//!     Dim startDate As Date
//!     Dim firstRenewal As Date
//!     Dim secondRenewal As Date
//!     
//!     startDate = Date
//!     firstRenewal = DateAdd("m", 12, startDate)   ' Annual renewal
//!     secondRenewal = DateAdd("m", 24, startDate)  ' Second year
//!     
//!     MsgBox "Start: " & startDate & vbCrLf & _
//!            "First renewal: " & firstRenewal & vbCrLf & _
//!            "Second renewal: " & secondRenewal
//! End Sub
//! ```
//!
//! ### Trial Period End
//!
//! ```vb
//! Function GetTrialEndDate(startDate As Date, trialDays As Integer) As Date
//!     GetTrialEndDate = DateAdd("d", trialDays, startDate)
//! End Function
//!
//! ' 30-day trial
//! Dim trialStart As Date
//! Dim trialEnd As Date
//! trialStart = Date
//! trialEnd = GetTrialEndDate(trialStart, 30)
//! ```
//!
//! ### Report Period Calculation
//!
//! ```vb
//! Function GetReportingPeriod(endDate As Date, months As Integer) As Date
//!     ' Calculate start date by going back specified months
//!     GetReportingPeriod = DateAdd("m", -months, endDate)
//! End Function
//!
//! ' Get start of 6-month period ending today
//! Dim periodStart As Date
//! periodStart = GetReportingPeriod(Date, 6)
//! ```
//!
//! ### Reminder Dates
//!
//! ```vb
//! Sub SetReminders(eventDate As Date)
//!     Dim oneWeekBefore As Date
//!     Dim oneDayBefore As Date
//!     Dim oneHourBefore As Date
//!     
//!     oneWeekBefore = DateAdd("d", -7, eventDate)
//!     oneDayBefore = DateAdd("d", -1, eventDate)
//!     oneHourBefore = DateAdd("h", -1, eventDate)
//!     
//!     ' Schedule reminders...
//! End Sub
//! ```
//!
//! ## Advanced Usage
//!
//! ### Business Days Calculation
//!
//! ```vb
//! Function AddBusinessDays(startDate As Date, days As Integer) As Date
//!     Dim result As Date
//!     Dim daysAdded As Integer
//!     Dim direction As Integer
//!     
//!     result = startDate
//!     direction = Sgn(days)
//!     daysAdded = 0
//!     
//!     Do While Abs(daysAdded) < Abs(days)
//!         result = DateAdd("d", direction, result)
//!         
//!         ' Skip weekends
//!         If Weekday(result) <> vbSaturday And Weekday(result) <> vbSunday Then
//!             daysAdded = daysAdded + direction
//!         End If
//!     Loop
//!     
//!     AddBusinessDays = result
//! End Function
//! ```
//!
//! ### Date Range Generator
//!
//! ```vb
//! Function GenerateDateSeries(startDate As Date, interval As String, _
//!                            count As Integer, step As Integer) As Variant
//!     Dim dates() As Date
//!     Dim i As Integer
//!     
//!     ReDim dates(0 To count - 1)
//!     
//!     For i = 0 To count - 1
//!         dates(i) = DateAdd(interval, i * step, startDate)
//!     Next i
//!     
//!     GenerateDateSeries = dates
//! End Function
//!
//! ' Generate 12 month-end dates
//! Dim monthEnds As Variant
//! monthEnds = GenerateDateSeries(#1/31/2025#, "m", 12, 1)
//! ```
//!
//! ### Fiscal Period Calculator
//!
//! ```vb
//! Function GetFiscalQuarterEnd(fiscalYearStart As Date, quarter As Integer) As Date
//!     Dim quarterStart As Date
//!     Dim quarterEnd As Date
//!     
//!     ' Calculate start of quarter
//!     quarterStart = DateAdd("m", (quarter - 1) * 3, fiscalYearStart)
//!     
//!     ' End is 3 months later minus 1 day
//!     quarterEnd = DateAdd("d", -1, DateAdd("m", 3, quarterStart))
//!     
//!     GetFiscalQuarterEnd = quarterEnd
//! End Function
//! ```
//!
//! ### Recurring Event Calculator
//!
//! ```vb
//! Function GetNextOccurrence(lastOccurrence As Date, frequency As String) As Date
//!     Select Case LCase(frequency)
//!         Case "daily"
//!             GetNextOccurrence = DateAdd("d", 1, lastOccurrence)
//!         Case "weekly"
//!             GetNextOccurrence = DateAdd("ww", 1, lastOccurrence)
//!         Case "biweekly"
//!             GetNextOccurrence = DateAdd("ww", 2, lastOccurrence)
//!         Case "monthly"
//!             GetNextOccurrence = DateAdd("m", 1, lastOccurrence)
//!         Case "quarterly"
//!             GetNextOccurrence = DateAdd("q", 1, lastOccurrence)
//!         Case "annually"
//!             GetNextOccurrence = DateAdd("yyyy", 1, lastOccurrence)
//!         Case Else
//!             GetNextOccurrence = lastOccurrence
//!     End Select
//! End Function
//! ```
//!
//! ### Time Zone Offset (Simple)
//!
//! ```vb
//! Function ConvertToTimeZone(localTime As Date, hourOffset As Integer) As Date
//!     ' Simple timezone conversion (doesn't account for DST)
//!     ConvertToTimeZone = DateAdd("h", hourOffset, localTime)
//! End Function
//!
//! ' Convert EST to PST (3 hours earlier)
//! Dim estTime As Date
//! Dim pstTime As Date
//! estTime = Now
//! pstTime = ConvertToTimeZone(estTime, -3)
//! ```
//!
//! ### Age Calculator with Precision
//!
//! ```vb
//! Function GetExactAge(birthDate As Date) As String
//!     Dim years As Integer
//!     Dim months As Integer
//!     Dim days As Integer
//!     Dim tempDate As Date
//!     
//!     ' Calculate years
//!     tempDate = birthDate
//!     years = 0
//!     Do While DateAdd("yyyy", years + 1, tempDate) <= Date
//!         years = years + 1
//!     Loop
//!     
//!     ' Calculate remaining months
//!     tempDate = DateAdd("yyyy", years, birthDate)
//!     months = 0
//!     Do While DateAdd("m", months + 1, tempDate) <= Date
//!         months = months + 1
//!     Loop
//!     
//!     ' Calculate remaining days
//!     tempDate = DateAdd("m", months, DateAdd("yyyy", years, birthDate))
//!     days = DateDiff("d", tempDate, Date)
//!     
//!     GetExactAge = years & " years, " & months & " months, " & days & " days"
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeDateAdd(interval As String, number As Long, _
//!                     dateValue As Date) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     ' Validate interval
//!     Select Case LCase(interval)
//!         Case "yyyy", "q", "m", "y", "d", "w", "ww", "h", "n", "s"
//!             SafeDateAdd = DateAdd(interval, number, dateValue)
//!         Case Else
//!             SafeDateAdd = Null  ' Invalid interval
//!     End Select
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeDateAdd = Null  ' Return Null on error
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 5** (Invalid procedure call): Invalid interval string
//! - **Error 13** (Type mismatch): Non-numeric number or non-date date parameter
//! - **Error 6** (Overflow): Result date is outside valid range (100-9999 AD)
//!
//! ## Performance Considerations
//!
//! - `DateAdd` is efficient for single date calculations
//! - For large date ranges, consider pre-calculating frequently used values
//! - Month and year additions are slightly slower than day additions
//! - Interval string comparison is case-insensitive but exact strings are faster
//!
//! ## Best Practices
//!
//! ### Use Named Constants for Intervals
//!
//! ```vb
//! ' Define constants for clarity
//! Const INTERVAL_YEAR As String = "yyyy"
//! Const INTERVAL_MONTH As String = "m"
//! Const INTERVAL_DAY As String = "d"
//! Const INTERVAL_HOUR As String = "h"
//!
//! ' Use in code
//! nextYear = DateAdd(INTERVAL_YEAR, 1, Date)
//! ```
//!
//! ### Validate Input Dates
//!
//! ```vb
//! Function AddDaysToDate(startDate As Variant, days As Integer) As Date
//!     If Not IsDate(startDate) Then
//!         Err.Raise 13, , "Invalid date"
//!     End If
//!     
//!     AddDaysToDate = DateAdd("d", days, CDate(startDate))
//! End Function
//! ```
//!
//! ### Handle Month-End Edge Cases
//!
//! ```vb
//! ' Be aware of month-end behavior
//! Dim jan31 As Date
//! jan31 = #1/31/2025#
//!
//! ' Adding 1 month gives Feb 28 (or 29)
//! Dim result As Date
//! result = DateAdd("m", 1, jan31)  ' Feb 28, 2025
//!
//! ' Adding 2 months gives Mar 31
//! result = DateAdd("m", 2, jan31)  ' Mar 31, 2025
//! ```
//!
//! ## Comparison with Other Date Functions
//!
//! ### `DateAdd` vs `DateDiff`
//!
//! ```vb
//! ' DateAdd - Adds interval to date, returns new date
//! Dim future As Date
//! future = DateAdd("d", 30, Date)
//!
//! ' DateDiff - Calculates interval between dates, returns number
//! Dim difference As Long
//! difference = DateDiff("d", Date, future)  ' Returns 30
//! ```
//!
//! ### `DateAdd` vs Simple Arithmetic
//!
//! ```vb
//! ' Simple arithmetic works for days
//! Dim tomorrow As Date
//! tomorrow = Date + 1  ' Same as DateAdd("d", 1, Date)
//!
//! ' But DateAdd is better for months/years
//! Dim nextMonth As Date
//! nextMonth = DateAdd("m", 1, Date)  ' Handles month-end correctly
//! ```
//!
//! ## Limitations
//!
//! - Date range limited to January 1, 100 through December 31, 9999
//! - No built-in support for business day calculations
//! - Doesn't handle holidays automatically
//! - Week starts on Sunday (cannot be customized)
//! - No built-in timezone support
//! - Daylight saving time handled by system, results may vary
//!
//! ## Related Functions
//!
//! - `DateDiff`: Returns the number of intervals between two dates
//! - `DatePart`: Returns a specified part of a date
//! - `DateSerial`: Creates a date from year, month, and day values
//! - `DateValue`: Converts a string to a date
//! - `Year`, `Month`, `Day`: Extract date components
//! - `Hour`, `Minute`, `Second`: Extract time components
//! - `Now`: Returns current date and time
//! - `Date`: Returns current date
//! - `Time`: Returns current time

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn dateadd_basic() {
        let source = r#"
result = DateAdd("d", 30, Date)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_years() {
        let source = r#"
nextYear = DateAdd("yyyy", 1, Date)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_months() {
        let source = r#"
nextMonth = DateAdd("m", 1, startDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_negative() {
        let source = r#"
pastDate = DateAdd("d", -7, Date)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_hours() {
        let source = r#"
later = DateAdd("h", 6, Now)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_minutes() {
        let source = r#"
later = DateAdd("n", 90, Now)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_seconds() {
        let source = r#"
later = DateAdd("s", 3600, Now)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_weeks() {
        let source = r#"
nextWeek = DateAdd("ww", 1, Date)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_quarters() {
        let source = r#"
nextQuarter = DateAdd("q", 1, Date)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_in_function() {
        let source = r#"
Function GetDueDate(invoice As Date) As Date
    GetDueDate = DateAdd("d", 30, invoice)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_with_literal() {
        let source = r#"
result = DateAdd("m", 6, #1/1/2025#)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_nested() {
        let source = r#"
result = DateAdd("d", -1, DateAdd("m", 3, startDate))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_in_if() {
        let source = r#"
If DateAdd("d", 30, startDate) > endDate Then
    MsgBox "Too late"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_with_variable_interval() {
        let source = r#"
Dim interval As String
interval = "m"
result = DateAdd(interval, 1, Date)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_in_loop() {
        let source = r#"
For i = 1 To 12
    dates(i) = DateAdd("m", i, startDate)
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_in_select_case() {
        let source = r#"
Select Case frequency
    Case "monthly"
        nextDate = DateAdd("m", 1, lastDate)
    Case "yearly"
        nextDate = DateAdd("yyyy", 1, lastDate)
End Select
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_with_format() {
        let source = r#"
formatted = Format(DateAdd("d", 7, Date), "yyyy-mm-dd")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_day_of_year() {
        let source = r#"
result = DateAdd("y", 1, Date)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_weekday() {
        let source = r#"
result = DateAdd("w", 1, Date)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_in_array() {
        let source = r#"
dates(0) = DateAdd("m", 0, startDate)
dates(1) = DateAdd("m", 1, startDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_multiple_calls() {
        let source = r#"
start = DateAdd("m", -1, Date)
finish = DateAdd("m", 1, Date)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_in_msgbox() {
        let source = r#"
MsgBox "Next week: " & DateAdd("ww", 1, Date)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_with_expression() {
        let source = r#"
result = DateAdd("d", days * 2, startDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_zero_interval() {
        let source = r#"
result = DateAdd("d", 0, Date)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn dateadd_large_number() {
        let source = r#"
result = DateAdd("d", 365, Date)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateAdd"));
        assert!(debug.contains("Identifier"));
    }
}
