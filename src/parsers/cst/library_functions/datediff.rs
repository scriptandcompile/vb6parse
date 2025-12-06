//! # `DateDiff` Function
//!
//! Returns a `Variant` (`Long`) specifying the number of time intervals between two specified dates.
//!
//! ## Syntax
//!
//! ```vb
//! DateDiff(interval, date1, date2[, firstdayofweek[, firstweekofyear]])
//! ```
//!
//! ## Parameters
//!
//! - **interval**: Required. `String` expression that is the interval of time you want to use
//!   to calculate the difference between date1 and date2. See Interval Settings for values.
//! - **date1**, **date2**: Required. `Variant` (`Date`) values that you want to use in the calculation.
//! - **firstdayofweek**: Optional. Constant that specifies the first day of the week.
//!   If not specified, Sunday is assumed. See `FirstDayOfWeek` Constants.
//! - **firstweekofyear**: Optional. Constant that specifies the first week of the year.
//!   If not specified, the first week is assumed to be the week containing January 1.
//!   See `FirstWeekOfYear` Constants.
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
//! ## `FirstDayOfWeek` Constants
//!
//! | Constant | Value | Description |
//! |----------|-------|-------------|
//! | vbUseSystem | 0 | Use system setting |
//! | vbSunday | 1 | Sunday (default) |
//! | vbMonday | 2 | Monday |
//! | vbTuesday | 3 | Tuesday |
//! | vbWednesday | 4 | Wednesday |
//! | vbThursday | 5 | Thursday |
//! | vbFriday | 6 | Friday |
//! | vbSaturday | 7 | Saturday |
//!
//! ## `FirstWeekOfYear` Constants
//!
//! | Constant | Value | Description |
//! |----------|-------|-------------|
//! | vbUseSystem | 0 | Use system setting |
//! | vbFirstJan1 | 1 | Start with week containing January 1 (default) |
//! | vbFirstFourDays | 2 | Start with week having at least 4 days in new year |
//! | vbFirstFullWeek | 3 | Start with first full week of the year |
//!
//! ## Return Value
//!
//! Returns a `Long` integer representing the number of intervals between the two dates.
//! The result is positive if date2 is later than date1, negative if date2 is earlier than date1,
//! and zero if they are equal.
//!
//! ## Remarks
//!
//! The `DateDiff` function is used to calculate the difference between two dates in the
//! specified time interval. The function counts the number of interval boundaries crossed
//! between the two dates.
//!
//! **Important Characteristics:**
//!
//! - Returns positive number if date2 > date1 (future date)
//! - Returns negative number if date2 < date1 (past date)
//! - Returns zero if date2 = date1 (same date/time)
//! - Counts interval boundaries, not elapsed time
//! - For "yyyy", crossing from Dec 31 to Jan 1 counts as 1 year
//! - For "m", crossing from Jan 31 to Feb 1 counts as 1 month
//! - For "ww", counts week boundaries (Sunday to Sunday by default)
//! - Day of year ("y") is equivalent to day ("d")
//! - Weekday ("w") is equivalent to day ("d")
//!
//! ## Boundary Counting vs Elapsed Time
//!
//! `DateDiff` counts boundaries crossed, not elapsed time:
//!
//! ```vb
//! ' Year example
//! DateDiff("yyyy", #12/31/2024#, #1/1/2025#)  ' Returns 1 (crossed 1 year boundary)
//! ' But only 1 day elapsed!
//!
//! ' Month example
//! DateDiff("m", #1/31/2025#, #2/1/2025#)  ' Returns 1 (crossed 1 month boundary)
//! ' But only 1 day elapsed!
//!
//! ' Day example (actual elapsed time)
//! DateDiff("d", #1/1/2025#, #1/31/2025#)  ' Returns 30 (30 days elapsed)
//! ```
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Calculate days between dates
//! Dim days As Long
//! days = DateDiff("d", #1/1/2025#, #1/31/2025#)
//! MsgBox "Days: " & days  ' Shows 30
//!
//! ' Calculate months between dates
//! Dim months As Long
//! months = DateDiff("m", #1/15/2025#, #6/15/2025#)
//! MsgBox "Months: " & months  ' Shows 5
//!
//! ' Calculate years between dates
//! Dim years As Long
//! years = DateDiff("yyyy", #1/1/2000#, #1/1/2025#)
//! MsgBox "Years: " & years  ' Shows 25
//! ```
//!
//! ### Age Calculation
//!
//! ```vb
//! Function CalculateAge(birthDate As Date) As Integer
//!     Dim age As Integer
//!     age = DateDiff("yyyy", birthDate, Date)
//!     
//!     ' Adjust if birthday hasn't occurred this year
//!     If DateSerial(Year(Date), Month(birthDate), Day(birthDate)) > Date Then
//!         age = age - 1
//!     End If
//!     
//!     CalculateAge = age
//! End Function
//! ```
//!
//! ### Days Until/Since Event
//!
//! ```vb
//! Function DaysUntilEvent(eventDate As Date) As Long
//!     DaysUntilEvent = DateDiff("d", Date, eventDate)
//! End Function
//!
//! ' Usage
//! Dim daysLeft As Long
//! daysLeft = DaysUntilEvent(#12/25/2025#)
//! If daysLeft > 0 Then
//!     MsgBox daysLeft & " days until Christmas"
//! ElseIf daysLeft < 0 Then
//!     MsgBox "Christmas was " & Abs(daysLeft) & " days ago"
//! Else
//!     MsgBox "Today is Christmas!"
//! End If
//! ```
//!
//! ## Common Patterns
//!
//! ### Elapsed Time Display
//!
//! ```vb
//! Function FormatElapsedTime(startTime As Date, endTime As Date) As String
//!     Dim hours As Long
//!     Dim minutes As Long
//!     Dim seconds As Long
//!     
//!     hours = DateDiff("h", startTime, endTime)
//!     minutes = DateDiff("n", startTime, endTime) Mod 60
//!     seconds = DateDiff("s", startTime, endTime) Mod 60
//!     
//!     FormatElapsedTime = hours & ":" & Format(minutes, "00") & ":" & Format(seconds, "00")
//! End Function
//! ```
//!
//! ### Working Days Calculator
//!
//! ```vb
//! Function CountWorkingDays(startDate As Date, endDate As Date) As Long
//!     Dim dayCount As Long
//!     Dim workDays As Long
//!     Dim currentDate As Date
//!     
//!     dayCount = DateDiff("d", startDate, endDate)
//!     workDays = 0
//!     
//!     For i = 0 To dayCount
//!         currentDate = DateAdd("d", i, startDate)
//!         If Weekday(currentDate) <> vbSaturday And Weekday(currentDate) <> vbSunday Then
//!             workDays = workDays + 1
//!         End If
//!     Next i
//!     
//!     CountWorkingDays = workDays
//! End Function
//! ```
//!
//! ### Overdue Indicator
//!
//! ```vb
//! Function GetOverdueDays(dueDate As Date) As Long
//!     Dim days As Long
//!     days = DateDiff("d", dueDate, Date)
//!     
//!     If days > 0 Then
//!         GetOverdueDays = days  ' Positive = overdue
//!     Else
//!         GetOverdueDays = 0     ' Not overdue
//!     End If
//! End Function
//! ```
//!
//! ### Subscription Status
//!
//! ```vb
//! Function GetSubscriptionStatus(startDate As Date, endDate As Date) As String
//!     Dim daysRemaining As Long
//!     
//!     daysRemaining = DateDiff("d", Date, endDate)
//!     
//!     Select Case daysRemaining
//!         Case Is < 0
//!             GetSubscriptionStatus = "Expired"
//!         Case 0 To 7
//!             GetSubscriptionStatus = "Expiring Soon (" & daysRemaining & " days)"
//!         Case 8 To 30
//!             GetSubscriptionStatus = "Active (" & daysRemaining & " days left)"
//!         Case Else
//!             GetSubscriptionStatus = "Active"
//!     End Select
//! End Function
//! ```
//!
//! ### Quarterly Report Period
//!
//! ```vb
//! Function GetQuartersBetween(startDate As Date, endDate As Date) As Integer
//!     GetQuartersBetween = DateDiff("q", startDate, endDate)
//! End Function
//!
//! ' Check if in same quarter
//! Function InSameQuarter(date1 As Date, date2 As Date) As Boolean
//!     InSameQuarter = (DateDiff("q", date1, date2) = 0)
//! End Function
//! ```
//!
//! ### Meeting Interval Tracker
//!
//! ```vb
//! Function WeeksSinceLastMeeting(lastMeeting As Date) As Long
//!     WeeksSinceLastMeeting = DateDiff("ww", lastMeeting, Date)
//! End Function
//!
//! Function IsMeetingDue(lastMeeting As Date, interval As Integer) As Boolean
//!     IsMeetingDue = (DateDiff("ww", lastMeeting, Date) >= interval)
//! End Function
//! ```
//!
//! ### Time Tracking
//!
//! ```vb
//! Sub LogSessionDuration(startTime As Date, endTime As Date)
//!     Dim hours As Long
//!     Dim minutes As Long
//!     
//!     hours = DateDiff("h", startTime, endTime)
//!     minutes = DateDiff("n", startTime, endTime) - (hours * 60)
//!     
//!     Debug.Print "Session duration: " & hours & "h " & minutes & "m"
//! End Sub
//! ```
//!
//! ### Age Range Categorization
//!
//! ```vb
//! Function GetAgeCategory(birthDate As Date) As String
//!     Dim age As Integer
//!     age = DateDiff("yyyy", birthDate, Date)
//!     
//!     ' Adjust for birthday not yet occurred
//!     If Month(Date) < Month(birthDate) Or _
//!        (Month(Date) = Month(birthDate) And Day(Date) < Day(birthDate)) Then
//!         age = age - 1
//!     End If
//!     
//!     Select Case age
//!         Case 0 To 12
//!             GetAgeCategory = "Child"
//!         Case 13 To 19
//!             GetAgeCategory = "Teenager"
//!         Case 20 To 64
//!             GetAgeCategory = "Adult"
//!         Case Else
//!             GetAgeCategory = "Senior"
//!     End Select
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Complete Time Breakdown
//!
//! ```vb
//! Type TimeBreakdown
//!     Years As Long
//!     Months As Long
//!     Days As Long
//!     Hours As Long
//!     Minutes As Long
//!     Seconds As Long
//! End Type
//!
//! Function GetDetailedDifference(startDate As Date, endDate As Date) As TimeBreakdown
//!     Dim result As TimeBreakdown
//!     Dim tempDate As Date
//!     
//!     ' Calculate years
//!     result.Years = DateDiff("yyyy", startDate, endDate)
//!     tempDate = DateAdd("yyyy", result.Years, startDate)
//!     If tempDate > endDate Then
//!         result.Years = result.Years - 1
//!         tempDate = DateAdd("yyyy", result.Years, startDate)
//!     End If
//!     
//!     ' Calculate months
//!     result.Months = DateDiff("m", tempDate, endDate)
//!     tempDate = DateAdd("m", result.Months, tempDate)
//!     If tempDate > endDate Then
//!         result.Months = result.Months - 1
//!         tempDate = DateAdd("m", result.Months, DateAdd("yyyy", result.Years, startDate))
//!     End If
//!     
//!     ' Calculate remaining time
//!     result.Days = DateDiff("d", tempDate, endDate)
//!     result.Hours = DateDiff("h", tempDate, endDate) Mod 24
//!     result.Minutes = DateDiff("n", tempDate, endDate) Mod 60
//!     result.Seconds = DateDiff("s", tempDate, endDate) Mod 60
//!     
//!     GetDetailedDifference = result
//! End Function
//! ```
//!
//! ### Week Number with Custom First Day
//!
//! ```vb
//! Function GetWeekNumber(dateValue As Date, startDay As VbDayOfWeek) As Long
//!     Dim yearStart As Date
//!     yearStart = DateSerial(Year(dateValue), 1, 1)
//!     GetWeekNumber = DateDiff("ww", yearStart, dateValue, startDay, vbFirstFourDays)
//! End Function
//!
//! ' Usage
//! Dim weekNum As Long
//! weekNum = GetWeekNumber(Date, vbMonday)  ' ISO week number (Monday start)
//! ```
//!
//! ### Performance Timer
//!
//! ```vb
//! Private m_startTime As Date
//!
//! Sub StartTimer()
//!     m_startTime = Now
//! End Sub
//!
//! Function GetElapsedMilliseconds() As Double
//!     Dim seconds As Long
//!     seconds = DateDiff("s", m_startTime, Now)
//!     
//!     ' VB6 doesn't support milliseconds directly
//!     ' This gives seconds as closest approximation
//!     GetElapsedMilliseconds = seconds * 1000
//! End Function
//! ```
//!
//! ### Date Range Validator
//!
//! ```vb
//! Function ValidateDateRange(startDate As Date, endDate As Date, _
//!                          maxDays As Long) As Boolean
//!     Dim daysDiff As Long
//!     
//!     ' Check date order
//!     If startDate > endDate Then
//!         ValidateDateRange = False
//!         Exit Function
//!     End If
//!     
//!     ' Check range limit
//!     daysDiff = DateDiff("d", startDate, endDate)
//!     ValidateDateRange = (daysDiff <= maxDays)
//! End Function
//! ```
//!
//! ### Fiscal Period Calculator
//!
//! ```vb
//! Function GetFiscalPeriodDifference(date1 As Date, date2 As Date, _
//!                                   fiscalYearStart As Integer) As Long
//!     ' Calculate fiscal months between dates
//!     ' fiscalYearStart = month number (e.g., 4 for April)
//!     
//!     Dim adjustedDate1 As Date
//!     Dim adjustedDate2 As Date
//!     
//!     ' Adjust dates to fiscal year basis
//!     adjustedDate1 = DateSerial(Year(date1), Month(date1) - fiscalYearStart + 1, Day(date1))
//!     adjustedDate2 = DateSerial(Year(date2), Month(date2) - fiscalYearStart + 1, Day(date2))
//!     
//!     GetFiscalPeriodDifference = DateDiff("m", adjustedDate1, adjustedDate2)
//! End Function
//! ```
//!
//! ### Batch Date Comparison
//!
//! ```vb
//! Function FindOldestDate(dates() As Date) As Date
//!     Dim i As Integer
//!     Dim oldest As Date
//!     
//!     oldest = dates(LBound(dates))
//!     
//!     For i = LBound(dates) + 1 To UBound(dates)
//!         If DateDiff("d", dates(i), oldest) > 0 Then
//!             oldest = dates(i)
//!         End If
//!     Next i
//!     
//!     FindOldestDate = oldest
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeDateDiff(interval As String, date1 As Variant, _
//!                      date2 As Variant) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     ' Validate dates
//!     If Not IsDate(date1) Or Not IsDate(date2) Then
//!         SafeDateDiff = Null
//!         Exit Function
//!     End If
//!     
//!     ' Validate interval
//!     Select Case LCase(interval)
//!         Case "yyyy", "q", "m", "y", "d", "w", "ww", "h", "n", "s"
//!             SafeDateDiff = DateDiff(interval, CDate(date1), CDate(date2))
//!         Case Else
//!             SafeDateDiff = Null
//!     End Select
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeDateDiff = Null
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 5** (Invalid procedure call): Invalid interval string
//! - **Error 13** (Type mismatch): Non-date values passed as date parameters
//! - **Error 6** (Overflow): Result exceeds Long integer range
//!
//! ## Performance Considerations
//!
//! - `DateDiff` is very fast for simple interval calculations
//! - Day ("d") calculations are fastest
//! - Month ("m") and year ("yyyy") require more computation
//! - Week calculations depend on `FirstDayOfWeek` and `FirstWeekOfYear` parameters
//! - For large datasets, cache `DateDiff` results when possible
//!
//! ## Best Practices
//!
//! ### Use Appropriate Intervals
//!
//! ```vb
//! ' Good - Use "d" for exact day count
//! days = DateDiff("d", startDate, endDate)
//!
//! ' Be careful - "yyyy" counts year boundaries, not elapsed years
//! years = DateDiff("yyyy", #12/31/2024#, #1/1/2025#)  ' Returns 1, but only 1 day!
//! ```
//!
//! ### Order Matters
//!
//! ```vb
//! ' Positive result - date2 is in future
//! diff = DateDiff("d", #1/1/2025#, #1/31/2025#)  ' Returns 30
//!
//! ' Negative result - date2 is in past
//! diff = DateDiff("d", #1/31/2025#, #1/1/2025#)  ' Returns -30
//! ```
//!
//! ### Handle Negative Results
//!
//! ```vb
//! Function GetAbsoluteDaysDifference(date1 As Date, date2 As Date) As Long
//!     GetAbsoluteDaysDifference = Abs(DateDiff("d", date1, date2))
//! End Function
//! ```
//!
//! ### Validate Date Order
//!
//! ```vb
//! Function CalculateDuration(startDate As Date, endDate As Date) As Long
//!     If startDate > endDate Then
//!         Err.Raise 5, , "Start date must be before end date"
//!     End If
//!     
//!     CalculateDuration = DateDiff("d", startDate, endDate)
//! End Function
//! ```
//!
//! ## Comparison with Other Functions
//!
//! ### `DateDiff` vs `DateAdd`
//!
//! ```vb
//! ' `DateDiff` - Calculate interval between dates (returns Long)
//! diff = DateDiff("d", #1/1/2025#, #1/31/2025#)  ' Returns 30
//!
//! ' `DateAdd` - Add interval to date (returns Date)
//! newDate = DateAdd("d", 30, #1/1/2025#)  ' Returns #1/31/2025#
//! ```
//!
//! ### `DateDiff` vs Subtraction
//!
//! ```vb
//! ' Subtraction gives days as Double
//! diff = #1/31/2025# - #1/1/2025#  ' Returns 30.0
//!
//! ' DateDiff gives days as Long
//! diff = DateDiff("d", #1/1/2025#, #1/31/2025#)  ' Returns 30
//!
//! ' DateDiff supports other intervals
//! months = DateDiff("m", #1/1/2025#, #6/1/2025#)  ' Returns 5
//! ```
//!
//! ## Limitations
//!
//! - Result must fit in Long integer range (-2,147,483,648 to 2,147,483,647)
//! - Week calculations depend on system or specified first day of week
//! - Counts boundaries crossed, not elapsed time (except for "d", "h", "n", "s")
//! - No built-in support for milliseconds
//! - No built-in support for business day calculations
//! - Cannot directly exclude holidays
//!
//! ## Related Functions
//!
//! - `DateAdd`: Adds a time interval to a date
//! - `DatePart`: Returns a specified part of a date
//! - `DateSerial`: Creates a date from year, month, and day values
//! - `Year`, `Month`, `Day`: Extract date components
//! - `Hour`, `Minute`, `Second`: Extract time components
//! - `Weekday`: Returns the day of the week
//! - `Now`: Returns current date and time
//! - `Date`: Returns current date
//! - `Time`: Returns current time

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn datediff_basic() {
        let source = r#"
days = DateDiff("d", startDate, endDate)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_years() {
        let source = r#"
years = DateDiff("yyyy", birthDate, Date)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_months() {
        let source = r#"
months = DateDiff("m", #1/1/2025#, #6/1/2025#)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_quarters() {
        let source = r#"
quarters = DateDiff("q", startDate, endDate)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_weeks() {
        let source = r#"
weeks = DateDiff("ww", lastMeeting, Date)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_hours() {
        let source = r#"
hours = DateDiff("h", startTime, endTime)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_minutes() {
        let source = r#"
minutes = DateDiff("n", startTime, endTime)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_seconds() {
        let source = r#"
seconds = DateDiff("s", startTime, Now)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_in_function() {
        let source = r#"
Function GetAge(birthDate As Date) As Integer
    GetAge = DateDiff("yyyy", birthDate, Date)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_with_firstdayofweek() {
        let source = r#"
weeks = DateDiff("ww", startDate, endDate, vbMonday)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_with_all_params() {
        let source = r#"
weeks = DateDiff("ww", startDate, endDate, vbMonday, vbFirstFourDays)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_in_if() {
        let source = r#"
If DateDiff("d", dueDate, Date) > 0 Then
    MsgBox "Overdue"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_negative_result() {
        let source = r#"
diff = DateDiff("d", endDate, startDate)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_with_abs() {
        let source = r#"
diff = Abs(DateDiff("d", date1, date2))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_in_select_case() {
        let source = r#"
Select Case DateDiff("d", dueDate, Date)
    Case Is < 0
        status = "Not due"
    Case 0
        status = "Due today"
    Case Else
        status = "Overdue"
End Select
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_day_of_year() {
        let source = r#"
days = DateDiff("y", startDate, endDate)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_weekday() {
        let source = r#"
days = DateDiff("w", startDate, endDate)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_in_loop() {
        let source = r#"
For i = 0 To count
    days(i) = DateDiff("d", startDate, dates(i))
Next i
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_comparison() {
        let source = r#"
If DateDiff("m", startDate, endDate) > 12 Then
    MsgBox "More than a year"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_modulo() {
        let source = r#"
minutes = DateDiff("n", startTime, endTime) Mod 60
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_in_msgbox() {
        let source = r#"
MsgBox "Days: " & DateDiff("d", startDate, endDate)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_multiple_calls() {
        let source = r#"
hours = DateDiff("h", startTime, endTime)
minutes = DateDiff("n", startTime, endTime) - (hours * 60)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_with_now() {
        let source = r#"
elapsed = DateDiff("s", startTime, Now)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_zero_result() {
        let source = r#"
diff = DateDiff("d", Date, Date)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn datediff_in_calculation() {
        let source = r#"
total = DateDiff("d", startDate, endDate) * rate
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateDiff"));
        assert!(debug.contains("Identifier"));
    }
}
