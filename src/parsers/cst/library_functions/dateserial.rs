//! # `DateSerial` Function
//!
//! Returns a `Variant` (`Date`) for a specified year, month, and day.
//!
//! ## Syntax
//!
//! ```vb
//! DateSerial(year, month, day)
//! ```
//!
//! ## Parameters
//!
//! - **year**: Required. `Integer` expression between 100 and 9999, inclusive, or a numeric
//!   expression. Values from 0 to 29 are interpreted as 2000-2029; values from 30 to 99
//!   are interpreted as 1930-1999.
//! - **month**: Required. `Integer` expression from 1 to 12, but can be any numeric expression
//!   representing months from -32,768 to 32,767. Month values outside 1-12 adjust the year
//!   accordingly.
//! - **day**: Required. `Integer` expression from 1 to 31, but can be any numeric expression
//!   representing days from -32,768 to 32,767. Day values outside the valid range adjust
//!   the month and year accordingly.
//!
//! ## Return Value
//!
//! Returns a `Variant` of subtype `Date` representing the specified date. The time portion
//! is set to midnight (00:00:00).
//!
//! ## Remarks
//!
//! The `DateSerial` function is used to construct a date value from individual year, month,
//! and day components. It's particularly useful for date calculations and building dates
//! programmatically.
//!
//! **Important Characteristics:**
//!
//! - Accepts values outside normal ranges and adjusts automatically
//! - Month values > 12 or < 1 adjust the year
//! - Day values outside valid range adjust the month
//! - Can use 0 or negative values for relative date calculations
//! - Two-digit years: 0-29 → 2000-2029, 30-99 → 1930-1999
//! - Always returns midnight (00:00:00) for time portion
//! - Invalid combinations return compile-time or runtime errors
//!
//! ## Range Adjustment Examples
//!
//! ```vb
//! ' Month adjustment
//! DateSerial(2025, 13, 1)    ' Returns 1/1/2026 (13th month = Jan next year)
//! DateSerial(2025, 0, 1)     ' Returns 12/1/2024 (0th month = Dec previous year)
//! DateSerial(2025, -1, 1)    ' Returns 11/1/2024 (month -1 = Nov previous year)
//!
//! ' Day adjustment
//! DateSerial(2025, 1, 32)    ' Returns 2/1/2025 (32nd day = Feb 1)
//! DateSerial(2025, 1, 0)     ' Returns 12/31/2024 (0th day = last day of prev month)
//! DateSerial(2025, 1, -1)    ' Returns 12/30/2024 (day -1)
//!
//! ' Combined adjustment
//! DateSerial(2025, 13, 32)   ' Returns 2/1/2026
//! ```
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Create a specific date
//! Dim birthday As Date
//! birthday = DateSerial(1990, 5, 15)  ' May 15, 1990
//!
//! ' Create date from variables
//! Dim y As Integer, m As Integer, d As Integer
//! y = 2025
//! m = 12
//! d = 25
//! Dim christmas As Date
//! christmas = DateSerial(y, m, d)
//!
//! ' Current year's date
//! Dim thisYear As Date
//! thisYear = DateSerial(Year(Date), 1, 1)  ' January 1 of current year
//! ```
//!
//! ### Last Day of Month
//!
//! ```vb
//! Function GetLastDayOfMonth(year As Integer, month As Integer) As Date
//!     ' Use day 0 of next month to get last day of current month
//!     GetLastDayOfMonth = DateSerial(year, month + 1, 0)
//! End Function
//!
//! ' Usage
//! Dim lastDay As Date
//! lastDay = GetLastDayOfMonth(2025, 2)  ' Feb 28, 2025 (or 29 in leap year)
//! ```
//!
//! ### First Day of Month
//!
//! ```vb
//! Function GetFirstDayOfMonth(someDate As Date) As Date
//!     GetFirstDayOfMonth = DateSerial(Year(someDate), Month(someDate), 1)
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Month Boundaries
//!
//! ```vb
//! Function GetMonthStart(someDate As Date) As Date
//!     GetMonthStart = DateSerial(Year(someDate), Month(someDate), 1)
//! End Function
//!
//! Function GetMonthEnd(someDate As Date) As Date
//!     GetMonthEnd = DateSerial(Year(someDate), Month(someDate) + 1, 0)
//! End Function
//!
//! ' Get entire month range
//! Dim startDate As Date
//! Dim endDate As Date
//! startDate = GetMonthStart(Date)
//! endDate = GetMonthEnd(Date)
//! ```
//!
//! ### Year Boundaries
//!
//! ```vb
//! Function GetYearStart(someDate As Date) As Date
//!     GetYearStart = DateSerial(Year(someDate), 1, 1)
//! End Function
//!
//! Function GetYearEnd(someDate As Date) As Date
//!     GetYearEnd = DateSerial(Year(someDate), 12, 31)
//! End Function
//! ```
//!
//! ### Quarter Boundaries
//!
//! ```vb
//! Function GetQuarterStart(year As Integer, quarter As Integer) As Date
//!     Dim month As Integer
//!     month = (quarter - 1) * 3 + 1
//!     GetQuarterStart = DateSerial(year, month, 1)
//! End Function
//!
//! Function GetQuarterEnd(year As Integer, quarter As Integer) As Date
//!     Dim month As Integer
//!     month = quarter * 3
//!     GetQuarterEnd = DateSerial(year, month + 1, 0)
//! End Function
//! ```
//!
//! ### Add Months Correctly
//!
//! ```vb
//! Function AddMonths(startDate As Date, months As Integer) As Date
//!     Dim y As Integer, m As Integer, d As Integer
//!     
//!     y = Year(startDate)
//!     m = Month(startDate)
//!     d = Day(startDate)
//!     
//!     ' Add months (DateSerial handles overflow)
//!     AddMonths = DateSerial(y, m + months, d)
//! End Function
//!
//! ' Handle day overflow gracefully
//! Function AddMonthsSafe(startDate As Date, months As Integer) As Date
//!     Dim y As Integer, m As Integer, d As Integer
//!     Dim lastDay As Date
//!     
//!     y = Year(startDate)
//!     m = Month(startDate)
//!     d = Day(startDate)
//!     
//!     ' Get last day of target month
//!     lastDay = DateSerial(y, m + months + 1, 0)
//!     
//!     ' Use smaller of original day or last day of month
//!     If d > Day(lastDay) Then
//!         d = Day(lastDay)
//!     End If
//!     
//!     AddMonthsSafe = DateSerial(y, m + months, d)
//! End Function
//! ```
//!
//! ### Leap Year Detection
//!
//! ```vb
//! Function IsLeapYear(year As Integer) As Boolean
//!     Dim feb29 As Date
//!     On Error Resume Next
//!     feb29 = DateSerial(year, 2, 29)
//!     IsLeapYear = (Err.Number = 0)
//! End Function
//! ```
//!
//! ### Days in Month
//!
//! ```vb
//! Function DaysInMonth(year As Integer, month As Integer) As Integer
//!     Dim lastDay As Date
//!     lastDay = DateSerial(year, month + 1, 0)
//!     DaysInMonth = Day(lastDay)
//! End Function
//! ```
//!
//! ### Birthday This Year
//!
//! ```vb
//! Function GetBirthdayThisYear(birthDate As Date) As Date
//!     GetBirthdayThisYear = DateSerial(Year(Date), Month(birthDate), Day(birthDate))
//! End Function
//!
//! Function HasBirthdayPassed(birthDate As Date) As Boolean
//!     HasBirthdayPassed = (GetBirthdayThisYear(birthDate) <= Date)
//! End Function
//! ```
//!
//! ### Week Start (Monday)
//!
//! ```vb
//! Function GetWeekStart(someDate As Date) As Date
//!     Dim offset As Integer
//!     offset = Weekday(someDate, vbMonday) - 1
//!     GetWeekStart = DateSerial(Year(someDate), Month(someDate), Day(someDate) - offset)
//! End Function
//! ```
//!
//! ### Generate Date Range
//!
//! ```vb
//! Function GenerateMonthStarts(year As Integer) As Variant
//!     Dim dates(1 To 12) As Date
//!     Dim i As Integer
//!     
//!     For i = 1 To 12
//!         dates(i) = DateSerial(year, i, 1)
//!     Next i
//!     
//!     GenerateMonthStarts = dates
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Fiscal Year Calculations
//!
//! ```vb
//! Function GetFiscalYearStart(calendarYear As Integer, fiscalStartMonth As Integer) As Date
//!     GetFiscalYearStart = DateSerial(calendarYear, fiscalStartMonth, 1)
//! End Function
//!
//! Function GetFiscalYearEnd(calendarYear As Integer, fiscalStartMonth As Integer) As Date
//!     ' Fiscal year end is day before next fiscal year starts
//!     GetFiscalYearEnd = DateSerial(calendarYear + 1, fiscalStartMonth, 0)
//! End Function
//!
//! Function GetCurrentFiscalYear(fiscalStartMonth As Integer) As Integer
//!     Dim currentMonth As Integer
//!     currentMonth = Month(Date)
//!     
//!     If currentMonth >= fiscalStartMonth Then
//!         GetCurrentFiscalYear = Year(Date)
//!     Else
//!         GetCurrentFiscalYear = Year(Date) - 1
//!     End If
//! End Function
//! ```
//!
//! ### Date Table Generator
//!
//! ```vb
//! Sub PopulateDateDimension(startYear As Integer, endYear As Integer)
//!     Dim y As Integer, m As Integer, d As Integer
//!     Dim currentDate As Date
//!     Dim rs As ADODB.Recordset
//!     
//!     Set rs = New ADODB.Recordset
//!     ' Open recordset...
//!     
//!     For y = startYear To endYear
//!         For m = 1 To 12
//!             Dim daysInMonth As Integer
//!             daysInMonth = Day(DateSerial(y, m + 1, 0))
//!             
//!             For d = 1 To daysInMonth
//!                 currentDate = DateSerial(y, m, d)
//!                 
//!                 rs.AddNew
//!                 rs("DateKey") = Format(currentDate, "yyyymmdd")
//!                 rs("FullDate") = currentDate
//!                 rs("Year") = y
//!                 rs("Quarter") = DatePart("q", currentDate)
//!                 rs("Month") = m
//!                 rs("Day") = d
//!                 rs("DayOfWeek") = Weekday(currentDate)
//!                 rs.Update
//!             Next d
//!         Next m
//!     Next y
//! End Sub
//! ```
//!
//! ### Anniversary Calculator
//!
//! ```vb
//! Function GetAnniversaryDate(originalDate As Date, yearsLater As Integer) As Date
//!     Dim y As Integer, m As Integer, d As Integer
//!     
//!     y = Year(originalDate)
//!     m = Month(originalDate)
//!     d = Day(originalDate)
//!     
//!     GetAnniversaryDate = DateSerial(y + yearsLater, m, d)
//! End Function
//!
//! ' Handle Feb 29 anniversaries
//! Function GetAnniversaryDateSafe(originalDate As Date, yearsLater As Integer) As Date
//!     Dim y As Integer, m As Integer, d As Integer
//!     
//!     y = Year(originalDate) + yearsLater
//!     m = Month(originalDate)
//!     d = Day(originalDate)
//!     
//!     ' For Feb 29, use Feb 28 in non-leap years
//!     If m = 2 And d = 29 Then
//!         If Not IsLeapYear(y) Then
//!             d = 28
//!         End If
//!     End If
//!     
//!     GetAnniversaryDateSafe = DateSerial(y, m, d)
//! End Function
//! ```
//!
//! ### Relative Date Builder
//!
//! ```vb
//! Function BuildRelativeDate(baseDate As Date, yearOffset As Integer, _
//!                          monthOffset As Integer, dayOffset As Integer) As Date
//!     BuildRelativeDate = DateSerial(Year(baseDate) + yearOffset, _
//!                                   Month(baseDate) + monthOffset, _
//!                                   Day(baseDate) + dayOffset)
//! End Function
//!
//! ' Get date 2 years, 3 months, and 5 days from now
//! Dim futureDate As Date
//! futureDate = BuildRelativeDate(Date, 2, 3, 5)
//! ```
//!
//! ### Easter Calculation (Simplified)
//!
//! ```vb
//! Function GetEasterSunday(year As Integer) As Date
//!     ' Simplified Meeus/Jones/Butcher algorithm
//!     Dim a As Integer, b As Integer, c As Integer
//!     Dim d As Integer, e As Integer, f As Integer
//!     Dim g As Integer, h As Integer, i As Integer
//!     Dim k As Integer, l As Integer, m As Integer
//!     Dim month As Integer, day As Integer
//!     
//!     a = year Mod 19
//!     b = year \ 100
//!     c = year Mod 100
//!     d = b \ 4
//!     e = b Mod 4
//!     f = (b + 8) \ 25
//!     g = (b - f + 1) \ 3
//!     h = (19 * a + b - d - g + 15) Mod 30
//!     i = c \ 4
//!     k = c Mod 4
//!     l = (32 + 2 * e + 2 * i - h - k) Mod 7
//!     m = (a + 11 * h + 22 * l) \ 451
//!     month = (h + l - 7 * m + 114) \ 31
//!     day = ((h + l - 7 * m + 114) Mod 31) + 1
//!     
//!     GetEasterSunday = DateSerial(year, month, day)
//! End Function
//! ```
//!
//! ### Business Month-End Handler
//!
//! ```vb
//! Function GetBusinessMonthEnd(year As Integer, month As Integer) As Date
//!     Dim lastDay As Date
//!     Dim dayOfWeek As Integer
//!     
//!     lastDay = DateSerial(year, month + 1, 0)
//!     dayOfWeek = Weekday(lastDay)
//!     
//!     ' If weekend, back up to Friday
//!     If dayOfWeek = vbSaturday Then
//!         lastDay = DateSerial(year, month + 1, -1)  ' Friday
//!     ElseIf dayOfWeek = vbSunday Then
//!         lastDay = DateSerial(year, month + 1, -2)  ' Friday
//!     End If
//!     
//!     GetBusinessMonthEnd = lastDay
//! End Function
//! ```
//!
//! ### Date Validator
//!
//! ```vb
//! Function IsValidDate(year As Integer, month As Integer, day As Integer) As Boolean
//!     On Error Resume Next
//!     Dim testDate As Date
//!     testDate = DateSerial(year, month, day)
//!     
//!     IsValidDate = (Err.Number = 0) And _
//!                   (Year(testDate) = year) And _
//!                   (Month(testDate) = month) And _
//!                   (Day(testDate) = day)
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeDateSerial(year As Integer, month As Integer, day As Integer) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     ' Validate ranges
//!     If year < 100 Or year > 9999 Then
//!         SafeDateSerial = Null
//!         Exit Function
//!     End If
//!     
//!     SafeDateSerial = DateSerial(year, month, day)
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeDateSerial = Null
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 5** (Invalid procedure call): Year outside 100-9999 range
//! - **Error 13** (Type mismatch): Non-numeric arguments
//! - **Error 6** (Overflow): Result date outside valid range
//!
//! ## Performance Considerations
//!
//! - `DateSerial` is very fast for date construction
//! - More efficient than parsing date strings
//! - Automatic range adjustment is performant
//! - No string formatting overhead
//! - Ideal for loop-based date generation
//!
//! ## Best Practices
//!
//! ### Use for Date Construction
//!
//! ```vb
//! ' Good - Clear and unambiguous
//! deadline = DateSerial(2025, 12, 31)
//!
//! ' Avoid - Locale-dependent
//! deadline = CDate("12/31/2025")  ' May fail in different locales
//! ```
//!
//! ### Leverage Range Adjustment
//!
//! ```vb
//! ' Use day 0 for last day of previous month
//! lastDayPrevMonth = DateSerial(year, month, 0)
//!
//! ' Use month 0 for last month of previous year
//! dec31 = DateSerial(year, 0, 31)
//! ```
//!
//! ### Validate Before Critical Operations
//!
//! ```vb
//! If IsValidDate(y, m, d) Then
//!     result = DateSerial(y, m, d)
//! Else
//!     MsgBox "Invalid date components"
//! End If
//! ```
//!
//! ### Extract Components for Manipulation
//!
//! ```vb
//! ' Extract, modify, rebuild
//! y = Year(someDate)
//! m = Month(someDate)
//! d = 1  ' First of month
//! newDate = DateSerial(y, m, d)
//! ```
//!
//! ## Comparison with Other Functions
//!
//! ### `DateSerial` vs Date Literals
//!
//! ```vb
//! ' DateSerial - Dynamic, programmatic
//! dt = DateSerial(Year(Date), 12, 25)
//!
//! ' Date Literal - Static, hardcoded
//! dt = #12/25/2025#
//! ```
//!
//! ### `DateSerial` vs `DateValue`
//!
//! ```vb
//! ' `DateSerial` - From numeric components
//! dt = DateSerial(2025, 12, 25)
//!
//! ' `DateValue` - From string representation
//! dt = DateValue("December 25, 2025")
//! ```
//!
//! ### `DateSerial` vs `DateAdd`
//!
//! ```vb
//! ' DateSerial - Absolute date construction
//! nextMonth = DateSerial(Year(Date), Month(Date) + 1, 1)
//!
//! ' DateAdd - Relative date calculation
//! nextMonth = DateAdd("m", 1, Date)
//! ```
//!
//! ## Limitations
//!
//! - Year must be between 100 and 9999
//! - Two-digit year interpretation fixed (0-29=2000-2029, 30-99=1930-1999)
//! - Always returns midnight (no time component)
//! - Cannot directly specify time components
//! - Invalid dates may raise runtime errors
//!
//! ## Related Functions
//!
//! - `DateValue`: Converts a string to a date
//! - `TimeSerial`: Creates a time from hour, minute, and second
//! - `DateAdd`: Adds a time interval to a date
//! - `Year`, `Month`, `Day`: Extract date components
//! - `Date`: Returns current system date
//! - `Now`: Returns current date and time
//! - `IsDate`: Tests if a value can be converted to a date
//! - `CDate`: Converts an expression to a Date

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_dateserial_basic() {
        let source = r#"
birthday = DateSerial(1990, 5, 15)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_with_variables() {
        let source = r#"
result = DateSerial(y, m, d)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_last_day_of_month() {
        let source = r#"
lastDay = DateSerial(2025, 2, 0)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_with_year_function() {
        let source = r#"
newYear = DateSerial(Year(Date), 1, 1)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_month_overflow() {
        let source = r#"
result = DateSerial(2025, 13, 1)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_day_overflow() {
        let source = r#"
result = DateSerial(2025, 1, 32)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_negative_month() {
        let source = r#"
result = DateSerial(2025, -1, 1)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_in_function() {
        let source = r#"
Function GetLastDay(y As Integer, m As Integer) As Date
    GetLastDay = DateSerial(y, m + 1, 0)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_with_expressions() {
        let source = r#"
result = DateSerial(Year(Date), Month(Date) + 1, 0)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_in_loop() {
        let source = r#"
For i = 1 To 12
    dates(i) = DateSerial(2025, i, 1)
Next i
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_quarter_start() {
        let source = r#"
quarterStart = DateSerial(year, (quarter - 1) * 3 + 1, 1)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_with_day_function() {
        let source = r#"
result = DateSerial(y, m, Day(someDate))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_in_comparison() {
        let source = r#"
If Date > DateSerial(2025, 12, 31) Then
    MsgBox "Past deadline"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_arithmetic() {
        let source = r#"
offset = DateSerial(y, m, d) - Date
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_with_constants() {
        let source = r#"
Const YEAR As Integer = 2025
result = DateSerial(YEAR, 1, 1)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_nested_in_format() {
        let source = r#"
formatted = Format(DateSerial(2025, 12, 25), "yyyy-mm-dd")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_in_select_case() {
        let source = r#"
Select Case Date
    Case DateSerial(2025, 1, 1)
        MsgBox "New Year"
End Select
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_multiple_calls() {
        let source = r#"
startDate = DateSerial(2025, 1, 1)
endDate = DateSerial(2025, 12, 31)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_with_datepart() {
        let source = r#"
result = DateSerial(DatePart("yyyy", Date), 1, 1)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_in_array() {
        let source = r#"
dates(0) = DateSerial(2025, 1, 1)
dates(1) = DateSerial(2025, 2, 1)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_week_start() {
        let source = r#"
weekStart = DateSerial(Year(Date), Month(Date), Day(Date) - offset)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_with_addition() {
        let source = r#"
anniversary = DateSerial(Year(original) + years, Month(original), Day(original))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_zero_day() {
        let source = r#"
lastMonth = DateSerial(2025, 2, 0)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_in_msgbox() {
        let source = r#"
MsgBox "Date: " & DateSerial(2025, 12, 25)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn test_dateserial_relative_calculation() {
        let source = r#"
result = DateSerial(Year(base) + yOffset, Month(base) + mOffset, Day(base) + dOffset)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("DateSerial"));
        assert!(debug.contains("Identifier"));
    }
}
