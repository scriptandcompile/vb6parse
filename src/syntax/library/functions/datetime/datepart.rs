//! # `DatePart` Function
//!
//! Returns a `Variant` (`Integer`) containing the specified part of a given date.
//!
//! ## Syntax
//!
//! ```vb
//! DatePart(interval, date[, firstdayofweek[, firstweekofyear]])
//! ```
//!
//! ## Parameters
//!
//! - **interval**: Required. `String` expression that is the interval of time you want to return.
//!   See the Interval Settings section for valid values.
//! - **date**: Required. `Variant` (`Date`) value that you want to evaluate.
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
//! | Setting | Description | Return Range |
//! |---------|-------------|--------------|
//! | "yyyy" | Year | 100-9999 |
//! | "q" | Quarter | 1-4 |
//! | "m" | Month | 1-12 |
//! | "y" | Day of year | 1-366 |
//! | "d" | Day | 1-31 |
//! | "w" | Weekday | 1-7 (Sunday=1) |
//! | "ww" | Week of year | 1-53 |
//! | "h" | Hour | 0-23 |
//! | "n" | Minute | 0-59 |
//! | "s" | Second | 0-59 |
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
//! Returns an `Integer` representing the specified part of the date. Returns `Null` if the date is `Null`.
//!
//! ## Remarks
//!
//! The `DatePart` function is used to extract a specific component from a date value.
//! It's particularly useful for date-based calculations, filtering, and grouping operations.
//!
//! **Important Characteristics:**
//!
//! - More flexible than `Year()`, `Month()`, or `Day()` functions.
//! - Can extract quarter, week, and day of year.
//! - Weekday numbering depends on `firstdayofweek` parameter.
//! - Week numbering depends on `firstweekofyear` parameter.
//! - Hours use 24-hour format (0-23).
//! - Sunday is 1 by default for weekday ("w").
//! - Compatible with SQL Server's `DATEPART` function
//!
//! ## Equivalent Simple Functions
//!
//! Some intervals have equivalent dedicated functions:
//! - `DatePart("yyyy", date)` = `Year(date)`
//! - `DatePart("m", date)` = `Month(date)`
//! - `DatePart("d", date)` = `Day(date)`
//! - `DatePart("w", date)` = `Weekday(date)`
//! - `DatePart("h", date)` = `Hour(date)`
//! - `DatePart("n", date)` = `Minute(date)`
//! - `DatePart("s", date)` = `Second(date)`
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! Dim testDate As Date
//! testDate = #3/15/2025 14:30:45#
//!
//! ' Extract various parts
//! MsgBox "Year: " & DatePart("yyyy", testDate)      ' 2025
//! MsgBox "Quarter: " & DatePart("q", testDate)      ' 1
//! MsgBox "Month: " & DatePart("m", testDate)        ' 3
//! MsgBox "Day: " & DatePart("d", testDate)          ' 15
//! MsgBox "Day of Year: " & DatePart("y", testDate)  ' 74
//! MsgBox "Weekday: " & DatePart("w", testDate)      ' Varies by day
//! MsgBox "Week: " & DatePart("ww", testDate)        ' Week number
//! MsgBox "Hour: " & DatePart("h", testDate)         ' 14
//! MsgBox "Minute: " & DatePart("n", testDate)       ' 30
//! MsgBox "Second: " & DatePart("s", testDate)       ' 45
//! ```
//!
//! ### Quarter Calculation
//!
//! ```vb
//! Function GetQuarter(dateValue As Date) As Integer
//!     GetQuarter = DatePart("q", dateValue)
//! End Function
//!
//! ' Usage
//! Dim currentQuarter As Integer
//! currentQuarter = GetQuarter(Date)
//! MsgBox "We are in Q" & currentQuarter
//! ```
//!
//! ### Week Number
//!
//! ```vb
//! Function GetWeekNumber(dateValue As Date) As Integer
//!     ' ISO week number (Monday start, 4-day rule)
//!     GetWeekNumber = DatePart("ww", dateValue, vbMonday, vbFirstFourDays)
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Fiscal Quarter Determination
//!
//! ```vb
//! Function GetFiscalQuarter(dateValue As Date, fiscalYearStart As Integer) As Integer
//!     ' fiscalYearStart is the month number (e.g., 4 for April)
//!     Dim currentMonth As Integer
//!     Dim adjustedMonth As Integer
//!     
//!     currentMonth = DatePart("m", dateValue)
//!     adjustedMonth = currentMonth - fiscalYearStart + 1
//!     
//!     If adjustedMonth <= 0 Then
//!         adjustedMonth = adjustedMonth + 12
//!     End If
//!     
//!     GetFiscalQuarter = Int((adjustedMonth - 1) / 3) + 1
//! End Function
//! ```
//!
//! ### Group By Time Period
//!
//! ```vb
//! Function GroupByPeriod(dateValue As Date, period As String) As String
//!     Select Case LCase(period)
//!         Case "year"
//!             GroupByPeriod = CStr(DatePart("yyyy", dateValue))
//!         Case "quarter"
//!             GroupByPeriod = DatePart("yyyy", dateValue) & "-Q" & DatePart("q", dateValue)
//!         Case "month"
//!             GroupByPeriod = DatePart("yyyy", dateValue) & "-" & Format(DatePart("m", dateValue), "00")
//!         Case "week"
//!             GroupByPeriod = DatePart("yyyy", dateValue) & "-W" & Format(DatePart("ww", dateValue), "00")
//!         Case Else
//!             GroupByPeriod = Format(dateValue, "yyyy-mm-dd")
//!     End Select
//! End Function
//! ```
//!
//! ### Day Name from Weekday
//!
//! ```vb
//! Function GetDayName(dateValue As Date) As String
//!     Select Case DatePart("w", dateValue)
//!         Case 1: GetDayName = "Sunday"
//!         Case 2: GetDayName = "Monday"
//!         Case 3: GetDayName = "Tuesday"
//!         Case 4: GetDayName = "Wednesday"
//!         Case 5: GetDayName = "Thursday"
//!         Case 6: GetDayName = "Friday"
//!         Case 7: GetDayName = "Saturday"
//!     End Select
//! End Function
//! ```
//!
//! ### Time of Day Category
//!
//! ```vb
//! Function GetTimeOfDay(dateValue As Date) As String
//!     Dim hour As Integer
//!     hour = DatePart("h", dateValue)
//!     
//!     Select Case hour
//!         Case 0 To 5
//!             GetTimeOfDay = "Night"
//!         Case 6 To 11
//!             GetTimeOfDay = "Morning"
//!         Case 12 To 17
//!             GetTimeOfDay = "Afternoon"
//!         Case 18 To 23
//!             GetTimeOfDay = "Evening"
//!     End Select
//! End Function
//! ```
//!
//! ### Business Hour Check
//!
//! ```vb
//! Function IsBusinessHours(checkTime As Date) As Boolean
//!     Dim hour As Integer
//!     Dim weekday As Integer
//!     
//!     hour = DatePart("h", checkTime)
//!     weekday = DatePart("w", checkTime)
//!     
//!     ' Monday-Friday, 9 AM - 5 PM
//!     If weekday >= 2 And weekday <= 6 Then  ' Mon-Fri
//!         If hour >= 9 And hour < 17 Then
//!             IsBusinessHours = True
//!         End If
//!     End If
//! End Function
//! ```
//!
//! ### Month Name Lookup
//!
//! ```vb
//! Function GetMonthName(dateValue As Date) As String
//!     Dim monthNames As Variant
//!     Dim monthNum As Integer
//!     
//!     monthNames = Array("January", "February", "March", "April", "May", "June", _
//!                       "July", "August", "September", "October", "November", "December")
//!     
//!     monthNum = DatePart("m", dateValue)
//!     GetMonthName = monthNames(monthNum - 1)
//! End Function
//! ```
//!
//! ### Quarter End Date
//!
//! ```vb
//! Function GetQuarterEnd(dateValue As Date) As Date
//!     Dim quarter As Integer
//!     Dim year As Integer
//!     Dim endMonth As Integer
//!     
//!     quarter = DatePart("q", dateValue)
//!     year = DatePart("yyyy", dateValue)
//!     endMonth = quarter * 3
//!     
//!     GetQuarterEnd = DateSerial(year, endMonth + 1, 0)  ' Last day of quarter
//! End Function
//! ```
//!
//! ### Data Binning by Hour
//!
//! ```vb
//! Function GetHourBucket(timestamp As Date) As String
//!     Dim hour As Integer
//!     hour = DatePart("h", timestamp)
//!     GetHourBucket = Format(hour, "00") & ":00"
//! End Function
//!
//! ' Use for grouping log entries
//! Sub AnalyzeLogs()
//!     Dim entry As Date
//!     Dim bucket As String
//!     
//!     For Each entry In logEntries
//!         bucket = GetHourBucket(entry)
//!         hourCounts(bucket) = hourCounts(bucket) + 1
//!     Next
//! End Sub
//! ```
//!
//! ## Advanced Usage
//!
//! ### ISO 8601 Week Number
//!
//! ```vb
//! Function GetISOWeekNumber(dateValue As Date) As Integer
//!     ' ISO 8601 week number: Monday start, 4-day rule
//!     GetISOWeekNumber = DatePart("ww", dateValue, vbMonday, vbFirstFourDays)
//! End Function
//!
//! Function GetISOYear(dateValue As Date) As Integer
//!     ' Year for ISO week (may differ from calendar year)
//!     Dim weekNum As Integer
//!     Dim month As Integer
//!     
//!     weekNum = GetISOWeekNumber(dateValue)
//!     month = DatePart("m", dateValue)
//!     
//!     If month = 1 And weekNum > 51 Then
//!         GetISOYear = DatePart("yyyy", dateValue) - 1
//!     ElseIf month = 12 And weekNum = 1 Then
//!         GetISOYear = DatePart("yyyy", dateValue) + 1
//!     Else
//!         GetISOYear = DatePart("yyyy", dateValue)
//!     End If
//! End Function
//! ```
//!
//! ### Dynamic Date Grouping
//!
//! ```vb
//! Function GetDateKey(dateValue As Date, granularity As String) As String
//!     Dim year As Integer
//!     Dim month As Integer
//!     Dim day As Integer
//!     Dim week As Integer
//!     Dim quarter As Integer
//!     
//!     year = DatePart("yyyy", dateValue)
//!     
//!     Select Case LCase(granularity)
//!         Case "year"
//!             GetDateKey = CStr(year)
//!         
//!         Case "quarter"
//!             quarter = DatePart("q", dateValue)
//!             GetDateKey = year & "Q" & quarter
//!         
//!         Case "month"
//!             month = DatePart("m", dateValue)
//!             GetDateKey = year & Format(month, "00")
//!         
//!         Case "week"
//!             week = DatePart("ww", dateValue, vbMonday)
//!             GetDateKey = year & "W" & Format(week, "00")
//!         
//!         Case "day"
//!             month = DatePart("m", dateValue)
//!             day = DatePart("d", dateValue)
//!             GetDateKey = year & Format(month, "00") & Format(day, "00")
//!         
//!         Case Else
//!             GetDateKey = Format(dateValue, "yyyymmdd")
//!     End Select
//! End Function
//! ```
//!
//! ### Custom Calendar System
//!
//! ```vb
//! Type CustomCalendar
//!     Year As Integer
//!     Period As Integer
//!     Week As Integer
//!     Day As Integer
//! End Type
//!
//! Function ConvertToCustomCalendar(dateValue As Date) As CustomCalendar
//!     Dim cal As CustomCalendar
//!     Dim yearStart As Date
//!     Dim dayOfYear As Integer
//!     
//!     cal.Year = DatePart("yyyy", dateValue)
//!     
//!     ' 13 periods of 4 weeks each
//!     yearStart = DateSerial(cal.Year, 1, 1)
//!     dayOfYear = DatePart("y", dateValue)
//!     
//!     cal.Week = Int((dayOfYear - 1) / 7) + 1
//!     cal.Period = Int((cal.Week - 1) / 4) + 1
//!     cal.Day = DatePart("w", dateValue, vbMonday)
//!     
//!     ConvertToCustomCalendar = cal
//! End Function
//! ```
//!
//! ### Time Series Aggregation
//!
//! ```vb
//! Function AggregateByInterval(dates() As Date, values() As Double, _
//!                             interval As String) As Collection
//!     Dim results As New Collection
//!     Dim i As Long
//!     Dim key As String
//!     Dim total As Double
//!     Dim count As Long
//!     
//!     For i = LBound(dates) To UBound(dates)
//!         key = GetDateKey(dates(i), interval)
//!         
//!         On Error Resume Next
//!         total = results(key)
//!         If Err.Number <> 0 Then
//!             results.Add values(i), key
//!         Else
//!             results.Remove key
//!             results.Add total + values(i), key
//!         End If
//!         On Error GoTo 0
//!     Next i
//!     
//!     Set AggregateByInterval = results
//! End Function
//! ```
//!
//! ### Shift Schedule Detector
//!
//! ```vb
//! Function GetShift(timestamp As Date) As String
//!     Dim hour As Integer
//!     Dim weekday As Integer
//!     
//!     hour = DatePart("h", timestamp)
//!     weekday = DatePart("w", timestamp)
//!     
//!     ' Weekend check
//!     If weekday = 1 Or weekday = 7 Then
//!         GetShift = "Weekend"
//!         Exit Function
//!     End If
//!     
//!     ' Shift determination
//!     Select Case hour
//!         Case 6 To 13
//!             GetShift = "Morning Shift"
//!         Case 14 To 21
//!             GetShift = "Afternoon Shift"
//!         Case Else
//!             GetShift = "Night Shift"
//!     End Select
//! End Function
//! ```
//!
//! ### Calendar Week Display
//!
//! ```vb
//! Function FormatCalendarWeek(dateValue As Date, Optional useISO As Boolean = False) As String
//!     Dim year As Integer
//!     Dim week As Integer
//!     
//!     If useISO Then
//!         year = GetISOYear(dateValue)
//!         week = GetISOWeekNumber(dateValue)
//!     Else
//!         year = DatePart("yyyy", dateValue)
//!         week = DatePart("ww", dateValue)
//!     End If
//!     
//!     FormatCalendarWeek = year & "-W" & Format(week, "00")
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeDatePart(interval As String, dateValue As Variant) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     ' Validate date
//!     If Not IsDate(dateValue) Then
//!         SafeDatePart = Null
//!         Exit Function
//!     End If
//!     
//!     ' Validate interval
//!     Select Case LCase(interval)
//!         Case "yyyy", "q", "m", "y", "d", "w", "ww", "h", "n", "s"
//!             SafeDatePart = DatePart(interval, CDate(dateValue))
//!         Case Else
//!             SafeDatePart = Null
//!     End Select
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeDatePart = Null
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 5** (Invalid procedure call): Invalid interval string
//! - **Error 13** (Type mismatch): Non-date value passed as date parameter
//!
//! ## Performance Considerations
//!
//! - `DatePart` is efficient for single extractions
//! - For multiple parts from same date, consider using dedicated functions:
//!   ```vb
//!   ' Less efficient
//!   y = DatePart("yyyy", d)
//!   m = DatePart("m", d)
//!   d = DatePart("d", d)
//!   
//!   ' More efficient
//!   y = Year(d)
//!   m = Month(d)
//!   d = Day(d)
//!   ```
//! - Week calculations are more expensive than other intervals
//! - Cache results when processing large datasets
//!
//! ## Best Practices
//!
//! ### Use Named Constants
//!
//! ```vb
//! ' Define interval constants
//! Const INTERVAL_YEAR As String = "yyyy"
//! Const INTERVAL_QUARTER As String = "q"
//! Const INTERVAL_MONTH As String = "m"
//! Const INTERVAL_WEEK As String = "ww"
//!
//! ' Use in code
//! quarter = DatePart(INTERVAL_QUARTER, Date)
//! ```
//!
//! ### Prefer Specific Functions for Simple Cases
//!
//! ```vb
//! ' Good - Use specific function
//! y = Year(someDate)
//!
//! ' Less clear - Using DatePart
//! y = DatePart("yyyy", someDate)
//! ```
//!
//! ### Be Aware of Weekday Numbering
//!
//! ```vb
//! ' Default: Sunday = 1
//! day = DatePart("w", Date)
//!
//! ' Explicit: Monday = 1
//! day = DatePart("w", Date, vbMonday)
//! ```
//!
//! ## Comparison with Other Functions
//!
//! ### `DatePart` vs Dedicated Functions
//!
//! ```vb
//! ' DatePart - Flexible, supports all intervals
//! quarter = DatePart("q", Date)
//! dayOfYear = DatePart("y", Date)
//!
//! ' Dedicated - Simpler, more readable for common cases
//! year = Year(Date)
//! month = Month(Date)
//! day = Day(Date)
//! weekday = Weekday(Date)
//! ```
//!
//! ## Limitations
//!
//! - No millisecond support
//! - Week numbering can be confusing with different standards (ISO vs US)
//! - Quarter calculation doesn't support fiscal quarters directly
//! - No built-in locale-aware day/month names
//! - `FirstWeekOfYear` affects week numbering interpretation
//!
//! ## Related Functions
//!
//! - `Year`: Returns the year part of a date
//! - `Month`: Returns the month part of a date
//! - `Day`: Returns the day part of a date
//! - `Weekday`: Returns the day of the week
//! - `Hour`: Returns the hour part of a time
//! - `Minute`: Returns the minute part of a time
//! - `Second`: Returns the second part of a time
//! - `DateAdd`: Adds a time interval to a date
//! - `DateDiff`: Returns the difference between two dates
//! - `DateSerial`: Creates a date from year, month, and day values
//! - `Format`: Formats a date as a string (alternative for custom formatting)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn datepart_basic() {
        let source = r#"
year = DatePart("yyyy", Date)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("year"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"yyyy\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                DateKeyword,
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
    fn datepart_quarter() {
        let source = r#"
quarter = DatePart("q", currentDate)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("quarter"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"q\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("currentDate"),
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
    fn datepart_month() {
        let source = r#"
month = DatePart("m", #3/15/2025#)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("month"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"m\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            LiteralExpression {
                                DateLiteral ("#3/15/2025#"),
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
    fn datepart_day() {
        let source = r#"
day = DatePart("d", Date)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("day"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"d\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                DateKeyword,
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
    fn datepart_day_of_year() {
        let source = r#"
dayOfYear = DatePart("y", Date)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("dayOfYear"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"y\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                DateKeyword,
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
    fn datepart_weekday() {
        let source = r#"
weekday = DatePart("w", Date)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("weekday"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"w\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                DateKeyword,
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
    fn datepart_week() {
        let source = r#"
week = DatePart("ww", Date)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("week"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"ww\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                DateKeyword,
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
    fn datepart_hour() {
        let source = r#"
hour = DatePart("h", Now)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("hour"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"h\""),
                            },
                        },
                        Comma,
                        Whitespace,
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
    fn datepart_minute() {
        let source = r#"
minute = DatePart("n", timestamp)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("minute"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"n\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("timestamp"),
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
    fn datepart_second() {
        let source = r#"
second = DatePart("s", Now)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("second"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"s\""),
                            },
                        },
                        Comma,
                        Whitespace,
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
    fn datepart_with_firstdayofweek() {
        let source = r#"
weekday = DatePart("w", Date, vbMonday)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("weekday"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"w\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                DateKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbMonday"),
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
    fn datepart_with_all_params() {
        let source = r#"
week = DatePart("ww", Date, vbMonday, vbFirstFourDays)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("week"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"ww\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                DateKeyword,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbMonday"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                Identifier ("vbFirstFourDays"),
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
    fn datepart_in_function() {
        let source = r#"
Function GetQuarter(d As Date) As Integer
    GetQuarter = DatePart("q", d)
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetQuarter"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("d"),
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
                            Identifier ("GetQuarter"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("DatePart"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"q\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("d"),
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
    fn datepart_in_select_case() {
        let source = r#"
Select Case DatePart("q", Date)
    Case 1
        MsgBox "Q1"
    Case 2
        MsgBox "Q2"
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
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"q\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                DateKeyword,
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
                    IntegerLiteral ("1"),
                    Newline,
                    StatementList {
                        Whitespace,
                        CallStatement {
                            Identifier ("MsgBox"),
                            Whitespace,
                            StringLiteral ("\"Q1\""),
                            Newline,
                        },
                        Whitespace,
                    },
                },
                CaseClause {
                    CaseKeyword,
                    Whitespace,
                    IntegerLiteral ("2"),
                    Newline,
                    StatementList {
                        Whitespace,
                        CallStatement {
                            Identifier ("MsgBox"),
                            Whitespace,
                            StringLiteral ("\"Q2\""),
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
    fn datepart_in_if() {
        let source = r#"
If DatePart("h", Now) >= 17 Then
    MsgBox "After hours"
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
                        Identifier ("DatePart"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                StringLiteralExpression {
                                    StringLiteral ("\"h\""),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                IdentifierExpression {
                                    Identifier ("Now"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    GreaterThanOrEqualOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("17"),
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
                        StringLiteral ("\"After hours\""),
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
    fn datepart_concatenation() {
        let source = r#"
key = DatePart("yyyy", Date) & "-" & DatePart("m", Date)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("key"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    BinaryExpression {
                        CallExpression {
                            Identifier ("DatePart"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"yyyy\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        DateKeyword,
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"-\""),
                        },
                    },
                    Whitespace,
                    Ampersand,
                    Whitespace,
                    CallExpression {
                        Identifier ("DatePart"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                StringLiteralExpression {
                                    StringLiteral ("\"m\""),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                IdentifierExpression {
                                    DateKeyword,
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
    fn datepart_in_calculation() {
        let source = r#"
endMonth = DatePart("q", Date) * 3
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("endMonth"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    CallExpression {
                        Identifier ("DatePart"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                StringLiteralExpression {
                                    StringLiteral ("\"q\""),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                IdentifierExpression {
                                    DateKeyword,
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    MultiplicationOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("3"),
                    },
                },
                Newline,
            },
        ]);
    }

    #[test]
    fn datepart_in_loop() {
        let source = r#"
For i = 1 To count
    months(i) = DatePart("m", dates(i))
Next i
"#;
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
                IdentifierExpression {
                    Identifier ("count"),
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        CallExpression {
                            Identifier ("months"),
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
                            Identifier ("DatePart"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"m\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    CallExpression {
                                        Identifier ("dates"),
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
    fn datepart_comparison() {
        let source = r#"
If DatePart("w", Date) = vbSaturday Then
    MsgBox "Weekend"
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
                        Identifier ("DatePart"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                StringLiteralExpression {
                                    StringLiteral ("\"w\""),
                                },
                            },
                            Comma,
                            Whitespace,
                            Argument {
                                IdentifierExpression {
                                    DateKeyword,
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    IdentifierExpression {
                        Identifier ("vbSaturday"),
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
                        StringLiteral ("\"Weekend\""),
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
    fn datepart_multiple_calls() {
        let source = r#"
y = DatePart("yyyy", Date)
m = DatePart("m", Date)
d = DatePart("d", Date)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("y"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"yyyy\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                DateKeyword,
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("m"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"m\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                DateKeyword,
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("d"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            StringLiteralExpression {
                                StringLiteral ("\"d\""),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                DateKeyword,
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
    fn datepart_in_msgbox() {
        let source = r#"
MsgBox "Quarter: " & DatePart("q", Date)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("MsgBox"),
                Whitespace,
                StringLiteral ("\"Quarter: \""),
                Whitespace,
                Ampersand,
                Whitespace,
                Identifier ("DatePart"),
                LeftParenthesis,
                StringLiteral ("\"q\""),
                Comma,
                Whitespace,
                DateKeyword,
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn datepart_with_format() {
        let source = r#"
formatted = Format(DatePart("m", Date), "00")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
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
                                Identifier ("DatePart"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        StringLiteralExpression {
                                            StringLiteral ("\"m\""),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            DateKeyword,
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
    fn datepart_nested_in_dateserial() {
        let source = r#"
quarterEnd = DateSerial(DatePart("yyyy", Date), DatePart("q", Date) * 3 + 1, 0)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("quarterEnd"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DateSerial"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("DatePart"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        StringLiteralExpression {
                                            StringLiteral ("\"yyyy\""),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            DateKeyword,
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BinaryExpression {
                                BinaryExpression {
                                    CallExpression {
                                        Identifier ("DatePart"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                StringLiteralExpression {
                                                    StringLiteral ("\"q\""),
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                IdentifierExpression {
                                                    DateKeyword,
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                    Whitespace,
                                    MultiplicationOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("3"),
                                    },
                                },
                                Whitespace,
                                AdditionOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("1"),
                                },
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
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
    fn datepart_with_variable_interval() {
        let source = r#"
Dim interval As String
interval = "q"
result = DatePart(interval, Date)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("interval"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("interval"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                StringLiteralExpression {
                    StringLiteral ("\"q\""),
                },
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("result"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("DatePart"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("interval"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            IdentifierExpression {
                                DateKeyword,
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
    fn datepart_range_check() {
        let source = r#"
If DatePart("h", Now) >= 9 And DatePart("h", Now) < 17 Then
    MsgBox "Business hours"
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
                    BinaryExpression {
                        CallExpression {
                            Identifier ("DatePart"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"h\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("Now"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        GreaterThanOrEqualOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("9"),
                        },
                    },
                    Whitespace,
                    AndKeyword,
                    Whitespace,
                    BinaryExpression {
                        CallExpression {
                            Identifier ("DatePart"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"h\""),
                                    },
                                },
                                Comma,
                                Whitespace,
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
                            IntegerLiteral ("17"),
                        },
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
                        StringLiteral ("\"Business hours\""),
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
}
