//! # Day Function
//!
//! Returns a whole number between 1 and 31, inclusive, representing the day of the month.
//!
//! ## Syntax
//!
//! ```vb
//! Day(date)
//! ```
//!
//! ## Parameters
//!
//! - **date**: Required. Any Variant, numeric expression, string expression, or any combination
//!   that can represent a date. If date contains Null, Null is returned.
//!
//! ## Return Value
//!
//! Returns an Integer representing the day of the month (1-31). Returns Null if the date
//! parameter contains Null.
//!
//! ## Remarks
//!
//! The `Day` function extracts the day component from a date value. It's one of the primary
//! date component extraction functions in VB6, along with `Year`, `Month`, `Weekday`, `Hour`,
//! `Minute`, and `Second`.
//!
//! **Important Characteristics:**
//!
//! - Returns values 1-31 depending on the month
//! - Works with Date variables, date literals, and date expressions
//! - Returns Null if input is Null (Variant behavior)
//! - Automatically handles different month lengths
//! - Time portion of datetime values is ignored
//! - Type mismatch error if argument cannot be interpreted as date
//! - Can be used with date arithmetic results
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Extract day from date literal
//! Dim d As Integer
//! d = Day(#1/15/2025#)  ' Returns 15
//!
//! ' From Date variable
//! Dim birthday As Date
//! birthday = #5/23/1990#
//! d = Day(birthday)  ' Returns 23
//!
//! ' From current date
//! d = Day(Date)  ' Returns current day
//! ```
//!
//! ### With Date Functions
//!
//! ```vb
//! ' Extract day from DateSerial result
//! Dim constructedDate As Date
//! constructedDate = DateSerial(2025, 12, 25)
//! Dim dayNum As Integer
//! dayNum = Day(constructedDate)  ' Returns 25
//!
//! ' From DateAdd calculation
//! Dim futureDate As Date
//! futureDate = DateAdd("d", 10, Date)
//! dayNum = Day(futureDate)
//! ```
//!
//! ### Date Validation
//!
//! ```vb
//! Function IsLastDayOfMonth(dateValue As Date) As Boolean
//!     Dim nextDay As Date
//!     nextDay = DateAdd("d", 1, dateValue)
//!     IsLastDayOfMonth = (Day(nextDay) = 1)
//! End Function
//!
//! ' Alternative using month comparison
//! Function IsLastDayOfMonth2(dateValue As Date) As Boolean
//!     IsLastDayOfMonth2 = (Day(dateValue + 1) = 1)
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Get Days in Month
//!
//! ```vb
//! Function DaysInMonth(dateValue As Date) As Integer
//!     ' Get last day of the month
//!     Dim firstOfNextMonth As Date
//!     firstOfNextMonth = DateSerial(Year(dateValue), Month(dateValue) + 1, 1)
//!     DaysInMonth = Day(firstOfNextMonth - 1)
//! End Function
//!
//! ' Alternative approach
//! Function DaysInMonth2(yr As Integer, mo As Integer) As Integer
//!     DaysInMonth2 = Day(DateSerial(yr, mo + 1, 0))
//! End Function
//! ```
//!
//! ### Extract Date Components
//!
//! ```vb
//! Sub DisplayDateParts(dateValue As Date)
//!     Dim yr As Integer, mo As Integer, dy As Integer
//!     
//!     yr = Year(dateValue)
//!     mo = Month(dateValue)
//!     dy = Day(dateValue)
//!     
//!     MsgBox "Year: " & yr & ", Month: " & mo & ", Day: " & dy
//! End Sub
//! ```
//!
//! ### Day-Based Filtering
//!
//! ```vb
//! Function IsFirstHalfOfMonth(dateValue As Date) As Boolean
//!     IsFirstHalfOfMonth = (Day(dateValue) <= 15)
//! End Function
//!
//! Function IsSecondHalfOfMonth(dateValue As Date) As Boolean
//!     IsSecondHalfOfMonth = (Day(dateValue) > 15)
//! End Function
//!
//! Function IsMonthStart(dateValue As Date) As Boolean
//!     IsMonthStart = (Day(dateValue) = 1)
//! End Function
//! ```
//!
//! ### Date Comparison by Day
//!
//! ```vb
//! Function SameDayOfMonth(date1 As Date, date2 As Date) As Boolean
//!     SameDayOfMonth = (Day(date1) = Day(date2))
//! End Function
//!
//! Function CompareDays(date1 As Date, date2 As Date) As Integer
//!     ' Returns: -1 if date1's day < date2's day
//!     '           0 if same day
//!     '           1 if date1's day > date2's day
//!     Dim d1 As Integer, d2 As Integer
//!     d1 = Day(date1)
//!     d2 = Day(date2)
//!     
//!     If d1 < d2 Then
//!         CompareDays = -1
//!     ElseIf d1 > d2 Then
//!         CompareDays = 1
//!     Else
//!         CompareDays = 0
//!     End If
//! End Function
//! ```
//!
//! ### Reconstruct Date with Modified Day
//!
//! ```vb
//! Function ChangeDay(originalDate As Date, newDay As Integer) As Date
//!     ' Create new date with same year/month but different day
//!     ChangeDay = DateSerial(Year(originalDate), Month(originalDate), newDay)
//! End Function
//!
//! ' Move to specific day of current month
//! Function MoveToDayOfMonth(dayNum As Integer) As Date
//!     MoveToDayOfMonth = DateSerial(Year(Date), Month(Date), dayNum)
//! End Function
//! ```
//!
//! ### Loop Through Days of Month
//!
//! ```vb
//! Sub ProcessAllDaysInMonth(yr As Integer, mo As Integer)
//!     Dim dayCount As Integer
//!     Dim currentDay As Date
//!     Dim i As Integer
//!     
//!     dayCount = Day(DateSerial(yr, mo + 1, 0))
//!     
//!     For i = 1 To dayCount
//!         currentDay = DateSerial(yr, mo, i)
//!         Debug.Print "Day " & Day(currentDay) & ": " & Format(currentDay, "dddd")
//!     Next i
//! End Sub
//! ```
//!
//! ### Pay Period Calculations
//!
//! ```vb
//! Function GetPayPeriod(dateValue As Date) As String
//!     ' Bi-monthly pay periods: 1-15 and 16-end
//!     If Day(dateValue) <= 15 Then
//!         GetPayPeriod = "First Half"
//!     Else
//!         GetPayPeriod = "Second Half"
//!     End If
//! End Function
//!
//! Function IsPayday(dateValue As Date) As Boolean
//!     Dim dy As Integer
//!     dy = Day(dateValue)
//!     ' Payday on 15th and last day of month
//!     IsPayday = (dy = 15) Or (Day(dateValue + 1) = 1)
//! End Function
//! ```
//!
//! ### Birthday/Anniversary Checks
//!
//! ```vb
//! Function IsBirthdayMonth(birthday As Date) As Boolean
//!     ' Check if current month/day matches birthday
//!     IsBirthdayMonth = (Month(Date) = Month(birthday)) And _
//!                       (Day(Date) = Day(birthday))
//! End Function
//!
//! Function DaysUntilBirthday(birthday As Date) As Integer
//!     Dim thisYearBirthday As Date
//!     thisYearBirthday = DateSerial(Year(Date), Month(birthday), Day(birthday))
//!     
//!     If thisYearBirthday < Date Then
//!         ' Birthday already passed this year
//!         thisYearBirthday = DateSerial(Year(Date) + 1, Month(birthday), Day(birthday))
//!     End If
//!     
//!     DaysUntilBirthday = DateDiff("d", Date, thisYearBirthday)
//! End Function
//! ```
//!
//! ### Data Grouping by Day
//!
//! ```vb
//! Function GetDayGroup(dateValue As Date) As String
//!     Dim dy As Integer
//!     dy = Day(dateValue)
//!     
//!     Select Case dy
//!         Case 1 To 10
//!             GetDayGroup = "Days 1-10"
//!         Case 11 To 20
//!             GetDayGroup = "Days 11-20"
//!         Case Else
//!             GetDayGroup = "Days 21-31"
//!     End Select
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Calendar Grid Generator
//!
//! ```vb
//! Sub GenerateMonthCalendar(yr As Integer, mo As Integer)
//!     Dim firstDay As Date
//!     Dim lastDay As Date
//!     Dim currentDate As Date
//!     Dim dayOfWeek As Integer
//!     Dim daysInMonth As Integer
//!     
//!     firstDay = DateSerial(yr, mo, 1)
//!     daysInMonth = Day(DateSerial(yr, mo + 1, 0))
//!     
//!     ' Print header
//!     Debug.Print "Sun Mon Tue Wed Thu Fri Sat"
//!     
//!     ' Print leading spaces
//!     dayOfWeek = Weekday(firstDay, vbSunday)
//!     Debug.Print String((dayOfWeek - 1) * 4, " ");
//!     
//!     ' Print days
//!     For i = 1 To daysInMonth
//!         currentDate = DateSerial(yr, mo, i)
//!         Debug.Print Format(Day(currentDate), "000") & " ";
//!         
//!         If Weekday(currentDate, vbSunday) = 7 Then
//!             Debug.Print  ' New line
//!         End If
//!     Next i
//! End Sub
//! ```
//!
//! ### Date Range Iterator
//!
//! ```vb
//! Function GetDateRange(startDate As Date, endDate As Date) As Variant
//!     ' Returns array of dates in range
//!     Dim dayCount As Long
//!     Dim dates() As Date
//!     Dim i As Long
//!     
//!     dayCount = DateDiff("d", startDate, endDate) + 1
//!     ReDim dates(0 To dayCount - 1)
//!     
//!     For i = 0 To dayCount - 1
//!         dates(i) = DateAdd("d", i, startDate)
//!         Debug.Print "Day " & Day(dates(i)) & ": " & dates(i)
//!     Next i
//!     
//!     GetDateRange = dates
//! End Function
//! ```
//!
//! ### Billing Cycle Calculator
//!
//! ```vb
//! Function CalculateBillingDay(startDate As Date, monthsElapsed As Integer) As Date
//!     ' Keep same day-of-month for billing cycle
//!     Dim billingDay As Integer
//!     billingDay = Day(startDate)
//!     
//!     CalculateBillingDay = DateSerial( _
//!         Year(DateAdd("m", monthsElapsed, startDate)), _
//!         Month(DateAdd("m", monthsElapsed, startDate)), _
//!         billingDay)
//! End Function
//!
//! Function AdjustForMonthEnd(targetDate As Date, originalDay As Integer) As Date
//!     ' Handle when original day doesn't exist in target month
//!     Dim targetMonth As Date
//!     Dim maxDayInMonth As Integer
//!     
//!     targetMonth = DateSerial(Year(targetDate), Month(targetDate), 1)
//!     maxDayInMonth = Day(DateSerial(Year(targetDate), Month(targetDate) + 1, 0))
//!     
//!     If originalDay > maxDayInMonth Then
//!         AdjustForMonthEnd = DateSerial(Year(targetDate), Month(targetDate), maxDayInMonth)
//!     Else
//!         AdjustForMonthEnd = DateSerial(Year(targetDate), Month(targetDate), originalDay)
//!     End If
//! End Function
//! ```
//!
//! ### Date Pattern Analyzer
//!
//! ```vb
//! Function AnalyzeDatePattern(dates() As Date) As String
//!     ' Determine if dates follow a pattern
//!     Dim i As Long
//!     Dim allSameDay As Boolean
//!     Dim firstDay As Integer
//!     
//!     If UBound(dates) < LBound(dates) Then
//!         AnalyzeDatePattern = "Empty"
//!         Exit Function
//!     End If
//!     
//!     firstDay = Day(dates(LBound(dates)))
//!     allSameDay = True
//!     
//!     For i = LBound(dates) + 1 To UBound(dates)
//!         If Day(dates(i)) <> firstDay Then
//!             allSameDay = False
//!             Exit For
//!         End If
//!     Next i
//!     
//!     If allSameDay Then
//!         AnalyzeDatePattern = "Monthly on day " & firstDay
//!     Else
//!         AnalyzeDatePattern = "Variable days"
//!     End If
//! End Function
//! ```
//!
//! ### Work Schedule Helper
//!
//! ```vb
//! Function IsWorkDay(dateValue As Date, workDays As String) As Boolean
//!     ' workDays: comma-separated list like "1,15,30"
//!     Dim dayList() As String
//!     Dim i As Integer
//!     Dim currentDay As Integer
//!     
//!     currentDay = Day(dateValue)
//!     dayList = Split(workDays, ",")
//!     
//!     For i = LBound(dayList) To UBound(dayList)
//!         If CInt(Trim(dayList(i))) = currentDay Then
//!             IsWorkDay = True
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     IsWorkDay = False
//! End Function
//! ```
//!
//! ### Leap Year Day Check
//!
//! ```vb
//! Function IsLeapYearDay(dateValue As Date) As Boolean
//!     ' Check if date is February 29
//!     IsLeapYearDay = (Month(dateValue) = 2) And (Day(dateValue) = 29)
//! End Function
//!
//! Function HasLeapYearDayBetween(startDate As Date, endDate As Date) As Boolean
//!     Dim yr As Integer
//!     Dim leapDay As Date
//!     
//!     For yr = Year(startDate) To Year(endDate)
//!         On Error Resume Next
//!         leapDay = DateSerial(yr, 2, 29)
//!         If Err.Number = 0 Then
//!             If leapDay >= startDate And leapDay <= endDate Then
//!                 HasLeapYearDayBetween = True
//!                 Exit Function
//!             End If
//!         End If
//!         Err.Clear
//!     Next yr
//!     
//!     HasLeapYearDayBetween = False
//! End Function
//! ```
//!
//! ### Database Query Builder
//!
//! ```vb
//! Function BuildDayFilter(targetDay As Integer) As String
//!     ' Build SQL WHERE clause for specific day of month
//!     BuildDayFilter = "Day(DateField) = " & targetDay
//! End Function
//!
//! Function GetRecordsForDayRange(startDay As Integer, endDay As Integer) As String
//!     GetRecordsForDayRange = "Day(DateField) BETWEEN " & startDay & " AND " & endDay
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeDay(dateValue As Variant) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     If IsNull(dateValue) Then
//!         SafeDay = Null
//!         Exit Function
//!     End If
//!     
//!     If Not IsDate(dateValue) Then
//!         SafeDay = Null
//!         Exit Function
//!     End If
//!     
//!     SafeDay = Day(dateValue)
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeDay = Null
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 13** (Type mismatch): Argument is not a valid date
//! - **Error 94** (Invalid use of Null): Using Null date without handling
//!
//! ## Performance Considerations
//!
//! - `Day` is a fast, direct component extraction
//! - More efficient than string parsing with `Format`
//! - No performance difference between extracting from Date vs Variant
//! - Cache results when using repeatedly on same date
//! - Combine with `Year` and `Month` for full date decomposition
//!
//! ## Best Practices
//!
//! ### Use for Component Extraction
//!
//! ```vb
//! ' Good - Direct component access
//! dayNum = Day(someDate)
//!
//! ' Avoid - String manipulation overhead
//! dayNum = CInt(Format(someDate, "dd"))
//! ```
//!
//! ### Validate Date Before Extraction
//!
//! ```vb
//! ' Good - Check for valid date
//! If IsDate(userInput) Then
//!     dayNum = Day(CDate(userInput))
//! End If
//!
//! ' Avoid - May cause type mismatch error
//! dayNum = Day(userInput)
//! ```
//!
//! ### Handle Null Values in Variants
//!
//! ```vb
//! ' Good - Check for Null
//! If Not IsNull(dateVariant) Then
//!     dayNum = Day(dateVariant)
//! End If
//!
//! ' Avoid - May propagate Null unexpectedly
//! dayNum = Day(dateVariant)
//! ```
//!
//! ### Combine with `DateSerial` for Date Manipulation
//!
//! ```vb
//! ' Good - Reconstruct date with modified component
//! newDate = DateSerial(Year(oldDate), Month(oldDate), 15)
//!
//! ' Good - Get last day of month
//! lastDay = Day(DateSerial(yr, mo + 1, 0))
//! ```
//!
//! ## Comparison with Other Functions
//!
//! ### `Day` vs `Format`
//!
//! ```vb
//! ' Day - Returns integer, fast
//! dayNum = Day(someDate)  ' Returns 15
//!
//! ' Format - Returns string, slower, more flexible
//! dayStr = Format(someDate, "dd")  ' Returns "15"
//! dayStr = Format(someDate, "d")   ' Returns "15" (no leading zero)
//! ```
//!
//! ### `Day` vs `DatePart`
//!
//! ```vb
//! ' Day - Specific function for day of month
//! dayNum = Day(someDate)
//!
//! ' DatePart - Generic function, can extract any part
//! dayNum = DatePart("d", someDate)  ' Same result
//! ```
//!
//! ### `Day` vs `Weekday`
//!
//! ```vb
//! ' Day - Day of month (1-31)
//! dayOfMonth = Day(#1/15/2025#)  ' Returns 15
//!
//! ' Weekday - Day of week (1-7)
//! dayOfWeek = Weekday(#1/15/2025#)  ' Returns 4 (Wednesday)
//! ```
//!
//! ## Limitations
//!
//! - Only returns day of month, not day of year or week
//! - Returns `Null` for `Null` input (`Variant` propagation)
//! - No built-in validation of day validity for given month
//! - Does not indicate if day is weekend, holiday, etc.
//! - Cannot distinguish between different months with same day
//!
//! ## Related Functions
//!
//! - `Year`: Extracts year component from date
//! - `Month`: Extracts month component from date
//! - `Weekday`: Returns day of week (1-7)
//! - `DatePart`: Generic date part extraction
//! - `DateSerial`: Constructs date from components
//! - `Format`: Formats date as string with custom patterns
//! - `Hour`, `Minute`, `Second`: Extract time components

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn day_basic() {
        let source = r"
d = Day(#1/15/2025#)
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_with_variable() {
        let source = r"
dayNum = Day(birthday)
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_current_date() {
        let source = r"
d = Day(Date)
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_with_dateserial() {
        let source = r"
dayNum = Day(DateSerial(2025, 12, 25))
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_in_function() {
        let source = r"
Function IsLastDayOfMonth(dt As Date) As Boolean
    IsLastDayOfMonth = (Day(dt + 1) = 1)
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_days_in_month() {
        let source = r"
daysInMonth = Day(DateSerial(yr, mo + 1, 0))
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_in_comparison() {
        let source = r#"
If Day(someDate) <= 15 Then
    MsgBox "First half"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_with_year_month() {
        let source = r"
yr = Year(dt)
mo = Month(dt)
dy = Day(dt)
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_equality_check() {
        let source = r"
sameDay = (Day(date1) = Day(date2))
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_in_dateserial_reconstruction() {
        let source = r"
newDate = DateSerial(Year(oldDate), Month(oldDate), Day(oldDate))
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_in_loop() {
        let source = r#"
For i = 1 To dayCount
    Debug.Print "Day " & Day(dates(i))
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_select_case() {
        let source = r#"
Select Case Day(someDate)
    Case 1 To 10
        MsgBox "Early"
    Case 11 To 20
        MsgBox "Mid"
    Case Else
        MsgBox "Late"
End Select
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_with_dateadd() {
        let source = r#"
dayNum = Day(DateAdd("d", 10, Date))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_birthday_check() {
        let source = r"
isToday = (Month(Date) = Month(birthday)) And (Day(Date) = Day(birthday))
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_pay_period() {
        let source = r#"
If Day(payDate) <= 15 Then
    period = "First Half"
Else
    period = "Second Half"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_leap_year_check() {
        let source = r"
isLeapDay = (Month(dt) = 2) And (Day(dt) = 29)
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_with_arithmetic() {
        let source = r"
nextDay = Day(currentDate + 1)
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_array_assignment() {
        let source = r"
days(i) = Day(dates(i))
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_msgbox() {
        let source = r#"
MsgBox "Day: " & Day(someDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_multiple_calls() {
        let source = r"
d1 = Day(date1)
d2 = Day(date2)
diff = d2 - d1
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_with_format() {
        let source = r#"
formatted = Format(DateSerial(2025, 1, Day(someDate)), "mm/dd/yyyy")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_validation() {
        let source = r#"
If Day(startDate) > Day(endDate) Then
    MsgBox "Check dates"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_database_filter() {
        let source = r#"
filter = "Day(DateField) = " & Day(targetDate)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_error_handling() {
        let source = r"
On Error Resume Next
result = Day(userInput)
If Err.Number <> 0 Then
    result = 0
End If
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn day_calendar_generator() {
        let source = r#"
For i = 1 To daysInMonth
    Debug.Print Format(Day(currentDate), "00")
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Day"));
        assert!(debug.contains("Identifier"));
    }
}
