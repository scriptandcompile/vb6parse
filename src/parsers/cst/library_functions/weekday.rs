//! VB6 `Weekday` Function
//!
//! The `Weekday` function returns a Variant (Integer) specifying a whole number representing the day of the week.
//!
//! ## Syntax
//! ```vb6
//! Weekday(date[, firstdayofweek])
//! ```
//!
//! ## Parameters
//! - `date`: Required. Variant, numeric expression, string expression, or any combination that can represent a date. If `date` contains Null, Null is returned.
//! - `firstdayofweek`: Optional. A constant that specifies the first day of the week. If not specified, vbSunday is assumed.
//!
//! ### FirstDayOfWeek Constants
//! - `vbUseSystem` (0): Use the system setting
//! - `vbSunday` (1): Sunday (default)
//! - `vbMonday` (2): Monday
//! - `vbTuesday` (3): Tuesday
//! - `vbWednesday` (4): Wednesday
//! - `vbThursday` (5): Thursday
//! - `vbFriday` (6): Friday
//! - `vbSaturday` (7): Saturday
//!
//! ## Returns
//! Returns a `Variant (Integer)` from 1 to 7 representing the day of the week:
//! - When `firstdayofweek` is `vbSunday` (default): 1=Sunday, 2=Monday, 3=Tuesday, 4=Wednesday, 5=Thursday, 6=Friday, 7=Saturday
//! - When `firstdayofweek` is `vbMonday`: 1=Monday, 2=Tuesday, 3=Wednesday, 4=Thursday, 5=Friday, 6=Saturday, 7=Sunday
//! - The numbering adjusts based on which day is specified as the first day of the week
//!
//! ## Remarks
//! The `Weekday` function determines which day of the week a date falls on:
//!
//! - **Return value range**: Always 1-7, where 1 represents the first day of the week
//! - **Default first day**: Sunday (vbSunday = 1) if not specified
//! - **Null handling**: Returns Null if date contains Null
//! - **Date parsing**: Accepts various date formats (strings, numbers, Date types)
//! - **Zero-based alternative**: Use `(Weekday(date) - 1)` for 0-6 range if needed
//! - **System setting**: vbUseSystem (0) uses Windows regional settings
//! - **International support**: firstdayofweek allows cultural calendar preferences
//! - **Combine with DateSerial**: Calculate specific weekdays
//!
//! ### Understanding Return Values
//! The return value changes based on `firstdayofweek`:
//! - With `vbSunday` (default): Sunday=1, Monday=2, ..., Saturday=7
//! - With `vbMonday`: Monday=1, Tuesday=2, ..., Sunday=7
//! - With `vbTuesday`: Tuesday=1, Wednesday=2, ..., Monday=7
//! - And so on for other starting days
//!
//! ### Common Usage Patterns
//! - **Weekend detection**: `Weekday(date) = vbSaturday Or Weekday(date) = vbSunday`
//! - **Weekday detection**: `Weekday(date) >= vbMonday And Weekday(date) <= vbFriday`
//! - **Specific day check**: `Weekday(date) = vbMonday`
//! - **ISO week (Monday start)**: Use `Weekday(date, vbMonday)`
//!
//! ## Typical Uses
//! 1. **Weekend Detection**: Determine if a date is Saturday or Sunday
//! 2. **Business Day Calculation**: Check if date is a weekday
//! 3. **Schedule Planning**: Plan events for specific days of the week
//! 4. **Report Grouping**: Group data by day of week
//! 5. **Calendar Display**: Format calendars with proper day alignment
//! 6. **Work Week Calculations**: Calculate working days
//! 7. **Day Name Display**: Get day name with WeekdayName function
//! 8. **Date Validation**: Ensure dates fall on specific days
//!
//! ## Basic Examples
//!
//! ### Example 1: Check if Date is Weekend
//! ```vb6
//! Function IsWeekend(checkDate As Date) As Boolean
//!     Dim dayNum As Integer
//!     dayNum = Weekday(checkDate)
//!     IsWeekend = (dayNum = vbSaturday Or dayNum = vbSunday)
//! End Function
//!
//! ' Usage:
//! If IsWeekend(Date) Then
//!     MsgBox "It's the weekend!"
//! End If
//! ```
//!
//! ### Example 2: Get Day Name
//! ```vb6
//! Function GetDayName(checkDate As Date) As String
//!     Select Case Weekday(checkDate)
//!         Case vbSunday
//!             GetDayName = "Sunday"
//!         Case vbMonday
//!             GetDayName = "Monday"
//!         Case vbTuesday
//!             GetDayName = "Tuesday"
//!         Case vbWednesday
//!             GetDayName = "Wednesday"
//!         Case vbThursday
//!             GetDayName = "Thursday"
//!         Case vbFriday
//!             GetDayName = "Friday"
//!         Case vbSaturday
//!             GetDayName = "Saturday"
//!     End Select
//! End Function
//! ```
//!
//! ### Example 3: Find Next Monday
//! ```vb6
//! Function GetNextMonday(fromDate As Date) As Date
//!     Dim daysToAdd As Integer
//!     Dim currentDay As Integer
//!     
//!     currentDay = Weekday(fromDate)
//!     
//!     If currentDay = vbMonday Then
//!         daysToAdd = 7 ' Next Monday is 7 days away
//!     ElseIf currentDay < vbMonday Then
//!         daysToAdd = vbMonday - currentDay
//!     Else
//!         daysToAdd = 7 - (currentDay - vbMonday)
//!     End If
//!     
//!     GetNextMonday = fromDate + daysToAdd
//! End Function
//! ```
//!
//! ### Example 4: Count Weekdays in Month
//! ```vb6
//! Function CountWeekdaysInMonth(year As Integer, month As Integer) As Integer
//!     Dim currentDate As Date
//!     Dim lastDay As Integer
//!     Dim day As Integer
//!     Dim count As Integer
//!     
//!     lastDay = Day(DateSerial(year, month + 1, 0))
//!     count = 0
//!     
//!     For day = 1 To lastDay
//!         currentDate = DateSerial(year, month, day)
//!         If Weekday(currentDate) >= vbMonday And Weekday(currentDate) <= vbFriday Then
//!             count = count + 1
//!         End If
//!     Next day
//!     
//!     CountWeekdaysInMonth = count
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Is Business Day
//! ```vb6
//! Function IsBusinessDay(checkDate As Date) As Boolean
//!     Dim dayNum As Integer
//!     dayNum = Weekday(checkDate)
//!     IsBusinessDay = (dayNum >= vbMonday And dayNum <= vbFriday)
//! End Function
//! ```
//!
//! ### Pattern 2: Get Week Start (Monday)
//! ```vb6
//! Function GetWeekStart(anyDate As Date) As Date
//!     Dim dayNum As Integer
//!     dayNum = Weekday(anyDate, vbMonday)
//!     GetWeekStart = anyDate - (dayNum - 1)
//! End Function
//! ```
//!
//! ### Pattern 3: Get Week End (Sunday)
//! ```vb6
//! Function GetWeekEnd(anyDate As Date) As Date
//!     Dim dayNum As Integer
//!     dayNum = Weekday(anyDate, vbMonday)
//!     GetWeekEnd = anyDate + (7 - dayNum)
//! End Function
//! ```
//!
//! ### Pattern 4: Days Until Specific Weekday
//! ```vb6
//! Function DaysUntilWeekday(fromDate As Date, targetDay As Integer) As Integer
//!     Dim currentDay As Integer
//!     currentDay = Weekday(fromDate)
//!     
//!     If targetDay >= currentDay Then
//!         DaysUntilWeekday = targetDay - currentDay
//!     Else
//!         DaysUntilWeekday = 7 - (currentDay - targetDay)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 5: Is Specific Day
//! ```vb6
//! Function IsMonday(checkDate As Date) As Boolean
//!     IsMonday = (Weekday(checkDate) = vbMonday)
//! End Function
//!
//! Function IsFriday(checkDate As Date) As Boolean
//!     IsFriday = (Weekday(checkDate) = vbFriday)
//! End Function
//! ```
//!
//! ### Pattern 6: Add Business Days
//! ```vb6
//! Function AddBusinessDays(startDate As Date, daysToAdd As Integer) As Date
//!     Dim currentDate As Date
//!     Dim daysAdded As Integer
//!     
//!     currentDate = startDate
//!     daysAdded = 0
//!     
//!     Do While daysAdded < daysToAdd
//!         currentDate = currentDate + 1
//!         If Weekday(currentDate) >= vbMonday And Weekday(currentDate) <= vbFriday Then
//!             daysAdded = daysAdded + 1
//!         End If
//!     Loop
//!     
//!     AddBusinessDays = currentDate
//! End Function
//! ```
//!
//! ### Pattern 7: Group by Day of Week
//! ```vb6
//! Function GetDayOfWeekGroup(checkDate As Date) As String
//!     Select Case Weekday(checkDate)
//!         Case vbMonday, vbWednesday, vbFriday
//!             GetDayOfWeekGroup = "MWF"
//!         Case vbTuesday, vbThursday
//!             GetDayOfWeekGroup = "TTh"
//!         Case vbSaturday, vbSunday
//!             GetDayOfWeekGroup = "Weekend"
//!     End Select
//! End Function
//! ```
//!
//! ### Pattern 8: Calculate Days in Same Week
//! ```vb6
//! Function AreSameWeek(date1 As Date, date2 As Date) As Boolean
//!     Dim week1Start As Date
//!     Dim week2Start As Date
//!     
//!     week1Start = date1 - (Weekday(date1, vbMonday) - 1)
//!     week2Start = date2 - (Weekday(date2, vbMonday) - 1)
//!     
//!     AreSameWeek = (week1Start = week2Start)
//! End Function
//! ```
//!
//! ### Pattern 9: Next Occurrence of Weekday
//! ```vb6
//! Function NextOccurrenceOf(targetDay As Integer, Optional fromDate As Date) As Date
//!     Dim startDate As Date
//!     Dim daysToAdd As Integer
//!     
//!     If fromDate = 0 Then startDate = Date Else startDate = fromDate
//!     
//!     daysToAdd = DaysUntilWeekday(startDate, targetDay)
//!     If daysToAdd = 0 Then daysToAdd = 7 ' Next occurrence
//!     
//!     NextOccurrenceOf = startDate + daysToAdd
//! End Function
//! ```
//!
//! ### Pattern 10: Weekend Days in Range
//! ```vb6
//! Function CountWeekendDays(startDate As Date, endDate As Date) As Integer
//!     Dim currentDate As Date
//!     Dim count As Integer
//!     
//!     count = 0
//!     currentDate = startDate
//!     
//!     Do While currentDate <= endDate
//!         If Weekday(currentDate) = vbSaturday Or Weekday(currentDate) = vbSunday Then
//!             count = count + 1
//!         End If
//!         currentDate = currentDate + 1
//!     Loop
//!     
//!     CountWeekendDays = count
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Business Day Calculator Class
//! ```vb6
//! ' Class: BusinessDayCalculator
//! ' Calculates business days excluding weekends and holidays
//! Option Explicit
//!
//! Private m_Holidays() As Date
//! Private m_HolidayCount As Integer
//!
//! Public Sub Initialize()
//!     m_HolidayCount = 0
//!     ReDim m_Holidays(0 To 9)
//! End Sub
//!
//! Public Sub AddHoliday(holidayDate As Date)
//!     If m_HolidayCount > UBound(m_Holidays) Then
//!         ReDim Preserve m_Holidays(0 To UBound(m_Holidays) * 2)
//!     End If
//!     
//!     m_Holidays(m_HolidayCount) = holidayDate
//!     m_HolidayCount = m_HolidayCount + 1
//! End Sub
//!
//! Public Function IsBusinessDay(checkDate As Date) As Boolean
//!     Dim i As Integer
//!     
//!     ' Check if weekend
//!     If Weekday(checkDate) = vbSaturday Or Weekday(checkDate) = vbSunday Then
//!         IsBusinessDay = False
//!         Exit Function
//!     End If
//!     
//!     ' Check if holiday
//!     For i = 0 To m_HolidayCount - 1
//!         If Int(m_Holidays(i)) = Int(checkDate) Then
//!             IsBusinessDay = False
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     IsBusinessDay = True
//! End Function
//!
//! Public Function CountBusinessDays(startDate As Date, endDate As Date) As Integer
//!     Dim currentDate As Date
//!     Dim count As Integer
//!     
//!     count = 0
//!     currentDate = startDate
//!     
//!     Do While currentDate <= endDate
//!         If IsBusinessDay(currentDate) Then
//!             count = count + 1
//!         End If
//!         currentDate = currentDate + 1
//!     Loop
//!     
//!     CountBusinessDays = count
//! End Function
//!
//! Public Function AddBusinessDays(startDate As Date, daysToAdd As Integer) As Date
//!     Dim currentDate As Date
//!     Dim daysAdded As Integer
//!     
//!     currentDate = startDate
//!     daysAdded = 0
//!     
//!     Do While daysAdded < daysToAdd
//!         currentDate = currentDate + 1
//!         If IsBusinessDay(currentDate) Then
//!             daysAdded = daysAdded + 1
//!         End If
//!     Loop
//!     
//!     AddBusinessDays = currentDate
//! End Function
//! ```
//!
//! ### Example 2: Week Navigator Module
//! ```vb6
//! ' Module: WeekNavigator
//! ' Navigate by weeks with different starting days
//! Option Explicit
//!
//! Public Function GetWeekStart(anyDate As Date, _
//!                             Optional firstDayOfWeek As VbDayOfWeek = vbMonday) As Date
//!     Dim dayNum As Integer
//!     dayNum = Weekday(anyDate, firstDayOfWeek)
//!     GetWeekStart = anyDate - (dayNum - 1)
//! End Function
//!
//! Public Function GetWeekEnd(anyDate As Date, _
//!                           Optional firstDayOfWeek As VbDayOfWeek = vbMonday) As Date
//!     Dim dayNum As Integer
//!     dayNum = Weekday(anyDate, firstDayOfWeek)
//!     GetWeekEnd = anyDate + (7 - dayNum)
//! End Function
//!
//! Public Function GetPreviousWeekStart(anyDate As Date, _
//!                                     Optional firstDayOfWeek As VbDayOfWeek = vbMonday) As Date
//!     GetPreviousWeekStart = GetWeekStart(anyDate - 7, firstDayOfWeek)
//! End Function
//!
//! Public Function GetNextWeekStart(anyDate As Date, _
//!                                 Optional firstDayOfWeek As VbDayOfWeek = vbMonday) As Date
//!     GetNextWeekStart = GetWeekStart(anyDate + 7, firstDayOfWeek)
//! End Function
//!
//! Public Function GetDaysInWeek(weekStartDate As Date, _
//!                              Optional firstDayOfWeek As VbDayOfWeek = vbMonday) As Date()
//!     Dim days(0 To 6) As Date
//!     Dim i As Integer
//!     Dim startDate As Date
//!     
//!     startDate = GetWeekStart(weekStartDate, firstDayOfWeek)
//!     
//!     For i = 0 To 6
//!         days(i) = startDate + i
//!     Next i
//!     
//!     GetDaysInWeek = days
//! End Function
//! ```
//!
//! ### Example 3: Schedule Analyzer Class
//! ```vb6
//! ' Class: ScheduleAnalyzer
//! ' Analyzes date schedules and patterns
//! Option Explicit
//!
//! Public Function GetDayDistribution(dates() As Date) As Variant
//!     Dim distribution(1 To 7) As Integer
//!     Dim i As Long
//!     Dim dayNum As Integer
//!     
//!     For i = LBound(dates) To UBound(dates)
//!         dayNum = Weekday(dates(i))
//!         distribution(dayNum) = distribution(dayNum) + 1
//!     Next i
//!     
//!     GetDayDistribution = distribution
//! End Function
//!
//! Public Function GetMostCommonDay(dates() As Date) As Integer
//!     Dim distribution As Variant
//!     Dim maxCount As Integer
//!     Dim mostCommonDay As Integer
//!     Dim i As Integer
//!     
//!     distribution = GetDayDistribution(dates)
//!     maxCount = 0
//!     mostCommonDay = vbSunday
//!     
//!     For i = 1 To 7
//!         If distribution(i) > maxCount Then
//!             maxCount = distribution(i)
//!             mostCommonDay = i
//!         End If
//!     Next i
//!     
//!     GetMostCommonDay = mostCommonDay
//! End Function
//!
//! Public Function FilterByWeekday(dates() As Date, targetDay As Integer) As Date()
//!     Dim result() As Date
//!     Dim count As Long
//!     Dim i As Long
//!     
//!     count = 0
//!     For i = LBound(dates) To UBound(dates)
//!         If Weekday(dates(i)) = targetDay Then
//!             ReDim Preserve result(0 To count)
//!             result(count) = dates(i)
//!             count = count + 1
//!         End If
//!     Next i
//!     
//!     FilterByWeekday = result
//! End Function
//!
//! Public Function IsRecurringPattern(dates() As Date, dayOfWeek As Integer) As Boolean
//!     Dim i As Long
//!     
//!     IsRecurringPattern = True
//!     For i = LBound(dates) To UBound(dates)
//!         If Weekday(dates(i)) <> dayOfWeek Then
//!             IsRecurringPattern = False
//!             Exit Function
//!         End If
//!     Next i
//! End Function
//! ```
//!
//! ### Example 4: Calendar Generator Module
//! ```vb6
//! ' Module: CalendarGenerator
//! ' Generates calendar grids and displays
//! Option Explicit
//!
//! Public Function GenerateMonthCalendar(year As Integer, month As Integer, _
//!                                      Optional firstDayOfWeek As VbDayOfWeek = vbSunday) As Variant
//!     Dim calendar(0 To 5, 0 To 6) As Variant
//!     Dim firstDay As Date
//!     Dim lastDay As Date
//!     Dim currentDate As Date
//!     Dim row As Integer
//!     Dim col As Integer
//!     Dim dayNum As Integer
//!     
//!     firstDay = DateSerial(year, month, 1)
//!     lastDay = DateSerial(year, month + 1, 0)
//!     
//!     ' Initialize calendar with empty values
//!     For row = 0 To 5
//!         For col = 0 To 6
//!             calendar(row, col) = ""
//!         Next col
//!     Next row
//!     
//!     ' Fill in the days
//!     currentDate = firstDay
//!     Do While currentDate <= lastDay
//!         dayNum = Weekday(currentDate, firstDayOfWeek) - 1
//!         row = (Day(currentDate) + Weekday(firstDay, firstDayOfWeek) - 2) \ 7
//!         col = dayNum
//!         calendar(row, col) = Day(currentDate)
//!         currentDate = currentDate + 1
//!     Loop
//!     
//!     GenerateMonthCalendar = calendar
//! End Function
//!
//! Public Function GetWeekNumbers(year As Integer, month As Integer) As Integer()
//!     Dim firstDay As Date
//!     Dim lastDay As Date
//!     Dim currentDate As Date
//!     Dim weekNums() As Integer
//!     Dim count As Integer
//!     
//!     firstDay = DateSerial(year, month, 1)
//!     lastDay = DateSerial(year, month + 1, 0)
//!     
//!     count = 0
//!     currentDate = firstDay
//!     
//!     Do While currentDate <= lastDay
//!         If Weekday(currentDate, vbMonday) = 1 Then
//!             ReDim Preserve weekNums(0 To count)
//!             weekNums(count) = DatePart("ww", currentDate, vbMonday, vbFirstFourDays)
//!             count = count + 1
//!         End If
//!         currentDate = currentDate + 1
//!     Loop
//!     
//!     GetWeekNumbers = weekNums
//! End Function
//! ```
//!
//! ## Error Handling
//! The `Weekday` function can raise the following errors:
//!
//! - **Error 13 (Type mismatch)**: If date argument cannot be interpreted as a date
//! - **Error 5 (Invalid procedure call or argument)**: If firstdayofweek is not between 0 and 7
//! - **Error 6 (Overflow)**: If date value is too large or too small
//!
//! ## Performance Notes
//! - Very fast operation - direct calculation from date value
//! - No performance difference between firstdayofweek values
//! - Constant time O(1) operation
//! - Can be called repeatedly without performance concerns
//! - Consider caching result if used multiple times with same date
//!
//! ## Best Practices
//! 1. **Use named constants** (vbMonday, vbSunday, etc.) instead of numeric values for clarity
//! 2. **Specify firstdayofweek** explicitly when week start matters (especially for ISO weeks)
//! 3. **Cache results** when calling repeatedly with same date
//! 4. **Validate date input** before calling Weekday to avoid errors
//! 5. **Use WeekdayName** function to get localized day names
//! 6. **Document assumptions** about which day starts the week in your code
//! 7. **Consider time zones** when working with date/time values
//! 8. **Test boundary cases** (month/year boundaries, leap years)
//! 9. **Use with DateAdd** for complex date arithmetic
//! 10. **Handle Null values** explicitly when working with Variant dates
//!
//! ## Comparison Table
//!
//! | Function | Purpose | Returns | Notes |
//! |----------|---------|---------|-------|
//! | `Weekday` | Get day of week number | Integer (1-7) | Configurable week start |
//! | `WeekdayName` | Get day name | String | Localized, abbreviated option |
//! | `Day` | Get day of month | Integer (1-31) | Calendar day number |
//! | `DatePart` | Get date part | Variant | More general, includes "w" for weekday |
//!
//! ## Platform Notes
//! - Available in VB6, VBA, and VBScript
//! - Behavior consistent across platforms
//! - Returns Integer (can be stored in Integer or Long)
//! - System locale affects vbUseSystem behavior
//! - Date range: January 1, 100 to December 31, 9999
//!
//! ## Limitations
//! - Cannot directly return day name (use WeekdayName for that)
//! - Return value always 1-7 (no zero-based option)
//! - Cannot specify custom week numbering systems
//! - Does not account for holidays or business days
//! - No built-in ISO 8601 week numbering (use DatePart("ww", ...) for that)
//! - Cannot calculate week of year (use DatePart for that)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_weekday_basic() {
        let source = r#"
Sub Test()
    dayNum = Weekday(Date)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_variable_assignment() {
        let source = r#"
Sub Test()
    Dim day As Integer
    day = Weekday(checkDate)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
        assert!(debug.contains("checkDate"));
    }

    #[test]
    fn test_weekday_with_firstdayofweek() {
        let source = r#"
Sub Test()
    dayNum = Weekday(myDate, vbMonday)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_if_statement() {
        let source = r#"
Sub Test()
    If Weekday(Date) = vbSaturday Then
        MsgBox "It's Saturday!"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_select_case() {
        let source = r#"
Sub Test()
    Select Case Weekday(checkDate)
        Case vbMonday
            DoMonday
        Case vbFriday
            DoFriday
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_function_return() {
        let source = r#"
Function IsWeekend(checkDate As Date) As Boolean
    IsWeekend = (Weekday(checkDate) = vbSaturday Or Weekday(checkDate) = vbSunday)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_for_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 31
        If Weekday(DateSerial(2024, 1, i)) = vbMonday Then
            count = count + 1
        End If
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_comparison() {
        let source = r#"
Sub Test()
    If Weekday(date1) = Weekday(date2) Then
        SameDay
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Day of week: " & Weekday(Date)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_function_argument() {
        let source = r#"
Sub Test()
    Call ProcessDay(Weekday(currentDate))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "Weekday: " & Weekday(Date)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_do_while() {
        let source = r#"
Sub Test()
    Do While Weekday(currentDate) <> vbMonday
        currentDate = currentDate + 1
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_do_until() {
        let source = r#"
Sub Test()
    Do Until Weekday(testDate) = vbFriday
        testDate = testDate + 1
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_while_wend() {
        let source = r#"
Sub Test()
    While Weekday(dt) >= vbMonday And Weekday(dt) <= vbFriday
        dt = dt + 1
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_iif() {
        let source = r#"
Sub Test()
    category = IIf(Weekday(dt) = vbSaturday Or Weekday(dt) = vbSunday, "Weekend", "Weekday")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_with_statement() {
        let source = r#"
Sub Test()
    With dateInfo
        .DayNumber = Weekday(.TheDate)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_parentheses() {
        let source = r#"
Sub Test()
    result = (Weekday(myDate) - 1)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    dayNum = Weekday(userDate)
    If Err.Number <> 0 Then
        dayNum = 0
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_property_assignment() {
        let source = r#"
Sub Test()
    obj.WeekdayNumber = Weekday(obj.EventDate)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_arithmetic() {
        let source = r#"
Sub Test()
    daysUntilMonday = vbMonday - Weekday(currentDate)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_array_assignment() {
        let source = r#"
Sub Test()
    weekdays(i) = Weekday(dates(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_print_statement() {
        let source = r#"
Sub Test()
    Print #1, Weekday(reportDate)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_class_usage() {
        let source = r#"
Sub Test()
    Set calendar = New CalendarControl
    calendar.StartDay = Weekday(calendar.FirstDate, vbMonday)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_elseif() {
        let source = r#"
Sub Test()
    If x = 1 Then
        y = 1
    ElseIf Weekday(dt) = vbMonday Then
        y = 2
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_range_check() {
        let source = r#"
Sub Test()
    If Weekday(dt) >= vbMonday And Weekday(dt) <= vbFriday Then
        IsWeekday = True
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_or_condition() {
        let source = r#"
Sub Test()
    If Weekday(dt) = vbSaturday Or Weekday(dt) = vbSunday Then
        IsWeekend = True
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_switch() {
        let source = r#"
Sub Test()
    category = Switch(Weekday(dt) = vbSunday, "Rest", Weekday(dt) = vbMonday, "Start", True, "Other")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }

    #[test]
    fn test_weekday_dateadd() {
        let source = r#"
Sub Test()
    nextWeek = DateAdd("d", 7 - Weekday(dt, vbMonday), dt)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Weekday"));
    }
}
