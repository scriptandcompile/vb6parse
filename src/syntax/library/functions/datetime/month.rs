//! # Month Function
//!
//! Returns a Variant (Integer) specifying a whole number between 1 and 12, inclusive,
//! representing the month of the year.
//!
//! ## Syntax
//!
//! ```vb
//! Month(date)
//! ```
//!
//! ## Parameters
//!
//! - **date** (Required) - Any Variant, numeric expression, string expression, or any combination
//!   that can represent a date. If date contains Null, Null is returned.
//!
//! ## Return Value
//!
//! Returns a **Variant (Integer)** from 1 to 12 representing the month of the year:
//! - 1 = January
//! - 2 = February
//! - 3 = March
//! - 4 = April
//! - 5 = May
//! - 6 = June
//! - 7 = July
//! - 8 = August
//! - 9 = September
//! - 10 = October
//! - 11 = November
//! - 12 = December
//!
//! ## Remarks
//!
//! The Month function extracts the month component from a date value. It is commonly used
//! for date calculations, filtering data by month, generating reports, and fiscal year processing.
//!
//! ### Key Characteristics:
//! - Returns an Integer from 1 (January) to 12 (December)
//! - Works with Date literals, Date variables, and string expressions that can be converted to dates
//! - If the date parameter contains Null, the function returns Null
//! - The time portion of the date value is ignored; only the date portion is used
//! - Can be used with Now, Date, or any valid date expression
//! - Type mismatch error (Error 13) occurs if the argument cannot be interpreted as a date
//! - Independent of the day and year components of the date
//!
//! ### Common Use Cases:
//! - Extract month from date for display or calculation
//! - Filter records by month
//! - Group data by month for reporting
//! - Calculate fiscal periods (quarters, fiscal years)
//! - Determine season based on month
//! - Validate date ranges
//! - Calculate months between dates
//! - Generate month-based file names or identifiers
//!
//! ## Typical Uses
//!
//! 1. **Extract Month from Date** - Get the month number from any date value
//! 2. **Filter by Month** - Select records from a specific month
//! 3. **Month-based Reports** - Generate monthly summaries and reports
//! 4. **Fiscal Year Calculations** - Determine fiscal quarters and periods
//! 5. **Season Determination** - Map month to season (Winter, Spring, Summer, Fall)
//! 6. **Date Validation** - Verify dates fall within expected month ranges
//! 7. **Monthly Aggregation** - Group transactions or data by month
//! 8. **File Organization** - Create month-based folder or file naming schemes
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Get current month
//! Dim currentMonth As Integer
//! currentMonth = Month(Now)
//! ' If today is November 23, 2025, returns 11
//! ```
//!
//! ```vb
//! ' Example 2: Extract month from date literal
//! Dim birthMonth As Integer
//! birthMonth = Month(#3/15/1990#)
//! ' Returns 3 (March)
//! ```
//!
//! ```vb
//! ' Example 3: Filter records by month
//! Dim orderDate As Date
//! Dim orderMonth As Integer
//!
//! orderDate = rs("OrderDate")
//! orderMonth = Month(orderDate)
//!
//! If orderMonth = 12 Then
//!     Debug.Print "December order"
//! End If
//! ```
//!
//! ```vb
//! ' Example 4: Determine fiscal quarter
//! Dim transactionDate As Date
//! Dim fiscalQuarter As Integer
//!
//! transactionDate = #5/15/2025#
//!
//! Select Case Month(transactionDate)
//!     Case 1, 2, 3
//!         fiscalQuarter = 1
//!     Case 4, 5, 6
//!         fiscalQuarter = 2
//!     Case 7, 8, 9
//!         fiscalQuarter = 3
//!     Case 10, 11, 12
//!         fiscalQuarter = 4
//! End Select
//!
//! Debug.Print "Q" & fiscalQuarter
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Safe month extraction with Null handling
//! Function SafeMonth(dateValue As Variant) As Variant
//!     If IsNull(dateValue) Then
//!         SafeMonth = Null
//!     ElseIf Not IsDate(dateValue) Then
//!         SafeMonth = Null
//!     Else
//!         SafeMonth = Month(dateValue)
//!     End If
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 2: Get month name from number
//! Function GetMonthName(dateValue As Date) As String
//!     Dim monthNames(1 To 12) As String
//!     monthNames(1) = "January"
//!     monthNames(2) = "February"
//!     monthNames(3) = "March"
//!     monthNames(4) = "April"
//!     monthNames(5) = "May"
//!     monthNames(6) = "June"
//!     monthNames(7) = "July"
//!     monthNames(8) = "August"
//!     monthNames(9) = "September"
//!     monthNames(10) = "October"
//!     monthNames(11) = "November"
//!     monthNames(12) = "December"
//!     
//!     GetMonthName = monthNames(Month(dateValue))
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 3: Check if date is in current month
//! Function IsCurrentMonth(dateValue As Date) As Boolean
//!     IsCurrentMonth = (Month(dateValue) = Month(Date))
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 4: Get fiscal quarter (Oct-Dec = Q1)
//! Function GetFiscalQuarter(dateValue As Date, fiscalStartMonth As Integer) As Integer
//!     Dim monthNum As Integer
//!     Dim adjustedMonth As Integer
//!     
//!     monthNum = Month(dateValue)
//!     adjustedMonth = monthNum - fiscalStartMonth + 1
//!     
//!     If adjustedMonth <= 0 Then
//!         adjustedMonth = adjustedMonth + 12
//!     End If
//!     
//!     GetFiscalQuarter = ((adjustedMonth - 1) \ 3) + 1
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 5: Calculate months between two dates
//! Function MonthsBetween(startDate As Date, endDate As Date) As Integer
//!     Dim years As Integer
//!     Dim months As Integer
//!     
//!     years = Year(endDate) - Year(startDate)
//!     months = Month(endDate) - Month(startDate)
//!     
//!     MonthsBetween = (years * 12) + months
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 6: Get season from month
//! Function GetSeason(dateValue As Date) As String
//!     Select Case Month(dateValue)
//!         Case 12, 1, 2
//!             GetSeason = "Winter"
//!         Case 3, 4, 5
//!             GetSeason = "Spring"
//!         Case 6, 7, 8
//!             GetSeason = "Summer"
//!         Case 9, 10, 11
//!             GetSeason = "Fall"
//!     End Select
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 7: Format month with leading zero
//! Function FormatMonth(dateValue As Date) As String
//!     FormatMonth = Format(Month(dateValue), "00")
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 8: Check if same month and year
//! Function IsSameMonthYear(date1 As Date, date2 As Date) As Boolean
//!     IsSameMonthYear = (Month(date1) = Month(date2)) And _
//!                       (Year(date1) = Year(date2))
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 9: Get first day of month
//! Function GetFirstDayOfMonth(dateValue As Date) As Date
//!     GetFirstDayOfMonth = DateSerial(Year(dateValue), Month(dateValue), 1)
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 10: Get last day of month
//! Function GetLastDayOfMonth(dateValue As Date) As Date
//!     Dim nextMonth As Date
//!     nextMonth = DateSerial(Year(dateValue), Month(dateValue) + 1, 1)
//!     GetLastDayOfMonth = DateAdd("d", -1, nextMonth)
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Monthly Sales Report Generator
//!
//! ```vb
//! ' Class: MonthlySalesReport
//! ' Generates comprehensive monthly sales analysis
//!
//! Option Explicit
//!
//! Private m_year As Integer
//! Private m_month As Integer
//! Private m_sales() As Double
//! Private m_recordCount As Long
//!
//! Public Sub Initialize(targetYear As Integer, targetMonth As Integer)
//!     m_year = targetYear
//!     m_month = targetMonth
//!     m_recordCount = 0
//!     ReDim m_sales(0)
//! End Sub
//!
//! Public Sub ProcessRecordset(rs As ADODB.Recordset)
//!     Dim saleDate As Date
//!     Dim saleAmount As Double
//!     
//!     rs.MoveFirst
//!     Do While Not rs.EOF
//!         saleDate = rs("SaleDate")
//!         
//!         If Year(saleDate) = m_year And Month(saleDate) = m_month Then
//!             saleAmount = rs("Amount")
//!             
//!             ReDim Preserve m_sales(m_recordCount)
//!             m_sales(m_recordCount) = saleAmount
//!             m_recordCount = m_recordCount + 1
//!         End If
//!         
//!         rs.MoveNext
//!     Loop
//! End Sub
//!
//! Public Function GetTotalSales() As Double
//!     Dim i As Long
//!     Dim total As Double
//!     
//!     total = 0
//!     For i = 0 To m_recordCount - 1
//!         total = total + m_sales(i)
//!     Next i
//!     
//!     GetTotalSales = total
//! End Function
//!
//! Public Function GetAverageSale() As Double
//!     If m_recordCount = 0 Then
//!         GetAverageSale = 0
//!     Else
//!         GetAverageSale = GetTotalSales() / m_recordCount
//!     End If
//! End Function
//!
//! Public Function GenerateReport() As String
//!     Dim report As String
//!     Dim monthName As String
//!     
//!     monthName = Format(DateSerial(m_year, m_month, 1), "mmmm")
//!     
//!     report = "Sales Report - " & monthName & " " & m_year & vbCrLf
//!     report = report & String(50, "-") & vbCrLf
//!     report = report & "Total Sales: " & Format(GetTotalSales(), "$#,##0.00") & vbCrLf
//!     report = report & "Number of Transactions: " & m_recordCount & vbCrLf
//!     report = report & "Average Sale: " & Format(GetAverageSale(), "$#,##0.00") & vbCrLf
//!     
//!     GenerateReport = report
//! End Function
//! ```
//!
//! ### Example 2: Fiscal Calendar Manager
//!
//! ```vb
//! ' Class: FiscalCalendar
//! ' Manages fiscal year calculations with custom start month
//!
//! Option Explicit
//!
//! Private m_fiscalStartMonth As Integer  ' 1-12
//!
//! Public Sub Initialize(fiscalStartMonth As Integer)
//!     If fiscalStartMonth < 1 Or fiscalStartMonth > 12 Then
//!         Err.Raise 5, "FiscalCalendar", "Invalid fiscal start month"
//!     End If
//!     m_fiscalStartMonth = fiscalStartMonth
//! End Sub
//!
//! Public Function GetFiscalYear(calendarDate As Date) As Integer
//!     Dim calendarYear As Integer
//!     Dim calendarMonth As Integer
//!     
//!     calendarYear = Year(calendarDate)
//!     calendarMonth = Month(calendarDate)
//!     
//!     If calendarMonth >= m_fiscalStartMonth Then
//!         GetFiscalYear = calendarYear
//!     Else
//!         GetFiscalYear = calendarYear - 1
//!     End If
//! End Function
//!
//! Public Function GetFiscalQuarter(calendarDate As Date) As Integer
//!     Dim monthNum As Integer
//!     Dim monthsFromStart As Integer
//!     
//!     monthNum = Month(calendarDate)
//!     monthsFromStart = monthNum - m_fiscalStartMonth
//!     
//!     If monthsFromStart < 0 Then
//!         monthsFromStart = monthsFromStart + 12
//!     End If
//!     
//!     GetFiscalQuarter = (monthsFromStart \ 3) + 1
//! End Function
//!
//! Public Function GetFiscalPeriod(calendarDate As Date) As Integer
//!     ' Returns 1-12 for each month of fiscal year
//!     Dim monthNum As Integer
//!     Dim period As Integer
//!     
//!     monthNum = Month(calendarDate)
//!     period = monthNum - m_fiscalStartMonth + 1
//!     
//!     If period <= 0 Then
//!         period = period + 12
//!     End If
//!     
//!     GetFiscalPeriod = period
//! End Function
//!
//! Public Function GetQuarterStartDate(calendarDate As Date) As Date
//!     Dim fiscalYear As Integer
//!     Dim quarter As Integer
//!     Dim quarterStartMonth As Integer
//!     
//!     fiscalYear = GetFiscalYear(calendarDate)
//!     quarter = GetFiscalQuarter(calendarDate)
//!     
//!     quarterStartMonth = m_fiscalStartMonth + ((quarter - 1) * 3)
//!     If quarterStartMonth > 12 Then
//!         quarterStartMonth = quarterStartMonth - 12
//!         fiscalYear = fiscalYear + 1
//!     End If
//!     
//!     GetQuarterStartDate = DateSerial(fiscalYear, quarterStartMonth, 1)
//! End Function
//!
//! Public Function FormatFiscalPeriod(calendarDate As Date) As String
//!     Dim fy As Integer
//!     Dim period As Integer
//!     
//!     fy = GetFiscalYear(calendarDate)
//!     period = GetFiscalPeriod(calendarDate)
//!     
//!     FormatFiscalPeriod = "FY" & fy & "-P" & Format(period, "00")
//! End Function
//! ```
//!
//! ### Example 3: Date Range Validator
//!
//! ```vb
//! ' Module: DateRangeValidator
//! ' Validates dates against month-based constraints
//!
//! Option Explicit
//!
//! Public Function IsInMonthRange(testDate As Date, _
//!                                startMonth As Integer, _
//!                                endMonth As Integer) As Boolean
//!     Dim testMonth As Integer
//!     testMonth = Month(testDate)
//!     
//!     If startMonth <= endMonth Then
//!         ' Same year range (e.g., March to September)
//!         IsInMonthRange = (testMonth >= startMonth) And (testMonth <= endMonth)
//!     Else
//!         ' Wraps year boundary (e.g., November to February)
//!         IsInMonthRange = (testMonth >= startMonth) Or (testMonth <= endMonth)
//!     End If
//! End Function
//!
//! Public Function IsValidBusinessCycle(startDate As Date, endDate As Date) As Boolean
//!     Dim monthsDiff As Integer
//!     
//!     monthsDiff = ((Year(endDate) - Year(startDate)) * 12) + _
//!                  (Month(endDate) - Month(startDate))
//!     
//!     ' Business cycle should be at least 3 months, max 18 months
//!     IsValidBusinessCycle = (monthsDiff >= 3) And (monthsDiff <= 18)
//! End Function
//!
//! Public Function GetMonthsOverlap(start1 As Date, end1 As Date, _
//!                                  start2 As Date, end2 As Date) As Integer
//!     Dim overlapStart As Date
//!     Dim overlapEnd As Date
//!     Dim monthsOverlap As Integer
//!     
//!     ' Determine overlap period
//!     If start1 > start2 Then
//!         overlapStart = start1
//!     Else
//!         overlapStart = start2
//!     End If
//!     
//!     If end1 < end2 Then
//!         overlapEnd = end1
//!     Else
//!         overlapEnd = end2
//!     End If
//!     
//!     If overlapStart > overlapEnd Then
//!         GetMonthsOverlap = 0
//!     Else
//!         monthsOverlap = ((Year(overlapEnd) - Year(overlapStart)) * 12) + _
//!                        (Month(overlapEnd) - Month(overlapStart)) + 1
//!         GetMonthsOverlap = monthsOverlap
//!     End If
//! End Function
//!
//! Public Function IsQuarterEnd(testDate As Date) As Boolean
//!     Dim monthNum As Integer
//!     Dim lastDay As Date
//!     
//!     monthNum = Month(testDate)
//!     
//!     ' Check if month is quarter-end month
//!     If monthNum Mod 3 <> 0 Then
//!         IsQuarterEnd = False
//!         Exit Function
//!     End If
//!     
//!     ' Check if it's the last day of the month
//!     lastDay = DateSerial(Year(testDate), monthNum + 1, 1) - 1
//!     IsQuarterEnd = (Day(testDate) = Day(lastDay))
//! End Function
//!
//! Public Function GetNextQuarterStart(fromDate As Date) As Date
//!     Dim currentMonth As Integer
//!     Dim nextQuarterMonth As Integer
//!     Dim targetYear As Integer
//!     
//!     currentMonth = Month(fromDate)
//!     targetYear = Year(fromDate)
//!     
//!     ' Calculate next quarter start month
//!     nextQuarterMonth = ((currentMonth - 1) \ 3 + 1) * 3 + 1
//!     
//!     If nextQuarterMonth > 12 Then
//!         nextQuarterMonth = nextQuarterMonth - 12
//!         targetYear = targetYear + 1
//!     End If
//!     
//!     GetNextQuarterStart = DateSerial(targetYear, nextQuarterMonth, 1)
//! End Function
//! ```
//!
//! ### Example 4: Subscription Manager
//!
//! ```vb
//! ' Class: SubscriptionManager
//! ' Tracks monthly recurring subscriptions
//!
//! Option Explicit
//!
//! Private Type Subscription
//!     CustomerID As String
//!     StartDate As Date
//!     EndDate As Date
//!     MonthlyFee As Double
//!     Active As Boolean
//! End Type
//!
//! Private m_subscriptions() As Subscription
//! Private m_count As Long
//!
//! Public Sub AddSubscription(customerID As String, startDate As Date, _
//!                            endDate As Date, monthlyFee As Double)
//!     ReDim Preserve m_subscriptions(m_count)
//!     
//!     m_subscriptions(m_count).CustomerID = customerID
//!     m_subscriptions(m_count).StartDate = startDate
//!     m_subscriptions(m_count).EndDate = endDate
//!     m_subscriptions(m_count).MonthlyFee = monthlyFee
//!     m_subscriptions(m_count).Active = True
//!     
//!     m_count = m_count + 1
//! End Sub
//!
//! Public Function GetMonthlyRevenue(targetYear As Integer, _
//!                                   targetMonth As Integer) As Double
//!     Dim i As Long
//!     Dim revenue As Double
//!     Dim checkDate As Date
//!     
//!     checkDate = DateSerial(targetYear, targetMonth, 15) ' Mid-month
//!     revenue = 0
//!     
//!     For i = 0 To m_count - 1
//!         If m_subscriptions(i).Active Then
//!             If IsSubscriptionActive(m_subscriptions(i), checkDate) Then
//!                 revenue = revenue + m_subscriptions(i).MonthlyFee
//!             End If
//!         End If
//!     Next i
//!     
//!     GetMonthlyRevenue = revenue
//! End Function
//!
//! Private Function IsSubscriptionActive(sub As Subscription, _
//!                                       checkDate As Date) As Boolean
//!     IsSubscriptionActive = (checkDate >= sub.StartDate) And _
//!                           (checkDate <= sub.EndDate)
//! End Function
//!
//! Public Function GetActiveSubscriptions(targetYear As Integer, _
//!                                        targetMonth As Integer) As Long
//!     Dim i As Long
//!     Dim count As Long
//!     Dim checkDate As Date
//!     
//!     checkDate = DateSerial(targetYear, targetMonth, 15)
//!     count = 0
//!     
//!     For i = 0 To m_count - 1
//!         If m_subscriptions(i).Active Then
//!             If IsSubscriptionActive(m_subscriptions(i), checkDate) Then
//!                 count = count + 1
//!             End If
//!         End If
//!     Next i
//!     
//!     GetActiveSubscriptions = count
//! End Function
//!
//! Public Function CalculateLifetimeMonths(customerID As String) As Integer
//!     Dim i As Long
//!     Dim totalMonths As Integer
//!     Dim monthsActive As Integer
//!     
//!     totalMonths = 0
//!     
//!     For i = 0 To m_count - 1
//!         If m_subscriptions(i).CustomerID = customerID Then
//!             monthsActive = ((Year(m_subscriptions(i).EndDate) - _
//!                            Year(m_subscriptions(i).StartDate)) * 12) + _
//!                           (Month(m_subscriptions(i).EndDate) - _
//!                            Month(m_subscriptions(i).StartDate)) + 1
//!             totalMonths = totalMonths + monthsActive
//!         End If
//!     Next i
//!     
//!     CalculateLifetimeMonths = totalMonths
//! End Function
//!
//! Public Sub GenerateAnnualReport(targetYear As Integer)
//!     Dim monthNum As Integer
//!     Dim revenue As Double
//!     
//!     Debug.Print "Annual Subscription Report - " & targetYear
//!     Debug.Print String(60, "-")
//!     
//!     For monthNum = 1 To 12
//!         revenue = GetMonthlyRevenue(targetYear, monthNum)
//!         Debug.Print Format(DateSerial(targetYear, monthNum, 1), "mmmm") & ": " & _
//!                    Format(revenue, "$#,##0.00") & " (" & _
//!                    GetActiveSubscriptions(targetYear, monthNum) & " subs)"
//!     Next monthNum
//! End Sub
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! On Error Resume Next
//! monthValue = Month(dateInput)
//! If Err.Number = 13 Then
//!     MsgBox "Invalid date format"
//! ElseIf Err.Number <> 0 Then
//!     MsgBox "Error extracting month: " & Err.Description
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Performance Considerations
//!
//! - `Month()` is a fast operation - extracting a component from an internal date value
//! - For repeated calls with the same date, consider caching the result
//! - When filtering large recordsets by month, use database queries when possible
//! - `Month()` can be called millions of times without significant performance impact
//! - Combining with `Year()` and `Day()` is efficient for complete date decomposition
//!
//! ## Best Practices
//!
//! 1. **Validate input** - Use `IsDate()` to check if the value can be converted to a date before calling `Month()`
//! 2. **Handle Null** - Check for Null values when working with Variant date parameters
//! 3. **Use with Format** - Combine `Month()` with `Format()` for display purposes (leading zeros, month names)
//! 4. **Combine with Year** - When filtering or comparing, consider both `Month()` and `Year()` for accuracy
//! 5. **Fiscal year awareness** - Remember that fiscal years may not align with calendar months
//! 6. **Use constants** - Define constants for month numbers to improve code readability
//! 7. **Document assumptions** - Clearly state whether code expects calendar or fiscal months
//! 8. **Consider `MonthName()`** - Use `MonthName()` function for getting month names instead of arrays
//! 9. **Date arithmetic** - Use `DateSerial()` with `Month()` for date calculations
//! 10. **Test edge cases** - Test with leap years, year boundaries, and Null values
//!
//! ## Comparison with Other Date Functions
//!
//! | Function | Returns | Range | Use Case |
//! |----------|---------|-------|----------|
//! | **Month** | Month number | 1-12 | Extract month component |
//! | **Day** | Day of month | 1-31 | Extract day component |
//! | **Year** | Year | e.g., 2025 | Extract year component |
//! | **Weekday** | Day of week | 1-7 | Determine day of week |
//! | **`DatePart`** | Any date component | Varies | General date part extraction |
//! | **`MonthName`** | Month name | String | Get month name (not number) |
//!
//! ## Platform Notes
//!
//! - Available in VBA (Excel, Access, Word, etc.)
//! - Available in VB6
//! - Available in `VBScript`
//! - Returns Integer type (not Long)
//! - Consistent across all VB6/VBA platforms
//! - Uses system's regional settings for date interpretation (when parsing strings)
//!
//! ## Limitations
//!
//! - Cannot extract month from time-only values (use Date data type)
//! - Returns Null if input is Null (not an error)
//! - Type mismatch error if input cannot be interpreted as a date
//! - Always returns 1-12, no support for zero-based indexing
//! - Does not provide month name (use `MonthName()` or `Format()` for that)
//! - Time component is ignored (only date portion matters)
//!
//! ## Related Functions
//!
//! - **Year** - Returns the year component of a date (1-9999)
//! - **Day** - Returns the day component of a date (1-31)
//! - **Weekday** - Returns the day of the week (1-7)
//! - **Hour** - Returns the hour component of a time (0-23)
//! - **Minute** - Returns the minute component of a time (0-59)
//! - **Second** - Returns the second component of a time (0-59)
//! - **`DatePart`** - Returns a specified part of a date (flexible)
//! - **`DateSerial`** - Creates a date from year, month, and day components
//! - **`MonthName`** - Returns the name of the month as a string
//! - **Format** - Formats a date with custom patterns including month
//!
//! ## VB6 Parser Notes
//!
//! Month is parsed as a regular function call (`CallExpression`). This module exists primarily
//! for documentation purposes to provide comprehensive reference material for VB6 developers
//! working with date calculations and month extraction operations.

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn month_basic() {
        let source = r"
Dim m As Integer
m = Month(Now)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_variable_assignment() {
        let source = r"
Dim currentMonth As Integer
currentMonth = Month(Date)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_date_literal() {
        let source = r"
Dim m As Integer
m = Month(#3/15/2025#)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_if_statement() {
        let source = r#"
If Month(orderDate) = 12 Then
    MsgBox "December order"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_function_return() {
        let source = r"
Function GetCurrentMonth() As Integer
    GetCurrentMonth = Month(Date)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_select_case() {
        let source = r"
Select Case Month(transactionDate)
    Case 1, 2, 3
        quarter = 1
    Case 4, 5, 6
        quarter = 2
    Case 7, 8, 9
        quarter = 3
    Case 10, 11, 12
        quarter = 4
End Select
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_debug_print() {
        let source = r"
Debug.Print Month(Now)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_with_statement() {
        let source = r"
With employeeRecord
    .HireMonth = Month(.HireDate)
End With
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_elseif() {
        let source = r"
If x > 0 Then
    y = 1
ElseIf Month(startDate) > 6 Then
    y = 2
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_parentheses() {
        let source = r"
Dim m As Integer
m = (Month(Date))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_iif() {
        let source = r#"
Dim result As String
result = IIf(Month(Date) = 12, "December", "Other")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_class_usage() {
        let source = r"
Private m_month As Integer

Public Sub SetMonth()
    m_month = Month(Now)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_function_argument() {
        let source = r"
Call ProcessMonth(Month(Date))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_property_assignment() {
        let source = r"
Set obj = New DateInfo
obj.CurrentMonth = Month(Date)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_array_assignment() {
        let source = r"
Dim months(10) As Integer
Dim i As Integer
months(i) = Month(Date)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_for_loop() {
        let source = r"
Dim i As Integer
For i = 0 To 10
    monthValues(i) = Month(dates(i))
Next i
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_while_wend() {
        let source = r#"
While Month(currentDate) <= 6
    currentDate = DateAdd("m", 1, currentDate)
Wend
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_do_while() {
        let source = r#"
Do While Month(endDate) < 12
    endDate = DateAdd("m", 1, endDate)
Loop
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_do_until() {
        let source = r#"
Do Until Month(targetDate) = 1
    targetDate = DateAdd("m", 1, targetDate)
Loop
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_msgbox() {
        let source = r#"
MsgBox "Current month: " & Month(Now)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_concatenation() {
        let source = r#"
Dim dateStr As String
dateStr = Year(Date) & "/" & Month(Date) & "/" & Day(Date)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_comparison() {
        let source = r#"
If Month(date1) = Month(date2) Then
    MsgBox "Same month"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_format() {
        let source = r#"
Dim formatted As String
formatted = Format(Month(Date), "00")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_arithmetic() {
        let source = r"
Dim quarter As Integer
quarter = ((Month(Date) - 1) \ 3) + 1
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_label_caption() {
        let source = r#"
lblMonth.Caption = "Month: " & CStr(Month(Date))
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_calculation() {
        let source = r"
Dim fiscalPeriod As Integer
fiscalPeriod = Month(Date) - 9
If fiscalPeriod <= 0 Then fiscalPeriod = fiscalPeriod + 12
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn month_mod_operator() {
        let source = r#"
If Month(Date) Mod 3 = 0 Then
    MsgBox "Quarter end month"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/month",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
