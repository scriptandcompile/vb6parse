//! VB6 `Year` Function
//!
//! The `Year` function returns an Integer representing the year of a specified date.
//!
//! ## Syntax
//! ```vb6
//! Year(date)
//! ```
//!
//! ## Parameters
//! - `date`: Required. Any Variant, numeric expression, string expression, or any combination that can represent a date. If `date` contains Null, Null is returned.
//!
//! ## Returns
//! Returns an `Integer` representing the year (a whole number between 100 and 9999, inclusive).
//!
//! ## Remarks
//! The `Year` function extracts the year component from a date value:
//!
//! - **Return range**: Returns values from 100 to 9999
//! - **Null handling**: If the date argument is Null, the function returns Null
//! - **Date validation**: Invalid dates cause Error 13 (Type mismatch)
//! - **String dates**: Accepts string representations of dates (e.g., "12/25/2023")
//! - **Numeric dates**: Accepts numeric date values (serial dates)
//! - **Date literals**: Accepts date literals (e.g., #12/25/2023#)
//! - **Current year**: Use `Year(Date)` or `Year(Now)` to get current year
//! - **Leap year calculation**: Can be used to determine leap years
//! - **Year arithmetic**: Often combined with DateSerial for date calculations
//! - **Four-digit years**: Always returns full four-digit year (not two-digit)
//!
//! ### Date Components Family
//! The `Year` function is part of a family of date component extraction functions:
//! - `Year(date)` - Returns the year (100-9999)
//! - `Month(date)` - Returns the month (1-12)
//! - `Day(date)` - Returns the day (1-31)
//! - `Weekday(date)` - Returns the day of week (1-7)
//! - `Hour(date)` - Returns the hour (0-23)
//! - `Minute(date)` - Returns the minute (0-59)
//! - `Second(date)` - Returns the second (0-59)
//!
//! ### Leap Year Detection
//! ```vb6
//! Function IsLeapYear(yr As Integer) As Boolean
//!     IsLeapYear = ((yr Mod 4 = 0) And (yr Mod 100 <> 0)) Or (yr Mod 400 = 0)
//! End Function
//! ```
//!
//! ### Combining with DateSerial
//! ```vb6
//! ' Get first day of current year
//! firstDay = DateSerial(Year(Date), 1, 1)
//!
//! ' Get last day of current year
//! lastDay = DateSerial(Year(Date), 12, 31)
//! ```
//!
//! ## Typical Uses
//! 1. **Extract Year**: Get the year component from a date
//! 2. **Age Calculation**: Calculate age from birth date
//! 3. **Fiscal Year**: Determine fiscal year for financial reporting
//! 4. **Year Filtering**: Filter data by year
//! 5. **Year Validation**: Validate year ranges in data entry
//! 6. **Archive Organization**: Organize files/records by year
//! 7. **Year Comparison**: Compare dates across different years
//! 8. **Report Grouping**: Group data by year for reports
//!
//! ## Basic Examples
//!
//! ### Example 1: Get Current Year
//! ```vb6
//! Sub ShowCurrentYear()
//!     Dim currentYear As Integer
//!     currentYear = Year(Date)
//!     MsgBox "Current year: " & currentYear
//! End Sub
//! ```
//!
//! ### Example 2: Calculate Age
//! ```vb6
//! Function CalculateAge(birthDate As Date) As Integer
//!     Dim age As Integer
//!     age = Year(Date) - Year(birthDate)
//!     
//!     ' Adjust if birthday hasn't occurred this year
//!     If Month(Date) < Month(birthDate) Or _
//!        (Month(Date) = Month(birthDate) And Day(Date) < Day(birthDate)) Then
//!         age = age - 1
//!     End If
//!     
//!     CalculateAge = age
//! End Function
//! ```
//!
//! ### Example 3: Get Years Between Dates
//! ```vb6
//! Function YearsBetween(startDate As Date, endDate As Date) As Integer
//!     YearsBetween = Year(endDate) - Year(startDate)
//! End Function
//! ```
//!
//! ### Example 4: Check If Leap Year
//! ```vb6
//! Function IsLeapYear(dt As Date) As Boolean
//!     Dim yr As Integer
//!     yr = Year(dt)
//!     IsLeapYear = ((yr Mod 4 = 0) And (yr Mod 100 <> 0)) Or (yr Mod 400 = 0)
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: Get First Day of Year
//! ```vb6
//! Function GetFirstDayOfYear(dt As Date) As Date
//!     GetFirstDayOfYear = DateSerial(Year(dt), 1, 1)
//! End Function
//! ```
//!
//! ### Pattern 2: Get Last Day of Year
//! ```vb6
//! Function GetLastDayOfYear(dt As Date) As Date
//!     GetLastDayOfYear = DateSerial(Year(dt), 12, 31)
//! End Function
//! ```
//!
//! ### Pattern 3: Same Year Check
//! ```vb6
//! Function IsSameYear(date1 As Date, date2 As Date) As Boolean
//!     IsSameYear = (Year(date1) = Year(date2))
//! End Function
//! ```
//!
//! ### Pattern 4: Get Fiscal Year
//! ```vb6
//! Function GetFiscalYear(dt As Date, fiscalStartMonth As Integer) As Integer
//!     If Month(dt) >= fiscalStartMonth Then
//!         GetFiscalYear = Year(dt)
//!     Else
//!         GetFiscalYear = Year(dt) - 1
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 5: Format Year Display
//! ```vb6
//! Function FormatYearDisplay(dt As Date) As String
//!     FormatYearDisplay = "Year " & Year(dt)
//! End Function
//! ```
//!
//! ### Pattern 6: Year Difference
//! ```vb6
//! Function GetYearDifference(startDate As Date, endDate As Date) As Integer
//!     GetYearDifference = Abs(Year(endDate) - Year(startDate))
//! End Function
//! ```
//!
//! ### Pattern 7: Validate Year Range
//! ```vb6
//! Function IsYearInRange(dt As Date, minYear As Integer, maxYear As Integer) As Boolean
//!     Dim yr As Integer
//!     yr = Year(dt)
//!     IsYearInRange = (yr >= minYear And yr <= maxYear)
//! End Function
//! ```
//!
//! ### Pattern 8: Get Years Until Date
//! ```vb6
//! Function YearsUntil(targetDate As Date) As Integer
//!     YearsUntil = Year(targetDate) - Year(Date)
//! End Function
//! ```
//!
//! ### Pattern 9: Add Years to Date
//! ```vb6
//! Function AddYears(dt As Date, years As Integer) As Date
//!     AddYears = DateSerial(Year(dt) + years, Month(dt), Day(dt))
//! End Function
//! ```
//!
//! ### Pattern 10: Get Year-to-Date Range
//! ```vb6
//! Sub GetYTDRange(ByRef startDate As Date, ByRef endDate As Date)
//!     startDate = DateSerial(Year(Date), 1, 1)
//!     endDate = Date
//! End Sub
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Age Calculator Class
//! ```vb6
//! ' Class: AgeCalculator
//! ' Calculates precise age with various options
//! Option Explicit
//!
//! Public Function GetAge(birthDate As Date, Optional asOfDate As Variant) As Integer
//!     Dim referenceDate As Date
//!     Dim age As Integer
//!     
//!     If IsMissing(asOfDate) Then
//!         referenceDate = Date
//!     Else
//!         referenceDate = CDate(asOfDate)
//!     End If
//!     
//!     age = Year(referenceDate) - Year(birthDate)
//!     
//!     ' Adjust if birthday hasn't occurred yet
//!     If Month(referenceDate) < Month(birthDate) Or _
//!        (Month(referenceDate) = Month(birthDate) And _
//!         Day(referenceDate) < Day(birthDate)) Then
//!         age = age - 1
//!     End If
//!     
//!     GetAge = age
//! End Function
//!
//! Public Function GetAgeInYearsAndMonths(birthDate As Date) As String
//!     Dim years As Integer
//!     Dim months As Integer
//!     Dim tempDate As Date
//!     
//!     years = GetAge(birthDate)
//!     tempDate = DateSerial(Year(birthDate) + years, Month(birthDate), Day(birthDate))
//!     months = DateDiff("m", tempDate, Date)
//!     
//!     GetAgeInYearsAndMonths = years & " years, " & months & " months"
//! End Function
//!
//! Public Function WillBeBirthdayThisYear(birthDate As Date) As Boolean
//!     Dim birthdayThisYear As Date
//!     birthdayThisYear = DateSerial(Year(Date), Month(birthDate), Day(birthDate))
//!     WillBeBirthdayThisYear = (birthdayThisYear >= Date)
//! End Function
//!
//! Public Function GetAgeAtDate(birthDate As Date, targetDate As Date) As Integer
//!     GetAgeAtDate = GetAge(birthDate, targetDate)
//! End Function
//! ```
//!
//! ### Example 2: Fiscal Year Manager Module
//! ```vb6
//! ' Module: FiscalYearManager
//! ' Manages fiscal year calculations
//! Option Explicit
//!
//! Private m_FiscalStartMonth As Integer
//!
//! Public Sub SetFiscalYearStart(startMonth As Integer)
//!     If startMonth < 1 Or startMonth > 12 Then
//!         Err.Raise 5, , "Start month must be between 1 and 12"
//!     End If
//!     m_FiscalStartMonth = startMonth
//! End Sub
//!
//! Public Function GetFiscalYear(dt As Date) As Integer
//!     If m_FiscalStartMonth = 0 Then m_FiscalStartMonth = 1
//!     
//!     If Month(dt) >= m_FiscalStartMonth Then
//!         GetFiscalYear = Year(dt)
//!     Else
//!         GetFiscalYear = Year(dt) - 1
//!     End If
//! End Function
//!
//! Public Function GetFiscalYearStart(fiscalYear As Integer) As Date
//!     If m_FiscalStartMonth = 0 Then m_FiscalStartMonth = 1
//!     GetFiscalYearStart = DateSerial(fiscalYear, m_FiscalStartMonth, 1)
//! End Function
//!
//! Public Function GetFiscalYearEnd(fiscalYear As Integer) As Date
//!     Dim endMonth As Integer
//!     Dim endYear As Integer
//!     
//!     If m_FiscalStartMonth = 0 Then m_FiscalStartMonth = 1
//!     
//!     endMonth = m_FiscalStartMonth - 1
//!     If endMonth = 0 Then endMonth = 12
//!     
//!     If m_FiscalStartMonth = 1 Then
//!         endYear = fiscalYear
//!     Else
//!         endYear = fiscalYear + 1
//!     End If
//!     
//!     GetFiscalYearEnd = DateSerial(endYear, endMonth, Day(DateSerial(endYear, endMonth + 1, 0)))
//! End Function
//!
//! Public Function FormatFiscalYear(fiscalYear As Integer) As String
//!     If m_FiscalStartMonth = 1 Then
//!         FormatFiscalYear = "FY" & fiscalYear
//!     Else
//!         FormatFiscalYear = "FY" & fiscalYear & "-" & Right$(CStr(fiscalYear + 1), 2)
//!     End If
//! End Function
//! ```
//!
//! ### Example 3: Year Range Analyzer Class
//! ```vb6
//! ' Class: YearRangeAnalyzer
//! ' Analyzes year ranges in date collections
//! Option Explicit
//!
//! Public Function GetYearRange(dates() As Date) As String
//!     Dim minYear As Integer
//!     Dim maxYear As Integer
//!     Dim i As Long
//!     Dim yr As Integer
//!     
//!     If UBound(dates) < LBound(dates) Then
//!         GetYearRange = "No dates"
//!         Exit Function
//!     End If
//!     
//!     minYear = Year(dates(LBound(dates)))
//!     maxYear = minYear
//!     
//!     For i = LBound(dates) To UBound(dates)
//!         yr = Year(dates(i))
//!         If yr < minYear Then minYear = yr
//!         If yr > maxYear Then maxYear = yr
//!     Next i
//!     
//!     If minYear = maxYear Then
//!         GetYearRange = CStr(minYear)
//!     Else
//!         GetYearRange = minYear & "-" & maxYear
//!     End If
//! End Function
//!
//! Public Function GetYearDistribution(dates() As Date) As Collection
//!     Dim distribution As New Collection
//!     Dim i As Long
//!     Dim yr As Integer
//!     Dim yearKey As String
//!     Dim count As Long
//!     
//!     For i = LBound(dates) To UBound(dates)
//!         yr = Year(dates(i))
//!         yearKey = CStr(yr)
//!         
//!         On Error Resume Next
//!         count = distribution(yearKey)
//!         If Err.Number <> 0 Then
//!             distribution.Add 1, yearKey
//!             Err.Clear
//!         Else
//!             distribution.Remove yearKey
//!             distribution.Add count + 1, yearKey
//!         End If
//!         On Error GoTo 0
//!     Next i
//!     
//!     Set GetYearDistribution = distribution
//! End Function
//!
//! Public Function GetMostCommonYear(dates() As Date) As Integer
//!     Dim distribution As Collection
//!     Dim maxCount As Long
//!     Dim maxYear As Integer
//!     Dim yr As Variant
//!     
//!     Set distribution = GetYearDistribution(dates)
//!     
//!     maxCount = 0
//!     For Each yr In distribution
//!         If distribution(CStr(yr)) > maxCount Then
//!             maxCount = distribution(CStr(yr))
//!             maxYear = CInt(yr)
//!         End If
//!     Next yr
//!     
//!     GetMostCommonYear = maxYear
//! End Function
//!
//! Public Function GetUniqueYears(dates() As Date) As Collection
//!     Dim years As New Collection
//!     Dim i As Long
//!     Dim yr As Integer
//!     Dim yearKey As String
//!     
//!     For i = LBound(dates) To UBound(dates)
//!         yr = Year(dates(i))
//!         yearKey = CStr(yr)
//!         
//!         On Error Resume Next
//!         years.Add yr, yearKey
//!         On Error GoTo 0
//!     Next i
//!     
//!     Set GetUniqueYears = years
//! End Function
//! ```
//!
//! ### Example 4: Date Archive Organizer Module
//! ```vb6
//! ' Module: DateArchiveOrganizer
//! ' Organizes files and data by year
//! Option Explicit
//!
//! Public Function GetArchivePath(baseFolder As String, dt As Date) As String
//!     Dim yearFolder As String
//!     yearFolder = baseFolder
//!     If Right$(yearFolder, 1) <> "\" Then yearFolder = yearFolder & "\"
//!     yearFolder = yearFolder & Year(dt) & "\"
//!     GetArchivePath = yearFolder
//! End Function
//!
//! Public Function CreateYearlyArchiveFolders(baseFolder As String, _
//!                                            startYear As Integer, _
//!                                            endYear As Integer) As Long
//!     Dim yr As Integer
//!     Dim folderPath As String
//!     Dim count As Long
//!     
//!     count = 0
//!     For yr = startYear To endYear
//!         folderPath = baseFolder
//!         If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
//!         folderPath = folderPath & yr & "\"
//!         
//!         On Error Resume Next
//!         MkDir folderPath
//!         If Err.Number = 0 Then count = count + 1
//!         Err.Clear
//!         On Error GoTo 0
//!     Next yr
//!     
//!     CreateYearlyArchiveFolders = count
//! End Function
//!
//! Public Function GetYearFromFilename(filename As String) As Integer
//!     Dim parts() As String
//!     Dim part As Variant
//!     Dim yr As Integer
//!     
//!     parts = Split(filename, "_")
//!     For Each part In parts
//!         If IsNumeric(part) Then
//!             yr = CInt(part)
//!             If yr >= 1900 And yr <= 9999 Then
//!                 GetYearFromFilename = yr
//!                 Exit Function
//!             End If
//!         End If
//!     Next part
//!     
//!     GetYearFromFilename = 0
//! End Function
//!
//! Public Function GenerateYearlyReport(data As Collection, reportYear As Integer) As String
//!     Dim report As String
//!     Dim item As Variant
//!     Dim itemDate As Date
//!     Dim count As Long
//!     
//!     report = "Year " & reportYear & " Report" & vbCrLf
//!     report = report & String$(50, "=") & vbCrLf
//!     
//!     count = 0
//!     For Each item In data
//!         itemDate = CDate(item)
//!         If Year(itemDate) = reportYear Then
//!             count = count + 1
//!         End If
//!     Next item
//!     
//!     report = report & "Total items: " & count & vbCrLf
//!     GenerateYearlyReport = report
//! End Function
//! ```
//!
//! ## Error Handling
//! The `Year` function can raise the following errors:
//!
//! - **Error 13 (Type mismatch)**: If the argument cannot be interpreted as a date
//! - **Error 5 (Invalid procedure call)**: If the date is outside the valid range
//! - **Returns Null**: If the input is Null (not an error)
//!
//! ## Performance Notes
//! - Very fast operation - direct extraction from date value
//! - Constant time O(1) complexity
//! - No performance penalty for different date formats
//! - Safe to call repeatedly in loops
//! - Consider caching if used extensively with same date
//!
//! ## Best Practices
//! 1. **Validate input** before calling if date source is uncertain
//! 2. **Handle Null** explicitly when working with nullable date fields
//! 3. **Use DateSerial** with Year for date construction/manipulation
//! 4. **Combine with Month/Day** for complete date component extraction
//! 5. **Cache results** when using same date repeatedly
//! 6. **Use fiscal year functions** for business date calculations
//! 7. **Consider leap years** when performing year-based calculations
//! 8. **Use DateDiff** for accurate year differences accounting for partial years
//! 9. **Test edge cases** like end-of-year dates and leap year boundaries
//! 10. **Document assumptions** about calendar systems and year ranges
//!
//! ## Comparison Table
//!
//! | Function | Returns | Range | Purpose |
//! |----------|---------|-------|---------|
//! | `Year` | Integer | 100-9999 | Year component |
//! | `Month` | Integer | 1-12 | Month component |
//! | `Day` | Integer | 1-31 | Day component |
//! | `DatePart` | Variant | Varies | Any date part |
//! | `Format$` | String | N/A | Formatted date |
//!
//! ## Platform Notes
//! - Available in VB6, VBA, and VBScript
//! - Consistent behavior across platforms
//! - Always returns four-digit year (not affected by regional settings)
//! - Date range: January 1, 100 to December 31, 9999
//! - Dates before year 100 or after year 9999 cause errors
//!
//! ## Limitations
//! - Cannot return two-digit year (always four digits)
//! - Cannot handle dates before year 100
//! - Cannot handle dates after year 9999
//! - No built-in fiscal year calculation (requires custom function)
//! - Does not account for different calendar systems
//! - No built-in leap year detection (requires separate function)
//! - Cannot extract century separately (must calculate from year)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_year_basic() {
        let source = r#"
Sub Test()
    currentYear = Year(Date)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_variable_assignment() {
        let source = r#"
Sub Test()
    Dim yr As Integer
    yr = Year(someDate)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
        assert!(debug.contains("someDate"));
    }

    #[test]
    fn test_year_function_return() {
        let source = r#"
Function GetYear(dt As Date) As Integer
    GetYear = Year(dt)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_if_statement() {
        let source = r#"
Sub Test()
    If Year(dt) = 2023 Then
        Process
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_comparison() {
        let source = r#"
Sub Test()
    If Year(date1) > Year(date2) Then
        Later
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_arithmetic() {
        let source = r#"
Sub Test()
    age = Year(Date) - Year(birthDate)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Year: " & Year(Now)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print Year(targetDate)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_for_loop() {
        let source = r#"
Sub Test()
    For i = Year(startDate) To Year(endDate)
        ProcessYear i
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_select_case() {
        let source = r#"
Sub Test()
    Select Case Year(dt)
        Case 2020
            DoLeapYear
        Case 2021
            DoNormal
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_function_argument() {
        let source = r#"
Sub Test()
    Call ProcessYear(Year(recordDate))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_dateserial() {
        let source = r#"
Sub Test()
    firstDay = DateSerial(Year(dt), 1, 1)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_property_assignment() {
        let source = r#"
Sub Test()
    obj.Year = Year(obj.Date)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_with_statement() {
        let source = r#"
Sub Test()
    With dateInfo
        .Year = Year(.DateValue)
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_array_assignment() {
        let source = r#"
Sub Test()
    years(i) = Year(dates(i))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_concatenation() {
        let source = r#"
Sub Test()
    display = "Year " & Year(dt) & " Report"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_parentheses() {
        let source = r#"
Sub Test()
    result = (Year(dt))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    yr = Year(userInput)
    If Err.Number <> 0 Then
        yr = 0
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_print_statement() {
        let source = r#"
Sub Test()
    Print #1, Year(recordDate)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_class_usage() {
        let source = r#"
Sub Test()
    Set analyzer = New YearAnalyzer
    analyzer.CurrentYear = Year(Date)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_elseif() {
        let source = r#"
Sub Test()
    If x = 1 Then
        y = 1
    ElseIf Year(dt) = 2023 Then
        y = 2
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_do_while() {
        let source = r#"
Sub Test()
    Do While Year(dt) < 2025
        dt = DateAdd("yyyy", 1, dt)
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_do_until() {
        let source = r#"
Sub Test()
    Do Until Year(dt) >= targetYear
        dt = DateAdd("yyyy", 1, dt)
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_while_wend() {
        let source = r#"
Sub Test()
    While Year(dt) < 2030
        ProcessYear Year(dt)
        dt = DateAdd("yyyy", 1, dt)
    Wend
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_iif() {
        let source = r#"
Sub Test()
    display = IIf(Year(dt) = Year(Date), "This year", "Other year")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_mod_operation() {
        let source = r#"
Sub Test()
    isLeap = (Year(dt) Mod 4 = 0)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_format() {
        let source = r#"
Sub Test()
    formatted = Format$(dt, "mmmm d, ") & Year(dt)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }

    #[test]
    fn test_year_cstr() {
        let source = r#"
Sub Test()
    yearString = CStr(Year(dt))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Year"));
    }
}
