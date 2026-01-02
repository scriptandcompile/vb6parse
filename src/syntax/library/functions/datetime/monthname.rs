//! # `MonthName` Function
//!
//! Returns a String indicating the specified month.
//!
//! ## Syntax
//!
//! ```vb
//! MonthName(month, [abbreviate])
//! ```
//!
//! ## Parameters
//!
//! - **month** (Required) - Numeric designation of the month. For example, January is 1, February is 2, and so on.
//! - **abbreviate** (Optional) - Boolean value that indicates if the month name is to be abbreviated. If omitted, the default is False, which means the month name is not abbreviated.
//!
//! ## Return Value
//!
//! Returns a **String** representing the month name:
//! - If abbreviate is False or omitted: Full month name (e.g., "January", "February")
//! - If abbreviate is True: Abbreviated month name (e.g., "Jan", "Feb")
//!
//! ## Remarks
//!
//! The `MonthName` function returns the localized name of the month based on the system's regional settings.
//! It is commonly used for displaying month names in user interfaces, reports, and formatted output.
//!
//! ### Key Characteristics:
//! - Returns localized month names based on system regional settings
//! - Month parameter must be between 1 and 12
//! - Error 5 (Invalid procedure call) if month is less than 1 or greater than 12
//! - Abbreviate parameter is optional; defaults to False (full name)
//! - Abbreviated names are typically 3 characters but may vary by locale
//! - Much simpler than maintaining arrays of month names
//! - Can be combined with `Month()` function for date formatting
//! - Respects system locale for internationalization
//!
//! ### Common Use Cases:
//! - Display month names in user interfaces
//! - Format dates for reports and output
//! - Create drop-down lists of months
//! - Generate calendar displays
//! - Internationalized date formatting
//! - Chart and graph labels
//! - File naming with month names
//! - Data export with readable month names
//!
//! ## Typical Uses
//!
//! 1. **Display Month Names** - Convert month numbers to readable names
//! 2. **Report Headers** - Format report titles with month names
//! 3. **Calendar Controls** - Populate month selectors
//! 4. **Data Export** - Include readable month names in exports
//! 5. **User Interface** - Display current or selected month
//! 6. **Chart Labels** - Label chart axes with month names
//! 7. **File Naming** - Create month-based file names
//! 8. **Internationalization** - Automatic locale-based month names
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Get full month name
//! Dim monthName As String
//! monthName = MonthName(3)
//! ' Returns "March"
//! ```
//!
//! ```vb
//! ' Example 2: Get abbreviated month name
//! Dim shortMonth As String
//! shortMonth = MonthName(11, True)
//! ' Returns "Nov"
//! ```
//!
//! ```vb
//! ' Example 3: Display current month
//! Dim currentMonthName As String
//! currentMonthName = MonthName(Month(Date))
//! MsgBox "Current month is " & currentMonthName
//! ```
//!
//! ```vb
//! ' Example 4: Format date with month name
//! Dim formattedDate As String
//! Dim d As Date
//! d = #5/15/2025#
//! formattedDate = MonthName(Month(d)) & " " & Day(d) & ", " & Year(d)
//! ' Returns "May 15, 2025"
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Safe month name with validation
//! Function SafeMonthName(monthNumber As Integer, _
//!                        Optional abbreviate As Boolean = False) As String
//!     If monthNumber < 1 Or monthNumber > 12 Then
//!         SafeMonthName = ""
//!     Else
//!         SafeMonthName = MonthName(monthNumber, abbreviate)
//!     End If
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 2: Populate month combo box
//! Sub PopulateMonthCombo(cbo As ComboBox, Optional abbreviate As Boolean = False)
//!     Dim i As Integer
//!     cbo.Clear
//!     For i = 1 To 12
//!         cbo.AddItem MonthName(i, abbreviate)
//!     Next i
//! End Sub
//! ```
//!
//! ```vb
//! ' Pattern 3: Get month name from date
//! Function GetMonthNameFromDate(dateValue As Date, _
//!                               Optional abbreviate As Boolean = False) As String
//!     GetMonthNameFromDate = MonthName(Month(dateValue), abbreviate)
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 4: Format date range with month names
//! Function FormatDateRange(startDate As Date, endDate As Date) As String
//!     Dim startMonth As String
//!     Dim endMonth As String
//!     
//!     startMonth = MonthName(Month(startDate), True)
//!     endMonth = MonthName(Month(endDate), True)
//!     
//!     FormatDateRange = startMonth & " " & Day(startDate) & " - " & _
//!                       endMonth & " " & Day(endDate) & ", " & Year(endDate)
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 5: Create month-based filename
//! Function GetMonthlyFileName(fileDate As Date, fileType As String) As String
//!     Dim monthAbbrev As String
//!     monthAbbrev = MonthName(Month(fileDate), True)
//!     GetMonthlyFileName = fileType & "_" & Year(fileDate) & "_" & monthAbbrev & ".dat"
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 6: Build month selection array
//! Function GetMonthArray(fullNames As Boolean) As String()
//!     Dim months() As String
//!     Dim i As Integer
//!     
//!     ReDim months(1 To 12)
//!     For i = 1 To 12
//!         months(i) = MonthName(i, Not fullNames)
//!     Next i
//!     
//!     GetMonthArray = months
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 7: Format month/year string
//! Function FormatMonthYear(dateValue As Date, abbreviate As Boolean) As String
//!     FormatMonthYear = MonthName(Month(dateValue), abbreviate) & " " & Year(dateValue)
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 8: Get fiscal month name
//! Function GetFiscalMonthName(fiscalMonth As Integer, abbreviate As Boolean) As String
//!     ' Fiscal month 1 might be October (calendar month 10)
//!     Dim calendarMonth As Integer
//!     Dim fiscalStartMonth As Integer
//!     
//!     fiscalStartMonth = 10 ' October
//!     calendarMonth = fiscalMonth + fiscalStartMonth - 1
//!     
//!     If calendarMonth > 12 Then
//!         calendarMonth = calendarMonth - 12
//!     End If
//!     
//!     GetFiscalMonthName = MonthName(calendarMonth, abbreviate)
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 9: Create report header
//! Function CreateMonthlyReportHeader(reportDate As Date) As String
//!     Dim header As String
//!     header = "Monthly Report - " & MonthName(Month(reportDate)) & " " & Year(reportDate)
//!     CreateMonthlyReportHeader = header
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 10: Compare month names
//! Function IsSameMonth(date1 As Date, date2 As Date) As Boolean
//!     IsSameMonth = (MonthName(Month(date1)) = MonthName(Month(date2))) And _
//!                   (Year(date1) = Year(date2))
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Calendar Generator
//!
//! ```vb
//! ' Class: CalendarGenerator
//! ' Generates formatted calendar displays
//!
//! Option Explicit
//!
//! Private m_year As Integer
//! Private m_month As Integer
//! Private m_firstDayOfWeek As VbDayOfWeek
//!
//! Public Sub Initialize(calYear As Integer, calMonth As Integer)
//!     m_year = calYear
//!     m_month = calMonth
//!     m_firstDayOfWeek = vbSunday
//! End Sub
//!
//! Public Function GetMonthHeader(abbreviate As Boolean) As String
//!     GetMonthHeader = MonthName(m_month, abbreviate) & " " & m_year
//! End Function
//!
//! Public Function GenerateCalendar() As String
//!     Dim calendar As String
//!     Dim firstDay As Date
//!     Dim lastDay As Date
//!     Dim currentDay As Date
//!     Dim dayOfWeek As Integer
//!     Dim weekRow As String
//!     Dim dayCount As Integer
//!     
//!     ' Create header
//!     calendar = String(28, " ") & GetMonthHeader(False) & vbCrLf
//!     calendar = calendar & "Su  Mo  Tu  We  Th  Fr  Sa" & vbCrLf
//!     
//!     ' Get first and last day of month
//!     firstDay = DateSerial(m_year, m_month, 1)
//!     lastDay = DateSerial(m_year, m_month + 1, 0)
//!     
//!     ' Start with padding for first week
//!     dayOfWeek = Weekday(firstDay, m_firstDayOfWeek)
//!     weekRow = String((dayOfWeek - 1) * 4, " ")
//!     
//!     ' Fill in days
//!     currentDay = firstDay
//!     Do While currentDay <= lastDay
//!         weekRow = weekRow & Format(Day(currentDay), "00") & "  "
//!         dayOfWeek = dayOfWeek + 1
//!         
//!         If dayOfWeek > 7 Then
//!             calendar = calendar & weekRow & vbCrLf
//!             weekRow = ""
//!             dayOfWeek = 1
//!         End If
//!         
//!         currentDay = DateAdd("d", 1, currentDay)
//!     Loop
//!     
//!     If Len(weekRow) > 0 Then
//!         calendar = calendar & weekRow & vbCrLf
//!     End If
//!     
//!     GenerateCalendar = calendar
//! End Function
//!
//! Public Function GetAllMonthsForYear(abbreviate As Boolean) As String()
//!     Dim months() As String
//!     Dim i As Integer
//!     
//!     ReDim months(1 To 12)
//!     For i = 1 To 12
//!         months(i) = MonthName(i, abbreviate)
//!     Next i
//!     
//!     GetAllMonthsForYear = months
//! End Function
//! ```
//!
//! ### Example 2: Report Generator with Month Names
//!
//! ```vb
//! ' Class: MonthlyReportFormatter
//! ' Formats reports with localized month names
//!
//! Option Explicit
//!
//! Private m_useAbbreviations As Boolean
//! Private m_reportData As Collection
//!
//! Public Sub Initialize(useAbbrev As Boolean)
//!     m_useAbbreviations = useAbbrev
//!     Set m_reportData = New Collection
//! End Sub
//!
//! Public Sub AddMonthData(monthNum As Integer, dataValue As Double)
//!     Dim entry As String
//!     entry = MonthName(monthNum, m_useAbbreviations) & ":" & dataValue
//!     m_reportData.Add entry
//! End Sub
//!
//! Public Function GenerateReport() As String
//!     Dim report As String
//!     Dim item As Variant
//!     Dim i As Integer
//!     
//!     report = "Monthly Summary Report" & vbCrLf
//!     report = report & String(50, "-") & vbCrLf
//!     
//!     For Each item In m_reportData
//!         report = report & item & vbCrLf
//!     Next item
//!     
//!     GenerateReport = report
//! End Function
//!
//! Public Function FormatDateWithMonth(dateValue As Date) As String
//!     FormatDateWithMonth = MonthName(Month(dateValue), m_useAbbreviations) & _
//!                          " " & Day(dateValue) & ", " & Year(dateValue)
//! End Function
//!
//! Public Function GetQuarterMonths(quarter As Integer) As String
//!     Dim months As String
//!     Dim startMonth As Integer
//!     Dim i As Integer
//!     
//!     startMonth = ((quarter - 1) * 3) + 1
//!     
//!     For i = 0 To 2
//!         If i > 0 Then months = months & ", "
//!         months = months & MonthName(startMonth + i, m_useAbbreviations)
//!     Next i
//!     
//!     GetQuarterMonths = months
//! End Function
//! ```
//!
//! ### Example 3: Internationalized Date Formatter
//!
//! ```vb
//! ' Module: InternationalDateFormatter
//! ' Provides locale-aware date formatting using MonthName
//!
//! Option Explicit
//!
//! Public Enum DateFormat
//!     dfLong = 0          ' "January 15, 2025"
//!     dfMedium = 1        ' "Jan 15, 2025"
//!     dfShort = 2         ' "01/15/2025"
//!     dfMonthYear = 3     ' "January 2025"
//!     dfMonthYearShort = 4 ' "Jan 2025"
//! End Enum
//!
//! Public Function FormatDate(dateValue As Date, formatType As DateFormat) As String
//!     Select Case formatType
//!         Case dfLong
//!             FormatDate = MonthName(Month(dateValue)) & " " & _
//!                         Day(dateValue) & ", " & Year(dateValue)
//!         
//!         Case dfMedium
//!             FormatDate = MonthName(Month(dateValue), True) & " " & _
//!                         Day(dateValue) & ", " & Year(dateValue)
//!         
//!         Case dfShort
//!             FormatDate = Format(dateValue, "mm/dd/yyyy")
//!         
//!         Case dfMonthYear
//!             FormatDate = MonthName(Month(dateValue)) & " " & Year(dateValue)
//!         
//!         Case dfMonthYearShort
//!             FormatDate = MonthName(Month(dateValue), True) & " " & Year(dateValue)
//!     End Select
//! End Function
//!
//! Public Function FormatDateRange(startDate As Date, endDate As Date, _
//!                                abbreviate As Boolean) As String
//!     Dim result As String
//!     
//!     If Year(startDate) = Year(endDate) Then
//!         If Month(startDate) = Month(endDate) Then
//!             ' Same month and year
//!             result = MonthName(Month(startDate), abbreviate) & " " & _
//!                     Day(startDate) & "-" & Day(endDate) & ", " & Year(startDate)
//!         Else
//!             ' Different months, same year
//!             result = MonthName(Month(startDate), abbreviate) & " " & Day(startDate) & " - " & _
//!                     MonthName(Month(endDate), abbreviate) & " " & Day(endDate) & ", " & _
//!                     Year(endDate)
//!         End If
//!     Else
//!         ' Different years
//!         result = MonthName(Month(startDate), abbreviate) & " " & Day(startDate) & ", " & _
//!                 Year(startDate) & " - " & _
//!                 MonthName(Month(endDate), abbreviate) & " " & Day(endDate) & ", " & _
//!                 Year(endDate)
//!     End If
//!     
//!     FormatDateRange = result
//! End Function
//!
//! Public Function GetMonthNames(abbreviate As Boolean) As String()
//!     Dim months(1 To 12) As String
//!     Dim i As Integer
//!     
//!     For i = 1 To 12
//!         months(i) = MonthName(i, abbreviate)
//!     Next i
//!     
//!     GetMonthNames = months
//! End Function
//!
//! Public Function ParseMonthName(monthString As String) As Integer
//!     Dim i As Integer
//!     
//!     ' Try full names first
//!     For i = 1 To 12
//!         If UCase(MonthName(i)) = UCase(monthString) Then
//!             ParseMonthName = i
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     ' Try abbreviated names
//!     For i = 1 To 12
//!         If UCase(MonthName(i, True)) = UCase(monthString) Then
//!             ParseMonthName = i
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     ParseMonthName = 0 ' Not found
//! End Function
//! ```
//!
//! ### Example 4: Chart Label Generator
//!
//! ```vb
//! ' Class: ChartLabelGenerator
//! ' Generates chart labels with month names
//!
//! Option Explicit
//!
//! Private m_startMonth As Integer
//! Private m_monthCount As Integer
//! Private m_abbreviate As Boolean
//!
//! Public Sub Initialize(startMonth As Integer, monthCount As Integer, _
//!                       abbreviate As Boolean)
//!     If startMonth < 1 Or startMonth > 12 Then
//!         Err.Raise 5, "ChartLabelGenerator", "Invalid start month"
//!     End If
//!     
//!     m_startMonth = startMonth
//!     m_monthCount = monthCount
//!     m_abbreviate = abbreviate
//! End Sub
//!
//! Public Function GetLabels() As String()
//!     Dim labels() As String
//!     Dim i As Integer
//!     Dim currentMonth As Integer
//!     
//!     ReDim labels(0 To m_monthCount - 1)
//!     
//!     For i = 0 To m_monthCount - 1
//!         currentMonth = m_startMonth + i
//!         
//!         ' Wrap around if exceeds 12
//!         If currentMonth > 12 Then
//!             currentMonth = currentMonth - 12
//!         End If
//!         
//!         labels(i) = MonthName(currentMonth, m_abbreviate)
//!     Next i
//!     
//!     GetLabels = labels
//! End Function
//!
//! Public Function GetLabelWithYear(monthNum As Integer, yearNum As Integer) As String
//!     GetLabelWithYear = MonthName(monthNum, m_abbreviate) & " '" & _
//!                        Right(CStr(yearNum), 2)
//! End Function
//!
//! Public Function GetQuarterLabel(quarter As Integer, yearNum As Integer) As String
//!     Dim firstMonth As Integer
//!     firstMonth = ((quarter - 1) * 3) + 1
//!     
//!     GetQuarterLabel = "Q" & quarter & " (" & _
//!                      MonthName(firstMonth, True) & "-" & _
//!                      MonthName(firstMonth + 2, True) & " " & _
//!                      yearNum & ")"
//! End Function
//!
//! Public Function GetYearLabels(abbreviate As Boolean) As String()
//!     Dim labels(1 To 12) As String
//!     Dim i As Integer
//!     
//!     For i = 1 To 12
//!         labels(i) = MonthName(i, abbreviate)
//!     Next i
//!     
//!     GetYearLabels = labels
//! End Function
//!
//! Public Function FormatMultiYearLabel(monthNum As Integer, _
//!                                     startYear As Integer, _
//!                                     endYear As Integer) As String
//!     FormatMultiYearLabel = MonthName(monthNum, m_abbreviate) & " " & _
//!                           startYear & "-" & endYear
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! On Error Resume Next
//! monthName = MonthName(monthNum)
//! If Err.Number = 5 Then
//!     MsgBox "Invalid month number. Must be between 1 and 12."
//! ElseIf Err.Number <> 0 Then
//!     MsgBox "Error getting month name: " & Err.Description
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Performance Considerations
//!
//! - `MonthName()` is very fast - simple lookup operation
//! - For repeated calls, consider caching month name arrays
//! - No performance difference between full and abbreviated names
//! - Locale lookup may add minimal overhead but is negligible
//! - More efficient than maintaining hardcoded month name arrays
//! - Can be called millions of times without performance concerns
//!
//! ## Best Practices
//!
//! 1. **Validate month numbers** - Always check that month is 1-12 before calling
//! 2. **Use for localization** - Leverages system locale for international support
//! 3. **Combine with `Month()`** - Extract month number from dates, then get name
//! 4. **Consider abbreviations** - Use abbreviated names for space-constrained displays
//! 5. **Cache for UI** - Populate drop-downs once, not on every refresh
//! 6. **Document format** - Clearly state whether full or abbreviated names are expected
//! 7. **Test locales** - Test with different regional settings for international apps
//! 8. **Handle errors** - Wrap in error handling for robust code
//! 9. **Use constants** - Define month number constants for readability
//! 10. **Prefer over arrays** - Use `MonthName()` instead of hardcoded name arrays
//!
//! ## Comparison with Other Approaches
//!
//! | Approach | Pros | Cons |
//! |----------|------|------|
//! | **`MonthName()`** | Automatic localization, simple, no maintenance | Requires VBA/VB6, slightly slower than array |
//! | **Hardcoded array** | Fast, full control | No localization, maintenance burden |
//! | **`Format()`** | Flexible formatting | Returns different formats, not just name |
//! | **`DatePart()`** | Returns number, not name | Need to convert to name separately |
//!
//! ## Localization Notes
//!
//! The `MonthName` function returns month names based on the system's regional settings:
//! - **English (US)**: "January", "February", etc. / "Jan", "Feb", etc.
//! - **Spanish**: "enero", "febrero", etc. / "ene", "feb", etc.
//! - **French**: "janvier", "février", etc. / "janv", "févr", etc.
//! - **German**: "Januar", "Februar", etc. / "Jan", "Feb", etc.
//! - **Japanese**: "1月", "2月", etc.
//! - And many more locales...
//!
//! ## Platform Notes
//!
//! - Available in VBA (Excel, Access, Word, etc.)
//! - Available in VB6
//! - **Not available in `VBScript`** (`VBScript` lacks this function)
//! - Returns String type
//! - Consistent across all VBA platforms
//! - Uses Control Panel regional settings
//! - First introduced in VB6 and VBA 6.0
//!
//! ## Limitations
//!
//! - Month parameter must be 1-12 (no 0 or 13+)
//! - Error 5 if month is out of valid range
//! - Returns names based on system locale only (cannot specify locale)
//! - Abbreviated length varies by locale (usually 3 chars, but not always)
//! - Not available in `VBScript`
//! - Cannot customize the returned names
//!
//! ## Related Functions
//!
//! - **Month** - Returns the month number (1-12) from a date
//! - **`WeekdayName`** - Returns the name of the weekday (similar function for days)
//! - **Format** - Can format dates with month names using format strings
//! - **`DatePart`** - Returns various parts of a date (including month number)
//! - **Year** - Returns the year component of a date
//! - **Day** - Returns the day component of a date
//!
//! ## VB6 Parser Notes
//!
//! `MonthName` is parsed as a regular function call (`CallExpression`). This module exists primarily
//! for documentation purposes to provide comprehensive reference material for VB6 developers
//! working with date formatting and display operations requiring localized month names.

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::parsers::cst::ConcreteSyntaxTree;
    #[test]
    fn monthname_basic() {
        let source = r"
Dim name As String
name = MonthName(3)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                NameKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            NameStatement {
                NameKeyword,
                Whitespace,
                EqualityOperator,
                Whitespace,
                Identifier ("MonthName"),
                LeftParenthesis,
                IntegerLiteral ("3"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn monthname_variable_assignment() {
        let source = r"
Dim monthName As String
monthName = MonthName(Month(Date))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("monthName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("monthName"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("MonthName"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Month"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            DateKeyword,
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
        ]);
    }

    #[test]
    fn monthname_abbreviated() {
        let source = r"
Dim shortName As String
shortName = MonthName(11, True)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("shortName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("shortName"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("MonthName"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("11"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword,
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
    fn monthname_if_statement() {
        let source = r#"
If MonthName(Month(Date)) = "November" Then
    MsgBox "It's November"
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
                        Identifier ("MonthName"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                CallExpression {
                                    Identifier ("Month"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                DateKeyword,
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    StringLiteralExpression {
                        StringLiteral ("\"November\""),
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
                        StringLiteral ("\"It's November\""),
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
    fn monthname_function_return() {
        let source = r"
Function GetCurrentMonthName() As String
    GetCurrentMonthName = MonthName(Month(Date))
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetCurrentMonthName"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetCurrentMonthName"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MonthName"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Month"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    DateKeyword,
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
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn monthname_concatenation() {
        let source = r#"
Dim dateStr As String
dateStr = MonthName(Month(Date)) & " " & Day(Date) & ", " & Year(Date)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("dateStr"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("dateStr"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    BinaryExpression {
                        BinaryExpression {
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("MonthName"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            CallExpression {
                                                Identifier ("Month"),
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        IdentifierExpression {
                                                            DateKeyword,
                                                        },
                                                    },
                                                },
                                                RightParenthesis,
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                StringLiteralExpression {
                                    StringLiteral ("\" \""),
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Day"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            DateKeyword,
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\", \""),
                        },
                    },
                    Whitespace,
                    Ampersand,
                    Whitespace,
                    CallExpression {
                        Identifier ("Year"),
                        LeftParenthesis,
                        ArgumentList {
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
    fn monthname_debug_print() {
        let source = r"
Debug.Print MonthName(5)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("Debug"),
                PeriodOperator,
                PrintKeyword,
                Whitespace,
                Identifier ("MonthName"),
                LeftParenthesis,
                IntegerLiteral ("5"),
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn monthname_with_statement() {
        let source = r"
With reportData
    .MonthDisplay = MonthName(Month(.ReportDate), True)
End With
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            WithStatement {
                WithKeyword,
                Whitespace,
                Identifier ("reportData"),
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            PeriodOperator,
                        },
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("MonthDisplay"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            CallExpression {
                                Identifier ("MonthName"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        CallExpression {
                                            Identifier ("Month"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    IdentifierExpression {
                                                        PeriodOperator,
                                                    },
                                                },
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                    CallStatement {
                        Identifier ("ReportDate"),
                        RightParenthesis,
                        Comma,
                        Whitespace,
                        TrueKeyword,
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                WithKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn monthname_select_case() {
        let source = r#"
Select Case MonthName(Month(Date), True)
    Case "Jan", "Feb", "Mar"
        MsgBox "Q1"
    Case "Apr", "May", "Jun"
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
                    Identifier ("MonthName"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Month"),
                                LeftParenthesis,
                                ArgumentList {
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
                            BooleanLiteralExpression {
                                TrueKeyword,
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
                    StringLiteral ("\"Jan\""),
                    Comma,
                    Whitespace,
                    StringLiteral ("\"Feb\""),
                    Comma,
                    Whitespace,
                    StringLiteral ("\"Mar\""),
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
                    StringLiteral ("\"Apr\""),
                    Comma,
                    Whitespace,
                    StringLiteral ("\"May\""),
                    Comma,
                    Whitespace,
                    StringLiteral ("\"Jun\""),
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
    fn monthname_elseif() {
        let source = r#"
If x > 0 Then
    y = 1
ElseIf MonthName(m) = "December" Then
    y = 2
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
                    IdentifierExpression {
                        Identifier ("x"),
                    },
                    Whitespace,
                    GreaterThanOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("0"),
                    },
                },
                Whitespace,
                ThenKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("y"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("1"),
                        },
                        Newline,
                    },
                },
                ElseIfClause {
                    ElseIfKeyword,
                    Whitespace,
                    BinaryExpression {
                        CallExpression {
                            Identifier ("MonthName"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("m"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"December\""),
                        },
                    },
                    Whitespace,
                    ThenKeyword,
                    Newline,
                    StatementList {
                        Whitespace,
                        AssignmentStatement {
                            IdentifierExpression {
                                Identifier ("y"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("2"),
                            },
                            Newline,
                        },
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
    fn monthname_parentheses() {
        let source = r"
Dim name As String
name = (MonthName(6, False))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                NameKeyword,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            NameStatement {
                NameKeyword,
                Whitespace,
                EqualityOperator,
                Whitespace,
                LeftParenthesis,
                Identifier ("MonthName"),
                LeftParenthesis,
                IntegerLiteral ("6"),
                Comma,
                Whitespace,
                FalseKeyword,
                RightParenthesis,
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn monthname_iif() {
        let source = r"
Dim display As String
display = IIf(useShort, MonthName(m, True), MonthName(m, False))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("display"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("display"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("IIf"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("useShort"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            CallExpression {
                                Identifier ("MonthName"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("m"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        BooleanLiteralExpression {
                                            TrueKeyword,
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            CallExpression {
                                Identifier ("MonthName"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("m"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        BooleanLiteralExpression {
                                            FalseKeyword,
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
        ]);
    }

    #[test]
    fn monthname_class_usage() {
        let source = r"
Private m_monthName As String

Public Sub UpdateMonth()
    m_monthName = MonthName(Month(Now), True)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                PrivateKeyword,
                Whitespace,
                Identifier ("m_monthName"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            Newline,
            SubStatement {
                PublicKeyword,
                Whitespace,
                SubKeyword,
                Whitespace,
                Identifier ("UpdateMonth"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("m_monthName"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("MonthName"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    CallExpression {
                                        Identifier ("Month"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("Now"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    BooleanLiteralExpression {
                                        TrueKeyword,
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
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn monthname_function_argument() {
        let source = r"
Call DisplayMonth(MonthName(m, True))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                CallKeyword,
                Whitespace,
                Identifier ("DisplayMonth"),
                LeftParenthesis,
                Identifier ("MonthName"),
                LeftParenthesis,
                Identifier ("m"),
                Comma,
                Whitespace,
                TrueKeyword,
                RightParenthesis,
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn monthname_property_assignment() {
        let source = r"
Set obj = New Calendar
obj.CurrentMonth = MonthName(Month(Date))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SetStatement {
                SetKeyword,
                Whitespace,
                Identifier ("obj"),
                Whitespace,
                EqualityOperator,
                Whitespace,
                NewKeyword,
                Whitespace,
                Identifier ("Calendar"),
                Newline,
            },
            AssignmentStatement {
                MemberAccessExpression {
                    Identifier ("obj"),
                    PeriodOperator,
                    Identifier ("CurrentMonth"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("MonthName"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Month"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            DateKeyword,
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
        ]);
    }

    #[test]
    fn monthname_array_assignment() {
        let source = r"
Dim monthNames(12) As String
Dim i As Integer
monthNames(i) = MonthName(i, True)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("monthNames"),
                LeftParenthesis,
                NumericLiteralExpression {
                    IntegerLiteral ("12"),
                },
                RightParenthesis,
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("i"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
            AssignmentStatement {
                CallExpression {
                    Identifier ("monthNames"),
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
                    Identifier ("MonthName"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            IdentifierExpression {
                                Identifier ("i"),
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            BooleanLiteralExpression {
                                TrueKeyword,
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
    fn monthname_for_loop() {
        let source = r"
Dim i As Integer
For i = 1 To 12
    cboMonth.AddItem MonthName(i)
Next i
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("i"),
                Whitespace,
                AsKeyword,
                Whitespace,
                IntegerKeyword,
                Newline,
            },
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
                NumericLiteralExpression {
                    IntegerLiteral ("12"),
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("cboMonth"),
                        PeriodOperator,
                        Identifier ("AddItem"),
                        Whitespace,
                        Identifier ("MonthName"),
                        LeftParenthesis,
                        Identifier ("i"),
                        RightParenthesis,
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
    fn monthname_while_wend() {
        let source = r"
While m <= 12
    Debug.Print MonthName(m)
    m = m + 1
Wend
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            WhileStatement {
                WhileKeyword,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("m"),
                    },
                    Whitespace,
                    LessThanOrEqualOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("12"),
                    },
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("MonthName"),
                        LeftParenthesis,
                        Identifier ("m"),
                        RightParenthesis,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("m"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("m"),
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
                        },
                        Newline,
                    },
                },
                WendKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn monthname_do_while() {
        let source = r"
Do While i < 12
    months(i) = MonthName(i)
    i = i + 1
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DoStatement {
                DoKeyword,
                Whitespace,
                WhileKeyword,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("i"),
                    },
                    Whitespace,
                    LessThanOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("12"),
                    },
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
                            Identifier ("MonthName"),
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
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("i"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("i"),
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
                        },
                        Newline,
                    },
                },
                LoopKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn monthname_do_until() {
        let source = r"
Do Until i > 12
    list.AddItem MonthName(i, True)
    i = i + 1
Loop
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DoStatement {
                DoKeyword,
                Whitespace,
                UntilKeyword,
                Whitespace,
                BinaryExpression {
                    IdentifierExpression {
                        Identifier ("i"),
                    },
                    Whitespace,
                    GreaterThanOperator,
                    Whitespace,
                    NumericLiteralExpression {
                        IntegerLiteral ("12"),
                    },
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("list"),
                        PeriodOperator,
                        Identifier ("AddItem"),
                        Whitespace,
                        Identifier ("MonthName"),
                        LeftParenthesis,
                        Identifier ("i"),
                        Comma,
                        Whitespace,
                        TrueKeyword,
                        RightParenthesis,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("i"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("i"),
                            },
                            Whitespace,
                            AdditionOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
                        },
                        Newline,
                    },
                },
                LoopKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn monthname_msgbox() {
        let source = r#"
MsgBox "Current month: " & MonthName(Month(Now))
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("MsgBox"),
                Whitespace,
                StringLiteral ("\"Current month: \""),
                Whitespace,
                Ampersand,
                Whitespace,
                Identifier ("MonthName"),
                LeftParenthesis,
                Identifier ("Month"),
                LeftParenthesis,
                Identifier ("Now"),
                RightParenthesis,
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn monthname_comparison() {
        let source = r#"
If MonthName(m1) = MonthName(m2) Then
    MsgBox "Same month name"
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
                        Identifier ("MonthName"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("m1"),
                                },
                            },
                        },
                        RightParenthesis,
                    },
                    Whitespace,
                    EqualityOperator,
                    Whitespace,
                    CallExpression {
                        Identifier ("MonthName"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("m2"),
                                },
                            },
                        },
                        RightParenthesis,
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
                        StringLiteral ("\"Same month name\""),
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
    fn monthname_ucase() {
        let source = r"
Dim upper As String
upper = UCase(MonthName(3, True))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("upper"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("upper"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("UCase"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("MonthName"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        NumericLiteralExpression {
                                            IntegerLiteral ("3"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        BooleanLiteralExpression {
                                            TrueKeyword,
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
        ]);
    }

    #[test]
    fn monthname_label_caption() {
        let source = r"
lblMonth.Caption = MonthName(Month(Date))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            AssignmentStatement {
                MemberAccessExpression {
                    Identifier ("lblMonth"),
                    PeriodOperator,
                    Identifier ("Caption"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("MonthName"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("Month"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            DateKeyword,
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
        ]);
    }

    #[test]
    fn monthname_combo_additem() {
        let source = r"
cboMonths.AddItem MonthName(i, False)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            CallStatement {
                Identifier ("cboMonths"),
                PeriodOperator,
                Identifier ("AddItem"),
                Whitespace,
                Identifier ("MonthName"),
                LeftParenthesis,
                Identifier ("i"),
                Comma,
                Whitespace,
                FalseKeyword,
                RightParenthesis,
                Newline,
            },
        ]);
    }

    #[test]
    fn monthname_format_string() {
        let source = r#"
Dim formatted As String
formatted = MonthName(Month(d)) & " " & Format(Day(d), "00") & ", " & Year(d)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("formatted"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("formatted"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                BinaryExpression {
                    BinaryExpression {
                        BinaryExpression {
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("MonthName"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            CallExpression {
                                                Identifier ("Month"),
                                                LeftParenthesis,
                                                ArgumentList {
                                                    Argument {
                                                        IdentifierExpression {
                                                            Identifier ("d"),
                                                        },
                                                    },
                                                },
                                                RightParenthesis,
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                StringLiteralExpression {
                                    StringLiteral ("\" \""),
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("Format"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        CallExpression {
                                            Identifier ("Day"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    IdentifierExpression {
                                                        Identifier ("d"),
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
                        },
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\", \""),
                        },
                    },
                    Whitespace,
                    Ampersand,
                    Whitespace,
                    CallExpression {
                        Identifier ("Year"),
                        LeftParenthesis,
                        ArgumentList {
                            Argument {
                                IdentifierExpression {
                                    Identifier ("d"),
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
    fn monthname_left_function() {
        let source = r"
Dim firstLetter As String
firstLetter = Left(MonthName(m), 1)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            DimStatement {
                DimKeyword,
                Whitespace,
                Identifier ("firstLetter"),
                Whitespace,
                AsKeyword,
                Whitespace,
                StringKeyword,
                Newline,
            },
            AssignmentStatement {
                IdentifierExpression {
                    Identifier ("firstLetter"),
                },
                Whitespace,
                EqualityOperator,
                Whitespace,
                CallExpression {
                    Identifier ("Left"),
                    LeftParenthesis,
                    ArgumentList {
                        Argument {
                            CallExpression {
                                Identifier ("MonthName"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("m"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Comma,
                        Whitespace,
                        Argument {
                            NumericLiteralExpression {
                                IntegerLiteral ("1"),
                            },
                        },
                    },
                    RightParenthesis,
                },
                Newline,
            },
        ]);
    }
}
