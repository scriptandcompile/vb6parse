//! # `DateValue` Function
//!
//! Returns a `Variant` (`Date`) containing the date represented by a string expression.
//!
//! ## Syntax
//!
//! ```vb
//! DateValue(date)
//! ```
//!
//! ## Parameters
//!
//! - **date**: Required. `String` expression representing a date from January 1, 100 through
//!   December 31, 9999. Can also be any expression that can represent a date, a time, or
//!   both a date and time, in that range.
//!
//! ## Return Value
//!
//! Returns a `Variant` of subtype `Date`. If the string includes valid time information, it's
//! not returned as part of the date (time is set to midnight). Returns `Null` if the string
//! cannot be converted to a valid date.
//!
//! ## Remarks
//!
//! The `DateValue` function is used to convert string representations of dates into actual
//! `Date` values. It recognizes dates according to the system locale settings.
//!
//! **Important Characteristics:**
//!
//! - Interprets strings according to system locale settings
//! - Recognizes various date formats (MM/DD/YYYY, Month DD, YYYY, etc.)
//! - Strips time information if present (returns date portion only)
//! - Two-digit years: 0-29 → 2000-2029, 30-99 → 1930-1999
//! - Month names can be spelled out or abbreviated
//! - Accepts dates in various formats depending on locale
//! - Returns midnight (00:00:00) for time portion
//! - Case-insensitive for month names
//!
//! ## Recognized Date Formats
//!
//! `DateValue` recognizes many formats (locale-dependent):
//!
//! ```vb
//! ' Numeric formats
//! DateValue("1/15/2025")        ' MM/DD/YYYY (US)
//! DateValue("15/1/2025")        ' DD/MM/YYYY (UK)
//! DateValue("2025-01-15")       ' ISO format
//! DateValue("1-15-2025")        ' With dashes
//!
//! ' Text formats
//! DateValue("January 15, 2025") ' Full month name
//! DateValue("Jan 15, 2025")     ' Abbreviated month
//! DateValue("15 January 2025")  ' Different order
//! DateValue("15-Jan-2025")      ' Mixed format
//!
//! ' Short formats
//! DateValue("1/15/25")          ' Two-digit year
//! DateValue("Jan 15")           ' Assumes current year
//! ```
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Convert string to date
//! Dim birthday As Date
//! birthday = DateValue("5/15/1990")
//! MsgBox birthday
//!
//! ' With month name
//! Dim holiday As Date
//! holiday = DateValue("December 25, 2025")
//!
//! ' ISO format
//! Dim isoDate As Date
//! isoDate = DateValue("2025-01-15")
//! ```
//!
//! ### Parse User Input
//!
//! ```vb
//! Function ParseDate(userInput As String) As Variant
//!     On Error Resume Next
//!     ParseDate = DateValue(userInput)
//!     
//!     If Err.Number <> 0 Then
//!         ParseDate = Null
//!     End If
//! End Function
//!
//! ' Usage
//! Dim inputDate As Variant
//! inputDate = ParseDate(txtDate.Text)
//! If IsNull(inputDate) Then
//!     MsgBox "Invalid date format"
//! End If
//! ```
//!
//! ### Strip Time from `DateTime`
//!
//! ```vb
//! Function GetDateOnly(dateTime As Variant) As Date
//!     ' Convert to string and back to strip time
//!     GetDateOnly = DateValue(CStr(dateTime))
//! End Function
//!
//! ' Alternative using Int
//! Function GetDateOnly2(dateTime As Date) As Date
//!     GetDateOnly2 = Int(dateTime)
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Validate Date String
//!
//! ```vb
//! Function IsValidDateString(dateStr As String) As Boolean
//!     On Error Resume Next
//!     Dim testDate As Date
//!     testDate = DateValue(dateStr)
//!     IsValidDateString = (Err.Number = 0)
//! End Function
//!
//! ' Usage
//! If IsValidDateString(txtDate.Text) Then
//!     MsgBox "Valid date"
//! Else
//!     MsgBox "Invalid date"
//! End If
//! ```
//!
//! ### Parse Various Formats
//!
//! ```vb
//! Function TryParseDateFormats(dateStr As String) As Variant
//!     Dim formats As Variant
//!     Dim i As Integer
//!     
//!     formats = Array("MM/DD/YYYY", "DD/MM/YYYY", "YYYY-MM-DD", _
//!                    "Month DD, YYYY", "DD Month YYYY")
//!     
//!     On Error Resume Next
//!     TryParseDateFormats = DateValue(dateStr)
//!     
//!     If Err.Number = 0 Then Exit Function
//!     
//!     ' Try with current year if not specified
//!     TryParseDateFormats = DateValue(dateStr & " " & Year(Date))
//! End Function
//! ```
//!
//! ### Import Data with Date Parsing
//!
//! ```vb
//! Sub ImportDataWithDates(filePath As String)
//!     Dim line As String
//!     Dim fields() As String
//!     Dim recordDate As Date
//!     
//!     Open filePath For Input As #1
//!     
//!     Do Until EOF(1)
//!         Line Input #1, line
//!         fields = Split(line, ",")
//!         
//!         On Error Resume Next
//!         recordDate = DateValue(fields(0))
//!         
//!         If Err.Number = 0 Then
//!             ' Process valid date record
//!             ProcessRecord recordDate, fields
//!         Else
//!             ' Log invalid date
//!             LogError "Invalid date: " & fields(0)
//!         End If
//!         
//!         On Error GoTo 0
//!     Loop
//!     
//!     Close #1
//! End Sub
//! ```
//!
//! ### Date Range Validator
//!
//! ```vb
//! Function ValidateDateRange(startStr As String, endStr As String) As Boolean
//!     Dim startDate As Date
//!     Dim endDate As Date
//!     
//!     On Error Resume Next
//!     startDate = DateValue(startStr)
//!     If Err.Number <> 0 Then Exit Function
//!     
//!     endDate = DateValue(endStr)
//!     If Err.Number <> 0 Then Exit Function
//!     
//!     ValidateDateRange = (startDate <= endDate)
//! End Function
//! ```
//!
//! ### Convert Text File Dates
//!
//! ```vb
//! Function ConvertTextDate(textDate As String) As Date
//!     ' Convert various text formats to standard date
//!     ConvertTextDate = DateValue(textDate)
//! End Function
//!
//! Sub ProcessLogFile()
//!     Dim logDate As Date
//!     Dim dateStr As String
//!     
//!     dateStr = "Jan 15, 2025"
//!     logDate = ConvertTextDate(dateStr)
//!     
//!     MsgBox Format(logDate, "yyyy-mm-dd")
//! End Sub
//! ```
//!
//! ### Flexible Date Parser
//!
//! ```vb
//! Function SmartDateParse(input As String) As Variant
//!     Dim result As Variant
//!     
//!     On Error Resume Next
//!     
//!     ' Try as-is
//!     result = DateValue(input)
//!     If Err.Number = 0 Then
//!         SmartDateParse = result
//!         Exit Function
//!     End If
//!     Err.Clear
//!     
//!     ' Try adding current year
//!     result = DateValue(input & ", " & Year(Date))
//!     If Err.Number = 0 Then
//!         SmartDateParse = result
//!         Exit Function
//!     End If
//!     
//!     ' Return Null if unparseable
//!     SmartDateParse = Null
//! End Function
//! ```
//!
//! ### Database Date Import
//!
//! ```vb
//! Sub ImportDatabaseDates(rs As ADODB.Recordset)
//!     Dim dateField As String
//!     Dim parsedDate As Date
//!     
//!     Do Until rs.EOF
//!         dateField = rs("DateField").Value
//!         
//!         On Error Resume Next
//!         parsedDate = DateValue(dateField)
//!         
//!         If Err.Number = 0 Then
//!             ' Update with parsed date
//!             rs("DateField") = parsedDate
//!             rs.Update
//!         End If
//!         
//!         rs.MoveNext
//!     Loop
//! End Sub
//! ```
//!
//! ### Form Date Validation
//!
//! ```vb
//! Function ValidateFormDate(ByRef txt As TextBox, fieldName As String) As Boolean
//!     Dim testDate As Date
//!     
//!     On Error Resume Next
//!     testDate = DateValue(txt.Text)
//!     
//!     If Err.Number <> 0 Then
//!         MsgBox "Invalid " & fieldName & " format", vbExclamation
//!         txt.SetFocus
//!         ValidateFormDate = False
//!     Else
//!         ValidateFormDate = True
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Multi-Locale Date Parser
//!
//! ```vb
//! Function ParseInternationalDate(dateStr As String, locale As String) As Variant
//!     ' This is simplified - VB6 doesn't have full locale switching
//!     ' Would need Windows API for true locale switching
//!     
//!     On Error Resume Next
//!     
//!     Select Case UCase(locale)
//!         Case "US"
//!             ' Try MM/DD/YYYY first
//!             ParseInternationalDate = DateValue(dateStr)
//!         
//!         Case "UK", "EU"
//!             ' Parse with assumption of DD/MM/YYYY
//!             ' Would need custom parsing logic
//!             ParseInternationalDate = DateValue(dateStr)
//!         
//!         Case "ISO"
//!             ' YYYY-MM-DD format
//!             ParseInternationalDate = DateValue(dateStr)
//!         
//!         Case Else
//!             ParseInternationalDate = DateValue(dateStr)
//!     End Select
//!     
//!     If Err.Number <> 0 Then
//!         ParseInternationalDate = Null
//!     End If
//! End Function
//! ```
//!
//! ### Batch Date Conversion
//!
//! ```vb
//! Function ConvertDateArray(dateStrings() As String) As Variant
//!     Dim dates() As Date
//!     Dim i As Long
//!     Dim validCount As Long
//!     
//!     ReDim dates(LBound(dateStrings) To UBound(dateStrings))
//!     validCount = 0
//!     
//!     For i = LBound(dateStrings) To UBound(dateStrings)
//!         On Error Resume Next
//!         dates(i) = DateValue(dateStrings(i))
//!         
//!         If Err.Number = 0 Then
//!             validCount = validCount + 1
//!         End If
//!         Err.Clear
//!     Next i
//!     
//!     ConvertDateArray = dates
//! End Function
//! ```
//!
//! ### Date String Normalizer
//!
//! ```vb
//! Function NormalizeDateString(input As String, outputFormat As String) As String
//!     Dim parsedDate As Date
//!     
//!     On Error Resume Next
//!     parsedDate = DateValue(input)
//!     
//!     If Err.Number = 0 Then
//!         NormalizeDateString = Format(parsedDate, outputFormat)
//!     Else
//!         NormalizeDateString = ""
//!     End If
//! End Function
//!
//! ' Usage: Convert various formats to ISO
//! Dim normalized As String
//! normalized = NormalizeDateString("Jan 15, 2025", "yyyy-mm-dd")  ' Returns "2025-01-15"
//! ```
//!
//! ### Excel Date Converter
//!
//! ```vb
//! Function ExcelDateToVBDate(excelDateStr As String) As Variant
//!     ' Excel stores dates as numbers, but when exported may be text
//!     Dim dateVal As Variant
//!     
//!     On Error Resume Next
//!     
//!     ' Try as text date
//!     dateVal = DateValue(excelDateStr)
//!     
//!     If Err.Number <> 0 Then
//!         Err.Clear
//!         ' Try as Excel serial number
//!         If IsNumeric(excelDateStr) Then
//!             dateVal = CDate(CDbl(excelDateStr))
//!         Else
//!             dateVal = Null
//!         End If
//!     End If
//!     
//!     ExcelDateToVBDate = dateVal
//! End Function
//! ```
//!
//! ### Calendar Date Picker Helper
//!
//! ```vb
//! Function ParseCalendarInput(input As String) As Variant
//!     ' Handle various calendar input formats
//!     Dim result As Date
//!     
//!     On Error Resume Next
//!     
//!     ' Remove extra whitespace
//!     input = Trim(input)
//!     
//!     ' Try direct conversion
//!     result = DateValue(input)
//!     If Err.Number = 0 Then
//!         ParseCalendarInput = result
//!         Exit Function
//!     End If
//!     
//!     ' Try common substitutions
//!     If LCase(input) = "today" Then
//!         ParseCalendarInput = Date
//!     ElseIf LCase(input) = "yesterday" Then
//!         ParseCalendarInput = Date - 1
//!     ElseIf LCase(input) = "tomorrow" Then
//!         ParseCalendarInput = Date + 1
//!     Else
//!         ParseCalendarInput = Null
//!     End If
//! End Function
//! ```
//!
//! ### Report Date Filter
//!
//! ```vb
//! Function BuildDateFilter(fromStr As String, toStr As String) As String
//!     Dim fromDate As Date
//!     Dim toDate As Date
//!     
//!     On Error Resume Next
//!     fromDate = DateValue(fromStr)
//!     toDate = DateValue(toStr)
//!     
//!     If Err.Number = 0 Then
//!         BuildDateFilter = "DateField >= #" & fromDate & "# AND DateField <= #" & toDate & "#"
//!     Else
//!         BuildDateFilter = ""
//!     End If
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeDateValue(dateStr As String) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     ' Validate input
//!     If Len(Trim(dateStr)) = 0 Then
//!         SafeDateValue = Null
//!         Exit Function
//!     End If
//!     
//!     SafeDateValue = DateValue(dateStr)
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeDateValue = Null
//! End Function
//!
//! Function SafeDateValueWithDefault(dateStr As String, defaultDate As Date) As Date
//!     On Error Resume Next
//!     SafeDateValueWithDefault = DateValue(dateStr)
//!     
//!     If Err.Number <> 0 Then
//!         SafeDateValueWithDefault = defaultDate
//!     End If
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 13** (Type mismatch): `String` cannot be recognized as a date
//! - **Error 5** (Invalid procedure call): `Date` is outside valid range
//!
//! ## Performance Considerations
//!
//! - `DateValue` involves string parsing, slower than `DateSerial`
//! - Use `DateSerial` when constructing dates from numeric components
//! - Cache parsed dates when processing large datasets
//! - Locale detection adds overhead
//! - For known formats, custom parsing may be faster
//!
//! ## Best Practices
//!
//! ### Always Validate User Input
//!
//! ```vb
//! ' Good - Validate before use
//! On Error Resume Next
//! userDate = DateValue(txtInput.Text)
//! If Err.Number <> 0 Then
//!     MsgBox "Please enter a valid date"
//!     Exit Sub
//! End If
//!
//! ' Avoid - Assuming input is valid
//! userDate = DateValue(txtInput.Text)  ' May crash
//! ```
//!
//! ### Use `IsDate` for Pre-validation
//!
//! ```vb
//! If IsDate(txtInput.Text) Then
//!     processDate = DateValue(txtInput.Text)
//! Else
//!     MsgBox "Invalid date"
//! End If
//! ```
//!
//! ### Prefer `DateSerial` for Programmatic Dates
//!
//! ```vb
//! ' Good - Fast and unambiguous
//! dt = DateSerial(2025, 12, 25)
//!
//! ' Less ideal - String parsing overhead
//! dt = DateValue("12/25/2025")
//! ```
//!
//! ### Be Aware of Locale Issues
//!
//! ```vb
//! ' US locale: MM/DD/YYYY
//! dt = DateValue("3/5/2025")    ' March 5 in US
//!
//! ' UK locale: DD/MM/YYYY
//! dt = DateValue("3/5/2025")    ' May 3 in UK
//!
//! ' Use unambiguous formats when possible
//! dt = DateValue("2025-03-05")  ' ISO format, clearer
//! dt = DateValue("March 5, 2025")  ' Text format, clearer
//! ```
//!
//! ## Comparison with Other Functions
//!
//! ### `DateValue` vs `DateSerial`
//!
//! ```vb
//! ' DateValue - From string representation
//! dt = DateValue("12/25/2025")
//!
//! ' DateSerial - From numeric components (faster, more reliable)
//! dt = DateSerial(2025, 12, 25)
//! ```
//!
//! ### `DateValue` vs `CDate`
//!
//! ```vb
//! ' DateValue - Returns date portion only (strips time)
//! dt = DateValue("12/25/2025 3:30 PM")  ' Returns 12/25/2025 00:00:00
//!
//! ' CDate - Preserves time information
//! dt = CDate("12/25/2025 3:30 PM")      ' Returns 12/25/2025 15:30:00
//! ```
//!
//! ### `DateValue` vs `IsDate`
//!
//! ```vb
//! ' IsDate - Tests if string can be converted (returns Boolean)
//! If IsDate("12/25/2025") Then...
//!
//! ' DateValue - Actually converts (returns Date or error)
//! dt = DateValue("12/25/2025")
//! ```
//!
//! ## Limitations
//!
//! - Locale-dependent interpretation can cause unexpected results
//! - Cannot directly parse custom date formats
//! - Limited control over parsing rules
//! - Strips time information (use `CDate` to preserve time)
//! - Two-digit year interpretation fixed (0-29=2000-2029, 30-99=1930-1999)
//! - Error handling required for user input
//!
//! ## Related Functions
//!
//! - `CDate`: Converts expression to `Date` (preserves time)
//! - `DateSerial`: Creates date from year, month, day (numeric)
//! - `IsDate`: Tests if expression can be converted to date
//! - `Format`: Formats date as `String` (opposite direction)
//! - `TimeValue`: Returns time portion from `String`
//! - `Year`, `Month`, `Day`: Extract date components
//! - `Date`: Returns current system date
//! - `CVDate`: Converts expression to `Date` (legacy function)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn datevalue_basic() {
        let source = r#"
result = DateValue("1/15/2025")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_with_variable() {
        let source = r"
birthday = DateValue(userInput)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_month_name() {
        let source = r#"
holiday = DateValue("December 25, 2025")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_iso_format() {
        let source = r#"
dt = DateValue("2025-01-15")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_in_function() {
        let source = r"
Function ParseDate(input As String) As Date
    ParseDate = DateValue(input)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_with_isdate() {
        let source = r"
If IsDate(txtDate.Text) Then
    result = DateValue(txtDate.Text)
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_in_comparison() {
        let source = r#"
If DateValue(startDate) > Date Then
    MsgBox "Future date"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_with_format() {
        let source = r#"
formatted = Format(DateValue(dateStr), "yyyy-mm-dd")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_error_handling() {
        let source = r#"
On Error Resume Next
result = DateValue(userInput)
If Err.Number <> 0 Then
    MsgBox "Invalid date"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_in_loop() {
        let source = r"
For i = 1 To count
    dates(i) = DateValue(dateStrings(i))
Next i
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_with_trim() {
        let source = r"
cleanDate = DateValue(Trim(input))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_in_select_case() {
        let source = r#"
Select Case DateValue(inputDate)
    Case Date
        MsgBox "Today"
    Case Else
        MsgBox "Other day"
End Select
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_with_cstr() {
        let source = r"
dateOnly = DateValue(CStr(dateTime))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_range_validation() {
        let source = r#"
startDate = DateValue(txtStart.Text)
endDate = DateValue(txtEnd.Text)
If startDate > endDate Then
    MsgBox "Invalid range"
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_with_year() {
        let source = r"
y = Year(DateValue(dateStr))
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_in_array_assignment() {
        let source = r#"
dates(0) = DateValue("1/1/2025")
dates(1) = DateValue("12/31/2025")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_database_field() {
        let source = r#"
rs("DateField") = DateValue(importedDate)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_concatenation() {
        let source = r#"
msg = "Date: " & DateValue(input)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_with_datediff() {
        let source = r#"
days = DateDiff("d", DateValue(start), DateValue(finish))
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_multiple_calls() {
        let source = r"
d1 = DateValue(str1)
d2 = DateValue(str2)
diff = d2 - d1
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_in_msgbox() {
        let source = r#"
MsgBox "Parsed: " & DateValue(userInput)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_textbox_validation() {
        let source = r"
testDate = DateValue(txtDate.Text)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_with_isnull() {
        let source = r"
result = DateValue(input)
If Not IsNull(result) Then
    Process result
End If
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_filter_building() {
        let source = r##"
filter = "Date >= #" & DateValue(startStr) & "#"
"##;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn datevalue_abbreviated_month() {
        let source = r#"
dt = DateValue("Jan 15, 2025")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/library/functions/datetime/datevalue",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
