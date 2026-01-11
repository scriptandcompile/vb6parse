//! # `Format` Function
//!
//! Returns a `Variant` (`String`) containing an expression formatted according to instructions contained in a format expression.
//!
//! ## Syntax
//!
//! ```vb
//! Format(expression[, format[, firstdayofweek[, firstweekofyear]]])
//! ```
//!
//! ## Parameters
//!
//! - **expression**: Required. Any valid expression to be formatted.
//! - **format**: Optional. A valid named or user-defined format expression.
//! - **firstdayofweek**: Optional. A constant that specifies the first day of the week (vbSunday=1, vbMonday=2, etc.).
//! - **firstweekofyear**: Optional. A constant that specifies the first week of the year.
//!
//! ## Return Value
//!
//! Returns a `Variant` of subtype `String` containing the formatted expression. If format is omitted,
//! `Format` returns a string similar to the `Str` function.
//!
//! ## Remarks
//!
//! The `Format` function is one of VB6's most versatile functions for converting values to
//! formatted strings. It supports numeric formatting, date/time formatting, and string formatting
//! with extensive control over output appearance.
//!
//! **Important Characteristics:**
//!
//! - Supports named formats (e.g., "General Number", "Currency", "Short Date")
//! - Supports user-defined format strings with placeholders
//! - Numeric formats: 0, #, ., ,, %, E, etc.
//! - Date/time formats: d, m, y, h, n, s, etc.
//! - String formats: @, &, <, >, !
//! - Returns empty string if expression is Null
//! - Locale-aware (respects regional settings)
//! - Can create custom formats with multiple sections
//! - Supports scientific notation
//! - Handles positive, negative, zero, and null values separately
//!
//! ## Named Formats
//!
//! ### Numeric Named Formats
//! - **General Number**: No thousand separator; displays as entered
//! - **Currency**: Thousand separator, 2 decimal places, currency symbol
//! - **Fixed**: At least one digit left of decimal, 2 digits right
//! - **Standard**: Thousand separator, 2 decimal places
//! - **Percent**: Multiplies by 100, adds percent sign, 2 decimal places
//! - **Scientific**: Standard scientific notation
//!
//! ### Date/Time Named Formats
//! - **General Date**: Short date and time if time ≠ midnight
//! - **Long Date**: Full date (e.g., "Wednesday, January 1, 2025")
//! - **Medium Date**: Date with abbreviated month (e.g., "01-Jan-25")
//! - **Short Date**: Locale-specific short date (e.g., "1/1/2025")
//! - **Long Time**: Full time with seconds (e.g., "1:30:00 PM")
//! - **Medium Time**: Time in 12-hour format (e.g., "1:30 PM")
//! - **Short Time**: Time in 24-hour format (e.g., "13:30")
//!
//! ## Typical Uses
//!
//! - Format numbers with specific decimal places
//! - Add thousand separators to large numbers
//! - Format currency values
//! - Display percentages
//! - Format dates and times for display
//! - Align text in reports
//! - Create fixed-width output
//! - Zero-pad numbers
//! - Convert values to specific string representations
//!
//! ## Examples
//!
//! ### Basic Numeric Formatting
//!
//! ```vb
//! Dim value As Double
//! value = 1234.567
//!
//! ' Named formats
//! Debug.Print Format(value, "General Number")  ' 1234.567
//! Debug.Print Format(value, "Currency")        ' $1,234.57
//! Debug.Print Format(value, "Fixed")           ' 1234.57
//! Debug.Print Format(value, "Standard")        ' 1,234.57
//! Debug.Print Format(value, "Percent")         ' 123456.70%
//! Debug.Print Format(value, "Scientific")      ' 1.23E+03
//!
//! ' Custom formats
//! Debug.Print Format(value, "0.00")            ' 1234.57
//! Debug.Print Format(value, "#,##0.00")        ' 1,234.57
//! Debug.Print Format(value, "000000.00")       ' 001234.57
//! ```
//!
//! ### Date/Time Formatting
//!
//! ```vb
//! Dim dt As Date
//! dt = #1/15/2025 3:45:30 PM#
//!
//! ' Named formats
//! Debug.Print Format(dt, "General Date")       ' 1/15/2025 3:45:30 PM
//! Debug.Print Format(dt, "Long Date")          ' Wednesday, January 15, 2025
//! Debug.Print Format(dt, "Short Date")         ' 1/15/2025
//! Debug.Print Format(dt, "Long Time")          ' 3:45:30 PM
//! Debug.Print Format(dt, "Short Time")         ' 15:45
//!
//! ' Custom formats
//! Debug.Print Format(dt, "yyyy-mm-dd")         ' 2025-01-15
//! Debug.Print Format(dt, "dd/mm/yyyy")         ' 15/01/2025
//! Debug.Print Format(dt, "hh:nn:ss")           ' 03:45:30
//! Debug.Print Format(dt, "dddd, mmmm d, yyyy") ' Wednesday, January 15, 2025
//! ```
//!
//! ### String Formatting
//!
//! ```vb
//! Dim text As String
//! text = "hello"
//!
//! Debug.Print Format(text, ">")                ' HELLO (uppercase)
//! Debug.Print Format(text, "<")                ' hello (lowercase)
//! Debug.Print Format(text, "@@@@@@@@@@")       ' 00000hello (right-aligned)
//! ```
//!
//! ## Common Patterns
//!
//! ### Format Currency with Symbol
//!
//! ```vb
//! Function FormatCurrency(amount As Double) As String
//!     FormatCurrency = Format(amount, "$#,##0.00")
//! End Function
//!
//! ' Usage
//! Debug.Print FormatCurrency(1234.56)  ' $1,234.56
//! Debug.Print FormatCurrency(-500)     ' $-500.00
//! ```
//!
//! ### Zero-Pad Numbers
//!
//! ```vb
//! Function PadWithZeros(num As Long, totalDigits As Integer) As String
//!     Dim formatStr As String
//!     formatStr = String(totalDigits, "0")
//!     PadWithZeros = Format(num, formatStr)
//! End Function
//!
//! ' Usage
//! Debug.Print PadWithZeros(42, 6)      ' 000042
//! Debug.Print PadWithZeros(7, 3)       ' 007
//! ```
//!
//! ### Format File Sizes
//!
//! ```vb
//! Function FormatFileSize(bytes As Long) As String
//!     Const KB = 1024
//!     Const MB = 1048576
//!     Const GB = 1073741824
//!     
//!     If bytes >= GB Then
//!         FormatFileSize = Format(bytes / GB, "0.00") & " GB"
//!     ElseIf bytes >= MB Then
//!         FormatFileSize = Format(bytes / MB, "0.00") & " MB"
//!     ElseIf bytes >= KB Then
//!         FormatFileSize = Format(bytes / KB, "0.00") & " KB"
//!     Else
//!         FormatFileSize = Format(bytes, "#,##0") & " bytes"
//!     End If
//! End Function
//! ```
//!
//! ### Format Phone Numbers
//!
//! ```vb
//! Function FormatPhoneNumber(phoneNum As String) As String
//!     ' Remove non-numeric characters
//!     Dim cleaned As String
//!     Dim i As Long
//!     
//!     For i = 1 To Len(phoneNum)
//!         If IsNumeric(Mid(phoneNum, i, 1)) Then
//!             cleaned = cleaned & Mid(phoneNum, i, 1)
//!         End If
//!     Next i
//!     
//!     ' Format as (XXX) XXX-XXXX
//!     If Len(cleaned) = 10 Then
//!         FormatPhoneNumber = "(" & Left(cleaned, 3) & ") " & _
//!                             Mid(cleaned, 4, 3) & "-" & _
//!                             Right(cleaned, 4)
//!     Else
//!         FormatPhoneNumber = phoneNum
//!     End If
//! End Function
//! ```
//!
//! ### Custom Date Display
//!
//! ```vb
//! Function FormatDateFriendly(dt As Date) As String
//!     Dim daysDiff As Long
//!     daysDiff = DateDiff("d", dt, Date)
//!     
//!     Select Case daysDiff
//!         Case 0
//!             FormatDateFriendly = "Today at " & Format(dt, "h:nn AM/PM")
//!         Case 1
//!             FormatDateFriendly = "Yesterday at " & Format(dt, "h:nn AM/PM")
//!         Case 2 To 6
//!             FormatDateFriendly = Format(dt, "dddd") & " at " & _
//!                                  Format(dt, "h:nn AM/PM")
//!         Case Else
//!             FormatDateFriendly = Format(dt, "mmmm d, yyyy")
//!     End Select
//! End Function
//! ```
//!
//! ### Format Percentages
//!
//! ```vb
//! Function FormatPercent(value As Double, decimals As Integer) As String
//!     Dim formatStr As String
//!     formatStr = "0." & String(decimals, "0") & "%"
//!     FormatPercent = Format(value, formatStr)
//! End Function
//!
//! ' Usage
//! Debug.Print FormatPercent(0.1234, 2)   ' 12.34%
//! Debug.Print FormatPercent(0.5, 0)      ' 50%
//! ```
//!
//! ### Format Decimal Places
//!
//! ```vb
//! Function FormatDecimal(value As Double, places As Integer) As String
//!     Dim formatStr As String
//!     If places = 0 Then
//!         formatStr = "0"
//!     Else
//!         formatStr = "0." & String(places, "0")
//!     End If
//!     FormatDecimal = Format(value, formatStr)
//! End Function
//!
//! ' Usage
//! Debug.Print FormatDecimal(3.14159, 2)  ' 3.14
//! Debug.Print FormatDecimal(100.7, 0)    ' 101
//! ```
//!
//! ### Align Text in Reports
//!
//! ```vb
//! Sub PrintReport()
//!     Debug.Print "Name", "Amount"
//!     Debug.Print String(40, "-")
//!     
//!     Debug.Print "Item 1", Format(1234.56, "$#,##0.00")
//!     Debug.Print "Item 2", Format(78.9, "$#,##0.00")
//!     Debug.Print "Item 3", Format(10000, "$#,##0.00")
//! End Sub
//! ```
//!
//! ### Format Elapsed Time
//!
//! ```vb
//! Function FormatElapsedTime(seconds As Long) As String
//!     Dim hours As Long
//!     Dim minutes As Long
//!     Dim secs As Long
//!     
//!     hours = seconds \ 3600
//!     minutes = (seconds Mod 3600) \ 60
//!     secs = seconds Mod 60
//!     
//!     FormatElapsedTime = Format(hours, "00") & ":" & _
//!                         Format(minutes, "00") & ":" & _
//!                         Format(secs, "00")
//! End Function
//!
//! ' Usage
//! Debug.Print FormatElapsedTime(3661)    ' 01:01:01
//! ```
//!
//! ### Scientific Notation
//!
//! ```vb
//! Function FormatScientific(value As Double, decimals As Integer) As String
//!     Dim formatStr As String
//!     formatStr = "0." & String(decimals, "0") & "E+00"
//!     FormatScientific = Format(value, formatStr)
//! End Function
//!
//! ' Usage
//! Debug.Print FormatScientific(1234567, 2)  ' 1.23E+06
//! ```
//!
//! ### Conditional Formatting
//!
//! ```vb
//! Function FormatNumber(value As Double) As String
//!     ' Format: positive;negative;zero;null
//!     FormatNumber = Format(value, "$#,##0.00;($#,##0.00);-;N/A")
//! End Function
//!
//! ' Usage
//! Debug.Print FormatNumber(1234.5)    ' $1,234.50
//! Debug.Print FormatNumber(-500)      ' ($500.00)
//! Debug.Print FormatNumber(0)         ' -
//! ```
//!
//! ## Advanced Usage
//!
//! ### Multi-Section Format Strings
//!
//! ```vb
//! ' Format: positive;negative;zero;null
//! Function FormatWithSigns(value As Variant) As String
//!     If IsNull(value) Then
//!         FormatWithSigns = "N/A"
//!     Else
//!         FormatWithSigns = Format(value, "+#,##0.00;-#,##0.00;Zero")
//!     End If
//! End Function
//!
//! ' Color-coded text (for display in rich text)
//! Function FormatColorCoded(value As Double) As String
//!     ' Positive=green, Negative=red, Zero=black
//!     If value > 0 Then
//!         FormatColorCoded = "[Green]" & Format(value, "$#,##0.00")
//!     ElseIf value < 0 Then
//!         FormatColorCoded = "[Red]" & Format(value, "($#,##0.00)")
//!     Else
//!         FormatColorCoded = "[Black]$0.00"
//!     End If
//! End Function
//! ```
//!
//! ### Invoice Number Formatting
//!
//! ```vb
//! Function FormatInvoiceNumber(invoiceNum As Long, prefix As String) As String
//!     FormatInvoiceNumber = prefix & "-" & Format(invoiceNum, "000000")
//! End Function
//!
//! ' Usage
//! Debug.Print FormatInvoiceNumber(123, "INV")    ' INV-000123
//! Debug.Print FormatInvoiceNumber(45678, "PO")   ' PO-045678
//! ```
//!
//! ### Custom Number Format Builder
//!
//! ```vb
//! Function BuildNumberFormat(decimals As Integer, _
//!                            Optional useSeparator As Boolean = True, _
//!                            Optional prefix As String = "", _
//!                            Optional suffix As String = "") As String
//!     Dim formatStr As String
//!     
//!     ' Build integer part
//!     If useSeparator Then
//!         formatStr = "#,##0"
//!     Else
//!         formatStr = "0"
//!     End If
//!     
//!     ' Add decimal part
//!     If decimals > 0 Then
//!         formatStr = formatStr & "." & String(decimals, "0")
//!     End If
//!     
//!     ' Add prefix/suffix
//!     formatStr = prefix & formatStr & suffix
//!     
//!     BuildNumberFormat = formatStr
//! End Function
//!
//! ' Usage
//! Debug.Print Format(1234.5, BuildNumberFormat(2, True, "$", ""))  ' $1,234.50
//! ```
//!
//! ### Format for SQL Queries
//!
//! ```vb
//! Function FormatForSQL(value As Variant) As String
//!     Select Case VarType(value)
//!         Case vbString
//!             ' Escape single quotes
//!             FormatForSQL = "'" & Replace(CStr(value), "'", "''") & "'"
//!         Case vbDate
//!             ' Format as ISO date
//!             FormatForSQL = "'" & Format(value, "yyyy-mm-dd hh:nn:ss") & "'"
//!         Case vbNull
//!             FormatForSQL = "NULL"
//!         Case vbBoolean
//!             If CBool(value) Then
//!                 FormatForSQL = "1"
//!             Else
//!                 FormatForSQL = "0"
//!             End If
//!         Case Else
//!             FormatForSQL = CStr(value)
//!     End Select
//! End Function
//! ```
//!
//! ### Dynamic Column Width Formatting
//!
//! ```vb
//! Function FormatColumn(value As Variant, width As Integer, _
//!                       Optional align As String = "left") As String
//!     Dim formatted As String
//!     formatted = CStr(value)
//!     
//!     If Len(formatted) > width Then
//!         formatted = Left(formatted, width - 3) & "..."
//!     Else
//!         Select Case LCase(align)
//!             Case "right"
//!                 formatted = Space(width - Len(formatted)) & formatted
//!             Case "center"
//!                 Dim leftPad As Integer
//!                 leftPad = (width - Len(formatted)) \ 2
//!                 formatted = Space(leftPad) & formatted & _
//!                             Space(width - Len(formatted) - leftPad)
//!             Case Else  ' left
//!                 formatted = formatted & Space(width - Len(formatted))
//!         End Select
//!     End If
//!     
//!     FormatColumn = formatted
//! End Function
//! ```
//!
//! ### Localized Date/Time Formatting
//!
//! ```vb
//! Function FormatLocalizedDate(dt As Date, style As String) As String
//!     ' Uses system locale settings
//!     Select Case LCase(style)
//!         Case "short"
//!             FormatLocalizedDate = Format(dt, "Short Date")
//!         Case "long"
//!             FormatLocalizedDate = Format(dt, "Long Date")
//!         Case "iso"
//!             FormatLocalizedDate = Format(dt, "yyyy-mm-dd")
//!         Case "filename"
//!             ' Safe for filenames
//!             FormatLocalizedDate = Format(dt, "yyyy-mm-dd_hhnnss")
//!         Case Else
//!             FormatLocalizedDate = Format(dt, "General Date")
//!     End Select
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeFormat(value As Variant, formatStr As String) As String
//!     On Error GoTo ErrorHandler
//!     
//!     If IsNull(value) Then
//!         SafeFormat = ""
//!     Else
//!         SafeFormat = Format(value, formatStr)
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 13  ' Type mismatch
//!             Debug.Print "Format error: Invalid data type for format string"
//!             SafeFormat = CStr(value)
//!         Case 5   ' Invalid procedure call
//!             Debug.Print "Format error: Invalid format string"
//!             SafeFormat = CStr(value)
//!         Case Else
//!             Debug.Print "Format error " & Err.Number & ": " & Err.Description
//!             SafeFormat = ""
//!     End Select
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 13** (Type Mismatch): Value type incompatible with format string
//! - **Error 5** (Invalid procedure call): Invalid format string syntax
//! - **Error 6** (Overflow): Value too large for format
//!
//! ## Performance Considerations
//!
//! - Format is relatively fast for simple conversions
//! - Complex format strings take longer to parse
//! - Avoid calling Format in tight loops if possible
//! - Consider caching format strings used repeatedly
//! - String concatenation with Format can be slow
//! - For simple conversions, CStr/Str may be faster
//!
//! ## Best Practices
//!
//! ### Use Named Formats When Possible
//!
//! ```vb
//! ' Good - Clear and locale-aware
//! formatted = Format(amount, "Currency")
//!
//! ' Less portable - Hard-coded currency symbol
//! formatted = Format(amount, "$#,##0.00")
//! ```
//!
//! ### Store Format Strings as Constants
//!
//! ```vb
//! Const FMT_CURRENCY = "$#,##0.00"
//! Const FMT_PERCENT = "0.00%"
//! Const FMT_DATE_ISO = "yyyy-mm-dd"
//!
//! formatted = Format(value, FMT_CURRENCY)
//! ```
//!
//! ### Handle Null Values
//!
//! ```vb
//! ' Good - Check for Null
//! If Not IsNull(value) Then
//!     formatted = Format(value, "0.00")
//! Else
//!     formatted = "N/A"
//! End If
//! ```
//!
//! ## Comparison with Other Functions
//!
//! ### `Format` vs `FormatNumber`
//!
//! ```vb
//! ' Format - More flexible, custom formats
//! result = Format(1234.567, "#,##0.00")
//!
//! ' FormatNumber - Simpler, fewer options
//! result = FormatNumber(1234.567, 2)
//! ```
//!
//! ### `Format` vs `FormatCurrency`
//!
//! ```vb
//! ' Format - Custom currency format
//! result = Format(1234.56, "$#,##0.00")
//!
//! ' FormatCurrency - Uses system locale
//! result = FormatCurrency(1234.56)
//! ```
//!
//! ### `Format` vs `FormatDateTime`
//!
//! ```vb
//! ' Format - Custom date format
//! result = Format(Now, "yyyy-mm-dd hh:nn:ss")
//!
//! ' FormatDateTime - Named formats only
//! result = FormatDateTime(Now, vbShortDate)
//! ```
//!
//! ## Limitations
//!
//! - Format strings are not validated until runtime
//! - Limited regex or pattern matching capabilities
//! - Cannot format arrays or objects directly
//! - No built-in support for custom cultures
//! - Some format characters are locale-dependent
//! - Return value is always `String` (`Variant`)
//!
//! ## Format String Reference
//!
//! ### Numeric Format Characters
//! - **0**: Digit placeholder (displays 0 if no digit)
//! - **#**: Digit placeholder (displays nothing if no digit)
//! - **.**: Decimal placeholder
//! - **,**: Thousand separator
//! - **%**: Percentage placeholder (multiplies by 100)
//! - **E+**, **E-**, **e+**, **e-**: Scientific notation
//!
//! ### Date/Time Format Characters
//! - **d**, **dd**: Day (1-31, 01-31)
//! - **ddd**, **dddd**: Day name (Mon, Monday)
//! - **m**, **mm**: Month (1-12, 01-12)
//! - **mmm**, **mmmm**: Month name (Jan, January)
//! - **yy**, **yyyy**: Year (25, 2025)
//! - **h**, **hh**: Hour (0-23, 00-23)
//! - **n**, **nn**: Minute (0-59, 00-59)
//! - **s**, **ss**: Second (0-59, 00-59)
//! - **AM/PM**, **am/pm**: 12-hour time indicator
//!
//! ## Related Functions
//!
//! - `FormatNumber`: Formats a number with specified decimal places
//! - `FormatCurrency`: Formats a number as currency
//! - `FormatPercent`: Formats a number as percentage
//! - `FormatDateTime`: Formats a date/time value
//! - `Str`: Converts number to string
//! - `CStr`: Converts expression to string

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn format_basic() {
        let source = r#"
result = Format(value, "0.00")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_named() {
        let source = r#"
formatted = Format(amount, "Currency")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_date() {
        let source = r#"
dateStr = Format(Now, "yyyy-mm-dd")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_percent() {
        let source = r#"
pct = Format(value, "0.00%")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_thousands() {
        let source = r##"
formatted = Format(value, "#,##0.00")
"##;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_zero_pad() {
        let source = r#"
padded = Format(num, "000000")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_debug_print() {
        let source = r#"
Debug.Print Format(value, "General Number")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_concatenation() {
        let source = r#"
msg = "Amount: " & Format(total, "$#,##0.00")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_in_function() {
        let source = r#"
Function FormatCurrency(amount As Double) As String
    FormatCurrency = Format(amount, "$#,##0.00")
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_division() {
        let source = r#"
sizeStr = Format(bytes / 1024, "0.00") & " KB"
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_datetime_fields() {
        let source = r#"
timeStr = Format(dt, "h:nn AM/PM")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_scientific() {
        let source = r#"
sci = Format(value, "0.00E+00")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_if_statement() {
        let source = r#"
If value > 0 Then
    result = Format(value, "$#,##0.00")
Else
    result = Format(value, "($#,##0.00)")
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_select_case() {
        let source = r#"
Select Case style
    Case "short"
        result = Format(dt, "Short Date")
    Case "long"
        result = Format(dt, "Long Date")
End Select
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_error_handling() {
        let source = r"
On Error GoTo ErrorHandler
formatted = Format(value, formatStr)
Exit Function
ErrorHandler:
    formatted = CStr(value)
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_isnull_check() {
        let source = r#"
If Not IsNull(value) Then
    formatted = Format(value, "0.00")
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_string_builder() {
        let source = r#"
formatStr = String(totalDigits, "0")
result = Format(num, formatStr)
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_elapsed_time() {
        let source = r#"
result = Format(hours, "00") & ":" & Format(minutes, "00") & ":" & Format(secs, "00")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_uppercase() {
        let source = r#"
upper = Format(text, ">")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_lowercase() {
        let source = r#"
lower = Format(text, "<")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_invoice() {
        let source = r#"
invoiceNum = prefix & "-" & Format(num, "000000")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_conditional() {
        let source = r#"
result = Format(value, "$#,##0.00;($#,##0.00);-;N/A")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_for_loop() {
        let source = r#"
For i = 1 To 10
    Debug.Print Format(i, "00")
Next i
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_assignment() {
        let source = r#"
lstBox.AddItem Format(items(i), "$#,##0.00")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_comparison() {
        let source = r#"
If Format(dt, "yyyy-mm-dd") = targetDate Then
    found = True
End If
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn format_multiline() {
        let source = r#"
report = "Total: " & Format(total, "$#,##0.00") & vbCrLf & _
         "Tax: " & Format(tax, "$#,##0.00") & vbCrLf & _
         "Grand Total: " & Format(total + tax, "$#,##0.00")
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/string/format");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}
