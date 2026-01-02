//! # `Date$` Function
//!
//! Returns the current system date as a `String`. The dollar sign suffix (`$`) explicitly
//! indicates that this function returns a `String` type (not a `Variant`).
//!
//! ## Syntax
//!
//! ```vb
//! Date$
//! ```
//!
//! ## Parameters
//!
//! None. The `Date$` function takes no parameters.
//!
//! ## Return Value
//!
//! Returns a `String` containing the current system date. The format depends on the system's
//! regional settings (typically "mm/dd/yyyy" in US or "dd/mm/yyyy" in other regions). The
//! return value is always a `String` type (never `Variant`).
//!
//! ## Remarks
//!
//! - The `Date$` function always returns a `String`, while `Date` (without `$`) returns a `Variant` of subtype `Date`.
//! - Returns only the date portion (no time information).
//! - Uses system date from computer's clock.
//! - Date format depends on system locale/regional settings.
//! - Common formats: "mm/dd/yyyy" (US), "dd/mm/yyyy" (Europe), "yyyy/mm/dd" (ISO).
//! - The string representation may include leading zeros (e.g., "01/05/2025").
//! - For better performance when you need a string, use `Date$` instead of `Date`.
//! - Cannot be used to set the system date (unlike `Date` statement).
//!
//! ## Typical Uses
//!
//! 1. **Date stamping** - Add date stamps to log entries, files, or records
//! 2. **Display formatting** - Show current date to users
//! 3. **File naming** - Include date in filenames
//! 4. **Logging** - Record when events occurred
//! 5. **Report generation** - Add date headers to reports
//! 6. **Audit trails** - Track when data was created or modified
//! 7. **String concatenation** - Combine date with other text
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Get current date as string
//! Dim dateStr As String
//! dateStr = Date$
//! ```
//!
//! ```vb
//! ' Example 2: Display current date
//! MsgBox "Today is: " & Date$
//! ```
//!
//! ```vb
//! ' Example 3: Create date stamp
//! Dim stamp As String
//! stamp = "Report generated on " & Date$
//! ```
//!
//! ```vb
//! ' Example 4: Simple assignment
//! currentDate = Date$
//! ```
//!
//! ## Common Patterns
//!
//! ### File Naming with Date
//! ```vb
//! Function CreateDateStampedFilename(baseName As String) As String
//!     Dim dateStr As String
//!     Dim cleanDate As String
//!     
//!     ' Get date and remove slashes
//!     dateStr = Date$
//!     cleanDate = Replace$(dateStr, "/", "")
//!     
//!     CreateDateStampedFilename = baseName & "_" & cleanDate & ".txt"
//! End Function
//! ```
//!
//! ### Log Entry with Date
//! ```vb
//! Sub WriteLogEntry(message As String)
//!     Dim logFile As Integer
//!     Dim logEntry As String
//!     
//!     logFile = FreeFile
//!     Open "application.log" For Append As #logFile
//!     
//!     logEntry = Date$ & " - " & message
//!     Print #logFile, logEntry
//!     
//!     Close #logFile
//! End Sub
//! ```
//!
//! ### Date-Based Conditional Logic
//! ```vb
//! Sub CheckDate()
//!     Dim todayStr As String
//!     todayStr = Date$
//!     
//!     ' Simple string comparison (locale-dependent)
//!     If todayStr = "12/25/2025" Then
//!         MsgBox "Merry Christmas!"
//!     End If
//! End Sub
//! ```
//!
//! ### Report Header
//! ```vb
//! Function CreateReportHeader(title As String) As String
//!     Dim header As String
//!     header = String$(60, "=") & vbCrLf
//!     header = header & title & vbCrLf
//!     header = header & "Generated: " & Date$ & vbCrLf
//!     header = header & String$(60, "=") & vbCrLf
//!     CreateReportHeader = header
//! End Function
//! ```
//!
//! ### Date Display in Status Bar
//! ```vb
//! Sub UpdateStatusBar()
//!     Form1.StatusBar.Panels(1).Text = "Date: " & Date$
//! End Sub
//! ```
//!
//! ### Backup File Naming
//! ```vb
//! Function GetBackupFilename(originalFile As String) As String
//!     Dim baseName As String
//!     Dim extension As String
//!     Dim dotPos As Integer
//!     Dim dateStr As String
//!     
//!     dotPos = InStrRev(originalFile, ".")
//!     If dotPos > 0 Then
//!         baseName = Left$(originalFile, dotPos - 1)
//!         extension = Mid$(originalFile, dotPos)
//!     Else
//!         baseName = originalFile
//!         extension = ""
//!     End If
//!     
//!     ' Clean date string for filename
//!     dateStr = Replace$(Date$, "/", "-")
//!     
//!     GetBackupFilename = baseName & "_backup_" & dateStr & extension
//! End Function
//! ```
//!
//! ### Daily Log File
//! ```vb
//! Function GetDailyLogFilename() As String
//!     Dim dateStr As String
//!     dateStr = Replace$(Date$, "/", "")
//!     GetDailyLogFilename = "log_" & dateStr & ".txt"
//! End Function
//! ```
//!
//! ### Date Validation (Simple)
//! ```vb
//! Function IsToday(dateStr As String) As Boolean
//!     IsToday = (dateStr = Date$)
//! End Function
//! ```
//!
//! ### Combining Date and Time
//! ```vb
//! Function GetDateTimeStamp() As String
//!     GetDateTimeStamp = Date$ & " " & Time$
//! End Function
//! ```
//!
//! ### Data Export Header
//! ```vb
//! Sub ExportData()
//!     Dim exportFile As Integer
//!     
//!     exportFile = FreeFile
//!     Open "export.csv" For Output As #exportFile
//!     
//!     ' Write header with date
//!     Print #exportFile, "Data Export - " & Date$
//!     Print #exportFile, "Name,Value,Status"
//!     
//!     ' Export data...
//!     
//!     Close #exportFile
//! End Sub
//! ```
//!
//! ## Related Functions
//!
//! - `Date`: Returns current date as `Variant` instead of `String`
//! - `Now`: Returns current date and time
//! - `Time$`: Returns current time as `String`
//! - `Format$`: Formats dates with custom patterns
//! - `Year`: Extracts year from date
//! - `Month`: Extracts month from date
//! - `Day`: Extracts day from date
//! - `DateSerial`: Creates date from year, month, day
//! - `DateValue`: Converts string to date
//!
//! ## Best Practices
//!
//! 1. Use `Format$` instead of `Date$` when you need specific date formats
//! 2. Be aware that `Date$` format depends on system locale settings
//! 3. For file naming, clean the date string (remove or replace slashes)
//! 4. Use `Date$` instead of `Date` when you need a string result
//! 5. For date comparisons, use `Date` (Variant) instead of `Date$` (String)
//! 6. Don't assume a specific date format - it varies by locale
//! 7. For consistent formatting, use `Format$(Date, "yyyy-mm-dd")`
//! 8. Test with different regional settings if your app is international
//! 9. Store dates in consistent format (ISO 8601 recommended)
//! 10. Use `DateValue` to parse date strings reliably
//!
//! ## Performance Considerations
//!
//! - `Date$` is slightly more efficient than `Date` when you need a string
//! - System date/time calls are fast but not free
//! - Cache the result if you need it multiple times in quick succession
//! - For high-frequency logging, consider caching the date string
//!
//! ## Locale Considerations
//!
//! The format of `Date$` varies by system locale:
//!
//! | Locale | Example Format | Sample Output |
//! |--------|----------------|---------------|
//! | US (English) | mm/dd/yyyy | "12/25/2025" |
//! | UK (English) | dd/mm/yyyy | "25/12/2025" |
//! | Germany | dd.mm.yyyy | "25.12.2025" |
//! | Japan | yyyy/mm/dd | "2025/12/25" |
//! | France | dd/mm/yyyy | "25/12/2025" |
//!
//! ## Common Pitfalls
//!
//! 1. **String Comparison**: Comparing `Date$` strings directly is locale-dependent and unreliable
//!    ```vb
//!    ' BAD - locale-dependent
//!    If Date$ = "12/25/2025" Then
//!    
//!    ' GOOD - use Date variants
//!    If Date = #12/25/2025# Then
//!    ```
//!
//! 2. **Date Parsing**: Don't parse `Date$` manually - use `DateValue` instead
//!    ```vb
//!    ' BAD - fragile parsing
//!    parts = Split(Date$, "/")
//!    
//!    ' GOOD - use built-in functions
//!    currentYear = Year(Date)
//!    currentMonth = Month(Date)
//!    ```
//!
//! 3. **Filename Safety**: Date strings may contain invalid filename characters
//!    ```vb
//!    ' BAD - slashes invalid in filenames
//!    filename = "report_" & Date$ & ".txt"
//!    
//!    ' GOOD - replace invalid characters
//!    filename = "report_" & Replace$(Date$, "/", "-") & ".txt"
//!    ```
//!
//! ## Limitations
//!
//! - Cannot be used to set the system date (use `Date` statement for that)
//! - Format is system-dependent and cannot be directly controlled
//! - No time information included (use `Now` or `Time$` for time)
//! - String comparison of dates is unreliable across locales
//! - Cannot specify date format (use `Format$` for custom formats)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn date_dollar_simple() {
        let source = r"
Sub Main()
    dateStr = Date$
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("dateStr"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("Date$"),
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
    fn date_dollar_assignment() {
        let source = r"
Sub Main()
    Dim currentDate As String
    currentDate = Date$
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("currentDate"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("currentDate"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("Date$"),
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
    fn date_dollar_concatenation() {
        let source = r#"
Sub Main()
    stamp = "Report: " & Date$
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("stamp"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"Report: \""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("Date$"),
                            },
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
    fn date_dollar_in_condition() {
        let source = r#"
Sub Main()
    If Date$ = "12/25/2025" Then
        MsgBox "Christmas!"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("Date$"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"12/25/2025\""),
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
                                StringLiteral ("\"Christmas!\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
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
    fn date_dollar_with_replace() {
        let source = r#"
Function GetFilename() As String
    cleanDate = Replace$(Date$, "/", "")
    GetFilename = "file_" & cleanDate & ".txt"
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetFilename"),
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
                            Identifier ("cleanDate"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Replace"),
                            DollarSign,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("Date$"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"/\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"\""),
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
                            Identifier ("GetFilename"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                StringLiteralExpression {
                                    StringLiteral ("\"file_\""),
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("cleanDate"),
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\".txt\""),
                            },
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
    fn date_dollar_log_entry() {
        let source = r#"
Sub WriteLog(message As String)
    logEntry = Date$ & " - " & message
    Print #1, logEntry
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("WriteLog"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("message"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("logEntry"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("Date$"),
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                StringLiteralExpression {
                                    StringLiteral ("\" - \""),
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("message"),
                            },
                        },
                        Newline,
                    },
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        Identifier ("logEntry"),
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
    fn date_dollar_report_header() {
        let source = r#"
Function CreateHeader() As String
    header = "Report" & vbCrLf
    header = header & "Date: " & Date$ & vbCrLf
    CreateHeader = header
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CreateHeader"),
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
                            Identifier ("header"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"Report\""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("vbCrLf"),
                            },
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("header"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("header"),
                                    },
                                    Whitespace,
                                    Ampersand,
                                    Whitespace,
                                    StringLiteralExpression {
                                        StringLiteral ("\"Date: \""),
                                    },
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("Date$"),
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("vbCrLf"),
                            },
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("CreateHeader"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("header"),
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
    fn date_dollar_status_bar() {
        let source = r#"
Sub UpdateStatus()
    StatusBar.Text = "Date: " & Date$
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("UpdateStatus"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        MemberAccessExpression {
                            Identifier ("StatusBar"),
                            PeriodOperator,
                            TextKeyword,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"Date: \""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("Date$"),
                            },
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
    fn date_dollar_backup_filename() {
        let source = r#"
Function GetBackup() As String
    dateStr = Replace$(Date$, "/", "-")
    GetBackup = "backup_" & dateStr & ".bak"
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetBackup"),
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
                            Identifier ("dateStr"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Replace"),
                            DollarSign,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("Date$"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"/\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"-\""),
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
                            Identifier ("GetBackup"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                StringLiteralExpression {
                                    StringLiteral ("\"backup_\""),
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("dateStr"),
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\".bak\""),
                            },
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
    fn date_dollar_daily_log() {
        let source = r#"
Function GetLogFile() As String
    dateStr = Replace$(Date$, "/", "")
    GetLogFile = "log_" & dateStr & ".txt"
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetLogFile"),
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
                            Identifier ("dateStr"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Replace"),
                            DollarSign,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("Date$"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"/\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"\""),
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
                            Identifier ("GetLogFile"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                StringLiteralExpression {
                                    StringLiteral ("\"log_\""),
                                },
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("dateStr"),
                                },
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\".txt\""),
                            },
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
    fn date_dollar_validation() {
        let source = r"
Function IsToday(dateStr As String) As Boolean
    IsToday = (dateStr = Date$)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("IsToday"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("dateStr"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    StringKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                BooleanKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("IsToday"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("dateStr"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("Date$"),
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
    fn date_dollar_with_time() {
        let source = r#"
Function GetTimestamp() As String
    GetTimestamp = Date$ & " " & Time$
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetTimestamp"),
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
                            Identifier ("GetTimestamp"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("Date$"),
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
                            IdentifierExpression {
                                TimeKeyword,
                                DollarSign,
                            },
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
    fn date_dollar_export_header() {
        let source = r#"
Sub ExportData()
    Print #1, "Export - " & Date$
    Print #1, "Name,Value"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ExportData"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Export - \""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Date$"),
                        Newline,
                    },
                    PrintStatement {
                        Whitespace,
                        PrintKeyword,
                        Whitespace,
                        Octothorpe,
                        IntegerLiteral ("1"),
                        Comma,
                        Whitespace,
                        StringLiteral ("\"Name,Value\""),
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
    fn date_dollar_debug_print() {
        let source = r#"
Sub Main()
    Debug.Print "Current date: " & Date$
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        StringLiteral ("\"Current date: \""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("Date$"),
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
    fn date_dollar_multiple_uses() {
        let source = r#"
Sub Main()
    d1 = Date$
    d2 = Date$
    If d1 = d2 Then MsgBox "Same"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("d1"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("Date$"),
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("d2"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("Date$"),
                        },
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("d1"),
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("d2"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Whitespace,
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Same\""),
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
    fn date_dollar_function_call() {
        let source = r"
Function GetCurrentDate() As String
    GetCurrentDate = Date$
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("GetCurrentDate"),
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
                            Identifier ("GetCurrentDate"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("Date$"),
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
    fn date_dollar_with_format() {
        let source = r#"
Sub Main()
    formatted = Format$(Date$, "Long Date")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("formatted"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Format$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("Date$"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Long Date\""),
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
    fn date_dollar_select_case() {
        let source = r#"
Sub Main()
    Select Case Date$
        Case "01/01/2025"
            mode = 1
        Case Else
            mode = 0
    End Select
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("Date$"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\"01/01/2025\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("mode"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseElseClause {
                            CaseKeyword,
                            Whitespace,
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                AssignmentStatement {
                                    IdentifierExpression {
                                        Identifier ("mode"),
                                    },
                                    Whitespace,
                                    EqualityOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("0"),
                                    },
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
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
    fn date_dollar_instr() {
        let source = r#"
Sub Main()
    If InStr(Date$, "/") > 0 Then
        hasSlash = True
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("InStr"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("Date$"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        StringLiteralExpression {
                                            StringLiteral ("\"/\""),
                                        },
                                    },
                                },
                                RightParenthesis,
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
                                    Identifier ("hasSlash"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BooleanLiteralExpression {
                                    TrueKeyword,
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
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
    fn date_dollar_len() {
        let source = r"
Sub Main()
    dateLen = Len(Date$)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Main"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("dateLen"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            LenKeyword,
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("Date$"),
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
}
