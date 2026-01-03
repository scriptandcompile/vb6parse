//! # `Format$` Function
//!
//! Returns a `String` formatted according to instructions contained in a format expression.
//!
//! ## Syntax
//!
//! ```vb6
//! Format$(expression[, format[, firstdayofweek[, firstweekofyear]]])
//! ```
//!
//! ## Parameters
//!
//! - `expression`: Required. Any valid expression.
//! - `format`: Optional. A valid named or user-defined format expression.
//! - `firstdayofweek`: Optional. A constant that specifies the first day of the week.
//! - `firstweekofyear`: Optional. A constant that specifies the first week of the year.
//!
//! ## Return Value
//!
//! Returns a `String` containing the formatted representation of the expression. If `format` is omitted, `Format$` returns a string similar to `Str$`.
//!
//! ## Remarks
//!
//! The `Format$` function is one of the most versatile functions in VB6, allowing you to format numbers, dates, times, and strings according to predefined or custom format expressions.
//!
//! You can use one of the predefined named formats or create user-defined formats with special characters that specify how the value should be displayed.
//!
//! ### Named Numeric Formats
//! - `General Number`: Display number with no thousand separator
//! - `Currency`: Display number with thousand separator and two decimal places
//! - `Fixed`: Display at least one digit to the left and two digits to the right of decimal
//! - `Standard`: Display number with thousand separator
//! - `Percent`: Display number multiplied by 100 with percent sign
//! - `Scientific`: Use standard scientific notation
//! - `Yes/No`: Display No if number is 0; otherwise display Yes
//! - `True/False`: Display False if number is 0; otherwise display True
//! - `On/Off`: Display Off if number is 0; otherwise display On
//!
//! ### Named Date/Time Formats
//! - `General Date`: Display date and/or time
//! - `Long Date`: Display date according to long date format
//! - `Medium Date`: Display date using medium date format
//! - `Short Date`: Display date using short date format
//! - `Long Time`: Display time using long time format (includes hours, minutes, seconds)
//! - `Medium Time`: Display time in 12-hour format using hours and minutes and AM/PM
//! - `Short Time`: Display time using 24-hour format (hh:mm)
//!
//! ### User-Defined Number Format Characters
//! - `0`: Digit placeholder. Display digit or zero
//! - `#`: Digit placeholder. Display digit or nothing
//! - `.`: Decimal placeholder
//! - `%`: Percentage placeholder
//! - `,`: Thousand separator
//! - `E- E+ e- e+`: Scientific notation
//! - `- + $ ( )`: Display literal character
//! - `\`: Display next character as literal
//!
//! ### User-Defined Date/Time Format Characters
//! - `c`: Display date as `ddddd` and time as `ttttt`
//! - `d`: Display day as number without leading zero (1-31)
//! - `dd`: Display day as number with leading zero (01-31)
//! - `ddd`: Display day as abbreviation (Sun-Sat)
//! - `dddd`: Display day as full name (Sunday-Saturday)
//! - `m`: Display month as number without leading zero (1-12)
//! - `mm`: Display month as number with leading zero (01-12)
//! - `mmm`: Display month as abbreviation (Jan-Dec)
//! - `mmmm`: Display month as full name (January-December)
//! - `yy`: Display year as 2-digit number (00-99)
//! - `yyyy`: Display year as 4-digit number (100-9999)
//! - `h`: Display hour as number without leading zero (0-23)
//! - `hh`: Display hour as number with leading zero (00-23)
//! - `n`: Display minute as number without leading zero (0-59)
//! - `nn`: Display minute as number with leading zero (00-59)
//! - `s`: Display second as number without leading zero (0-59)
//! - `ss`: Display second as number with leading zero (00-59)
//! - `AM/PM`: Use 12-hour clock and display uppercase AM/PM
//!
//! ### User-Defined String Format Characters
//! - `@`: Character placeholder. Display character or space
//! - `&`: Character placeholder. Display character or nothing
//! - `<`: Force lowercase
//! - `>`: Force uppercase
//!
//! ## Typical Uses
//!
//! ### Example 1: Formatting Currency
//! ```vb6
//! Dim amount As Double
//! amount = 1234.56
//! Text1.Text = Format$(amount, "Currency")  ' "$1,234.56"
//! ```
//!
//! ### Example 2: Custom Number Format
//! ```vb6
//! Dim value As Double
//! value = 1234.5
//! result = Format$(value, "0000.00")  ' "1234.50"
//! ```
//!
//! ### Example 3: Date Formatting
//! ```vb6
//! Dim today As Date
//! today = Now
//! dateStr = Format$(today, "Long Date")
//! ```
//!
//! ### Example 4: Custom Date Format
//! ```vb6
//! dateStr = Format$(Now, "yyyy-mm-dd")  ' "2024-01-15"
//! ```
//!
//! ## Common Usage Patterns
//!
//! ### Formatting as Percentage
//! ```vb6
//! Dim rate As Double
//! rate = 0.075
//! display = Format$(rate, "0.00%")  ' "7.50%"
//! ```
//!
//! ### Zero-Padded Numbers
//! ```vb6
//! Dim id As Integer
//! id = 42
//! idStr = Format$(id, "000000")  ' "000042"
//! ```
//!
//! ### Phone Number Formatting
//! ```vb6
//! Dim phone As String
//! phone = "5551234567"
//! formatted = Format$(phone, "(@@@) @@@-@@@@")  ' "(555) 123-4567"
//! ```
//!
//! ### Time Formatting
//! ```vb6
//! Dim currentTime As Date
//! currentTime = Now
//! timeStr = Format$(currentTime, "hh:nn:ss AM/PM")
//! ```
//!
//! ### Scientific Notation
//! ```vb6
//! Dim bigNum As Double
//! bigNum = 12345678
//! sciStr = Format$(bigNum, "0.00E+00")  ' "1.23E+07"
//! ```
//!
//! ### File Timestamp
//! ```vb6
//! filename = "backup_" & Format$(Now, "yyyymmdd_hhnnss") & ".dat"
//! ```
//!
//! ### Accounting Format
//! ```vb6
//! balance = Format$(amount, "#,##0.00;(#,##0.00)")
//! ' Positive: "1,234.56"
//! ' Negative: "(1,234.56)"
//! ```
//!
//! ### Leading Zeros for Dates
//! ```vb6
//! monthStr = Format$(Month(Date), "00")  ' "01" to "12"
//! dayStr = Format$(Day(Date), "00")      ' "01" to "31"
//! ```
//!
//! ### Conditional Formatting
//! ```vb6
//! ' Format: positive;negative;zero
//! result = Format$(value, "+0.00;-0.00;Zero")
//! ```
//!
//! ### Uppercase/Lowercase Conversion
//! ```vb6
//! upperName = Format$("john doe", ">")      ' "JOHN DOE"
//! lowerName = Format$("JOHN DOE", "<")      ' "john doe"
//! ```
//!
//! ## Related Functions
//!
//! - `Format`: Variant version of `Format$`
//! - `Str$`: Converts a number to a string
//! - `CStr`: Converts an expression to a string
//! - `FormatNumber`: Formats a number with specific options
//! - `FormatCurrency`: Formats a number as currency
//! - `FormatDateTime`: Formats a date/time value
//! - `FormatPercent`: Formats a number as a percentage
//!
//! ## Best Practices
//!
//! 1. Use named formats for common formatting tasks (clearer intent)
//! 2. Cache format strings if using the same format repeatedly
//! 3. Test custom format strings with edge cases (zero, negative, very large/small)
//! 4. Use `@` instead of `&` in string formats when you want spaces preserved
//! 5. Remember that `m` vs `mm` depends on context (month vs minute)
//! 6. Use four-digit years (`yyyy`) to avoid Y2K-style issues
//! 7. Consider locale settings when using named formats
//! 8. Use semicolons to specify different formats for positive, negative, and zero
//! 9. Escape literal characters with backslash or quotes when needed
//! 10. Be aware that `Format$` returns a string - convert back if needed
//!
//! ## Performance Considerations
//!
//! - Named formats are slightly faster than complex user-defined formats
//! - Avoid calling `Format$` in tight loops if possible (cache results)
//! - For simple zero-padding, `String$` + `Right$` may be faster
//! - `Format$` is slower than simple string concatenation
//! - Consider using `FormatNumber`, `FormatCurrency`, etc. for specific tasks
//!
//! ## Locale Considerations
//!
//! | Aspect | Behavior |
//! |--------|----------|
//! | Currency Symbol | Uses system locale currency symbol |
//! | Decimal Separator | Uses locale decimal separator (. or ,) |
//! | Thousand Separator | Uses locale thousand separator |
//! | Date Format | Named date formats use locale settings |
//! | Day/Month Names | Uses locale language for names |
//! | AM/PM Designators | Uses locale AM/PM strings |
//! | First Day of Week | Can be overridden with parameter |
//! | First Week of Year | Can be overridden with parameter |
//!
//! ## Common Pitfalls
//!
//! - Using `m` for minutes instead of `n` (m means month)
//! - Forgetting that `Format$` always returns a string
//! - Not escaping literal characters in format strings
//! - Assuming `#` and `0` behave the same (they don't)
//! - Using comma as decimal separator in code (always use period)
//! - Not handling empty strings or null values
//! - Forgetting that format strings are case-sensitive
//! - Using named formats that don't exist (causes error)
//!
//! ## Limitations
//!
//! - Cannot create truly custom named formats
//! - Limited control over locale-specific formatting
//! - No built-in format for ISO 8601 dates (must use `yyyy-mm-ddThh:nn:ss`)
//! - Cannot format arrays or objects directly
//! - Some format combinations may produce unexpected results
//! - Maximum string length limitations apply to output
//! - Cannot use for binary or hexadecimal display (use `Hex$` or `Oct$`)

#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn format_dollar_simple() {
        let source = r#"
Sub Main()
    result = Format$(123.45, "Currency")
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
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Format$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        SingleLiteral,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Currency\""),
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
    fn format_dollar_assignment() {
        let source = r#"
Sub Main()
    Dim formatted As String
    formatted = Format$(Now, "yyyy-mm-dd")
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
                                        Identifier ("Now"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"yyyy-mm-dd\""),
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
    fn format_dollar_number_format() {
        let source = r#"
Sub Main()
    numStr = Format$(1234.5, "0000.00")
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
                            Identifier ("numStr"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Format$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        SingleLiteral,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"0000.00\""),
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
    fn format_dollar_percentage() {
        let source = r#"
Sub Main()
    Dim rate As Double
    rate = 0.075
    display = Format$(rate, "0.00%")
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
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("rate"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("rate"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            SingleLiteral,
                        },
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("display"),
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
                                        Identifier ("rate"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"0.00%\""),
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
    fn format_dollar_zero_padding() {
        let source = r#"
Sub Main()
    idStr = Format$(42, "000000")
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
                            Identifier ("idStr"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Format$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("42"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"000000\""),
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
    fn format_dollar_date_long() {
        let source = r#"
Sub Main()
    dateStr = Format$(Date, "Long Date")
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
                            Identifier ("dateStr"),
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
                                        DateKeyword,
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
    fn format_dollar_time_custom() {
        let source = r#"
Sub Main()
    timeStr = Format$(Now, "hh:nn:ss AM/PM")
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
                            Identifier ("timeStr"),
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
                                        Identifier ("Now"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"hh:nn:ss AM/PM\""),
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
    fn format_dollar_filename() {
        let source = r#"
Sub Main()
    filename = "backup_" & Format$(Now, "yyyymmdd_hhnnss") & ".dat"
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
                            Identifier ("filename"),
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
                                CallExpression {
                                    Identifier ("Format$"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("Now"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            StringLiteralExpression {
                                                StringLiteral ("\"yyyymmdd_hhnnss\""),
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
                                StringLiteral ("\".dat\""),
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
    fn format_dollar_scientific() {
        let source = r#"
Sub Main()
    sciStr = Format$(12345678, "0.00E+00")
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
                            Identifier ("sciStr"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Format$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("12345678"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"0.00E+00\""),
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
    fn format_dollar_accounting() {
        let source = r##"
Sub Main()
    balance = Format$(amount, "#,##0.00;(#,##0.00)")
End Sub
"##;
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
                            Identifier ("balance"),
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
                                        Identifier ("amount"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"#,##0.00;(#,##0.00)\""),
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
    fn format_dollar_uppercase() {
        let source = r#"
Sub Main()
    upperName = Format$("john doe", ">")
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
                            Identifier ("upperName"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Format$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"john doe\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\">\""),
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
    fn format_dollar_lowercase() {
        let source = r#"
Sub Main()
    lowerName = Format$("JOHN DOE", "<")
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
                            Identifier ("lowerName"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Format$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"JOHN DOE\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"<\""),
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
    fn format_dollar_in_condition() {
        let source = r#"
Sub Main()
    If Format$(value, "0.00") = "0.00" Then
        Debug.Print "Zero value"
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
                                Identifier ("Format$"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("value"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        StringLiteralExpression {
                                            StringLiteral ("\"0.00\""),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            EqualityOperator,
                            Whitespace,
                            StringLiteralExpression {
                                StringLiteral ("\"0.00\""),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                StringLiteral ("\"Zero value\""),
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
    fn format_dollar_multiple_uses() {
        let source = r#"
Sub Main()
    d = Format$(Date, "yyyy-mm-dd")
    t = Format$(Time, "hh:nn:ss")
    dt = d & " " & t
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
                            Identifier ("d"),
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
                                        DateKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"yyyy-mm-dd\""),
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
                            Identifier ("t"),
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
                                        TimeKeyword,
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"hh:nn:ss\""),
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
                            Identifier ("dt"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("d"),
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
                                Identifier ("t"),
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
    fn format_dollar_in_function() {
        let source = r#"
Function FormatCurrency(amount As Double) As String
    FormatCurrency = Format$(amount, "Currency")
End Function
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("FormatCurrency"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("amount"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
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
                            Identifier ("FormatCurrency"),
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
                                        Identifier ("amount"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Currency\""),
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
    fn format_dollar_phone_number() {
        let source = r#"
Sub Main()
    phone = "5551234567"
    formatted = Format$(phone, "(@@@) @@@-@@@@")
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
                            Identifier ("phone"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        StringLiteralExpression {
                            StringLiteral ("\"5551234567\""),
                        },
                        Newline,
                    },
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
                                        Identifier ("phone"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"(@@@) @@@-@@@@\""),
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
    fn format_dollar_conditional_format() {
        let source = r#"
Sub Main()
    result = Format$(value, "+0.00;-0.00;Zero")
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
                            Identifier ("result"),
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
                                        Identifier ("value"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"+0.00;-0.00;Zero\""),
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
    fn format_dollar_select_case() {
        let source = r#"
Sub Main()
    formatted = Format$(amount, "Currency")
    Select Case formatted
        Case "$0.00"
            Debug.Print "Empty"
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
                                        Identifier ("amount"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Currency\""),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("formatted"),
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            StringLiteral ("\"$0.00\""),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("Debug"),
                                    PeriodOperator,
                                    PrintKeyword,
                                    Whitespace,
                                    StringLiteral ("\"Empty\""),
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
    fn format_dollar_with_len() {
        let source = r#"
Sub Main()
    str = Format$(123, "000000")
    length = Len(str)
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
                            Identifier ("str"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Format$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("123"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"000000\""),
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
                            Identifier ("length"),
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
                                        Identifier ("str"),
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
    fn format_dollar_expression_arg() {
        let source = r#"
Sub Main()
    result = Format$(x + y, "0.00")
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
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("Format$"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        IdentifierExpression {
                                            Identifier ("x"),
                                        },
                                        Whitespace,
                                        AdditionOperator,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("y"),
                                        },
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"0.00\""),
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
