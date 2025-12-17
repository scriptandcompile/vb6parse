//! # `UCase$` Function
//!
//! Returns a `String` that has been converted to uppercase.
//! The "$" suffix indicates this function returns a `String` type.
//!
//! ## Syntax
//!
//! ```vb
//! UCase$(string)
//! ```
//!
//! ## Parameters
//!
//! - **string**: Required. Any valid string expression. If `string` contains `Null`, `Null` is returned.
//!
//! ## Returns
//!
//! Returns a `String` with all lowercase letters converted to uppercase. Numbers and punctuation
//! are unchanged.
//!
//! ## Remarks
//!
//! - `UCase$` converts all lowercase letters in a string to uppercase.
//! - The "$" suffix explicitly indicates the function returns a `String` type rather than a `Variant`.
//! - Only lowercase letters (a-z) are affected; uppercase letters and non-alphabetic characters remain unchanged.
//! - `UCase$` is functionally equivalent to `UCase`, but `UCase$` returns a `String` while `UCase` can return a `Variant`.
//! - For better performance when you know the result is a string, use `UCase$`.
//! - If the argument is `Null`, the function returns `Null`.
//! - The conversion is based on the system locale settings.
//! - For international characters, the behavior depends on the current code page.
//! - The inverse function is `LCase$`, which converts strings to lowercase.
//! - Common use cases include display formatting, SQL keywords, and creating constants.
//!
//! ## Typical Uses
//!
//! 1. **Display formatting** - Format text for display in uppercase
//! 2. **SQL keyword generation** - Create SQL queries with uppercase keywords
//! 3. **Constant generation** - Generate uppercase constant names
//! 4. **File path normalization** - Normalize file paths for case-insensitive systems
//! 5. **Acronym formatting** - Format acronyms and abbreviations
//! 6. **Header text** - Create uppercase headers for reports
//! 7. **Code generation** - Generate uppercase identifiers in code
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Simple conversion
//! Dim result As String
//! result = UCase$("hello")  ' Returns "HELLO"
//! ```
//!
//! ```vb
//! ' Example 2: Mixed case
//! Dim text As String
//! text = UCase$("Hello World")  ' Returns "HELLO WORLD"
//! ```
//!
//! ```vb
//! ' Example 3: With numbers and punctuation
//! Dim mixed As String
//! mixed = UCase$("abc123!@#")  ' Returns "ABC123!@#"
//! ```
//!
//! ```vb
//! ' Example 4: Already uppercase
//! Dim upper As String
//! upper = UCase$("ALREADY UPPERCASE")  ' Returns "ALREADY UPPERCASE"
//! ```
//!
//! ## Common Patterns
//!
//! ### SQL Keyword Formatting
//! ```vb
//! Function BuildSQLQuery(table As String, field As String) As String
//!     BuildSQLQuery = UCase$("SELECT") & " * " & UCase$("FROM") & " " & table
//! End Function
//! ```
//!
//! ### Constant Name Generator
//! ```vb
//! Function GenerateConstantName(baseName As String) As String
//!     GenerateConstantName = UCase$(Replace(baseName, " ", "_"))
//! End Function
//! ```
//!
//! ### Acronym Formatter
//! ```vb
//! Function FormatAcronym(text As String) As String
//!     FormatAcronym = UCase$(text)
//! End Function
//! ```
//!
//! ### Header Text Generator
//! ```vb
//! Function CreateHeader(title As String) As String
//!     CreateHeader = String$(Len(title), "=") & vbCrLf & _
//!                    UCase$(title) & vbCrLf & _
//!                    String$(Len(title), "=")
//! End Function
//! ```
//!
//! ### Case-Insensitive Command Comparison
//! ```vb
//! Function ProcessCommand(cmd As String) As Boolean
//!     Select Case UCase$(Trim$(cmd))
//!         Case "START"
//!             ProcessCommand = StartService()
//!         Case "STOP"
//!             ProcessCommand = StopService()
//!         Case "RESTART"
//!             ProcessCommand = RestartService()
//!         Case Else
//!             ProcessCommand = False
//!     End Select
//! End Function
//! ```
//!
//! ### File Extension Normalization
//! ```vb
//! Function NormalizeExtension(filename As String) As String
//!     Dim ext As String
//!     ext = Right$(filename, 4)
//!     If UCase$(ext) = ".TXT" Then
//!         NormalizeExtension = "Text File"
//!     End If
//! End Function
//! ```
//!
//! ### Environment Variable Names
//! ```vb
//! Function GetEnvironmentVar(varName As String) As String
//!     GetEnvironmentVar = Environ$(UCase$(varName))
//! End Function
//! ```
//!
//! ### Registry Key Normalization
//! ```vb
//! Function NormalizeRegistryKey(keyName As String) As String
//!     NormalizeRegistryKey = UCase$(Trim$(keyName))
//! End Function
//! ```
//!
//! ### Display Name Formatting
//! ```vb
//! Function FormatDisplayName(firstName As String, lastName As String) As String
//!     FormatDisplayName = UCase$(lastName) & ", " & firstName
//! End Function
//! ```
//!
//! ### Code Template Generator
//! ```vb
//! Function GenerateEnumMember(memberName As String) As String
//!     GenerateEnumMember = "    " & UCase$(memberName) & " = " & counter
//! End Function
//! ```
//!
//! ## Advanced Examples
//!
//! ### SQL Query Builder with Uppercase Keywords
//! ```vb
//! Function BuildComplexQuery(table As String, fields As String, whereClause As String) As String
//!     Dim sql As String
//!     
//!     sql = UCase$("SELECT") & " " & fields & " "
//!     sql = sql & UCase$("FROM") & " " & table
//!     
//!     If Len(whereClause) > 0 Then
//!         sql = sql & " " & UCase$("WHERE") & " " & whereClause
//!     End If
//!     
//!     BuildComplexQuery = sql
//! End Function
//! ```
//!
//! ### Configuration File Writer
//! ```vb
//! Sub WriteConfigSection(fileNum As Integer, sectionName As String, settings As Collection)
//!     Dim key As Variant
//!     
//!     Print #fileNum, "[" & UCase$(sectionName) & "]"
//!     
//!     For Each key In settings
//!         Print #fileNum, UCase$(key) & "=" & settings(key)
//!     Next key
//!     
//!     Print #fileNum, ""
//! End Sub
//! ```
//!
//! ### Report Header Generator
//! ```vb
//! Function GenerateReportHeader(reportTitle As String, reportDate As String) As String
//!     Dim header As String
//!     Dim separator As String
//!     
//!     separator = String$(60, "=")
//!     header = separator & vbCrLf
//!     header = header & Space$((60 - Len(reportTitle)) \ 2) & UCase$(reportTitle) & vbCrLf
//!     header = header & Space$((60 - Len(reportDate)) \ 2) & reportDate & vbCrLf
//!     header = header & separator & vbCrLf
//!     
//!     GenerateReportHeader = header
//! End Function
//! ```
//!
//! ### Macro Name Validator
//! ```vb
//! Function ValidateMacroName(macroName As String) As String
//!     Dim validName As String
//!     Dim i As Long
//!     Dim char As String
//!     
//!     ' Convert to uppercase and remove invalid characters
//!     validName = UCase$(macroName)
//!     
//!     For i = 1 To Len(validName)
//!         char = Mid$(validName, i, 1)
//!         If (char >= "A" And char <= "Z") Or _
//!            (char >= "0" And char <= "9") Or _
//!            char = "_" Then
//!             ValidateMacroName = ValidateMacroName & char
//!         End If
//!     Next i
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function SafeUCase(text As String) As String
//!     On Error GoTo ErrorHandler
//!     
//!     If IsNull(text) Then
//!         SafeUCase = ""
//!         Exit Function
//!     End If
//!     
//!     SafeUCase = UCase$(text)
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeUCase = ""
//! End Function
//! ```
//!
//! ## Performance Notes
//!
//! - `UCase$` is a fast operation with minimal overhead
//! - For large strings, the performance is linear with string length
//! - `UCase$` (returns `String`) is slightly faster than `UCase` (returns `Variant`)
//! - When formatting multiple strings, consider caching uppercase versions if reused
//! - For very large datasets, consider using database-level text functions
//!
//! ## Best Practices
//!
//! 1. **Use for display** - Convert to uppercase when formatting for display
//! 2. **Prefer `UCase$` over `UCase`** - Use `UCase$` when you know the result is a string
//! 3. **SQL keywords** - Use uppercase for SQL keywords to improve readability
//! 4. **Handle Null** - Check for `Null` values before calling `UCase$`
//! 5. **Combine with Trim** - Often useful to combine `UCase$` with `Trim$` for cleaner output
//! 6. **Document intent** - Make it clear when uppercase conversion is for display vs. comparison
//! 7. **Consider locale** - Be aware that conversion may vary by system locale
//!
//! ## Comparison with Related Functions
//!
//! | Function | Return Type | Conversion | Use Case |
//! |----------|-------------|------------|----------|
//! | `UCase` | Variant | To uppercase | When working with Variant types |
//! | `UCase$` | String | To uppercase | When result is definitely a string |
//! | `LCase` | Variant | To lowercase | Convert to lowercase (Variant) |
//! | `LCase$` | String | To lowercase | Convert to lowercase (String) |
//! | `StrConv` | String | Various conversions | Complex case conversions |
//!
//! ## Common Use Cases
//!
//! ### HTTP Header Names
//! ```vb
//! Function FormatHTTPHeader(headerName As String, headerValue As String) As String
//!     FormatHTTPHeader = UCase$(headerName) & ": " & headerValue
//! End Function
//! ```
//!
//! ### Database Column Names
//! ```vb
//! Function GetColumnName(fieldName As String) As String
//!     GetColumnName = UCase$(Replace(fieldName, " ", "_"))
//! End Function
//! ```
//!
//! ### License Key Formatting
//! ```vb
//! Function FormatLicenseKey(key As String) As String
//!     ' Format as XXXX-XXXX-XXXX-XXXX
//!     Dim upperKey As String
//!     upperKey = UCase$(Replace(key, "-", ""))
//!     
//!     FormatLicenseKey = Mid$(upperKey, 1, 4) & "-" & _
//!                        Mid$(upperKey, 5, 4) & "-" & _
//!                        Mid$(upperKey, 9, 4) & "-" & _
//!                        Mid$(upperKey, 13, 4)
//! End Function
//! ```
//!
//! ### Command Line Argument Parsing
//! ```vb
//! Function ParseArgument(arg As String) As String
//!     If Left$(arg, 1) = "/" Or Left$(arg, 1) = "-" Then
//!         ParseArgument = UCase$(Mid$(arg, 2))
//!     Else
//!         ParseArgument = UCase$(arg)
//!     End If
//! End Function
//! ```
//!
//! ## Platform Notes
//!
//! - On Windows, `UCase$` respects the system locale for character conversion
//! - Behavior may vary for extended ASCII and international characters
//! - For ASCII characters (a-z), behavior is consistent across all platforms
//! - Some characters may convert differently depending on the active code page
//! - Modern Windows systems handle Unicode characters in `UCase$` operations
//!
//! ## Limitations
//!
//! - Conversion is based on system locale; may not work as expected for all Unicode characters
//! - Returns `Null` if the input is `Null` (unlike some other string functions that error)
//! - Does not handle advanced Unicode normalization or case folding
//! - For true Unicode case folding, more sophisticated methods may be needed
//! - Some special characters may not convert in all locales

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn ucase_dollar_simple() {
        let source = r#"
Sub Test()
    result = UCase$("hello")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_mixed_case() {
        let source = r#"
Sub Test()
    text = UCase$("Hello World")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_with_numbers() {
        let source = r#"
Sub Test()
    mixed = UCase$("abc123")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_sql_query() {
        let source = r#"
Function BuildSQLQuery(table As String) As String
    BuildSQLQuery = UCase$("SELECT") & " * FROM " & table
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_constant_name() {
        let source = r#"
Function GenerateConstantName(baseName As String) As String
    GenerateConstantName = UCase$(Replace(baseName, " ", "_"))
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_acronym() {
        let source = r#"
Function FormatAcronym(text As String) As String
    FormatAcronym = UCase$(text)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_header() {
        let source = r#"
Function CreateHeader(title As String) As String
    CreateHeader = UCase$(title)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_command_processing() {
        let source = r#"
Function ProcessCommand(cmd As String) As Boolean
    Select Case UCase$(Trim$(cmd))
        Case "START"
            ProcessCommand = True
    End Select
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_file_extension() {
        let source = r#"
Function NormalizeExtension(filename As String) As String
    Dim ext As String
    ext = Right$(filename, 4)
    If UCase$(ext) = ".TXT" Then
        NormalizeExtension = "Text File"
    End If
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_environment_var() {
        let source = r#"
Function GetEnvironmentVar(varName As String) As String
    GetEnvironmentVar = Environ$(UCase$(varName))
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_registry_key() {
        let source = r#"
Function NormalizeRegistryKey(keyName As String) As String
    NormalizeRegistryKey = UCase$(Trim$(keyName))
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_display_name() {
        let source = r#"
Function FormatDisplayName(lastName As String) As String
    FormatDisplayName = UCase$(lastName)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_enum_member() {
        let source = r#"
Function GenerateEnumMember(memberName As String) As String
    GenerateEnumMember = UCase$(memberName)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_complex_query() {
        let source = r#"
Function BuildComplexQuery(table As String, fields As String) As String
    Dim sql As String
    sql = UCase$("SELECT") & " " & fields
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_config_writer() {
        let source = r#"
Sub WriteConfigSection(sectionName As String)
    section = UCase$(sectionName)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_report_header() {
        let source = r#"
Function GenerateReportHeader(reportTitle As String) As String
    GenerateReportHeader = UCase$(reportTitle)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_macro_validator() {
        let source = r#"
Function ValidateMacroName(macroName As String) As String
    ValidateMacroName = UCase$(macroName)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_safe_wrapper() {
        let source = r#"
Function SafeUCase(text As String) As String
    SafeUCase = UCase$(text)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_http_header() {
        let source = r#"
Function FormatHTTPHeader(headerName As String, headerValue As String) As String
    FormatHTTPHeader = UCase$(headerName) & ": " & headerValue
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }

    #[test]
    fn ucase_dollar_column_name() {
        let source = r#"
Function GetColumnName(fieldName As String) As String
    GetColumnName = UCase$(Replace(fieldName, " ", "_"))
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("UCase$"));
    }
}
