//! # `Command$` Function
//!
//! Returns the argument portion of the command line used to launch Microsoft Visual Basic or an
//! executable program developed with Visual Basic. The dollar sign suffix (`$`) explicitly
//! indicates that this function returns a `String` type (not a `Variant`).
//!
//! ## Syntax
//!
//! ```vb
//! Command$()
//! ```
//!
//! ## Parameters
//!
//! None. The `Command$` function takes no arguments.
//!
//! ## Return Value
//!
//! Returns a `String` containing the command-line arguments passed to the program. If no arguments
//! were passed, returns an empty string (""). The return value is always a `String` type (never `Variant`).
//!
//! ## Remarks
//!
//! - The `Command$` function always returns a `String`, while `Command` (without `$`) can return a `Variant`.
//! - Returns only the arguments, not the executable path or name.
//! - Arguments are returned as a single string, exactly as passed to the application.
//! - Multiple arguments are separated by spaces (unless quoted).
//! - Quoted strings preserve internal spaces but quotes may be included in the result.
//! - Leading and trailing spaces are typically trimmed by the system.
//! - Returns empty string ("") if no arguments were provided.
//! - Case is preserved as entered on the command line.
//! - For better performance when you know the result is a string, use `Command$` instead of `Command`.
//!
//! ## Command Line Processing
//!
//! When an application is launched with:
//! ```text
//! MyApp.exe /debug file.txt "long filename.doc"
//! ```
//!
//! `Command$()` returns:
//! ```text
//! /debug file.txt "long filename.doc"
//! ```
//!
//! ## Typical Uses
//!
//! 1. **Processing startup parameters** - Read switches and configuration flags
//! 2. **File path handling** - Accept file paths to open at startup
//! 3. **Debug modes** - Enable special debugging or logging modes
//! 4. **Automation** - Support scripted or automated workflows
//! 5. **Configuration** - Pass runtime configuration without config files
//! 6. **Batch processing** - Process multiple files or operations
//! 7. **Integration** - Allow other applications to control behavior
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Get command line arguments
//! Sub Main()
//!     Dim cmdLine As String
//!     cmdLine = Command$()
//!     MsgBox "Arguments: " & cmdLine
//! End Sub
//! ```
//!
//! ```vb
//! ' Example 2: Check if arguments provided
//! Sub Main()
//!     If Command$() <> "" Then
//!         MsgBox "Arguments: " & Command$()
//!     Else
//!         MsgBox "No arguments"
//!     End If
//! End Sub
//! ```
//!
//! ```vb
//! ' Example 3: Simple file opener
//! Sub Main()
//!     Dim filename As String
//!     filename = Trim$(Command$())
//!     If filename <> "" Then
//!         OpenFile filename
//!     End If
//! End Sub
//! ```
//!
//! ```vb
//! ' Example 4: Check for debug mode
//! Sub Main()
//!     If InStr(Command$(), "/debug") > 0 Then
//!         EnableDebugMode
//!     End If
//! End Sub
//! ```
//!
//! ## Common Patterns
//!
//! ### Processing Multiple Arguments
//! ```vb
//! Function ParseArguments() As Collection
//!     Dim args As String
//!     Dim result As New Collection
//!     Dim parts() As String
//!     
//!     args = Command$()
//!     If args = "" Then
//!         Set ParseArguments = result
//!         Exit Function
//!     End If
//!     
//!     parts = Split(args, " ")
//!     Dim i As Integer
//!     For i = LBound(parts) To UBound(parts)
//!         If Trim$(parts(i)) <> "" Then
//!             result.Add Trim$(parts(i))
//!         End If
//!     Next i
//!     
//!     Set ParseArguments = result
//! End Function
//! ```
//!
//! ### Processing Switches and Parameters
//! ```vb
//! Sub Main()
//!     Dim args As String
//!     args = Command$()
//!     
//!     ' Check for various switches
//!     If InStr(args, "/debug") > 0 Then
//!         App.LogMode = 1
//!     End If
//!     
//!     If InStr(args, "/silent") > 0 Then
//!         App.SilentMode = True
//!     End If
//!     
//!     If InStr(args, "/verbose") > 0 Then
//!         App.VerboseMode = True
//!     End If
//! End Sub
//! ```
//!
//! ### Opening File from Command Line
//! ```vb
//! Sub Main()
//!     Dim filename As String
//!     filename = Trim$(Command$())
//!     
//!     If filename <> "" Then
//!         ' Remove surrounding quotes if present
//!         If Left$(filename, 1) = Chr$(34) Then
//!             filename = Mid$(filename, 2)
//!         End If
//!         If Right$(filename, 1) = Chr$(34) Then
//!             filename = Left$(filename, Len(filename) - 1)
//!         End If
//!         
//!         ' Verify file exists and open it
//!         If Dir$(filename) <> "" Then
//!             LoadDocument filename
//!         Else
//!             MsgBox "File not found: " & filename
//!         End If
//!     End If
//! End Sub
//! ```
//!
//! ### Named Parameter Extraction
//! ```vb
//! Function GetParameter(paramName As String) As String
//!     Dim args As String
//!     Dim pos As Integer
//!     Dim endPos As Integer
//!     Dim result As String
//!     
//!     args = " " & Command$() & " "
//!     pos = InStr(1, args, "/" & paramName & ":", vbTextCompare)
//!     
//!     If pos > 0 Then
//!         pos = pos + Len(paramName) + 2
//!         endPos = InStr(pos, args, " ")
//!         If endPos > pos Then
//!             result = Mid$(args, pos, endPos - pos)
//!         End If
//!     End If
//!     
//!     GetParameter = result
//! End Function
//! ```
//!
//! ### Logging Startup Arguments
//! ```vb
//! Sub Main()
//!     Dim args As String
//!     Dim logFile As Integer
//!     
//!     args = Command$()
//!     
//!     logFile = FreeFile
//!     Open App.Path & "\startup.log" For Append As #logFile
//!     Print #logFile, Now & " - Started with args: " & args
//!     Close #logFile
//! End Sub
//! ```
//!
//! ### Configuration from Command Line
//! ```vb
//! Sub Main()
//!     Dim args As String
//!     args = UCase$(Command$())
//!     
//!     ' Set configuration based on arguments
//!     If InStr(args, "/SERVER:") > 0 Then
//!         App.ServerName = GetParameter("server")
//!     End If
//!     
//!     If InStr(args, "/PORT:") > 0 Then
//!         App.Port = Val(GetParameter("port"))
//!     End If
//!     
//!     If InStr(args, "/USER:") > 0 Then
//!         App.UserName = GetParameter("user")
//!     End If
//! End Sub
//! ```
//!
//! ### Help Display
//! ```vb
//! Sub Main()
//!     Dim args As String
//!     args = LCase$(Trim$(Command$()))
//!     
//!     If args = "/?" Or args = "-?" Or args = "/help" Or args = "-help" Then
//!         DisplayHelp
//!         End
//!     End If
//! End Sub
//!
//! Sub DisplayHelp()
//!     Dim helpText As String
//!     helpText = "Usage: MyApp [options]" & vbCrLf
//!     helpText = helpText & "/debug    - Enable debug mode" & vbCrLf
//!     helpText = helpText & "/silent   - Run in silent mode" & vbCrLf
//!     helpText = helpText & "/file:xxx - Open specified file" & vbCrLf
//!     MsgBox helpText
//! End Sub
//! ```
//!
//! ### Batch Mode Processing
//! ```vb
//! Sub Main()
//!     Dim args As String
//!     Dim files() As String
//!     Dim i As Integer
//!     
//!     args = Command$()
//!     
//!     If InStr(args, "/batch") > 0 Then
//!         ' Parse file list
//!         files = Split(Replace$(args, "/batch", ""), " ")
//!         
//!         For i = LBound(files) To UBound(files)
//!             If Trim$(files(i)) <> "" Then
//!                 ProcessFile Trim$(files(i))
//!             End If
//!         Next i
//!         
//!         End  ' Exit after batch processing
//!     End If
//! End Sub
//! ```
//!
//! ### Error Recovery
//! ```vb
//! Sub Main()
//!     On Error GoTo ErrorHandler
//!     
//!     Dim args As String
//!     args = Command$()
//!     
//!     ' Process arguments
//!     If args <> "" Then
//!         ProcessCommandLine args
//!     End If
//!     
//!     Exit Sub
//!     
//! ErrorHandler:
//!     MsgBox "Error processing command line: " & args & vbCrLf & _
//!            "Error: " & Err.Description
//!     End
//! End Sub
//! ```
//!
//! ### Case-Insensitive Switch Detection
//! ```vb
//! Function HasSwitch(switchName As String) As Boolean
//!     Dim args As String
//!     args = " " & UCase$(Command$()) & " "
//!     switchName = " /" & UCase$(switchName) & " "
//!     HasSwitch = (InStr(args, switchName) > 0)
//! End Function
//! ```
//!
//! ## Related Functions
//!
//! - `Command`: Returns command-line arguments as `Variant` instead of `String`
//! - `App.Path`: Returns the path where the application executable is located
//! - `App.EXEName`: Returns the name of the executable file
//! - `Environ$`: Returns environment variable values
//!
//! ## Best Practices
//!
//! 1. Always trim the result to remove leading/trailing spaces
//! 2. Handle the case where no arguments are provided (empty string)
//! 3. Use case-insensitive comparison for switches and parameters
//! 4. Document expected command-line format in your application
//! 5. Validate arguments before using them
//! 6. Provide meaningful error messages for invalid arguments
//! 7. Consider implementing a `/help` or `/?` switch
//! 8. Use `Command$` instead of `Command` for better performance
//! 9. Be careful with quoted strings - they may include the quotes
//! 10. Log startup arguments for debugging and support purposes
//!
//! ## Performance Considerations
//!
//! - `Command$` is slightly more efficient than `Command` because it avoids `Variant` overhead
//! - The function is typically called once at startup, so performance is rarely a concern
//! - Parsing complex command lines can be slow; cache the result if needed multiple times
//! - Consider using a dedicated command-line parser for complex argument processing
//!
//! ## Platform Notes
//!
//! - Command-line argument handling is consistent across Windows platforms
//! - Maximum command-line length varies by Windows version (typically 8191 characters)
//! - Arguments are passed by the operating system when the executable is launched
//! - VB6 IDE does not allow setting command-line arguments for debugging
//! - Use a shortcut or command prompt to test command-line arguments
//!
//! ## Security Considerations
//!
//! 1. Never execute command-line arguments directly as code
//! 2. Validate all file paths before accessing files
//! 3. Sanitize arguments before using in SQL queries or shell commands
//! 4. Limit accepted argument values to known good values when possible
//! 5. Log suspicious or malformed command-line arguments
//!
//! ## Limitations
//!
//! - Returns arguments as a single string (manual parsing required for multiple arguments)
//! - Does not provide the executable path or name (use `App.Path` and `App.EXEName`)
//! - No built-in support for named parameters (must implement custom parsing)
//! - Quote handling is system-dependent and may vary
//! - Cannot distinguish between missing arguments and empty string argument

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn command_dollar_simple() {
        let source = r#"
Sub Main()
    args = Command$()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_assignment() {
        let source = r#"
Sub Main()
    Dim cmdLine As String
    cmdLine = Command$()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_in_condition() {
        let source = r#"
Sub Main()
    If Command$() <> "" Then
        MsgBox "Arguments provided"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_with_trim() {
        let source = r#"
Sub Main()
    filename = Trim$(Command$())
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_with_instr() {
        let source = r#"
Sub Main()
    If InStr(Command$(), "/debug") > 0 Then
        EnableDebug
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_parse_args() {
        let source = r#"
Function ParseArguments() As Collection
    Dim args As String
    args = Command$()
    If args = "" Then Exit Function
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_multiple_checks() {
        let source = r#"
Sub Main()
    Dim args As String
    args = Command$()
    If InStr(args, "/debug") > 0 Then Debug = True
    If InStr(args, "/silent") > 0 Then Silent = True
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_file_opener() {
        let source = r#"
Sub Main()
    Dim filename As String
    filename = Trim$(Command$())
    If filename <> "" Then
        OpenFile filename
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_with_split() {
        let source = r#"
Sub Main()
    Dim parts() As String
    parts = Split(Command$(), " ")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_logging() {
        let source = r#"
Sub Main()
    Dim logFile As Integer
    logFile = FreeFile
    Print #logFile, "Args: " & Command$()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_ucase() {
        let source = r#"
Sub Main()
    Dim args As String
    args = UCase$(Command$())
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_lcase() {
        let source = r#"
Sub Main()
    args = LCase$(Trim$(Command$()))
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_help_check() {
        let source = r#"
Sub Main()
    If Command$() = "/?" Then
        DisplayHelp
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_replace() {
        let source = r#"
Sub Main()
    args = Replace$(Command$(), "/batch", "")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_concatenation() {
        let source = r#"
Sub Main()
    msg = "Started with: " & Command$()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_function_call() {
        let source = r#"
Function HasSwitch(switchName As String) As Boolean
    Dim args As String
    args = UCase$(Command$())
    HasSwitch = (InStr(args, switchName) > 0)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_empty_check() {
        let source = r#"
Sub Main()
    If Len(Command$()) = 0 Then
        MsgBox "No arguments"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_direct_print() {
        let source = r#"
Sub Main()
    Debug.Print Command$()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }

    #[test]
    fn command_dollar_select_case() {
        let source = r#"
Sub Main()
    Select Case LCase$(Command$())
        Case "/debug"
            DebugMode = True
        Case "/release"
            DebugMode = False
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("Command$"));
    }
}
