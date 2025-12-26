//! # `CurDir$` Function
//!
//! Returns a `String` representing the current path for the specified drive or the default drive.
//! The dollar sign suffix (`$`) explicitly indicates that this function returns a `String` type
//! (not a `Variant`).
//!
//! ## Syntax
//!
//! ```vb
//! CurDir$[(drive)]
//! ```
//!
//! ## Parameters
//!
//! - **`drive`**: Optional. `String` expression that specifies an existing drive. If no drive is
//!   specified or if drive is a zero-length string (""), `CurDir$` returns the path for the
//!   current drive. The drive parameter can be just the drive letter (e.g., "C") or include
//!   a colon (e.g., "C:").
//!
//! ## Return Value
//!
//! Returns a `String` containing the current directory path for the specified drive. The returned
//! path does not include a trailing backslash unless the current directory is the root directory.
//! The return value is always a `String` type (never `Variant`).
//!
//! ## Remarks
//!
//! - The `CurDir$` function always returns a `String`, while `CurDir` (without `$`) can return a `Variant`.
//! - Without arguments, returns current directory of current drive.
//! - With drive specified, returns current directory of that drive.
//! - Does not include trailing backslash (except for root directory).
//! - Drive parameter is case-insensitive ("C" and "c" are equivalent).
//! - Each drive maintains its own current directory in Windows.
//! - On Windows, returns full path (e.g., "C:\Windows\System32").
//! - Root directory returns drive with backslash (e.g., "C:\").
//! - For better performance when you know the result is a string, use `CurDir$` instead of `CurDir`.
//!
//! ## Drive Specification
//!
//! The drive parameter can be specified in several ways:
//! - `CurDir$()` - Current drive
//! - `CurDir$("")` - Current drive
//! - `CurDir$("C")` - Drive C
//! - `CurDir$("C:")` - Drive C
//! - `CurDir$("D")` - Drive D
//!
//! ## Typical Uses
//!
//! 1. **Directory context** - Determine current working directory
//! 2. **Path building** - Construct full paths from relative paths
//! 3. **Directory restoration** - Save and restore directory state
//! 4. **File operations** - Locate files relative to current directory
//! 5. **Logging** - Record current directory for debugging
//! 6. **Validation** - Verify expected working directory
//! 7. **Multi-drive operations** - Work with multiple drives simultaneously
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Get current directory
//! Dim currentDir As String
//! currentDir = CurDir$()
//! ```
//!
//! ```vb
//! ' Example 2: Get current directory of specific drive
//! Dim cDrive As String
//! cDrive = CurDir$("C")
//! ```
//!
//! ```vb
//! ' Example 3: Check if in expected directory
//! If CurDir$() = "C:\MyApp" Then
//!     MsgBox "In correct directory"
//! End If
//! ```
//!
//! ```vb
//! ' Example 4: Display current directory
//! MsgBox "Current directory: " & CurDir$()
//! ```
//!
//! ## Common Patterns
//!
//! ### Save and Restore Directory
//! ```vb
//! Sub ProcessInDifferentDirectory(targetDir As String)
//!     Dim savedDir As String
//!     
//!     ' Save current directory
//!     savedDir = CurDir$()
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     ' Change to target directory
//!     ChDir targetDir
//!     
//!     ' Do work in target directory
//!     ProcessFiles
//!     
//!     ' Restore original directory
//!     ChDir savedDir
//!     Exit Sub
//!     
//! ErrorHandler:
//!     ' Always restore directory even on error
//!     ChDir savedDir
//!     Err.Raise Err.Number, , Err.Description
//! End Sub
//! ```
//!
//! ### Building Full Path from Relative Path
//! ```vb
//! Function GetFullPath(relativePath As String) As String
//!     Dim currentDir As String
//!     currentDir = CurDir$()
//!     
//!     ' Check if current dir already ends with backslash
//!     If Right$(currentDir, 1) = "\" Then
//!         GetFullPath = currentDir & relativePath
//!     Else
//!         GetFullPath = currentDir & "\" & relativePath
//!     End If
//! End Function
//! ```
//!
//! ### Ensuring Correct Working Directory
//! ```vb
//! Sub EnsureWorkingDirectory(expectedDir As String)
//!     Dim currentDir As String
//!     currentDir = CurDir$()
//!     
//!     If UCase$(currentDir) <> UCase$(expectedDir) Then
//!         ChDir expectedDir
//!     End If
//! End Sub
//! ```
//!
//! ### Multi-Drive Directory Tracking
//! ```vb
//! Function GetAllDriveDirs() As Collection
//!     Dim drives() As String
//!     Dim result As New Collection
//!     Dim i As Integer
//!     
//!     drives = Array("C", "D", "E", "F")
//!     
//!     For i = LBound(drives) To UBound(drives)
//!         On Error Resume Next
//!         result.Add CurDir$(drives(i)), drives(i)
//!         On Error GoTo 0
//!     Next i
//!     
//!     Set GetAllDriveDirs = result
//! End Function
//! ```
//!
//! ### Path Validation
//! ```vb
//! Function IsValidDirectory(dirPath As String) As Boolean
//!     Dim savedDir As String
//!     Dim result As Boolean
//!     
//!     savedDir = CurDir$()
//!     result = False
//!     
//!     On Error Resume Next
//!     ChDir dirPath
//!     If Err.Number = 0 Then
//!         result = True
//!         ChDir savedDir
//!     End If
//!     On Error GoTo 0
//!     
//!     IsValidDirectory = result
//! End Function
//! ```
//!
//! ### Log File Path Construction
//! ```vb
//! Function GetLogFilePath() As String
//!     Dim logDir As String
//!     Dim logFile As String
//!     
//!     logDir = CurDir$()
//!     logFile = "application_" & Format$(Now, "yyyymmdd") & ".log"
//!     
//!     If Right$(logDir, 1) = "\" Then
//!         GetLogFilePath = logDir & logFile
//!     Else
//!         GetLogFilePath = logDir & "\" & logFile
//!     End If
//! End Function
//! ```
//!
//! ### Working Directory Report
//! ```vb
//! Sub ReportCurrentDirectories()
//!     Dim report As String
//!     
//!     report = "Current Directory: " & CurDir$() & vbCrLf
//!     
//!     On Error Resume Next
//!     report = report & "C: Drive: " & CurDir$("C") & vbCrLf
//!     report = report & "D: Drive: " & CurDir$("D") & vbCrLf
//!     On Error GoTo 0
//!     
//!     MsgBox report, vbInformation, "Directory Report"
//! End Sub
//! ```
//!
//! ### Relative File Search
//! ```vb
//! Function FindFileInCurrentDir(filename As String) As String
//!     Dim fullPath As String
//!     Dim currentDir As String
//!     
//!     currentDir = CurDir$()
//!     
//!     If Right$(currentDir, 1) = "\" Then
//!         fullPath = currentDir & filename
//!     Else
//!         fullPath = currentDir & "\" & filename
//!     End If
//!     
//!     If Dir$(fullPath) <> "" Then
//!         FindFileInCurrentDir = fullPath
//!     Else
//!         FindFileInCurrentDir = ""
//!     End If
//! End Function
//! ```
//!
//! ### Directory Depth Calculator
//! ```vb
//! Function GetDirectoryDepth() As Integer
//!     Dim path As String
//!     Dim depth As Integer
//!     Dim i As Integer
//!     
//!     path = CurDir$()
//!     depth = 0
//!     
//!     For i = 1 To Len(path)
//!         If Mid$(path, i, 1) = "\" Then
//!             depth = depth + 1
//!         End If
//!     Next i
//!     
//!     GetDirectoryDepth = depth - 1  ' Subtract 1 for root backslash
//! End Function
//! ```
//!
//! ### Safe Directory Operation Wrapper
//! ```vb
//! Function ExecuteInDirectory(dirPath As String, operation As String) As Boolean
//!     Dim savedDir As String
//!     Dim success As Boolean
//!     
//!     savedDir = CurDir$()
//!     success = False
//!     
//!     On Error GoTo ErrorHandler
//!     
//!     ' Change to target directory
//!     ChDir dirPath
//!     
//!     ' Execute operation based on parameter
//!     Select Case operation
//!         Case "CLEANUP"
//!             DeleteTempFiles
//!         Case "BACKUP"
//!             BackupFiles
//!         Case "SCAN"
//!             ScanFiles
//!     End Select
//!     
//!     success = True
//!     
//! ErrorHandler:
//!     ' Always restore directory
//!     ChDir savedDir
//!     ExecuteInDirectory = success
//! End Function
//! ```
//!
//! ## Related Functions
//!
//! - `CurDir`: Returns current directory as `Variant` instead of `String`
//! - `ChDir`: Changes the current directory
//! - `ChDrive`: Changes the current drive
//! - `App.Path`: Returns the path where the application executable is located
//! - `Dir$`: Returns files or directories matching a pattern
//! - `MkDir`: Creates a new directory
//! - `RmDir`: Removes an empty directory
//!
//! ## Best Practices
//!
//! 1. Always save current directory before changing it
//! 2. Use error handling when changing directories
//! 3. Restore original directory in error handlers
//! 4. Use `App.Path` for application-relative paths instead of relying on current directory
//! 5. Check for trailing backslash when building paths
//! 6. Use `CurDir$` instead of `CurDir` for better performance
//! 7. Be aware that current directory can be changed by other code
//! 8. Document assumptions about current directory in your code
//! 9. Use absolute paths when possible to avoid directory dependencies
//! 10. Test code with different working directories
//!
//! ## Performance Considerations
//!
//! - `CurDir$` is slightly more efficient than `CurDir` because it avoids `Variant` overhead
//! - Directory queries are fast system calls
//! - Cache the result if you need to use it multiple times in quick succession
//! - Avoid excessive directory changes as they affect all operations
//!
//! ## Platform Notes
//!
//! - Windows maintains separate current directories for each drive
//! - Network paths (UNC paths) are supported (e.g., "\\\\server\\share\\folder")
//! - Long path names (> 260 characters) may cause issues in VB6
//! - Current directory is process-specific and thread-safe
//! - Some operations (like file dialogs) may change the current directory
//!
//! ## Error Conditions
//!
//! - If the specified drive does not exist, an error occurs (Error 68: Device unavailable)
//! - If the drive is not ready (e.g., no disk in drive), an error occurs (Error 71: Disk not ready)
//! - Network drives that are disconnected will cause errors
//!
//! ## Security Considerations
//!
//! 1. Don't assume the current directory for security-sensitive operations
//! 2. Use absolute paths for configuration and data files
//! 3. Be aware that current directory can be manipulated by attackers
//! 4. Validate directory paths before changing to them
//! 5. Use `App.Path` for application resources rather than current directory
//!
//! ## Limitations
//!
//! - Cannot set the current directory (use `ChDir` for that)
//! - Path is limited to `MAX_PATH` characters (typically 260) in VB6
//! - Does not resolve symbolic links or junctions
//! - Drive letter parameter is Windows-specific

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn curdir_dollar_no_args() {
        let source = r"
Sub Main()
    currentDir = CurDir$()
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_with_drive() {
        let source = r#"
Sub Main()
    cDrive = CurDir$("C")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_with_drive_colon() {
        let source = r#"
Sub Main()
    dDrive = CurDir$("D:")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_in_condition() {
        let source = r#"
Sub Main()
    If CurDir$() = "C:\MyApp" Then
        MsgBox "Correct directory"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_save_restore() {
        let source = r#"
Sub Main()
    Dim savedDir As String
    savedDir = CurDir$()
    ChDir "C:\Temp"
    ChDir savedDir
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_path_building() {
        let source = r#"
Function GetFullPath(relativePath As String) As String
    GetFullPath = CurDir$() & "\" & relativePath
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_with_right() {
        let source = r#"
Sub Main()
    If Right$(CurDir$(), 1) = "\" Then
        path = CurDir$() & "file.txt"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_msgbox() {
        let source = r#"
Sub Main()
    msg = "Current: " & CurDir$()
    MsgBox msg
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_ucase() {
        let source = r#"
Sub Main()
    If UCase$(CurDir$()) = "C:\MYAPP" Then
        Process
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_multiple_drives() {
        let source = r#"
Sub Main()
    Dim c As String, d As String
    c = CurDir$("C")
    d = CurDir$("D")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_error_handler() {
        let source = r#"
Sub Main()
    Dim savedDir As String
    savedDir = CurDir$()
    On Error GoTo ErrorHandler
    ChDir "C:\NewPath"
    Exit Sub
ErrorHandler:
    ChDir savedDir
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_with_dir() {
        let source = r#"
Function FindFile(filename As String) As String
    fullPath = CurDir$() & "\" & filename
    If Dir$(fullPath) <> "" Then
        FindFile = fullPath
    End If
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_depth_calc() {
        let source = r"
Function GetDepth() As Integer
    Dim path As String
    path = CurDir$()
    GetDepth = Len(path)
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_log_path() {
        let source = r#"
Function GetLogPath() As String
    GetLogPath = CurDir$() & "\app.log"
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_validation() {
        let source = r"
Sub EnsureDirectory(expectedDir As String)
    If CurDir$() <> expectedDir Then
        ChDir expectedDir
    End If
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_format() {
        let source = r#"
Sub Main()
    logFile = CurDir$() & "\log_" & Format$(Now, "yyyymmdd") & ".txt"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_debug_print() {
        let source = r#"
Sub Main()
    Debug.Print "Current Dir: " & CurDir$()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_select_case() {
        let source = r#"
Sub Main()
    Select Case UCase$(CurDir$())
        Case "C:\WINDOWS"
            mode = 1
        Case "C:\TEMP"
            mode = 2
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_empty_string() {
        let source = r#"
Sub Main()
    currentDir = CurDir$("")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }

    #[test]
    fn curdir_dollar_variable_drive() {
        let source = r#"
Sub Main()
    Dim drive As String
    drive = "C"
    path = CurDir$(drive)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("Identifier") && debug.contains("CurDir$"));
    }
}
