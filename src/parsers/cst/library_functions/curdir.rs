//! # `CurDir` Function
//!
//! Returns a `String` representing the current path for the specified drive or the default drive.
//!
//! ## Syntax
//!
//! ```vb
//! CurDir[(drive)]
//! ```
//!
//! ## Parameters
//!
//! - **`drive`**: Optional. `String` expression that specifies an existing drive. If no drive is
//!   specified or if drive is a zero-length string (""), `CurDir` returns the path for the
//!   current drive. The drive parameter can be just the drive letter (e.g., "C") or include
//!   a colon (e.g., "C:").
//!
//! ## Return Value
//!
//! Returns a `String` containing the current directory path for the specified drive. The returned
//! path does not include a trailing backslash unless the current directory is the root directory.
//!
//! ## Remarks
//!
//! The `CurDir` function returns the current working directory for a specified drive. This is
//! useful for:
//!
//! - Determining the current directory before changing it
//! - Building relative file paths
//! - Saving and restoring directory context
//! - File path validation
//! - Log file location determination
//!
//! **Important Characteristics:**
//!
//! - Without arguments, returns current directory of current drive
//! - With drive specified, returns current directory of that drive
//! - Does not include trailing backslash (except for root directory)
//! - Drive parameter is case-insensitive
//! - Each drive maintains its own current directory
//! - On Windows, returns full path (e.g., "C:\Windows\System32")
//! - Root directory returns drive with backslash (e.g., "C:\")
//!
//! ## Drive Specification
//!
//! The drive parameter can be specified in several ways:
//! - `CurDir()` - Current drive
//! - `CurDir("")` - Current drive
//! - `CurDir("C")` - Drive C
//! - `CurDir("C:")` - Drive C
//! - `CurDir("D")` - Drive D
//!
//! ## Examples
//!
//! ### Basic Usage
//!
//! ```vb
//! ' Get current directory of current drive
//! Dim currentDir As String
//! currentDir = CurDir()  ' Returns something like "C:\Users\Username\Documents"
//!
//! ' Get current directory of specific drive
//! Dim cDrive As String
//! cDrive = CurDir("C")  ' Returns current directory on C: drive
//!
//! Dim dDrive As String
//! dDrive = CurDir("D:")  ' Returns current directory on D: drive
//! ```
//!
//! ### Save and Restore Directory
//!
//! ```vb
//! Sub ProcessInDifferentDirectory(targetDir As String)
//!     Dim savedDir As String
//!     
//!     ' Save current directory
//!     savedDir = CurDir()
//!     
//!     ' Change to target directory
//!     ChDir targetDir
//!     
//!     ' Do work in target directory
//!     ProcessFiles
//!     
//!     ' Restore original directory
//!     ChDir savedDir
//! End Sub
//! ```
//!
//! ### Building Relative Paths
//!
//! ```vb
//! Function GetFullPath(relativePath As String) As String
//!     ' Combine current directory with relative path
//!     If Right(CurDir(), 1) = "\" Then
//!         GetFullPath = CurDir() & relativePath
//!     Else
//!         GetFullPath = CurDir() & "\" & relativePath
//!     End If
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Check if at Root Directory
//!
//! ```vb
//! Function IsRootDirectory() As Boolean
//!     Dim currentPath As String
//!     currentPath = CurDir()
//!     
//!     ' Root directory ends with backslash (e.g., "C:\")
//!     IsRootDirectory = (Len(currentPath) = 3 And Right(currentPath, 1) = "\")
//! End Function
//! ```
//!
//! ### Get Current Drive Letter
//!
//! ```vb
//! Function GetCurrentDrive() As String
//!     Dim currentPath As String
//!     currentPath = CurDir()
//!     
//!     ' Extract drive letter (first character)
//!     GetCurrentDrive = Left(currentPath, 1)
//! End Function
//! ```
//!
//! ### Ensure Trailing Backslash
//!
//! ```vb
//! Function EnsureTrailingBackslash(path As String) As String
//!     If Right(path, 1) <> "\" Then
//!         EnsureTrailingBackslash = path & "\"
//!     Else
//!         EnsureTrailingBackslash = path
//!     End If
//! End Function
//!
//! ' Usage
//! Dim dirPath As String
//! dirPath = EnsureTrailingBackslash(CurDir())
//! ```
//!
//! ### Directory Context Manager
//!
//! ```vb
//! Type DirectoryContext
//!     SavedDirectory As String
//! End Type
//!
//! Function PushDirectory(newDir As String) As DirectoryContext
//!     Dim ctx As DirectoryContext
//!     ctx.SavedDirectory = CurDir()
//!     ChDir newDir
//!     PushDirectory = ctx
//! End Function
//!
//! Sub PopDirectory(ctx As DirectoryContext)
//!     ChDir ctx.SavedDirectory
//! End Sub
//!
//! ' Usage
//! Dim ctx As DirectoryContext
//! ctx = PushDirectory("C:\Temp")
//! ' Do work...
//! PopDirectory ctx
//! ```
//!
//! ### Multi-Drive Path Tracking
//!
//! ```vb
//! Function GetAllDrivePaths() As Collection
//!     Dim paths As New Collection
//!     Dim drives() As String
//!     Dim i As Integer
//!     
//!     drives = Array("C", "D", "E", "F")
//!     
//!     On Error Resume Next
//!     For i = LBound(drives) To UBound(drives)
//!         paths.Add CurDir(drives(i)), drives(i)
//!     Next i
//!     On Error GoTo 0
//!     
//!     Set GetAllDrivePaths = paths
//! End Function
//! ```
//!
//! ### Log File in Current Directory
//!
//! ```vb
//! Function GetLogFilePath() As String
//!     Dim currentDir As String
//!     Dim logFile As String
//!     
//!     currentDir = CurDir()
//!     logFile = "application.log"
//!     
//!     If Right(currentDir, 1) = "\" Then
//!         GetLogFilePath = currentDir & logFile
//!     Else
//!         GetLogFilePath = currentDir & "\" & logFile
//!     End If
//! End Function
//! ```
//!
//! ### Temporary Directory Operations
//!
//! ```vb
//! Sub ProcessInTempDirectory()
//!     Dim originalDir As String
//!     Dim tempDir As String
//!     
//!     originalDir = CurDir()
//!     tempDir = Environ("TEMP")
//!     
//!     On Error GoTo Cleanup
//!     
//!     ChDir tempDir
//!     
//!     ' Process files in temp directory
//!     ProcessTempFiles
//!     
//! Cleanup:
//!     ChDir originalDir
//! End Sub
//! ```
//!
//! ### Validate Relative Path
//!
//! ```vb
//! Function IsRelativePath(path As String) As Boolean
//!     ' Check if path is relative (doesn't start with drive letter)
//!     IsRelativePath = (InStr(path, ":") = 0)
//! End Function
//!
//! Function ResolveRelativePath(relativePath As String) As String
//!     If IsRelativePath(relativePath) Then
//!         ResolveRelativePath = CurDir() & "\" & relativePath
//!     Else
//!         ResolveRelativePath = relativePath
//!     End If
//! End Function
//! ```
//!
//! ### Directory Breadcrumb Trail
//!
//! ```vb
//! Function GetDirectoryParts() As String()
//!     Dim currentDir As String
//!     Dim parts() As String
//!     
//!     currentDir = CurDir()
//!     
//!     ' Remove drive letter and colon
//!     If InStr(currentDir, ":") > 0 Then
//!         currentDir = Mid(currentDir, 3)
//!     End If
//!     
//!     ' Remove leading backslash
//!     If Left(currentDir, 1) = "\" Then
//!         currentDir = Mid(currentDir, 2)
//!     End If
//!     
//!     ' Split by backslash
//!     If Len(currentDir) > 0 Then
//!         parts = Split(currentDir, "\")
//!     End If
//!     
//!     GetDirectoryParts = parts
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Cross-Drive File Operations
//!
//! ```vb
//! Sub CopyFileToAnotherDrive(sourceFile As String, targetDrive As String)
//!     Dim sourceDrive As String
//!     Dim targetPath As String
//!     
//!     ' Get current directory on target drive
//!     On Error Resume Next
//!     targetPath = CurDir(targetDrive)
//!     
//!     If Err.Number = 0 Then
//!         ' Build target file path
//!         If Right(targetPath, 1) <> "\" Then
//!             targetPath = targetPath & "\"
//!         End If
//!         
//!         FileCopy sourceFile, targetPath & Dir(sourceFile)
//!     End If
//! End Sub
//! ```
//!
//! ### Directory Stack Implementation
//!
//! ```vb
//! Private dirStack As Collection
//!
//! Sub InitDirectoryStack()
//!     Set dirStack = New Collection
//! End Sub
//!
//! Sub PushDir(Optional newDir As String)
//!     If dirStack Is Nothing Then InitDirectoryStack
//!     
//!     ' Save current directory
//!     dirStack.Add CurDir()
//!     
//!     ' Change to new directory if specified
//!     If Len(newDir) > 0 Then
//!         ChDir newDir
//!     End If
//! End Sub
//!
//! Sub PopDir()
//!     If dirStack Is Nothing Then Exit Sub
//!     If dirStack.Count = 0 Then Exit Sub
//!     
//!     ' Restore previous directory
//!     ChDir dirStack(dirStack.Count)
//!     dirStack.Remove dirStack.Count
//! End Sub
//! ```
//!
//! ### Smart Path Concatenation
//!
//! ```vb
//! Function CombinePaths(ParamArray paths() As Variant) As String
//!     Dim result As String
//!     Dim i As Integer
//!     Dim part As String
//!     
//!     If UBound(paths) < LBound(paths) Then
//!         ' No paths provided, return current directory
//!         CombinePaths = CurDir()
//!         Exit Function
//!     End If
//!     
//!     result = CStr(paths(LBound(paths)))
//!     
//!     For i = LBound(paths) + 1 To UBound(paths)
//!         part = CStr(paths(i))
//!         
//!         ' Remove leading backslash from part
//!         If Left(part, 1) = "\" Then part = Mid(part, 2)
//!         
//!         ' Add backslash if needed
//!         If Right(result, 1) <> "\" Then result = result & "\"
//!         
//!         result = result & part
//!     Next i
//!     
//!     CombinePaths = result
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! Function GetCurrentDirectorySafe(Optional drive As String = "") As String
//!     On Error GoTo ErrorHandler
//!     
//!     If Len(drive) = 0 Then
//!         GetCurrentDirectorySafe = CurDir()
//!     Else
//!         GetCurrentDirectorySafe = CurDir(drive)
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 68  ' Device unavailable
//!             MsgBox "Drive " & drive & " is not available.", vbExclamation
//!         Case 71  ' Disk not ready
//!             MsgBox "Drive " & drive & " is not ready.", vbExclamation
//!         Case Else
//!             MsgBox "Error getting current directory: " & Err.Description, vbCritical
//!     End Select
//!     
//!     GetCurrentDirectorySafe = ""
//! End Function
//! ```
//!
//! ### Common Errors
//!
//! - **Error 68** (Device unavailable): Specified drive does not exist
//! - **Error 71** (Disk not ready): Drive exists but is not ready (e.g., no CD in drive)
//! - **Error 5** (Invalid procedure call): Invalid drive specification
//!
//! ## Performance Considerations
//!
//! - `CurDir` is a fast function with minimal overhead
//! - Results can be cached if directory won't change during execution
//! - Accessing network drives may have latency
//! - For frequently used paths, cache the result in a variable
//!
//! ## Best Practices
//!
//! ### Always Restore Directory
//!
//! ```vb
//! Sub SafeDirectoryOperation()
//!     Dim savedDir As String
//!     savedDir = CurDir()
//!     
//!     On Error GoTo Cleanup
//!     
//!     ' Change directory and do work
//!     ChDir "C:\Temp"
//!     ProcessFiles
//!     
//! Cleanup:
//!     ChDir savedDir
//! End Sub
//! ```
//!
//! ### Use Absolute Paths When Possible
//!
//! ```vb
//! ' Instead of relying on current directory:
//! Open "data.txt" For Input As #1  ' Depends on CurDir
//!
//! ' Use absolute paths:
//! Open "C:\MyApp\Data\data.txt" For Input As #1  ' Explicit path
//! ```
//!
//! ### Validate Drive Before Use
//!
//! ```vb
//! Function IsDriveAvailable(drive As String) As Boolean
//!     On Error Resume Next
//!     Dim test As String
//!     test = CurDir(drive)
//!     IsDriveAvailable = (Err.Number = 0)
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ## Platform Considerations
//!
//! - **Windows**: Returns paths with backslashes (e.g., "C:\Windows")
//! - **Drive letters**: Windows-specific concept
//! - **Network paths**: `UNC` paths (\\server\share) not supported by `CurDir`
//! - **Long paths**: Paths longer than 260 characters may cause issues
//! - **Case sensitivity**: Windows file system is case-insensitive
//!
//! ## Limitations
//!
//! - Returns only local drive paths, not `UNC` network paths
//! - Cannot set the current directory (use `ChDir` for that)
//! - Drive must be available and ready
//! - Does not validate that the returned path still exists
//! - Each drive remembers its own current directory independently
//! - Does not work with drives that don't have current directory concept
//!
//! ## Related Functions
//!
//! - `ChDir`: Changes the current directory
//! - `ChDrive`: Changes the current drive
//! - `Dir`: Returns files/directories matching a pattern
//! - `MkDir`: Creates a new directory
//! - `RmDir`: Removes an empty directory
//! - `App.Path`: Returns the path where the application started

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn curdir_basic() {
        let source = r#"
currentDir = CurDir()
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_with_drive() {
        let source = r#"
path = CurDir("C")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_with_drive_colon() {
        let source = r#"
path = CurDir("C:")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_in_assignment() {
        let source = r#"
Dim savedDir As String
savedDir = CurDir()
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_save_restore() {
        let source = r#"
savedDir = CurDir()
ChDir "C:\Temp"
ChDir savedDir
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_in_function() {
        let source = r#"
Function GetCurrentPath() As String
    GetCurrentPath = CurDir()
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_with_concatenation() {
        let source = r#"
fullPath = CurDir() & "\data.txt"
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_in_if_statement() {
        let source = r#"
If Right(CurDir(), 1) = "\" Then
    ProcessRoot
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_with_right() {
        let source = r#"
If Right(CurDir(), 1) <> "\" Then
    path = CurDir() & "\"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_with_left() {
        let source = r#"
drive = Left(CurDir(), 1)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_multiple_drives() {
        let source = r#"
cPath = CurDir("C")
dPath = CurDir("D")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_with_error_handling() {
        let source = r#"
On Error Resume Next
path = CurDir("X")
If Err.Number <> 0 Then
    MsgBox "Drive not available"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_in_msgbox() {
        let source = r#"
MsgBox "Current directory: " & CurDir()
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_with_variable() {
        let source = r#"
Dim drv As String
drv = "C"
path = CurDir(drv)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_in_select_case() {
        let source = r#"
Select Case CurDir()
    Case "C:\"
        ProcessRoot
    Case Else
        ProcessOther
End Select
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_with_len() {
        let source = r#"
If Len(CurDir()) = 3 Then
    isRoot = True
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_in_loop() {
        let source = r#"
For i = 1 To 5
    path = CurDir() & "\file" & i & ".txt"
Next i
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_with_instr() {
        let source = r#"
If InStr(CurDir(), "Windows") > 0 Then
    ProcessWindowsDir
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_empty_string() {
        let source = r#"
currentPath = CurDir("")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_in_print() {
        let source = r#"
Print "Current directory: "; CurDir()
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_comparison() {
        let source = r#"
If CurDir() = "C:\Windows" Then
    ProcessWindows
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_in_do_loop() {
        let source = r#"
Do While CurDir() <> "C:\"
    ChDir ".."
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_with_mid() {
        let source = r#"
pathPart = Mid(CurDir(), 4)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_in_sub() {
        let source = r#"
Sub SaveCurrentDir()
    savedPath = CurDir()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }

    #[test]
    fn curdir_with_whitespace() {
        let source = r#"
path = CurDir( )
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let debug = tree.debug_tree();
        assert!(debug.contains("CurDir"));
        assert!(debug.contains("Identifier"));
    }
}
