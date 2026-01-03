//! # `RmDir` Statement
//!
//! Removes an empty directory or folder.
//!
//! ## Syntax
//!
//! ```vb
//! RmDir path
//! ```
//!
//! - `path`: Required. String expression that identifies the directory or folder to be removed. May include drive.
//!   If no drive is specified, `RmDir` removes the directory or folder on the current drive.
//!
//! ## Remarks
//!
//! - An error occurs if you try to use `RmDir` on a directory containing files. Use the Kill statement to delete all files before attempting to remove a directory.
//! - An error also occurs if you try to remove a directory that doesn't exist.
//! - The directory must be empty (contain no files or subdirectories) before it can be removed.
//! - The `path` argument can include absolute or relative paths.
//! - On Windows systems, both forward slashes (/) and backslashes (\) can be used as path separators.
//! - The directory name can include the drive letter.
//! - You cannot remove the current directory. You must change to a parent or different directory first.
//! - UNC paths are supported on network drives.
//! - To remove a directory tree, you must remove all subdirectories first (working from innermost to outermost).
//!
//! ## Examples
//!
//! ```vb
//! ' Remove a directory in the current directory
//! RmDir "OldFolder"
//!
//! ' Remove a directory with full path
//! RmDir "C:\Temp\TempFiles"
//!
//! ' Remove a directory on another drive
//! RmDir "D:\Data\Archive"
//!
//! ' Remove nested directories (must remove innermost first)
//! RmDir "C:\Temp\Logs\Archive"
//! RmDir "C:\Temp\Logs"
//! RmDir "C:\Temp"
//!
//! ' Remove directory on network drive
//! RmDir "\\Server\Share\OldFolder"
//!
//! ' Safe removal with error handling
//! On Error Resume Next
//! RmDir "C:\Temp\ToDelete"
//! If Err.Number <> 0 Then
//!     MsgBox "Could not remove directory"
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Common Errors
//!
//! - **Error 75**: Path/File access error - directory contains files or subdirectories
//! - **Error 76**: Path not found - directory doesn't exist
//! - **Error 5**: Invalid procedure call - trying to remove current directory
//!
//! ## Reference
//!
//! [RmDir Statement - Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/rmdir-statement)

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses an `RmDir` statement.
    pub(crate) fn parse_rmdir_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::RmDirStatement);
    }
}

#[cfg(test)]
mod tests {
    use crate::*;

    // RmDir statement tests
    #[test]
    fn rmdir_simple() {
        let source = r#"
Sub Test()
    RmDir "OldFolder"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
        assert!(debug.contains("RmDirKeyword"));
        assert!(debug.contains("OldFolder"));
    }

    #[test]
    fn rmdir_at_module_level() {
        let source = r#"RmDir "C:\Temp""#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
    }

    #[test]
    fn rmdir_with_full_path() {
        let source = r#"
Sub Test()
    RmDir "C:\Temp\TempFiles"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
        assert!(debug.contains("RmDirKeyword"));
    }

    #[test]
    fn rmdir_with_variable() {
        let source = r#"
Sub Test()
    Dim folderPath As String
    folderPath = "C:\TempData"
    RmDir folderPath
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
        assert!(debug.contains("folderPath"));
    }

    #[test]
    fn rmdir_with_drive_letter() {
        let source = r#"
Sub Test()
    RmDir "D:\Archive"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
    }

    #[test]
    fn rmdir_relative_path() {
        let source = r#"
Sub Test()
    RmDir "SubFolder"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
    }

    #[test]
    fn rmdir_parent_directory() {
        let source = r#"
Sub Test()
    RmDir "..\TempFolder"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
    }

    #[test]
    fn rmdir_unc_path() {
        let source = r#"
Sub Test()
    RmDir "\\Server\Share\TempFolder"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
    }

    #[test]
    fn rmdir_in_if_statement() {
        let source = r#"
Sub Test()
    If folderExists Then
        RmDir "C:\Temp"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn rmdir_inline_if() {
        let source = r"
Sub Test()
    If shouldDelete Then RmDir tempPath
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
    }

    #[test]
    fn rmdir_with_comment() {
        let source = r#"
Sub Test()
    RmDir "C:\Temp" ' Remove temporary directory
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn rmdir_preserves_whitespace() {
        let source = r#"    RmDir    "C:\Temp"    "#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert!(cst.text().contains("RmDir"));

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
    }

    #[test]
    fn rmdir_with_error_handling() {
        let source = r#"
Sub CleanupDirectories()
    On Error Resume Next
    RmDir "C:\Temp\Session1"
    If Err.Number <> 0 Then
        MsgBox "Could not remove directory"
    End If
    On Error GoTo 0
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
        assert!(debug.contains("OnErrorStatement"));
    }

    #[test]
    fn rmdir_multiple_directories() {
        let source = r#"
Sub Test()
    RmDir "C:\Temp\Folder1"
    RmDir "C:\Temp\Folder2"
    RmDir "C:\Temp\Folder3"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("RmDirStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn rmdir_in_loop() {
        let source = r#"
Sub RemoveTemporaryFolders()
    Dim i As Integer
    For i = 1 To 10
        RmDir "C:\Temp\Session" & i
    Next i
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
        assert!(debug.contains("ForStatement"));
    }

    #[test]
    fn rmdir_nested_removal() {
        let source = r#"
Sub RemoveNestedDirectories()
    ' Remove innermost first
    RmDir "C:\Project\Temp\Cache\Old"
    RmDir "C:\Project\Temp\Cache"
    RmDir "C:\Project\Temp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("RmDirStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn rmdir_with_concatenation() {
        let source = r#"
Sub Test()
    Dim basePath As String
    basePath = "C:\Temp\"
    RmDir basePath & "OldData"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
    }

    #[test]
    fn rmdir_after_kill() {
        let source = r#"
Sub RemoveDirectoryWithFiles()
    On Error Resume Next
    Kill "C:\Temp\Session\*.*"
    RmDir "C:\Temp\Session"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
        assert!(debug.contains("KillStatement"));
    }

    #[test]
    fn rmdir_with_function_result() {
        let source = r"
Sub Test()
    RmDir GetTempPath()
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
        assert!(debug.contains("GetTempPath"));
    }

    #[test]
    fn rmdir_in_select_case() {
        let source = r#"
Sub Test()
    Select Case action
        Case "cleanup"
            RmDir "C:\Temp"
        Case "archive"
            RmDir "C:\Archive\Old"
        Case Else
            RmDir defaultPath
    End Select
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("RmDirStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn rmdir_with_do_loop() {
        let source = r"
Sub Test()
    Dim dirExists As Boolean
    dirExists = True
    Do While dirExists
        On Error Resume Next
        RmDir tempDir
        dirExists = (Err.Number = 0)
    Loop
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
        assert!(debug.contains("DoStatement"));
    }

    #[test]
    fn rmdir_with_chdir() {
        let source = r#"
Sub Test()
    ChDir "C:\"
    RmDir "C:\OldTemp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
        assert!(debug.contains("ChDirStatement"));
    }

    #[test]
    fn rmdir_conditional_removal() {
        let source = r#"
Sub Test()
    If Dir("C:\Temp", vbDirectory) <> "" Then
        RmDir "C:\Temp"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
    }

    #[test]
    fn rmdir_with_line_continuation() {
        let source = r#"
Sub Test()
    RmDir _
        "C:\VeryLongPathName\SubFolder\TempData"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
    }

    #[test]
    fn rmdir_in_class_module() {
        let source = r"
Private Sub Class_Terminate()
    On Error Resume Next
    RmDir m_tempDirectory
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.cls", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
    }

    #[test]
    fn rmdir_with_app_path() {
        let source = r#"
Sub Test()
    RmDir App.Path & "\TempCache"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
    }

    #[test]
    fn rmdir_recursive_cleanup() {
        let source = r#"
Sub RecursiveRemove()
    On Error Resume Next
    Kill "C:\Temp\Data\*.*"
    RmDir "C:\Temp\Data"
    Kill "C:\Temp\*.*"
    RmDir "C:\Temp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("RmDirStatement").count();
        assert_eq!(count, 2);
    }

    #[test]
    fn rmdir_with_environ() {
        let source = r#"
Sub Test()
    RmDir Environ("TEMP") & "\MyApp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
    }

    #[test]
    fn rmdir_after_mkdir() {
        let source = r#"
Sub Test()
    MkDir "C:\TempWork"
    ' Do some work
    RmDir "C:\TempWork"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
        assert!(debug.contains("MkDirStatement"));
    }

    #[test]
    fn rmdir_with_explicit_error_check() {
        let source = r"
Function RemoveDirectory(path As String) As Boolean
    On Error GoTo ErrorHandler
    RmDir path
    RemoveDirectory = True
    Exit Function
ErrorHandler:
    RemoveDirectory = False
End Function
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
        assert!(debug.contains("FunctionStatement"));
    }

    #[test]
    fn rmdir_forward_slash_path() {
        let source = r#"
Sub Test()
    RmDir "C:/Temp/Session"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("RmDirStatement"));
    }
}
