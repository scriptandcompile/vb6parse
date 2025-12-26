//! # `SetAttr` Statement
//!
//! Sets attribute information for a file.
//!
//! ## Syntax
//!
//! ```vb
//! SetAttr pathname, attributes
//! ```
//!
//! ## Parts
//!
//! - **pathname**: Required. String expression that specifies a file name. May include directory or folder, and drive.
//! - **attributes**: Required. Numeric expression or constant specifying the file attributes. Sum of the values of the file attribute constants.
//!
//! ## File Attribute Constants
//!
//! | Constant | Value | Description |
//! |----------|-------|-------------|
//! | vbNormal | 0 | Normal (no attributes set) |
//! | vbReadOnly | 1 | Read-only file attribute |
//! | vbHidden | 2 | Hidden file attribute |
//! | vbSystem | 4 | System file attribute |
//! | vbArchive | 32 | File has changed since last backup |
//!
//! ## Remarks
//!
//! - **Combining Attributes**: You can combine attributes by adding their values together (e.g., `vbReadOnly + vbHidden = 3`).
//! - **File Must Exist**: A run-time error occurs if the file specified by pathname doesn't exist.
//! - **Pathname Validation**: Pathname can be a fully qualified path or a relative path. Wildcard characters (* and ?) are not supported.
//! - **Cannot Set Directory Attribute**: You cannot use `SetAttr` to set the directory (vbDirectory = 16) attribute. Use `MkDir` and `RmDir` instead.
//! - **Volume Label**: You cannot use `SetAttr` to set the volume label (vbVolume = 8) attribute.
//! - **Read-Only Directories**: `SetAttr` cannot change the read-only status of a directory; it only works with files.
//! - **Error Handling**: Use error handling to trap potential errors like file not found, permission denied, or invalid attributes.
//! - **`GetAttr` Function**: Use `GetAttr` to retrieve current file attributes before modifying them with `SetAttr`.
//! - **Clearing Attributes**: To clear an attribute, set the file to vbNormal (0) or use a combination that excludes the unwanted attribute.
//!
//! ## Examples
//!
//! ### Set File to Read-Only
//!
//! ```vb
//! SetAttr "C:\MyFile.txt", vbReadOnly
//! ```
//!
//! ### Set File to Hidden
//!
//! ```vb
//! SetAttr "C:\Data\Secret.dat", vbHidden
//! ```
//!
//! ### Combine Multiple Attributes
//!
//! ```vb
//! ' Set file to read-only and hidden
//! SetAttr "C:\Config.ini", vbReadOnly + vbHidden
//! ```
//!
//! ### Clear All Attributes (Normal)
//!
//! ```vb
//! SetAttr "C:\MyFile.txt", vbNormal
//! ```
//!
//! ### Set Archive Attribute
//!
//! ```vb
//! SetAttr "C:\Backup\Data.dat", vbArchive
//! ```
//!
//! ### Using Variables
//!
//! ```vb
//! Dim fileName As String
//! Dim attrs As Integer
//!
//! fileName = "C:\Data\MyFile.txt"
//! attrs = vbReadOnly + vbArchive
//! SetAttr fileName, attrs
//! ```
//!
//! ### Toggle Read-Only Attribute
//!
//! ```vb
//! Dim currentAttrs As Integer
//! Dim filePath As String
//!
//! filePath = "C:\MyFile.txt"
//! currentAttrs = GetAttr(filePath)
//!
//! If currentAttrs And vbReadOnly Then
//!     ' Remove read-only
//!     SetAttr filePath, currentAttrs And Not vbReadOnly
//! Else
//!     ' Add read-only
//!     SetAttr filePath, currentAttrs Or vbReadOnly
//! End If
//! ```
//!
//! ### Set System File
//!
//! ```vb
//! SetAttr "C:\Windows\system.dat", vbSystem
//! ```
//!
//! ### Set Multiple Files in a Loop
//!
//! ```vb
//! Dim i As Integer
//! For i = 1 To 10
//!     SetAttr "C:\Files\File" & i & ".txt", vbReadOnly
//! Next i
//! ```
//!
//! ### With Error Handling
//!
//! ```vb
//! On Error Resume Next
//! SetAttr "C:\MyFile.txt", vbReadOnly
//! If Err.Number <> 0 Then
//!     MsgBox "Could not set file attributes: " & Err.Description
//! End If
//! On Error GoTo 0
//! ```
//!
//! ### Using App.Path
//!
//! ```vb
//! SetAttr App.Path & "\Config.ini", vbHidden
//! ```
//!
//! ### Preserve Existing Attributes While Adding New Ones
//!
//! ```vb
//! Dim filePath As String
//! Dim currentAttrs As Integer
//!
//! filePath = "C:\MyFile.txt"
//! currentAttrs = GetAttr(filePath)
//!
//! ' Add hidden attribute while preserving others
//! SetAttr filePath, currentAttrs Or vbHidden
//! ```
//!
//! ### Remove Specific Attribute
//!
//! ```vb
//! Dim filePath As String
//! Dim currentAttrs As Integer
//!
//! filePath = "C:\MyFile.txt"
//! currentAttrs = GetAttr(filePath)
//!
//! ' Remove hidden attribute while preserving others
//! SetAttr filePath, currentAttrs And Not vbHidden
//! ```
//!
//! ### Using Numeric Values
//!
//! ```vb
//! SetAttr "C:\MyFile.txt", 1  ' Same as vbReadOnly
//! SetAttr "C:\MyFile.txt", 3  ' Read-only + Hidden (1 + 2)
//! SetAttr "C:\MyFile.txt", 35 ' Read-only + Hidden + Archive (1 + 2 + 32)
//! ```
//!
//! ### Conditional Attribute Setting
//!
//! ```vb
//! If FileIsImportant Then
//!     SetAttr filePath, vbReadOnly + vbArchive
//! Else
//!     SetAttr filePath, vbNormal
//! End If
//! ```
//!
//! ## Common Errors
//!
//! - **Error 53**: File not found - occurs if the pathname doesn't exist
//! - **Error 5**: Invalid procedure call or argument - occurs if attributes value is invalid
//! - **Error 70**: Permission denied - occurs if you don't have write access to the file
//! - **Error 75**: Path/File access error - occurs if the file is open or locked
//!
//! ## Important Notes
//!
//! - **File Must Be Closed**: The file should not be open when you use `SetAttr`.
//! - **Permissions Required**: You must have appropriate permissions to change file attributes.
//! - **Network Files**: `SetAttr` works with network files if you have appropriate permissions.
//! - **UNC Paths**: `SetAttr` supports UNC (Universal Naming Convention) paths like "\\\\Server\\Share\\File.txt".
//! - **Attribute Persistence**: File attributes persist after the application closes; they're stored in the file system.
//! - **Read-Only Files**: To modify a read-only file, you must first remove the read-only attribute, make changes, then restore it.
//! - **`GetAttr` Complement**: Always use `GetAttr` to retrieve current attributes before modifying them to avoid unintentionally removing existing attributes.
//!
//! ## Best Practices
//!
//! - Use error handling when working with `SetAttr` as file operations can fail for many reasons
//! - Use `GetAttr` before `SetAttr` to preserve existing attributes you don't want to change
//! - Use symbolic constants (vbReadOnly, etc.) instead of numeric values for better code readability
//! - Check file existence using `Dir()` before calling `SetAttr`
//! - Be cautious when setting system attributes as they can affect system stability
//! - Document why specific attributes are being set, especially for hidden or system files
//! - Consider user permissions when setting attributes on shared or network files
//!
//! ## See Also
//!
//! - `GetAttr` function (retrieve file attributes)
//! - `Dir` function (check if file exists)
//! - `Kill` statement (delete files)
//! - `Name` statement (rename files)
//! - `FileCopy` statement (copy files)
//!
//! ## References
//!
//! - [SetAttr Statement - Microsoft Docs](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/setattr-statement)

use crate::parsers::cst::Parser;
use crate::parsers::syntaxkind::SyntaxKind;

impl Parser<'_> {
    /// Parses a `SetAttr` statement.
    pub(crate) fn parse_setattr_statement(&mut self) {
        self.parse_simple_builtin_statement(SyntaxKind::SetAttrStatement);
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn setattr_simple() {
        let source = r#"
Sub Test()
    SetAttr "C:\MyFile.txt", vbReadOnly
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
        assert!(debug.contains("SetAttrKeyword"));
    }

    #[test]
    fn setattr_at_module_level() {
        let source = "SetAttr \"C:\\File.txt\", vbNormal\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_readonly() {
        let source = r#"
Sub Test()
    SetAttr "C:\MyFile.txt", vbReadOnly
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
        assert!(debug.contains("vbReadOnly"));
    }

    #[test]
    fn setattr_hidden() {
        let source = r#"
Sub Test()
    SetAttr "C:\Data\Secret.dat", vbHidden
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
        assert!(debug.contains("vbHidden"));
    }

    #[test]
    fn setattr_combined_attributes() {
        let source = r#"
Sub Test()
    SetAttr "C:\Config.ini", vbReadOnly + vbHidden
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
        assert!(debug.contains("vbReadOnly"));
        assert!(debug.contains("vbHidden"));
    }

    #[test]
    fn setattr_normal() {
        let source = r#"
Sub Test()
    SetAttr "C:\MyFile.txt", vbNormal
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
        assert!(debug.contains("vbNormal"));
    }

    #[test]
    fn setattr_archive() {
        let source = r#"
Sub Test()
    SetAttr "C:\Backup\Data.dat", vbArchive
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
        assert!(debug.contains("vbArchive"));
    }

    #[test]
    fn setattr_with_variables() {
        let source = r"
Sub Test()
    SetAttr fileName, attrs
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
        assert!(debug.contains("fileName"));
        assert!(debug.contains("attrs"));
    }

    #[test]
    fn setattr_system() {
        let source = r#"
Sub Test()
    SetAttr "C:\Windows\system.dat", vbSystem
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
        assert!(debug.contains("vbSystem"));
    }

    #[test]
    fn setattr_numeric_value() {
        let source = r#"
Sub Test()
    SetAttr "C:\MyFile.txt", 1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_with_app_path() {
        let source = r#"
Sub Test()
    SetAttr App.Path & "\Config.ini", vbHidden
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_inside_if_statement() {
        let source = r"
If FileExists Then
    SetAttr filePath, vbReadOnly
End If
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_inside_loop() {
        let source = r#"
For i = 1 To 10
    SetAttr "C:\Files\File" & i & ".txt", vbReadOnly
Next i
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_with_comment() {
        let source = r#"
Sub Test()
    SetAttr "C:\MyFile.txt", vbReadOnly ' Make read-only
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
        assert!(debug.contains("' Make read-only"));
    }

    #[test]
    fn setattr_preserves_whitespace() {
        let source = "SetAttr   \"File.txt\"  ,   vbNormal\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_with_getattr() {
        let source = r"
Sub Test()
    currentAttrs = GetAttr(filePath)
    SetAttr filePath, currentAttrs Or vbHidden
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
        assert!(debug.contains("GetAttr"));
    }

    #[test]
    fn setattr_in_select_case() {
        let source = r"
Select Case fileType
    Case 1
        SetAttr filePath, vbReadOnly
    Case 2
        SetAttr filePath, vbHidden
End Select
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_multiple_on_same_line() {
        let source = "SetAttr \"File1.txt\", vbNormal: SetAttr \"File2.txt\", vbReadOnly\n";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_in_with_block() {
        let source = r"
With fileObj
    SetAttr .Path, vbArchive
End With
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_in_sub() {
        let source = r#"
Sub MakeReadOnly()
    SetAttr "C:\MyFile.txt", vbReadOnly
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_in_function() {
        let source = r"
Function SetFileAttributes(path As String) As Boolean
    SetAttr path, vbReadOnly + vbArchive
    SetFileAttributes = True
End Function
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_unc_path() {
        let source = r#"
Sub Test()
    SetAttr "\\Server\Share\File.txt", vbReadOnly
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_in_class_module() {
        let source = r"
Private filePath As String

Public Sub SetReadOnly()
    SetAttr filePath, vbReadOnly
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.cls", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_with_line_continuation() {
        let source = r#"
Sub Test()
    SetAttr _
        "C:\MyFile.txt", vbReadOnly
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_with_concatenation() {
        let source = r#"
Sub Test()
    SetAttr "C:\Data\" & fileName & ".txt", vbHidden
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_toggle_readonly() {
        let source = r"
Sub Test()
    If currentAttrs And vbReadOnly Then
        SetAttr filePath, currentAttrs And Not vbReadOnly
    Else
        SetAttr filePath, currentAttrs Or vbReadOnly
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_error_handling() {
        let source = r#"
On Error Resume Next
SetAttr "C:\MyFile.txt", vbReadOnly
If Err.Number <> 0 Then
    MsgBox "Error"
End If
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_with_dir_function() {
        let source = r#"
Sub Test()
    If Dir(filePath) <> "" Then
        SetAttr filePath, vbReadOnly
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_multiple_calls() {
        let source = r#"
Sub Test()
    SetAttr "File1.txt", vbReadOnly
    SetAttr "File2.txt", vbHidden
    SetAttr "File3.txt", vbArchive
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }

    #[test]
    fn setattr_conditional() {
        let source = r"
Sub Test()
    If FileIsImportant Then
        SetAttr filePath, vbReadOnly + vbArchive
    Else
        SetAttr filePath, vbNormal
    End If
End Sub
";
        let cst = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("SetAttrStatement"));
    }
}
