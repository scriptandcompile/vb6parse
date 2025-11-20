//! Built-in VB6 statement parsing.
//!
//! This module handles parsing of VB6 built-in system statements:
//! - AppActivate: Activate an application window
//! - Beep: Sound a tone through the computer's speaker
//! - ChDir: Change the current directory or folder
//! - ChDrive: Change the current drive
//! - Close: Close files opened with the Open statement
//! - Date: Set the current system date
//! - DeleteSetting: Delete a section or key setting from the Windows registry
//! - Error: Generate a run-time error
//! - FileCopy: Copy a file
//! - Get: Read data from an open disk file into a variable
//! - Input: Read data from an open sequential file
//! - Reset: Close all disk files opened using the Open statement

use crate::language::VB6Token;
use crate::parsers::SyntaxKind;

use super::Parser;

impl<'a> Parser<'a> {
    /// Check if the current token is a built-in statement keyword.
    pub(super) fn is_builtin_statement_keyword(&self) -> bool {
        matches!(
            self.current_token(),
            Some(VB6Token::AppActivateKeyword)
                | Some(VB6Token::BeepKeyword)
                | Some(VB6Token::ChDirKeyword)
                | Some(VB6Token::ChDriveKeyword)
                | Some(VB6Token::CloseKeyword)
                | Some(VB6Token::DateKeyword)
                | Some(VB6Token::DeleteSettingKeyword)
                | Some(VB6Token::ErrorKeyword)
                | Some(VB6Token::FileCopyKeyword)
                | Some(VB6Token::GetKeyword)
                | Some(VB6Token::InputKeyword)
                | Some(VB6Token::ResetKeyword)
        )
    }

    /// Dispatch built-in statement parsing to the appropriate parser.
    pub(super) fn parse_builtin_statement(&mut self) {
        match self.current_token() {
            Some(VB6Token::AppActivateKeyword) => {
                // VB6 AppActivate statement syntax:
                // - AppActivate title[, wait]
                //
                // Activates an application window.
                //
                // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/appactivate-statement)
                self.parse_simple_builtin_statement(SyntaxKind::AppActivateStatement);
            }
            Some(VB6Token::BeepKeyword) => {
                // VB6 Beep statement syntax:
                // - Beep
                //
                // Emits a standard system beep sound.
                //
                // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/beep-statement)
                self.parse_simple_builtin_statement(SyntaxKind::BeepStatement);
            }
            Some(VB6Token::ChDirKeyword) => {
                // VB6 ChDir statement syntax:
                // - ChDir path
                //
                // Changes the current directory or folder.
                //
                // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chdir-statement)
                self.parse_simple_builtin_statement(SyntaxKind::ChDirStatement);
            }
            Some(VB6Token::ChDriveKeyword) => {
                // VB6 ChDrive statement syntax:
                // - ChDrive drive
                //
                // Changes the current drive.
                //
                // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chdrive-statement)
                self.parse_simple_builtin_statement(SyntaxKind::ChDriveStatement);
            }
            Some(VB6Token::CloseKeyword) => {
                // VB6 Close statement syntax:
                // - Close [filenumberlist]
                //
                // Closes input or output files opened using the Open statement.
                //
                // filenumberlist: Optional. One or more file numbers using the syntax:
                // [[#]filenumber] [, [#]filenumber] ...
                //
                // If filenumberlist is omitted, all active files opened by the Open statement are closed.
                //
                // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/close-statement)
                self.parse_simple_builtin_statement(SyntaxKind::CloseStatement);
            }
            Some(VB6Token::DateKeyword) => {
                // VB6 Date statement syntax:
                // - Date = dateexpression
                //
                // Sets the current system date.
                //
                // dateexpression: Required. Any expression that can represent a date.
                //
                // Note: The Date statement is used to set the date. To retrieve the current date,
                // use the Date function.
                //
                // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/date-statement)
                self.parse_simple_builtin_statement(SyntaxKind::DateStatement);
            }
            Some(VB6Token::DeleteSettingKeyword) => {
                // VB6 DeleteSetting statement syntax:
                // - DeleteSetting appname, section[, key]
                //
                // Deletes a section or key setting from an application's entry in the Windows registry.
                //
                // The DeleteSetting statement syntax has these named arguments:
                //
                // | Part     | Description |
                // |----------|-------------|
                // | appname  | Required. String expression containing the name of the application or project to which the section or key setting applies. |
                // | section  | Required. String expression containing the name of the section from which the key setting is being deleted. If only appname and section are provided, the specified section is deleted along with all related key settings. |
                // | key      | Optional. String expression containing the name of the key setting being deleted. |
                //
                // Examples:
                // - DeleteSetting "MyApp", "Startup" (deletes entire Startup section)
                // - DeleteSetting "MyApp", "Startup", "Left" (deletes Left key from Startup section)
                // - DeleteSetting App.ProductName, "FileFilter" (deletes FileFilter section)
                //
                // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/deletesetting-statement)
                self.parse_simple_builtin_statement(SyntaxKind::DeleteSettingStatement);
            }
            Some(VB6Token::ErrorKeyword) => {
                // VB6 Error statement syntax:
                // - Error errornumber
                //
                // Generates a run-time error; can be used instead of the Err.Raise method.
                //
                // The Error statement syntax has this part:
                //
                // | Part          | Description |
                // |---------------|-------------|
                // | errornumber   | Required. Any valid error number. |
                //
                // Remarks:
                // - The Error statement is supported for backward compatibility.
                // - In new code, use the Err object's Raise method to generate run-time errors.
                // - If errornumber is defined, the Error statement calls the error handler after the properties
                //   of the Err object are assigned the following default values:
                //   * Err.Number: The value specified as the argument to the Error statement
                //   * Err.Source: The name of the current Visual Basic project
                //   * Err.Description: String expression corresponding to the return value of the Error function
                //     for the specified Number, if this string exists
                //
                // Examples:
                // ```vb
                // Error 11  ' Generate "Division by zero" error
                // Error 53  ' Generate "File not found" error
                // Error vbObjectError + 1000  ' Generate custom error
                // ```
                //
                // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/error-statement)
                self.parse_simple_builtin_statement(SyntaxKind::ErrorStatement);
            }
            Some(VB6Token::FileCopyKeyword) => {
                // VB6 FileCopy statement syntax:
                // - FileCopy source, destination
                //
                // Copies a file.
                //
                // The FileCopy statement syntax has these named arguments:
                //
                // | Part          | Description |
                // |---------------|-------------|
                // | source        | Required. String expression that specifies a file name. May include directory or folder, and drive. |
                // | destination   | Required. String expression that specifies a file name. May include directory or folder, and drive. |
                //
                // Remarks:
                // - If you try to use the FileCopy statement on a currently open file, an error occurs.
                // - FileCopy can copy files between directories/folders and between drives.
                // - Both source and destination can include path information (drive and directory/folder).
                // - If destination specifies a directory/folder that doesn't exist, FileCopy creates it.
                //
                // Examples:
                // ```vb
                // FileCopy "C:\SOURCE.TXT", "C:\DEST.TXT"
                // FileCopy oldFile, newFile
                // FileCopy App.Path & "\data.dat", "C:\Backup\data.dat"
                // ```
                //
                // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filecopy-statement)
                self.parse_simple_builtin_statement(SyntaxKind::FileCopyStatement);
            }
            Some(VB6Token::GetKeyword) => {
                // VB6 Get statement syntax:
                // - Get [#]filenumber, [recnumber], varname
                //
                // Reads data from an open disk file into a variable.
                //
                // The Get statement syntax has these parts:
                //
                // | Part          | Description |
                // |---------------|-------------|
                // | filenumber    | Required. Any valid file number. |
                // | recnumber     | Optional. Variant (Long). Record number (Random mode files) or byte number (Binary mode files) at which reading begins. |
                // | varname       | Required. Valid variable name into which data is read. |
                //
                // Remarks:
                // - Get is used with files opened in Binary or Random mode.
                // - For files opened in Random mode, the record length specified in the Open statement determines the number of bytes read.
                // - For files opened in Binary mode, Get reads any number of bytes.
                // - The first record or byte in a file is at position 1, the second at position 2, and so on.
                // - If you omit recnumber, the next record or byte following the last Get or Put statement (or pointed to by the last Seek function) is read.
                // - You must include delimiting commas, for example: Get #1, , myVariable
                // - For files opened in Random mode, the following rules apply:
                //   * If the length of the data being read is less than the length specified in the Len clause, subsequent records on disk are aligned on record-length boundaries.
                //   * The space between the end of one record and the beginning of the next is padded with existing file contents.
                //   * If the variable being read is a variable-length string, Get reads a 2-byte descriptor containing the string length and then reads the string data.
                // - For files opened in Binary mode, all the Random rules apply, except:
                //   * The Len clause in the Open statement has no effect.
                //   * Get reads the data contiguously, with no padding between records.
                //
                // Examples:
                // ```vb
                // Get #1, , myRecord
                // Get #1, recordNumber, customerData
                // Get fileNum, , buffer
                // ```
                //
                // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/get-statement)
                self.parse_simple_builtin_statement(SyntaxKind::GetStatement);
            }
            Some(VB6Token::InputKeyword) => {
                // VB6 Input statement syntax:
                // - Input #filenumber, varlist
                //
                // Reads data from an open sequential file and assigns the data to variables.
                //
                // The Input # statement syntax has these parts:
                //
                // | Part          | Description |
                // |---------------|-------------|
                // | filenumber    | Required. Any valid file number. |
                // | varlist       | Required. Comma-delimited list of variables that are assigned values read from the file. Variables can't be arrays or object variables. However, variables that describe an element of an array or user-defined type may be used. |
                //
                // Remarks:
                // - Data read with Input # is usually written to a file with Write #.
                // - Use this statement only with files opened in Input or Binary mode.
                // - The Input # statement reads data items from a sequential file and assigns them to variables.
                // - Data items in the file must appear in the same order as the variables in varlist and be separated by commas.
                // - If the data item to be read is a quoted string, Input # strips the quotation marks.
                // - Input # is typically used to read data that was written to a file using the Write # statement.
                // - For files opened for Binary access, Input # reads all the bytes it needs to complete the varlist.
                // - If end of file is reached before all variables are filled, an error occurs.
                //
                // Examples:
                // ```vb
                // Input #1, name, age
                // Input #fileNum, x, y, z
                // Input #1, firstName, lastName, address
                // ```
                //
                // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/input-statement)
                self.parse_simple_builtin_statement(SyntaxKind::InputStatement);
            }
            Some(VB6Token::ResetKeyword) => {
                // VB6 Reset statement syntax:
                // - Reset
                //
                // Closes all disk files opened using the Open statement.
                //
                // The Reset statement closes all active files opened by the Open statement
                // and writes the contents of all file buffers to disk.
                //
                // Use Reset to ensure all file data is written to disk before ending your program.
                // This is particularly important in programs that may terminate abnormally.
                //
                // [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/reset-statement)
                self.parse_simple_builtin_statement(SyntaxKind::ResetStatement);
            }
            _ => {}
        }
    }

    /// Generic parser for built-in statements that follow the pattern:
    /// - Keyword [arguments]
    ///
    /// All built-in statements in this module share the same structure:
    /// 1. Set parsing_header to false
    /// 2. Start a syntax node of the given kind
    /// 3. Consume the keyword token
    /// 4. Consume everything until newline (arguments/parameters)
    /// 5. Consume the newline
    /// 6. Finish the syntax node
    pub(super) fn parse_simple_builtin_statement(&mut self, kind: SyntaxKind) {
        // if we are now parsing a built-in statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(kind.to_raw());

        // Consume the keyword
        self.consume_token();

        // Consume everything until newline (arguments/parameters)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node();
    }
}

#[cfg(test)]
mod test {
    use crate::*;

    #[test]
    fn appactivate_simple() {
        let source = r#"
Sub Test()
    AppActivate "MyApp"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("AppActivateStatement"));
        assert!(debug.contains("AppActivateKeyword"));
    }

    #[test]
    fn appactivate_with_variable() {
        let source = r#"
Sub Test()
    AppActivate lstTopWin.Text
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("AppActivateStatement"));
    }

    #[test]
    fn appactivate_with_wait_parameter() {
        let source = r#"
Sub Test()
    AppActivate "Calculator", True
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("AppActivateStatement"));
    }

    #[test]
    fn appactivate_with_title_variable() {
        let source = r#"
Sub Test()
    AppActivate sTitle
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("AppActivateStatement"));
    }

    #[test]
    fn appactivate_preserves_whitespace() {
        let source = r#"
Sub Test()
    AppActivate   "MyApp"  ,  False
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("AppActivateStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn multiple_appactivate_statements() {
        let source = r#"
Sub Test()
    AppActivate "App1"
    AppActivate "App2"
    AppActivate windowTitle
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("AppActivateStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn appactivate_in_if_statement() {
        let source = r#"
Sub Test()
    If condition Then
        AppActivate "MyApp"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("AppActivateStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn appactivate_inline_if() {
        let source = r#"
Sub Test()
    If windowExists Then AppActivate windowTitle
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("AppActivateStatement"));
    }

    #[test]
    fn appactivate_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    AppActivate lstTopWin.Text
    If Err Then MsgBox "AppActivate error: " & Err
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("AppActivateStatement"));
    }

    #[test]
    fn appactivate_at_module_level() {
        let source = r#"
AppActivate "MyApp"
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("AppActivateStatement"));
    }

    #[test]
    fn beep_simple() {
        let source = r#"
Sub Test()
    Beep
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("BeepStatement"));
        assert!(debug.contains("BeepKeyword"));
    }

    #[test]
    fn beep_in_if_statement() {
        let source = r#"
Sub Test()
    If condition Then
        Beep
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("BeepStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn beep_inline_if() {
        let source = r#"
Sub Test()
    If error Then Beep
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("BeepStatement"));
    }

    #[test]
    fn multiple_beep_statements() {
        let source = r#"
Sub Test()
    Beep
    Beep
    Beep
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("BeepStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn beep_with_comment() {
        let source = r#"
Sub Test()
    Beep ' Alert user
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("BeepStatement"));
        assert!(debug.contains("EndOfLineComment"));
    }

    #[test]
    fn beep_in_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 3
        Beep
    Next i
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("BeepStatement"));
        assert!(debug.contains("ForStatement"));
    }

    #[test]
    fn beep_in_select_case() {
        let source = r#"
Sub Test()
    Select Case value
        Case 1
            Beep
        Case Else
            Beep
    End Select
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("BeepStatement").count();
        assert_eq!(count, 2);
    }

    #[test]
    fn beep_preserves_whitespace() {
        let source = r#"
Sub Test()
    Beep    ' Extra spaces
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("BeepStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn beep_at_module_level() {
        let source = r#"
Beep
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("BeepStatement"));
    }

    #[test]
    fn beep_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Beep
    If Err Then MsgBox "Error occurred"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("BeepStatement"));
    }

    #[test]
    fn chdir_simple_string_literal() {
        let source = r#"
Sub Test()
    ChDir "C:\Windows"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_with_variable() {
        let source = r#"
Sub Test()
    ChDir myPath
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_with_app_path() {
        let source = r#"
Sub Test()
    ChDir App.Path
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_with_expression() {
        let source = r#"
Sub Test()
    ChDir GetPath() & "\subdir"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_in_if_statement() {
        let source = r#"
Sub Test()
    If dirExists Then ChDir newPath
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_at_module_level() {
        let source = r#"
ChDir "C:\Temp"
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_with_comment() {
        let source = r#"
Sub Test()
    ChDir basePath ' Change to base directory
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
        assert!(debug.contains("EndOfLineComment"));
    }

    #[test]
    fn chdir_multiple_in_sequence() {
        let source = r#"
Sub Test()
    ChDir "C:\Windows"
    ChDir "C:\Temp"
    ChDir originalPath
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let chdir_count = debug.matches("ChDirStatement").count();
        assert_eq!(chdir_count, 3, "Expected 3 ChDir statements");
    }

    #[test]
    fn chdir_in_multiline_if() {
        let source = r#"
Sub Test()
    If pathValid Then
        ChDir newPath
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_with_parentheses() {
        let source = r#"
Sub Test()
    ChDir (basePath)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_with_parentheses_without_space() {
        let source = r#"
Sub Test()
    ChDir(basePath)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdrive_simple_string_literal() {
        let source = r#"
Sub Test()
    ChDrive "C:"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_with_variable() {
        let source = r#"
Sub Test()
    ChDrive myDrive
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_with_app_path() {
        let source = r#"
Sub Test()
    ChDrive App.Path
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_with_left_function() {
        let source = r#"
Sub Test()
    ChDrive Left(sInitDir, 1)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_in_if_statement() {
        let source = r#"
Sub Test()
    If driveValid Then ChDrive newDrive
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_at_module_level() {
        let source = r#"
ChDrive "D:"
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_with_comment() {
        let source = r#"
Sub Test()
    ChDrive driveLetter ' Change to specified drive
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
        assert!(debug.contains("EndOfLineComment"));
    }

    #[test]
    fn chdrive_multiple_in_sequence() {
        let source = r#"
Sub Test()
    ChDrive "C:"
    ChDrive "D:"
    ChDrive originalDrive
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let chdrive_count = debug.matches("ChDriveStatement").count();
        assert_eq!(chdrive_count, 3, "Expected 3 ChDrive statements");
    }

    #[test]
    fn chdrive_in_multiline_if() {
        let source = r#"
Sub Test()
    If driveExists Then
        ChDrive targetDrive
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_with_parentheses() {
        let source = r#"
Sub Test()
    ChDrive (Left$(sInitDir, 1))
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_with_expression() {
        let source = r#"
Sub Test()
    ChDrive Left(theZtmPath, 1)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_and_chdir_together() {
        let source = r#"
Sub Test()
    ChDrive "C:"
    ChDir "C:\Windows"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn close_all_files() {
        let source = r#"
Sub Test()
    Close
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("CloseKeyword"));
    }

    #[test]
    fn close_single_file() {
        let source = r#"
Sub Test()
    Close #1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("CloseKeyword"));
    }

    #[test]
    fn close_single_file_without_hash() {
        let source = r#"
Sub Test()
    Close 1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("CloseKeyword"));
    }

    #[test]
    fn close_multiple_files() {
        let source = r#"
Sub Test()
    Close #1, #2, #3
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("CloseKeyword"));
    }

    #[test]
    fn close_with_variable() {
        let source = r#"
Sub Test()
    Close fileNum
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("CloseKeyword"));
    }

    #[test]
    fn close_with_hash_variable() {
        let source = r#"
Sub Test()
    Close #fileNum
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("CloseKeyword"));
    }

    #[test]
    fn close_multiple_files_mixed() {
        let source = r#"
Sub Test()
    Close #1, fileNum2, #3
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("CloseKeyword"));
    }

    #[test]
    fn close_preserves_whitespace() {
        let source = r#"
Sub Test()
    Close   #1  ,  #2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn multiple_close_statements() {
        let source = r#"
Sub Test()
    Close #1
    Close #2
    Close
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("CloseStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn close_in_if_statement() {
        let source = r#"
Sub Test()
    If fileOpen Then
        Close #1
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn close_inline_if() {
        let source = r#"
Sub Test()
    If fileOpen Then Close #fileNum
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
    }

    #[test]
    fn close_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Close #1
    If Err Then MsgBox "Error closing file"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
    }

    #[test]
    fn close_at_module_level() {
        let source = r#"
Close #1
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("CloseStatement"));
    }

    #[test]
    fn date_simple() {
        let source = r#"
Sub Test()
    Date = #1/1/2024#
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
        assert!(debug.contains("DateKeyword"));
    }

    #[test]
    fn date_with_variable() {
        let source = r#"
Sub Test()
    Date = newDate
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
        assert!(debug.contains("DateKeyword"));
    }

    #[test]
    fn date_with_function_call() {
        let source = r#"
Sub Test()
    Date = DateSerial(2024, 1, 1)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
        assert!(debug.contains("DateKeyword"));
    }

    #[test]
    fn date_with_string_expression() {
        let source = r#"
Sub Test()
    Date = "January 1, 2024"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
        assert!(debug.contains("DateKeyword"));
    }

    #[test]
    fn date_with_expression() {
        let source = r#"
Sub Test()
    Date = Now() + 7
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
        assert!(debug.contains("DateKeyword"));
    }

    #[test]
    fn date_preserves_whitespace() {
        let source = r#"
Sub Test()
    Date   =   #1/1/2024#
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn multiple_date_statements() {
        let source = r#"
Sub Test()
    Date = #1/1/2024#
    Date = #2/1/2024#
    Date = #3/1/2024#
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let count = debug.matches("DateStatement").count();
        assert_eq!(count, 3);
    }

    #[test]
    fn date_in_if_statement() {
        let source = r#"
Sub Test()
    If resetDate Then
        Date = #1/1/2024#
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
        assert!(debug.contains("IfStatement"));
    }

    #[test]
    fn date_inline_if() {
        let source = r#"
Sub Test()
    If resetDate Then Date = #1/1/2024#
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
    }

    #[test]
    fn date_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Date = userDate
    If Err Then MsgBox "Invalid date"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
    }

    #[test]
    fn deletesetting_with_section_only() {
        // Test DeleteSetting with appname and section (deletes entire section)
        let source = r#"
Sub Test()
    DeleteSetting "MyApp", "Startup"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_with_key() {
        // Test DeleteSetting with appname, section, and key
        let source = r#"
Sub Test()
    DeleteSetting "MyApp", "Startup", "Left"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_with_app_productname() {
        // Test DeleteSetting using App.ProductName
        let source = r#"
Sub Test()
    DeleteSetting App.ProductName, "FileFilter"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_with_constants() {
        // Test DeleteSetting with constants
        let source = r#"
Sub Test()
    DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Left"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_multiple_calls() {
        // Test multiple DeleteSetting calls
        let source = r#"
Sub Test()
    DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Left"
    DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Top"
    DeleteSetting REGISTRY_KEY, "Settings", "frmPost.Height"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.matches("DeleteSettingStatement").count() >= 3);
    }

    #[test]
    fn deletesetting_with_variables() {
        // Test DeleteSetting with variables
        let source = r#"
Sub Test()
    Dim appName As String
    Dim sectionName As String
    appName = "MyApp"
    sectionName = "Settings"
    DeleteSetting appName, sectionName
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_in_loop() {
        // Test DeleteSetting in a loop
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 1 To 10
        DeleteSetting "MyApp", "Item" & i
    Next i
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_with_concatenation() {
        // Test DeleteSetting with string concatenation
        let source = r#"
Sub Test()
    DeleteSetting "MyApp", "Section" & Num, "Key" & Index
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_in_if_statement() {
        // Test DeleteSetting in conditional
        let source = r#"
Sub Test()
    If ResetSettings Then
        DeleteSetting "MyApp", "Preferences"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_with_function_call() {
        // Test DeleteSetting with function call as argument
        let source = r#"
Sub Test()
    DeleteSetting GetAppName(), GetSection(), GetKey()
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_with_parentheses() {
        // Test DeleteSetting with parentheses around arguments
        let source = r#"
Sub Test()
    DeleteSetting ("MyApp"), ("Settings"), ("WindowState")
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]
    fn deletesetting_with_error_handling() {
        // Test DeleteSetting with error handling
        let source = r#"
Sub Test()
    On Error Resume Next
    DeleteSetting "MyApp", "Settings"
    If Err Then MsgBox "Error deleting setting"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DeleteSettingStatement"));
    }

    #[test]

    fn date_at_module_level() {
        let source = r#"
Date = #1/1/2024#
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
    }

    #[test]
    fn date_with_dateadd() {
        let source = r#"
Sub Test()
    Date = DateAdd("d", 30, Date)
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("DateStatement"));
    }

    // Reset statement tests
    #[test]
    fn reset_simple() {
        let source = r#"
Sub Test()
    Reset
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
        assert!(debug.contains("ResetKeyword"));
    }

    #[test]
    fn reset_at_module_level() {
        let source = "Reset\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
    }

    #[test]
    fn reset_in_if_statement() {
        let source = r#"
Sub CleanupFiles()
    If CloseAllFiles Then
        Reset
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
    }

    #[test]
    fn reset_preserves_whitespace() {
        let source = "    Reset    \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Reset    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
    }

    #[test]
    fn reset_with_comment() {
        let source = r#"
Sub Test()
    Reset ' Close all open files
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn reset_inline_if() {
        let source = r#"
Sub Test()
    If shouldClose Then Reset
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
    }

    #[test]
    fn reset_in_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        Reset
    Next i
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
    }

    #[test]
    fn reset_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Reset
    If Err.Number <> 0 Then
        MsgBox "Error closing files"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
    }

    #[test]
    fn multiple_reset_statements() {
        let source = r#"
Sub Test()
    Reset
    DoSomething
    Reset
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let reset_count = debug.matches("ResetStatement").count();
        assert_eq!(reset_count, 2);
    }

    #[test]
    fn reset_after_file_operations() {
        let source = r#"
Sub Test()
    Open "test.txt" For Output As #1
    Print #1, "data"
    Close #1
    Reset
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ResetStatement"));
    }

    // Error statement tests
    #[test]
    fn error_simple() {
        let source = r#"
Sub Test()
    Error 11
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
        assert!(debug.contains("ErrorKeyword"));
    }

    #[test]
    fn error_at_module_level() {
        let source = "Error 53\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn error_with_literal() {
        let source = r#"
Sub Test()
    Error 9
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn error_with_expression() {
        let source = r#"
Sub Test()
    Error vbObjectError + 1000
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
        assert!(debug.contains("vbObjectError"));
    }

    #[test]
    fn error_with_variable() {
        let source = r#"
Sub Test()
    Error errorNumber
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
        assert!(debug.contains("errorNumber"));
    }

    #[test]
    fn error_preserves_whitespace() {
        let source = "    Error    11    \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Error    11    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn error_with_comment() {
        let source = r#"
Sub Test()
    Error 11 ' Division by zero
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn error_in_if_statement() {
        let source = r#"
Sub Test()
    If shouldFail Then
        Error 5
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn error_inline_if() {
        let source = r#"
Sub Test()
    If invalidData Then Error 13
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn error_in_select_case() {
        let source = r#"
Sub Test()
    Select Case errorType
        Case 1
            Error 11
        Case 2
            Error 13
    End Select
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let error_count = debug.matches("ErrorStatement").count();
        assert_eq!(error_count, 2);
    }

    #[test]
    fn error_with_error_handler() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    DoSomething
    Exit Sub
ErrorHandler:
    Error 1000
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn error_custom_number() {
        let source = r#"
Sub Test()
    Error 32000
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("ErrorStatement"));
    }

    #[test]
    fn multiple_error_statements() {
        let source = r#"
Sub Test()
    Error 1
    DoSomething
    Error 2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let error_count = debug.matches("ErrorStatement").count();
        assert_eq!(error_count, 2);
    }

    // FileCopy statement tests
    #[test]
    fn filecopy_simple() {
        let source = r#"
Sub Test()
    FileCopy "source.txt", "dest.txt"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("FileCopyStatement"));
        assert!(debug.contains("FileCopyKeyword"));
    }

    #[test]
    fn filecopy_at_module_level() {
        let source = "FileCopy \"old.dat\", \"new.dat\"\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("FileCopyStatement"));
    }

    #[test]
    fn filecopy_with_paths() {
        let source = r#"
Sub Test()
    FileCopy "C:\\SOURCE.TXT", "C:\\DEST.TXT"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("FileCopyStatement"));
    }

    #[test]
    fn filecopy_with_variables() {
        let source = r#"
Sub Test()
    FileCopy oldFile, newFile
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("FileCopyStatement"));
        assert!(debug.contains("oldFile"));
        assert!(debug.contains("newFile"));
    }

    #[test]
    fn filecopy_with_expressions() {
        let source = r#"
Sub Test()
    FileCopy App.Path & "\\data.dat", "C:\\Backup\\data.dat"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("FileCopyStatement"));
        assert!(debug.contains("App"));
    }

    #[test]
    fn filecopy_preserves_whitespace() {
        let source = "    FileCopy    \"a.txt\"  ,  \"b.txt\"    \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    FileCopy    \"a.txt\"  ,  \"b.txt\"    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("FileCopyStatement"));
    }

    #[test]
    fn filecopy_with_comment() {
        let source = r#"
Sub Test()
    FileCopy "source.dat", "dest.dat" ' Backup file
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("FileCopyStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn filecopy_in_if_statement() {
        let source = r#"
Sub Test()
    If needBackup Then
        FileCopy "data.txt", "backup.txt"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("FileCopyStatement"));
    }

    #[test]
    fn filecopy_inline_if() {
        let source = r#"
Sub Test()
    If createBackup Then FileCopy "file.txt", "file.bak"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("FileCopyStatement"));
    }

    #[test]
    fn filecopy_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    FileCopy "source.dat", "dest.dat"
    If Err.Number <> 0 Then
        MsgBox "Error copying file"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("FileCopyStatement"));
    }

    #[test]
    fn filecopy_in_loop() {
        let source = r#"
Sub Test()
    For i = 1 To 10
        FileCopy "file" & i & ".txt", "backup" & i & ".txt"
    Next i
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("FileCopyStatement"));
    }

    #[test]
    fn multiple_filecopy_statements() {
        let source = r#"
Sub Test()
    FileCopy "file1.txt", "backup1.txt"
    FileCopy "file2.txt", "backup2.txt"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let filecopy_count = debug.matches("FileCopyStatement").count();
        assert_eq!(filecopy_count, 2);
    }

    #[test]
    fn filecopy_network_paths() {
        let source = r#"
Sub Test()
    FileCopy "\\\\server\\share\\file.dat", "C:\\local\\file.dat"
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("FileCopyStatement"));
    }

    // Get statement tests
    #[test]
    fn get_simple() {
        let source = r#"
Sub Test()
    Get #1, , myRecord
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
        assert!(debug.contains("GetKeyword"));
    }

    #[test]
    fn get_at_module_level() {
        let source = "Get #1, , myData\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn get_with_record_number() {
        let source = r#"
Sub Test()
    Get #1, recordNumber, customerData
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
        assert!(debug.contains("recordNumber"));
    }

    #[test]
    fn get_with_file_variable() {
        let source = r#"
Sub Test()
    Get fileNum, , buffer
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
        assert!(debug.contains("fileNum"));
    }

    #[test]
    fn get_with_hash_symbol() {
        let source = r#"
Sub Test()
    Get #fileNumber, position, data
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn get_preserves_whitespace() {
        let source = "    Get    #1  ,  ,  myVar    \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Get    #1  ,  ,  myVar    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn get_with_comment() {
        let source = r#"
Sub Test()
    Get #1, , myRecord ' Read next record
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn get_in_if_statement() {
        let source = r#"
Sub Test()
    If Not EOF(1) Then
        Get #1, , myData
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn get_inline_if() {
        let source = r#"
Sub Test()
    If hasData Then Get #1, , buffer
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn get_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Get #1, , myRecord
    If Err.Number <> 0 Then
        MsgBox "Error reading file"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn get_in_loop() {
        let source = r#"
Sub Test()
    Do While Not EOF(1)
        Get #1, , myRecord
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    #[test]
    fn multiple_get_statements() {
        let source = r#"
Sub Test()
    Get #1, , record1
    Get #1, , record2
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let get_count = debug.matches("GetStatement").count();
        assert_eq!(get_count, 2);
    }

    #[test]
    fn get_binary_mode() {
        let source = r#"
Sub Test()
    Dim buffer As String * 512
    Get #1, , buffer
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("GetStatement"));
    }

    // Input statement tests
    #[test]
    fn input_simple() {
        let source = r#"
Sub Test()
    Input #1, name, age
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("InputStatement"));
        assert!(debug.contains("InputKeyword"));
    }

    #[test]
    fn input_at_module_level() {
        let source = "Input #1, myData\n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.root_kind(), SyntaxKind::Root);
        assert_eq!(cst.child_count(), 1);

        let debug = cst.debug_tree();
        assert!(debug.contains("InputStatement"));
    }

    #[test]
    fn input_multiple_variables() {
        let source = r#"
Sub Test()
    Input #1, firstName, lastName, age, address
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("InputStatement"));
        assert!(debug.contains("firstName"));
    }

    #[test]
    fn input_with_file_variable() {
        let source = r#"
Sub Test()
    Input #fileNum, x, y, z
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("InputStatement"));
        assert!(debug.contains("fileNum"));
    }

    #[test]
    fn input_preserves_whitespace() {
        let source = "    Input    #1  ,  name  ,  age    \n";
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        assert_eq!(cst.text(), "    Input    #1  ,  name  ,  age    \n");

        let debug = cst.debug_tree();
        assert!(debug.contains("InputStatement"));
    }

    #[test]
    fn input_with_comment() {
        let source = r#"
Sub Test()
    Input #1, name, age ' Read person data
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("InputStatement"));
        assert!(debug.contains("Comment"));
    }

    #[test]
    fn input_in_if_statement() {
        let source = r#"
Sub Test()
    If Not EOF(1) Then
        Input #1, myData
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("InputStatement"));
    }

    #[test]
    fn input_inline_if() {
        let source = r#"
Sub Test()
    If hasData Then Input #1, buffer
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("InputStatement"));
    }

    #[test]
    fn input_with_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Input #1, name, age
    If Err.Number <> 0 Then
        MsgBox "Error reading file"
    End If
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("InputStatement"));
    }

    #[test]
    fn input_in_loop() {
        let source = r#"
Sub Test()
    Do While Not EOF(1)
        Input #1, myRecord
    Loop
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("InputStatement"));
    }

    #[test]
    fn multiple_input_statements() {
        let source = r#"
Sub Test()
    Input #1, header
    Input #1, data
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        let input_count = debug.matches("InputStatement").count();
        assert_eq!(input_count, 2);
    }

    #[test]
    fn input_sequential_file() {
        let source = r#"
Sub Test()
    Open "data.txt" For Input As #1
    Input #1, name, age, city
    Close #1
End Sub
"#;
        let cst = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();

        let debug = cst.debug_tree();
        assert!(debug.contains("InputStatement"));
    }
}
