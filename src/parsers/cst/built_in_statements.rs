//! Built-in VB6 statement parsing.
//!
//! This module handles parsing of VB6 built-in system statements:
//! - AppActivate: Activate an application window
//! - Beep: Sound a tone through the computer's speaker
//! - ChDir: Change the current directory or folder
//! - ChDrive: Change the current drive

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("AppActivateStatement"));
    }

    #[test]
    fn appactivate_at_module_level() {
        let source = r#"
AppActivate "MyApp"
"#;
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("BeepStatement"));
        assert!(debug.contains("Whitespace"));
    }

    #[test]
    fn beep_at_module_level() {
        let source = r#"
Beep
"#;
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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
        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }

    #[test]
    fn chdir_at_module_level() {
        let source = r#"
ChDir "C:\Temp"
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
    }

    #[test]
    fn chdrive_at_module_level() {
        let source = r#"
ChDrive "D:"
"#;

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

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

        let mut source_stream = SourceStream::new("test.bas", source);
        let result = tokenize(&mut source_stream);
        let token_stream = result.result.expect("Tokenization should succeed");
        let cst = parse(token_stream);

        let debug = cst.debug_tree();
        assert!(debug.contains("ChDriveStatement"));
        assert!(debug.contains("ChDriveKeyword"));
        assert!(debug.contains("ChDirStatement"));
        assert!(debug.contains("ChDirKeyword"));
    }
}
