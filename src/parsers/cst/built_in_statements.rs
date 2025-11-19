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
    /// Parse an AppActivate statement.
    ///
    /// VB6 AppActivate statement syntax:
    /// - AppActivate title[, wait]
    ///
    /// Activates an application window.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/appactivate-statement)
    pub(super) fn parse_appactivate_statement(&mut self) {
        // if we are now parsing an AppActivate statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::AppActivateStatement.to_raw());

        // Consume "AppActivate" keyword
        self.consume_token();

        // Consume everything until newline (title and optional wait parameter)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // AppActivateStatement
    }

    /// Parse a Beep statement.
    ///
    /// VB6 Beep statement syntax:
    /// - Beep
    ///
    /// Sounds a tone through the computer's speaker.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/beep-statement)
    pub(super) fn parse_beep_statement(&mut self) {
        // if we are now parsing a beep statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::BeepStatement.to_raw());

        // Consume "Beep" keyword
        self.consume_token();

        // Consume any whitespace and comments until newline
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // BeepStatement
    }

    /// Parse a ChDir statement.
    ///
    /// VB6 ChDir statement syntax:
    /// - ChDir path
    ///
    /// Changes the current directory or folder.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chdir-statement)
    pub(super) fn parse_chdir_statement(&mut self) {
        // if we are now parsing a ChDir statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder.start_node(SyntaxKind::ChDirStatement.to_raw());

        // Consume "ChDir" keyword
        self.consume_token();

        // Consume everything until newline (the path parameter)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // ChDirStatement
    }

    /// Parse a ChDrive statement.
    ///
    /// VB6 ChDrive statement syntax:
    /// - ChDrive drive
    ///
    /// Changes the current drive.
    ///
    /// [Reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chdrive-statement)
    pub(super) fn parse_chdrive_statement(&mut self) {
        // if we are now parsing a ChDrive statement, we are no longer in the header.
        self.parsing_header = false;

        self.builder
            .start_node(SyntaxKind::ChDriveStatement.to_raw());

        // Consume "ChDrive" keyword
        self.consume_token();

        // Consume everything until newline (the drive parameter)
        self.consume_until(VB6Token::Newline);

        // Consume the newline
        if self.at_token(VB6Token::Newline) {
            self.consume_token();
        }

        self.builder.finish_node(); // ChDriveStatement
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
}
